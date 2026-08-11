using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Diagnostics;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.OpenAI;
using ExcelReportBuilder.Agent.Tools;
using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.Agent.Execution;

public sealed class GuardedAgentRunner
{
    private readonly IOpenAiCompatibleClient _client;
    private readonly IAgentEventSink _eventSink;
    private readonly IAgentHostToolBridge _hostToolBridge;
    private readonly TimeSpan _heartbeatInterval;
    private readonly Func<DateTimeOffset> _utcNow;
    private long _sequence;

    public GuardedAgentRunner(
        IOpenAiCompatibleClient client,
        IAgentEventSink? eventSink = null,
        TimeSpan? heartbeatInterval = null,
        Func<DateTimeOffset>? utcNow = null,
        IAgentHostToolBridge? hostToolBridge = null)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _eventSink = eventSink ?? NullAgentEventSink.Instance;
        _hostToolBridge = hostToolBridge ?? UnavailableAgentHostToolBridge.Instance;
        _heartbeatInterval = heartbeatInterval ?? TimeSpan.FromSeconds(5);
        if (_heartbeatInterval <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(heartbeatInterval));
        }

        _utcNow = utcNow ?? (() => DateTimeOffset.UtcNow);
    }

    public async Task<AgentRunResult> RunAsync(
        AgentJobRequest request,
        CancellationToken cancellationToken)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        _sequence = 0;

        await ProgressAsync(
            request.JobId,
            AgentProgressStage.Accepted,
            "The worker accepted the report request.",
            null,
            null,
            cancellationToken).ConfigureAwait(false);

        await ProgressAsync(
            request.JobId,
            AgentProgressStage.ValidatingInput,
            "Checking prompt, data description, endpoint policy, and repair limits.",
            null,
            null,
            cancellationToken).ConfigureAwait(false);
        AgentRequestValidator.Validate(request);
        await CheckpointAsync(
            request.JobId,
            request.WorkbookId,
            AgentProgressStage.ValidatingInput,
            0,
            "Validated bounded input.",
            cancellationToken).ConfigureAwait(false);

        await ProgressAsync(
            request.JobId,
            AgentProgressStage.DiscoveringModels,
            "Reading available models from the endpoint.",
            null,
            null,
            cancellationToken).ConfigureAwait(false);
        ModelDiscoveryResult? discovery = null;
        try
        {
            discovery = await WithHeartbeatAsync(
                _client.DiscoverModelsAsync(request.Endpoint, cancellationToken),
                request.JobId,
                AgentProgressStage.DiscoveringModels,
                "Still waiting for the endpoint model list.",
                cancellationToken).ConfigureAwait(false);
        }
        catch (AgentEndpointException)
        {
            // Model discovery is optional. The configured model is verified by
            // the first authenticated completion request instead.
        }

        if (discovery != null &&
            !discovery.ModelIds.Any(id => string.Equals(id, request.Endpoint.Model, StringComparison.Ordinal)))
        {
            throw new AgentRunException(
                "configured_model_unavailable",
                "The configured model was not returned by the models endpoint.",
                false);
        }

        await CheckpointAsync(
            request.JobId,
            request.WorkbookId,
            AgentProgressStage.DiscoveringModels,
            0,
            discovery == null
                ? "Model discovery is unavailable; the configured model will be verified by completion."
                : "Confirmed the configured model is available.",
            cancellationToken).ConfigureAwait(false);

        var completedTools = new HashSet<string>(StringComparer.Ordinal);
        var acceptedCallIds = new HashSet<string>(StringComparer.Ordinal);
        var acceptedCalls = new List<ValidatedAgentToolCall>();
        var acceptedResults = new List<HostToolResultSummary>();
        string? workflowInstruction = request.ResumeFromCheckpointId == null
            ? null
            : "The host supplied a resumable checkpoint. Reconfirm the workflow from the current bounded request before continuing.";
        var repairCycles = 0;
        var workflowTurns = 0;
        var maximumWorkflowTurns = (request.MaxRepairCycles + 1) * 10;

        while (workflowTurns < maximumWorkflowTurns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            workflowTurns++;
            await ProgressAsync(
                request.JobId,
                AgentProgressStage.RequestingProposal,
                repairCycles == 0
                    ? "Requesting the next bounded report workflow tool call."
                    : "Requesting the next tool call in the repaired report workflow.",
                workflowTurns,
                maximumWorkflowTurns,
                cancellationToken).ConfigureAwait(false);

            var proposal = await WithHeartbeatAsync(
                _client.RequestToolProposalAsync(request, workflowInstruction, cancellationToken),
                request.JobId,
                AgentProgressStage.RequestingProposal,
                "The model is still preparing the next bounded workflow tool call.",
                cancellationToken).ConfigureAwait(false);

            await ProgressAsync(
                request.JobId,
                AgentProgressStage.ValidatingProposal,
                "Validating tool names, JSON shapes, source columns, and allowlisted values.",
                workflowTurns,
                maximumWorkflowTurns,
                cancellationToken).ConfigureAwait(false);
            var validation = AgentToolCallValidator.Validate(
                proposal.ToolCalls,
                request.Data,
                requireCompleteWorkflow: false);
            if (validation.IsValid && validation.ToolCalls.Count != 1)
            {
                validation = new AgentToolValidationResult
                {
                    IsValid = false,
                    ErrorCode = "parallel_tool_calls_not_allowed",
                    RepairInstruction = "Return exactly one workflow tool call, then wait for its deterministic host result before choosing the next tool.",
                };
            }

            if (validation.IsValid)
            {
                var toolCall = validation.ToolCalls[0];
                var flowError = acceptedCallIds.Contains(toolCall.Id)
                    ? "Use a new tool-call ID after every deterministic host round trip."
                    : ValidateNextWorkflowTool(toolCall.Name, completedTools);
                if (flowError != null)
                {
                    validation = new AgentToolValidationResult
                    {
                        IsValid = false,
                        ErrorCode = "workflow_order_invalid",
                        RepairInstruction = flowError,
                    };
                }
            }

            if (!validation.IsValid)
            {
                if (repairCycles >= request.MaxRepairCycles)
                {
                    throw RepairLimitReached();
                }

                repairCycles++;
                workflowInstruction = validation.RepairInstruction;
                completedTools.Clear();
                acceptedCallIds.Clear();
                acceptedCalls.Clear();
                acceptedResults.Clear();
                await PublishRepairAsync(
                    request,
                    repairCycles,
                    "Rejected an invalid workflow tool call before host execution.",
                    cancellationToken).ConfigureAwait(false);
                continue;
            }

            var acceptedCall = validation.ToolCalls[0];
            await CheckpointAsync(
                request.JobId,
                request.WorkbookId,
                AgentProgressStage.ValidatingProposal,
                repairCycles,
                "Validated workflow tool call " + acceptedCall.Name + ".",
                cancellationToken).ConfigureAwait(false);

            var hostOutcome = await InvokeHostToolsAsync(
                request,
                validation.ToolCalls,
                repairCycles,
                cancellationToken).ConfigureAwait(false);
            if (!hostOutcome.Succeeded)
            {
                if (repairCycles >= request.MaxRepairCycles)
                {
                    throw RepairLimitReached();
                }

                repairCycles++;
                workflowInstruction = hostOutcome.RepairInstruction;
                completedTools.Clear();
                acceptedCallIds.Clear();
                acceptedCalls.Clear();
                acceptedResults.Clear();
                await PublishRepairAsync(
                    request,
                    repairCycles,
                    "A deterministic host gate rejected the workflow.",
                    cancellationToken).ConfigureAwait(false);
                continue;
            }

            completedTools.Add(acceptedCall.Name);
            acceptedCallIds.Add(acceptedCall.Id);
            acceptedCalls.Add(acceptedCall);
            acceptedResults.AddRange(hostOutcome.Summaries);
            if (string.Equals(acceptedCall.Name, AgentToolNames.ProposePeriodMapping, StringComparison.Ordinal))
            {
                ApplyAcceptedPeriodMapping(request.Data, acceptedCall.ArgumentsJson);
            }
            else if (string.Equals(acceptedCall.Name, AgentToolNames.ProposeTransforms, StringComparison.Ordinal))
            {
                ApplyAcceptedTransforms(request.Data, acceptedCall.ArgumentsJson);
            }
            if (string.Equals(acceptedCall.Name, AgentToolNames.FinalChangeSummary, StringComparison.Ordinal))
            {
                await ProgressAsync(
                    request.JobId,
                    AgentProgressStage.Completed,
                    "The deterministic host accepted the guarded workflow and checks.",
                    workflowTurns,
                    maximumWorkflowTurns,
                    cancellationToken).ConfigureAwait(false);

                return new AgentRunResult
                {
                    JobId = request.JobId,
                    WorkbookId = request.WorkbookId,
                    Model = string.IsNullOrWhiteSpace(proposal.Model) ? request.Endpoint.Model : proposal.Model,
                    RepairCyclesUsed = repairCycles,
                    ToolCalls = acceptedCalls,
                    HostToolResults = acceptedResults,
                };
            }

            workflowInstruction = BuildContinuationInstruction(
                acceptedCall,
                hostOutcome.Results[0],
                completedTools,
                request);
        }

        throw new AgentRunException(
            "workflow_turn_limit_reached",
            "The model did not finish the guarded workflow within the bounded tool-turn limit.",
            false);
    }

    private static void ApplyAcceptedPeriodMapping(AgentDataSnapshot data, string argumentsJson)
    {
        using (var document = JsonDocument.Parse(argumentsJson))
        {
            JsonElement root = document.RootElement;
            if (!string.Equals(
                    root.GetProperty("mode").GetString(),
                    "widePeriods",
                    StringComparison.Ordinal))
            {
                return;
            }

            JsonElement mappings = root.GetProperty("mappings");
            var sourceFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasMetrics = false;
            foreach (var mapping in mappings.EnumerateArray())
            {
                sourceFields.Add(mapping.GetProperty("sourceField").GetString() ?? string.Empty);
                hasMetrics = hasMetrics ||
                    (mapping.TryGetProperty("metric", out var metric) &&
                     !string.IsNullOrWhiteSpace(metric.GetString()));
            }

            data.Fields.RemoveAll(field => sourceFields.Contains(field.Name));
            UpsertField(data.Fields, "Period", AgentFieldType.Date, false);
            UpsertField(data.Fields, "Value", AgentFieldType.Number, true);
            if (hasMetrics)
            {
                UpsertField(data.Fields, "Metric", AgentFieldType.Text, false);
            }

            foreach (var row in data.SampleRows)
            {
                row.Values.RemoveAll(value => sourceFields.Contains(value.Field));
                UpsertSample(row.Values, "Period");
                UpsertSample(row.Values, "Value");
                if (hasMetrics) UpsertSample(row.Values, "Metric");
            }

            try
            {
                data.RowCount = checked(data.RowCount * Math.Max(1, mappings.GetArrayLength()));
            }
            catch (OverflowException)
            {
                data.RowCount = long.MaxValue;
            }
        }
    }

    private static void UpsertField(
        ICollection<AgentField> fields,
        string name,
        AgentFieldType type,
        bool allowsBlank)
    {
        var existing = fields.FirstOrDefault(field =>
            string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            existing.Type = type;
            existing.AllowsBlank = allowsBlank;
            return;
        }

        fields.Add(new AgentField
        {
            Name = name,
            Type = type,
            AllowsBlank = allowsBlank
        });
    }

    private static void UpsertSample(ICollection<AgentSampleValue> values, string field)
    {
        if (values.Any(value => string.Equals(value.Field, field, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        values.Add(new AgentSampleValue { Field = field, Value = null });
    }

    private static void ApplyAcceptedTransforms(AgentDataSnapshot data, string argumentsJson)
    {
        using (var document = JsonDocument.Parse(argumentsJson))
        {
            foreach (var transform in document.RootElement.GetProperty("transforms").EnumerateArray())
            {
                var kind = transform.GetProperty("kind").GetString() ?? string.Empty;
                var sourceField = transform.GetProperty("sourceField").GetString() ?? string.Empty;
                var outputField = transform.GetProperty("outputField").GetString() ?? string.Empty;
                var source = data.Fields.FirstOrDefault(field =>
                    string.Equals(field.Name, sourceField, StringComparison.OrdinalIgnoreCase));
                if (source == null)
                {
                    continue;
                }

                var outputType = kind == "convertNumber"
                    ? AgentFieldType.Number
                    : kind == "convertDate"
                        ? AgentFieldType.Date
                        : kind == "addArithmeticColumn"
                            ? AgentFieldType.Number
                            : kind == "derivePeriodPart"
                                ? AgentFieldType.Text
                        : source.Type;
                data.Fields.RemoveAll(field =>
                    string.Equals(field.Name, outputField, StringComparison.OrdinalIgnoreCase));
                bool createsAdditionalField = string.Equals(
                        kind,
                        "addArithmeticColumn",
                        StringComparison.Ordinal) ||
                    string.Equals(kind, "derivePeriodPart", StringComparison.Ordinal);
                if (!createsAdditionalField &&
                    !string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase))
                {
                    data.Fields.Remove(source);
                }

                data.Fields.Add(new AgentField
                {
                    Name = outputField,
                    Type = outputType,
                    AllowsBlank = source.AllowsBlank
                });
                if (createsAdditionalField)
                {
                    foreach (var row in data.SampleRows)
                    {
                        row.Values.RemoveAll(value =>
                            string.Equals(value.Field, outputField, StringComparison.OrdinalIgnoreCase));
                        row.Values.Add(new AgentSampleValue
                        {
                            Field = outputField,
                            Value = null
                        });
                    }
                }
                else
                {
                    foreach (var row in data.SampleRows)
                    {
                        var sample = row.Values.FirstOrDefault(value =>
                            string.Equals(value.Field, sourceField, StringComparison.OrdinalIgnoreCase));
                        if (sample != null)
                        {
                            sample.Field = outputField;
                            ApplySampleTransform(sample, transform, kind);
                        }
                    }
                }
            }
        }
    }

    private static void ApplySampleTransform(
        AgentSampleValue sample,
        JsonElement transform,
        string kind)
    {
        if (sample.Value == null) return;
        if (string.Equals(kind, "trimText", StringComparison.Ordinal))
        {
            sample.Value = sample.Value.Trim();
            return;
        }

        if (!string.Equals(kind, "mapValues", StringComparison.Ordinal)) return;
        foreach (var mapping in transform.GetProperty("mappings").EnumerateArray())
        {
            string from = mapping.GetProperty("from").GetString() ?? string.Empty;
            if (string.Equals(sample.Value, from, StringComparison.Ordinal))
            {
                sample.Value = mapping.GetProperty("to").GetString();
                return;
            }
        }
    }

    private async Task PublishRepairAsync(
        AgentJobRequest request,
        int repairCycles,
        string checkpointMessage,
        CancellationToken cancellationToken)
    {
        await ProgressAsync(
            request.JobId,
            AgentProgressStage.RepairingProposal,
            "Preparing a bounded validation-driven repair cycle.",
            repairCycles,
            request.MaxRepairCycles,
            cancellationToken).ConfigureAwait(false);
        await CheckpointAsync(
            request.JobId,
            request.WorkbookId,
            AgentProgressStage.RepairingProposal,
            repairCycles,
            checkpointMessage,
            cancellationToken).ConfigureAwait(false);
    }

    private static AgentRunException RepairLimitReached()
    {
        return new AgentRunException(
            "repair_limit_reached",
            "The model did not produce a valid checked report workflow within the bounded repair-cycle limit.",
            false);
    }

    private static string? ValidateNextWorkflowTool(
        string toolName,
        HashSet<string> completedTools)
    {
        if (completedTools.Contains(toolName))
        {
            return "Do not repeat an accepted workflow tool call. Continue with the next required tool.";
        }

        if (!completedTools.Contains(AgentToolNames.ProposeReportSpec))
        {
            if (string.Equals(toolName, AgentToolNames.ProposePeriodMapping, StringComparison.Ordinal))
            {
                return completedTools.Contains(AgentToolNames.ProposeTransforms)
                    ? "Period mapping must be proposed before transforms. Restart with the report specification."
                    : null;
            }

            if (string.Equals(toolName, AgentToolNames.ProposeTransforms, StringComparison.Ordinal) ||
                string.Equals(toolName, AgentToolNames.ProposeReportSpec, StringComparison.Ordinal))
            {
                return null;
            }

            return "Propose the report specification before validation, managed draft build, checks, or final summary.";
        }

        if (!completedTools.Contains(AgentToolNames.ValidateSpec))
        {
            return string.Equals(toolName, AgentToolNames.ValidateSpec, StringComparison.Ordinal)
                ? null
                : "Call validate_spec after the accepted report specification.";
        }

        if (!completedTools.Contains(AgentToolNames.RequestManagedDraftBuild))
        {
            return string.Equals(toolName, AgentToolNames.RequestManagedDraftBuild, StringComparison.Ordinal)
                ? null
                : "Call request_managed_draft_build after deterministic specification validation.";
        }

        if (!completedTools.Contains(AgentToolNames.RunChecks))
        {
            return string.Equals(toolName, AgentToolNames.RunChecks, StringComparison.Ordinal)
                ? null
                : "Call run_checks after the managed draft build.";
        }

        return string.Equals(toolName, AgentToolNames.FinalChangeSummary, StringComparison.Ordinal)
            ? null
            : "Call final_change_summary only after independent checks pass.";
    }

    private static string BuildContinuationInstruction(
        ValidatedAgentToolCall toolCall,
        HostToolResultRequest hostResult,
        HashSet<string> completedTools,
        AgentJobRequest job)
    {
        var boundedResult = hostResult.ResultJson.Length > 4096
            ? hostResult.ResultJson.Substring(0, 4096)
            : hostResult.ResultJson;
        boundedResult = DiagnosticRedactor.Redact(
            boundedResult,
            new[] { job.UserPrompt, job.Endpoint.ApiKey });
        var next = NextRequiredTool(completedTools);
        return "The deterministic host accepted " + toolCall.Name +
               " with outcome " + hostResult.OutcomeCode +
               " and bounded structured result " + boundedResult +
               ". Return exactly one call to " + next + " and wait for its host result.";
    }

    private static string NextRequiredTool(HashSet<string> completedTools)
    {
        if (!completedTools.Contains(AgentToolNames.ProposeReportSpec)) return AgentToolNames.ProposeReportSpec;
        if (!completedTools.Contains(AgentToolNames.ValidateSpec)) return AgentToolNames.ValidateSpec;
        if (!completedTools.Contains(AgentToolNames.RequestManagedDraftBuild)) return AgentToolNames.RequestManagedDraftBuild;
        if (!completedTools.Contains(AgentToolNames.RunChecks)) return AgentToolNames.RunChecks;
        return AgentToolNames.FinalChangeSummary;
    }

    private async Task<HostToolBatchOutcome> InvokeHostToolsAsync(
        AgentJobRequest job,
        IReadOnlyList<ValidatedAgentToolCall> toolCalls,
        int repairCycle,
        CancellationToken cancellationToken)
    {
        var outcome = new HostToolBatchOutcome();
        var repair = new StringBuilder(
            "The deterministic host rejected part of the workflow. Correct the proposal using these bounded outcomes: ");

        foreach (var toolCall in toolCalls)
        {
            await ProgressAsync(
                job.JobId,
                AgentProgressStage.AwaitingHostTool,
                HostToolProgressMessage(toolCall.Name),
                null,
                null,
                cancellationToken).ConfigureAwait(false);
            var hostRequest = new HostToolRequestEvent
            {
                JobId = job.JobId,
                WorkbookId = job.WorkbookId,
                ToolCallId = toolCall.Id,
                ToolName = toolCall.Name,
                ArgumentsJson = toolCall.ArgumentsJson,
                RequestedAtUtc = _utcNow(),
            };

            HostToolResultRequest hostResult;
            try
            {
                hostResult = await WithHeartbeatAsync(
                    _hostToolBridge.InvokeAsync(hostRequest, cancellationToken),
                    job.JobId,
                    AgentProgressStage.AwaitingHostTool,
                    "Still waiting for the deterministic host tool result.",
                    cancellationToken).ConfigureAwait(false);
                HostToolResultValidator.Validate(hostRequest, hostResult);
            }
            catch (AgentInputValidationException exception)
            {
                hostResult = new HostToolResultRequest
                {
                    JobId = job.JobId,
                    ToolCallId = toolCall.Id,
                    Succeeded = false,
                    OutcomeCode = exception.Code,
                    ResultJson = "{}",
                };
            }

            await ProgressAsync(
                job.JobId,
                AgentProgressStage.ProcessingHostResult,
                "Received and checked the deterministic result for " + toolCall.Name + ".",
                null,
                null,
                cancellationToken).ConfigureAwait(false);
            await CheckpointAsync(
                job.JobId,
                job.WorkbookId,
                AgentProgressStage.ProcessingHostResult,
                repairCycle,
                "Processed deterministic host result for " + toolCall.Name + ".",
                cancellationToken).ConfigureAwait(false);

            var accepted = hostResult.Succeeded && hostResult.CheckFailures.Count == 0;
            outcome.Results.Add(hostResult);
            outcome.Summaries.Add(new HostToolResultSummary
            {
                ToolCallId = toolCall.Id,
                ToolName = toolCall.Name,
                Succeeded = accepted,
                OutcomeCode = hostResult.OutcomeCode,
                CheckFailureCount = hostResult.CheckFailures.Count,
            });

            if (accepted) continue;
            outcome.Succeeded = false;
            repair.Append(toolCall.Name).Append('=').Append(hostResult.OutcomeCode).Append(". ");
            foreach (var failure in hostResult.CheckFailures)
            {
                repair.Append(failure.Code)
                    .Append(": ")
                    .Append(DiagnosticRedactor.Redact(
                        failure.Message,
                        new[] { job.UserPrompt, job.Endpoint.ApiKey }))
                    .Append(". ");
            }

            // Never continue to a managed draft build or later workflow step
            // after an earlier deterministic gate rejects the proposal.
            break;
        }

        if (outcome.Succeeded)
        {
            outcome.RepairInstruction = string.Empty;
            return outcome;
        }

        outcome.RepairInstruction = repair.Length > 4096
            ? repair.ToString(0, 4096)
            : repair.ToString();
        return outcome;
    }

    private static string HostToolProgressMessage(string toolName)
    {
        switch (toolName)
        {
            case AgentToolNames.ValidateSpec:
                return "Waiting for deterministic report specification validation.";
            case AgentToolNames.RequestManagedDraftBuild:
                return "Waiting for the host to build the explicitly managed draft.";
            case AgentToolNames.RunChecks:
                return "Waiting for independent managed-draft checks.";
            default:
                return "Waiting for the host to process " + toolName + ".";
        }
    }

    private async Task<T> WithHeartbeatAsync<T>(
        Task<T> operation,
        string jobId,
        AgentProgressStage stage,
        string heartbeatMessage,
        CancellationToken cancellationToken)
    {
        while (!operation.IsCompleted)
        {
            var delay = Task.Delay(_heartbeatInterval, cancellationToken);
            var completed = await Task.WhenAny(operation, delay).ConfigureAwait(false);
            if (completed == operation)
            {
                break;
            }

            cancellationToken.ThrowIfCancellationRequested();
            await ProgressAsync(
                jobId,
                stage,
                heartbeatMessage,
                null,
                null,
                cancellationToken).ConfigureAwait(false);
        }

        return await operation.ConfigureAwait(false);
    }

    private Task ProgressAsync(
        string jobId,
        AgentProgressStage stage,
        string message,
        int? completedUnits,
        int? totalUnits,
        CancellationToken cancellationToken)
    {
        return _eventSink.PublishProgressAsync(
            new AgentProgressEvent
            {
                JobId = jobId,
                Sequence = Interlocked.Increment(ref _sequence),
                Stage = stage,
                Message = message,
                CompletedUnits = completedUnits,
                TotalUnits = totalUnits,
                OccurredAtUtc = _utcNow(),
            },
            cancellationToken);
    }

    private Task CheckpointAsync(
        string jobId,
        string workbookId,
        AgentProgressStage stage,
        int completedRepairCycles,
        string lastCompletedStep,
        CancellationToken cancellationToken)
    {
        var sequence = Interlocked.Increment(ref _sequence);
        return _eventSink.PublishCheckpointAsync(
            new AgentCheckpointEvent
            {
                JobId = jobId,
                WorkbookId = workbookId,
                Sequence = sequence,
                CheckpointId = Guid.NewGuid().ToString("N"),
                Stage = stage,
                CompletedRepairCycles = completedRepairCycles,
                LastCompletedStep = lastCompletedStep,
                OccurredAtUtc = _utcNow(),
            },
            cancellationToken);
    }

    private sealed class HostToolBatchOutcome
    {
        public bool Succeeded { get; set; } = true;

        public string RepairInstruction { get; set; } = string.Empty;

        public List<HostToolResultSummary> Summaries { get; } = new List<HostToolResultSummary>();

        public List<HostToolResultRequest> Results { get; } = new List<HostToolResultRequest>();
    }
}

public sealed class AgentRunException : Exception
{
    public AgentRunException(string code, string message, bool retryable, Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        Retryable = retryable;
    }

    public string Code { get; }

    public bool Retryable { get; }
}
