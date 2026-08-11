using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.Agent.Models;

public enum AgentProgressStage
{
    Accepted,
    ValidatingInput,
    DiscoveringModels,
    RequestingProposal,
    ValidatingProposal,
    AwaitingHostTool,
    ProcessingHostResult,
    RepairingProposal,
    Cancelling,
    Completed,
    Cancelled,
    Failed,
}

public sealed class AgentProgressEvent
{
    public string JobId { get; set; } = string.Empty;

    public long Sequence { get; set; }

    public AgentProgressStage Stage { get; set; }

    public string Message { get; set; } = string.Empty;

    public int? CompletedUnits { get; set; }

    public int? TotalUnits { get; set; }

    public DateTimeOffset OccurredAtUtc { get; set; }
}

public sealed class AgentCheckpointEvent
{
    public string JobId { get; set; } = string.Empty;

    public string WorkbookId { get; set; } = string.Empty;

    public long Sequence { get; set; }

    public string CheckpointId { get; set; } = string.Empty;

    public AgentProgressStage Stage { get; set; }

    public int CompletedRepairCycles { get; set; }

    public string LastCompletedStep { get; set; } = string.Empty;

    public DateTimeOffset OccurredAtUtc { get; set; }
}

public interface IAgentEventSink
{
    Task PublishProgressAsync(AgentProgressEvent progress, CancellationToken cancellationToken);

    Task PublishCheckpointAsync(AgentCheckpointEvent checkpoint, CancellationToken cancellationToken);
}

public sealed class NullAgentEventSink : IAgentEventSink
{
    public static NullAgentEventSink Instance { get; } = new NullAgentEventSink();

    private NullAgentEventSink()
    {
    }

    public Task PublishProgressAsync(AgentProgressEvent progress, CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }

    public Task PublishCheckpointAsync(AgentCheckpointEvent checkpoint, CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }
}

public sealed class HelloRequest
{
    public string ClientName { get; set; } = string.Empty;

    public List<string> SupportedProtocolVersions { get; set; } = new List<string>();
}

public sealed class HelloAcknowledgement
{
    public string ProtocolVersion { get; set; } = string.Empty;

    public string WorkerVersion { get; set; } = string.Empty;

    public bool CurrentUserOnlyPipe { get; set; }
}

public sealed class StartJobRequest
{
    public AgentJobRequest Job { get; set; } = new AgentJobRequest();
}

public sealed class HostToolRequestEvent
{
    public string JobId { get; set; } = string.Empty;

    public string WorkbookId { get; set; } = string.Empty;

    public string ToolCallId { get; set; } = string.Empty;

    public string ToolName { get; set; } = string.Empty;

    public string ArgumentsJson { get; set; } = "{}";

    public DateTimeOffset RequestedAtUtc { get; set; }
}

public sealed class HostToolResultRequest
{
    public string JobId { get; set; } = string.Empty;

    public string ToolCallId { get; set; } = string.Empty;

    public bool Succeeded { get; set; }

    public string OutcomeCode { get; set; } = string.Empty;

    /// <summary>
    /// Bounded structured result from deterministic host code. Never include
    /// workbook values, paths, prompts, credentials, or exception details.
    /// </summary>
    public string ResultJson { get; set; } = "{}";

    public List<HostCheckFailure> CheckFailures { get; set; } = new List<HostCheckFailure>();
}

public sealed class HostCheckFailure
{
    public string Code { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;
}

public sealed class ResumeMetadataRequest
{
    public string? WorkbookId { get; set; }
}

public sealed class ResumeMetadataResponse
{
    public List<AgentResumeMetadata> Jobs { get; set; } = new List<AgentResumeMetadata>();
}

public sealed class AgentResumeMetadata
{
    public string JobId { get; set; } = string.Empty;

    public string WorkbookId { get; set; } = string.Empty;

    public string CheckpointId { get; set; } = string.Empty;

    public AgentProgressStage Stage { get; set; }

    public int CompletedRepairCycles { get; set; }

    public string LastCompletedStep { get; set; } = string.Empty;

    public string? FailureCode { get; set; }

    public bool CanResume { get; set; }

    public DateTimeOffset UpdatedAtUtc { get; set; }
}

public sealed class CancelJobRequest
{
    public string JobId { get; set; } = string.Empty;
}

public sealed class CancelJobAcknowledgement
{
    public string JobId { get; set; } = string.Empty;

    public bool CancellationRequested { get; set; }
}

public sealed class EndpointProbeRequest
{
    public AgentEndpointSettings Endpoint { get; set; } = new AgentEndpointSettings();
}

public sealed class JobCompletedEvent
{
    public AgentRunResult Result { get; set; } = new AgentRunResult();
}

public sealed class JobFailedEvent
{
    public string JobId { get; set; } = string.Empty;

    public RedactedDiagnostic Diagnostic { get; set; } = new RedactedDiagnostic();

    public AgentResumeMetadata? ResumeMetadata { get; set; }
}

public sealed class JobCancelledEvent
{
    public string JobId { get; set; } = string.Empty;

    public string Message { get; set; } = "The job was cancelled.";
}

public sealed class ProtocolErrorEvent
{
    public string Code { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;
}

public sealed class RedactedDiagnostic
{
    public string Code { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;

    public bool Retryable { get; set; }
}
