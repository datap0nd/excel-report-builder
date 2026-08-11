using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Execution;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.OpenAI;
using ExcelReportBuilder.Agent.Tools;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class GuardedAgentRunnerTests
{
    [Fact]
    public async Task RunAsync_RepairsRejectedProposalAndEmitsCheckpoints()
    {
        var client = new FakeClient();
        client.Proposals.Enqueue(new AgentModelProposal
        {
            Model = AgentDefaults.Model,
            ToolCalls = { new AgentToolCall { Id = "bad", Name = "run_shell", ArgumentsJson = "{}" } },
        });
        EnqueueWorkflow(client);
        var sink = new RecordingSink();
        var bridge = new FakeHostToolBridge();
        var runner = new GuardedAgentRunner(client, sink, hostToolBridge: bridge);

        var result = await runner.RunAsync(SyntheticJob.Create(2), CancellationToken.None);

        Assert.Equal(1, result.RepairCyclesUsed);
        Assert.Equal(5, result.ToolCalls.Count);
        Assert.Equal(5, result.HostToolResults.Count);
        Assert.Equal(6, client.ProposalRequests);
        Assert.Contains(client.RepairInstructions, value =>
            value != null && value.Contains("one call to validate_spec", StringComparison.Ordinal));
        Assert.Contains(sink.Progress, item => item.Stage == AgentProgressStage.RepairingProposal);
        Assert.Contains(sink.Checkpoints, item => item.CompletedRepairCycles == 1);
        Assert.Equal(AgentProgressStage.Completed, sink.Progress.Last().Stage);
        Assert.True(sink.Sequences.SequenceEqual(sink.Sequences.OrderBy(value => value)));
    }

    [Fact]
    public async Task RunAsync_StopsAtBoundedRepairLimit()
    {
        var client = new FakeClient { AlwaysInvalid = true };
        var runner = new GuardedAgentRunner(client);

        var error = await Assert.ThrowsAsync<AgentRunException>(() =>
            runner.RunAsync(SyntheticJob.Create(2), CancellationToken.None));

        Assert.Equal("repair_limit_reached", error.Code);
        Assert.Equal(3, client.ProposalRequests);
    }

    [Fact]
    public async Task RunAsync_CheckFailureDrivesBoundedModelRepair()
    {
        var client = new FakeClient();
        foreach (var proposal in WorkflowProposals().Take(4)) client.Proposals.Enqueue(proposal);
        EnqueueWorkflow(client);
        var bridge = new FakeHostToolBridge { FailFirstChecksCall = true };
        var runner = new GuardedAgentRunner(client, hostToolBridge: bridge);

        var result = await runner.RunAsync(SyntheticJob.Create(2), CancellationToken.None);

        Assert.Equal(1, result.RepairCyclesUsed);
        Assert.Contains(client.RepairInstructions, value =>
            value != null && value.Contains("period_coverage_failed", StringComparison.Ordinal));
        Assert.Equal(9, bridge.Invocations.Count);
        Assert.Equal(AgentToolNames.RunChecks, bridge.Invocations[3].ToolName);
        Assert.Equal(AgentToolNames.ProposeReportSpec, bridge.Invocations[4].ToolName);
        Assert.Equal(5, result.HostToolResults.Count);
    }

    [Fact]
    public async Task RunAsync_RequiresHostRoundTripBeforeEveryNextTool()
    {
        var client = new FakeClient();
        var parallelProposal = new AgentModelProposal { Model = AgentDefaults.Model };
        parallelProposal.ToolCalls.AddRange(ToolCallValidatorTests.ApprovedWorkflow());
        client.Proposals.Enqueue(parallelProposal);
        EnqueueWorkflow(client);
        var bridge = new FakeHostToolBridge();
        var runner = new GuardedAgentRunner(client, hostToolBridge: bridge);

        var result = await runner.RunAsync(SyntheticJob.Create(1), CancellationToken.None);

        Assert.Equal(1, result.RepairCyclesUsed);
        Assert.Equal(5, bridge.Invocations.Count);
        Assert.Contains(client.RepairInstructions, value =>
            value != null && value.Contains("exactly one workflow tool call", StringComparison.Ordinal));
    }

    [Fact]
    public async Task RunAsync_UsesCallerCancellationWithoutGlobalTimeout()
    {
        var client = new FakeClient { BlockProposalUntilCancellation = true };
        var sink = new RecordingSink();
        var runner = new GuardedAgentRunner(
            client,
            sink,
            TimeSpan.FromMilliseconds(5),
            hostToolBridge: new FakeHostToolBridge());
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(80));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            runner.RunAsync(SyntheticJob.Create(), cancellation.Token));

        Assert.Contains(
            sink.Progress,
            item => item.Stage == AgentProgressStage.RequestingProposal &&
                    item.Message.Contains("still", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task RunAsync_ProjectsAcceptedTransformFieldsIntoTheNextBoundedTurn()
    {
        var client = new FakeClient();
        client.Proposals.Enqueue(new AgentModelProposal
        {
            Model = AgentDefaults.Model,
            ToolCalls =
            {
                new AgentToolCall
                {
                    Id = "transform-call",
                    Name = AgentToolNames.ProposeTransforms,
                    ArgumentsJson =
                        "{\"transforms\":[{\"kind\":\"mapValues\",\"sourceField\":\"Department\"," +
                        "\"outputField\":\"Reporting Group\",\"mappings\":[{\"from\":\"Operations\",\"to\":\"Core\"}]}]}"
                }
            }
        });
        var workflow = WorkflowProposals().ToList();
        workflow[0].ToolCalls[0].ArgumentsJson = workflow[0].ToolCalls[0].ArgumentsJson.Replace(
            "Department",
            "Reporting Group",
            StringComparison.Ordinal);
        foreach (var proposal in workflow) client.Proposals.Enqueue(proposal);
        var job = SyntheticJob.Create();
        var runner = new GuardedAgentRunner(client, hostToolBridge: new FakeHostToolBridge());

        var result = await runner.RunAsync(job, CancellationToken.None);

        Assert.Equal(6, result.ToolCalls.Count);
        Assert.Contains(job.Data.Fields, field => field.Name == "Reporting Group");
        Assert.DoesNotContain(job.Data.Fields, field => field.Name == "Department");
        Assert.Equal("Core", job.Data.SampleRows[0].Values.Single(value =>
            value.Field == "Reporting Group").Value);
    }

    private static void EnqueueWorkflow(FakeClient client)
    {
        foreach (var proposal in WorkflowProposals()) client.Proposals.Enqueue(proposal);
    }

    private static IEnumerable<AgentModelProposal> WorkflowProposals()
    {
        foreach (var toolCall in ToolCallValidatorTests.ApprovedWorkflow())
        {
            yield return new AgentModelProposal
            {
                Model = AgentDefaults.Model,
                ToolCalls = { toolCall },
            };
        }
    }

    private sealed class FakeClient : IOpenAiCompatibleClient
    {
        public Queue<AgentModelProposal> Proposals { get; } = new();

        public bool AlwaysInvalid { get; set; }

        public bool BlockProposalUntilCancellation { get; set; }

        public int ProposalRequests { get; private set; }

        public List<string?> RepairInstructions { get; } = new();

        public Task<ModelDiscoveryResult> DiscoverModelsAsync(
            AgentEndpointSettings settings,
            CancellationToken cancellationToken)
        {
            return Task.FromResult(new ModelDiscoveryResult
            {
                ModelIds = { settings.Model },
                SelectedModel = settings.Model,
            });
        }

        public async Task<AgentModelProposal> RequestToolProposalAsync(
            AgentJobRequest request,
            string? repairInstruction,
            CancellationToken cancellationToken)
        {
            ProposalRequests++;
            RepairInstructions.Add(repairInstruction);
            if (BlockProposalUntilCancellation)
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            }

            if (AlwaysInvalid)
            {
                return new AgentModelProposal
                {
                    Model = request.Endpoint.Model,
                    ToolCalls = { new AgentToolCall { Id = "bad-" + ProposalRequests, Name = "read_file", ArgumentsJson = "{}" } },
                };
            }

            return Proposals.Dequeue();
        }

        public Task<EndpointProbeResult> CheckToolCallingAsync(
            AgentEndpointSettings settings,
            CancellationToken cancellationToken)
        {
            throw new NotSupportedException();
        }
    }

    private sealed class FakeHostToolBridge : IAgentHostToolBridge
    {
        private bool _failedChecks;

        public bool FailFirstChecksCall { get; set; }

        public List<HostToolRequestEvent> Invocations { get; } = new();

        public Task<HostToolResultRequest> InvokeAsync(
            HostToolRequestEvent request,
            CancellationToken cancellationToken)
        {
            Invocations.Add(request);
            if (FailFirstChecksCall && !_failedChecks && request.ToolName == AgentToolNames.RunChecks)
            {
                _failedChecks = true;
                return Task.FromResult(new HostToolResultRequest
                {
                    JobId = request.JobId,
                    ToolCallId = request.ToolCallId,
                    Succeeded = false,
                    OutcomeCode = "checks_failed",
                    ResultJson = "{}",
                    CheckFailures =
                    {
                        new HostCheckFailure
                        {
                            Code = "period_coverage_failed",
                            Message = "One expected synthetic period was missing.",
                        },
                    },
                });
            }

            return Task.FromResult(new HostToolResultRequest
            {
                JobId = request.JobId,
                ToolCallId = request.ToolCallId,
                Succeeded = true,
                OutcomeCode = "accepted",
                ResultJson = "{}",
            });
        }
    }

    private sealed class RecordingSink : IAgentEventSink
    {
        private readonly object _gate = new();

        public List<AgentProgressEvent> Progress { get; } = new();

        public List<AgentCheckpointEvent> Checkpoints { get; } = new();

        public List<long> Sequences { get; } = new();

        public Task PublishProgressAsync(AgentProgressEvent progress, CancellationToken cancellationToken)
        {
            lock (_gate)
            {
                Progress.Add(progress);
                Sequences.Add(progress.Sequence);
            }
            return Task.CompletedTask;
        }

        public Task PublishCheckpointAsync(AgentCheckpointEvent checkpoint, CancellationToken cancellationToken)
        {
            lock (_gate)
            {
                Checkpoints.Add(checkpoint);
                Sequences.Add(checkpoint.Sequence);
            }
            return Task.CompletedTask;
        }
    }
}
