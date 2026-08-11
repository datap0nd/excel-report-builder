using System.Text.Json;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Worker;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class WorkerStateTests
{
    [Fact]
    public void ActiveWorkbookRegistry_AllowsOnlyOneJobPerWorkbook()
    {
        var registry = new ActiveWorkbookJobRegistry();

        Assert.True(registry.TryStart("workbook-1", "job-1"));
        Assert.False(registry.TryStart("workbook-1", "job-2"));
        Assert.True(registry.TryStart("workbook-2", "job-3"));
        registry.Complete("workbook-1", "different-job");
        Assert.False(registry.TryStart("workbook-1", "job-4"));
        registry.Complete("workbook-1", "job-1");
        Assert.True(registry.TryStart("workbook-1", "job-4"));
    }

    [Fact]
    public async Task CheckpointStore_PersistsOnlyResumeMetadataAndExposesFailure()
    {
        var directory = Path.Combine(Path.GetTempPath(), "erb-agent-checkpoint-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new LocalAppDataJobCheckpointStore(directory);
            var checkpoint = new AgentCheckpointEvent
            {
                JobId = "job-synthetic-1",
                WorkbookId = "workbook-synthetic-1",
                CheckpointId = "checkpoint-synthetic-1",
                Stage = AgentProgressStage.ProcessingHostResult,
                CompletedRepairCycles = 1,
                LastCompletedStep = "Processed deterministic host check result.",
                OccurredAtUtc = DateTimeOffset.Parse("2026-01-01T00:00:00Z"),
            };

            await store.SaveCheckpointAsync(checkpoint, CancellationToken.None);
            var failure = await store.MarkFailureAsync(
                checkpoint.JobId,
                checkpoint.WorkbookId,
                "worker_interrupted",
                CancellationToken.None);

            Assert.True(failure.CanResume);
            Assert.Equal(checkpoint.CheckpointId, failure.CheckpointId);
            Assert.Equal("worker_interrupted", failure.FailureCode);
            var listed = await store.ListAsync(checkpoint.WorkbookId, CancellationToken.None);
            Assert.Single(listed);

            var path = Assert.Single(Directory.GetFiles(directory, "*.json"));
            using var document = JsonDocument.Parse(await File.ReadAllTextAsync(path));
            var propertyNames = document.RootElement.EnumerateObject().Select(item => item.Name).ToArray();
            Assert.DoesNotContain("userPrompt", propertyNames);
            Assert.DoesNotContain("endpoint", propertyNames);
            Assert.DoesNotContain("apiKey", propertyNames);
            Assert.DoesNotContain("data", propertyNames);
            Assert.DoesNotContain("toolArguments", propertyNames);
            Assert.DoesNotContain("toolResult", propertyNames);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }

    [Fact]
    public async Task CheckpointStore_DeletesSuccessfulJobMetadata()
    {
        var directory = Path.Combine(Path.GetTempPath(), "erb-agent-checkpoint-" + Guid.NewGuid().ToString("N"));
        try
        {
            var store = new LocalAppDataJobCheckpointStore(directory);
            var checkpoint = new AgentCheckpointEvent
            {
                JobId = "job-synthetic-2",
                WorkbookId = "workbook-synthetic-2",
                CheckpointId = "checkpoint-synthetic-2",
                Stage = AgentProgressStage.ValidatingProposal,
                LastCompletedStep = "Validated synthetic proposal.",
                OccurredAtUtc = DateTimeOffset.UtcNow,
            };
            await store.SaveCheckpointAsync(checkpoint, CancellationToken.None);

            await store.DeleteAsync(checkpoint.JobId, checkpoint.WorkbookId, CancellationToken.None);

            Assert.Null(await store.LoadAsync(checkpoint.JobId, checkpoint.WorkbookId, CancellationToken.None));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, true);
        }
    }
}
