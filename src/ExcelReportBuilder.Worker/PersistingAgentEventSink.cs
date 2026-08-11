using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Worker;

internal sealed class PersistingAgentEventSink : IAgentEventSink
{
    private readonly PipeConnectionWriter _writer;
    private readonly IJobCheckpointStore _checkpointStore;

    public PersistingAgentEventSink(
        PipeConnectionWriter writer,
        IJobCheckpointStore checkpointStore)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
        _checkpointStore = checkpointStore ?? throw new ArgumentNullException(nameof(checkpointStore));
    }

    public Task PublishProgressAsync(AgentProgressEvent progress, CancellationToken cancellationToken)
    {
        return _writer.PublishProgressAsync(progress, cancellationToken);
    }

    public async Task PublishCheckpointAsync(
        AgentCheckpointEvent checkpoint,
        CancellationToken cancellationToken)
    {
        await _checkpointStore.SaveCheckpointAsync(checkpoint, cancellationToken).ConfigureAwait(false);
        await _writer.PublishCheckpointAsync(checkpoint, cancellationToken).ConfigureAwait(false);
    }
}
