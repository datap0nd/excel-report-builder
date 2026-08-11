using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;

namespace ExcelReportBuilder.Worker;

internal sealed class PipeConnectionWriter : IAgentEventSink, IDisposable
{
    private readonly Stream _stream;
    private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);
    private readonly CancellationToken _connectionCancellation;

    public PipeConnectionWriter(Stream stream, CancellationToken connectionCancellation)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _connectionCancellation = connectionCancellation;
    }

    public Task PublishProgressAsync(AgentProgressEvent progress, CancellationToken cancellationToken)
    {
        return SendAsync(
            AgentMessageType.Progress,
            progress.JobId,
            progress,
            cancellationToken);
    }

    public Task PublishCheckpointAsync(AgentCheckpointEvent checkpoint, CancellationToken cancellationToken)
    {
        return SendAsync(
            AgentMessageType.Checkpoint,
            checkpoint.JobId,
            checkpoint,
            cancellationToken);
    }

    public async Task SendAsync<TPayload>(
        AgentMessageType messageType,
        string correlationId,
        TPayload payload,
        CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            _connectionCancellation);
        await _gate.WaitAsync(linked.Token).ConfigureAwait(false);
        try
        {
            var envelope = AgentProtocol.Create(messageType, correlationId, payload);
            await PipeJsonProtocol.WriteAsync(_stream, envelope, linked.Token).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public void Dispose()
    {
        _gate.Dispose();
    }
}
