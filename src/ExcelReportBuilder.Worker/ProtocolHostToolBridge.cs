using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Execution;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;

namespace ExcelReportBuilder.Worker;

internal sealed class ProtocolHostToolBridge : IAgentHostToolBridge
{
    private readonly PipeConnectionWriter _writer;
    private readonly ConcurrentDictionary<string, TaskCompletionSource<HostToolResultRequest>> _pending =
        new ConcurrentDictionary<string, TaskCompletionSource<HostToolResultRequest>>(StringComparer.Ordinal);

    public ProtocolHostToolBridge(PipeConnectionWriter writer)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public async Task<HostToolResultRequest> InvokeAsync(
        HostToolRequestEvent request,
        CancellationToken cancellationToken)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        var key = Key(request.JobId, request.ToolCallId);
        var completion = new TaskCompletionSource<HostToolResultRequest>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        if (!_pending.TryAdd(key, completion))
        {
            throw new InvalidOperationException("A deterministic host tool request with this ID is already pending.");
        }

        using var registration = cancellationToken.Register(() => completion.TrySetCanceled(cancellationToken));
        try
        {
            await _writer.SendAsync(
                AgentMessageType.HostToolRequest,
                request.JobId,
                request,
                cancellationToken).ConfigureAwait(false);
            return await completion.Task.ConfigureAwait(false);
        }
        finally
        {
            _pending.TryRemove(key, out _);
        }
    }

    public bool TryComplete(HostToolResultRequest result)
    {
        if (result == null) return false;
        return _pending.TryGetValue(Key(result.JobId, result.ToolCallId), out var completion) &&
               completion.TrySetResult(result);
    }

    public void CancelAll(CancellationToken cancellationToken)
    {
        foreach (var pending in _pending.Values)
        {
            pending.TrySetCanceled(cancellationToken);
        }
    }

    private static string Key(string jobId, string toolCallId)
    {
        return jobId + "\0" + toolCallId;
    }
}
