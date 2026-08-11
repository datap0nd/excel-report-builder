using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Execution;

public interface IAgentHostToolBridge
{
    Task<HostToolResultRequest> InvokeAsync(
        HostToolRequestEvent request,
        CancellationToken cancellationToken);
}

public sealed class UnavailableAgentHostToolBridge : IAgentHostToolBridge
{
    public static UnavailableAgentHostToolBridge Instance { get; } = new UnavailableAgentHostToolBridge();

    private UnavailableAgentHostToolBridge()
    {
    }

    public Task<HostToolResultRequest> InvokeAsync(
        HostToolRequestEvent request,
        CancellationToken cancellationToken)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        return Task.FromResult(new HostToolResultRequest
        {
            JobId = request.JobId,
            ToolCallId = request.ToolCallId,
            Succeeded = false,
            OutcomeCode = "host_tool_bridge_unavailable",
            ResultJson = "{}",
        });
    }
}
