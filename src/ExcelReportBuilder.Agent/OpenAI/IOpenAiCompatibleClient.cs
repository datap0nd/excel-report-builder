using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.OpenAI;

public interface IOpenAiCompatibleClient
{
    Task<ModelDiscoveryResult> DiscoverModelsAsync(
        AgentEndpointSettings settings,
        CancellationToken cancellationToken);

    Task<AgentModelProposal> RequestToolProposalAsync(
        AgentJobRequest request,
        string? repairInstruction,
        CancellationToken cancellationToken);

    Task<EndpointProbeResult> CheckToolCallingAsync(
        AgentEndpointSettings settings,
        CancellationToken cancellationToken);
}
