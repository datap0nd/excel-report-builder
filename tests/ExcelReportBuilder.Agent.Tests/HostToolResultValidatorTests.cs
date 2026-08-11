using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class HostToolResultValidatorTests
{
    [Fact]
    public void Validate_AcceptsBoundedDeterministicResult()
    {
        var request = Request();
        var result = new HostToolResultRequest
        {
            JobId = request.JobId,
            ToolCallId = request.ToolCallId,
            Succeeded = false,
            OutcomeCode = "checks_failed",
            ResultJson = "{\"passed\":false}",
            CheckFailures =
            {
                new HostCheckFailure { Code = "totals_mismatch", Message = "Synthetic totals did not match." },
            },
        };

        HostToolResultValidator.Validate(request, result);
    }

    [Fact]
    public void Validate_RejectsMismatchedPendingRequest()
    {
        var request = Request();
        var result = new HostToolResultRequest
        {
            JobId = request.JobId,
            ToolCallId = "different-call",
            OutcomeCode = "accepted",
            ResultJson = "{}",
        };

        var error = Assert.Throws<AgentInputValidationException>(() =>
            HostToolResultValidator.Validate(request, result));

        Assert.Equal("host_tool_result_mismatch", error.Code);
    }

    [Fact]
    public void Validate_RejectsUnboundedOrUnstructuredResult()
    {
        var request = Request();
        var result = new HostToolResultRequest
        {
            JobId = request.JobId,
            ToolCallId = request.ToolCallId,
            OutcomeCode = "accepted",
            ResultJson = "[]",
        };

        Assert.Throws<AgentInputValidationException>(() =>
            HostToolResultValidator.Validate(request, result));
    }

    private static HostToolRequestEvent Request()
    {
        return new HostToolRequestEvent
        {
            JobId = "job-1",
            WorkbookId = "workbook-1",
            ToolCallId = "tool-call-1",
            ToolName = "run_checks",
            ArgumentsJson = "{}",
        };
    }
}
