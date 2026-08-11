using ExcelReportBuilder.Agent.Execution;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class AgentJobIdentityTests
{
    [Fact]
    public void Create_IsStableForTheExactBoundedRequest()
    {
        var first = SyntheticJob.Create();
        var second = SyntheticJob.Create();

        Assert.Equal(AgentJobIdentity.Create(first), AgentJobIdentity.Create(second));
        Assert.StartsWith("job_", AgentJobIdentity.Create(first), StringComparison.Ordinal);
    }

    [Fact]
    public void Create_ChangesWhenPromptDataOrWorkbookChanges()
    {
        var baseline = SyntheticJob.Create();
        string expected = AgentJobIdentity.Create(baseline);

        var promptChanged = SyntheticJob.Create();
        promptChanged.UserPrompt += " Add a filter.";
        var dataChanged = SyntheticJob.Create();
        dataChanged.Data.SampleRows[0].Values[0].Value = "Different";
        var workbookChanged = SyntheticJob.Create();
        workbookChanged.WorkbookId = "workbook-synthetic-002";

        Assert.NotEqual(expected, AgentJobIdentity.Create(promptChanged));
        Assert.NotEqual(expected, AgentJobIdentity.Create(dataChanged));
        Assert.NotEqual(expected, AgentJobIdentity.Create(workbookChanged));
    }

    [Fact]
    public void Create_DoesNotExposeWorkbookOrPromptText()
    {
        var request = SyntheticJob.Create();

        string identity = AgentJobIdentity.Create(request);

        Assert.DoesNotContain(request.WorkbookId, identity, StringComparison.Ordinal);
        Assert.DoesNotContain(request.UserPrompt, identity, StringComparison.Ordinal);
    }
}
