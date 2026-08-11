using ExcelReportBuilder.Agent.Execution;
using System.Reflection;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class WorkbookAgentJobLeaseTests
{
    [Fact]
    public void TryAcquire_AllowsOnlyOneLeaseForTheSameWorkbook()
    {
        string workbookId = "workbook_" + Guid.NewGuid().ToString("N");

        Assert.True(WorkbookAgentJobLease.TryAcquire(workbookId, out var first));
        try
        {
            Assert.False(WorkbookAgentJobLease.TryAcquire(workbookId, out var second));
            Assert.Null(second);
        }
        finally
        {
            first!.Dispose();
        }

        Assert.True(WorkbookAgentJobLease.TryAcquire(workbookId, out var afterRelease));
        afterRelease!.Dispose();
    }

    [Fact]
    public void LeaseName_DoesNotExposeWorkbookOrUserText()
    {
        const string workbookId = "workbook-sensitive-display-name";

        MethodInfo method = typeof(WorkbookAgentJobLease).GetMethod(
            "CreateLeaseName",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        string leaseName = (string)method.Invoke(null, new object[] { workbookId })!;

        Assert.DoesNotContain(workbookId, leaseName, StringComparison.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(Environment.UserName))
        {
            Assert.DoesNotContain(Environment.UserName, leaseName, StringComparison.OrdinalIgnoreCase);
        }
    }
}
