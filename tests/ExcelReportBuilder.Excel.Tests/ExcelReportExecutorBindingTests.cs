using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ExcelReportExecutorBindingTests
{
    [Fact]
    public void Rejects_a_stale_plan_before_accessing_excel()
    {
        var specification = new ReportSpecV1
        {
            Id = "report",
            OwnershipId = "owned_report",
            Name = "Report"
        };
        var plan = new ReportBuildPlan
        {
            SpecificationId = specification.Id,
            OwnershipId = specification.OwnershipId,
            SchemaVersion = specification.SchemaVersion,
            SpecificationHash = "stale"
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new ExcelReportExecutor().BuildManagedDraft(new object(), specification, plan));

        Assert.Contains("exact validated specification", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_a_plan_mutated_after_planning_before_accessing_excel()
    {
        var specification = new ReportSpecV1
        {
            Id = "report",
            OwnershipId = "owned_report",
            Name = "Report"
        };
        var plan = new ReportBuildPlan
        {
            SpecificationId = specification.Id,
            OwnershipId = specification.OwnershipId,
            SchemaVersion = specification.SchemaVersion,
            SpecificationHash = ReportSpecDigest.Compute(specification)
        };
        plan.PlanHash = ReportBuildPlanDigest.Compute(plan);
        plan.Source.ProjectedRows = 12;

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new ExcelReportExecutor().BuildManagedDraft(new object(), specification, plan));

        Assert.Contains("exact validated specification", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
