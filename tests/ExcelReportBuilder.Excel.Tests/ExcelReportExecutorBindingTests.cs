using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Rendering;
using System.Runtime.InteropServices;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ExcelReportExecutorBindingTests
{
    [Theory]
    [InlineData(false, false, false, true, 0, false, false, false, true)]
    [InlineData(false, true, false, true, 0, false, false, false, false)]
    [InlineData(false, false, true, true, 0, false, false, false, false)]
    [InlineData(false, false, false, false, 0, false, false, false, false)]
    [InlineData(false, false, false, true, 1, true, false, true, true)]
    [InlineData(false, false, false, true, 1, true, false, false, false)]
    [InlineData(false, false, false, true, 1, false, false, false, true)]
    [InlineData(false, false, false, true, 1, false, true, false, false)]
    [InlineData(true, false, false, true, 0, false, false, false, false)]
    public void Hidden_row_grand_totals_use_only_additive_leaf_output_cells_for_reconciliation(
        bool showRowGrandTotals,
        bool rowIsSubtotal,
        bool outputIsSliced,
        bool isAggregateMeasure,
        int columnFieldCount,
        bool hasExplicitGrandColumn,
        bool columnIsSubtotal,
        bool columnIsGrandTotal,
        bool expected)
    {
        Assert.Equal(
            expected,
            ExcelReportExecutor.IsDenseOutputTotalContribution(
                showRowGrandTotals,
                rowIsSubtotal,
                outputIsSliced,
                isAggregateMeasure,
                columnFieldCount,
                hasExplicitGrandColumn,
                columnIsSubtotal,
                columnIsGrandTotal));
    }

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

    [Fact]
    public void Com_failure_during_independent_pivot_read_fails_closed()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            ExcelReportExecutor.ReadPivotAggregateExpectation(
                new ThrowingPivot(),
                "Amount",
                Array.Empty<PivotFilterItem>()));

        Assert.Contains("independently", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unsupported_independent_pivot_filter_depth_fails_closed()
    {
        var filters = Enumerable.Range(
                1,
                ExcelReportExecutor.MaximumIndependentPivotFilterPairs + 1)
            .Select(index => new PivotFilterItem
            {
                Field = "Field" + index,
                Value = index
            })
            .ToList();

        var exception = Assert.Throws<NotSupportedException>(() =>
            ExcelReportExecutor.ReadPivotAggregateExpectation(
                new ThrowingPivot(),
                "Amount",
                filters));

        Assert.Contains("fourteen", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    public sealed class ThrowingPivot
    {
        public object GetPivotData(string dataFieldCaption)
        {
            throw new COMException("Synthetic PivotTable read failure.");
        }
    }
}
