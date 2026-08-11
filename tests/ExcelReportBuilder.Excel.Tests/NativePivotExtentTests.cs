using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class NativePivotExtentTests
{
    [Fact]
    public void Rejects_a_native_pivot_larger_than_its_owned_extent()
    {
        var block = new DenseReportBlockPlan
        {
            OwnedRange = new OwnedRangePlan { RowCount = 10, ColumnCount = 5 }
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            NativePivotTableExecutor.DemandDimensionsWithinOwnedRange(block, 11, 5));

        Assert.Contains("owned extent", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Accepts_a_native_pivot_inside_its_owned_extent()
    {
        var block = new DenseReportBlockPlan
        {
            OwnedRange = new OwnedRangePlan { RowCount = 10, ColumnCount = 5 }
        };

        NativePivotTableExecutor.DemandDimensionsWithinOwnedRange(block, 10, 5);
    }

    [Fact]
    public void Lets_a_hidden_dense_support_pivot_exceed_the_visible_block_extent()
    {
        var block = new DenseReportBlockPlan
        {
            OutputMode = ReportOutputMode.DenseGrid,
            OwnedRange = new OwnedRangePlan { RowCount = 10, ColumnCount = 5 }
        };

        NativePivotTableExecutor.DemandDimensionsWithinOwnedRange(block, 1000, 20);
    }
}
