using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ManagedOutputLayoutPlannerTests
{
    [Fact]
    public void Groups_blocks_on_one_logical_sheet_and_separates_other_outputs()
    {
        var blocks = new List<DenseReportBlockPlan>
        {
            Block("summary", "Report", "B3"),
            Block("detail", " report ", "J3"),
            Block("appendix", "Appendix", "A1")
        };

        var outputs = ManagedOutputLayoutPlanner.Group("report-id", blocks);

        Assert.Equal(2, outputs.Count);
        Assert.Equal(2, outputs.Single(output => output.Blocks.Count == 2).Blocks.Count);
        Assert.All(outputs, output =>
            Assert.Equal(ManagedObjectKind.DraftWorksheet, output.DraftIdentity.Kind));
        Assert.NotEqual(outputs[0].DraftIdentity.ObjectId, outputs[1].DraftIdentity.ObjectId);
    }

    [Fact]
    public void Derives_idempotent_case_insensitive_identities_for_each_output_lifecycle()
    {
        var first = ManagedOutputIdentity.Draft("report-id", "Report");
        var second = ManagedOutputIdentity.Draft("report-id", " report ");
        var published = ManagedOutputIdentity.Published("report-id", "REPORT");
        var rollback = ManagedOutputIdentity.Rollback("report-id", "Report");

        Assert.Equal(first.ObjectId, second.ObjectId);
        Assert.Equal(first.ObjectId, published.ObjectId);
        Assert.Equal(first.ObjectId, rollback.ObjectId);
        Assert.NotEqual(first.MarkerValue, published.MarkerValue);
        Assert.NotEqual(published.MarkerValue, rollback.MarkerValue);
    }

    private static DenseReportBlockPlan Block(string id, string worksheet, string anchor)
    {
        return new DenseReportBlockPlan
        {
            BlockId = id,
            OwnershipId = "owned_" + id,
            WorksheetName = worksheet,
            AnchorCell = anchor
        };
    }
}
