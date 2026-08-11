using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class CanonicalRoutingTests
{
    [Theory]
    [InlineData(0, CanonicalBackend.Worksheet)]
    [InlineData(1_048_575, CanonicalBackend.Worksheet)]
    [InlineData(1_048_576, CanonicalBackend.DataModel)]
    [InlineData(12_000_000, CanonicalBackend.DataModel)]
    public void Router_never_sends_oversized_results_to_a_worksheet(
        long projectedRows,
        CanonicalBackend expected)
    {
        var router = new CanonicalDestinationRouter();

        Assert.Equal(expected, router.Choose(projectedRows));
    }

    [Fact]
    public void Restricted_query_policy_accepts_current_workbook_only()
    {
        var policy = new RestrictedQueryFormulaPolicy();

        policy.DemandWorkbookOnly(
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source");
    }

    [Theory]
    [InlineData("let Source = File.Contents(\"data\") in Source")]
    [InlineData("let Source = Web.Contents(\"https://example.test\") in Source")]
    [InlineData("let Source = Sql.Database(\"server\", \"db\") in Source")]
    public void Restricted_query_policy_rejects_external_access(string formula)
    {
        var policy = new RestrictedQueryFormulaPolicy();

        Assert.Throws<InvalidOperationException>(() => policy.DemandWorkbookOnly(formula));
    }
}
