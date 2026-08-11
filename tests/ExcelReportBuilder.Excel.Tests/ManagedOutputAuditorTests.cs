using ExcelReportBuilder.Excel.Rendering;
using ExcelReportBuilder.Excel.Validation;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ManagedOutputAuditorTests
{
    [Fact]
    public void Repeated_measure_ids_are_reconciled_per_block_even_when_global_totals_cancel()
    {
        var first = new ManagedBlockMeasureKey("first", "amount");
        var second = new ManagedBlockMeasureKey("second", "amount");
        var pivotTotals = new Dictionary<ManagedBlockMeasureKey, decimal>
        {
            [first] = 100m,
            [second] = 50m
        };
        var outputTotals = new Dictionary<ManagedBlockMeasureKey, decimal>
        {
            [first] = 99m,
            [second] = 51m
        };

        var checks = ManagedOutputAuditor.ReconcileBlockTotals(
            new[] { first, second },
            pivotTotals,
            outputTotals,
            0m);

        Assert.Equal(2, checks.Count);
        Assert.All(checks, check => Assert.Equal(CheckOutcome.Failed, check.Outcome));
        Assert.Equal(
            ManagedOutputAuditor.AggregateByMeasure(pivotTotals)["amount"],
            ManagedOutputAuditor.AggregateByMeasure(outputTotals)["amount"]);
    }

    [Fact]
    public void Missing_output_total_fails_only_its_owned_block_check()
    {
        var first = new ManagedBlockMeasureKey("first", "amount");
        var second = new ManagedBlockMeasureKey("second", "amount");

        var checks = ManagedOutputAuditor.ReconcileBlockTotals(
            new[] { first, second },
            new Dictionary<ManagedBlockMeasureKey, decimal>
            {
                [first] = 100m,
                [second] = 20m
            },
            new Dictionary<ManagedBlockMeasureKey, decimal>
            {
                [first] = 100m
            },
            0m);

        Assert.Equal(CheckOutcome.Passed, checks.Single(check => check.CheckId.Contains("first")).Outcome);
        Assert.Equal(CheckOutcome.Failed, checks.Single(check => check.CheckId.Contains("second")).Outcome);
    }

    [Theory]
    [InlineData(10.0, false, true)]
    [InlineData(10.0000005, false, true)]
    [InlineData(10.01, false, false)]
    [InlineData(null, false, false)]
    [InlineData(10.0, true, false)]
    public void Numeric_formula_expectations_are_compared_with_bounded_tolerance(
        double? actual,
        bool hasExcelError,
        bool expected)
    {
        Assert.Equal(
            expected,
            ManagedOutputAuditor.FormulaValueMatches(
                DenseFormulaExpectation.Number(10m),
                actual,
                hasExcelError));
    }

    [Fact]
    public void Blank_and_error_formula_expectations_are_distinct()
    {
        Assert.True(ManagedOutputAuditor.FormulaValueMatches(
            DenseFormulaExpectation.Blank(),
            string.Empty,
            false));
        Assert.False(ManagedOutputAuditor.FormulaValueMatches(
            DenseFormulaExpectation.Blank(),
            null,
            true));
        Assert.True(ManagedOutputAuditor.FormulaValueMatches(
            DenseFormulaExpectation.Error(),
            2042,
            true));
        Assert.False(ManagedOutputAuditor.FormulaValueMatches(
            DenseFormulaExpectation.Error(),
            10m,
            false));
        Assert.False(ManagedOutputAuditor.IsExcelErrorDisplay("########"));
        Assert.True(ManagedOutputAuditor.IsExcelErrorDisplay("#DIV/0!"));
    }

    [Fact]
    public void Unchanged_formula_with_changed_numeric_value_fails_the_independent_audit()
    {
        var formula = SafeFormulaFactory.FromTypedMeasure("=1+1");

        var result = ManagedOutputAuditor.AuditFormulaCell(
            formula,
            DenseFormulaExpectation.Number(2m),
            "=1+1",
            "3",
            3d);

        Assert.False(result.FormulaChanged);
        Assert.False(result.HasExcelError);
        Assert.False(result.ValueMatches);
    }
}
