using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class FormulaBuilderTests
{
    [Fact]
    public void GetPivotData_formula_escapes_all_text_as_literals()
    {
        var builder = new GetPivotDataFormulaBuilder();

        var formula = builder.Build(
            "Net \"value\"",
            "Managed ' pivot",
            "$A$3",
            new[]
            {
                new PivotFilterItem { Field = "Region", Value = "A\"B" },
                new PivotFilterItem { Field = "Period", Value = new DateTime(2026, 1, 1) }
            });

        Assert.Equal(
            "=IFERROR(GETPIVOTDATA(\"Net \"\"value\"\"\",'Managed '' pivot'!$A$3,\"Region\",\"A\"\"B\",\"Period\",DATE(2026,1,1)),\"\")",
            formula);
    }

    [Fact]
    public void Formula_builder_rejects_formula_like_anchor()
    {
        var builder = new GetPivotDataFormulaBuilder();

        Assert.Throws<ArgumentException>(() => builder.Build("Value", "Pivot", "A1+NOW()"));
    }

    [Fact]
    public void Safe_divide_returns_blank_for_zero_or_missing_denominator()
    {
        var builder = new GetPivotDataFormulaBuilder();

        Assert.Equal("=IF(OR(B2=0,B2=\"\"),\"\",A2/B2)", builder.SafeDivide("A2", "B2"));
    }

    [Fact]
    public void GetPivotData_preserves_datetime_filter_time()
    {
        var formula = new GetPivotDataFormulaBuilder().Build(
            "Value",
            "Managed pivot",
            "$A$3",
            new[]
            {
                new PivotFilterItem
                {
                    Field = "Timestamp",
                    Value = new DateTime(2026, 1, 2, 13, 14, 15, 250)
                }
            });

        Assert.Contains("DATE(2026,1,2)+TIME(13,14,15.25)", formula, StringComparison.Ordinal);
    }
}
