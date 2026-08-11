using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ExcelLiteralTextTests
{
    [Theory]
    [InlineData("=HYPERLINK(\"https://example.test\")")]
    [InlineData("+SUM(1,1)")]
    [InlineData("-2+3")]
    [InlineData("@SUM(1,1)")]
    [InlineData("  =1+1")]
    [InlineData("\t=1+1")]
    [InlineData("\r=1+1")]
    public void Prefixes_formula_like_untrusted_labels(string value)
    {
        Assert.True(ExcelLiteralText.CouldBeInterpretedAsFormula(value));
        Assert.Equal("'" + value, ExcelLiteralText.Prepare(value));
    }

    [Theory]
    [InlineData("North")]
    [InlineData("1+1")]
    [InlineData("")]
    [InlineData("   ")]
    public void Preserves_inert_labels(string value)
    {
        Assert.False(ExcelLiteralText.CouldBeInterpretedAsFormula(value));
        Assert.Equal(value, ExcelLiteralText.Prepare(value));
    }
}
