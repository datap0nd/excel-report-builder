using ExcelReportBuilder.Excel.PivotPlus.NamedSets;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotNamedSetFormulaTransportTests
{
    [Fact]
    public void Encodes_one_exact_excel_apostrophe_envelope()
    {
        const string raw = "{[Product].[SKU].&[Director's Cut]}";

        string encoded = PivotNamedSetFormulaTransport.EncodeForExcel(raw);

        Assert.Equal("'{[Product].[SKU].&[Director''s Cut]}'", encoded);
    }

    [Theory]
    [InlineData("{[Sales].[Region].DefaultMember}", "{[Sales].[Region].DefaultMember}")]
    [InlineData("'{[Sales].[Region].DefaultMember}'", "{[Sales].[Region].DefaultMember}")]
    [InlineData("'{[Product].[SKU].&[Director''s Cut]}'", "{[Product].[SKU].&[Director's Cut]}")]
    public void Decodes_only_raw_or_one_reversible_envelope(
        string readback,
        string expected)
    {
        Assert.True(PivotNamedSetFormulaTransport.TryDecodeReadback(
            readback,
            out string raw));
        Assert.Equal(expected, raw);
    }

    [Theory]
    [InlineData("'{[Sales].[Region]}")]
    [InlineData("{[Sales].[Region]}'")]
    [InlineData("'{[Product].[SKU].&[Director's Cut]}'")]
    [InlineData("={[Sales].[Region]}")]
    [InlineData("")]
    public void Rejects_malformed_or_unsupported_readback(string value)
    {
        Assert.False(PivotNamedSetFormulaTransport.TryDecodeReadback(
            value,
            out _));
    }

    [Fact]
    public void Exact_readback_does_not_normalize_case_or_whitespace()
    {
        const string expected = "{[Sales].[Region].DefaultMember}";

        Assert.Throws<InvalidOperationException>(() =>
            PivotNamedSetFormulaTransport.DemandExactReadback(
                "{ [Sales].[Region].DefaultMember }",
                expected));
        Assert.Throws<InvalidOperationException>(() =>
            PivotNamedSetFormulaTransport.DemandExactReadback(
                "{[sales].[Region].DefaultMember}",
                expected));
    }
}
