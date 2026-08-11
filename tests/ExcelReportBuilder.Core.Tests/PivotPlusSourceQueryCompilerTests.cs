using System;
using ExcelReportBuilder.Core.PivotPlus;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotPlusSourceQueryCompilerTests
{
    [Fact]
    public void TableQueryReadsOnlyTheNamedWorkbookObject()
    {
        string query = PivotPlusSourceQueryCompiler.Compile(
            "Sales_Data",
            PivotPlusWorkbookObjectKind.Table);

        Assert.Equal(
            "let\n    Source = Excel.CurrentWorkbook(){[Name=\"Sales_Data\"]}[Content]\nin\n    Source",
            query);
        Assert.DoesNotContain("File.", query, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Web.", query, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void NamedRangePromotesExactlyOneHeaderRow()
    {
        string query = PivotPlusSourceQueryCompiler.Compile(
            "Selected_Source",
            PivotPlusWorkbookObjectKind.NamedRange);

        Assert.Contains("Excel.CurrentWorkbook()", query, StringComparison.Ordinal);
        Assert.Contains("Table.PromoteHeaders", query, StringComparison.Ordinal);
        Assert.Contains("Culture = \"en-US\"", query, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Sheet 1!A1:B2")]
    [InlineData("Source\"]}[Content], Web.Contents(\"https://invalid\")")]
    [InlineData("a-b")]
    [InlineData(@"folder\source")]
    [InlineData("file:source")]
    [InlineData("C:source")]
    public void RejectsNamesOutsideTheWorkbookObjectGrammar(string name)
    {
        Assert.Throws<ArgumentException>(() =>
            PivotPlusSourceQueryCompiler.Compile(
                name,
                PivotPlusWorkbookObjectKind.Table));
    }

    [Fact]
    public void RejectsNetworkShareShapeWithoutEmbeddingARepositoryPath()
    {
        string name = new string('\\', 2) + "host" + '\\' + "object";

        Assert.Throws<ArgumentException>(() =>
            PivotPlusSourceQueryCompiler.Compile(
                name,
                PivotPlusWorkbookObjectKind.Table));
    }

    [Fact]
    public void RejectsUnknownSourceKind()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PivotPlusSourceQueryCompiler.Compile(
                "Sales_Data",
                (PivotPlusWorkbookObjectKind)99));
    }
}
