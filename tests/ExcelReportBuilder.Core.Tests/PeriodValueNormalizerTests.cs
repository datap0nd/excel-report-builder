using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PeriodValueNormalizerTests
{
    [Theory]
    [InlineData("202601", 2026, 1)]
    [InlineData("2026-02", 2026, 2)]
    [InlineData("Mar-26", 2026, 3)]
    [InlineData("April 2026", 2026, 4)]
    [InlineData("2026 May", 2026, 5)]
    public void Normalizes_supported_month_values(string value, int year, int month)
    {
        DateTime result = PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Month);

        Assert.Equal(new DateTime(year, month, 1), result);
    }

    [Theory]
    [InlineData(202601)]
    [InlineData(202601d)]
    public void Normalizes_compact_numeric_month_values(object value)
    {
        DateTime result = PeriodValueNormalizer.Normalize(
            value,
            expectedGrain: PeriodGrain.Month);

        Assert.Equal(new DateTime(2026, 1, 1), result);
    }

    [Theory]
    [InlineData("Jan--26")]
    [InlineData("2026//01")]
    [InlineData("Q1__26")]
    [InlineData("Jan - 26")]
    public void Rejects_malformed_repeated_period_delimiters(string value)
    {
        Assert.Throws<ArgumentException>(() =>
            PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Month));
    }

    [Fact]
    public void Repeated_whitespace_remains_a_supported_period_separator()
    {
        DateTime result = PeriodValueNormalizer.Normalize(
            "Jan   26",
            expectedGrain: PeriodGrain.Month);

        Assert.Equal(new DateTime(2026, 1, 1), result);
    }

    [Theory]
    [InlineData(202600)]
    [InlineData(202613)]
    public void Rejects_invalid_six_digit_numeric_periods(int value)
    {
        Assert.Throws<ArgumentException>(() =>
            PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Month));
    }

    [Theory]
    [InlineData("Q1 2026", 1)]
    [InlineData("2026-Q2", 4)]
    [InlineData("Q3-26", 7)]
    [InlineData("26 Q4", 10)]
    public void Normalizes_supported_quarter_values(string value, int firstMonth)
    {
        DateTime result = PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Quarter);

        Assert.Equal(new DateTime(2026, firstMonth, 1), result);
    }

    [Fact]
    public void Uses_an_explicit_year_for_yearless_months_and_quarters()
    {
        Assert.Equal(
            new DateTime(2026, 1, 1),
            PeriodValueNormalizer.Normalize("Jan", 2026, PeriodGrain.Month));
        Assert.Equal(
            new DateTime(2026, 4, 1),
            PeriodValueNormalizer.Normalize("Q2", 2026, PeriodGrain.Quarter));
    }

    [Fact]
    public void Never_infers_a_missing_year()
    {
        Assert.Throws<InvalidOperationException>(() =>
            PeriodValueNormalizer.Normalize("January", expectedGrain: PeriodGrain.Month));
        Assert.Throws<InvalidOperationException>(() =>
            PeriodValueNormalizer.Normalize("Q1", expectedGrain: PeriodGrain.Quarter));
    }

    [Fact]
    public void Rejects_ambiguous_dates_and_grain_mismatches()
    {
        Assert.Throws<ArgumentException>(() =>
            PeriodValueNormalizer.Normalize("01/02/2026"));
        Assert.Throws<InvalidOperationException>(() =>
            PeriodValueNormalizer.Normalize("Jan 2026", expectedGrain: PeriodGrain.Quarter));
    }

    [Fact]
    public void Canonicalizes_excel_dates_to_the_declared_grain()
    {
        var value = new DateTime(2026, 5, 17, 14, 30, 0);

        Assert.Equal(
            new DateTime(2026, 5, 17),
            PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Day));
        Assert.Equal(
            new DateTime(2026, 5, 1),
            PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Month));
        Assert.Equal(
            new DateTime(2026, 4, 1),
            PeriodValueNormalizer.Normalize(value, expectedGrain: PeriodGrain.Quarter));
    }
}
