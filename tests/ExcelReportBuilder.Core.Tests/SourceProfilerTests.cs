using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

public sealed class SourceProfilerTests
{
    [Fact]
    public void Profiles_synthetic_types_without_parsing_ambiguous_dates()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Units", "Amount", "Ambiguous" },
            new object?[][]
            {
                new object?[] { "2026-01-01", 2, 12.5m, "01/02/2026" },
                new object?[] { "2026-02-01", 3L, 20m, "02/03/2026" },
                new object?[] { null, 4, null, null }
            });

        Assert.Equal(3, profile.RowCount);
        Assert.Equal(SourceValueType.Date, profile.FindColumn("Period")!.InferredType);
        Assert.Equal(1, profile.FindColumn("Period")!.BlankCount);
        Assert.Equal(SourceValueType.WholeNumber, profile.FindColumn("Units")!.InferredType);
        Assert.Equal(SourceValueType.DecimalNumber, profile.FindColumn("Amount")!.InferredType);
        Assert.Equal(SourceValueType.Text, profile.FindColumn("Ambiguous")!.InferredType);
    }

    [Fact]
    public void Records_duplicate_headers_and_ragged_rows_instead_of_hiding_them()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "region" },
            new object?[][] { new object?[] { "North" } });

        Assert.Contains(profile.Issues, issue => issue.Code == SourceProfileIssueCode.DuplicateHeader);
        Assert.Contains(profile.Issues, issue => issue.Code == SourceProfileIssueCode.RaggedRow);
    }

    [Fact]
    public void Profiles_compact_month_year_and_quarter_period_text_without_culture_guessing()
    {
        var monthly = SourceProfiler.Profile(
            new[] { "Period" },
            new object?[][]
            {
                new object?[] { 202601 },
                new object?[] { "Feb-26" },
                new object?[] { "March 2026" }
            });
        var quarterly = SourceProfiler.Profile(
            new[] { "Period" },
            new object?[][]
            {
                new object?[] { "Q1 2026" },
                new object?[] { "2026-Q2" }
            });

        SourceColumnProfile monthColumn = monthly.FindColumn("Period")!;
        SourceColumnProfile quarterColumn = quarterly.FindColumn("Period")!;
        Assert.Equal(SourceValueType.Date, monthColumn.InferredType);
        Assert.Equal(3, monthColumn.DateLikeCount);
        Assert.Equal(3, monthColumn.MonthGrainCount);
        Assert.Equal(new DateTime(2026, 1, 1), monthColumn.MinimumDate);
        Assert.Equal(new DateTime(2026, 3, 1), monthColumn.MaximumDate);
        Assert.Equal(2, quarterColumn.QuarterGrainCount);
        Assert.Equal(new DateTime(2026, 4, 1), quarterColumn.MaximumDate);
    }

    [Fact]
    public void Uses_a_fixed_two_digit_year_window_and_keeps_yearless_periods_unresolved()
    {
        var resolved = SourceProfiler.Profile(
            new[] { "Period" },
            new object?[][]
            {
                new object?[] { "Jan-29" },
                new object?[] { "Feb-30" }
            });
        var unresolved = SourceProfiler.Profile(
            new[] { "Period" },
            new object?[][]
            {
                new object?[] { "Jan" },
                new object?[] { "Q2" }
            });

        SourceColumnProfile resolvedColumn = resolved.FindColumn("Period")!;
        SourceColumnProfile unresolvedColumn = unresolved.FindColumn("Period")!;
        Assert.Equal(new DateTime(1930, 2, 1), resolvedColumn.MinimumDate);
        Assert.Equal(new DateTime(2029, 1, 1), resolvedColumn.MaximumDate);
        Assert.Equal(0, unresolvedColumn.DateLikeCount);
        Assert.Equal(2, unresolvedColumn.PeriodLikeWithoutYearCount);
        Assert.Equal(1d, unresolvedColumn.PeriodLikeRatio);
        Assert.Null(unresolvedColumn.MinimumDate);
    }

    [Theory]
    [InlineData("202613")]
    [InlineData("2026-13")]
    [InlineData("Q5 2026")]
    [InlineData("01/02/2026")]
    public void Does_not_accept_invalid_or_ambiguous_period_text(string value)
    {
        var profile = SourceProfiler.Profile(
            new[] { "Candidate" },
            new object?[][] { new object?[] { value } });

        SourceColumnProfile column = profile.FindColumn("Candidate")!;
        Assert.Equal(0, column.DateLikeCount);
        Assert.Equal(0, column.PeriodLikeWithoutYearCount);
        Assert.NotEqual(SourceValueType.Date, column.InferredType);
        Assert.NotEqual(SourceValueType.DateTime, column.InferredType);
    }

    [Fact]
    public void Profiles_and_projects_one_hundred_thousand_synthetic_rows_in_normal_ci()
    {
        var profile = SourceProfiler.Profile(
            SyntheticScaleData.Headers,
            SyntheticScaleData.CreateRows(100_000));
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            ReportingYear = 2026
        };
        for (var month = 1; month <= 12; month++)
        {
            mapping.Columns.Add(new PeriodColumnMapping { SourceColumn = "Month" + month, Month = month });
        }

        var projection = RowProjectionCalculator.Project(profile.RowCount, mapping);

        Assert.Equal(100_000, profile.RowCount);
        Assert.Empty(profile.Issues);
        Assert.Equal(1_200_000, projection.ProjectedRows);
        Assert.Equal(SourceLoadRoute.DataModel, projection.Route);
    }

    [Fact]
    public void Full_worksheet_scale_profile_generator_is_opt_in()
    {
        if (!string.Equals(
            Environment.GetEnvironmentVariable(SyntheticScaleData.FullScaleEnvironmentVariable),
            "1",
            StringComparison.Ordinal))
        {
            return;
        }

        var rows = SyntheticScaleData.CreateRows((int)RowProjection.MaximumWorksheetDataRows);
        var profile = SourceProfiler.Profile(SyntheticScaleData.Headers, rows);

        Assert.Equal(RowProjection.MaximumWorksheetDataRows, profile.RowCount);
        Assert.Empty(profile.Issues);
    }
}
