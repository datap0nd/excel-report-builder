using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PeriodDetectorTests
{
    [Fact]
    public void Detects_a_unique_long_date_column()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Region", "Amount" },
            new object?[][]
            {
                new object?[] { new DateTime(2026, 1, 1), "North", 10m },
                new object?[] { new DateTime(2026, 2, 1), "South", 20m }
            });

        var result = PeriodDetector.Detect(profile);

        Assert.Equal(PeriodLayoutKind.LongDateColumn, result.Kind);
        Assert.Equal("Period", result.DateColumn);
        Assert.False(result.IsAmbiguous);
        Assert.Equal(PeriodMappingKind.LongDateColumn, result.ToPeriodMapping().Kind);
    }

    [Fact]
    public void Does_not_choose_between_multiple_date_columns()
    {
        var profile = SourceProfiler.Profile(
            new[] { "OrderDate", "PostingDate", "Amount" },
            new object?[][]
            {
                new object?[] { new DateTime(2026, 1, 1), new DateTime(2026, 1, 2), 10m },
                new object?[] { new DateTime(2026, 2, 1), new DateTime(2026, 2, 2), 20m }
            });

        var result = PeriodDetector.Detect(profile);

        Assert.True(result.IsAmbiguous);
        Assert.Null(result.DateColumn);
        Assert.Contains(result.Issues, issue => issue.Code == PeriodDetectionIssueCode.MultipleDateColumns);
    }

    [Fact]
    public void Month_headers_without_year_require_explicit_input()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "Jan", "February", "Mar" },
            new object?[][] { new object?[] { "North", 10m, 20m, 30m } });

        var unresolved = PeriodDetector.Detect(profile);
        var resolved = PeriodDetector.Detect(profile, 2026);

        Assert.Equal(PeriodLayoutKind.MonthHeaders, unresolved.Kind);
        Assert.True(unresolved.RequiresReportingYear);
        Assert.True(unresolved.IsAmbiguous);
        Assert.All(unresolved.HeaderMatches, match => Assert.Null(match.CanonicalPeriod));
        Assert.False(resolved.IsAmbiguous);
        Assert.All(resolved.HeaderMatches, match => Assert.Equal(2026, match.CanonicalPeriod!.Value.Year));
    }

    [Fact]
    public void Detects_complete_metric_month_headers_and_rejects_incomplete_matrix()
    {
        var complete = SourceProfiler.Profile(
            new[] { "Region", "Revenue Jan 2026", "Cost Jan 2026", "Revenue Feb 2026", "Cost Feb 2026" },
            new object?[][] { new object?[] { "North", 10m, 4m, 11m, 5m } });
        var incomplete = SourceProfiler.Profile(
            new[] { "Region", "Revenue Jan 2026", "Cost Jan 2026", "Revenue Feb 2026" },
            new object?[][] { new object?[] { "North", 10m, 4m, 11m } });

        var completeResult = PeriodDetector.Detect(complete);
        var incompleteResult = PeriodDetector.Detect(incomplete);

        Assert.Equal(PeriodLayoutKind.MetricMonthHeaders, completeResult.Kind);
        Assert.False(completeResult.IsAmbiguous);
        Assert.Equal(new[] { "Cost", "Revenue" }, completeResult.HeaderMatches.Select(item => item.Metric).Distinct().Order());
        Assert.Contains(
            incompleteResult.Issues,
            issue => issue.Code == PeriodDetectionIssueCode.IncompleteMetricPeriodMatrix);
    }

    [Fact]
    public void Detects_compact_year_month_headers()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "202601", "202602", "202603" },
            new object?[][] { new object?[] { "North", 10m, 20m, 30m } });

        var result = PeriodDetector.Detect(profile);

        Assert.Equal(PeriodLayoutKind.MonthHeaders, result.Kind);
        Assert.Equal(PeriodGrain.Month, result.Grain);
        Assert.False(result.IsAmbiguous);
        Assert.Equal(new[] { 1, 2, 3 }, result.HeaderMatches.Select(match => match.Month));
        Assert.All(result.HeaderMatches, match => Assert.Equal(2026, match.Year));
    }

    [Fact]
    public void Expands_two_digit_years_with_a_fixed_excel_compatible_cutoff()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "Jan-29", "Feb-30" },
            new object?[][] { new object?[] { "North", 10m, 20m } });

        var result = PeriodDetector.Detect(profile);

        Assert.False(result.IsAmbiguous);
        Assert.Equal(2029, result.HeaderMatches.Single(match => match.Month == 1).Year);
        Assert.Equal(1930, result.HeaderMatches.Single(match => match.Month == 2).Year);
    }

    [Fact]
    public void Detects_quarter_headers_at_quarter_grain()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "Q1 2026", "2026-Q2", "Q3-2026", "2026 Q4" },
            new object?[][] { new object?[] { "North", 10m, 20m, 30m, 40m } });

        var result = PeriodDetector.Detect(profile);
        var mapping = result.ToPeriodMapping();

        Assert.Equal(PeriodLayoutKind.MonthHeaders, result.Kind);
        Assert.Equal(PeriodGrain.Quarter, result.Grain);
        Assert.Equal(PeriodGrain.Quarter, mapping.Grain);
        Assert.Equal(new[] { 1, 4, 7, 10 }, result.HeaderMatches.Select(match => match.Month));
        Assert.Equal(
            new[]
            {
                new DateTime(2026, 1, 1),
                new DateTime(2026, 4, 1),
                new DateTime(2026, 7, 1),
                new DateTime(2026, 10, 1)
            },
            result.HeaderMatches.Select(match => match.CanonicalPeriod!.Value));
    }

    [Fact]
    public void Detects_complete_metric_quarter_headers()
    {
        var profile = SourceProfiler.Profile(
            new[]
            {
                "Region", "Sales Q1 2026", "Qty Q1 2026",
                "2026-Q2 Sales", "2026-Q2 Qty"
            },
            new object?[][] { new object?[] { "North", 10m, 2m, 12m, 3m } });

        var result = PeriodDetector.Detect(profile);

        Assert.Equal(PeriodLayoutKind.MetricMonthHeaders, result.Kind);
        Assert.Equal(PeriodGrain.Quarter, result.Grain);
        Assert.False(result.IsAmbiguous);
        Assert.Equal(new[] { "Qty", "Sales" }, result.HeaderMatches.Select(match => match.Metric).Distinct().Order());
    }

    [Fact]
    public void Quarter_headers_without_year_require_explicit_input()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "Q1", "Q2" },
            new object?[][] { new object?[] { "North", 10m, 20m } });

        var unresolved = PeriodDetector.Detect(profile);
        var resolved = PeriodDetector.Detect(profile, 2026);

        Assert.True(unresolved.RequiresReportingYear);
        Assert.True(unresolved.IsAmbiguous);
        Assert.All(unresolved.HeaderMatches, match => Assert.Null(match.CanonicalPeriod));
        Assert.False(resolved.IsAmbiguous);
        Assert.All(resolved.HeaderMatches, match => Assert.Equal(2026, match.CanonicalPeriod!.Value.Year));
    }

    [Fact]
    public void Rejects_mixed_month_and_quarter_headers()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Region", "Jan 2026", "Q2 2026" },
            new object?[][] { new object?[] { "North", 10m, 20m } });

        var result = PeriodDetector.Detect(profile);

        Assert.True(result.IsAmbiguous);
        Assert.Contains(result.Issues, issue => issue.Code == PeriodDetectionIssueCode.MixedPeriodGrains);
    }

    [Fact]
    public void Detects_supported_month_period_values_in_a_long_column()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Amount" },
            new object?[][]
            {
                new object?[] { 202601, 10m },
                new object?[] { "Feb-26", 20m },
                new object?[] { "March 2026", 30m },
                new object?[] { "2026-04", 40m }
            });

        var result = PeriodDetector.Detect(profile);

        Assert.Equal(PeriodLayoutKind.LongDateColumn, result.Kind);
        Assert.Equal("Period", result.DateColumn);
        Assert.Equal(PeriodGrain.Month, result.Grain);
        Assert.False(result.IsAmbiguous);
        Assert.Equal(4, profile.FindColumn("Period")!.DateLikeCount);
    }

    [Fact]
    public void Detects_supported_quarter_period_values_in_a_long_column()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Amount" },
            new object?[][]
            {
                new object?[] { "Q1 2026", 10m },
                new object?[] { "2026-Q2", 20m }
            });

        var result = PeriodDetector.Detect(profile);

        Assert.Equal(PeriodLayoutKind.LongDateColumn, result.Kind);
        Assert.Equal(PeriodGrain.Quarter, result.Grain);
        Assert.False(result.IsAmbiguous);
    }

    [Fact]
    public void Long_month_names_without_year_are_never_inferred()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Amount" },
            new object?[][]
            {
                new object?[] { "Jan", 10m },
                new object?[] { "February", 20m }
            });

        var unresolved = PeriodDetector.Detect(profile);
        var resolved = PeriodDetector.Detect(profile, 2026);

        Assert.Equal(PeriodLayoutKind.LongDateColumn, unresolved.Kind);
        Assert.True(unresolved.RequiresReportingYear);
        Assert.True(unresolved.IsAmbiguous);
        Assert.Throws<InvalidOperationException>(() => unresolved.ToPeriodMapping());
        Assert.False(resolved.IsAmbiguous);
        Assert.Equal(2026, resolved.ToPeriodMapping().ReportingYear);
    }

    [Fact]
    public void Rejects_mixed_grains_in_a_long_period_column()
    {
        var profile = SourceProfiler.Profile(
            new[] { "Period", "Amount" },
            new object?[][]
            {
                new object?[] { "Jan 2026", 10m },
                new object?[] { "Q2 2026", 20m }
            });

        var result = PeriodDetector.Detect(profile);

        Assert.True(result.IsAmbiguous);
        Assert.Contains(result.Issues, issue => issue.Code == PeriodDetectionIssueCode.MixedPeriodGrains);
    }
}
