using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

public sealed class WidePeriodNormalizerTests
{
    [Fact]
    public void Metric_month_normalization_preserves_every_value_and_total()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MetricMonthHeaders,
            ReportingYear = 2026,
            KeyColumns = { "Region" },
            Columns =
            {
                Map("Revenue Jan", 1, "Revenue"),
                Map("Cost Jan", 1, "Cost"),
                Map("Revenue Feb", 2, "Revenue"),
                Map("Cost Feb", 2, "Cost")
            }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?>
            {
                ["Region"] = "North",
                ["Revenue Jan"] = 10m,
                ["Cost Jan"] = 4m,
                ["Revenue Feb"] = 11m,
                ["Cost Feb"] = 5m
            },
            new Dictionary<string, object?>
            {
                ["Region"] = "South",
                ["Revenue Jan"] = 20m,
                ["Cost Jan"] = 8m,
                ["Revenue Feb"] = 22m,
                ["Cost Feb"] = 9m
            }
        };

        var normalized = WidePeriodNormalizer.Normalize(source, mapping);

        Assert.Equal(8, normalized.Count);
        Assert.Equal(89m, normalized.Sum(row => (decimal)row.Value!));
        Assert.Equal(89m, source.Sum(row => mapping.Columns.Sum(column => (decimal)row[column.SourceColumn]!)));
        Assert.All(normalized, row => Assert.Equal(2026, row.Period.Year));
    }

    [Fact]
    public void Wide_month_normalization_preserves_a_blank_cell_as_a_canonical_row()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            ReportingYear = 2026,
            KeyColumns = { "Region" },
            Columns =
            {
                Map("Jan", 1, null),
                Map("Feb", 2, null)
            }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?>
            {
                ["Region"] = "North",
                ["Jan"] = null,
                ["Feb"] = 12m
            }
        };

        var normalized = WidePeriodNormalizer.Normalize(source, mapping);

        Assert.Equal(2, normalized.Count);
        NormalizedPeriodValue january = Assert.Single(
            normalized,
            row => row.SourceColumn == "Jan");
        Assert.Equal(new DateTime(2026, 1, 1), january.Period);
        Assert.Null(january.Value);
        Assert.Equal("North", january.Keys["Region"]);
        Assert.Equal(12m, normalized.Where(row => row.Value != null).Sum(row => (decimal)row.Value!));
    }

    [Fact]
    public void Metric_month_normalization_keeps_blank_cells_in_the_multi_metric_matrix()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MetricMonthHeaders,
            ReportingYear = 2026,
            KeyColumns = { "Region" },
            Columns =
            {
                Map("Revenue Jan", 1, "Revenue"),
                Map("Cost Jan", 1, "Cost"),
                Map("Revenue Feb", 2, "Revenue"),
                Map("Cost Feb", 2, "Cost")
            }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?>
            {
                ["Region"] = "North",
                ["Revenue Jan"] = 10m,
                ["Cost Jan"] = null,
                ["Revenue Feb"] = 11m,
                ["Cost Feb"] = 5m
            }
        };

        var normalized = WidePeriodNormalizer.Normalize(source, mapping);

        Assert.Equal(4, normalized.Count);
        Assert.Equal(2, normalized.Count(row => row.Metric == "Revenue"));
        Assert.Equal(2, normalized.Count(row => row.Metric == "Cost"));
        NormalizedPeriodValue blankCost = Assert.Single(
            normalized,
            row => row.SourceColumn == "Cost Jan");
        Assert.Equal("Cost", blankCost.Metric);
        Assert.Null(blankCost.Value);
        Assert.Equal(26m, normalized.Where(row => row.Value != null).Sum(row => (decimal)row.Value!));
    }

    [Fact]
    public void Missing_year_is_never_inferred()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            Columns = { Map("Jan", 1, null) }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?> { ["Jan"] = 10m }
        };

        Assert.Throws<InvalidOperationException>(() => WidePeriodNormalizer.Normalize(source, mapping));
    }

    [Fact]
    public void Quarter_normalization_uses_quarter_start_dates_and_preserves_totals()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            Grain = PeriodGrain.Quarter,
            KeyColumns = { "Region" },
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Q1 2026", Month = 1, Year = 2026 },
                new PeriodColumnMapping { SourceColumn = "2026-Q2", Month = 4, Year = 2026 }
            }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?>
            {
                ["Region"] = "North",
                ["Q1 2026"] = 30m,
                ["2026-Q2"] = 45m
            },
            new Dictionary<string, object?>
            {
                ["Region"] = "South",
                ["Q1 2026"] = 20m,
                ["2026-Q2"] = 25m
            }
        };

        var normalized = WidePeriodNormalizer.Normalize(source, mapping);

        Assert.Equal(4, normalized.Count);
        Assert.Equal(new[] { 1, 4 }, normalized.Select(row => row.Period.Month).Distinct().Order());
        Assert.Equal(120m, normalized.Sum(row => (decimal)row.Value!));
        Assert.Equal(
            120m,
            source.Sum(row => mapping.Columns.Sum(column => (decimal)row[column.SourceColumn]!)));
    }

    [Fact]
    public void Quarter_normalization_rejects_non_start_months()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            Grain = PeriodGrain.Quarter,
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Bad quarter", Month = 2, Year = 2026 }
            }
        };
        var source = new IReadOnlyDictionary<string, object?>[]
        {
            new Dictionary<string, object?> { ["Bad quarter"] = 10m }
        };

        Assert.Throws<ArgumentException>(() => WidePeriodNormalizer.Normalize(source, mapping));
    }

    private static PeriodColumnMapping Map(string source, int month, string? metric)
    {
        return new PeriodColumnMapping { SourceColumn = source, Month = month, Metric = metric };
    }
}
