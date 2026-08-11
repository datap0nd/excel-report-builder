using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

public sealed class RowProjectionTests
{
    [Fact]
    public void Month_headers_expand_once_per_header_and_route_without_truncation()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            ReportingYear = 2026,
            Columns =
            {
                Map("Jan", 1),
                Map("Feb", 2),
                Map("Mar", 3)
            }
        };

        var worksheet = RowProjectionCalculator.Project(100, mapping);
        var dataModel = RowProjectionCalculator.Project(400_000, mapping);

        Assert.Equal(300, worksheet.ProjectedRows);
        Assert.Equal(SourceLoadRoute.Worksheet, worksheet.Route);
        Assert.Equal(1_200_000, dataModel.ProjectedRows);
        Assert.Equal(SourceLoadRoute.DataModel, dataModel.Route);
        Assert.False(dataModel.Route == SourceLoadRoute.Worksheet);
    }

    [Fact]
    public void Metric_month_headers_expand_once_per_source_metric_period_column()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MetricMonthHeaders,
            ReportingYear = 2026,
            Columns =
            {
                Map("Revenue Jan", 1, "Revenue"),
                Map("Cost Jan", 1, "Cost"),
                Map("Revenue Feb", 2, "Revenue"),
                Map("Cost Feb", 2, "Cost")
            }
        };

        var result = RowProjectionCalculator.Project(12, mapping);

        Assert.Equal(4, result.ExpansionFactor);
        Assert.Equal(48, result.ProjectedRows);
    }

    [Fact]
    public void Projection_overflow_routes_to_data_model()
    {
        var mapping = new PeriodMappingSpec
        {
            Kind = PeriodMappingKind.MonthHeaders,
            ReportingYear = 2026,
            Columns = { Map("Jan", 1), Map("Feb", 2) }
        };

        var result = RowProjectionCalculator.Project(long.MaxValue, mapping);

        Assert.Equal(long.MaxValue, result.ProjectedRows);
        Assert.Equal(SourceLoadRoute.DataModel, result.Route);
    }

    private static PeriodColumnMapping Map(string source, int month, string? metric = null)
    {
        return new PeriodColumnMapping { SourceColumn = source, Month = month, Metric = metric };
    }
}
