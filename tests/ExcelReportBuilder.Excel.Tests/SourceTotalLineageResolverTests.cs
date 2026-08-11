using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Excel.Validation;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class SourceTotalLineageResolverTests
{
    private readonly SourceTotalLineageResolver _resolver = new();

    [Fact]
    public void Resolve_TracesRenameBackToRawSource()
    {
        var specification = new ReportSpecV1();
        specification.Transforms.Add(new RenameColumnTransform
        {
            Id = "rename",
            From = "Raw Amount",
            To = "Amount"
        });

        IReadOnlyList<string> result = _resolver.Resolve(
            specification,
            Sum("Amount"));

        Assert.Equal(new[] { "Raw Amount" }, result);
    }

    [Fact]
    public void Resolve_RejectsDerivedOrValueChangingLineage()
    {
        var derived = new ReportSpecV1();
        derived.Transforms.Add(new AddArithmeticColumnTransform
        {
            Id = "derived",
            OutputColumn = "Variance"
        });
        var converted = new ReportSpecV1();
        converted.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "convert",
            Column = "Amount",
            DataType = ColumnDataType.DecimalNumber
        });

        Assert.Empty(_resolver.Resolve(derived, Sum("Variance")));
        Assert.Empty(_resolver.Resolve(converted, Sum("Amount")));
    }

    [Fact]
    public void Resolve_MapsWideCanonicalValueToEveryUniqueRawPeriodColumn()
    {
        var specification = new ReportSpecV1
        {
            PeriodMapping = new PeriodMappingSpec
            {
                Kind = PeriodMappingKind.MonthHeaders,
                ValueColumnName = "Value",
                Columns =
                {
                    new PeriodColumnMapping { SourceColumn = "Jan", Month = 1, Year = 2026 },
                    new PeriodColumnMapping { SourceColumn = "Feb", Month = 2, Year = 2026 }
                }
            }
        };
        specification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize",
            PeriodMappingId = "periods"
        });

        IReadOnlyList<string> result = _resolver.Resolve(specification, Sum("Value"));

        Assert.Equal(new[] { "Jan", "Feb" }, result);
    }

    private static AggregateMeasureExpression Sum(string field)
    {
        return new AggregateMeasureExpression
        {
            Field = field,
            Function = AggregateFunction.Sum
        };
    }
}
