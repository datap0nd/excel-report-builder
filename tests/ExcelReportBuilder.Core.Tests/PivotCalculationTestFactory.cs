using ExcelReportBuilder.Core.PivotPlus.Calculations;

namespace ExcelReportBuilder.Core.Tests;

internal static class PivotCalculationTestFactory
{
    public static PivotModelSchema Schema()
    {
        return new PivotModelSchema(new[]
        {
            new PivotModelTableSchema(
                "fact",
                "Fact Sales",
                new[]
                {
                    new PivotModelFieldSchema("amount", "Amount", PivotModelDataType.DecimalNumber),
                    new PivotModelFieldSchema("units", "Units", PivotModelDataType.WholeNumber),
                    new PivotModelFieldSchema("price", "Unit Price", PivotModelDataType.Currency),
                    new PivotModelFieldSchema("weight", "Weight", PivotModelDataType.DecimalNumber),
                    new PivotModelFieldSchema(
                        "period",
                        "Month Start",
                        PivotModelDataType.Date,
                        new[]
                        {
                            new PivotModelMember("jan", PivotScalarValue.Date(new DateTime(2026, 1, 1))),
                            new PivotModelMember("feb", PivotScalarValue.Date(new DateTime(2026, 2, 1))),
                            new PivotModelMember("mar", PivotScalarValue.Date(new DateTime(2026, 3, 1)))
                        }),
                    new PivotModelFieldSchema(
                        "scenario",
                        "Scenario",
                        PivotModelDataType.Text,
                        new[]
                        {
                            new PivotModelMember("actual", PivotScalarValue.Text("Actual")),
                            new PivotModelMember("plan", PivotScalarValue.Text("Plan"))
                        }),
                    new PivotModelFieldSchema(
                        "region",
                        "Region",
                        PivotModelDataType.Text,
                        new[]
                        {
                            new PivotModelMember("north", PivotScalarValue.Text("North")),
                            new PivotModelMember("south", PivotScalarValue.Text("South"))
                        }),
                    new PivotModelFieldSchema(
                        "department",
                        "Department",
                        PivotModelDataType.Text,
                        new[]
                        {
                            new PivotModelMember("consumer", PivotScalarValue.Text("Consumer")),
                            new PivotModelMember("enterprise", PivotScalarValue.Text("Enterprise"))
                        }),
                    new PivotModelFieldSchema("active", "Is Active", PivotModelDataType.Boolean),
                    new PivotModelFieldSchema("transaction_date", "Transaction Date", PivotModelDataType.Date)
                }),
            new PivotModelTableSchema(
                "product",
                "Dim Product",
                new[]
                {
                    new PivotModelFieldSchema(
                        "family",
                        "Product Family",
                        PivotModelDataType.Text,
                        new[]
                        {
                            new PivotModelMember("core", PivotScalarValue.Text("Core")),
                            new PivotModelMember("plus", PivotScalarValue.Text("Plus"))
                        })
                })
        });
    }

    public static PivotPeriodDefinition MonthlyPeriods(bool includeMarch = true)
    {
        var coverage = new List<PivotPeriodCoverageMember>
        {
            CoverageMonth(1, "jan") ,
            CoverageMonth(2, "feb")
        };
        if (includeMarch)
        {
            coverage.Add(CoverageMonth(3, "mar"));
        }

        return new PivotPeriodDefinition(
            new PivotPeriodSource(
                "period",
                PivotPeriodGrain.Month,
                PivotPeriodCoverageStatus.Complete,
                coverage,
                scenarioFieldId: "scenario"),
            new[]
            {
                Slice("actual_jan", "Actual Jan", PivotPeriodGrain.Month, 1, "actual"),
                Slice("actual_feb", "Actual Feb", PivotPeriodGrain.Month, 2, "actual"),
                Slice("actual_mar", "Actual Mar", PivotPeriodGrain.Month, 3, "actual"),
                Slice("actual_q1", "Q1 Actual", PivotPeriodGrain.Quarter, 1, "actual"),
                Slice("plan_q1", "Q1 Plan", PivotPeriodGrain.Quarter, 1, "plan")
            });
    }

    public static PivotMeasureFormat Whole()
    {
        return new PivotMeasureFormat(PivotMeasureFormatKind.WholeNumber, 0, true);
    }

    public static PivotMeasureFormat Decimal(int places = 2)
    {
        return new PivotMeasureFormat(PivotMeasureFormatKind.DecimalNumber, places, true);
    }

    public static PivotMeasureFormat Currency(string marker = "EUR")
    {
        return new PivotMeasureFormat(PivotMeasureFormatKind.Currency, 2, true, marker);
    }

    public static PivotMeasureFormat Percentage(int places = 1)
    {
        return new PivotMeasureFormat(PivotMeasureFormatKind.Percentage, places, false);
    }

    public static PivotMeasureFormat PercentagePoints(int places = 1)
    {
        return new PivotMeasureFormat(PivotMeasureFormatKind.PercentagePoints, places, false);
    }

    public static PivotMeasureDefinition Measure(
        string id,
        string caption,
        PivotCalculationExpression expression,
        PivotMeasureFormat? format = null,
        string homeTableId = "fact")
    {
        return new PivotMeasureDefinition(
            id,
            caption,
            homeTableId,
            format ?? Decimal(),
            expression);
    }

    public static PivotAggregateExpression Sum(string? periodSliceId = null)
    {
        return new PivotAggregateExpression(
            "amount",
            PivotCalculationAggregateFunction.Sum,
            periodSliceId);
    }

    public static PivotMeasureSetDefinition Set(
        IEnumerable<PivotMeasureDefinition> measures,
        PivotPeriodDefinition? periods = null,
        PivotModelSchema? schema = null)
    {
        return new PivotMeasureSetDefinition(schema ?? Schema(), measures, periods);
    }

    private static PivotPeriodCoverageMember CoverageMonth(int month, string memberId)
    {
        return new PivotPeriodCoverageMember(
            new PivotPeriodPoint(PivotPeriodGrain.Month, 2026, month),
            PivotFilterValue.FromMember(memberId),
            new[] { "actual", "plan" });
    }

    private static PivotPeriodSlice Slice(
        string id,
        string caption,
        PivotPeriodGrain grain,
        int ordinal,
        string scenario)
    {
        return new PivotPeriodSlice(
            id,
            caption,
            new PivotPeriodPoint(grain, 2026, ordinal),
            scenario,
            PivotSliceFilterMode.ReplaceAxisContext);
    }
}
