using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Rendering;
using ExcelReportBuilder.Excel.Validation;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotMeasureEvaluatorTests
{
    [Fact]
    public void Independently_evaluates_filtered_typed_ratio_and_caches_pivot_reads()
    {
        var regionFilter = new MeasureFilterSpec
        {
            Field = "Region",
            Operator = MeasureFilterOperator.In,
            Values = { ScalarValue.FromText("East"), ScalarValue.FromText("West") }
        };
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["rate"] = new MeasureDefinition
            {
                Id = "rate",
                Label = "Rate",
                Expression = new RatioMeasureExpression
                {
                    Numerator = new FilteredAggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum,
                        Filters = { regionFilter }
                    },
                    Denominator = new AggregateMeasureExpression
                    {
                        Field = "Units",
                        Function = AggregateFunction.Sum
                    },
                    OnZero = ZeroDenominatorBehavior.Blank
                }
            }
        };
        var pivot = new PivotBuildResult
        {
            WorksheetName = "Managed pivot",
            AnchorCell = "$A$3",
            DataFields = new[]
            {
                Descriptor("rate", "Amount", "Amount input", new[] { regionFilter }),
                Descriptor("rate", "Units", "Units input")
            }
        };
        var reads = 0;
        var evaluator = new PivotMeasureEvaluator(
            measures,
            pivot,
            (caption, filters) =>
            {
                reads++;
                if (caption == "Units input")
                {
                    Assert.Empty(filters);
                    return DenseFormulaExpectation.Number(20m);
                }

                var region = Assert.Single(filters);
                Assert.Equal("Region", region.Field);
                return DenseFormulaExpectation.Number(
                    string.Equals(region.Value as string, "East", StringComparison.Ordinal) ? 60m : 40m);
            });

        var first = evaluator.EvaluateAcrossMemberSets(
            "rate",
            EmptyMemberSets(),
            EmptyPeriodSets());
        var second = evaluator.EvaluateAcrossMemberSets(
            "rate",
            EmptyMemberSets(),
            EmptyPeriodSets());

        Assert.Equal(DenseFormulaExpectationKind.Number, first.Kind);
        Assert.Equal(5m, first.NumericValue);
        Assert.Equal(5m, second.NumericValue);
        Assert.Equal(3, reads);
    }

    [Fact]
    public void Zero_denominator_returns_blank_without_reading_the_numerator()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["rate"] = new MeasureDefinition
            {
                Id = "rate",
                Label = "Rate",
                Expression = new SafeDivideMeasureExpression
                {
                    Numerator = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    },
                    Denominator = new AggregateMeasureExpression
                    {
                        Field = "Units",
                        Function = AggregateFunction.Sum
                    },
                    OnZero = ZeroDenominatorBehavior.Blank
                }
            }
        };
        var pivot = new PivotBuildResult
        {
            DataFields = new[]
            {
                Descriptor("rate", "Amount", "Amount input"),
                Descriptor("rate", "Units", "Units input")
            }
        };
        var numeratorRead = false;
        var evaluator = new PivotMeasureEvaluator(
            measures,
            pivot,
            (caption, _) =>
            {
                if (caption == "Amount input")
                {
                    numeratorRead = true;
                    return DenseFormulaExpectation.Error();
                }

                return DenseFormulaExpectation.Number(0m);
            });

        var result = evaluator.EvaluateAcrossMemberSets(
            "rate",
            EmptyMemberSets(),
            EmptyPeriodSets());

        Assert.Equal(DenseFormulaExpectationKind.Blank, result.Kind);
        Assert.False(numeratorRead);
    }

    [Fact]
    public void Share_of_parent_removes_only_the_deepest_row_member()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["share"] = new MeasureDefinition
            {
                Id = "share",
                Label = "Share",
                Expression = new ShareMeasureExpression
                {
                    Part = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    },
                    Whole = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    },
                    Scope = ShareDenominatorScope.Parent,
                    OnZero = ZeroDenominatorBehavior.Blank
                }
            }
        };
        var pivot = new PivotBuildResult
        {
            DataFields = new[] { Descriptor("share", "Amount", "Amount input") }
        };
        var evaluator = new PivotMeasureEvaluator(
            measures,
            pivot,
            (_, filters) => DenseFormulaExpectation.Number(
                filters.Any(filter => filter.Field == "Category") ? 25m : 100m));
        IReadOnlyList<IReadOnlyList<PivotFilterItem>> members = new[]
        {
            (IReadOnlyList<PivotFilterItem>)new[]
            {
                new PivotFilterItem { Field = "Region", Value = "East" },
                new PivotFilterItem { Field = "Category", Value = "A" },
                new PivotFilterItem { Field = "Period", Value = new DateTime(2026, 1, 1) }
            }
        };

        var result = evaluator.EvaluateAcrossMemberSets(
            "share",
            members,
            EmptyPeriodSets(),
            new[] { "Region", "Category" });

        Assert.Equal(DenseFormulaExpectationKind.Number, result.Kind);
        Assert.Equal(0.25m, result.NumericValue);
    }

    [Fact]
    public void Missing_aggregate_descriptor_fails_closed()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["amount"] = new MeasureDefinition
            {
                Id = "amount",
                Label = "Amount",
                Expression = new AggregateMeasureExpression
                {
                    Field = "Amount",
                    Function = AggregateFunction.Sum
                }
            }
        };
        var evaluator = new PivotMeasureEvaluator(
            measures,
            new PivotBuildResult(),
            (_, _) => DenseFormulaExpectation.Number(1m));

        var exception = Assert.Throws<InvalidOperationException>(() =>
            evaluator.EvaluateAcrossMemberSets(
                "amount",
                EmptyMemberSets(),
                EmptyPeriodSets()));

        Assert.Contains("independent validation", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Aggregate_reader_error_fails_closed_instead_of_becoming_a_formula_blank()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["amount"] = new MeasureDefinition
            {
                Id = "amount",
                Label = "Amount",
                Expression = new AggregateMeasureExpression
                {
                    Field = "Amount",
                    Function = AggregateFunction.Sum
                }
            }
        };
        var evaluator = new PivotMeasureEvaluator(
            measures,
            new PivotBuildResult
            {
                DataFields = new[] { Descriptor("amount", "Amount", "Amount input") }
            },
            (_, _) => DenseFormulaExpectation.Error());

        var exception = Assert.Throws<InvalidOperationException>(() =>
            evaluator.EvaluateAcrossMemberSets(
                "amount",
                EmptyMemberSets(),
                EmptyPeriodSets()));

        Assert.Contains("independently", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static PivotDataFieldDescriptor Descriptor(
        string measureId,
        string field,
        string caption,
        IReadOnlyList<MeasureFilterSpec>? filters = null)
    {
        return new PivotDataFieldDescriptor
        {
            MeasureId = measureId,
            ComponentId = measureId + "_" + field,
            Role = AggregateComponentRole.Input,
            SourceField = field,
            Function = AggregateFunction.Sum,
            PivotCaption = caption,
            Filters = filters ?? Array.Empty<MeasureFilterSpec>()
        };
    }

    private static IReadOnlyList<IReadOnlyList<PivotFilterItem>> EmptyMemberSets()
    {
        return new[] { (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>() };
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> EmptyPeriodSets()
    {
        return new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>(
            StringComparer.OrdinalIgnoreCase);
    }
}
