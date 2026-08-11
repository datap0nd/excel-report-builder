using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class WeightedAggregateTests
{
    [Fact]
    public void Preserves_sum_product_and_weight_totals_over_the_same_rows()
    {
        var result = WeightedAggregateCalculator.Calculate(new[]
        {
            new WeightedObservation { Value = 10m, Weight = 2m },
            new WeightedObservation { Value = 20m, Weight = 1m },
            new WeightedObservation { Value = null, Weight = 99m }
        });

        Assert.Equal(40m, result.Numerator);
        Assert.Equal(3m, result.Denominator);
        Assert.Equal(40m / 3m, result.Value);
        Assert.Equal(2, result.IncludedRows);
        Assert.Equal(1, result.ExcludedRows);
    }

    [Fact]
    public void Validator_requires_an_owned_typed_product_column()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Transforms.RemoveAll(transform => transform.Id == "derive_weighted_units");

        var validation = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(validation.Issues, issue => issue.Code == "WEIGHTED_PRODUCT_TRANSFORM_REQUIRED");
    }

    [Fact]
    public void Zero_denominator_behavior_is_explicit()
    {
        var observations = new[]
        {
            new WeightedObservation { Value = 10m, Weight = 0m }
        };

        Assert.Null(WeightedAggregateCalculator.Calculate(observations).Value);
        Assert.Equal(0m, WeightedAggregateCalculator.Calculate(observations, ZeroDenominatorBehavior.Zero).Value);
        Assert.Throws<DivideByZeroException>(() =>
            WeightedAggregateCalculator.Calculate(observations, ZeroDenominatorBehavior.Error));
    }

    [Fact]
    public void Validator_requires_sum_for_both_weighted_components()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var weighted = (WeightedAggregateMeasureExpression)spec.Measures
            .Single(measure => measure.Id == "weighted_rate").Expression;
        ((FilteredAggregateMeasureExpression)weighted.Numerator).Function = AggregateFunction.Average;

        var validation = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(validation.Issues, issue => issue.Code == "WEIGHTED_SUM_REQUIRED");
    }
}
