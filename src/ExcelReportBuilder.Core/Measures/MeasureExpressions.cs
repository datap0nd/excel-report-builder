using System.Collections.Generic;

namespace ExcelReportBuilder.Core.Measures
{
    public enum MeasureValueType
    {
        WholeNumber,
        Number,
        Currency,
        Percentage
    }

    public enum MeasureExpressionKind
    {
        Aggregate,
        FilteredAggregate,
        WeightedAggregate,
        Reference,
        Constant,
        Binary,
        SafeDivide,
        Ratio,
        Difference,
        Share
    }

    public enum AggregateFunction
    {
        Sum,
        Count,
        DistinctCount,
        Average,
        Minimum,
        Maximum
    }

    public enum BinaryMeasureOperator
    {
        Add,
        Subtract,
        Multiply,
        Divide
    }

    public enum DifferenceKind
    {
        Absolute,
        Percentage,
        PercentagePoints
    }

    public enum ZeroDenominatorBehavior
    {
        Blank,
        Zero,
        Error
    }

    public enum ShareDenominatorScope
    {
        Explicit,
        Parent,
        FilteredReportTotal
    }

    public enum MeasureFilterOperator
    {
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        In,
        NotIn,
        IsBlank,
        IsNotBlank
    }

    public sealed class MeasureDefinition
    {
        public string Id { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;

        public MeasureValueType ValueType { get; set; } = MeasureValueType.Number;

        public MeasureExpression Expression { get; set; } = new ConstantMeasureExpression();

        public string? NumberFormat { get; set; }
    }

    /// <summary>
    /// A typed expression graph. Every node declares its result type, allowing
    /// deterministic validation before any Excel object is touched.
    /// </summary>
    public abstract class MeasureExpression
    {
        public abstract MeasureExpressionKind Kind { get; }

        public MeasureValueType ResultType { get; set; } = MeasureValueType.Number;
    }

    public sealed class AggregateMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Aggregate;

        public string Field { get; set; } = string.Empty;

        public AggregateFunction Function { get; set; } = AggregateFunction.Sum;

        public string? PeriodSliceId { get; set; }
    }

    public sealed class FilteredAggregateMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.FilteredAggregate;

        public string Field { get; set; } = string.Empty;

        public AggregateFunction Function { get; set; } = AggregateFunction.Sum;

        public List<MeasureFilterSpec> Filters { get; set; } = new List<MeasureFilterSpec>();

        public string? PeriodSliceId { get; set; }
    }

    public sealed class WeightedAggregateMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.WeightedAggregate;

        /// <summary>
        /// The aggregate of a deterministic row-level value-times-weight column.
        /// The column itself is produced by an AddArithmeticColumnTransform.
        /// </summary>
        public MeasureExpression Numerator { get; set; } = new AggregateMeasureExpression();

        /// <summary>
        /// The aggregate of the weight column over the identical filter scope.
        /// </summary>
        public MeasureExpression Denominator { get; set; } = new AggregateMeasureExpression();

        public ZeroDenominatorBehavior OnZero { get; set; } = ZeroDenominatorBehavior.Blank;
    }

    public sealed class MeasureFilterSpec
    {
        public string Field { get; set; } = string.Empty;

        public MeasureFilterOperator Operator { get; set; }

        public List<Specifications.ScalarValue> Values { get; set; } = new List<Specifications.ScalarValue>();
    }

    public sealed class ReferenceMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Reference;

        public string MeasureId { get; set; } = string.Empty;
    }

    public sealed class ConstantMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Constant;

        public decimal Value { get; set; }
    }

    public sealed class BinaryMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Binary;

        public BinaryMeasureOperator Operator { get; set; }

        public MeasureExpression Left { get; set; } = new ConstantMeasureExpression();

        public MeasureExpression Right { get; set; } = new ConstantMeasureExpression();

        public bool ReturnBlankOnZeroDenominator { get; set; } = true;
    }

    public sealed class SafeDivideMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.SafeDivide;

        public MeasureExpression Numerator { get; set; } = new ConstantMeasureExpression();

        public MeasureExpression Denominator { get; set; } = new ConstantMeasureExpression();

        public ZeroDenominatorBehavior OnZero { get; set; } = ZeroDenominatorBehavior.Blank;

        public bool AsPercentage { get; set; }
    }

    public sealed class RatioMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Ratio;

        public MeasureExpression Numerator { get; set; } = new ConstantMeasureExpression();

        public MeasureExpression Denominator { get; set; } = new ConstantMeasureExpression();

        public ZeroDenominatorBehavior OnZero { get; set; } = ZeroDenominatorBehavior.Blank;
    }

    public sealed class DifferenceMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Difference;

        public DifferenceKind DifferenceKind { get; set; }

        public MeasureExpression Current { get; set; } = new ConstantMeasureExpression();

        public MeasureExpression Baseline { get; set; } = new ConstantMeasureExpression();

        public ZeroDenominatorBehavior OnZero { get; set; } = ZeroDenominatorBehavior.Blank;
    }

    public sealed class ShareMeasureExpression : MeasureExpression
    {
        public override MeasureExpressionKind Kind => MeasureExpressionKind.Share;

        public MeasureExpression Part { get; set; } = new ConstantMeasureExpression();

        public MeasureExpression Whole { get; set; } = new ConstantMeasureExpression();

        public ZeroDenominatorBehavior OnZero { get; set; } = ZeroDenominatorBehavior.Blank;

        public ShareDenominatorScope Scope { get; set; } = ShareDenominatorScope.Explicit;
    }
}
