using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Rendering
{
    /// <summary>
    /// Converts the typed measure graph to host-owned Excel formulas. It has no
    /// entry point for arbitrary formula text.
    /// </summary>
    public sealed class MeasureFormulaCompiler
    {
        internal const int MaximumExpandedAggregateTerms = 512;

        private readonly GetPivotDataFormulaBuilder pivotFormulaBuilder = new GetPivotDataFormulaBuilder();

        public SafeExcelFormula Compile(
            string measureId,
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            PivotBuildResult pivot,
            IReadOnlyList<PivotFilterItem>? memberFilters = null,
            IReadOnlyList<string>? rowFieldOrder = null)
        {
            var filters = memberFilters ?? Array.Empty<PivotFilterItem>();
            return CompileAcrossMemberSets(
                measureId,
                measures,
                pivot,
                new[] { filters },
                new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>(StringComparer.OrdinalIgnoreCase),
                rowFieldOrder);
        }

        public SafeExcelFormula CompileAcrossMemberSets(
            string measureId,
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            PivotBuildResult pivot,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> periodFilterSets,
            IReadOnlyList<string>? rowFieldOrder = null)
        {
            if (!measures.TryGetValue(measureId, out var definition))
            {
                throw new InvalidOperationException("The requested measure is not defined.");
            }

            if (memberFilterSets == null || memberFilterSets.Count == 0)
            {
                throw new InvalidOperationException("At least one member-filter set is required.");
            }

            if (periodFilterSets == null)
            {
                throw new ArgumentNullException(nameof(periodFilterSets));
            }

            var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var expression = CompileExpression(
                definition.Expression,
                measureId,
                measures,
                pivot,
                memberFilterSets,
                periodFilterSets,
                rowFieldOrder ?? Array.Empty<string>(),
                active);
            return SafeFormulaFactory.FromTypedMeasure(
                RequiresErrorPropagation(definition.Expression, measures, new HashSet<string>(StringComparer.OrdinalIgnoreCase))
                    ? "=" + expression
                    : "=IFERROR(" + expression + ",\"\")");
        }

        private string CompileExpression(
            MeasureExpression expression,
            string ownerMeasureId,
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            PivotBuildResult pivot,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> periodFilterSets,
            IReadOnlyList<string> rowFieldOrder,
            ISet<string> active)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    return AggregateExpression(
                        ownerMeasureId,
                        aggregate.Field,
                        aggregate.Function,
                        Array.Empty<MeasureFilterSpec>(),
                        aggregate.PeriodSliceId,
                        pivot,
                        memberFilterSets,
                        periodFilterSets);
                case FilteredAggregateMeasureExpression filtered:
                    return AggregateExpression(
                        ownerMeasureId,
                        filtered.Field,
                        filtered.Function,
                        filtered.Filters,
                        filtered.PeriodSliceId,
                        pivot,
                        memberFilterSets,
                        periodFilterSets);
                case WeightedAggregateMeasureExpression weighted:
                    return SafeDivide(
                        CompileExpression(weighted.Numerator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        CompileExpression(weighted.Denominator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        weighted.OnZero);
                case ReferenceMeasureExpression reference:
                    if (!measures.TryGetValue(reference.MeasureId, out var referenced))
                    {
                        throw new InvalidOperationException("A referenced measure is not defined.");
                    }

                    if (!active.Add(reference.MeasureId))
                    {
                        throw new InvalidOperationException("The measure graph contains a cycle.");
                    }

                    try
                    {
                        return CompileExpression(
                            referenced.Expression,
                            ownerMeasureId,
                            measures,
                            pivot,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active);
                    }
                    finally
                    {
                        active.Remove(reference.MeasureId);
                    }
                case ConstantMeasureExpression constant:
                    return constant.Value.ToString(CultureInfo.InvariantCulture);
                case BinaryMeasureExpression binary:
                {
                    var left = CompileExpression(binary.Left, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active);
                    var right = CompileExpression(binary.Right, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active);
                    switch (binary.Operator)
                    {
                        case BinaryMeasureOperator.Add: return "(" + left + "+" + right + ")";
                        case BinaryMeasureOperator.Subtract: return "(" + left + "-" + right + ")";
                        case BinaryMeasureOperator.Multiply: return "(" + left + "*" + right + ")";
                        case BinaryMeasureOperator.Divide:
                            return binary.ReturnBlankOnZeroDenominator
                                ? SafeDivide(left, right, ZeroDenominatorBehavior.Blank)
                                : "(" + left + "/" + right + ")";
                        default: throw new NotSupportedException("The binary measure operator is not supported.");
                    }
                }
                case SafeDivideMeasureExpression divide:
                    return SafeDivide(
                        CompileExpression(divide.Numerator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        CompileExpression(divide.Denominator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        divide.OnZero);
                case RatioMeasureExpression ratio:
                    return SafeDivide(
                        CompileExpression(ratio.Numerator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        CompileExpression(ratio.Denominator, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        ratio.OnZero);
                case DifferenceMeasureExpression difference:
                {
                    var current = CompileExpression(difference.Current, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active);
                    var baseline = CompileExpression(difference.Baseline, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active);
                    switch (difference.DifferenceKind)
                    {
                        case DifferenceKind.Absolute: return "(" + current + "-" + baseline + ")";
                        case DifferenceKind.Percentage:
                            return SafeDivide("(" + current + "-" + baseline + ")", baseline, difference.OnZero);
                        case DifferenceKind.PercentagePoints:
                            return "(" + current + "-" + baseline + ")";
                        default: throw new NotSupportedException("The difference kind is not supported.");
                    }
                }
                case ShareMeasureExpression share:
                    IReadOnlyList<IReadOnlyList<PivotFilterItem>> denominatorMemberSets =
                        ResolveShareDenominatorMemberSets(
                            share.Scope,
                            memberFilterSets,
                            rowFieldOrder);
                    return SafeDivide(
                        CompileExpression(share.Part, ownerMeasureId, measures, pivot, memberFilterSets, periodFilterSets, rowFieldOrder, active),
                        CompileExpression(share.Whole, ownerMeasureId, measures, pivot, denominatorMemberSets, periodFilterSets, rowFieldOrder, active),
                        share.OnZero);
                default:
                    throw new NotSupportedException("The typed measure expression is not supported.");
            }
        }

        private static IReadOnlyList<IReadOnlyList<PivotFilterItem>> ResolveShareDenominatorMemberSets(
            ShareDenominatorScope scope,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyList<string> rowFieldOrder)
        {
            if (scope == ShareDenominatorScope.Explicit)
            {
                return memberFilterSets;
            }

            if (rowFieldOrder.Count == 0)
            {
                throw new InvalidOperationException(
                    "Share of parent or report total requires an explicit Rows hierarchy.");
            }

            var rowFields = new HashSet<string>(rowFieldOrder, StringComparer.OrdinalIgnoreCase);
            var result = new List<IReadOnlyList<PivotFilterItem>>();
            foreach (IReadOnlyList<PivotFilterItem> memberSet in memberFilterSets)
            {
                string? fieldToRemove = null;
                if (scope == ShareDenominatorScope.Parent)
                {
                    for (var index = rowFieldOrder.Count - 1; index >= 0; index--)
                    {
                        string candidate = rowFieldOrder[index];
                        if (memberSet.Any(filter => string.Equals(
                            filter.Field,
                            candidate,
                            StringComparison.OrdinalIgnoreCase)))
                        {
                            fieldToRemove = candidate;
                            break;
                        }
                    }
                }

                result.Add(memberSet
                    .Where(filter => scope == ShareDenominatorScope.FilteredReportTotal
                        ? !rowFields.Contains(filter.Field)
                        : !string.Equals(filter.Field, fieldToRemove, StringComparison.OrdinalIgnoreCase))
                    .ToList());
            }

            return result;
        }

        private string AggregateExpression(
            string ownerMeasureId,
            string field,
            AggregateFunction function,
            IReadOnlyList<MeasureFilterSpec> measureFilters,
            string? periodSliceId,
            PivotBuildResult pivot,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> periodFilterSets)
        {
            var descriptor = pivot.DataFields.FirstOrDefault(candidate =>
                string.Equals(candidate.MeasureId, ownerMeasureId, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(candidate.SourceField, field, StringComparison.OrdinalIgnoreCase) &&
                candidate.Function == function &&
                string.Equals(candidate.PeriodSliceId, periodSliceId, StringComparison.OrdinalIgnoreCase) &&
                FiltersEquivalent(candidate.Filters, measureFilters));
            if (descriptor == null)
            {
                throw new InvalidOperationException("The native PivotTable is missing a required aggregate component.");
            }

            var filterSets = ExpandFilters(measureFilters);
            if (filterSets.Count == 0)
            {
                filterSets.Add(new List<PivotFilterItem>());
            }

            IReadOnlyList<IReadOnlyList<PivotFilterItem>> sliceSets =
                new[] { (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>() };
            if (!string.IsNullOrWhiteSpace(periodSliceId))
            {
                if (!periodFilterSets.TryGetValue(periodSliceId!, out sliceSets))
                {
                    throw new InvalidOperationException("A measure references an unresolved period slice.");
                }
            }

            var termCount = checked((long)memberFilterSets.Count * sliceSets.Count * filterSets.Count);
            if (termCount > MaximumExpandedAggregateTerms)
            {
                throw new InvalidOperationException(
                    "The typed measure expands to too many PivotTable terms for one managed formula.");
            }

            var expressions = new List<string>();
            foreach (var memberSet in memberFilterSets)
            {
                foreach (var sliceSet in sliceSets)
                {
                    foreach (var filterSet in filterSets)
                    {
                        var filters = new List<PivotFilterItem>(memberSet);
                        filters.AddRange(sliceSet);
                        filters.AddRange(filterSet);
                        expressions.Add(pivotFormulaBuilder.BuildExpression(
                            descriptor.PivotCaption,
                            pivot.WorksheetName,
                            pivot.AnchorCell,
                            filters));
                    }
                }
            }

            return expressions.Count == 1
                ? expressions[0]
                : "(" + string.Join("+", expressions) + ")";
        }

        private static List<List<PivotFilterItem>> ExpandFilters(IReadOnlyList<MeasureFilterSpec> filters)
        {
            var sets = new List<List<PivotFilterItem>> { new List<PivotFilterItem>() };
            foreach (var filter in filters)
            {
                IReadOnlyList<ScalarValue> values;
                switch (filter.Operator)
                {
                    case MeasureFilterOperator.Equal:
                    case MeasureFilterOperator.In:
                        values = filter.Values;
                        break;
                    case MeasureFilterOperator.IsBlank:
                        values = new[] { ScalarValue.Null() };
                        break;
                    case MeasureFilterOperator.IsNotBlank:
                    case MeasureFilterOperator.NotEqual:
                    case MeasureFilterOperator.NotIn:
                    case MeasureFilterOperator.GreaterThan:
                    case MeasureFilterOperator.GreaterThanOrEqual:
                    case MeasureFilterOperator.LessThan:
                    case MeasureFilterOperator.LessThanOrEqual:
                        throw new NotSupportedException(
                            "This filtered aggregate requires a dedicated filtered pivot and cannot be rendered from the shared pivot.");
                    default:
                        throw new NotSupportedException("The measure filter operator is not supported.");
                }

                if (values.Count == 0)
                {
                    throw new InvalidOperationException("The measure filter requires at least one value.");
                }

                if ((long)sets.Count * values.Count > MaximumExpandedAggregateTerms)
                {
                    throw new InvalidOperationException(
                        "The filtered aggregate expands to too many bounded member combinations.");
                }

                var expanded = new List<List<PivotFilterItem>>();
                foreach (var existing in sets)
                {
                    foreach (var value in values)
                    {
                        var copy = new List<PivotFilterItem>(existing)
                        {
                            new PivotFilterItem { Field = filter.Field, Value = ScalarValueObject(value) }
                        };
                        expanded.Add(copy);
                    }
                }

                sets = expanded;
            }

            return sets;
        }

        private static string SafeDivide(string numerator, string denominator, ZeroDenominatorBehavior behavior)
        {
            switch (behavior)
            {
                case ZeroDenominatorBehavior.Blank:
                    return "IF(OR((" + denominator + ")=0,(" + denominator + ")=\"\"),\"\",(" + numerator + ")/(" + denominator + "))";
                case ZeroDenominatorBehavior.Zero:
                    return "IF(OR((" + denominator + ")=0,(" + denominator + ")=\"\"),0,(" + numerator + ")/(" + denominator + "))";
                case ZeroDenominatorBehavior.Error:
                    return "((" + numerator + ")/(" + denominator + "))";
                default:
                    throw new NotSupportedException("The zero-denominator behavior is not supported.");
            }
        }

        private static object? ScalarValueObject(ScalarValue value)
        {
            switch (value.Kind)
            {
                case ScalarValueKind.Null: return null;
                case ScalarValueKind.Text: return value.Text;
                case ScalarValueKind.Number: return value.Number;
                case ScalarValueKind.Boolean: return value.Boolean;
                case ScalarValueKind.Date:
                case ScalarValueKind.DateTime: return value.Temporal;
                default: throw new NotSupportedException("The scalar filter value is not supported.");
            }
        }

        private static bool FiltersEquivalent(
            IReadOnlyList<MeasureFilterSpec> left,
            IReadOnlyList<MeasureFilterSpec> right)
        {
            if (left.Count != right.Count)
            {
                return false;
            }

            return left.Select(CanonicalFilterKey)
                .OrderBy(value => value, StringComparer.Ordinal)
                .SequenceEqual(
                    right.Select(CanonicalFilterKey).OrderBy(value => value, StringComparer.Ordinal),
                    StringComparer.Ordinal);
        }

        private static string CanonicalFilterKey(MeasureFilterSpec filter)
        {
            return (filter.Field ?? string.Empty).ToUpperInvariant() + "|" + filter.Operator + "|" +
                string.Join(",", filter.Values.Select(ScalarKey).OrderBy(value => value, StringComparer.Ordinal));
        }

        private static string ScalarKey(ScalarValue value)
        {
            return value.Kind + ":" + (value.Text
                ?? (value.Number.HasValue ? value.Number.Value.ToString(CultureInfo.InvariantCulture) : null)
                ?? (value.Boolean.HasValue ? value.Boolean.Value.ToString() : null)
                ?? (value.Temporal.HasValue ? value.Temporal.Value.ToString("O", CultureInfo.InvariantCulture) : string.Empty));
        }

        private static bool RequiresErrorPropagation(
            MeasureExpression expression,
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            ISet<string> active)
        {
            switch (expression)
            {
                case WeightedAggregateMeasureExpression weighted:
                    return weighted.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(weighted.Numerator, measures, active) ||
                        RequiresErrorPropagation(weighted.Denominator, measures, active);
                case ReferenceMeasureExpression reference:
                    if (!measures.TryGetValue(reference.MeasureId, out var referenced) ||
                        !active.Add(reference.MeasureId))
                    {
                        return false;
                    }

                    try
                    {
                        return RequiresErrorPropagation(referenced.Expression, measures, active);
                    }
                    finally
                    {
                        active.Remove(reference.MeasureId);
                    }
                case BinaryMeasureExpression binary:
                    return (binary.Operator == BinaryMeasureOperator.Divide && !binary.ReturnBlankOnZeroDenominator) ||
                        RequiresErrorPropagation(binary.Left, measures, active) ||
                        RequiresErrorPropagation(binary.Right, measures, active);
                case SafeDivideMeasureExpression divide:
                    return divide.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(divide.Numerator, measures, active) ||
                        RequiresErrorPropagation(divide.Denominator, measures, active);
                case RatioMeasureExpression ratio:
                    return ratio.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(ratio.Numerator, measures, active) ||
                        RequiresErrorPropagation(ratio.Denominator, measures, active);
                case DifferenceMeasureExpression difference:
                    return (difference.DifferenceKind == DifferenceKind.Percentage &&
                            difference.OnZero == ZeroDenominatorBehavior.Error) ||
                        RequiresErrorPropagation(difference.Current, measures, active) ||
                        RequiresErrorPropagation(difference.Baseline, measures, active);
                case ShareMeasureExpression share:
                    return share.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(share.Part, measures, active) ||
                        RequiresErrorPropagation(share.Whole, measures, active);
                default:
                    return false;
            }
        }
    }
}
