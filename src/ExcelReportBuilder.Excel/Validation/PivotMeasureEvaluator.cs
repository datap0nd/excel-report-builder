using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Validation
{
    /// <summary>
    /// Independently evaluates the typed measure graph from native PivotTable
    /// aggregate inputs. Formula text is never inspected or executed here.
    /// </summary>
    internal sealed class PivotMeasureEvaluator
    {
        private readonly IReadOnlyDictionary<string, MeasureDefinition> measures;
        private readonly PivotBuildResult pivot;
        private readonly Func<string, IReadOnlyList<PivotFilterItem>, DenseFormulaExpectation> aggregateReader;
        private readonly Dictionary<string, DenseFormulaExpectation> aggregateCache =
            new Dictionary<string, DenseFormulaExpectation>(StringComparer.Ordinal);

        public PivotMeasureEvaluator(
            IReadOnlyDictionary<string, MeasureDefinition> measures,
            PivotBuildResult pivot,
            Func<string, IReadOnlyList<PivotFilterItem>, DenseFormulaExpectation> aggregateReader)
        {
            this.measures = measures ?? throw new ArgumentNullException(nameof(measures));
            this.pivot = pivot ?? throw new ArgumentNullException(nameof(pivot));
            this.aggregateReader = aggregateReader ?? throw new ArgumentNullException(nameof(aggregateReader));
        }

        public DenseFormulaExpectation EvaluateAcrossMemberSets(
            string measureId,
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

            var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { measureId };
            var result = EvaluateExpression(
                definition.Expression,
                measureId,
                memberFilterSets,
                periodFilterSets,
                rowFieldOrder ?? Array.Empty<string>(),
                active);
            if (result.Kind == DenseFormulaExpectationKind.Error &&
                !RequiresErrorPropagation(
                    definition.Expression,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase) { measureId }))
            {
                return DenseFormulaExpectation.Blank();
            }

            return result;
        }

        private DenseFormulaExpectation EvaluateExpression(
            MeasureExpression expression,
            string ownerMeasureId,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> periodFilterSets,
            IReadOnlyList<string> rowFieldOrder,
            ISet<string> active)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    return EvaluateAggregate(
                        ownerMeasureId,
                        aggregate.Field,
                        aggregate.Function,
                        Array.Empty<MeasureFilterSpec>(),
                        aggregate.PeriodSliceId,
                        memberFilterSets,
                        periodFilterSets);
                case FilteredAggregateMeasureExpression filtered:
                    return EvaluateAggregate(
                        ownerMeasureId,
                        filtered.Field,
                        filtered.Function,
                        filtered.Filters,
                        filtered.PeriodSliceId,
                        memberFilterSets,
                        periodFilterSets);
                case WeightedAggregateMeasureExpression weighted:
                    return EvaluateDivision(
                        () => EvaluateExpression(
                            weighted.Numerator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        () => EvaluateExpression(
                            weighted.Denominator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
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
                        return EvaluateExpression(
                            referenced.Expression,
                            ownerMeasureId,
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
                    return DenseFormulaExpectation.Number(constant.Value);
                case BinaryMeasureExpression binary:
                    if (binary.Operator == BinaryMeasureOperator.Divide)
                    {
                        return EvaluateDivision(
                            () => EvaluateExpression(
                                binary.Left,
                                ownerMeasureId,
                                memberFilterSets,
                                periodFilterSets,
                                rowFieldOrder,
                                active),
                            () => EvaluateExpression(
                                binary.Right,
                                ownerMeasureId,
                                memberFilterSets,
                                periodFilterSets,
                                rowFieldOrder,
                                active),
                            binary.ReturnBlankOnZeroDenominator
                                ? ZeroDenominatorBehavior.Blank
                                : ZeroDenominatorBehavior.Error);
                    }

                    return EvaluateBinary(
                        EvaluateExpression(
                            binary.Left,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        EvaluateExpression(
                            binary.Right,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        binary.Operator);
                case SafeDivideMeasureExpression divide:
                    return EvaluateDivision(
                        () => EvaluateExpression(
                            divide.Numerator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        () => EvaluateExpression(
                            divide.Denominator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        divide.OnZero);
                case RatioMeasureExpression ratio:
                    return EvaluateDivision(
                        () => EvaluateExpression(
                            ratio.Numerator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        () => EvaluateExpression(
                            ratio.Denominator,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        ratio.OnZero);
                case DifferenceMeasureExpression difference:
                    return EvaluateDifference(
                        difference,
                        ownerMeasureId,
                        memberFilterSets,
                        periodFilterSets,
                        rowFieldOrder,
                        active);
                case ShareMeasureExpression share:
                    var denominatorMemberSets = ResolveShareDenominatorMemberSets(
                        share.Scope,
                        memberFilterSets,
                        rowFieldOrder);
                    return EvaluateDivision(
                        () => EvaluateExpression(
                            share.Part,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        () => EvaluateExpression(
                            share.Whole,
                            ownerMeasureId,
                            denominatorMemberSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active),
                        share.OnZero);
                default:
                    throw new NotSupportedException(
                        "The typed measure expression cannot be independently evaluated.");
            }
        }

        private DenseFormulaExpectation EvaluateDifference(
            DifferenceMeasureExpression difference,
            string ownerMeasureId,
            IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberFilterSets,
            IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>> periodFilterSets,
            IReadOnlyList<string> rowFieldOrder,
            ISet<string> active)
        {
            if (difference.DifferenceKind == DifferenceKind.Percentage)
            {
                return EvaluateDivision(
                    () =>
                    {
                        var current = EvaluateExpression(
                            difference.Current,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active);
                        var baseline = EvaluateExpression(
                            difference.Baseline,
                            ownerMeasureId,
                            memberFilterSets,
                            periodFilterSets,
                            rowFieldOrder,
                            active);
                        return EvaluateBinary(current, baseline, BinaryMeasureOperator.Subtract);
                    },
                    () => EvaluateExpression(
                        difference.Baseline,
                        ownerMeasureId,
                        memberFilterSets,
                        periodFilterSets,
                        rowFieldOrder,
                        active),
                    difference.OnZero);
            }

            if (difference.DifferenceKind != DifferenceKind.Absolute &&
                difference.DifferenceKind != DifferenceKind.PercentagePoints)
            {
                throw new NotSupportedException("The difference kind is not supported.");
            }

            return EvaluateBinary(
                EvaluateExpression(
                    difference.Current,
                    ownerMeasureId,
                    memberFilterSets,
                    periodFilterSets,
                    rowFieldOrder,
                    active),
                EvaluateExpression(
                    difference.Baseline,
                    ownerMeasureId,
                    memberFilterSets,
                    periodFilterSets,
                    rowFieldOrder,
                    active),
                BinaryMeasureOperator.Subtract);
        }

        private DenseFormulaExpectation EvaluateAggregate(
            string ownerMeasureId,
            string field,
            AggregateFunction function,
            IReadOnlyList<MeasureFilterSpec> measureFilters,
            string? periodSliceId,
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
                throw new InvalidOperationException(
                    "The native PivotTable is missing a required aggregate component for independent validation.");
            }

            var filterSets = ExpandFilters(measureFilters);
            if (filterSets.Count == 0)
            {
                filterSets.Add(new List<PivotFilterItem>());
            }

            IReadOnlyList<IReadOnlyList<PivotFilterItem>> sliceSets =
                new[] { (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>() };
            if (!string.IsNullOrWhiteSpace(periodSliceId) &&
                !periodFilterSets.TryGetValue(periodSliceId!, out sliceSets))
            {
                throw new InvalidOperationException(
                    "A measure references an unresolved period slice during independent validation.");
            }

            var termCount = checked((long)memberFilterSets.Count * sliceSets.Count * filterSets.Count);
            if (termCount > MeasureFormulaCompiler.MaximumExpandedAggregateTerms)
            {
                throw new InvalidOperationException(
                    "The typed measure expands to too many PivotTable terms for independent validation.");
            }

            var total = 0m;
            try
            {
                foreach (var memberSet in memberFilterSets)
                {
                    foreach (var sliceSet in sliceSets)
                    {
                        foreach (var filterSet in filterSets)
                        {
                            var filters = new List<PivotFilterItem>(memberSet);
                            filters.AddRange(sliceSet);
                            filters.AddRange(filterSet);
                            var value = ReadAggregate(descriptor.PivotCaption, filters);
                            if (value.Kind == DenseFormulaExpectationKind.Error)
                            {
                                throw new InvalidOperationException(
                                    "A PivotTable aggregate could not be read independently.");
                            }

                            if (value.Kind == DenseFormulaExpectationKind.Number)
                            {
                                total = checked(total + value.NumericValue!.Value);
                            }
                        }
                    }
                }
            }
            catch (OverflowException)
            {
                return DenseFormulaExpectation.Error();
            }

            return DenseFormulaExpectation.Number(total);
        }

        private DenseFormulaExpectation ReadAggregate(
            string pivotCaption,
            IReadOnlyList<PivotFilterItem> filters)
        {
            var key = AggregateCacheKey(pivotCaption, filters);
            if (!aggregateCache.TryGetValue(key, out var value))
            {
                value = aggregateReader(pivotCaption, filters) ??
                    throw new InvalidOperationException(
                        "The PivotTable aggregate reader returned no independent result.");
                aggregateCache.Add(key, value);
            }

            return value;
        }

        private static DenseFormulaExpectation EvaluateBinary(
            DenseFormulaExpectation left,
            DenseFormulaExpectation right,
            BinaryMeasureOperator operation)
        {
            if (left.Kind == DenseFormulaExpectationKind.Error ||
                right.Kind == DenseFormulaExpectationKind.Error)
            {
                return DenseFormulaExpectation.Error();
            }

            var leftNumber = ExcelArithmeticValue(left);
            var rightNumber = ExcelArithmeticValue(right);
            try
            {
                switch (operation)
                {
                    case BinaryMeasureOperator.Add:
                        return DenseFormulaExpectation.Number(checked(leftNumber + rightNumber));
                    case BinaryMeasureOperator.Subtract:
                        return DenseFormulaExpectation.Number(checked(leftNumber - rightNumber));
                    case BinaryMeasureOperator.Multiply:
                        return DenseFormulaExpectation.Number(checked(leftNumber * rightNumber));
                    default:
                        throw new NotSupportedException("The binary measure operator is not supported.");
                }
            }
            catch (OverflowException)
            {
                return DenseFormulaExpectation.Error();
            }
        }

        private static DenseFormulaExpectation EvaluateDivision(
            Func<DenseFormulaExpectation> numeratorFactory,
            Func<DenseFormulaExpectation> denominatorFactory,
            ZeroDenominatorBehavior behavior)
        {
            var denominator = denominatorFactory();
            if (denominator.Kind == DenseFormulaExpectationKind.Error)
            {
                return denominator;
            }

            if (denominator.Kind == DenseFormulaExpectationKind.Blank ||
                denominator.NumericValue.GetValueOrDefault() == 0m)
            {
                switch (behavior)
                {
                    case ZeroDenominatorBehavior.Blank:
                        return DenseFormulaExpectation.Blank();
                    case ZeroDenominatorBehavior.Zero:
                        return DenseFormulaExpectation.Number(0m);
                    case ZeroDenominatorBehavior.Error:
                        return DenseFormulaExpectation.Error();
                    default:
                        throw new NotSupportedException(
                            "The zero-denominator behavior is not supported.");
                }
            }

            var numerator = numeratorFactory();
            if (numerator.Kind == DenseFormulaExpectationKind.Error)
            {
                return numerator;
            }

            try
            {
                return DenseFormulaExpectation.Number(
                    ExcelArithmeticValue(numerator) / denominator.NumericValue!.Value);
            }
            catch (DivideByZeroException)
            {
                return DenseFormulaExpectation.Error();
            }
            catch (OverflowException)
            {
                return DenseFormulaExpectation.Error();
            }
        }

        private static decimal ExcelArithmeticValue(DenseFormulaExpectation value)
        {
            return value.Kind == DenseFormulaExpectationKind.Number
                ? value.NumericValue!.Value
                : 0m;
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
            foreach (var memberSet in memberFilterSets)
            {
                string? fieldToRemove = null;
                if (scope == ShareDenominatorScope.Parent)
                {
                    for (var index = rowFieldOrder.Count - 1; index >= 0; index--)
                    {
                        var candidate = rowFieldOrder[index];
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
                            "This filtered aggregate cannot be independently evaluated from the shared pivot.");
                    default:
                        throw new NotSupportedException("The measure filter operator is not supported.");
                }

                if (values.Count == 0)
                {
                    throw new InvalidOperationException("The measure filter requires at least one value.");
                }

                if ((long)sets.Count * values.Count > MeasureFormulaCompiler.MaximumExpandedAggregateTerms)
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
                ?? (value.Temporal.HasValue
                    ? value.Temporal.Value.ToString("O", CultureInfo.InvariantCulture)
                    : string.Empty));
        }

        private static string AggregateCacheKey(
            string pivotCaption,
            IReadOnlyList<PivotFilterItem> filters)
        {
            return LengthPrefixed(pivotCaption.ToUpperInvariant()) + string.Concat(
                filters.Select(FilterItemKey)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .Select(LengthPrefixed));
        }

        private static string FilterItemKey(PivotFilterItem filter)
        {
            var value = filter.Value;
            string canonical;
            if (value == null)
            {
                canonical = "NULL";
            }
            else if (value is DateTime date)
            {
                canonical = "DATE:" + date.ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is IFormattable formattable)
            {
                canonical = value.GetType().FullName + ":" +
                    formattable.ToString(null, CultureInfo.InvariantCulture);
            }
            else
            {
                canonical = value.GetType().FullName + ":" + value;
            }

            return LengthPrefixed((filter.Field ?? string.Empty).ToUpperInvariant()) +
                   LengthPrefixed(canonical);
        }

        private static string LengthPrefixed(string value)
        {
            return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
        }

        private bool RequiresErrorPropagation(MeasureExpression expression, ISet<string> active)
        {
            switch (expression)
            {
                case WeightedAggregateMeasureExpression weighted:
                    return weighted.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(weighted.Numerator, active) ||
                        RequiresErrorPropagation(weighted.Denominator, active);
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
                        return RequiresErrorPropagation(referenced.Expression, active);
                    }
                    finally
                    {
                        active.Remove(reference.MeasureId);
                    }
                case BinaryMeasureExpression binary:
                    return (binary.Operator == BinaryMeasureOperator.Divide &&
                            !binary.ReturnBlankOnZeroDenominator) ||
                        RequiresErrorPropagation(binary.Left, active) ||
                        RequiresErrorPropagation(binary.Right, active);
                case SafeDivideMeasureExpression divide:
                    return divide.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(divide.Numerator, active) ||
                        RequiresErrorPropagation(divide.Denominator, active);
                case RatioMeasureExpression ratio:
                    return ratio.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(ratio.Numerator, active) ||
                        RequiresErrorPropagation(ratio.Denominator, active);
                case DifferenceMeasureExpression difference:
                    return (difference.DifferenceKind == DifferenceKind.Percentage &&
                            difference.OnZero == ZeroDenominatorBehavior.Error) ||
                        RequiresErrorPropagation(difference.Current, active) ||
                        RequiresErrorPropagation(difference.Baseline, active);
                case ShareMeasureExpression share:
                    return share.OnZero == ZeroDenominatorBehavior.Error ||
                        RequiresErrorPropagation(share.Part, active) ||
                        RequiresErrorPropagation(share.Whole, active);
                default:
                    return false;
            }
        }
    }
}
