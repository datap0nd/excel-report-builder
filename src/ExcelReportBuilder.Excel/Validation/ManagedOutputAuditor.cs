using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Validation
{
    internal sealed class ManagedFormulaCellAuditResult
    {
        public bool FormulaChanged { get; set; }

        public bool HasExcelError { get; set; }

        public bool ValueMatches { get; set; }
    }

    internal readonly struct ManagedBlockMeasureKey : IEquatable<ManagedBlockMeasureKey>
    {
        public ManagedBlockMeasureKey(string blockId, string measureId)
        {
            if (string.IsNullOrWhiteSpace(blockId))
            {
                throw new ArgumentException("A managed block identifier is required.", nameof(blockId));
            }

            if (string.IsNullOrWhiteSpace(measureId))
            {
                throw new ArgumentException("A measure identifier is required.", nameof(measureId));
            }

            BlockId = blockId;
            MeasureId = measureId;
        }

        public string BlockId { get; }

        public string MeasureId { get; }

        public bool Equals(ManagedBlockMeasureKey other)
        {
            return string.Equals(BlockId, other.BlockId, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(MeasureId, other.MeasureId, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj)
        {
            return obj is ManagedBlockMeasureKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (StringComparer.OrdinalIgnoreCase.GetHashCode(BlockId) * 397) ^
                       StringComparer.OrdinalIgnoreCase.GetHashCode(MeasureId);
            }
        }

        public override string ToString()
        {
            return BlockId + ":" + MeasureId;
        }
    }

    internal static class ManagedOutputAuditor
    {
        internal const decimal FormulaTolerance = 0.000001m;
        private const decimal FormulaRelativeTolerance = 0.000000000000001m;

        public static IReadOnlyDictionary<string, decimal> AggregateByMeasure(
            IReadOnlyDictionary<ManagedBlockMeasureKey, decimal> scopedTotals)
        {
            if (scopedTotals == null) throw new ArgumentNullException(nameof(scopedTotals));

            var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var total in scopedTotals)
            {
                result[total.Key.MeasureId] = result.TryGetValue(total.Key.MeasureId, out var existing)
                    ? checked(existing + total.Value)
                    : total.Value;
            }

            return result;
        }

        public static IReadOnlyList<CheckResult> ReconcileBlockTotals(
            IEnumerable<ManagedBlockMeasureKey> requiredTotals,
            IReadOnlyDictionary<ManagedBlockMeasureKey, decimal> pivotTotals,
            IReadOnlyDictionary<ManagedBlockMeasureKey, decimal> outputTotals,
            decimal tolerance)
        {
            if (requiredTotals == null) throw new ArgumentNullException(nameof(requiredTotals));
            if (pivotTotals == null) throw new ArgumentNullException(nameof(pivotTotals));
            if (outputTotals == null) throw new ArgumentNullException(nameof(outputTotals));
            if (tolerance < 0m) throw new ArgumentOutOfRangeException(nameof(tolerance));

            var results = new List<CheckResult>();
            foreach (var key in requiredTotals
                         .Distinct()
                         .OrderBy(item => item.BlockId, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(item => item.MeasureId, StringComparer.OrdinalIgnoreCase))
            {
                var hasPivot = pivotTotals.TryGetValue(key, out var expected);
                var hasOutput = outputTotals.TryGetValue(key, out var actual);
                var passed = hasPivot && hasOutput && AreClose(expected, actual, tolerance);
                results.Add(new CheckResult
                {
                    CheckId = "mandatory-block-total:" + key.BlockId + ":" + key.MeasureId,
                    Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                    Message = passed
                        ? "The managed output total reconciles with its PivotTable for block '" +
                          key.BlockId + "' and Value '" + key.MeasureId + "'."
                        : !hasPivot
                            ? "The PivotTable total is unavailable for managed block '" +
                              key.BlockId + "' and Value '" + key.MeasureId + "'."
                            : !hasOutput
                                ? "The output total is unavailable for managed block '" +
                                  key.BlockId + "' and Value '" + key.MeasureId + "'."
                                : "The managed output total does not reconcile with its PivotTable for block '" +
                                  key.BlockId + "' and Value '" + key.MeasureId + "'.",
                    Expected = hasPivot ? expected : (decimal?)null,
                    Actual = hasOutput ? actual : (decimal?)null
                });
            }

            return results;
        }

        public static bool FormulaValueMatches(
            DenseFormulaExpectation expectation,
            object? actualValue,
            bool hasExcelError,
            decimal tolerance = FormulaTolerance)
        {
            if (expectation == null) throw new ArgumentNullException(nameof(expectation));
            if (tolerance < 0m) throw new ArgumentOutOfRangeException(nameof(tolerance));

            switch (expectation.Kind)
            {
                case DenseFormulaExpectationKind.Error:
                    return hasExcelError;
                case DenseFormulaExpectationKind.Blank:
                    return !hasExcelError && IsBlank(actualValue);
                case DenseFormulaExpectationKind.Number:
                    if (hasExcelError || IsBlank(actualValue) || !expectation.NumericValue.HasValue)
                    {
                        return false;
                    }

                    try
                    {
                        var actual = Convert.ToDecimal(actualValue, CultureInfo.InvariantCulture);
                        return FormulaValuesAreClose(
                            expectation.NumericValue.Value,
                            actual,
                            tolerance);
                    }
                    catch (Exception exception) when (
                        exception is FormatException ||
                        exception is InvalidCastException ||
                        exception is OverflowException)
                    {
                        return false;
                    }
                default:
                    throw new NotSupportedException("The dense formula expectation kind is not supported.");
            }
        }

        public static ManagedFormulaCellAuditResult AuditFormulaCell(
            SafeExcelFormula expectedFormula,
            DenseFormulaExpectation expectedValue,
            string actualFormula,
            string displayedValue,
            object? actualValue,
            decimal tolerance = FormulaTolerance)
        {
            if (expectedFormula == null) throw new ArgumentNullException(nameof(expectedFormula));
            if (expectedValue == null) throw new ArgumentNullException(nameof(expectedValue));

            var hasExcelError = IsExcelErrorDisplay(displayedValue);
            return new ManagedFormulaCellAuditResult
            {
                FormulaChanged = !string.Equals(
                    actualFormula ?? string.Empty,
                    expectedFormula.Value,
                    StringComparison.Ordinal),
                HasExcelError = hasExcelError,
                ValueMatches = FormulaValueMatches(
                    expectedValue,
                    actualValue,
                    hasExcelError,
                    tolerance)
            };
        }

        private static bool IsBlank(object? value)
        {
            return value == null ||
                   string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.InvariantCulture));
        }

        internal static bool IsExcelErrorDisplay(string? displayedValue)
        {
            var value = (displayedValue ?? string.Empty).Trim();
            return value.Length > 1 &&
                   value[0] == '#' &&
                   value.Any(character => character != '#');
        }

        private static bool AreClose(decimal expected, decimal actual, decimal tolerance)
        {
            return Math.Abs(expected - actual) <= tolerance;
        }

        private static bool FormulaValuesAreClose(
            decimal expected,
            decimal actual,
            decimal absoluteTolerance)
        {
            var magnitude = expected == decimal.MinValue ? decimal.MaxValue : Math.Abs(expected);
            var scaledTolerance = magnitude * FormulaRelativeTolerance;
            return AreClose(expected, actual, Math.Max(absoluteTolerance, scaledTolerance));
        }
    }
}
