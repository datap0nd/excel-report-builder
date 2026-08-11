using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Validation
{
    public enum CheckOutcome
    {
        Passed,
        Failed,
        Warning
    }

    public sealed class CheckResult
    {
        public string CheckId { get; set; } = string.Empty;

        public CheckOutcome Outcome { get; set; }

        public string Message { get; set; } = string.Empty;

        public decimal? Expected { get; set; }

        public decimal? Actual { get; set; }

    }

    public sealed class ReconciliationSnapshot
    {
        public long SourceRows { get; set; }

        public long ProjectedNormalizedRows { get; set; }

        /// <summary>
        /// Exact final row count produced by independently applying the closed
        /// typed transform grammar to the complete current source. Required
        /// whenever planning requires independent post-transform evidence.
        /// </summary>
        public long? ExpectedPostTransformNormalizedRows { get; set; }

        public long ActualNormalizedRows { get; set; }

        public IReadOnlyDictionary<string, decimal> SourceTotals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.Ordinal);

        public IReadOnlyDictionary<string, decimal> NormalizedTotals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.Ordinal);

        public IReadOnlyDictionary<string, decimal> PivotTotals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.Ordinal);

        public IReadOnlyDictionary<string, decimal> OutputTotals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.Ordinal);

        public IReadOnlyDictionary<string, decimal> OutputMinimums { get; set; } =
            new Dictionary<string, decimal>(StringComparer.Ordinal);

        public IReadOnlyDictionary<string, long> MissingRequiredValues { get; set; } =
            new Dictionary<string, long>(StringComparer.Ordinal);
    }

    public sealed class ReportReconciler
    {
        public IReadOnlyList<CheckResult> Reconcile(ReconciliationSnapshot snapshot, decimal tolerance)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            if (tolerance < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(tolerance));
            }

            var checks = new List<CheckResult>();
            checks.Add(new CheckResult
            {
                CheckId = "normalized-row-count",
                Outcome = snapshot.ProjectedNormalizedRows == snapshot.ActualNormalizedRows
                    ? CheckOutcome.Passed
                    : CheckOutcome.Failed,
                Message = snapshot.ProjectedNormalizedRows == snapshot.ActualNormalizedRows
                    ? "Normalized row count matches the projection."
                    : "Normalized row count differs from the projection.",
                Expected = snapshot.ProjectedNormalizedRows,
                Actual = snapshot.ActualNormalizedRows
            });

            foreach (var sourceTotal in snapshot.SourceTotals)
            {
                AddComparison(checks, sourceTotal.Key, "source-to-normalized", sourceTotal.Value, snapshot.NormalizedTotals, tolerance);
                AddComparison(checks, sourceTotal.Key, "source-to-pivot", sourceTotal.Value, snapshot.PivotTotals, tolerance);
                AddComparison(checks, sourceTotal.Key, "source-to-output", sourceTotal.Value, snapshot.OutputTotals, tolerance);
            }

            return checks;
        }

        public IReadOnlyList<CheckResult> Reconcile(
            ReconciliationSnapshot snapshot,
            IReadOnlyList<BuildCheckPlan> plan)
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (plan == null) throw new ArgumentNullException(nameof(plan));

            var results = new List<CheckResult>();
            foreach (var check in plan)
            {
                switch (check.Kind)
                {
                    case ReportCheckKind.NoTruncation:
                        results.Add(RowCountResult(check, snapshot));
                        break;
                    case ReportCheckKind.TotalPreservation:
                    {
                        if (check.EvaluationScope == CheckEvaluationScope.RenderedOutput)
                        {
                            AddRenderedOutputComparisons(results, check, snapshot);
                            break;
                        }

                        var measures = string.IsNullOrWhiteSpace(check.MeasureId)
                            ? snapshot.SourceTotals.Keys.OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList()
                            : new List<string> { check.MeasureId! };
                        if (measures.Count == 0)
                        {
                            results.Add(new CheckResult
                            {
                                CheckId = check.Id,
                                Outcome = CheckOutcome.Passed,
                                Message = "No additive Values are configured. Total preservation is not applicable; the mandatory row-preservation check remains in force."
                            });
                            break;
                        }

                        foreach (var measure in measures)
                        {
                            if (!snapshot.SourceTotals.TryGetValue(measure, out var expected))
                            {
                                results.Add(Unavailable(check.Id + ":" + measure, "The source total is unavailable."));
                                continue;
                            }

                            AddComparison(results, measure, check.Id + ":source-to-normalized", expected, snapshot.NormalizedTotals, check.Tolerance);
                        }

                        break;
                    }
                    case ReportCheckKind.Balance:
                        results.Add(BalanceResult(check, snapshot));
                        break;
                    case ReportCheckKind.NonNegative:
                        results.Add(NonNegativeResult(check, snapshot));
                        break;
                    case ReportCheckKind.RequiredValues:
                        results.Add(RequiredValuesResult(check, snapshot));
                        break;
                    default:
                        results.Add(Unavailable(check.Id, "The configured check kind is not supported."));
                        break;
                }
            }

            return results;
        }

        private static void AddRenderedOutputComparisons(
            ICollection<CheckResult> results,
            BuildCheckPlan check,
            ReconciliationSnapshot snapshot)
        {
            var measures = string.IsNullOrWhiteSpace(check.MeasureId)
                ? snapshot.PivotTotals.Keys
                    .Where(snapshot.OutputTotals.ContainsKey)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList()
                : new List<string> { check.MeasureId! };
            if (measures.Count == 0)
            {
                results.Add(new CheckResult
                {
                    CheckId = check.Id,
                    Outcome = CheckOutcome.Passed,
                    Message = "No directly comparable aggregate Values are configured. Calculated Values remain covered by managed formula integrity checks."
                });
                return;
            }

            foreach (string measure in measures)
            {
                if (!snapshot.PivotTotals.TryGetValue(measure, out decimal expected))
                {
                    results.Add(Unavailable(check.Id + ":" + measure, "The filtered PivotTable total is unavailable."));
                    continue;
                }

                AddComparison(
                    results,
                    measure,
                    check.Id + ":pivot-to-output",
                    expected,
                    snapshot.OutputTotals,
                    check.Tolerance);
            }
        }

        private static CheckResult RowCountResult(BuildCheckPlan check, ReconciliationSnapshot snapshot)
        {
            var requiresIndependentCount = check.RowCountExpectation ==
                RowCountExpectation.ExactPostTransformCount ||
                check.RowCountExpectation == RowCountExpectation.AtMostProjection;
            if (requiresIndependentCount && !snapshot.ExpectedPostTransformNormalizedRows.HasValue)
            {
                return Unavailable(
                    check.Id,
                    "The exact independently audited post-transform row count is unavailable.");
            }

            var expected = requiresIndependentCount
                ? snapshot.ExpectedPostTransformNormalizedRows!.Value
                : snapshot.ProjectedNormalizedRows;
            var passed = expected == snapshot.ActualNormalizedRows;
            return new CheckResult
            {
                CheckId = check.Id,
                Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = passed
                    ? requiresIndependentCount
                        ? "Normalized row count matches the independently audited post-transform count."
                        : "Normalized row count matches the projection."
                    : requiresIndependentCount
                        ? "Normalized row count differs from the independently audited post-transform count."
                        : "Normalized row count differs from the projection.",
                Expected = expected,
                Actual = snapshot.ActualNormalizedRows
            };
        }

        private static CheckResult BalanceResult(BuildCheckPlan check, ReconciliationSnapshot snapshot)
        {
            if (string.IsNullOrWhiteSpace(check.MeasureId) || string.IsNullOrWhiteSpace(check.ComparedMeasureId) ||
                !snapshot.OutputTotals.TryGetValue(check.MeasureId!, out var left) ||
                !snapshot.OutputTotals.TryGetValue(check.ComparedMeasureId!, out var right))
            {
                return Unavailable(check.Id, "The configured balance totals are unavailable.");
            }

            var passed = Math.Abs(left - right) <= check.Tolerance;
            return new CheckResult
            {
                CheckId = check.Id,
                Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = passed ? "The configured balance reconciles." : "The configured balance does not reconcile.",
                Expected = right,
                Actual = left
            };
        }

        private static CheckResult NonNegativeResult(BuildCheckPlan check, ReconciliationSnapshot snapshot)
        {
            if (string.IsNullOrWhiteSpace(check.MeasureId) ||
                !snapshot.OutputMinimums.TryGetValue(check.MeasureId!, out var minimum))
            {
                return Unavailable(check.Id, "The output minimum is unavailable for the non-negative check.");
            }

            var passed = minimum >= -check.Tolerance;
            return new CheckResult
            {
                CheckId = check.Id,
                Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = passed ? "All checked output values are non-negative." : "A checked output value is negative.",
                Expected = 0m,
                Actual = minimum
            };
        }

        private static CheckResult RequiredValuesResult(BuildCheckPlan check, ReconciliationSnapshot snapshot)
        {
            if (string.IsNullOrWhiteSpace(check.MeasureId) ||
                !snapshot.MissingRequiredValues.TryGetValue(check.MeasureId!, out var missing))
            {
                return Unavailable(check.Id, "The missing-value count is unavailable for the required-values check.");
            }

            return new CheckResult
            {
                CheckId = check.Id,
                Outcome = missing == 0 ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = missing == 0
                    ? "All required output values are present."
                    : "One or more required output values are missing.",
                Expected = 0m,
                Actual = missing
            };
        }

        private static CheckResult Unavailable(string checkId, string message)
        {
            return new CheckResult
            {
                CheckId = checkId,
                Outcome = CheckOutcome.Failed,
                Message = message
            };
        }

        private static void AddComparison(
            ICollection<CheckResult> checks,
            string measure,
            string scope,
            decimal expected,
            IReadOnlyDictionary<string, decimal> actuals,
            decimal tolerance)
        {
            if (!actuals.TryGetValue(measure, out var actual))
            {
                checks.Add(new CheckResult
                {
                    CheckId = scope + ":" + measure,
                    Outcome = CheckOutcome.Failed,
                    Message = "The " + scope + " check has no value for " + measure + ".",
                    Expected = expected
                });
                return;
            }

            var passed = Math.Abs(expected - actual) <= tolerance;
            checks.Add(new CheckResult
            {
                CheckId = scope + ":" + measure,
                Outcome = passed ? CheckOutcome.Passed : CheckOutcome.Failed,
                Message = passed
                    ? "The " + scope + " total reconciles for " + measure + "."
                    : "The " + scope + " total does not reconcile for " + measure + ".",
                Expected = expected,
                Actual = actual
            });
        }
    }
}
