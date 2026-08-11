using ExcelReportBuilder.Excel.Validation;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ReconciliationTests
{
    [Fact]
    public void Reconciler_checks_every_stage_and_row_expansion()
    {
        var reconciler = new ReportReconciler();
        var totals = new Dictionary<string, decimal> { ["value"] = 120m };
        var snapshot = new ReconciliationSnapshot
        {
            SourceRows = 10,
            ProjectedNormalizedRows = 20,
            ActualNormalizedRows = 20,
            SourceTotals = totals,
            NormalizedTotals = totals,
            PivotTotals = totals,
            OutputTotals = totals
        };

        var checks = reconciler.Reconcile(snapshot, 0.001m);

        Assert.Equal(4, checks.Count);
        Assert.All(checks, check => Assert.Equal(CheckOutcome.Passed, check.Outcome));
    }

    [Fact]
    public void Reconciler_fails_missing_and_changed_totals()
    {
        var reconciler = new ReportReconciler();
        var snapshot = new ReconciliationSnapshot
        {
            ProjectedNormalizedRows = 12,
            ActualNormalizedRows = 11,
            SourceTotals = new Dictionary<string, decimal> { ["value"] = 100m },
            NormalizedTotals = new Dictionary<string, decimal> { ["value"] = 99m },
            PivotTotals = new Dictionary<string, decimal>(),
            OutputTotals = new Dictionary<string, decimal> { ["value"] = 100m }
        };

        var checks = reconciler.Reconcile(snapshot, 0.01m);

        Assert.Equal(3, checks.Count(check => check.Outcome == CheckOutcome.Failed));
    }

    [Fact]
    public void Planned_checks_use_their_own_tolerances_and_never_disappear()
    {
        var snapshot = new ReconciliationSnapshot
        {
            ProjectedNormalizedRows = 10,
            ActualNormalizedRows = 10,
            SourceTotals = new Dictionary<string, decimal> { ["amount"] = 100m },
            NormalizedTotals = new Dictionary<string, decimal> { ["amount"] = 100.01m },
            PivotTotals = new Dictionary<string, decimal> { ["amount"] = 100.01m },
            OutputTotals = new Dictionary<string, decimal> { ["amount"] = 100.01m }
        };
        var plan = new[]
        {
            new BuildCheckPlan { Id = "rows", Kind = ReportCheckKind.NoTruncation, Mandatory = true },
            new BuildCheckPlan
            {
                Id = "totals",
                Kind = ReportCheckKind.TotalPreservation,
                MeasureId = "amount",
                Tolerance = 0.02m,
                Mandatory = true
            },
            new BuildCheckPlan
            {
                Id = "required",
                Kind = ReportCheckKind.RequiredValues,
                MeasureId = "amount"
            }
        };

        var results = new ReportReconciler().Reconcile(snapshot, plan);

        Assert.Equal(3, results.Count);
        Assert.Equal(2, results.Count(result => result.Outcome == CheckOutcome.Passed));
        Assert.Contains(results, result => result.CheckId == "required" && result.Outcome == CheckOutcome.Failed);
    }

    [Fact]
    public void Total_preservation_is_not_applicable_when_no_additive_value_exists()
    {
        var snapshot = new ReconciliationSnapshot
        {
            ProjectedNormalizedRows = 10,
            ActualNormalizedRows = 10
        };
        var plan = new[]
        {
            new BuildCheckPlan { Id = "rows", Kind = ReportCheckKind.NoTruncation, Mandatory = true },
            new BuildCheckPlan { Id = "totals", Kind = ReportCheckKind.TotalPreservation, Mandatory = true }
        };

        var results = new ReportReconciler().Reconcile(snapshot, plan);

        Assert.All(results, result => Assert.Equal(CheckOutcome.Passed, result.Outcome));
        Assert.Contains(results, result =>
            result.CheckId == "totals" && result.Message.Contains("not applicable", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Planned_canonical_check_does_not_compare_unfiltered_source_to_filtered_output()
    {
        var snapshot = new ReconciliationSnapshot
        {
            SourceTotals = new Dictionary<string, decimal> { ["amount"] = 100m },
            NormalizedTotals = new Dictionary<string, decimal> { ["amount"] = 100m },
            PivotTotals = new Dictionary<string, decimal> { ["amount"] = 40m },
            OutputTotals = new Dictionary<string, decimal> { ["amount"] = 40m }
        };
        var plan = new[]
        {
            new BuildCheckPlan
            {
                Id = "canonical",
                Kind = ReportCheckKind.TotalPreservation,
                EvaluationScope = CheckEvaluationScope.CanonicalData
            },
            new BuildCheckPlan
            {
                Id = "rendered",
                Kind = ReportCheckKind.TotalPreservation,
                EvaluationScope = CheckEvaluationScope.RenderedOutput
            }
        };

        var results = new ReportReconciler().Reconcile(snapshot, plan);

        Assert.Equal(2, results.Count);
        Assert.All(results, result => Assert.Equal(CheckOutcome.Passed, result.Outcome));
        Assert.Contains(results, result => result.CheckId.Contains("source-to-normalized", StringComparison.Ordinal));
        Assert.Contains(results, result => result.CheckId.Contains("pivot-to-output", StringComparison.Ordinal));
    }

    [Fact]
    public void Rendered_output_check_fails_when_filtered_pivot_and_output_differ()
    {
        var snapshot = new ReconciliationSnapshot
        {
            PivotTotals = new Dictionary<string, decimal> { ["amount"] = 40m },
            OutputTotals = new Dictionary<string, decimal> { ["amount"] = 39m }
        };
        var plan = new[]
        {
            new BuildCheckPlan
            {
                Id = "rendered",
                Kind = ReportCheckKind.TotalPreservation,
                EvaluationScope = CheckEvaluationScope.RenderedOutput,
                Tolerance = 0.01m
            }
        };

        CheckResult result = Assert.Single(new ReportReconciler().Reconcile(snapshot, plan));

        Assert.Equal(CheckOutcome.Failed, result.Outcome);
    }
}
