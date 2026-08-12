using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Semantics;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotSemanticMutationServiceTests
{
    [Fact]
    public void Combined_transaction_runs_artifacts_layout_deletes_one_refresh_and_verify_in_order()
    {
        var log = new List<string>();
        var pivot = new FakePivot { ManualUpdate = false };
        IReadOnlyList<PivotMutationStep> steps = new[]
        {
            Step("measure-upsert", log),
            Step("set-upsert", log),
            Step("layout", log),
            Step("set-delete", log),
            Step("measure-delete", log)
        };

        PivotSemanticMutationService.ExecutePrepared(
            new PivotMutationCoordinator(),
            pivot,
            steps,
            () => log.Add("refresh"),
            () => log.Add("verify"));

        Assert.Equal(
            new[]
            {
                "measure-upsert",
                "set-upsert",
                "layout",
                "set-delete",
                "measure-delete",
                "refresh",
                "verify"
            },
            log);
        Assert.False(pivot.ManualUpdate);
        Assert.Equal(1, log.Count(item => item == "refresh"));
    }

    [Fact]
    public void Combined_transaction_rolls_back_in_exact_reverse_order_with_one_rollback_refresh()
    {
        var log = new List<string>();
        var pivot = new FakePivot { ManualUpdate = false };
        IReadOnlyList<PivotMutationStep> steps = new[]
        {
            Step("measure-upsert", log),
            Step("set-upsert", log),
            Step("layout", log),
            new(
                "set-delete",
                () =>
                {
                    log.Add("set-delete");
                    throw new InvalidOperationException("boom");
                },
                () => log.Add("undo-set-delete"))
        };

        PivotMutationException failure = Assert.Throws<PivotMutationException>(() =>
            PivotSemanticMutationService.ExecutePrepared(
                new PivotMutationCoordinator(),
                pivot,
                steps,
                () => log.Add("refresh"),
                () => log.Add("verify")));

        Assert.True(failure.RollbackCompleted);
        Assert.Equal(
            new[]
            {
                "measure-upsert",
                "set-upsert",
                "layout",
                "set-delete",
                "undo-set-delete",
                "undo-layout",
                "undo-set-upsert",
                "undo-measure-upsert",
                "refresh"
            },
            log);
        Assert.False(pivot.ManualUpdate);
        Assert.Equal(1, log.Count(item => item == "refresh"));
    }

    [Fact]
    public void Undo_orders_restores_before_layout_and_removals_after_layout()
    {
        var log = new List<string>();
        IReadOnlyList<PivotMutationStep> steps =
            PivotSemanticMutationService.BuildUndoSteps(
                new[] { Step("restore-measure", log) },
                new[] { Step("forward-delete-old-set", log) },
                Step("restore-layout", log),
                new[] { Step("forward-create-new-set", log) },
                new[] { Step("delete-created-measure", log) });

        PivotSemanticMutationService.ExecutePrepared(
            new PivotMutationCoordinator(),
            new FakePivot(),
            steps,
            () => log.Add("refresh"),
            () => log.Add("verify"));

        Assert.Equal(
            new[]
            {
                "restore-measure",
                "undo-forward-delete-old-set",
                "restore-layout",
                "undo-forward-create-new-set",
                "delete-created-measure",
                "refresh",
                "verify"
            },
            log);
    }

    private static PivotMutationStep Step(string name, ICollection<string> log)
    {
        return new PivotMutationStep(
            name,
            () => log.Add(name),
            () => log.Add("undo-" + name));
    }

    public sealed class FakePivot
    {
        public bool ManualUpdate { get; set; }
    }
}
