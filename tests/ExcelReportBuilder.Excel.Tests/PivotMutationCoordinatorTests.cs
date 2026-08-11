using ExcelReportBuilder.Excel.PivotPlus;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotMutationCoordinatorTests
{
    [Fact]
    public void BatchesStepsRefreshesOnceAndVerifies()
    {
        var pivot = new FakePivot { ManualUpdate = false };
        var calls = new List<string>();
        var coordinator = new PivotMutationCoordinator();

        coordinator.Execute(
            pivot,
            new[]
            {
                Step("rows", calls),
                Step("columns", calls)
            },
            () => calls.Add("refresh:" + pivot.ManualUpdate),
            () => calls.Add("verify"));

        Assert.Equal(
            new[] { "apply:rows", "apply:columns", "refresh:False", "verify" },
            calls);
        Assert.False(pivot.ManualUpdate);
    }

    [Fact]
    public void RollsBackAppliedStepsInReverseOrder()
    {
        var pivot = new FakePivot { ManualUpdate = false };
        var calls = new List<string>();
        var coordinator = new PivotMutationCoordinator();

        PivotMutationException exception = Assert.Throws<PivotMutationException>(() =>
            coordinator.Execute(
                pivot,
                new[]
                {
                    Step("one", calls),
                    Step("two", calls),
                    new PivotMutationStep(
                        "three",
                        () => throw new InvalidOperationException("failed"),
                        () => calls.Add("rollback:three"))
                },
                () => calls.Add("refresh:" + pivot.ManualUpdate),
                () => calls.Add("verify")));

        Assert.Equal("three", exception.FailedStep);
        Assert.True(exception.RollbackCompleted);
        Assert.Equal(
            new[]
            {
                "apply:one",
                "apply:two",
                "rollback:three",
                "rollback:two",
                "rollback:one",
                "refresh:False"
            },
            calls);
        Assert.False(pivot.ManualUpdate);
    }

    [Fact]
    public void VerificationFailureAlsoRollsBack()
    {
        var pivot = new FakePivot { ManualUpdate = true };
        var calls = new List<string>();
        var coordinator = new PivotMutationCoordinator();

        PivotMutationException exception = Assert.Throws<PivotMutationException>(() =>
            coordinator.Execute(
                pivot,
                new[] { Step("layout", calls) },
                () => calls.Add("refresh:" + pivot.ManualUpdate),
                () => throw new InvalidOperationException("wrong layout")));

        Assert.Equal("verify", exception.FailedStep);
        Assert.True(exception.RollbackCompleted);
        Assert.Equal(
            new[] { "apply:layout", "refresh:True", "rollback:layout", "refresh:True" },
            calls);
        Assert.True(pivot.ManualUpdate);
    }

    [Fact]
    public void RejectsReentrantMutation()
    {
        var pivot = new FakePivot();
        var coordinator = new PivotMutationCoordinator();

        PivotMutationException outer = Assert.Throws<PivotMutationException>(() =>
            coordinator.Execute(
                pivot,
                new[]
                {
                    new PivotMutationStep(
                        "outer",
                        () => coordinator.Execute(
                            pivot,
                            Array.Empty<PivotMutationStep>(),
                            () => { },
                            () => { }),
                        () => { })
                },
                () => { },
                () => { }));

        Assert.IsType<InvalidOperationException>(outer.InnerException);
        Assert.False(pivot.ManualUpdate);
    }

    [Fact]
    public void ReportsIncompleteRollback()
    {
        var pivot = new FakePivot();
        var coordinator = new PivotMutationCoordinator();

        PivotMutationException exception = Assert.Throws<PivotMutationException>(() =>
            coordinator.Execute(
                pivot,
                new[]
                {
                    new PivotMutationStep(
                        "one",
                        () => { },
                        () => throw new InvalidOperationException("rollback failed")),
                    new PivotMutationStep(
                        "two",
                        () => throw new InvalidOperationException("apply failed"),
                        () => { })
                },
                () => { },
                () => { }));

        Assert.False(exception.RollbackCompleted);
        Assert.IsType<AggregateException>(exception.InnerException);
    }

    [Fact]
    public void RollsBackAFailingStepThatMutatedBeforeThrowing()
    {
        var pivot = new FakePivot();
        var value = 0;
        var coordinator = new PivotMutationCoordinator();

        PivotMutationException exception = Assert.Throws<PivotMutationException>(() =>
            coordinator.Execute(
                pivot,
                new[]
                {
                    new PivotMutationStep(
                        "partial",
                        () =>
                        {
                            value = 1;
                            throw new InvalidOperationException("Excel failed after mutation");
                        },
                        () => value = 0)
                },
                () => { },
                () => { }));

        Assert.True(exception.RollbackCompleted);
        Assert.Equal(0, value);
    }

    private static PivotMutationStep Step(string name, ICollection<string> calls)
    {
        return new PivotMutationStep(
            name,
            () => calls.Add("apply:" + name),
            () => calls.Add("rollback:" + name));
    }

    public sealed class FakePivot
    {
        public bool ManualUpdate { get; set; }
    }
}
