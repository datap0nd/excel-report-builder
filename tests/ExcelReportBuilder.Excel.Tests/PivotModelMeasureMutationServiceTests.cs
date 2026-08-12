using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Measures;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotModelMeasureMutationServiceTests
{
    [Fact]
    public void Apply_CreatesMeasuresInDependencyOrderAndAppliesExactInterleavedValues()
    {
        var events = new List<string>();
        PivotDaxCompilation compilation = Compile(
            Measure("actual", "Actual", Sum()),
            Measure(
                "variance",
                "Variance",
                new PivotDifferenceExpression(
                    new PivotMeasureReferenceExpression("actual"),
                    Sum())));
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0");
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected(existing)),
            events);
        var ownership = new RecordingOwnershipStore(Metadata(), events);
        var identity = new RecordingIdentityResolver();
        var service = Service(gateway, ownership, identity);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                new PivotMeasureValuePlacement(1, "actual"),
                ExistingPlacement(2, existing),
                new PivotMeasureValuePlacement(3, "variance")
            },
            PivotValuesAxis.Columns,
            valuesPosition: 1);

        PivotModelMeasureApplyResult result = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            placement);

        Assert.Equal(PivotModelMeasureApplyStatus.Applied, result.Status);
        Assert.Equal(2, result.Created);
        Assert.Equal(0, result.Updated);
        Assert.Equal(0, result.Deleted);
        Assert.True(result.UndoAvailable);
        Assert.Equal(1, gateway.RefreshCalls);
        Assert.Equal(1, identity.PersistCalls);
        Assert.True(events.IndexOf("journal") < events.FindIndex(item => item.StartsWith("create:", StringComparison.Ordinal)));
        Assert.Equal(
            compilation.CreationSequence.Select(item => "create:" + item.GeneratedMeasureName),
            events.Where(item => item.StartsWith("create:", StringComparison.Ordinal)));
        Assert.Equal(
            new[]
            {
                compilation.Measures[0].GeneratedMeasureName,
                null,
                compilation.Measures[1].GeneratedMeasureName
            },
            gateway.State.SelectedPivot.DataFields.Select(item => item.ModelMeasureName));
        Assert.Equal(PivotValuesAxis.Columns, gateway.State.SelectedPivot.ValuesAxis);
        Assert.NotNull(ownership.Base.Undo);
        Assert.Null(ownership.Pending);
        Assert.All(ownership.Base.Artifacts, artifact =>
            Assert.StartsWith("measure.host.v1:sha256:", artifact.Fingerprint, StringComparison.Ordinal));
    }

    [Fact]
    public void Apply_UpdatesOwnedMeasureInPlaceAndLeavesOtherPivotUntouched()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(
            Measure(
                "metric",
                "Metric",
                new PivotAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Average)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        ModelPivotUsageSnapshot selected = Selected(GeneratedField(prior, 1));
        ModelPivotUsageSnapshot other = OtherPivot(ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0"));
        var gateway = new RecordingGateway(Snapshot(new[] { prior }, selected, other));
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)));
        var service = Service(gateway, ownership);
        var placement = GeneratedPlacement("metric");

        PivotModelMeasureApplyResult result = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            placement);

        Assert.Equal(1, result.Updated);
        Assert.Equal(0, result.Created);
        Assert.Equal(0, result.Deleted);
        Assert.Contains("update:Metric", gateway.Events);
        Assert.DoesNotContain("create:Metric", gateway.Events);
        Assert.DoesNotContain("delete:Metric", gateway.Events);
        Assert.Equal(
            PivotModelMeasureCanonical.CreatePivotFingerprint(other),
            PivotModelMeasureCanonical.CreatePivotFingerprint(
                gateway.State.PivotUsages.Single(item => !item.IsSelectedTarget)));
    }

    [Fact]
    public void Apply_RejectsUnownedNameCollisionBeforeIdentityOrJournal()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        LiveModelMeasureSnapshot collision = Live(
            compilation.Measures.Single(),
            descriptionOverride: "user authored");
        var gateway = new RecordingGateway(
            Snapshot(new[] { collision }, Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var identity = new RecordingIdentityResolver();
        var service = Service(gateway, ownership, identity);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        Assert.Contains("unowned", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, identity.PersistCalls);
        Assert.Equal(0, ownership.JournalCalls);
        Assert.DoesNotContain(gateway.Events, item => item.StartsWith("create:", StringComparison.Ordinal));
    }

    [Fact]
    public void Apply_RequiresEveryCurrentUnownedValueExactlyOnce()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0");
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected(existing)));
        var ownership = new RecordingOwnershipStore(Metadata());
        var identity = new RecordingIdentityResolver();
        var service = Service(gateway, ownership, identity);

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        Assert.Contains("unowned", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, identity.PersistCalls);
        Assert.Equal(0, ownership.JournalCalls);
    }

    [Fact]
    public void Apply_BlocksChangedMeasureUsedByAnotherPivotBeforeMutation()
    {
        PivotDaxCompilation priorCompilation = Compile(Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(Measure(
            "metric",
            "Metric",
            new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Maximum)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        var gateway = new RecordingGateway(Snapshot(
            new[] { prior },
            Selected(GeneratedField(prior, 1)),
            OtherPivot(GeneratedField(prior, 1))));
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)));
        var identity = new RecordingIdentityResolver();
        var service = Service(gateway, ownership, identity);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                desiredCompilation,
                GeneratedPlacement("metric")));

        Assert.Contains("another PivotTable", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, identity.PersistCalls);
        Assert.Equal(0, ownership.JournalCalls);
    }

    [Fact]
    public void Apply_DeletesDependentsBeforeDependencies()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("base", "Base", Sum()),
            Measure(
                "dependent",
                "Dependent",
                new PivotMeasureReferenceExpression("base")),
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        PivotDaxCompilation desiredCompilation = Compile(
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        LiveModelMeasureSnapshot[] prior = priorCompilation.Measures
            .Select(item => Live(item))
            .ToArray();
        var gateway = new RecordingGateway(Snapshot(
            prior,
            Selected(prior.Select((item, index) => GeneratedField(item, index + 1)).ToArray())));
        var ownership = new RecordingOwnershipStore(Metadata(prior.Select(Artifact).ToArray()));
        var service = Service(gateway, ownership);

        PivotModelMeasureApplyResult result = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            GeneratedPlacement("keep"));

        Assert.Equal(2, result.Deleted);
        Assert.Equal(
            new[] { "delete:Dependent", "delete:Base" },
            gateway.Events.Where(item => item.StartsWith("delete:", StringComparison.Ordinal)));
    }

    [Fact]
    public void Apply_BlocksDeleteWhenUserMeasureDependsOnOwnedMeasure()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("remove", "Remove", Sum()),
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        PivotDaxCompilation desiredCompilation = Compile(
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        LiveModelMeasureSnapshot[] owned = priorCompilation.Measures
            .Select(item => Live(item))
            .ToArray();
        LiveModelMeasureSnapshot user = UserMeasure("User Metric", "=[Remove] + 1");
        var gateway = new RecordingGateway(Snapshot(
            owned.Concat(new[] { user }),
            Selected(owned.Select((item, index) => GeneratedField(item, index + 1)).ToArray())));
        var ownership = new RecordingOwnershipStore(Metadata(owned.Select(Artifact).ToArray()));
        var service = Service(gateway, ownership);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                desiredCompilation,
                GeneratedPlacement("keep")));

        Assert.Contains("depend", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, ownership.JournalCalls);
    }

    [Fact]
    public void Apply_UpdatesOwnedDependentToRemoveReferenceBeforeDeletingDependency()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("base", "Base", Sum()),
            Measure("dependent", "Dependent", new PivotMeasureReferenceExpression("base")),
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        PivotDaxCompilation desiredCompilation = Compile(
            Measure("dependent", "Dependent", Sum()),
            Measure("keep", "Keep", new PivotAggregateExpression(
                "units",
                PivotCalculationAggregateFunction.Sum)));
        LiveModelMeasureSnapshot[] prior = priorCompilation.Measures
            .Select(item => Live(item))
            .ToArray();
        var gateway = new RecordingGateway(Snapshot(
            prior,
            Selected(prior.Select((item, index) => GeneratedField(item, index + 1)).ToArray())));
        var ownership = new RecordingOwnershipStore(Metadata(prior.Select(Artifact).ToArray()));
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, "dependent"),
                new PivotMeasureValuePlacement(2, "keep")
            },
            PivotValuesAxis.Columns,
            1);

        PivotModelMeasureApplyResult result = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            placement);

        Assert.True(result.Updated >= 1);
        Assert.Equal(1, result.Deleted);
        string[] orderedEvents = gateway.Events.ToArray();
        Assert.True(
            Array.IndexOf(orderedEvents, "update:Dependent") <
            Array.IndexOf(orderedEvents, "delete:Base"));
    }

    [Fact]
    public void Apply_BlocksFreshNameThatMayChangeExistingUserMeasureBinding()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        LiveModelMeasureSnapshot user = UserMeasure("User Metric", "=[Metric] + 1");
        var gateway = new RecordingGateway(Snapshot(new[] { user }, Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        Assert.Contains("depend", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, ownership.JournalCalls);
    }

    [Fact]
    public void Apply_RefreshFailureRollsBackExcelAndClearsFreshJournal()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0");
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(existing));
        var gateway = new RecordingGateway(initial)
        {
            ThrowOnRefreshCall = 1
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, existing),
                new PivotMeasureValuePlacement(2, "metric")
            },
            PivotValuesAxis.Columns,
            1);

        PivotModelMeasureMutationException exception = Assert.Throws<PivotModelMeasureMutationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                placement));

        Assert.True(exception.RollbackCompleted);
        Assert.False(exception.RecoveryRequired);
        Assert.Equal(2, gateway.RefreshCalls);
        Assert.Empty(gateway.State.Measures);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Null(ownership.Pending);
        Assert.Equal(1, ownership.RestoreCalls);
    }

    [Fact]
    public void Apply_CommitFailureRetainsPendingAndIdenticalRetryConvergesWithoutDuplicateCreate()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var events = new List<string>();
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()),
            events);
        var ownership = new RecordingOwnershipStore(Metadata(), events)
        {
            ThrowOnCommit = true
        };
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException exception = Assert.Throws<PivotModelMeasureMutationException>(() =>
            service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        Assert.True(exception.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Single(gateway.State.Measures);
        Assert.Single(gateway.Events, item => item == "create:Metric");

        ownership.ThrowOnCommit = false;
        PivotModelMeasureApplyResult recovered = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));

        Assert.Equal(PivotModelMeasureApplyStatus.Applied, recovered.Status);
        Assert.Single(gateway.Events, item => item == "create:Metric");
        Assert.Contains("update:Metric", gateway.Events);
        Assert.Null(ownership.Pending);
        Assert.Single(ownership.Base.Artifacts);
        Assert.True(recovered.UndoAvailable);

        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);
        Assert.Empty(gateway.State.Measures);
        Assert.Empty(ownership.Base.Artifacts);
    }

    [Fact]
    public void Apply_CommitFailureRetryAcceptsExactAlreadyFinalRepeatedValuesAndPreservesOriginalUndo()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        ModelDataFieldSnapshot first = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            position: 1);
        ModelDataFieldSnapshot second = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            position: 2);
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(first, second));
        var gateway = new RecordingGateway(initial);
        var ownership = new RecordingOwnershipStore(Metadata())
        {
            ThrowOnCommit = true
        };
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, second),
                new PivotMeasureValuePlacement(2, "metric"),
                ExistingPlacement(3, first)
            },
            PivotValuesAxis.Columns,
            1);

        Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            placement));
        Assert.Equal(1, gateway.Events.Count(item => item == "placement"));

        ownership.ThrowOnCommit = false;
        PivotModelMeasureApplyResult recovered = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            placement);

        Assert.True(recovered.UndoAvailable);
        Assert.Equal(1, gateway.Events.Count(item => item == "placement"));
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void Apply_PartialPlacementRollbackRetryRepairsFromSessionPreviewAndPreservesUndo()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        ModelDataFieldSnapshot first = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            position: 1);
        ModelDataFieldSnapshot second = ExistingField(
            "[Measures].[Sum of Cost]",
            "Sum of Cost",
            "$#,##0",
            position: 2);
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(first, second));
        var gateway = new RecordingGateway(initial)
        {
            ThrowAfterPlacementCall = 1,
            ThrowAfterRestorePlacementCall = 1
        };
        gateway.RestorePlacementMutation = (host, call) =>
        {
            if (call != 1) return;
            ModelPivotUsageSnapshot selected = host.State.SelectedPivot;
            host.ReplaceSelectedForTest(new ModelPivotUsageSnapshot(
                selected.WorksheetName,
                selected.PivotTableName,
                isSelectedTarget: true,
                selected.DataFields,
                PivotValuesAxis.Rows,
                valuesPosition: 2));
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, second),
                new PivotMeasureValuePlacement(2, "metric"),
                ExistingPlacement(3, first)
            },
            PivotValuesAxis.Columns,
            1);

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                placement));
        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.NotEqual(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);

        gateway.ThrowAfterPlacementCall = 0;
        gateway.ThrowAfterRestorePlacementCall = 0;
        PivotModelMeasureApplyResult recovered = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            placement);

        Assert.True(recovered.UndoAvailable);
        Assert.Contains("restore-placement", gateway.Events);
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void Apply_PostRefreshMeasureDriftIsNotBlessedAsOwned()
    {
        PivotDaxCompilation priorCompilation = Compile(Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(Measure(
            "metric",
            "Metric",
            new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Average)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        var gateway = new RecordingGateway(Snapshot(
            new[] { prior },
            Selected(GeneratedField(prior, 1))));
        gateway.RefreshMutation = (host, call) =>
        {
            if (call != 1) return;
            LiveModelMeasureSnapshot updated = host.State.Measures.Single();
            host.ReplaceMeasure(WithFormula(updated, updated.Formula + " + 0"));
        };
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)));
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException exception =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                desiredCompilation,
                GeneratedPlacement("metric")));

        Assert.True(exception.RollbackCompleted);
        Assert.False(exception.RecoveryRequired);
        Assert.Equal(prior.LiveFingerprint, gateway.State.Measures.Single().LiveFingerprint);
        Assert.Null(ownership.Pending);
        Assert.Equal(prior.LiveFingerprint, ownership.Base.Artifacts.Single().Fingerprint);
    }

    [Fact]
    public void Apply_PostRefreshGeneratedValueDisplayDriftRollsBackExactLayout()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        LiveModelMeasureSnapshot metric = Live(compilation.Measures.Single());
        ModelDataFieldSnapshot generated = GeneratedField(metric, 1);
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            position: 2);
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            new[] { metric },
            Selected(generated, existing));
        var gateway = new RecordingGateway(initial);
        gateway.RefreshMutation = (host, call) =>
        {
            if (call != 1) return;
            ModelPivotUsageSnapshot selected = host.State.SelectedPivot;
            ModelDataFieldSnapshot changed = WithDisplay(
                selected.DataFields.Single(field => field.IsModelMeasure),
                "Changed caption",
                "0.0000");
            host.ReplaceSelectedForTest(new ModelPivotUsageSnapshot(
                selected.WorksheetName,
                selected.PivotTableName,
                isSelectedTarget: true,
                selected.DataFields.Select(field => field.IsModelMeasure ? changed : field),
                selected.ValuesAxis,
                selected.ValuesPosition));
        };
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(metric)));
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, existing),
                new PivotMeasureValuePlacement(2, "metric")
            },
            PivotValuesAxis.Columns,
            1);

        PivotModelMeasureMutationException exception =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                placement));

        Assert.True(exception.RollbackCompleted);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Null(ownership.Pending);
    }

    [Fact]
    public void Apply_FinalRetryDoesNotBlessGeneratedValueDisplayDriftFromOriginalPreview()
    {
        PivotDaxCompilation priorCompilation = Compile(Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(Measure(
            "metric",
            "Metric",
            new PivotAggregateExpression(
                "amount",
                PivotCalculationAggregateFunction.Average)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        var gateway = new RecordingGateway(Snapshot(
            new[] { prior },
            Selected(GeneratedField(prior, 1))));
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)))
        {
            ThrowOnCommit = true
        };
        var service = Service(gateway, ownership);
        Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            GeneratedPlacement("metric")));
        ModelPivotUsageSnapshot selected = gateway.State.SelectedPivot;
        gateway.ReplaceSelectedForTest(new ModelPivotUsageSnapshot(
            selected.WorksheetName,
            selected.PivotTableName,
            isSelectedTarget: true,
            selected.DataFields.Select(field => WithDisplay(
                field,
                "Changed caption",
                "0.0000")),
            selected.ValuesAxis,
            selected.ValuesPosition));
        ownership.ThrowOnCommit = false;

        PivotModelMeasureMutationException retry =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                desiredCompilation,
                GeneratedPlacement("metric")));

        Assert.True(retry.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Equal(prior.LiveFingerprint, ownership.Base.Artifacts.Single().Fingerprint);
    }

    [Fact]
    public void Apply_RestartedPendingUpdateRefusesToGuessPriorGeneratedValueDisplay()
    {
        PivotDaxCompilation priorCompilation = Compile(Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(Measure(
            "metric",
            "Metric",
            new PivotAggregateExpression(
                "amount",
                PivotCalculationAggregateFunction.Average)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        var gateway = new RecordingGateway(Snapshot(
            new[] { prior },
            Selected(GeneratedField(prior, 1))));
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)))
        {
            ThrowOnCommit = true
        };
        var firstSession = Service(gateway, ownership);
        Assert.Throws<PivotModelMeasureMutationException>(() => firstSession.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            GeneratedPlacement("metric")));
        Assert.NotNull(ownership.Pending);
        ownership.ThrowOnCommit = false;
        var restarted = Service(gateway, ownership);

        InvalidOperationException retry = Assert.Throws<InvalidOperationException>(() =>
            restarted.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                desiredCompilation,
                GeneratedPlacement("metric")));

        Assert.Contains("session preview", retry.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(ownership.Pending);
        Assert.Equal(prior.LiveFingerprint, ownership.Base.Artifacts.Single().Fingerprint);
    }

    [Fact]
    public void Apply_CreateCommitThenThrowRetainsRecoveryOwnershipUntilExactRetry()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()))
        {
            ThrowAfterCreateName = compilation.Measures.Single().GeneratedMeasureName
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        Assert.False(failure.RollbackCompleted);
        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Single(gateway.State.Measures);

        gateway.ThrowAfterCreateName = null;
        PivotModelMeasureApplyResult recovered = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));

        Assert.Equal(PivotModelMeasureApplyStatus.Applied, recovered.Status);
        Assert.Single(gateway.State.Measures);
        Assert.Null(ownership.Pending);
        Assert.True(recovered.UndoAvailable);
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void Apply_RefreshSideEffectThenThrowDoesNotClearRecoveryJournal()
    {
        LiveModelMeasureSnapshot user = UserMeasure("User Metric", "=1");
        var gateway = new RecordingGateway(Snapshot(new[] { user }, Selected()))
        {
            ThrowOnRefreshCall = 1
        };
        gateway.RefreshMutation = (host, call) =>
        {
            if (call == 1)
            {
                host.ReplaceMeasure(WithFormula(user, "=2"));
            }
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                Compile(Measure("metric", "Metric", Sum())),
                GeneratedPlacement("metric")));

        Assert.False(failure.RollbackCompleted);
        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Equal("=2", gateway.State.Measures.Single().Formula);
        Assert.Equal(0, ownership.RestoreCalls);
    }

    [Fact]
    public void Apply_PostRefreshDependentUserMeasureCannotBeBlessedOrCleared()
    {
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        gateway.RefreshMutation = (host, call) =>
        {
            if (call == 1)
            {
                host.AddMeasure(UserMeasure("Dependent User Measure", "=[Metric]"));
            }
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                Compile(Measure("metric", "Metric", Sum())),
                GeneratedPlacement("metric")));

        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Contains(gateway.State.Measures, measure =>
            string.Equals(measure.Name, "Dependent User Measure", StringComparison.Ordinal));
        Assert.Empty(ownership.Base.Artifacts);
    }

    [Fact]
    public void Apply_PostRefreshSiblingPivotUsingChangedMeasureCannotBeBlessed()
    {
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        gateway.RefreshMutation = (host, call) =>
        {
            if (call != 1) return;
            LiveModelMeasureSnapshot metric = host.State.Measures.Single();
            host.AddOtherPivot(OtherPivot(GeneratedField(metric, 1)));
        };
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                Compile(Measure("metric", "Metric", Sum())),
                GeneratedPlacement("metric")));

        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Equal(2, gateway.State.PivotUsages.Count);
        Assert.Empty(ownership.Base.Artifacts);
    }

    [Fact]
    public void Apply_FailedPendingRetryKeepsJournalEvenWhenRetryStartRollbackCompletes()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata())
        {
            ThrowOnCommit = true
        };
        var service = Service(gateway, ownership);
        Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric")));
        Assert.NotNull(ownership.Pending);

        ownership.ThrowOnCommit = false;
        gateway.ThrowOnRefreshCall = gateway.RefreshCalls + 1;
        PivotModelMeasureMutationException retryFailure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                GeneratedPlacement("metric")));

        // The pending retry starts with the already-created measure. Its
        // logical-create rollback removes that measure, so the coordinator's
        // callback success is not an exact retry-start rollback. The service
        // correctly refuses to claim it as restored.
        Assert.False(retryFailure.RollbackCompleted);
        Assert.True(retryFailure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Empty(gateway.State.Measures);
        Assert.Equal(0, ownership.RestoreCalls);

        gateway.ThrowOnRefreshCall = 0;
        PivotModelMeasureApplyResult recovered = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));
        Assert.Equal(PivotModelMeasureApplyStatus.Applied, recovered.Status);
        Assert.Null(ownership.Pending);
    }

    [Fact]
    public void Apply_CommitAfterSaveAmbiguityPromotesOriginalUndoOnNoChangeRetry()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata())
        {
            CommitThenThrow = true
        };
        var service = Service(gateway, ownership);

        Assert.Throws<PivotModelMeasureMutationException>(() => service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric")));
        Assert.Null(ownership.Pending);
        Assert.Single(ownership.Base.Artifacts);

        ownership.CommitThenThrow = false;
        PivotModelMeasureApplyResult retry = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));

        Assert.Equal(PivotModelMeasureApplyStatus.NoChange, retry.Status);
        Assert.True(retry.UndoAvailable);
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void Undo_RestoresExactSessionSnapshotAndIsUnavailableAfterRestart()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0");
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(existing));
        var gateway = new RecordingGateway(initial);
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                new PivotMeasureValuePlacement(1, "metric"),
                ExistingPlacement(2, existing)
            },
            PivotValuesAxis.Columns,
            1);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            placement);

        service.Undo(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId);

        Assert.Empty(gateway.State.Measures);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Empty(ownership.Base.Artifacts);
        Assert.Null(ownership.Base.Undo);
        Assert.Throws<PivotModelMeasureUndoUnavailableException>(() =>
            service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId));

        var restarted = Service(gateway, ownership);
        Assert.Throws<PivotModelMeasureUndoUnavailableException>(() =>
            restarted.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId));
    }

    [Fact]
    public void Undo_RestoresUpdatedMeasureFromSessionOnlySnapshot()
    {
        PivotDaxCompilation priorCompilation = Compile(Measure("metric", "Metric", Sum()));
        PivotDaxCompilation desiredCompilation = Compile(Measure(
            "metric",
            "Metric",
            new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Minimum)));
        LiveModelMeasureSnapshot prior = Live(priorCompilation.Measures.Single());
        var gateway = new RecordingGateway(Snapshot(
            new[] { prior },
            Selected(GeneratedField(prior, 1))));
        var ownership = new RecordingOwnershipStore(Metadata(Artifact(prior)));
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            GeneratedPlacement("metric"));

        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        LiveModelMeasureSnapshot restored = Assert.Single(gateway.State.Measures);
        Assert.Equal(prior.LiveFingerprint, restored.LiveFingerprint);
        Assert.Equal(prior.Formula, restored.Formula);
        Assert.Equal(prior.Description, restored.Description);
        Assert.Equal(prior.LiveFingerprint, Assert.Single(ownership.Base.Artifacts).Fingerprint);
    }

    [Fact]
    public void Undo_CommitFailureRetryFinalizesWithoutRepeatingWorkbookMutation()
    {
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            GeneratedPlacement("metric"));
        ownership.ThrowOnCommit = true;

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Undo(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId));

        Assert.True(failure.RecoveryRequired);
        Assert.Empty(gateway.State.Measures);
        Assert.NotNull(ownership.Pending);
        int refreshes = gateway.RefreshCalls;
        int mutations = gateway.Events.Count(item =>
            item.StartsWith("delete:", StringComparison.Ordinal) ||
            item.StartsWith("restore:", StringComparison.Ordinal) ||
            item == "placement" || item == "restore-placement");

        ownership.ThrowOnCommit = false;
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        Assert.Equal(refreshes, gateway.RefreshCalls);
        Assert.Equal(
            mutations,
            gateway.Events.Count(item =>
                item.StartsWith("delete:", StringComparison.Ordinal) ||
                item.StartsWith("restore:", StringComparison.Ordinal) ||
                item == "placement" || item == "restore-placement"));
        Assert.Null(ownership.Pending);
        Assert.Empty(ownership.Base.Artifacts);
    }

    [Fact]
    public void Undo_IncompleteRollbackRetryRepairsIntermediateWorkbookAndFinalizes()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        string measureName = compilation.Measures.Single().GeneratedMeasureName;
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));

        gateway.ThrowOnRefreshCall = gateway.RefreshCalls + 1;
        gateway.ThrowBeforeRestoreName = measureName;
        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Undo(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId));

        Assert.False(failure.RollbackCompleted);
        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Empty(gateway.State.Measures);

        gateway.ThrowOnRefreshCall = 0;
        gateway.ThrowBeforeRestoreName = null;
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        Assert.Empty(gateway.State.Measures);
        Assert.Empty(gateway.State.SelectedPivot.DataFields);
        Assert.Empty(ownership.Base.Artifacts);
        Assert.Null(ownership.Pending);
        Assert.Throws<PivotModelMeasureUndoUnavailableException>(() => service.Undo(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId));
    }

    [Fact]
    public void Undo_PostRefreshSiblingPivotUsingChangedMeasureRetainsRecoveryJournal()
    {
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            GeneratedPlacement("metric"));
        LiveModelMeasureSnapshot metric = gateway.State.Measures.Single();
        gateway.RefreshMutation = (host, call) =>
        {
            if (call == 2)
            {
                host.AddOtherPivot(OtherPivot(GeneratedField(metric, 1)));
            }
        };

        PivotModelMeasureMutationException failure =
            Assert.Throws<PivotModelMeasureMutationException>(() => service.Undo(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId));

        Assert.True(failure.RecoveryRequired);
        Assert.NotNull(ownership.Pending);
        Assert.Equal(2, gateway.State.PivotUsages.Count);
        Assert.Single(ownership.Base.Artifacts);
    }

    [Fact]
    public void Undo_CommitAfterSaveAmbiguityIsRecognizedAsAlreadyFinalized()
    {
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            GeneratedPlacement("metric"));
        ownership.CommitThenThrow = true;

        Assert.Throws<PivotModelMeasureMutationException>(() => service.Undo(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId));
        Assert.Null(ownership.Pending);
        int refreshes = gateway.RefreshCalls;
        int commits = ownership.CommitCalls;

        ownership.CommitThenThrow = false;
        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        Assert.Equal(refreshes, gateway.RefreshCalls);
        Assert.Equal(commits, ownership.CommitCalls);
        Assert.Empty(ownership.Base.Artifacts);
    }

    [Fact]
    public void Undo_DeletesCreatedDependentsBeforeTheirDependencies()
    {
        PivotDaxCompilation compilation = Compile(
            Measure("base", "Base", Sum()),
            Measure("dependent", "Dependent", new PivotMeasureReferenceExpression("base")));
        var events = new List<string>();
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()),
            events);
        var service = Service(gateway, new RecordingOwnershipStore(Metadata(), events));
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            new PivotMeasurePlacementPlan(
                new[]
                {
                    new PivotMeasureValuePlacement(1, "base"),
                    new PivotMeasureValuePlacement(2, "dependent")
                },
                PivotValuesAxis.Columns,
                1));
        events.Clear();

        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        Assert.Equal(
            new[] { "delete:Dependent", "delete:Base" },
            events.Where(item => item.StartsWith("delete:", StringComparison.Ordinal)));
    }

    [Fact]
    public void Undo_RecreatesDeletedDependenciesBeforeTheirDependents()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("base", "Base", Sum()),
            Measure("dependent", "Dependent", new PivotMeasureReferenceExpression("base")),
            Measure(
                "keep",
                "Keep",
                new PivotAggregateExpression(
                    "units",
                    PivotCalculationAggregateFunction.Sum)));
        LiveModelMeasureSnapshot[] prior = priorCompilation.Measures
            .Select(item => Live(item))
            .ToArray();
        var events = new List<string>();
        var gateway = new RecordingGateway(Snapshot(
            prior,
            Selected(prior.Select((item, index) => GeneratedField(item, index + 1)).ToArray())),
            events);
        var ownership = new RecordingOwnershipStore(
            Metadata(prior.Select(Artifact).ToArray()),
            events);
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure(
                "keep",
                "Keep",
                new PivotAggregateExpression(
                    "units",
                    PivotCalculationAggregateFunction.Sum))),
            GeneratedPlacement("keep"));
        events.Clear();

        service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId);

        Assert.Equal(
            new[] { "restore:Base", "restore:Dependent" },
            events.Where(item => item.StartsWith("restore:", StringComparison.Ordinal)).Take(2));
    }

    [Fact]
    public void Undo_BlocksNewUserDependencyThroughUnchangedOwnedDependent()
    {
        PivotDaxCompilation priorCompilation = Compile(
            Measure("changed", "Changed", Sum()),
            Measure("dependent", "Dependent", new PivotMeasureReferenceExpression("changed")));
        PivotDaxCompilation desiredCompilation = Compile(
            Measure(
                "changed",
                "Changed",
                new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Average)),
            Measure("dependent", "Dependent", new PivotMeasureReferenceExpression("changed")));
        LiveModelMeasureSnapshot[] prior = priorCompilation.Measures
            .Select(item => Live(item))
            .ToArray();
        var gateway = new RecordingGateway(Snapshot(
            prior,
            Selected(prior.Select((item, index) => GeneratedField(item, index + 1)).ToArray())));
        var ownership = new RecordingOwnershipStore(Metadata(prior.Select(Artifact).ToArray()));
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desiredCompilation,
            new PivotMeasurePlacementPlan(
                new[]
                {
                    new PivotMeasureValuePlacement(1, "changed"),
                    new PivotMeasureValuePlacement(2, "dependent")
                },
                PivotValuesAxis.Columns,
                1));
        gateway.AddMeasure(UserMeasure("User", "=[Dependent] + 1"));
        int journalCalls = ownership.JournalCalls;

        PivotModelMeasureUndoUnavailableException exception =
            Assert.Throws<PivotModelMeasureUndoUnavailableException>(() => service.Undo(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId));

        Assert.Contains("depend", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(journalCalls, ownership.JournalCalls);
    }

    [Fact]
    public void Undo_BlocksWhenAnotherPivotStartsUsingCreatedMeasure()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));
        LiveModelMeasureSnapshot created = Assert.Single(gateway.State.Measures);
        gateway.AddOtherPivot(OtherPivot(GeneratedField(created, 1)));
        int journalCallsBeforeUndo = ownership.JournalCalls;

        PivotModelMeasureUndoUnavailableException exception =
            Assert.Throws<PivotModelMeasureUndoUnavailableException>(() =>
                service.Undo(gateway.Workbook, gateway.Pivot, Context(), SetupId));

        Assert.Contains("PivotTable", exception.Message, StringComparison.Ordinal);
        Assert.Equal(journalCallsBeforeUndo, ownership.JournalCalls);
        Assert.Single(gateway.State.Measures);
    }

    [Fact]
    public void PrepareParticipant_DefersRefreshJournalAndCommitForCombinedSemanticCoordinator()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var ownership = new RecordingOwnershipStore(Metadata());
        var identity = new RecordingIdentityResolver();
        var service = Service(gateway, ownership, identity);

        PivotModelMeasurePreparedMutation prepared = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"),
            ownership.Base,
            existingPending: null);

        Assert.Equal(0, gateway.RefreshCalls);
        Assert.Equal(0, ownership.JournalCalls);
        Assert.Equal(0, identity.PersistCalls);
        Assert.Empty(gateway.State.Measures);
        Assert.NotEmpty(prepared.Steps);

        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "measure and named-set plan");
        prepared.PrimeUndoContribution("apply_combined", combinedPlan);

        ModelMeasureWorkbookSnapshot? verified = null;
        new PivotMutationCoordinator().Execute(
            prepared.Target.PivotTable,
            prepared.Steps,
            prepared.Refresh,
            () => verified = prepared.Verify());

        Assert.NotNull(verified);
        Assert.Equal(1, gateway.RefreshCalls);
        PivotPlusOwnedArtifact artifact = Assert.Single(prepared.BuildArtifacts(verified!));
        Assert.StartsWith("measure.host.v1:sha256:", artifact.Fingerprint, StringComparison.Ordinal);
        Assert.Equal(0, ownership.JournalCalls);

        PivotModelMeasureUndoContribution? contribution = prepared.BuildUndoContribution(
            "apply_combined",
            combinedPlan,
            verified!);
        Assert.NotNull(contribution);
        Assert.Equal(combinedPlan, contribution!.PlanFingerprint);
        Assert.Equal(prepared.Before.SelectedPivotFingerprint, contribution.Before.SelectedPivotFingerprint);
        Assert.Equal(verified!.SelectedPivotFingerprint, contribution.After.SelectedPivotFingerprint);
        Assert.Single(contribution.AfterOwnedArtifacts);
        Assert.Single(prepared.UpsertSteps);
        Assert.Single(prepared.PlacementSteps);
        Assert.Empty(prepared.DeleteSteps);
    }

    [Fact]
    public void PreparedParticipants_ApplyAndUndoThroughOneRefreshPerCombinedBoundary()
    {
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected());
        var gateway = new RecordingGateway(initial);
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        PivotModelMeasurePreparedMutation apply = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            GeneratedPlacement("metric"),
            ownership.Base,
            existingPending: null);
        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "prepared apply and undo");
        apply.PrimeUndoContribution("apply_combined_undo", combinedPlan);
        ModelMeasureWorkbookSnapshot? after = null;

        new PivotMutationCoordinator().Execute(
            apply.Target.PivotTable,
            apply.Steps,
            apply.Refresh,
            () => after = apply.Verify());
        PivotModelMeasureUndoContribution? contribution = apply.BuildUndoContribution(
            "apply_combined_undo",
            combinedPlan,
            after!);
        Assert.NotNull(contribution);

        PivotModelMeasurePreparedUndo undo = service.PrepareUndoParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            contribution!);
        new PivotMutationCoordinator().Execute(
            undo.Target.PivotTable,
            undo.Steps,
            undo.Refresh,
            undo.Verify);

        Assert.Equal(2, gateway.RefreshCalls);
        Assert.Empty(gateway.State.Measures);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
    }

    [Fact]
    public void PreparedParticipant_CombinedFailureCanProveExactRollback()
    {
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected());
        var gateway = new RecordingGateway(initial);
        var service = Service(gateway, new RecordingOwnershipStore(Metadata()));
        PivotModelMeasurePreparedMutation prepared = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            GeneratedPlacement("metric"),
            Metadata(),
            existingPending: null);
        prepared.PrimeUndoContribution(
            "apply_combined_failure",
            PivotPlusFingerprint.Create("semantic.combined.v1", "failure"));
        var combinedSteps = prepared.Steps.Concat(new[]
        {
            new PivotMutationStep(
                "synthetic named-set failure",
                () => throw new InvalidOperationException("synthetic failure"),
                () => { })
        }).ToList();

        PivotMutationException failure = Assert.Throws<PivotMutationException>(() =>
            new PivotMutationCoordinator().Execute(
                prepared.Target.PivotTable,
                combinedSteps,
                prepared.Refresh,
                () => prepared.Verify()));

        Assert.True(failure.RollbackCompleted);
        prepared.VerifyRollback();
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void PrepareParticipant_RejectsBaseMetadataFromAnotherSetupOrTarget()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        var service = Service(gateway, new RecordingOwnershipStore(Metadata()));
        PivotPlusWorkbookMetadata wrongSetup = Metadata();
        wrongSetup.SetupId = "setup_other";
        PivotPlusWorkbookMetadata wrongTarget = Metadata();
        wrongTarget.TargetWorksheetName = "OtherSheet";

        Assert.Throws<InvalidOperationException>(() => service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"),
            wrongSetup,
            existingPending: null));
        Assert.Throws<InvalidOperationException>(() => service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"),
            wrongTarget,
            existingPending: null));
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void PrepareParticipant_ResumesMeasureSliceOfCombinedPendingJournal()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        PivotPlusWorkbookMetadata metadata = Metadata();
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotModelMeasurePreparedMutation fresh = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"),
            metadata,
            existingPending: null);
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = fresh.Pending.ApplyId,
            PlanFingerprint = PivotPlusFingerprint.Create(
                "semantic.combined.v1",
                "combined plan"),
            BeforePivotFingerprint = fresh.Pending.BeforePivotFingerprint,
            ExpectedPivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.combined-pivot.v1",
                "combined expected pivot"),
            Transitions = fresh.Pending.Transitions
                .Select(CloneTransition)
                .Concat(new[]
                {
                    new PivotPlusSemanticArtifactTransition
                    {
                        Kind = PivotPlusArtifactKind.NamedSet,
                        ArtifactId = "PP_Set_Q1",
                        Operation = PivotPlusSemanticArtifactOperation.Create,
                        PlannedDefinitionFingerprint = PivotPlusFingerprint.Create(
                            "namedset.definition.v1",
                            "Q1")
                    }
                })
                .ToList()
        };

        PivotModelMeasurePreparedMutation resumed = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"),
            metadata,
            combined,
            new PivotModelMeasureParticipantRetryBinding(
                combined.PlanFingerprint,
                fresh.Pending.PlanFingerprint));

        Assert.Equal(combined.ApplyId, resumed.Pending.ApplyId);
        Assert.Single(resumed.Pending.Transitions);
        Assert.All(resumed.Pending.Transitions, transition =>
            Assert.Equal(PivotPlusArtifactKind.Measure, transition.Kind));
        Assert.Single(resumed.UpsertSteps);
    }

    [Fact]
    public void PrepareParticipant_CombinedRetryRequiresExactCoordinatorPlanBinding()
    {
        PivotDaxCompilation compilation = Compile(
            Measure("first", "First", Sum()),
            Measure(
                "second",
                "Second",
                new PivotAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Average)));
        var originalPlacement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, "first"),
                new PivotMeasureValuePlacement(2, "second")
            },
            PivotValuesAxis.Columns,
            1);
        var changedPlacement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, "second"),
                new PivotMeasureValuePlacement(2, "first")
            },
            PivotValuesAxis.Columns,
            1);
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        PivotPlusWorkbookMetadata metadata = Metadata();
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotModelMeasurePreparedMutation fresh = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            originalPlacement,
            metadata,
            existingPending: null);
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = fresh.Pending.ApplyId,
            PlanFingerprint = PivotPlusFingerprint.Create(
                "semantic.combined.v1",
                "original combined placement"),
            BeforePivotFingerprint = fresh.Pending.BeforePivotFingerprint,
            ExpectedPivotFingerprint = fresh.Pending.ExpectedPivotFingerprint,
            Transitions = fresh.Pending.Transitions.Select(CloneTransition).ToList()
        };
        Assert.Throws<InvalidOperationException>(() => service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            changedPlacement,
            metadata,
            combined,
            new PivotModelMeasureParticipantRetryBinding(
                combined.PlanFingerprint,
                fresh.Pending.PlanFingerprint)));
        Assert.Empty(gateway.State.Measures);
    }

    [Fact]
    public void PrepareParticipant_ResumedCreateUpdateDeleteAndValuesKeepsOriginalUndoContribution()
    {
        PivotDaxCompilation priorKeep = Compile(Measure("keep", "Keep", Sum()));
        PivotDaxCompilation priorRemove = Compile(Measure("remove", "Remove", Sum()));
        LiveModelMeasureSnapshot keep = Live(priorKeep.Measures.Single());
        LiveModelMeasureSnapshot remove = Live(priorRemove.Measures.Single());
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            new[] { keep, remove },
            Selected(GeneratedField(keep, 1), GeneratedField(remove, 2)));
        var gateway = new RecordingGateway(initial);
        PivotPlusWorkbookMetadata metadata = Metadata(Artifact(keep), Artifact(remove));
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotDaxCompilation desired = Compile(
            Measure(
                "keep",
                "Keep",
                new PivotAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Average)),
            Measure("created", "Created", Sum()));
        var placement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, "created"),
                new PivotMeasureValuePlacement(2, "keep")
            },
            PivotValuesAxis.Columns,
            1);
        PivotModelMeasurePreparedMutation fresh = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desired,
            placement,
            metadata,
            existingPending: null);
        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "create update delete values");
        fresh.PrimeUndoContribution("apply_combined_resume", combinedPlan);
        ModelMeasureWorkbookSnapshot? firstAfter = null;
        new PivotMutationCoordinator().Execute(
            fresh.Target.PivotTable,
            fresh.Steps,
            fresh.Refresh,
            () => firstAfter = fresh.Verify());
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = "apply_combined_resume",
            PlanFingerprint = combinedPlan,
            BeforePivotFingerprint = fresh.Pending.BeforePivotFingerprint,
            ExpectedPivotFingerprint = fresh.Pending.ExpectedPivotFingerprint,
            Transitions = fresh.Pending.Transitions.Select(CloneTransition).ToList()
        };

        PivotModelMeasurePreparedMutation resumed = service.PrepareParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desired,
            placement,
            metadata,
            combined,
            new PivotModelMeasureParticipantRetryBinding(
                combinedPlan,
                fresh.Pending.PlanFingerprint));
        resumed.PrimeUndoContribution("apply_combined_resume", combinedPlan);
        ModelMeasureWorkbookSnapshot? retryAfter = null;
        new PivotMutationCoordinator().Execute(
            resumed.Target.PivotTable,
            resumed.Steps,
            resumed.Refresh,
            () => retryAfter = resumed.Verify());
        PivotModelMeasureUndoContribution? contribution = resumed.BuildUndoContribution(
            "apply_combined_resume",
            combinedPlan,
            retryAfter!);

        Assert.NotNull(contribution);
        Assert.Equal(initial.SelectedPivotFingerprint, contribution!.Before.SelectedPivotFingerprint);
        Assert.Equal(
            new[] { keep.LiveFingerprint, remove.LiveFingerprint }.OrderBy(value => value),
            contribution.BeforeOwnedArtifacts.Select(value => value.Fingerprint).OrderBy(value => value));
        PivotModelMeasurePreparedUndo undo = service.PrepareUndoParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            contribution);
        new PivotMutationCoordinator().Execute(
            undo.Target.PivotTable,
            undo.Steps,
            undo.Refresh,
            undo.Verify);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Equal(
            initial.Measures.Select(value => value.LiveFingerprint).OrderBy(value => value),
            gateway.State.Measures.Select(value => value.LiveFingerprint).OrderBy(value => value));
    }

    [Fact]
    public void Apply_DistinguishesRepeatedExistingValuesByPreviewPosition()
    {
        ModelDataFieldSnapshot first = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            1);
        ModelDataFieldSnapshot second = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0",
            2);
        var gateway = new RecordingGateway(Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(first, second)));
        var service = Service(gateway, new RecordingOwnershipStore(Metadata()));
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, second),
                new PivotMeasureValuePlacement(2, "metric"),
                ExistingPlacement(3, first)
            },
            PivotValuesAxis.Columns,
            1);

        PivotModelMeasureApplyResult result = service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            placement);

        Assert.Equal(PivotModelMeasureApplyStatus.Applied, result.Status);
        Assert.Equal(3, gateway.State.SelectedPivot.DataFields.Count);
        Assert.Equal(
            new[] { null, "Metric", null },
            gateway.State.SelectedPivot.DataFields.Select(field => field.ModelMeasureName));
    }

    [Fact]
    public void ProductionApply_MetadataNeverContainsCompiledDaxFormula()
    {
        var workbook = new PivotPlusPersistenceTests.FaultingWorkbook();
        var metadataStore = new PivotPlusWorkbookMetadataStore();
        metadataStore.Save(workbook, Metadata());
        PivotDaxCompilation compilation = Compile(
            Measure("metric", "Metric", new PivotDifferenceExpression(
                Sum(),
                new PivotAggregateExpression(
                    "units",
                    PivotCalculationAggregateFunction.Average))));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()),
            workbook: workbook);
        var service = new PivotModelMeasureMutationService(
            gateway,
            new PivotModelMeasureOwnershipStore(metadataStore),
            new RecordingIdentityResolver(),
            new PivotMutationCoordinator());

        service.Apply(
            workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            GeneratedPlacement("metric"));

        string xml = Assert.Single(workbook.CustomXMLParts.AllXml);
        Assert.DoesNotContain(compilation.Measures.Single().DaxFormula, xml, StringComparison.Ordinal);
        Assert.DoesNotContain("DaxFormula", xml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("measure.host.v1:sha256:", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ProductionOwnershipStore_CommitDoesNotOverwriteConcurrentActiveOwnershipChange()
    {
        var workbook = new PivotPlusPersistenceTests.FaultingWorkbook();
        var metadataStore = new PivotPlusWorkbookMetadataStore();
        metadataStore.Save(workbook, Metadata());
        var ownership = new PivotModelMeasureOwnershipStore(metadataStore);
        PivotPlusPendingSemanticApplyMetadata pending = PendingCreate("Metric");
        PivotModelMeasureOwnershipSession session = ownership.Journal(
            workbook,
            SetupId,
            Target(),
            pending);
        PivotPlusWorkbookMetadata changed = metadataStore.Load(workbook, SetupId)!;
        changed.Artifacts.Add(new PivotPlusOwnedArtifact
        {
            Kind = PivotPlusArtifactKind.NamedSet,
            ArtifactId = "set_user",
            Fingerprint = PivotPlusFingerprint.Create("namedset.host.v1", "changed")
        });
        // Simulate out-of-band Custom XML tampering/concurrent replacement;
        // the public Save path correctly refuses this state while pending.
        workbook.CustomXMLParts
            .SelectByNamespace(PivotPlusWorkbookMetadataStore.NamespaceUri)
            .Item(1)
            .Delete();
        workbook.CustomXMLParts.Seed(metadataStore.Serialize(changed));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            ownership.Commit(
                workbook,
                session,
                Array.Empty<PivotPlusOwnedArtifact>(),
                undo: null));

        Assert.Contains("ownership changed", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            metadataStore.Load(workbook, SetupId)!.Artifacts,
            item => item.ArtifactId == "set_user");
    }

    [Fact]
    public void ProductionOwnershipStore_RestoreDoesNotClearDifferentPendingJournal()
    {
        var workbook = new PivotPlusPersistenceTests.FaultingWorkbook();
        var metadataStore = new PivotPlusWorkbookMetadataStore();
        metadataStore.Save(workbook, Metadata());
        var ownership = new PivotModelMeasureOwnershipStore(metadataStore);
        PivotModelMeasureOwnershipSession session = ownership.Journal(
            workbook,
            SetupId,
            Target(),
            PendingCreate("Metric"));
        PivotPlusWorkbookMetadata changed = metadataStore.Load(workbook, SetupId)!;
        changed.PendingSemanticApply!.PlanFingerprint =
            PivotPlusFingerprint.Create("measure.plan.v1", "different plan");
        metadataStore.Save(workbook, changed);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            ownership.RestoreBase(workbook, session));

        Assert.Contains("journal changed", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(metadataStore.Load(workbook, SetupId)!.PendingSemanticApply);
    }

    [Fact]
    public void ArtifactParticipant_DelegatesValuesLayoutAndRefreshAndExposesTrustedBindings()
    {
        ModelDataFieldSnapshot existing = ExistingField(
            "[Measures].[Sum of Units]",
            "Sum of Units",
            "#,##0");
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected(existing));
        var events = new List<string>();
        var gateway = new RecordingGateway(initial, events);
        var service = Service(gateway, new RecordingOwnershipStore(Metadata()));
        PivotModelMeasureArtifactPreparedMutation prepared = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            Metadata(),
            existingPending: null);
        PivotModelMeasureArtifactBinding binding = prepared.DefinitionBindings["metric"];
        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "artifact participant layout delegation");
        prepared.PrimeUndoContribution("apply_artifact_delegate", combinedPlan);
        var layoutStep = new PivotMutationStep(
            "synthetic combined Values layout",
            () =>
            {
                events.Add("combined-layout");
                LiveModelMeasureSnapshot measure = Assert.Single(gateway.State.Measures);
                gateway.ReplaceSelectedForTest(Selected(
                    GeneratedField(measure, 1),
                    existing));
            },
            () => gateway.ReplaceSelectedForTest(initial.SelectedPivot));
        ModelMeasureWorkbookSnapshot? verified = null;

        new PivotMutationCoordinator().Execute(
            prepared.Target.PivotTable,
            prepared.UpsertSteps.Concat(new[] { layoutStep }).Concat(prepared.DeleteSteps).ToList(),
            prepared.Refresh,
            () => verified = prepared.Verify());

        Assert.NotNull(verified);
        Assert.Equal("metric", binding.DefinitionId);
        Assert.Equal("Metric", binding.HostMeasureName);
        Assert.StartsWith(
            "measure.definition.v1:sha256:",
            binding.DefinitionFingerprint,
            StringComparison.Ordinal);
        Assert.Empty(prepared.DeleteSteps);
        Assert.DoesNotContain("placement", gateway.Events);
        Assert.DoesNotContain("restore-placement", gateway.Events);
        Assert.Equal(0, gateway.PlacementCalls);
        Assert.Equal(0, gateway.RestorePlacementCalls);
        Assert.Equal(1, gateway.RefreshCalls);
        Assert.Equal(2, gateway.State.SelectedPivot.DataFields.Count);
        Assert.Single(prepared.BuildArtifacts(verified!));
        Assert.NotNull(prepared.BuildUndoContribution(
            "apply_artifact_delegate",
            combinedPlan,
            verified!));
    }

    [Fact]
    public void ArtifactParticipant_ResumesPartialCombinedLayoutWithoutUsingPivotHashesOrInventingUndo()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        PivotPlusWorkbookMetadata metadata = Metadata();
        var firstService = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotModelMeasureArtifactPreparedMutation fresh = firstService.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            metadata,
            existingPending: null);
        new PivotMutationCoordinator().Execute(
            fresh.Target.PivotTable,
            fresh.Steps,
            fresh.Refresh,
            () => fresh.Verify());
        LiveModelMeasureSnapshot created = Assert.Single(gateway.State.Measures);
        gateway.ReplaceSelectedForTest(Selected(
            ExistingField("[Measures].[User Value]", "User Value", "0.0"),
            GeneratedField(created, 2)));
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = fresh.ApplyId,
            PlanFingerprint = PivotPlusFingerprint.Create(
                "semantic.combined.v1",
                "partial layout retry"),
            BeforePivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.before-pivot.v1",
                "intentionally unrelated to the selected Values layout"),
            ExpectedPivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.expected-pivot.v1",
                "owned by the layout participant"),
            Transitions = fresh.Transitions.Select(CloneTransition).Concat(new[]
            {
                new PivotPlusSemanticArtifactTransition
                {
                    Kind = PivotPlusArtifactKind.NamedSet,
                    ArtifactId = "PP_Set_Q1",
                    Operation = PivotPlusSemanticArtifactOperation.Create,
                    PlannedDefinitionFingerprint = PivotPlusFingerprint.Create(
                        "namedset.definition.v1",
                        "Q1")
                }
            }).ToList()
        };
        var restartedService = Service(gateway, new RecordingOwnershipStore(metadata));

        PivotModelMeasureArtifactPreparedMutation resumed =
            restartedService.PrepareArtifactParticipant(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                SetupId,
                compilation,
                metadata,
                combined,
                new PivotModelMeasureParticipantRetryBinding(
                    combined.PlanFingerprint,
                    fresh.ParticipantPlanFingerprint));
        ModelMeasureWorkbookSnapshot? verified = null;
        new PivotMutationCoordinator().Execute(
            resumed.Target.PivotTable,
            resumed.Steps,
            resumed.Refresh,
            () => verified = resumed.Verify());

        Assert.Single(resumed.UpsertSteps);
        Assert.Empty(resumed.DeleteSteps);
        Assert.Equal(0, gateway.PlacementCalls);
        Assert.Equal(0, gateway.RestorePlacementCalls);
        Assert.Equal(2, gateway.State.SelectedPivot.DataFields.Count);
        Assert.Null(resumed.BuildUndoContribution(
            combined.ApplyId,
            combined.PlanFingerprint,
            verified!));
    }

    [Fact]
    public void ArtifactParticipant_SameSessionRetryReconcilesCreateUpdateDeleteAndKeepsOriginalUndo()
    {
        PivotDaxCompilation priorKeep = Compile(Measure("keep", "Keep", Sum()));
        PivotDaxCompilation priorRemove = Compile(Measure("remove", "Remove", Sum()));
        LiveModelMeasureSnapshot keep = Live(priorKeep.Measures.Single());
        LiveModelMeasureSnapshot remove = Live(priorRemove.Measures.Single());
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            new[] { keep, remove },
            Selected());
        var gateway = new RecordingGateway(initial);
        PivotPlusWorkbookMetadata metadata = Metadata(Artifact(keep), Artifact(remove));
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotDaxCompilation desired = Compile(
            Measure(
                "keep",
                "Keep",
                new PivotAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Average)),
            Measure("created", "Created", Sum()));
        PivotModelMeasureArtifactPreparedMutation fresh = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desired,
            metadata,
            existingPending: null);
        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "artifact create update delete retry");
        fresh.PrimeUndoContribution("apply_artifact_retry", combinedPlan);
        new PivotMutationCoordinator().Execute(
            fresh.Target.PivotTable,
            fresh.Steps,
            fresh.Refresh,
            () => fresh.Verify());
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = "apply_artifact_retry",
            PlanFingerprint = combinedPlan,
            BeforePivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.before-pivot.v1",
                "layout-owned before"),
            ExpectedPivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.expected-pivot.v1",
                "layout-owned after"),
            Transitions = fresh.Transitions.Select(CloneTransition).ToList()
        };

        PivotModelMeasureArtifactPreparedMutation resumed = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desired,
            metadata,
            combined,
            new PivotModelMeasureParticipantRetryBinding(
                combined.PlanFingerprint,
                fresh.ParticipantPlanFingerprint));
        ModelMeasureWorkbookSnapshot? verified = null;
        new PivotMutationCoordinator().Execute(
            resumed.Target.PivotTable,
            resumed.Steps,
            resumed.Refresh,
            () => verified = resumed.Verify());
        PivotModelMeasureArtifactUndoContribution contribution =
            resumed.BuildUndoContribution(
                combined.ApplyId,
                combined.PlanFingerprint,
                verified!)!;

        Assert.Equal(2, resumed.UpsertSteps.Count);
        Assert.Empty(resumed.DeleteSteps);
        Assert.Equal(
            initial.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value),
            contribution.Before.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value));
        Assert.Equal(
            new[] { keep.LiveFingerprint, remove.LiveFingerprint }.OrderBy(value => value),
            contribution.BeforeOwnedArtifacts.Select(artifact => artifact.Fingerprint).OrderBy(value => value));

        PivotModelMeasureArtifactPreparedUndo undo =
            service.PrepareArtifactUndoParticipant(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                contribution);
        new PivotMutationCoordinator().Execute(
            undo.Target.PivotTable,
            undo.Steps,
            undo.Refresh,
            undo.Verify);
        Assert.Equal(
            initial.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value),
            gateway.State.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value));
        Assert.Equal(0, gateway.PlacementCalls);
        Assert.Equal(0, gateway.RestorePlacementCalls);
    }

    [Fact]
    public void ArtifactParticipant_RejectsStaleParticipantRetryBindingBeforeMutation()
    {
        PivotDaxCompilation compilation = Compile(Measure("metric", "Metric", Sum()));
        var gateway = new RecordingGateway(
            Snapshot(Array.Empty<LiveModelMeasureSnapshot>(), Selected()));
        PivotPlusWorkbookMetadata metadata = Metadata();
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotModelMeasureArtifactPreparedMutation fresh = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            metadata,
            existingPending: null);
        var combined = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = fresh.ApplyId,
            PlanFingerprint = PivotPlusFingerprint.Create(
                "semantic.combined.v1",
                "stale participant binding"),
            BeforePivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.before-pivot.v1",
                "before"),
            ExpectedPivotFingerprint = PivotPlusFingerprint.Create(
                "semantic.expected-pivot.v1",
                "after"),
            Transitions = fresh.Transitions.Select(CloneTransition).ToList()
        };

        Assert.Throws<InvalidOperationException>(() => service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            compilation,
            metadata,
            combined,
            new PivotModelMeasureParticipantRetryBinding(
                combined.PlanFingerprint,
                PivotPlusFingerprint.Create(
                    "measure.artifact-plan.v1",
                    "stale"))));
        Assert.Empty(gateway.State.Measures);
        Assert.DoesNotContain(gateway.Events, item =>
            item.StartsWith("create:", StringComparison.Ordinal));
    }

    [Fact]
    public void ArtifactParticipant_RollbackProofIgnoresLayoutButRejectsArtifactDrift()
    {
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            Selected());
        var gateway = new RecordingGateway(initial);
        var service = Service(gateway, new RecordingOwnershipStore(Metadata()));
        PivotModelMeasureArtifactPreparedMutation prepared = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            Metadata(),
            existingPending: null);
        var layoutStep = new PivotMutationStep(
            "external layout mutation",
            () => gateway.ReplaceSelectedForTest(Selected(
                ExistingField("[Measures].[User Value]", "User Value", "0"))),
            () => { });
        var failureStep = new PivotMutationStep(
            "synthetic later participant failure",
            () => throw new InvalidOperationException("synthetic failure"),
            () => { });

        PivotMutationException failure = Assert.Throws<PivotMutationException>(() =>
            new PivotMutationCoordinator().Execute(
                prepared.Target.PivotTable,
                prepared.UpsertSteps.Concat(new[] { layoutStep, failureStep }).ToList(),
                prepared.Refresh,
                () => prepared.Verify()));

        Assert.True(failure.RollbackCompleted);
        Assert.Empty(gateway.State.Measures);
        Assert.NotEqual(
            initial.SelectedPivotFingerprint,
            gateway.State.SelectedPivotFingerprint);
        prepared.VerifyRollback();
        gateway.AddMeasure(UserMeasure("Unexpected", "=SUM('Facts'[Amount])"));
        Assert.Throws<InvalidOperationException>(prepared.VerifyRollback);
        Assert.Equal(0, gateway.RestorePlacementCalls);
    }

    [Fact]
    public void ArtifactParticipants_ApplyAndUndoArtifactPhasesWithoutLayoutMutation()
    {
        PivotDaxCompilation priorKeep = Compile(Measure("keep", "Keep", Sum()));
        PivotDaxCompilation priorRemove = Compile(Measure("remove", "Remove", Sum()));
        LiveModelMeasureSnapshot keep = Live(priorKeep.Measures.Single());
        LiveModelMeasureSnapshot remove = Live(priorRemove.Measures.Single());
        ModelMeasureWorkbookSnapshot initial = Snapshot(
            new[] { keep, remove },
            Selected(GeneratedField(keep, 1), GeneratedField(remove, 2)));
        var events = new List<string>();
        var gateway = new RecordingGateway(initial, events);
        PivotPlusWorkbookMetadata metadata = Metadata(Artifact(keep), Artifact(remove));
        var service = Service(gateway, new RecordingOwnershipStore(metadata));
        PivotDaxCompilation desired = Compile(
            Measure(
                "keep",
                "Keep",
                new PivotAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Average)),
            Measure("created", "Created", Sum()));
        PivotModelMeasureArtifactPreparedMutation apply = service.PrepareArtifactParticipant(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            desired,
            metadata,
            existingPending: null);
        string combinedPlan = PivotPlusFingerprint.Create(
            "semantic.combined.v1",
            "artifact apply and undo");
        apply.PrimeUndoContribution("apply_artifact_undo", combinedPlan);
        var applyLayout = new PivotMutationStep(
            "external final layout",
            () =>
            {
                events.Add("external-layout-apply");
                LiveModelMeasureSnapshot liveKeep = gateway.State.Measures.Single(item =>
                    item.Name == "Keep");
                LiveModelMeasureSnapshot liveCreated = gateway.State.Measures.Single(item =>
                    item.Name == "Created");
                gateway.ReplaceSelectedForTest(Selected(
                    GeneratedField(liveCreated, 1),
                    GeneratedField(liveKeep, 2)));
            },
            () => gateway.ReplaceSelectedForTest(initial.SelectedPivot));
        ModelMeasureWorkbookSnapshot? after = null;
        new PivotMutationCoordinator().Execute(
            apply.Target.PivotTable,
            apply.UpsertSteps.Concat(new[] { applyLayout }).Concat(apply.DeleteSteps).ToList(),
            apply.Refresh,
            () => after = apply.Verify());
        PivotModelMeasureArtifactUndoContribution contribution =
            apply.BuildUndoContribution(
                "apply_artifact_undo",
                combinedPlan,
                after!)!;

        PivotModelMeasureArtifactPreparedUndo undo =
            service.PrepareArtifactUndoParticipant(
                gateway.Workbook,
                gateway.Pivot,
                Context(),
                contribution);
        ModelPivotUsageSnapshot finalLayout = gateway.State.SelectedPivot;
        var undoLayout = new PivotMutationStep(
            "external prior layout",
            () =>
            {
                events.Add("external-layout-undo");
                gateway.ReplaceSelectedForTest(initial.SelectedPivot);
            },
            () => gateway.ReplaceSelectedForTest(finalLayout));
        new PivotMutationCoordinator().Execute(
            undo.Target.PivotTable,
            undo.UpsertSteps.Concat(new[] { undoLayout }).Concat(undo.DeleteSteps).ToList(),
            undo.Refresh,
            undo.Verify);

        Assert.Equal(2, apply.UpsertSteps.Count);
        Assert.Single(apply.DeleteSteps);
        Assert.Equal(2, undo.UpsertSteps.Count);
        Assert.Single(undo.DeleteSteps);
        Assert.Equal(2, gateway.RefreshCalls);
        Assert.Equal(0, gateway.PlacementCalls);
        Assert.Equal(0, gateway.RestorePlacementCalls);
        Assert.Equal(initial.SelectedPivotFingerprint, gateway.State.SelectedPivotFingerprint);
        Assert.Equal(
            initial.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value),
            gateway.State.Measures.Select(measure => measure.LiveFingerprint).OrderBy(value => value));
    }

    [Fact]
    public void Apply_StoresNativeUndoValuePositionsAsZeroBasedBoundaries()
    {
        ModelDataFieldSnapshot first = ExistingField(
            "[Measures].[First]",
            "First",
            "0",
            1);
        ModelDataFieldSnapshot last = ExistingField(
            "[Measures].[Last]",
            "Last",
            "0",
            256);
        var selected = new ModelPivotUsageSnapshot(
            "Sheet1",
            "PivotTable1",
            isSelectedTarget: true,
            new[] { first, last },
            PivotValuesAxis.Columns,
            valuesPosition: 1);
        var gateway = new RecordingGateway(Snapshot(
            Array.Empty<LiveModelMeasureSnapshot>(),
            selected));
        var ownership = new RecordingOwnershipStore(Metadata());
        var service = Service(gateway, ownership);
        var placement = new PivotMeasurePlacementPlan(
            new PivotMeasureValuePlacement[]
            {
                ExistingPlacement(1, first),
                ExistingPlacement(2, last),
                new PivotMeasureValuePlacement(3, "metric")
            },
            PivotValuesAxis.Columns,
            valuesPosition: 1);

        service.Apply(
            gateway.Workbook,
            gateway.Pivot,
            Context(),
            SetupId,
            Compile(Measure("metric", "Metric", Sum())),
            placement);

        Assert.NotNull(ownership.Base.Undo);
        Assert.Equal(
            new[] { 0, 255 },
            ownership.Base.Undo!.PreviousFieldPlacements
                .Select(field => field.Position)
                .OrderBy(position => position));
    }

    [Fact]
    public void DaxReferenceScanner_IgnoresStringsCommentsAndQuotedTableNames()
    {
        IReadOnlyCollection<string> references = DaxMeasureReferenceScanner.ReadPossibleReferences(
            "=\"[String]\" + 'Table [Not Measure]'[Column] + [Actual]]Name] " +
            "// [Line Comment]\n/* [Block Comment] */ + [Real Measure]");

        Assert.Contains("Column", references, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("Actual]Name", references, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("Real Measure", references, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("String", references, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("Line Comment", references, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain("Block Comment", references, StringComparer.OrdinalIgnoreCase);
    }

    private const string SetupId = "setup_measure";

    private static PivotModelMeasureMutationService Service(
        RecordingGateway gateway,
        RecordingOwnershipStore ownership,
        RecordingIdentityResolver? identity = null)
    {
        return new PivotModelMeasureMutationService(
            gateway,
            ownership,
            identity ?? new RecordingIdentityResolver(),
            new PivotMutationCoordinator());
    }

    private static PivotDaxCompilation Compile(params PivotMeasureDefinition[] measures)
    {
        return PivotDaxCompiler.Compile(new PivotMeasureSetDefinition(
            new PivotModelSchema(new[]
            {
                new PivotModelTableSchema(
                    "fact",
                    "Fact",
                    new[]
                    {
                        new PivotModelFieldSchema(
                            "amount",
                            "Amount",
                            PivotModelDataType.DecimalNumber),
                        new PivotModelFieldSchema(
                            "units",
                            "Units",
                            PivotModelDataType.WholeNumber)
                    })
            }),
            measures));
    }

    private static PivotMeasureDefinition Measure(
        string id,
        string caption,
        PivotCalculationExpression expression)
    {
        return new PivotMeasureDefinition(
            id,
            caption,
            "fact",
            new PivotMeasureFormat(
                PivotMeasureFormatKind.DecimalNumber,
                decimalPlaces: 2,
                useThousandsSeparator: true),
            expression);
    }

    private static PivotAggregateExpression Sum()
    {
        return new PivotAggregateExpression(
            "amount",
            PivotCalculationAggregateFunction.Sum);
    }

    private static PivotTableContext Context()
    {
        return new PivotTableContext(
            new PivotLayoutDefinition(
                Target(),
                new PivotSourceDescriptor(
                    PivotSourceKind.DataModel,
                    "Workbook Data Model",
                    PivotCapability.DataModel |
                    PivotCapability.ModelMeasures |
                    PivotCapability.NativeFieldPlacement |
                    PivotCapability.Refresh,
                    modelTableName: "Fact"),
                Array.Empty<PivotFieldDescriptor>(),
                Array.Empty<PivotFieldPlacement>(),
                clearAll: true),
            isConnected: true,
            sourceFieldsComplete: true);
    }

    private static PivotTargetIdentity Target()
    {
        return new PivotTargetIdentity("workbook_1", "Sheet1", "PivotTable1");
    }

    private static PivotMeasurePlacementPlan GeneratedPlacement(string definitionId)
    {
        return new PivotMeasurePlacementPlan(
            new[] { new PivotMeasureValuePlacement(1, definitionId) },
            PivotValuesAxis.Automatic,
            1);
    }

    private static PivotMeasureValuePlacement ExistingPlacement(
        int position,
        ModelDataFieldSnapshot field)
    {
        return new PivotMeasureValuePlacement(
            position,
            new PivotExistingDataFieldIdentity(
                field.UniqueName,
                field.CaptionFingerprint,
                PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                    field.NumberFormat),
                field.Position));
    }

    private static LiveModelMeasureSnapshot Live(
        OwnedPivotMeasureDefinition measure,
        string? descriptionOverride = null)
    {
        var format = Format(measure.Format);
        string description = descriptionOverride ??
            PivotModelMeasureCanonical.CreateDescriptionMarker(
                SetupId,
                measure.DefinitionId,
                measure.DefinitionFingerprint);
        const string lineage = "model.table.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        string fingerprint = PivotModelMeasureCanonical.CreateLiveFingerprint(
            measure.GeneratedMeasureName,
            measure.HomeTableName,
            lineage,
            measure.DaxFormula,
            description,
            format);
        return new LiveModelMeasureSnapshot(
            measure.GeneratedMeasureName,
            measure.HomeTableName,
            lineage,
            measure.DaxFormula,
            description,
            format,
            fingerprint);
    }

    private static LiveModelMeasureSnapshot UserMeasure(string name, string formula)
    {
        var format = new ModelMeasureFormatSnapshot(
            ExcelModelMeasureFormatKind.DecimalNumber,
            2,
            true);
        const string lineage = "model.table.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        string fingerprint = PivotModelMeasureCanonical.CreateLiveFingerprint(
            name,
            "Fact",
            lineage,
            formula,
            "user",
            format);
        return new LiveModelMeasureSnapshot(
            name,
            "Fact",
            lineage,
            formula,
            "user",
            format,
            fingerprint);
    }

    private static LiveModelMeasureSnapshot WithFormula(
        LiveModelMeasureSnapshot source,
        string formula)
    {
        string fingerprint = PivotModelMeasureCanonical.CreateLiveFingerprint(
            source.Name,
            source.AssociatedTableName,
            source.AssociatedTableLineageFingerprint,
            formula,
            source.Description,
            source.Format);
        return new LiveModelMeasureSnapshot(
            source.Name,
            source.AssociatedTableName,
            source.AssociatedTableLineageFingerprint,
            formula,
            source.Description,
            source.Format,
            fingerprint);
    }

    private static ModelDataFieldSnapshot WithDisplay(
        ModelDataFieldSnapshot source,
        string caption,
        string numberFormat)
    {
        return new ModelDataFieldSnapshot(
            source.UniqueName,
            caption,
            PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(caption),
            numberFormat,
            source.Position,
            source.IsModelMeasure,
            source.ModelMeasureName);
    }

    private static ModelMeasureFormatSnapshot Format(PivotMeasureFormat format)
    {
        return new ModelMeasureFormatSnapshot(
            format.Kind == PivotMeasureFormatKind.WholeNumber
                ? ExcelModelMeasureFormatKind.WholeNumber
                : ExcelModelMeasureFormatKind.DecimalNumber,
            format.DecimalPlaces,
            format.UseThousandsSeparator,
            format.CurrencySymbolOrCode);
    }

    private static PivotPlusOwnedArtifact Artifact(LiveModelMeasureSnapshot measure)
    {
        return new PivotPlusOwnedArtifact
        {
            Kind = PivotPlusArtifactKind.Measure,
            ArtifactId = measure.Name,
            Fingerprint = measure.LiveFingerprint
        };
    }

    private static PivotPlusSemanticArtifactTransition CloneTransition(
        PivotPlusSemanticArtifactTransition source)
    {
        return new PivotPlusSemanticArtifactTransition
        {
            Kind = source.Kind,
            ArtifactId = source.ArtifactId,
            Operation = source.Operation,
            BeforeLiveFingerprint = source.BeforeLiveFingerprint,
            PlannedDefinitionFingerprint = source.PlannedDefinitionFingerprint
        };
    }

    private static PivotPlusWorkbookMetadata Metadata(params PivotPlusOwnedArtifact[] artifacts)
    {
        return new PivotPlusWorkbookMetadata
        {
            SetupId = SetupId,
            TargetWorksheetName = Target().WorksheetName,
            TargetPivotTableName = Target().PivotTableName,
            Artifacts = artifacts.ToList()
        };
    }

    private static PivotPlusPendingSemanticApplyMetadata PendingCreate(string artifactId)
    {
        return new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = "apply_measure",
            PlanFingerprint = PivotPlusFingerprint.Create("measure.plan.v1", "plan"),
            BeforePivotFingerprint = PivotPlusFingerprint.Create("measure.pivot.v1", "before"),
            ExpectedPivotFingerprint = PivotPlusFingerprint.Create("measure.pivot.v1", "after"),
            Transitions = new List<PivotPlusSemanticArtifactTransition>
            {
                new PivotPlusSemanticArtifactTransition
                {
                    Kind = PivotPlusArtifactKind.Measure,
                    ArtifactId = artifactId,
                    Operation = PivotPlusSemanticArtifactOperation.Create,
                    PlannedDefinitionFingerprint = PivotPlusFingerprint.Create(
                        "measure.definition.v1",
                        artifactId)
                }
            }
        };
    }

    private static ModelDataFieldSnapshot ExistingField(
        string uniqueName,
        string caption,
        string numberFormat,
        int position = 1)
    {
        return new ModelDataFieldSnapshot(
            uniqueName,
            caption,
            PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(caption),
            numberFormat,
            position,
            isModelMeasure: false);
    }

    private static ModelDataFieldSnapshot GeneratedField(
        LiveModelMeasureSnapshot measure,
        int position)
    {
        return new ModelDataFieldSnapshot(
            MeasureUniqueName(measure.Name),
            measure.Name,
            PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(measure.Name),
            "#,##0.00",
            position,
            isModelMeasure: true,
            modelMeasureName: measure.Name);
    }

    private static ModelPivotUsageSnapshot Selected(params ModelDataFieldSnapshot[] fields)
    {
        return new ModelPivotUsageSnapshot(
            "Sheet1",
            "PivotTable1",
            isSelectedTarget: true,
            Reposition(fields),
            fields.Length >= 2 ? PivotValuesAxis.Columns : PivotValuesAxis.Automatic,
            1);
    }

    private static ModelPivotUsageSnapshot OtherPivot(params ModelDataFieldSnapshot[] fields)
    {
        return new ModelPivotUsageSnapshot(
            "Sheet2",
            "PivotTable2",
            isSelectedTarget: false,
            Reposition(fields),
            fields.Length >= 2 ? PivotValuesAxis.Columns : PivotValuesAxis.Automatic,
            1);
    }

    private static ModelMeasureWorkbookSnapshot Snapshot(
        IEnumerable<LiveModelMeasureSnapshot> measures,
        params ModelPivotUsageSnapshot[] usages)
    {
        ModelPivotUsageSnapshot selected = usages.Single(item => item.IsSelectedTarget);
        return new ModelMeasureWorkbookSnapshot(
            measures,
            usages,
            PivotModelMeasureCanonical.CreatePivotFingerprint(selected));
    }

    private static IReadOnlyList<ModelDataFieldSnapshot> Reposition(
        IEnumerable<ModelDataFieldSnapshot> fields)
    {
        return fields.Select((field, index) => new ModelDataFieldSnapshot(
            field.UniqueName,
            field.Caption,
            field.CaptionFingerprint,
            field.NumberFormat,
            index + 1,
            field.IsModelMeasure,
            field.ModelMeasureName)).ToList();
    }

    private static string MeasureUniqueName(string name)
    {
        return "[Measures].[" + name.Replace("]", "]]", StringComparison.Ordinal) + "]";
    }

    private sealed class RecordingGateway : IPivotModelMeasureGateway
    {
        private readonly List<string> events;

        public RecordingGateway(
            ModelMeasureWorkbookSnapshot initial,
            List<string>? sharedEvents = null,
            object? workbook = null)
        {
            State = initial;
            events = sharedEvents ?? new List<string>();
            Workbook = workbook ?? new object();
        }

        public object Workbook { get; }

        public FakePivot Pivot { get; } = new FakePivot();

        public ModelMeasureWorkbookSnapshot State { get; private set; }

        public IReadOnlyList<string> Events => events;

        public int RefreshCalls { get; private set; }

        public int ThrowOnRefreshCall { get; set; }

        public string? ThrowAfterCreateName { get; set; }

        public string? ThrowBeforeRestoreName { get; set; }

        public int ThrowAfterPlacementCall { get; set; }

        public int ThrowAfterRestorePlacementCall { get; set; }

        public int PlacementCalls { get; private set; }

        public int RestorePlacementCalls { get; private set; }

        public Action<RecordingGateway, int>? RefreshMutation { get; set; }

        public Action<RecordingGateway, int>? RestorePlacementMutation { get; set; }

        public BoundModelMeasureTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context)
        {
            Assert.Same(Workbook, workbook);
            Assert.Same(Pivot, pivotTable);
            events.Add("bind");
            return new BoundModelMeasureTarget(
                Workbook,
                Pivot,
                new object(),
                new object(),
                context.Definition.Target);
        }

        public ModelMeasureWorkbookSnapshot Capture(BoundModelMeasureTarget target)
        {
            events.Add("capture");
            return State;
        }

        public LiveModelMeasureSnapshot CreateMeasure(
            BoundModelMeasureTarget target,
            DesiredModelMeasure definition)
        {
            events.Add("create:" + definition.Name);
            if (State.Measures.Any(item => string.Equals(
                    item.Name,
                    definition.Name,
                    StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException("duplicate create");
            }

            LiveModelMeasureSnapshot value = Live(definition);
            SetMeasures(State.Measures.Concat(new[] { value }));
            if (string.Equals(
                    ThrowAfterCreateName,
                    definition.Name,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("create committed before failure");
            }

            return value;
        }

        public LiveModelMeasureSnapshot UpdateMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before,
            DesiredModelMeasure definition)
        {
            events.Add("update:" + definition.Name);
            DemandLive(before);
            LiveModelMeasureSnapshot value = Live(definition);
            SetMeasures(State.Measures.Select(item => string.Equals(
                    item.Name,
                    before.Name,
                    StringComparison.OrdinalIgnoreCase)
                ? value
                : item));
            return value;
        }

        public LiveModelMeasureSnapshot RestoreMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot before)
        {
            events.Add("restore:" + before.Name);
            if (string.Equals(
                    ThrowBeforeRestoreName,
                    before.Name,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("restore failed before mutation");
            }

            if (State.Measures.Any(item => string.Equals(
                    item.Name,
                    before.Name,
                    StringComparison.OrdinalIgnoreCase)))
            {
                SetMeasures(State.Measures.Select(item => string.Equals(
                        item.Name,
                        before.Name,
                        StringComparison.OrdinalIgnoreCase)
                    ? before
                    : item));
            }
            else
            {
                SetMeasures(State.Measures.Concat(new[] { before }));
            }

            return before;
        }

        public void DeleteMeasure(
            BoundModelMeasureTarget target,
            LiveModelMeasureSnapshot expected)
        {
            events.Add("delete:" + expected.Name);
            DemandLive(expected);
            SetMeasures(State.Measures.Where(item => !string.Equals(
                item.Name,
                expected.Name,
                StringComparison.OrdinalIgnoreCase)));
        }

        public void ApplyPlacement(
            BoundModelMeasureTarget target,
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelMeasureWorkbookSnapshot before)
        {
            events.Add("placement");
            PlacementCalls++;
            var prior = before.SelectedPivot.DataFields.ToDictionary(
                item => item.UniqueName + "\u001f" + item.CaptionFingerprint + "\u001f" + item.Position,
                StringComparer.OrdinalIgnoreCase);
            var fields = new List<ModelDataFieldSnapshot>();
            foreach (PivotMeasureValuePlacement item in placement.Values.OrderBy(value => value.Position))
            {
                if (item.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[item.DefinitionId!];
                    ModelDataFieldSnapshot? existing = before.SelectedPivot.DataFields.SingleOrDefault(field =>
                        string.Equals(
                            field.ModelMeasureName,
                            definition.Name,
                            StringComparison.OrdinalIgnoreCase));
                    fields.Add(existing == null
                        ? new ModelDataFieldSnapshot(
                            MeasureUniqueName(definition.Name),
                            definition.Name,
                            PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(definition.Name),
                            "#,##0.00",
                            item.Position,
                            isModelMeasure: true,
                            modelMeasureName: definition.Name)
                        : WithPosition(existing, item.Position));
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = item.ExistingDataField!;
                    fields.Add(WithPosition(
                        prior[identity.UniqueName + "\u001f" + identity.CurrentCaptionFingerprint +
                              "\u001f" + identity.CurrentPosition],
                        item.Position));
                }
            }

            ReplaceSelected(new ModelPivotUsageSnapshot(
                "Sheet1",
                "PivotTable1",
                isSelectedTarget: true,
                fields,
                placement.ValuesAxis,
                placement.ValuesPosition));
            if (PlacementCalls == ThrowAfterPlacementCall)
            {
                throw new InvalidOperationException("placement committed before failure");
            }
        }

        public void RestorePlacement(
            BoundModelMeasureTarget target,
            ModelPivotUsageSnapshot before)
        {
            events.Add("restore-placement");
            RestorePlacementCalls++;
            ReplaceSelected(before);
            RestorePlacementMutation?.Invoke(this, RestorePlacementCalls);
            if (RestorePlacementCalls == ThrowAfterRestorePlacementCall)
            {
                throw new InvalidOperationException("placement restore failed after mutation");
            }
        }

        public void Refresh(BoundModelMeasureTarget target)
        {
            RefreshCalls++;
            events.Add("refresh:" + RefreshCalls);
            RefreshMutation?.Invoke(this, RefreshCalls);
            if (RefreshCalls == ThrowOnRefreshCall)
            {
                throw new InvalidOperationException("refresh failed");
            }
        }

        public void AddOtherPivot(ModelPivotUsageSnapshot usage)
        {
            if (usage.IsSelectedTarget)
            {
                throw new ArgumentException("An external usage cannot replace the selected target.", nameof(usage));
            }

            State = new ModelMeasureWorkbookSnapshot(
                State.Measures,
                State.PivotUsages.Concat(new[] { usage }),
                State.SelectedPivotFingerprint);
        }

        public void AddMeasure(LiveModelMeasureSnapshot measure)
        {
            SetMeasures(State.Measures.Concat(new[] { measure }));
        }

        public void ReplaceMeasure(LiveModelMeasureSnapshot measure)
        {
            SetMeasures(State.Measures.Select(item => string.Equals(
                    item.Name,
                    measure.Name,
                    StringComparison.OrdinalIgnoreCase)
                ? measure
                : item));
        }

        public void ReplaceSelectedForTest(ModelPivotUsageSnapshot selected)
        {
            ReplaceSelected(selected);
        }

        private void DemandLive(LiveModelMeasureSnapshot expected)
        {
            LiveModelMeasureSnapshot current = State.Measures.Single(item => string.Equals(
                item.Name,
                expected.Name,
                StringComparison.OrdinalIgnoreCase));
            if (!string.Equals(current.LiveFingerprint, expected.LiveFingerprint, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("live fingerprint changed");
            }
        }

        private static LiveModelMeasureSnapshot Live(DesiredModelMeasure definition)
        {
            ModelMeasureFormatSnapshot format = Format(definition.Format);
            const string lineage = "model.table.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            string fingerprint = PivotModelMeasureCanonical.CreateLiveFingerprint(
                definition.Name,
                definition.HomeTableName,
                lineage,
                definition.Formula,
                definition.DescriptionMarker,
                format);
            return new LiveModelMeasureSnapshot(
                definition.Name,
                definition.HomeTableName,
                lineage,
                definition.Formula,
                definition.DescriptionMarker,
                format,
                fingerprint);
        }

        private void SetMeasures(IEnumerable<LiveModelMeasureSnapshot> measures)
        {
            State = new ModelMeasureWorkbookSnapshot(
                measures,
                State.PivotUsages,
                State.SelectedPivotFingerprint);
        }

        private void ReplaceSelected(ModelPivotUsageSnapshot selected)
        {
            State = new ModelMeasureWorkbookSnapshot(
                State.Measures,
                State.PivotUsages.Select(item => item.IsSelectedTarget ? selected : item),
                PivotModelMeasureCanonical.CreatePivotFingerprint(selected));
        }

        private static ModelDataFieldSnapshot WithPosition(
            ModelDataFieldSnapshot field,
            int position)
        {
            return new ModelDataFieldSnapshot(
                field.UniqueName,
                field.Caption,
                field.CaptionFingerprint,
                field.NumberFormat,
                position,
                field.IsModelMeasure,
                field.ModelMeasureName);
        }
    }

    private sealed class RecordingOwnershipStore : IPivotModelMeasureOwnershipStore
    {
        private readonly List<string> events;

        public RecordingOwnershipStore(
            PivotPlusWorkbookMetadata metadata,
            List<string>? sharedEvents = null)
        {
            Base = Clone(metadata);
            events = sharedEvents ?? new List<string>();
        }

        public PivotPlusWorkbookMetadata Base { get; private set; }

        public PivotPlusPendingSemanticApplyMetadata? Pending { get; private set; }

        public int JournalCalls { get; private set; }

        public int RestoreCalls { get; private set; }

        public bool ThrowOnCommit { get; set; }

        public bool CommitThenThrow { get; set; }

        public int CommitCalls { get; private set; }

        public PivotPlusWorkbookMetadata ReadBase(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            out PivotPlusPendingSemanticApplyMetadata? existingPending)
        {
            existingPending = Pending;
            return Clone(Base);
        }

        public PivotModelMeasureOwnershipSession Journal(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotPlusPendingSemanticApplyMetadata pending)
        {
            JournalCalls++;
            events.Add("journal");
            if (Pending != null && !string.Equals(
                    Pending.PlanFingerprint,
                    pending.PlanFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException("different pending plan");
            }

            Pending = pending;
            return new PivotModelMeasureOwnershipSession(
                Clone(Base),
                pending,
                resumed: existingPending());

            bool existingPending()
            {
                return JournalCalls > 1;
            }
        }

        public void Commit(
            object workbook,
            PivotModelMeasureOwnershipSession session,
            IReadOnlyList<PivotPlusOwnedArtifact> measures,
            PivotPlusUndoMetadata? undo)
        {
            events.Add("commit");
            CommitCalls++;
            if (ThrowOnCommit)
            {
                throw new InvalidOperationException("commit failed");
            }

            Base = Clone(session.BaseMetadata);
            Base.Artifacts = Base.Artifacts
                .Where(item => item.Kind != PivotPlusArtifactKind.Measure)
                .Concat(measures.Select(Clone))
                .ToList();
            Base.Undo = undo;
            Pending = null;
            if (CommitThenThrow)
            {
                throw new InvalidOperationException("commit reported failure after save");
            }
        }

        public void RestoreBase(object workbook, PivotModelMeasureOwnershipSession session)
        {
            events.Add("restore-base");
            RestoreCalls++;
            Base = Clone(session.BaseMetadata);
            Pending = null;
        }

        private static PivotPlusWorkbookMetadata Clone(PivotPlusWorkbookMetadata source)
        {
            return new PivotPlusWorkbookMetadata
            {
                SchemaVersion = source.SchemaVersion,
                SetupId = source.SetupId,
                TargetWorksheetName = source.TargetWorksheetName,
                TargetPivotTableName = source.TargetPivotTableName,
                Artifacts = source.Artifacts.Select(Clone).ToList(),
                Undo = source.Undo
            };
        }

        private static PivotPlusOwnedArtifact Clone(PivotPlusOwnedArtifact source)
        {
            return new PivotPlusOwnedArtifact
            {
                Kind = source.Kind,
                ArtifactId = source.ArtifactId,
                Fingerprint = source.Fingerprint
            };
        }
    }

    private sealed class RecordingIdentityResolver : IWorkbookIdentityResolver
    {
        public int PersistCalls { get; private set; }

        public string Resolve(object workbook)
        {
            return Target().WorkbookId;
        }

        public void Persist(object workbook, string expectedWorkbookId)
        {
            Assert.Equal(Target().WorkbookId, expectedWorkbookId);
            PersistCalls++;
        }
    }

    public sealed class FakePivot
    {
        public bool ManualUpdate { get; set; }
    }
}
