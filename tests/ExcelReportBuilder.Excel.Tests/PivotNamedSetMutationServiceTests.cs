using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotNamedSetMutationServiceTests
{
    private const string SetupId = "setup_1";
    private const string SourceFingerprint =
        "pivot.source.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    [Fact]
    public void Prepares_create_places_nothing_and_builds_the_final_owned_receipt()
    {
        PivotMdxCompilation compilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        var gateway = new RecordingGateway(SourceFingerprint);
        var service = new PivotNamedSetMutationService(gateway);

        PivotNamedSetPreparedMutation prepared = service.PrepareParticipant(
            new object(),
            new object(),
            Context(),
            SetupId,
            compilation,
            Metadata(),
            existingPending: null);

        PivotPlusSemanticArtifactTransition transition = Assert.Single(prepared.Transitions);
        Assert.Equal(PivotPlusSemanticArtifactOperation.Create, transition.Operation);
        Assert.Single(prepared.UpsertSteps);
        Assert.Empty(prepared.DeleteSteps);

        Run(prepared.UpsertSteps);
        Assert.Equal(0, Assert.Single(gateway.Artifacts).Orientation);
        gateway.SetOrientation(1);
        PivotNamedSetWorkbookSnapshot verified = prepared.Verify();
        PivotPlusOwnedArtifact receipt = Assert.Single(prepared.BuildArtifacts(verified));

        Assert.Equal(PivotPlusArtifactKind.NamedSet, receipt.Kind);
        Assert.Equal("[PivotTablePlus_setup_rows]", receipt.ArtifactId);
        Assert.Equal(Assert.Single(gateway.Artifacts).LiveFingerprint, receipt.Fingerprint);
        Assert.Equal(0, gateway.RefreshCalls);
    }

    [Fact]
    public void Refuses_to_replace_a_definition_under_the_same_owned_host_name()
    {
        PivotMdxCompilation original = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        PivotMdxCompilation changed = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(
                caption: "Changed",
                sourceFingerprint: SourceFingerprint);
        DesiredPivotNamedSet liveDefinition = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired(SetupId, original));
        var gateway = new RecordingGateway(SourceFingerprint);
        gateway.Artifacts.Add(RecordingGateway.Live(liveDefinition, 1));
        PivotPlusWorkbookMetadata metadata = Metadata();
        metadata.Artifacts.Add(new PivotPlusOwnedArtifact
        {
            Kind = PivotPlusArtifactKind.NamedSet,
            ArtifactId = liveDefinition.Name,
            Fingerprint = gateway.Artifacts.Single().LiveFingerprint
        });

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            new PivotNamedSetMutationService(gateway).PrepareParticipant(
                new object(),
                new object(),
                Context(),
                SetupId,
                changed,
                metadata,
                existingPending: null));

        Assert.Contains("new compiler-generated", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, gateway.CreateCalls);
        Assert.Equal(0, gateway.DeleteCalls);
    }

    [Fact]
    public void Deletes_only_after_layout_hides_the_old_set_and_rollback_restores_it_hidden()
    {
        PivotMdxCompilation oldCompilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        DesiredPivotNamedSet oldDefinition = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired(SetupId, oldCompilation));
        var gateway = new RecordingGateway(SourceFingerprint);
        LivePivotNamedSetSnapshot original = RecordingGateway.Live(oldDefinition, 1);
        gateway.Artifacts.Add(original);
        PivotPlusWorkbookMetadata metadata = Metadata();
        metadata.Artifacts.Add(new PivotPlusOwnedArtifact
        {
            Kind = PivotPlusArtifactKind.NamedSet,
            ArtifactId = original.Name,
            Fingerprint = original.LiveFingerprint
        });
        PivotMdxCompilation replacement = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(
                generatedName: "[PivotTablePlus_setup_rows_v2]",
                sourceFingerprint: SourceFingerprint);

        PivotNamedSetPreparedMutation prepared =
            new PivotNamedSetMutationService(gateway).PrepareParticipant(
                new object(),
                new object(),
                Context(),
                SetupId,
                replacement,
                metadata,
                existingPending: null);

        Run(prepared.UpsertSteps);
        Assert.Throws<InvalidOperationException>(() => Run(prepared.DeleteSteps));
        gateway.SetOrientation(0, original.Name);
        Run(prepared.DeleteSteps);
        Assert.DoesNotContain(gateway.Artifacts, item => item.Name == original.Name);

        prepared.DeleteSteps.Single().Rollback();
        LivePivotNamedSetSnapshot restored = gateway.Artifacts.Single(item =>
            item.Name == original.Name);
        Assert.Equal(0, restored.Orientation);
    }

    [Fact]
    public void Pending_transition_slice_must_match_exactly()
    {
        PivotMdxCompilation compilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        var pending = new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = "apply_1",
            PlanFingerprint = Fingerprint("plan"),
            BeforePivotFingerprint = Fingerprint("before"),
            ExpectedPivotFingerprint = Fingerprint("after"),
            Transitions = new List<PivotPlusSemanticArtifactTransition>
            {
                new()
                {
                    Kind = PivotPlusArtifactKind.NamedSet,
                    ArtifactId = "[wrong]",
                    Operation = PivotPlusSemanticArtifactOperation.Create,
                    PlannedDefinitionFingerprint = Fingerprint("definition")
                }
            }
        };

        Assert.Throws<InvalidOperationException>(() =>
            new PivotNamedSetMutationService(new RecordingGateway(SourceFingerprint))
                .PrepareParticipant(
                    new object(),
                    new object(),
                    Context(),
                    SetupId,
                    compilation,
                    Metadata(),
                pending));
    }

    [Fact]
    public void Pending_create_already_visible_is_replayed_without_a_duplicate()
    {
        PivotMdxCompilation compilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        DesiredPivotNamedSet desired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired(SetupId, compilation));
        var gateway = new RecordingGateway(SourceFingerprint);
        gateway.Artifacts.Add(RecordingGateway.Live(desired, 1));
        var transition = new PivotPlusSemanticArtifactTransition
        {
            Kind = PivotPlusArtifactKind.NamedSet,
            ArtifactId = desired.Name,
            Operation = PivotPlusSemanticArtifactOperation.Create,
            PlannedDefinitionFingerprint = desired.DefinitionFingerprint
        };
        PivotPlusPendingSemanticApplyMetadata pending = Pending(transition);

        PivotNamedSetPreparedMutation prepared =
            new PivotNamedSetMutationService(gateway).PrepareParticipant(
                new object(),
                new object(),
                Context(),
                SetupId,
                compilation,
                Metadata(),
                pending);
        Run(prepared.UpsertSteps);
        prepared.Verify();

        Assert.Equal(0, gateway.CreateCalls);
        Assert.Single(gateway.Artifacts);
    }

    [Fact]
    public void Pending_delete_already_absent_is_verified_without_guessing_a_restore()
    {
        PivotMdxCompilation oldCompilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(sourceFingerprint: SourceFingerprint);
        DesiredPivotNamedSet prior = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired(SetupId, oldCompilation));
        LivePivotNamedSetSnapshot priorLive = RecordingGateway.Live(prior, 1);
        PivotPlusWorkbookMetadata metadata = Metadata();
        metadata.Artifacts.Add(new PivotPlusOwnedArtifact
        {
            Kind = PivotPlusArtifactKind.NamedSet,
            ArtifactId = prior.Name,
            Fingerprint = priorLive.LiveFingerprint
        });
        PivotMdxCompilation desiredCompilation = PivotNamedSetCanonicalTests
            .CompileDefaultMemberSet(
                generatedName: "[PivotTablePlus_setup_rows_v2]",
                sourceFingerprint: SourceFingerprint);
        DesiredPivotNamedSet desired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired(SetupId, desiredCompilation));
        var gateway = new RecordingGateway(SourceFingerprint);
        gateway.Artifacts.Add(RecordingGateway.Live(desired, 1));
        PivotPlusPendingSemanticApplyMetadata pending = Pending(
            new PivotPlusSemanticArtifactTransition
            {
                Kind = PivotPlusArtifactKind.NamedSet,
                ArtifactId = desired.Name,
                Operation = PivotPlusSemanticArtifactOperation.Create,
                PlannedDefinitionFingerprint = desired.DefinitionFingerprint
            },
            new PivotPlusSemanticArtifactTransition
            {
                Kind = PivotPlusArtifactKind.NamedSet,
                ArtifactId = prior.Name,
                Operation = PivotPlusSemanticArtifactOperation.Delete,
                BeforeLiveFingerprint = priorLive.LiveFingerprint,
                PlannedDefinitionFingerprint = priorLive.LiveFingerprint
            });

        PivotNamedSetPreparedMutation prepared =
            new PivotNamedSetMutationService(gateway).PrepareParticipant(
                new object(),
                new object(),
                Context(),
                SetupId,
                desiredCompilation,
                metadata,
                pending);
        Run(prepared.UpsertSteps);
        Run(prepared.DeleteSteps);
        prepared.Verify();

        Assert.Equal(0, gateway.CreateCalls);
        Assert.Equal(0, gateway.DeleteCalls);
    }

    private static void Run(IEnumerable<PivotMutationStep> steps)
    {
        foreach (PivotMutationStep step in steps) step.Apply();
    }

    private static PivotPlusWorkbookMetadata Metadata()
    {
        return new PivotPlusWorkbookMetadata
        {
            SetupId = SetupId,
            TargetWorksheetName = "Sheet1",
            TargetPivotTableName = "PivotTable1"
        };
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
                    PivotCapability.NamedSets |
                    PivotCapability.AsymmetricAxes |
                    PivotCapability.Refresh,
                    modelTableName: "Fact"),
                Array.Empty<PivotFieldDescriptor>(),
                Array.Empty<PivotFieldPlacement>(),
                clearAll: true),
            isConnected: true,
            sourceFieldsComplete: true);
    }

    private static PivotTargetIdentity Target() =>
        new("workbook_1", "Sheet1", "PivotTable1");

    private static string Fingerprint(string value) =>
        PivotPlusFingerprint.Create("test.v1", value);

    private static PivotPlusPendingSemanticApplyMetadata Pending(
        params PivotPlusSemanticArtifactTransition[] transitions)
    {
        return new PivotPlusPendingSemanticApplyMetadata
        {
            ApplyId = "apply_1",
            PlanFingerprint = Fingerprint("plan"),
            BeforePivotFingerprint = Fingerprint("before"),
            ExpectedPivotFingerprint = Fingerprint("after"),
            Transitions = transitions.ToList()
        };
    }

    private sealed class RecordingGateway : IPivotNamedSetGateway
    {
        private const string Lineage =
            "namedset.model-lineage.v1:sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
        private readonly string sourceFingerprint;

        public RecordingGateway(string sourceFingerprint)
        {
            this.sourceFingerprint = sourceFingerprint;
        }

        public List<LivePivotNamedSetSnapshot> Artifacts { get; } = new();
        public int CreateCalls { get; private set; }
        public int DeleteCalls { get; private set; }
        public int RefreshCalls { get; private set; }

        public BoundPivotNamedSetTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context) =>
            new(workbook, pivotTable, new object(), new object(), Target());

        public PivotNamedSetSchemaDiscoveryResult DiscoverSchema(
            BoundPivotNamedSetTarget target) => throw new NotSupportedException();

        public PivotNamedSetWorkbookSnapshot Capture(BoundPivotNamedSetTarget target)
        {
            return new PivotNamedSetWorkbookSnapshot(
                new[]
                {
                    new PivotNamedSetPivotSnapshot(
                        "Sheet1",
                        "PivotTable1",
                        true,
                        Artifacts.ToArray(),
                        Artifacts.Select(item => new PivotCalculatedMemberReferenceSnapshot(
                            "Sheet1",
                            "PivotTable1",
                            item.Name,
                            1,
                            item.RawFormula,
                            true)),
                        true,
                        Fingerprint("pivot-" + Artifacts.Count))
                },
                sourceFingerprint,
                Lineage);
        }

        public LivePivotNamedSetSnapshot CreateSet(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition)
        {
            CreateCalls++;
            LivePivotNamedSetSnapshot live = Live(definition, 0);
            Artifacts.RemoveAll(item => item.Name == definition.Name);
            Artifacts.Add(live);
            return live;
        }

        public LivePivotNamedSetSnapshot ReplaceSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before,
            DesiredPivotNamedSet definition) => throw new NotSupportedException();

        public LivePivotNamedSetSnapshot RestoreSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before)
        {
            Artifacts.RemoveAll(item => item.Name == before.Name);
            Artifacts.Add(before);
            return before;
        }

        public void DeleteSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot expected)
        {
            DeleteCalls++;
            if (expected.Orientation != 0)
            {
                throw new InvalidOperationException("visible");
            }

            Artifacts.RemoveAll(item => item.Name == expected.Name);
        }

        public void SetOrientation(int orientation, string? name = null)
        {
            for (int index = 0; index < Artifacts.Count; index++)
            {
                if (name == null || Artifacts[index].Name == name)
                {
                    Artifacts[index] = WithOrientation(Artifacts[index], orientation);
                }
            }
        }

        public static LivePivotNamedSetSnapshot Live(
            DesiredPivotNamedSet definition,
            int orientation)
        {
            string live = PivotNamedSetCanonical.CreateLiveFingerprint(
                definition.SourceFingerprint,
                Lineage,
                definition.Name,
                PivotNamedSetPairState.Complete,
                definition.FormulaFingerprint,
                definition.DisplayFolderMarker,
                definition.Name,
                definition.Caption,
                1,
                3,
                definition.Dynamic,
                definition.FlattenHierarchies,
                definition.FlattenHierarchies,
                definition.HierarchizeDistinct,
                definition.HierarchizeDistinct,
                true,
                orientation,
                true);
            return new LivePivotNamedSetSnapshot(
                "Sheet1",
                "PivotTable1",
                true,
                definition.Name,
                PivotNamedSetPairState.Complete,
                definition.RawMdx,
                definition.FormulaFingerprint,
                definition.DisplayFolderMarker,
                definition.Name,
                definition.Caption,
                1,
                3,
                definition.Dynamic,
                definition.FlattenHierarchies,
                definition.FlattenHierarchies,
                definition.HierarchizeDistinct,
                definition.HierarchizeDistinct,
                true,
                orientation,
                true,
                definition.SourceFingerprint,
                Lineage,
                live);
        }

        private static LivePivotNamedSetSnapshot WithOrientation(
            LivePivotNamedSetSnapshot snapshot,
            int orientation)
        {
            string live = PivotNamedSetCanonical.CreateLiveFingerprint(
                snapshot.SourceFingerprint,
                snapshot.ModelLineageFingerprint,
                snapshot.Name,
                snapshot.PairState,
                snapshot.FormulaFingerprint,
                snapshot.DisplayFolder,
                snapshot.SourceName,
                snapshot.Caption,
                snapshot.CalculatedMemberType,
                snapshot.CubeFieldType,
                snapshot.Dynamic,
                snapshot.CalculatedMemberFlattenHierarchies,
                snapshot.CubeFieldFlattenHierarchies,
                snapshot.CalculatedMemberHierarchizeDistinct,
                snapshot.CubeFieldHierarchizeDistinct,
                snapshot.ShowInFieldList,
                orientation,
                snapshot.IsValid);
            return new LivePivotNamedSetSnapshot(
                snapshot.WorksheetName,
                snapshot.PivotTableName,
                snapshot.IsSelectedTarget,
                snapshot.Name,
                snapshot.PairState,
                snapshot.RawFormula,
                snapshot.FormulaFingerprint,
                snapshot.DisplayFolder,
                snapshot.SourceName,
                snapshot.Caption,
                snapshot.CalculatedMemberType,
                snapshot.CubeFieldType,
                snapshot.Dynamic,
                snapshot.CalculatedMemberFlattenHierarchies,
                snapshot.CubeFieldFlattenHierarchies,
                snapshot.CalculatedMemberHierarchizeDistinct,
                snapshot.CubeFieldHierarchizeDistinct,
                snapshot.ShowInFieldList,
                orientation,
                snapshot.IsValid,
                snapshot.SourceFingerprint,
                snapshot.ModelLineageFingerprint,
                live);
        }
    }
}
