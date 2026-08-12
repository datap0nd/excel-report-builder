using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    internal sealed class PivotNamedSetHostBinding
    {
        public PivotNamedSetHostBinding(
            string definitionId,
            string hostSetName,
            string definitionFingerprint)
        {
            DefinitionId = definitionId;
            HostSetName = hostSetName;
            DefinitionFingerprint = definitionFingerprint;
        }

        public string DefinitionId { get; }
        public string HostSetName { get; }
        public string DefinitionFingerprint { get; }
    }

    internal sealed class PivotNamedSetPreparedMutation
    {
        private readonly Func<PivotNamedSetWorkbookSnapshot> verify;
        private readonly Action verifyRollback;
        private readonly Func<PivotNamedSetWorkbookSnapshot, IReadOnlyList<PivotPlusOwnedArtifact>>
            buildArtifacts;

        public PivotNamedSetPreparedMutation(
            BoundPivotNamedSetTarget target,
            PivotNamedSetWorkbookSnapshot before,
            string participantPlanFingerprint,
            IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions,
            IReadOnlyDictionary<string, PivotNamedSetHostBinding> definitionBindings,
            IReadOnlyList<PivotMutationStep> upsertSteps,
            IReadOnlyList<PivotMutationStep> deleteSteps,
            Func<PivotNamedSetWorkbookSnapshot> verify,
            Action verifyRollback,
            Func<PivotNamedSetWorkbookSnapshot, IReadOnlyList<PivotPlusOwnedArtifact>>
                buildArtifacts)
        {
            Target = target;
            Before = before;
            ParticipantPlanFingerprint = participantPlanFingerprint;
            Transitions = transitions;
            DefinitionBindings = definitionBindings;
            UpsertSteps = upsertSteps;
            DeleteSteps = deleteSteps;
            this.verify = verify;
            this.verifyRollback = verifyRollback;
            this.buildArtifacts = buildArtifacts;
        }

        public BoundPivotNamedSetTarget Target { get; }
        public PivotNamedSetWorkbookSnapshot Before { get; }
        public string ParticipantPlanFingerprint { get; }
        public IReadOnlyList<PivotPlusSemanticArtifactTransition> Transitions { get; }
        public IReadOnlyDictionary<string, PivotNamedSetHostBinding> DefinitionBindings { get; }
        public IReadOnlyList<PivotMutationStep> UpsertSteps { get; }
        public IReadOnlyList<PivotMutationStep> DeleteSteps { get; }
        public bool IsNoChange => Transitions.Count == 0;

        public PivotNamedSetWorkbookSnapshot Verify() => verify();
        public void VerifyRollback() => verifyRollback();

        public IReadOnlyList<PivotPlusOwnedArtifact> BuildArtifacts(
            PivotNamedSetWorkbookSnapshot verified) => buildArtifacts(verified);
    }

    /// <summary>
    /// Artifact-only named-set participant. Axis placement and the single
    /// PivotTable refresh are owned by the composite semantic coordinator.
    /// Existing definitions are immutable: a changed definition receives a
    /// new compiler-generated host name, is placed first, and the old set is
    /// removed only after the layout has stopped using it.
    /// </summary>
    internal sealed class PivotNamedSetMutationService
    {
        private const int Hidden = 0;
        private const int Row = 1;
        private const int Column = 2;

        private readonly IPivotNamedSetGateway gateway;

        public PivotNamedSetMutationService()
            : this(new LateBoundPivotNamedSetGateway())
        {
        }

        internal PivotNamedSetMutationService(IPivotNamedSetGateway gateway)
        {
            this.gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
        }

        internal string ComputeParticipantPlanFingerprint(
            string setupId,
            PivotMdxCompilation compilation,
            IEnumerable<PivotPlusSemanticArtifactTransition> transitions)
        {
            IReadOnlyList<DesiredPivotNamedSet> desired =
                PivotNamedSetCompilationAdapter.CreateDesired(setupId, compilation);
            return CreatePlanFingerprint(compilation, desired, transitions);
        }

        internal PivotNamedSetPreparedMutation PrepareParticipant(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId,
            PivotMdxCompilation compilation,
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            string? expectedParticipantPlanFingerprint = null)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
            if (baseMetadata == null) throw new ArgumentNullException(nameof(baseMetadata));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");

            IReadOnlyList<DesiredPivotNamedSet> desired =
                PivotNamedSetCompilationAdapter.CreateDesired(setupId, compilation);
            BoundPivotNamedSetTarget target = gateway.Bind(workbook, pivotTable, context);
            DemandTarget(baseMetadata, setupId, target.Identity);
            PivotNamedSetWorkbookSnapshot before = gateway.Capture(target);
            DemandSource(before, compilation.SourceFingerprint);

            IReadOnlyList<PivotPlusOwnedArtifact> owned = baseMetadata.Artifacts
                .Where(artifact => artifact.Kind == PivotPlusArtifactKind.NamedSet)
                .ToList();
            DemandOwnedRecords(owned, before, existingPending);

            var desiredByName = desired.ToDictionary(
                item => item.Name,
                StringComparer.OrdinalIgnoreCase);
            var ownedByName = owned.ToDictionary(
                item => item.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var selectedByName = before.SelectedPivot.Artifacts.ToDictionary(
                item => item.Name,
                StringComparer.OrdinalIgnoreCase);
            var transitions = new List<PivotPlusSemanticArtifactTransition>();

            foreach (DesiredPivotNamedSet definition in desired)
            {
                selectedByName.TryGetValue(
                    definition.Name,
                    out LivePivotNamedSetSnapshot? live);
                if (!ownedByName.TryGetValue(
                        definition.Name,
                        out PivotPlusOwnedArtifact? active))
                {
                    if (live != null && !MatchesDesiredDefinition(live, definition))
                    {
                        throw new InvalidOperationException(
                            "An unowned named set already uses the generated identity.");
                    }

                    transitions.Add(new PivotPlusSemanticArtifactTransition
                    {
                        Kind = PivotPlusArtifactKind.NamedSet,
                        ArtifactId = definition.Name,
                        Operation = PivotPlusSemanticArtifactOperation.Create,
                        BeforeLiveFingerprint = string.Empty,
                        PlannedDefinitionFingerprint = definition.DefinitionFingerprint
                    });
                    continue;
                }

                bool activeTruth = live != null && string.Equals(
                    live.LiveFingerprint,
                    active.Fingerprint,
                    StringComparison.Ordinal);
                bool exactPendingFinal = existingPending != null &&
                    live != null &&
                    MatchesDesiredDefinition(live, definition) &&
                    live.Orientation == Orientation(definition.Axis) &&
                    existingPending.Transitions.Any(item =>
                        item.Kind == PivotPlusArtifactKind.NamedSet &&
                        item.Operation == PivotPlusSemanticArtifactOperation.Update &&
                        string.Equals(
                            item.ArtifactId,
                            definition.Name,
                            StringComparison.Ordinal) &&
                        string.Equals(
                            item.BeforeLiveFingerprint,
                            active.Fingerprint,
                            StringComparison.Ordinal));
                if (!activeTruth && !exactPendingFinal)
                {
                    throw new InvalidOperationException(
                        "An owned named set no longer matches its exact ownership receipt.");
                }

                if (!MatchesDesiredDefinition(live!, definition))
                {
                    throw new InvalidOperationException(
                        "A changed named-set definition must use a new compiler-generated host name.");
                }

                int expectedOrientation = Orientation(definition.Axis);
                if (live!.Orientation != expectedOrientation || exactPendingFinal)
                {
                    transitions.Add(new PivotPlusSemanticArtifactTransition
                    {
                        Kind = PivotPlusArtifactKind.NamedSet,
                        ArtifactId = definition.Name,
                        Operation = PivotPlusSemanticArtifactOperation.Update,
                        BeforeLiveFingerprint = active.Fingerprint,
                        PlannedDefinitionFingerprint = definition.DefinitionFingerprint
                    });
                }
            }

            foreach (PivotPlusOwnedArtifact active in owned.Where(
                         artifact => !desiredByName.ContainsKey(artifact.ArtifactId)))
            {
                transitions.Add(new PivotPlusSemanticArtifactTransition
                {
                    Kind = PivotPlusArtifactKind.NamedSet,
                    ArtifactId = active.ArtifactId,
                    Operation = PivotPlusSemanticArtifactOperation.Delete,
                    BeforeLiveFingerprint = active.Fingerprint,
                    PlannedDefinitionFingerprint = active.Fingerprint
                });
            }

            transitions = OrderTransitions(transitions);
            DemandPendingSlice(existingPending, transitions);
            string planFingerprint = CreatePlanFingerprint(
                compilation,
                desired,
                transitions);
            if (!string.IsNullOrEmpty(expectedParticipantPlanFingerprint) &&
                !string.Equals(
                    planFingerprint,
                    expectedParticipantPlanFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The named-set participant does not match the durable combined plan.");
            }

            var upsertSteps = new List<PivotMutationStep>();
            foreach (DesiredPivotNamedSet definition in desired.Where(definition =>
                         !ownedByName.ContainsKey(definition.Name)))
            {
                DesiredPivotNamedSet captured = definition;
                LivePivotNamedSetSnapshot? created = null;
                upsertSteps.Add(new PivotMutationStep(
                    "create named set " + captured.Name,
                    () =>
                    {
                        PivotNamedSetWorkbookSnapshot current = gateway.Capture(target);
                        LivePivotNamedSetSnapshot? existing = FindSelected(
                            current,
                            captured.Name);
                        if (existing != null &&
                            MatchesDesiredDefinition(existing, captured))
                        {
                            created = existing;
                            return;
                        }

                        created = gateway.CreateSet(target, captured);
                    },
                    () =>
                    {
                        PivotNamedSetWorkbookSnapshot current = gateway.Capture(target);
                        LivePivotNamedSetSnapshot? existing = FindSelected(
                            current,
                            captured.Name);
                        if (existing == null) return;
                        gateway.DeleteSet(target, DemandHidden(existing));
                    }));
            }

            var deleteSteps = new List<PivotMutationStep>();
            foreach (PivotPlusOwnedArtifact active in owned.Where(
                         artifact => !desiredByName.ContainsKey(artifact.ArtifactId)))
            {
                if (!selectedByName.TryGetValue(
                        active.ArtifactId,
                        out LivePivotNamedSetSnapshot? prior))
                {
                    // An exact pending Delete may already have committed. The
                    // durable journal preserves the before receipt; this retry
                    // only needs to verify the final absence and commit.
                    continue;
                }

                LivePivotNamedSetSnapshot hiddenPrior = WithOrientation(prior, Hidden);
                deleteSteps.Add(new PivotMutationStep(
                    "delete named set " + prior.Name,
                    () =>
                    {
                        PivotNamedSetWorkbookSnapshot current = gateway.Capture(target);
                        LivePivotNamedSetSnapshot? live = FindSelected(current, prior.Name);
                        if (live == null) return;
                        gateway.DeleteSet(target, DemandHidden(live));
                    },
                    () =>
                    {
                        PivotNamedSetWorkbookSnapshot current = gateway.Capture(target);
                        if (FindSelected(current, prior.Name) == null)
                        {
                            gateway.RestoreSet(target, hiddenPrior);
                        }
                    }));
            }

            var bindings = new ReadOnlyDictionary<string, PivotNamedSetHostBinding>(
                desired.ToDictionary(
                    item => item.DefinitionId,
                    item => new PivotNamedSetHostBinding(
                        item.DefinitionId,
                        item.Name,
                        item.DefinitionFingerprint),
                    StringComparer.OrdinalIgnoreCase));
            var readonlyTransitions = new ReadOnlyCollection<PivotPlusSemanticArtifactTransition>(
                transitions.Select(Clone).ToList());

            return new PivotNamedSetPreparedMutation(
                target,
                before,
                planFingerprint,
                readonlyTransitions,
                bindings,
                new ReadOnlyCollection<PivotMutationStep>(upsertSteps),
                new ReadOnlyCollection<PivotMutationStep>(deleteSteps),
                () =>
                {
                    PivotNamedSetWorkbookSnapshot after = gateway.Capture(target);
                    VerifyFinal(before, after, desired, owned);
                    return after;
                },
                () => DemandExactSnapshot(before, gateway.Capture(target)),
                verified => BuildArtifacts(verified, desired));
        }

        private static IReadOnlyList<PivotPlusOwnedArtifact> BuildArtifacts(
            PivotNamedSetWorkbookSnapshot verified,
            IReadOnlyList<DesiredPivotNamedSet> desired)
        {
            var result = new List<PivotPlusOwnedArtifact>();
            foreach (DesiredPivotNamedSet definition in desired)
            {
                LivePivotNamedSetSnapshot live = verified.SelectedPivot.Artifacts.Single(
                    artifact => string.Equals(
                        artifact.Name,
                        definition.Name,
                        StringComparison.OrdinalIgnoreCase));
                result.Add(new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.NamedSet,
                    ArtifactId = live.Name,
                    Fingerprint = live.LiveFingerprint
                });
            }

            return new ReadOnlyCollection<PivotPlusOwnedArtifact>(result);
        }

        private static void VerifyFinal(
            PivotNamedSetWorkbookSnapshot before,
            PivotNamedSetWorkbookSnapshot after,
            IReadOnlyList<DesiredPivotNamedSet> desired,
            IReadOnlyList<PivotPlusOwnedArtifact> priorOwned)
        {
            DemandSource(after, before.SourceFingerprint);
            var desiredByName = desired.ToDictionary(
                item => item.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (DesiredPivotNamedSet definition in desired)
            {
                LivePivotNamedSetSnapshot live = after.SelectedPivot.Artifacts.SingleOrDefault(
                    artifact => string.Equals(
                        artifact.Name,
                        definition.Name,
                        StringComparison.OrdinalIgnoreCase)) ??
                    throw new InvalidOperationException(
                        "Excel did not retain a generated named set.");
                if (!MatchesDesiredDefinition(live, definition) ||
                    live.Orientation != Orientation(definition.Axis))
                {
                    throw new InvalidOperationException(
                        "Excel did not retain the exact generated named-set definition and axis.");
                }
            }

            var removed = new HashSet<string>(
                priorOwned.Select(item => item.ArtifactId)
                    .Where(name => !desiredByName.ContainsKey(name)),
                StringComparer.OrdinalIgnoreCase);
            if (after.SelectedPivot.Artifacts.Any(artifact => removed.Contains(artifact.Name)))
            {
                throw new InvalidOperationException(
                    "Excel retained a named set scheduled for deletion.");
            }

            var changed = new HashSet<string>(
                priorOwned.Select(item => item.ArtifactId)
                    .Concat(desired.Select(item => item.Name)),
                StringComparer.OrdinalIgnoreCase);
            DemandUnrelatedInventory(before, after, changed);
        }

        private static void DemandUnrelatedInventory(
            PivotNamedSetWorkbookSnapshot before,
            PivotNamedSetWorkbookSnapshot after,
            ISet<string> changed)
        {
            if (before.Pivots.Count != after.Pivots.Count)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model PivotTable inventory changed during named-set Apply.");
            }

            foreach (PivotNamedSetPivotSnapshot priorPivot in before.Pivots)
            {
                PivotNamedSetPivotSnapshot currentPivot = after.Pivots.SingleOrDefault(pivot =>
                    string.Equals(pivot.WorksheetName, priorPivot.WorksheetName, StringComparison.Ordinal) &&
                    string.Equals(pivot.PivotTableName, priorPivot.PivotTableName, StringComparison.Ordinal)) ??
                    throw new InvalidOperationException(
                        "A Data Model PivotTable changed identity during named-set Apply.");
                IEnumerable<LivePivotNamedSetSnapshot> priorArtifacts = priorPivot.Artifacts
                    .Where(item => priorPivot.IsSelectedTarget && changed.Contains(item.Name)
                        ? false
                        : true);
                IEnumerable<LivePivotNamedSetSnapshot> currentArtifacts = currentPivot.Artifacts
                    .Where(item => currentPivot.IsSelectedTarget && changed.Contains(item.Name)
                        ? false
                        : true);
                if (!SnapshotsEqual(priorArtifacts, currentArtifacts))
                {
                    throw new InvalidOperationException(
                        "An unrelated named set changed during the semantic Apply.");
                }

                IEnumerable<PivotCalculatedMemberReferenceSnapshot> priorCalculated =
                    priorPivot.CalculatedMembers.Where(item =>
                        priorPivot.IsSelectedTarget && changed.Contains(item.Name)
                            ? false
                            : true);
                IEnumerable<PivotCalculatedMemberReferenceSnapshot> currentCalculated =
                    currentPivot.CalculatedMembers.Where(item =>
                        currentPivot.IsSelectedTarget && changed.Contains(item.Name)
                            ? false
                            : true);
                if (!CalculatedEqual(priorCalculated, currentCalculated))
                {
                    throw new InvalidOperationException(
                        "An unrelated calculated member changed during the semantic Apply.");
                }
            }
        }

        private static bool SnapshotsEqual(
            IEnumerable<LivePivotNamedSetSnapshot> left,
            IEnumerable<LivePivotNamedSetSnapshot> right)
        {
            string[] first = left.Select(item => item.Name + "\u001f" + item.LiveFingerprint)
                .OrderBy(item => item, StringComparer.Ordinal).ToArray();
            string[] second = right.Select(item => item.Name + "\u001f" + item.LiveFingerprint)
                .OrderBy(item => item, StringComparer.Ordinal).ToArray();
            return first.SequenceEqual(second, StringComparer.Ordinal);
        }

        private static bool CalculatedEqual(
            IEnumerable<PivotCalculatedMemberReferenceSnapshot> left,
            IEnumerable<PivotCalculatedMemberReferenceSnapshot> right)
        {
            string Key(PivotCalculatedMemberReferenceSnapshot item) =>
                item.Name + "\u001f" + item.Type + "\u001f" + item.RawFormula + "\u001f" +
                item.FormulaScanComplete;
            return left.Select(Key).OrderBy(item => item, StringComparer.Ordinal)
                .SequenceEqual(
                    right.Select(Key).OrderBy(item => item, StringComparer.Ordinal),
                    StringComparer.Ordinal);
        }

        private static void DemandExactSnapshot(
            PivotNamedSetWorkbookSnapshot expected,
            PivotNamedSetWorkbookSnapshot actual)
        {
            DemandSource(actual, expected.SourceFingerprint);
            DemandUnrelatedInventory(
                expected,
                actual,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }

        private static void DemandOwnedRecords(
            IReadOnlyList<PivotPlusOwnedArtifact> owned,
            PivotNamedSetWorkbookSnapshot snapshot,
            PivotPlusPendingSemanticApplyMetadata? pending)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact artifact in owned)
            {
                if (!names.Add(artifact.ArtifactId))
                {
                    throw new InvalidOperationException(
                        "Named-set ownership contains a duplicate artifact identity.");
                }

                LivePivotNamedSetSnapshot? live = snapshot.SelectedPivot.Artifacts.SingleOrDefault(
                    item => string.Equals(
                        item.Name,
                        artifact.ArtifactId,
                        StringComparison.OrdinalIgnoreCase));
                if (live != null && string.Equals(
                    live.LiveFingerprint,
                    artifact.Fingerprint,
                    StringComparison.Ordinal))
                {
                    continue;
                }

                bool recognizedPending = pending != null && pending.Transitions.Any(item =>
                    item.Kind == PivotPlusArtifactKind.NamedSet &&
                    item.Operation != PivotPlusSemanticArtifactOperation.Create &&
                    string.Equals(
                        item.ArtifactId,
                        artifact.ArtifactId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        item.BeforeLiveFingerprint,
                        artifact.Fingerprint,
                        StringComparison.Ordinal));
                if (!recognizedPending)
                {
                    throw new InvalidOperationException(
                        "An owned named set changed outside PivotTable+.");
                }
            }
        }

        private static void DemandTarget(
            PivotPlusWorkbookMetadata metadata,
            string setupId,
            PivotTargetIdentity target)
        {
            if (!string.Equals(metadata.SetupId, setupId, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    metadata.TargetWorksheetName,
                    target.WorksheetName,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    metadata.TargetPivotTableName,
                    target.PivotTableName,
                    StringComparison.Ordinal) ||
                metadata.RecoveryPhase != PivotPlusRecoveryPhase.None)
            {
                throw new InvalidOperationException(
                    "The named-set participant is not bound to the exact active setup target.");
            }
        }

        private static void DemandPendingSlice(
            PivotPlusPendingSemanticApplyMetadata? pending,
            IReadOnlyList<PivotPlusSemanticArtifactTransition> proposed)
        {
            if (pending == null) return;
            IReadOnlyList<PivotPlusSemanticArtifactTransition> existing = OrderTransitions(
                pending.Transitions.Where(item => item.Kind == PivotPlusArtifactKind.NamedSet));
            if (existing.Count != proposed.Count)
            {
                throw new InvalidOperationException(
                    "The named sets do not match the pending combined Apply.");
            }

            for (int index = 0; index < existing.Count; index++)
            {
                PivotPlusSemanticArtifactTransition left = existing[index];
                PivotPlusSemanticArtifactTransition right = proposed[index];
                if (left.Operation != right.Operation ||
                    !string.Equals(left.ArtifactId, right.ArtifactId, StringComparison.Ordinal) ||
                    !string.Equals(left.BeforeLiveFingerprint, right.BeforeLiveFingerprint, StringComparison.Ordinal) ||
                    !string.Equals(left.PlannedDefinitionFingerprint, right.PlannedDefinitionFingerprint, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The named sets do not match the exact pending transition slice.");
                }
            }
        }

        private static string CreatePlanFingerprint(
            PivotMdxCompilation compilation,
            IEnumerable<DesiredPivotNamedSet> desired,
            IEnumerable<PivotPlusSemanticArtifactTransition> transitions)
        {
            var canonical = new StringBuilder("namedset-participant-plan-v1");
            Append(canonical, compilation.CompilationFingerprint);
            Append(canonical, compilation.SourceFingerprint);
            foreach (DesiredPivotNamedSet item in desired.OrderBy(
                         value => value.DefinitionId,
                         StringComparer.Ordinal))
            {
                Append(canonical, item.DefinitionId);
                Append(canonical, item.Name);
                Append(canonical, item.DefinitionFingerprint);
                Append(canonical, item.FormulaFingerprint);
                Append(canonical, (int)item.Axis);
            }

            foreach (PivotPlusSemanticArtifactTransition transition in
                     OrderTransitions(transitions))
            {
                Append(canonical, (int)transition.Operation);
                Append(canonical, transition.ArtifactId);
                Append(canonical, transition.BeforeLiveFingerprint);
                Append(canonical, transition.PlannedDefinitionFingerprint);
            }

            return PivotPlusFingerprint.Create(
                "namedset.participant-plan.v1",
                canonical.ToString());
        }

        private static bool MatchesDesiredDefinition(
            LivePivotNamedSetSnapshot live,
            DesiredPivotNamedSet desired)
        {
            return live.IsSelectedTarget &&
                   live.PairState == PivotNamedSetPairState.Complete &&
                   string.Equals(live.SourceFingerprint, desired.SourceFingerprint, StringComparison.Ordinal) &&
                   string.Equals(live.Name, desired.Name, StringComparison.Ordinal) &&
                   string.Equals(live.SourceName, desired.Name, StringComparison.Ordinal) &&
                   string.Equals(live.Caption, desired.Caption, StringComparison.Ordinal) &&
                   string.Equals(live.RawFormula, desired.RawMdx, StringComparison.Ordinal) &&
                   string.Equals(live.FormulaFingerprint, desired.FormulaFingerprint, StringComparison.Ordinal) &&
                   string.Equals(live.DisplayFolder, desired.DisplayFolderMarker, StringComparison.Ordinal) &&
                   live.CalculatedMemberType == 1 &&
                   live.CubeFieldType == 3 &&
                   live.Dynamic == desired.Dynamic &&
                   live.CalculatedMemberFlattenHierarchies == desired.FlattenHierarchies &&
                   live.CubeFieldFlattenHierarchies == desired.FlattenHierarchies &&
                   live.CalculatedMemberHierarchizeDistinct == desired.HierarchizeDistinct &&
                   live.CubeFieldHierarchizeDistinct == desired.HierarchizeDistinct &&
                   live.ShowInFieldList == true &&
                   live.IsValid == true;
        }

        private static LivePivotNamedSetSnapshot DemandHidden(
            LivePivotNamedSetSnapshot snapshot)
        {
            if (snapshot.Orientation != Hidden)
            {
                throw new InvalidOperationException(
                    "A named set must be hidden by the layout phase before deletion.");
            }

            return snapshot;
        }

        private static LivePivotNamedSetSnapshot WithOrientation(
            LivePivotNamedSetSnapshot snapshot,
            int orientation)
        {
            string fingerprint = PivotNamedSetCanonical.CreateLiveFingerprint(
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
                fingerprint);
        }

        private static LivePivotNamedSetSnapshot? FindSelected(
            PivotNamedSetWorkbookSnapshot snapshot,
            string name)
        {
            return snapshot.SelectedPivot.Artifacts.SingleOrDefault(item => string.Equals(
                item.Name,
                name,
                StringComparison.OrdinalIgnoreCase));
        }

        private static int Orientation(PivotNamedSetAxis axis) =>
            axis == PivotNamedSetAxis.Row ? Row : Column;

        private static void DemandSource(
            PivotNamedSetWorkbookSnapshot snapshot,
            string expected)
        {
            if (!string.Equals(
                    snapshot.SourceFingerprint,
                    expected,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The live Data Model schema changed after named-set compilation.");
            }
        }

        private static List<PivotPlusSemanticArtifactTransition> OrderTransitions(
            IEnumerable<PivotPlusSemanticArtifactTransition> transitions)
        {
            return transitions.OrderBy(item => item.Kind)
                .ThenBy(item => item.ArtifactId, StringComparer.Ordinal)
                .ThenBy(item => item.Operation)
                .Select(Clone)
                .ToList();
        }

        private static PivotPlusSemanticArtifactTransition Clone(
            PivotPlusSemanticArtifactTransition transition)
        {
            return new PivotPlusSemanticArtifactTransition
            {
                Kind = transition.Kind,
                ArtifactId = transition.ArtifactId,
                Operation = transition.Operation,
                BeforeLiveFingerprint = transition.BeforeLiveFingerprint,
                PlannedDefinitionFingerprint = transition.PlannedDefinitionFingerprint
            };
        }

        private static void Append(StringBuilder builder, string value)
        {
            builder.Append('|').Append(value.Length).Append(':').Append(value);
        }

        private static void Append(StringBuilder builder, int value)
        {
            Append(builder, value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
