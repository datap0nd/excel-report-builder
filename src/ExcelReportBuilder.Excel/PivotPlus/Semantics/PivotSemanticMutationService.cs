using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Measures;
using ExcelReportBuilder.Excel.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Semantics
{
    public enum PivotSemanticApplyStatus
    {
        Applied,
        Recovered
    }

    public sealed class PivotSemanticApplyResult
    {
        public PivotSemanticApplyResult(
            string applyId,
            PivotSemanticApplyStatus status,
            int measureTransitionCount,
            int namedSetTransitionCount,
            bool undoAvailable)
        {
            ApplyId = applyId;
            Status = status;
            MeasureTransitionCount = measureTransitionCount;
            NamedSetTransitionCount = namedSetTransitionCount;
            UndoAvailable = undoAvailable;
        }

        public string ApplyId { get; }
        public PivotSemanticApplyStatus Status { get; }
        public int MeasureTransitionCount { get; }
        public int NamedSetTransitionCount { get; }
        public bool UndoAvailable { get; }
    }

    public sealed class PivotSemanticUndoResult
    {
        public PivotSemanticUndoResult(string applyId, bool recovered)
        {
            ApplyId = applyId;
            Recovered = recovered;
        }

        public string ApplyId { get; }
        public bool Recovered { get; }
    }

    public sealed class PivotSemanticMutationException : InvalidOperationException
    {
        internal PivotSemanticMutationException(
            string message,
            bool rollbackCompleted,
            bool recoveryRequired,
            Exception innerException)
            : base(message, innerException)
        {
            RollbackCompleted = rollbackCompleted;
            RecoveryRequired = recoveryRequired;
        }

        public bool RollbackCompleted { get; }
        public bool RecoveryRequired { get; }
    }

    /// <summary>
    /// Executes one typed Measure + NamedSet + layout transaction against one
    /// real Data Model PivotTable. It journals all semantic artifacts once,
    /// refreshes the selected PivotTable exactly once, verifies every layer,
    /// and commits ownership once. It never accepts raw DAX or MDX.
    /// </summary>
    public sealed class PivotSemanticMutationService
    {
        private readonly PivotModelMeasureMutationService measures;
        private readonly PivotNamedSetMutationService namedSets;
        private readonly LateBoundPivotSemanticLayoutGateway layoutGateway;
        private readonly PivotModelMeasureOwnershipStore ownership;
        private readonly IWorkbookIdentityResolver workbookIdentity;
        private readonly PivotMutationCoordinator coordinator;
        private readonly object synchronization = new object();
        private bool active;
        private static readonly ConditionalWeakTable<object, LayoutSeedLedger> LayoutSeeds =
            new ConditionalWeakTable<object, LayoutSeedLedger>();
        private static readonly ConditionalWeakTable<object, SemanticUndoLedger> UndoStates =
            new ConditionalWeakTable<object, SemanticUndoLedger>();

        public PivotSemanticMutationService()
            : this(
                new PivotModelMeasureMutationService(),
                new PivotNamedSetMutationService(),
                new LateBoundPivotSemanticLayoutGateway(),
                new PivotModelMeasureOwnershipStore(),
                new StoredWorkbookIdentityResolver(),
                new PivotMutationCoordinator())
        {
        }

        internal PivotSemanticMutationService(
            PivotModelMeasureMutationService measures,
            PivotNamedSetMutationService namedSets,
            LateBoundPivotSemanticLayoutGateway layoutGateway,
            PivotModelMeasureOwnershipStore ownership,
            IWorkbookIdentityResolver workbookIdentity,
            PivotMutationCoordinator coordinator)
        {
            this.measures = measures ?? throw new ArgumentNullException(nameof(measures));
            this.namedSets = namedSets ?? throw new ArgumentNullException(nameof(namedSets));
            this.layoutGateway = layoutGateway ??
                throw new ArgumentNullException(nameof(layoutGateway));
            this.ownership = ownership ?? throw new ArgumentNullException(nameof(ownership));
            this.workbookIdentity = workbookIdentity ??
                throw new ArgumentNullException(nameof(workbookIdentity));
            this.coordinator = coordinator ?? throw new ArgumentNullException(nameof(coordinator));
        }

        public PivotSemanticApplyResult Apply(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId,
            PivotDaxCompilation daxCompilation,
            PivotMdxCompilation mdxCompilation,
            PivotSemanticLayoutPlan layout)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (daxCompilation == null)
            {
                throw new ArgumentNullException(nameof(daxCompilation));
            }

            if (mdxCompilation == null)
            {
                throw new ArgumentNullException(nameof(mdxCompilation));
            }

            if (layout == null) throw new ArgumentNullException(nameof(layout));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            if (!mdxCompilation.HasExactMeasureDependencies(daxCompilation))
            {
                throw new InvalidOperationException(
                    "The named-set compilation is not bound to the exact supplied measure compilation.");
            }

            Enter();
            try
            {
                BoundPivotSemanticLayoutTarget layoutTarget = layoutGateway.Bind(
                    workbook,
                    pivotTable,
                    context);
                PivotSemanticLayoutSnapshot layoutBefore = layoutGateway.Capture(layoutTarget);
                PivotPlusWorkbookMetadata baseMetadata = ownership.ReadBase(
                    workbook,
                    setupId,
                    layoutTarget.Identity,
                    out PivotPlusPendingSemanticApplyMetadata? existingPending);
                PivotSemanticLayoutSnapshot layoutPreview = layoutBefore;
                if (existingPending != null &&
                    TryGetLayoutSeed(
                        workbook,
                        setupId,
                        existingPending.ApplyId,
                        existingPending.PlanFingerprint,
                        out PivotSemanticLayoutSnapshot? savedPreview) &&
                    savedPreview != null)
                {
                    layoutPreview = savedPreview;
                }

                PivotModelMeasureArtifactPreparedMutation measurePrepared;
                PivotNamedSetPreparedMutation namedSetPrepared;
                string layoutPlanFingerprint = CreateLayoutPlanFingerprint(layout);
                if (existingPending == null)
                {
                    measurePrepared = measures.PrepareArtifactParticipant(
                        workbook,
                        pivotTable,
                        context,
                        setupId,
                        daxCompilation,
                        baseMetadata,
                        existingPending: null);
                    namedSetPrepared = namedSets.PrepareParticipant(
                        workbook,
                        pivotTable,
                        context,
                        setupId,
                        mdxCompilation,
                        baseMetadata,
                        existingPending: null);
                }
                else
                {
                    string measurePlanFingerprint =
                        measures.ComputeArtifactParticipantPlanFingerprint(
                            setupId,
                            daxCompilation,
                            existingPending);
                    string namedSetPlanFingerprint =
                        namedSets.ComputeParticipantPlanFingerprint(
                            setupId,
                            mdxCompilation,
                            existingPending.Transitions.Where(item =>
                                item.Kind == PivotPlusArtifactKind.NamedSet));
                    string expectedCombined = CreateCombinedPlanFingerprint(
                        measurePlanFingerprint,
                        namedSetPlanFingerprint,
                        layoutPlanFingerprint,
                        daxCompilation,
                        mdxCompilation);
                    if (!string.Equals(
                            existingPending.PlanFingerprint,
                            expectedCombined,
                            StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "The requested semantic Apply does not match the durable pending plan.");
                    }

                    measurePrepared = measures.PrepareArtifactParticipant(
                        workbook,
                        pivotTable,
                        context,
                        setupId,
                        daxCompilation,
                        baseMetadata,
                        existingPending,
                        new PivotModelMeasureParticipantRetryBinding(
                            expectedCombined,
                            measurePlanFingerprint));
                    namedSetPrepared = namedSets.PrepareParticipant(
                        workbook,
                        pivotTable,
                        context,
                        setupId,
                        mdxCompilation,
                        baseMetadata,
                        existingPending,
                        namedSetPlanFingerprint);
                }

                DemandSameTarget(
                    layoutTarget.Identity,
                    measurePrepared.Target.Identity,
                    namedSetPrepared.Target.Identity);
                DemandNoUnsafeMdxMeasureReferences(
                    namedSetPrepared.Before,
                    measurePrepared.Transitions,
                    namedSetPrepared.DefinitionBindings.Values.Select(item => item.HostSetName));

                IReadOnlyDictionary<string, string> measureMappings =
                    new ReadOnlyDictionary<string, string>(
                        measurePrepared.DefinitionBindings.ToDictionary(
                            item => item.Key,
                            item => item.Value.HostMeasureName,
                            StringComparer.Ordinal));
                IReadOnlyDictionary<string, string> namedSetMappings =
                    new ReadOnlyDictionary<string, string>(
                        namedSetPrepared.DefinitionBindings.ToDictionary(
                            item => item.Key,
                            item => item.Value.HostSetName,
                            StringComparer.Ordinal));
                PivotSemanticPreparedPlacement layoutPrepared;
                if (existingPending != null && !string.Equals(
                        layoutBefore.LayoutFingerprint,
                        layoutPreview.LayoutFingerprint,
                        StringComparison.Ordinal))
                {
                    if (ReferenceEquals(layoutPreview, layoutBefore))
                    {
                        throw new PivotSemanticMutationException(
                            "The pending semantic layout needs its original session preview to recover safely.",
                            rollbackCompleted: false,
                            recoveryRequired: true,
                            new InvalidOperationException(
                                "The filter-preserving layout snapshot is unavailable."));
                    }

                    layoutPrepared = layoutGateway.PrepareRecoveredFinal(
                        layoutTarget,
                        layout,
                        namedSetMappings,
                        measureMappings,
                        layoutPreview);
                }
                else
                {
                    layoutPrepared = layoutGateway.Prepare(
                        layoutTarget,
                        layout,
                        namedSetMappings,
                        measureMappings,
                        layoutPreview);
                }

                string combinedPlanFingerprint = CreateCombinedPlanFingerprint(
                    measurePrepared.ParticipantPlanFingerprint,
                    namedSetPrepared.ParticipantPlanFingerprint,
                    layoutPlanFingerprint,
                    daxCompilation,
                    mdxCompilation);
                var transitions = measurePrepared.Transitions
                    .Concat(namedSetPrepared.Transitions)
                    .OrderBy(item => item.Kind)
                    .ThenBy(item => item.ArtifactId, StringComparer.Ordinal)
                    .Select(Clone)
                    .ToList();
                string applyId = existingPending?.ApplyId ??
                    "apply_" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
                var pending = new PivotPlusPendingSemanticApplyMetadata
                {
                    ApplyId = applyId,
                    PlanFingerprint = combinedPlanFingerprint,
                    BeforePivotFingerprint = layoutPreview.LayoutFingerprint,
                    ExpectedPivotFingerprint = PivotPlusFingerprint.Create(
                        "semantic.expected-layout.v1",
                        layoutPlanFingerprint),
                    Transitions = transitions
                };
                if (existingPending != null)
                {
                    DemandSamePending(existingPending, pending);
                }
                else
                {
                    RememberLayoutSeed(
                        workbook,
                        setupId,
                        applyId,
                        combinedPlanFingerprint,
                        layoutPreview);
                }

                measurePrepared.PrimeUndoContribution(applyId, combinedPlanFingerprint);
                workbookIdentity.Persist(workbook, layoutTarget.Identity.WorkbookId);
                PivotModelMeasureOwnershipSession session;
                try
                {
                    session = ownership.Journal(
                        workbook,
                        setupId,
                        layoutTarget.Identity,
                        pending);
                }
                catch (Exception failure)
                {
                    throw new PivotSemanticMutationException(
                        "The combined semantic journal could not be confirmed. Retry the identical Apply.",
                        rollbackCompleted: true,
                        recoveryRequired: true,
                        failure);
                }

                var steps = new List<PivotMutationStep>();
                steps.AddRange(measurePrepared.UpsertSteps);
                steps.AddRange(namedSetPrepared.UpsertSteps);
                steps.Add(new PivotMutationStep(
                    "apply complete PivotTable semantic layout",
                    layoutPrepared.Apply,
                    layoutPrepared.Restore));
                steps.AddRange(namedSetPrepared.DeleteSteps);
                steps.AddRange(measurePrepared.DeleteSteps);

                ModelMeasureWorkbookSnapshot? verifiedMeasures = null;
                PivotNamedSetWorkbookSnapshot? verifiedSets = null;
                PivotSemanticLayoutSnapshot? verifiedLayout = null;
                try
                {
                    ExecutePrepared(
                        coordinator,
                        pivotTable,
                        steps,
                        measurePrepared.Refresh,
                        () =>
                        {
                            layoutPrepared.VerifyDesired();
                            verifiedMeasures = measurePrepared.Verify();
                            verifiedSets = namedSetPrepared.Verify();
                            verifiedLayout = layoutGateway.Capture(layoutTarget);
                        });
                }
                catch (PivotMutationException failure)
                {
                    if (failure.RollbackCompleted)
                    {
                        try
                        {
                            measurePrepared.VerifyRollback();
                            namedSetPrepared.VerifyRollback();
                            ownership.RestoreBase(workbook, session);
                        }
                        catch (Exception recoveryFailure)
                        {
                            throw new PivotSemanticMutationException(
                                "The combined semantic Apply failed; exact rollback or journal cleanup could not be proven.",
                                rollbackCompleted: false,
                                recoveryRequired: true,
                                new AggregateException(failure, recoveryFailure));
                        }
                    }

                    throw new PivotSemanticMutationException(
                        failure.Message,
                        failure.RollbackCompleted,
                        recoveryRequired: !failure.RollbackCompleted,
                        failure);
                }

                if (verifiedMeasures == null || verifiedSets == null || verifiedLayout == null)
                {
                    throw new InvalidOperationException(
                        "The combined semantic verification did not produce exact host receipts.");
                }

                IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts =
                    measurePrepared.BuildArtifacts(verifiedMeasures)
                        .Concat(namedSetPrepared.BuildArtifacts(verifiedSets))
                        .ToList();
                PivotPlusUndoMetadata undo = BuildUndo(
                    applyId,
                    layoutPreview,
                    verifiedLayout,
                    transitions,
                    finalArtifacts);
                PivotModelMeasureArtifactUndoContribution? measureUndo =
                    measurePrepared.BuildUndoContribution(
                        applyId,
                        combinedPlanFingerprint,
                        verifiedMeasures);
                try
                {
                    ownership.CommitSemantic(
                        workbook,
                        session,
                        finalArtifacts,
                        undo);
                }
                catch (Exception failure)
                {
                    throw new PivotSemanticMutationException(
                        "Excel applied and verified the semantic plan, but ownership finalization was ambiguous. Retry the identical Apply.",
                        rollbackCompleted: false,
                        recoveryRequired: true,
                        failure);
                }

                RememberUndoState(
                    workbook,
                    setupId,
                    new SemanticUndoState(
                        applyId,
                        combinedPlanFingerprint,
                        layoutTarget.Identity,
                        session.BaseMetadata,
                        finalArtifacts,
                        transitions,
                        measureUndo,
                        namedSetPrepared,
                        layoutPrepared,
                        layoutTarget,
                        layoutPreview,
                        verifiedLayout,
                        measurePrepared.Refresh));

                return new PivotSemanticApplyResult(
                    applyId,
                    existingPending == null
                        ? PivotSemanticApplyStatus.Applied
                        : PivotSemanticApplyStatus.Recovered,
                    measurePrepared.Transitions.Count,
                    namedSetPrepared.Transitions.Count,
                    undoAvailable: true);
            }
            finally
            {
                Exit();
            }
        }

        public PivotSemanticUndoResult Undo(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            if (!TryGetUndoState(workbook, setupId, out SemanticUndoState? state) ||
                state == null)
            {
                throw new InvalidOperationException(
                    "Undo is unavailable after the Excel session that applied this semantic plan has ended.");
            }

            Enter();
            try
            {
                BoundPivotSemanticLayoutTarget liveLayoutTarget = layoutGateway.Bind(
                    workbook,
                    pivotTable,
                    context);
                DemandSameTarget(state.Target, liveLayoutTarget.Identity);
                PivotPlusWorkbookMetadata currentBase = ownership.ReadBase(
                    workbook,
                    setupId,
                    liveLayoutTarget.Identity,
                    out PivotPlusPendingSemanticApplyMetadata? existingPending);
                PivotModelMeasureArtifactPreparedUndo? measureUndo = state.MeasureUndo == null
                    ? null
                    : measures.PrepareArtifactUndoParticipant(
                        workbook,
                        pivotTable,
                        context,
                        state.MeasureUndo);

                PivotPlusPendingSemanticApplyMetadata pending = BuildUndoPending(
                    state,
                    currentBase,
                    existingPending);
                PivotModelMeasureOwnershipSession session;
                try
                {
                    session = ownership.Journal(
                        workbook,
                        setupId,
                        liveLayoutTarget.Identity,
                        pending);
                }
                catch (Exception failure)
                {
                    throw new PivotSemanticMutationException(
                        "The combined Undo journal could not be confirmed. Retry the same Undo.",
                        rollbackCompleted: true,
                        recoveryRequired: true,
                        failure);
                }

                PivotSemanticLayoutSnapshot currentLayout = layoutGateway.Capture(
                    liveLayoutTarget);
                if (string.Equals(
                        currentLayout.LayoutFingerprint,
                        state.BeforeLayout.LayoutFingerprint,
                        StringComparison.Ordinal))
                {
                    try
                    {
                        layoutGateway.VerifySnapshot(liveLayoutTarget, state.BeforeLayout);
                        state.NamedSets.VerifyRollback();
                        measureUndo?.Verify();
                        ownership.CommitSemanticState(workbook, session, state.BaseMetadata);
                        RemoveUndoState(workbook, setupId);
                        return new PivotSemanticUndoResult(state.ApplyId, recovered: true);
                    }
                    catch (Exception failure)
                    {
                        throw new PivotSemanticMutationException(
                            "The already-applied Undo could not be finalized safely.",
                            rollbackCompleted: false,
                            recoveryRequired: true,
                            failure);
                    }
                }

                if (!string.Equals(
                        currentLayout.LayoutFingerprint,
                        state.AfterLayout.LayoutFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new PivotSemanticMutationException(
                        "The PivotTable layout changed after Apply; session Undo is no longer safe.",
                        rollbackCompleted: false,
                        recoveryRequired: false,
                        new InvalidOperationException("The live layout matches neither Undo endpoint."));
                }

                IReadOnlyList<PivotMutationStep> steps = BuildUndoSteps(
                    measureUndo?.UpsertSteps,
                    state.NamedSets.DeleteSteps,
                    new PivotMutationStep(
                        "restore prior PivotTable semantic layout",
                        state.Layout.Restore,
                        state.Layout.Apply),
                    state.NamedSets.UpsertSteps,
                    measureUndo?.DeleteSteps);

                try
                {
                    ExecutePrepared(
                        coordinator,
                        pivotTable,
                        steps,
                        state.Refresh,
                        () =>
                        {
                            layoutGateway.VerifySnapshot(
                                liveLayoutTarget,
                                state.BeforeLayout);
                            state.NamedSets.VerifyRollback();
                            measureUndo?.Verify();
                        });
                }
                catch (PivotMutationException failure)
                {
                    if (failure.RollbackCompleted)
                    {
                        try
                        {
                            state.Layout.VerifyDesired();
                            state.NamedSets.Verify();
                            measureUndo?.VerifyRollback();
                            ownership.RestoreBase(workbook, session);
                        }
                        catch (Exception recoveryFailure)
                        {
                            throw new PivotSemanticMutationException(
                                "Combined Undo failed and its exact post-Apply state could not be restored.",
                                rollbackCompleted: false,
                                recoveryRequired: true,
                                new AggregateException(failure, recoveryFailure));
                        }
                    }

                    throw new PivotSemanticMutationException(
                        failure.Message,
                        failure.RollbackCompleted,
                        recoveryRequired: !failure.RollbackCompleted,
                        failure);
                }

                try
                {
                    ownership.CommitSemanticState(workbook, session, state.BaseMetadata);
                }
                catch (Exception failure)
                {
                    throw new PivotSemanticMutationException(
                        "Excel completed and verified Undo, but ownership finalization was ambiguous. Retry Undo.",
                        rollbackCompleted: false,
                        recoveryRequired: true,
                        failure);
                }

                RemoveUndoState(workbook, setupId);
                return new PivotSemanticUndoResult(state.ApplyId, recovered: false);
            }
            finally
            {
                Exit();
            }
        }

        internal static void ExecutePrepared(
            PivotMutationCoordinator coordinator,
            object pivotTable,
            IReadOnlyList<PivotMutationStep> steps,
            Action refresh,
            Action verify)
        {
            if (coordinator == null) throw new ArgumentNullException(nameof(coordinator));
            coordinator.Execute(pivotTable, steps, refresh, verify);
        }

        internal static IReadOnlyList<PivotMutationStep> BuildUndoSteps(
            IEnumerable<PivotMutationStep>? measureUpserts,
            IEnumerable<PivotMutationStep> namedSetDeletes,
            PivotMutationStep layoutRestore,
            IEnumerable<PivotMutationStep> namedSetUpserts,
            IEnumerable<PivotMutationStep>? measureDeletes)
        {
            if (namedSetDeletes == null)
            {
                throw new ArgumentNullException(nameof(namedSetDeletes));
            }

            if (layoutRestore == null)
            {
                throw new ArgumentNullException(nameof(layoutRestore));
            }

            if (namedSetUpserts == null)
            {
                throw new ArgumentNullException(nameof(namedSetUpserts));
            }

            var result = new List<PivotMutationStep>();
            if (measureUpserts != null) result.AddRange(measureUpserts);
            result.AddRange(Invert(namedSetDeletes));
            result.Add(layoutRestore);
            result.AddRange(Invert(namedSetUpserts));
            if (measureDeletes != null) result.AddRange(measureDeletes);
            return new ReadOnlyCollection<PivotMutationStep>(result);
        }

        private static PivotPlusPendingSemanticApplyMetadata BuildUndoPending(
            SemanticUndoState state,
            PivotPlusWorkbookMetadata currentBase,
            PivotPlusPendingSemanticApplyMetadata? existingPending)
        {
            var beforeByKey = state.BaseMetadata.Artifacts.ToDictionary(
                ArtifactKey,
                StringComparer.OrdinalIgnoreCase);
            var afterByKey = currentBase.Artifacts.ToDictionary(
                ArtifactKey,
                StringComparer.OrdinalIgnoreCase);
            var transitions = new List<PivotPlusSemanticArtifactTransition>();
            foreach (PivotPlusSemanticArtifactTransition forward in state.Transitions)
            {
                string key = ArtifactKey(forward.Kind, forward.ArtifactId);
                beforeByKey.TryGetValue(key, out PivotPlusOwnedArtifact? before);
                afterByKey.TryGetValue(key, out PivotPlusOwnedArtifact? after);
                PivotPlusSemanticArtifactOperation operation;
                string beforeLive;
                string planned;
                switch (forward.Operation)
                {
                    case PivotPlusSemanticArtifactOperation.Create:
                        if (after == null)
                        {
                            throw new InvalidOperationException(
                                "An Apply-created artifact is missing before Undo.");
                        }

                        operation = PivotPlusSemanticArtifactOperation.Delete;
                        beforeLive = after.Fingerprint;
                        planned = after.Fingerprint;
                        break;
                    case PivotPlusSemanticArtifactOperation.Delete:
                        if (before == null)
                        {
                            throw new InvalidOperationException(
                                "A deleted artifact has no exact prior ownership receipt.");
                        }

                        operation = PivotPlusSemanticArtifactOperation.Create;
                        beforeLive = string.Empty;
                        planned = before.Fingerprint;
                        break;
                    case PivotPlusSemanticArtifactOperation.Update:
                        if (before == null || after == null)
                        {
                            throw new InvalidOperationException(
                                "An updated artifact is missing an Undo endpoint.");
                        }

                        operation = PivotPlusSemanticArtifactOperation.Update;
                        beforeLive = after.Fingerprint;
                        planned = before.Fingerprint;
                        break;
                    default:
                        throw new InvalidOperationException(
                            "The semantic transition cannot be reversed.");
                }

                transitions.Add(new PivotPlusSemanticArtifactTransition
                {
                    Kind = forward.Kind,
                    ArtifactId = forward.ArtifactId,
                    Operation = operation,
                    BeforeLiveFingerprint = beforeLive,
                    PlannedDefinitionFingerprint = planned
                });
            }

            string undoPlan = PivotPlusFingerprint.Create(
                "semantic.undo-plan.v1",
                state.PlanFingerprint + "|" + state.BeforeLayout.LayoutFingerprint + "|" +
                state.AfterLayout.LayoutFingerprint);
            var proposed = new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = "undo_" + state.ApplyId,
                PlanFingerprint = undoPlan,
                BeforePivotFingerprint = state.AfterLayout.LayoutFingerprint,
                ExpectedPivotFingerprint = state.BeforeLayout.LayoutFingerprint,
                Transitions = transitions
                    .OrderBy(item => item.Kind)
                    .ThenBy(item => item.ArtifactId, StringComparer.Ordinal)
                    .ToList()
            };
            if (existingPending != null) DemandSamePending(existingPending, proposed);
            return proposed;
        }

        private static IReadOnlyList<PivotMutationStep> Invert(
            IEnumerable<PivotMutationStep> original)
        {
            return new ReadOnlyCollection<PivotMutationStep>(
                original.Reverse().Select(step => new PivotMutationStep(
                    "undo " + step.Name,
                    step.Rollback,
                    step.Apply)).ToList());
        }

        private static string ArtifactKey(PivotPlusOwnedArtifact artifact) =>
            ArtifactKey(artifact.Kind, artifact.ArtifactId);

        private static string ArtifactKey(PivotPlusArtifactKind kind, string artifactId) =>
            ((int)kind).ToString(CultureInfo.InvariantCulture) + "|" + artifactId;

        private static PivotPlusUndoMetadata BuildUndo(
            string applyId,
            PivotSemanticLayoutSnapshot before,
            PivotSemanticLayoutSnapshot after,
            IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions,
            IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts)
        {
            var createdNames = new HashSet<string>(
                transitions.Where(item =>
                        item.Operation == PivotPlusSemanticArtifactOperation.Create)
                    .Select(item => item.ArtifactId),
                StringComparer.OrdinalIgnoreCase);
            var placements = new List<PivotPlusUndoFieldPlacement>();
            placements.AddRange(before.Rows.Select(item => UndoPlacement(
                item.UniqueName,
                item.CaptionFingerprint,
                PivotPlusFieldArea.Row,
                item.Position)));
            placements.AddRange(before.Columns.Select(item => UndoPlacement(
                item.UniqueName,
                item.CaptionFingerprint,
                PivotPlusFieldArea.Column,
                item.Position)));
            placements.AddRange(before.Values.Select(item => UndoPlacement(
                item.UniqueName,
                item.CaptionFingerprint + "|" + item.NumberFormatFingerprint,
                PivotPlusFieldArea.Data,
                item.Position)));
            return new PivotPlusUndoMetadata
            {
                ApplyId = applyId,
                BeforePivotFingerprint = before.LayoutFingerprint,
                AfterPivotFingerprint = after.LayoutFingerprint,
                CreatedArtifacts = finalArtifacts
                    .Where(item => createdNames.Contains(item.ArtifactId))
                    .Select(CloneArtifact)
                    .ToList(),
                PreviousFieldPlacements = placements
            };
        }

        private static PivotPlusUndoFieldPlacement UndoPlacement(
            string uniqueName,
            string stateFingerprint,
            PivotPlusFieldArea area,
            int oneBasedPosition)
        {
            return new PivotPlusUndoFieldPlacement
            {
                FieldFingerprint = PivotPlusFingerprint.Create(
                    "semantic.undo-field.v1",
                    uniqueName + "\u001f" + stateFingerprint),
                Area = area,
                Position = oneBasedPosition - 1
            };
        }

        private static void DemandNoUnsafeMdxMeasureReferences(
            PivotNamedSetWorkbookSnapshot snapshot,
            IEnumerable<PivotPlusSemanticArtifactTransition> measureTransitions,
            IEnumerable<string> managedNamedSetNames)
        {
            string[] changedMeasureNames = measureTransitions
                .Where(item => item.Kind == PivotPlusArtifactKind.Measure)
                .Select(item => "[Measures].[" + item.ArtifactId.Replace("]", "]]" ) + "]")
                .ToArray();
            if (changedMeasureNames.Length == 0) return;
            var managedSets = new HashSet<string>(
                managedNamedSetNames,
                StringComparer.OrdinalIgnoreCase);
            foreach (PivotCalculatedMemberReferenceSnapshot calculated in
                     snapshot.Pivots.SelectMany(item => item.CalculatedMembers)
                         .Where(item => !managedSets.Contains(item.Name)))
            {
                foreach (string measure in changedMeasureNames)
                {
                    if (MdxNamedSetReferenceScanner.MightReference(
                            calculated.RawFormula,
                            measure))
                    {
                        throw new InvalidOperationException(
                            "An unowned calculated member may depend on a measure changed by this Apply.");
                    }
                }
            }
        }

        private static string CreateLayoutPlanFingerprint(PivotSemanticLayoutPlan plan)
        {
            var canonical = new StringBuilder("semantic-layout-plan-v1");
            Append(canonical, (int)plan.ValuesAxis);
            Append(canonical, plan.ValuesPosition);
            foreach (PivotSemanticAxisPlacement item in plan.Rows.OrderBy(value => value.Position))
            {
                Append(canonical, "row");
                AppendAxis(canonical, item);
            }

            foreach (PivotSemanticAxisPlacement item in plan.Columns.OrderBy(value => value.Position))
            {
                Append(canonical, "column");
                AppendAxis(canonical, item);
            }

            foreach (PivotSemanticValuePlacement item in plan.Values.OrderBy(value => value.Position))
            {
                Append(canonical, item.Position);
                if (item.IsGeneratedMeasure)
                {
                    Append(canonical, "generated");
                    Append(canonical, item.DefinitionId ?? string.Empty);
                }
                else
                {
                    PivotExistingSemanticValueIdentity identity = item.ExistingDataField!;
                    Append(canonical, "existing");
                    Append(canonical, identity.UniqueName);
                    Append(canonical, identity.CurrentCaptionFingerprint);
                    Append(canonical, identity.CurrentNumberFormatFingerprint);
                    Append(canonical, identity.CurrentPosition);
                }
            }

            return PivotPlusFingerprint.Create("semantic.layout-plan.v1", canonical.ToString());
        }

        private static void AppendAxis(StringBuilder canonical, PivotSemanticAxisPlacement item)
        {
            Append(canonical, item.Position);
            if (item.IsGeneratedNamedSet)
            {
                Append(canonical, "generated");
                Append(canonical, item.DefinitionId ?? string.Empty);
                return;
            }

            PivotExistingAxisFieldIdentity identity = item.ExistingField!;
            Append(canonical, "existing");
            Append(canonical, identity.UniqueName);
            Append(canonical, identity.CurrentCaptionFingerprint);
            Append(canonical, (int)identity.CurrentArea);
            Append(canonical, identity.CurrentPosition);
        }

        private static string CreateCombinedPlanFingerprint(
            string measurePlanFingerprint,
            string namedSetPlanFingerprint,
            string layoutPlanFingerprint,
            PivotDaxCompilation dax,
            PivotMdxCompilation mdx)
        {
            var canonical = new StringBuilder("semantic-combined-plan-v1");
            Append(canonical, measurePlanFingerprint);
            Append(canonical, namedSetPlanFingerprint);
            Append(canonical, layoutPlanFingerprint);
            foreach (OwnedPivotMeasureDefinition measure in dax.Measures.OrderBy(
                         item => item.DefinitionId,
                         StringComparer.Ordinal))
            {
                Append(canonical, measure.DefinitionFingerprint);
                Append(canonical, measure.FormulaFingerprint);
            }

            Append(canonical, mdx.CompilationFingerprint);
            return PivotPlusFingerprint.Create("semantic.combined-plan.v1", canonical.ToString());
        }

        private static void DemandSameTarget(
            params Core.PivotPlus.PivotTargetIdentity[] targets)
        {
            Core.PivotPlus.PivotTargetIdentity first = targets[0];
            if (targets.Skip(1).Any(target =>
                    !string.Equals(target.WorkbookId, first.WorkbookId, StringComparison.Ordinal) ||
                    !string.Equals(target.WorksheetName, first.WorksheetName, StringComparison.Ordinal) ||
                    !string.Equals(target.PivotTableName, first.PivotTableName, StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The semantic participants are not bound to one exact PivotTable target.");
            }
        }

        private static void DemandSamePending(
            PivotPlusPendingSemanticApplyMetadata existing,
            PivotPlusPendingSemanticApplyMetadata proposed)
        {
            bool equal = string.Equals(existing.ApplyId, proposed.ApplyId, StringComparison.Ordinal) &&
                         string.Equals(existing.PlanFingerprint, proposed.PlanFingerprint, StringComparison.Ordinal) &&
                         string.Equals(existing.BeforePivotFingerprint, proposed.BeforePivotFingerprint, StringComparison.Ordinal) &&
                         string.Equals(existing.ExpectedPivotFingerprint, proposed.ExpectedPivotFingerprint, StringComparison.Ordinal) &&
                         existing.Transitions.Count == proposed.Transitions.Count;
            if (equal)
            {
                IReadOnlyList<string> left = existing.Transitions.Select(TransitionKey)
                    .OrderBy(item => item, StringComparer.Ordinal).ToList();
                IReadOnlyList<string> right = proposed.Transitions.Select(TransitionKey)
                    .OrderBy(item => item, StringComparer.Ordinal).ToList();
                equal = left.SequenceEqual(right, StringComparer.Ordinal);
            }

            if (!equal)
            {
                throw new InvalidOperationException(
                    "The recomputed semantic journal differs from the durable pending Apply.");
            }
        }

        private static string TransitionKey(PivotPlusSemanticArtifactTransition item) =>
            ((int)item.Kind).ToString(CultureInfo.InvariantCulture) + "|" +
            ((int)item.Operation).ToString(CultureInfo.InvariantCulture) + "|" +
            item.ArtifactId + "|" + item.BeforeLiveFingerprint + "|" +
            item.PlannedDefinitionFingerprint;

        private static PivotPlusSemanticArtifactTransition Clone(
            PivotPlusSemanticArtifactTransition item) => new()
        {
            Kind = item.Kind,
            ArtifactId = item.ArtifactId,
            Operation = item.Operation,
            BeforeLiveFingerprint = item.BeforeLiveFingerprint,
            PlannedDefinitionFingerprint = item.PlannedDefinitionFingerprint
        };

        private static PivotPlusOwnedArtifact CloneArtifact(PivotPlusOwnedArtifact item) => new()
        {
            Kind = item.Kind,
            ArtifactId = item.ArtifactId,
            Fingerprint = item.Fingerprint
        };

        private static void Append(StringBuilder builder, string value)
        {
            builder.Append('|').Append(value.Length).Append(':').Append(value);
        }

        private static void Append(StringBuilder builder, int value) =>
            Append(builder, value.ToString(CultureInfo.InvariantCulture));

        private static void RememberLayoutSeed(
            object workbook,
            string setupId,
            string applyId,
            string planFingerprint,
            PivotSemanticLayoutSnapshot before)
        {
            LayoutSeedLedger ledger = LayoutSeeds.GetValue(
                workbook,
                _ => new LayoutSeedLedger());
            lock (ledger.Synchronization)
            {
                ledger.Items[setupId] = new LayoutSeed(
                    applyId,
                    planFingerprint,
                    before);
            }
        }

        private static bool TryGetLayoutSeed(
            object workbook,
            string setupId,
            string applyId,
            string planFingerprint,
            out PivotSemanticLayoutSnapshot? before)
        {
            before = null;
            if (!LayoutSeeds.TryGetValue(workbook, out LayoutSeedLedger? ledger))
            {
                return false;
            }

            lock (ledger.Synchronization)
            {
                if (!ledger.Items.TryGetValue(setupId, out LayoutSeed? seed) ||
                    !string.Equals(seed.ApplyId, applyId, StringComparison.Ordinal) ||
                    !string.Equals(
                        seed.PlanFingerprint,
                        planFingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                before = seed.Before;
                return true;
            }
        }

        private static void RememberUndoState(
            object workbook,
            string setupId,
            SemanticUndoState state)
        {
            SemanticUndoLedger ledger = UndoStates.GetValue(
                workbook,
                _ => new SemanticUndoLedger());
            lock (ledger.Synchronization) ledger.Items[setupId] = state;
        }

        private static bool TryGetUndoState(
            object workbook,
            string setupId,
            out SemanticUndoState? state)
        {
            state = null;
            if (!UndoStates.TryGetValue(workbook, out SemanticUndoLedger? ledger))
            {
                return false;
            }

            lock (ledger.Synchronization)
            {
                return ledger.Items.TryGetValue(setupId, out state);
            }
        }

        private static void RemoveUndoState(object workbook, string setupId)
        {
            if (!UndoStates.TryGetValue(workbook, out SemanticUndoLedger? ledger)) return;
            lock (ledger.Synchronization) ledger.Items.Remove(setupId);
        }

        private void Enter()
        {
            lock (synchronization)
            {
                if (active)
                {
                    throw new InvalidOperationException(
                        "A combined PivotTable+ semantic mutation is already active.");
                }

                active = true;
            }
        }

        private void Exit()
        {
            lock (synchronization) active = false;
        }

        private sealed class LayoutSeedLedger
        {
            public object Synchronization { get; } = new object();
            public Dictionary<string, LayoutSeed> Items { get; } =
                new Dictionary<string, LayoutSeed>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class LayoutSeed
        {
            public LayoutSeed(
                string applyId,
                string planFingerprint,
                PivotSemanticLayoutSnapshot before)
            {
                ApplyId = applyId;
                PlanFingerprint = planFingerprint;
                Before = before;
            }

            public string ApplyId { get; }
            public string PlanFingerprint { get; }
            public PivotSemanticLayoutSnapshot Before { get; }
        }

        private sealed class SemanticUndoLedger
        {
            public object Synchronization { get; } = new object();
            public Dictionary<string, SemanticUndoState> Items { get; } =
                new Dictionary<string, SemanticUndoState>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class SemanticUndoState
        {
            public SemanticUndoState(
                string applyId,
                string planFingerprint,
                Core.PivotPlus.PivotTargetIdentity target,
                PivotPlusWorkbookMetadata baseMetadata,
                IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts,
                IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions,
                PivotModelMeasureArtifactUndoContribution? measureUndo,
                PivotNamedSetPreparedMutation namedSets,
                PivotSemanticPreparedPlacement layout,
                BoundPivotSemanticLayoutTarget layoutTarget,
                PivotSemanticLayoutSnapshot beforeLayout,
                PivotSemanticLayoutSnapshot afterLayout,
                Action refresh)
            {
                ApplyId = applyId;
                PlanFingerprint = planFingerprint;
                Target = target;
                BaseMetadata = baseMetadata;
                FinalArtifacts = finalArtifacts;
                Transitions = transitions;
                MeasureUndo = measureUndo;
                NamedSets = namedSets;
                Layout = layout;
                LayoutTarget = layoutTarget;
                BeforeLayout = beforeLayout;
                AfterLayout = afterLayout;
                Refresh = refresh;
            }

            public string ApplyId { get; }
            public string PlanFingerprint { get; }
            public Core.PivotPlus.PivotTargetIdentity Target { get; }
            public PivotPlusWorkbookMetadata BaseMetadata { get; }
            public IReadOnlyList<PivotPlusOwnedArtifact> FinalArtifacts { get; }
            public IReadOnlyList<PivotPlusSemanticArtifactTransition> Transitions { get; }
            public PivotModelMeasureArtifactUndoContribution? MeasureUndo { get; }
            public PivotNamedSetPreparedMutation NamedSets { get; }
            public PivotSemanticPreparedPlacement Layout { get; }
            public BoundPivotSemanticLayoutTarget LayoutTarget { get; }
            public PivotSemanticLayoutSnapshot BeforeLayout { get; }
            public PivotSemanticLayoutSnapshot AfterLayout { get; }
            public Action Refresh { get; }
        }
    }
}
