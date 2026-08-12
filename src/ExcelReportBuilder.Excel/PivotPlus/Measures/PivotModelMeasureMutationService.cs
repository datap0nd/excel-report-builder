using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    /// <summary>
    /// Internal transaction participant for a future combined semantic Apply.
    /// It deliberately exposes no public DAX or COM surface: a higher-level
    /// Excel coordinator can concatenate these bounded steps with named-set
    /// steps, journal all transitions once, refresh once, verify each
    /// participant, and commit metadata once.
    /// </summary>
    internal sealed class PivotModelMeasurePreparedMutation
    {
        private readonly Action refresh;
        private readonly Func<ModelMeasureWorkbookSnapshot> verify;
        private readonly Action verifyRollback;
        private readonly Func<ModelMeasureWorkbookSnapshot, IReadOnlyList<PivotPlusOwnedArtifact>>
            buildArtifacts;
        private readonly Action<string, string> primeUndoContribution;
        private readonly Func<string, string, ModelMeasureWorkbookSnapshot,
            PivotModelMeasureUndoContribution?> buildUndoContribution;

        public PivotModelMeasurePreparedMutation(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            PivotPlusPendingSemanticApplyMetadata pending,
            IReadOnlyList<PivotMutationStep> upsertSteps,
            IReadOnlyList<PivotMutationStep> placementSteps,
            IReadOnlyList<PivotMutationStep> deleteSteps,
            bool isNoChange,
            int createCount,
            int updateCount,
            int deleteCount,
            Action refresh,
            Func<ModelMeasureWorkbookSnapshot> verify,
            Action verifyRollback,
            Func<ModelMeasureWorkbookSnapshot, IReadOnlyList<PivotPlusOwnedArtifact>> buildArtifacts,
            Action<string, string> primeUndoContribution,
            Func<string, string, ModelMeasureWorkbookSnapshot,
                PivotModelMeasureUndoContribution?> buildUndoContribution)
        {
            Target = target;
            Before = before;
            Pending = pending;
            UpsertSteps = upsertSteps;
            PlacementSteps = placementSteps;
            DeleteSteps = deleteSteps;
            Steps = new ReadOnlyCollection<PivotMutationStep>(
                upsertSteps.Concat(placementSteps).Concat(deleteSteps).ToList());
            IsNoChange = isNoChange;
            CreateCount = createCount;
            UpdateCount = updateCount;
            DeleteCount = deleteCount;
            this.refresh = refresh;
            this.verify = verify;
            this.verifyRollback = verifyRollback;
            this.buildArtifacts = buildArtifacts;
            this.primeUndoContribution = primeUndoContribution;
            this.buildUndoContribution = buildUndoContribution;
        }

        public BoundModelMeasureTarget Target { get; }

        public ModelMeasureWorkbookSnapshot Before { get; }

        public PivotPlusPendingSemanticApplyMetadata Pending { get; }

        public IReadOnlyList<PivotMutationStep> Steps { get; }

        /// <summary>Run before named-set upserts in a combined Apply.</summary>
        public IReadOnlyList<PivotMutationStep> UpsertSteps { get; }

        /// <summary>Run in the combined placement phase.</summary>
        public IReadOnlyList<PivotMutationStep> PlacementSteps { get; }

        /// <summary>Run after named-set deletes in a combined Apply.</summary>
        public IReadOnlyList<PivotMutationStep> DeleteSteps { get; }

        public bool IsNoChange { get; }

        public int CreateCount { get; }

        public int UpdateCount { get; }

        public int DeleteCount { get; }

        public void Refresh()
        {
            refresh();
        }

        public ModelMeasureWorkbookSnapshot Verify()
        {
            return verify();
        }

        /// <summary>
        /// Proves that a combined coordinator rollback restored the exact
        /// workbook snapshot captured when this participant was prepared.
        /// </summary>
        public void VerifyRollback()
        {
            verifyRollback();
        }

        public IReadOnlyList<PivotPlusOwnedArtifact> BuildArtifacts(
            ModelMeasureWorkbookSnapshot verified)
        {
            return buildArtifacts(verified);
        }

        /// <summary>
        /// Builds the session-only measure contribution to a combined semantic
        /// Undo. The higher-level coordinator supplies its combined Apply and
        /// plan identities; formulas remain only in this in-memory object.
        /// </summary>
        public void PrimeUndoContribution(
            string combinedApplyId,
            string combinedPlanFingerprint)
        {
            PivotPlusMetadataValidator.ValidateId(
                combinedApplyId,
                "combined semantic Apply identifier");
            PivotPlusMetadataValidator.ValidateFingerprint(
                combinedPlanFingerprint,
                "combined semantic plan fingerprint");
            primeUndoContribution(combinedApplyId, combinedPlanFingerprint);
        }

        public PivotModelMeasureUndoContribution? BuildUndoContribution(
            string combinedApplyId,
            string combinedPlanFingerprint,
            ModelMeasureWorkbookSnapshot verified)
        {
            PivotPlusMetadataValidator.ValidateId(
                combinedApplyId,
                "combined semantic Apply identifier");
            PivotPlusMetadataValidator.ValidateFingerprint(
                combinedPlanFingerprint,
                "combined semantic plan fingerprint");
            return buildUndoContribution(
                combinedApplyId,
                combinedPlanFingerprint,
                verified);
        }
    }

    /// <summary>
    /// Executable, session-only measure contribution to a combined semantic
    /// Undo. A higher-level coordinator can order these phase buckets with
    /// named-set work, refresh once, and verify both forward and rollback
    /// states without duplicating measure internals.
    /// </summary>
    internal sealed class PivotModelMeasurePreparedUndo
    {
        private readonly Action refresh;
        private readonly Action verify;
        private readonly Action verifyRollback;

        public PivotModelMeasurePreparedUndo(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            IReadOnlyList<PivotMutationStep> upsertSteps,
            IReadOnlyList<PivotMutationStep> placementSteps,
            IReadOnlyList<PivotMutationStep> deleteSteps,
            Action refresh,
            Action verify,
            Action verifyRollback)
        {
            Target = target;
            Before = before;
            UpsertSteps = upsertSteps;
            PlacementSteps = placementSteps;
            DeleteSteps = deleteSteps;
            Steps = new ReadOnlyCollection<PivotMutationStep>(
                upsertSteps.Concat(placementSteps).Concat(deleteSteps).ToList());
            this.refresh = refresh;
            this.verify = verify;
            this.verifyRollback = verifyRollback;
        }

        public BoundModelMeasureTarget Target { get; }

        public ModelMeasureWorkbookSnapshot Before { get; }

        /// <summary>Restore measures before named-set restoration.</summary>
        public IReadOnlyList<PivotMutationStep> UpsertSteps { get; }

        /// <summary>Restore the selected Values sequence in the shared placement phase.</summary>
        public IReadOnlyList<PivotMutationStep> PlacementSteps { get; }

        /// <summary>Remove Apply-created measures after named-set deletion.</summary>
        public IReadOnlyList<PivotMutationStep> DeleteSteps { get; }

        public IReadOnlyList<PivotMutationStep> Steps { get; }

        public void Refresh()
        {
            refresh();
        }

        public void Verify()
        {
            verify();
        }

        public void VerifyRollback()
        {
            verifyRollback();
        }
    }

    /// <summary>
    /// Session-only state contributed by the measure participant to a future
    /// combined Measure + NamedSet Undo. It is never serialized because its
    /// snapshots contain the exact prior host formulas needed for rollback.
    /// </summary>
    internal sealed class PivotModelMeasureUndoContribution
    {
        public PivotModelMeasureUndoContribution(
            string applyId,
            string planFingerprint,
            PivotTargetIdentity target,
            ModelMeasureWorkbookSnapshot before,
            ModelMeasureWorkbookSnapshot after,
            IReadOnlyList<PivotPlusOwnedArtifact> beforeOwnedArtifacts,
            IReadOnlyList<PivotPlusOwnedArtifact> afterOwnedArtifacts,
            PivotPlusUndoMetadata workbookUndo)
        {
            ApplyId = applyId;
            PlanFingerprint = planFingerprint;
            Target = target;
            Before = before;
            After = after;
            BeforeOwnedArtifacts = beforeOwnedArtifacts;
            AfterOwnedArtifacts = afterOwnedArtifacts;
            WorkbookUndo = workbookUndo;
        }

        public string ApplyId { get; }

        public string PlanFingerprint { get; }

        public PivotTargetIdentity Target { get; }

        public ModelMeasureWorkbookSnapshot Before { get; }

        public ModelMeasureWorkbookSnapshot After { get; }

        public IReadOnlyList<PivotPlusOwnedArtifact> BeforeOwnedArtifacts { get; }

        public IReadOnlyList<PivotPlusOwnedArtifact> AfterOwnedArtifacts { get; }

        public PivotPlusUndoMetadata WorkbookUndo { get; }
    }

    /// <summary>
    /// Trusted coordinator proof for replaying the measure slice of one
    /// combined Measure + NamedSet journal. The service verifies both the
    /// durable combined identity and its own freshly recomputed slice.
    /// </summary>
    internal sealed class PivotModelMeasureParticipantRetryBinding
    {
        public PivotModelMeasureParticipantRetryBinding(
            string combinedPlanFingerprint,
            string measurePlanFingerprint)
        {
            PivotPlusMetadataValidator.ValidateFingerprint(
                combinedPlanFingerprint,
                "combined semantic plan fingerprint");
            PivotPlusMetadataValidator.ValidateFingerprint(
                measurePlanFingerprint,
                "measure participant plan fingerprint");
            CombinedPlanFingerprint = combinedPlanFingerprint;
            MeasurePlanFingerprint = measurePlanFingerprint;
        }

        public string CombinedPlanFingerprint { get; }

        public string MeasurePlanFingerprint { get; }
    }

    /// <summary>
    /// Applies typed, compiler-produced model measures to one selected native
    /// Data Model PivotTable. This service never accepts arbitrary DAX.
    /// </summary>
    public sealed partial class PivotModelMeasureMutationService
    {
        private const int MaximumMeasures = 128;
        private const int MaximumValuePlacements = 256;

        private readonly IPivotModelMeasureGateway gateway;
        private readonly IPivotModelMeasureOwnershipStore ownership;
        private readonly IWorkbookIdentityResolver workbookIdentity;
        private readonly PivotMutationCoordinator coordinator;
        private readonly ConditionalWeakTable<object, UndoLedger> undoLedgers =
            new ConditionalWeakTable<object, UndoLedger>();
        private readonly object synchronization = new object();
        private bool applyActive;

        public PivotModelMeasureMutationService()
            : this(
                new LateBoundPivotModelMeasureGateway(),
                new PivotModelMeasureOwnershipStore(),
                new StoredWorkbookIdentityResolver(),
                new PivotMutationCoordinator())
        {
        }

        internal PivotModelMeasureMutationService(
            IPivotModelMeasureGateway gateway,
            IPivotModelMeasureOwnershipStore ownership,
            IWorkbookIdentityResolver workbookIdentity,
            PivotMutationCoordinator coordinator)
        {
            this.gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
            this.ownership = ownership ?? throw new ArgumentNullException(nameof(ownership));
            this.workbookIdentity = workbookIdentity ??
                throw new ArgumentNullException(nameof(workbookIdentity));
            this.coordinator = coordinator ?? throw new ArgumentNullException(nameof(coordinator));
        }

        public PivotModelMeasureApplyResult Apply(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId,
            PivotDaxCompilation compilation,
            PivotMeasurePlacementPlan placement)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
            if (placement == null) throw new ArgumentNullException(nameof(placement));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");

            Enter();
            try
            {
                IReadOnlyList<DesiredModelMeasure> definitions = CompileDefinitions(
                    setupId,
                    compilation);
                IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById =
                    definitions.ToDictionary(
                        definition => definition.DefinitionId,
                        StringComparer.OrdinalIgnoreCase);
                BoundModelMeasureTarget target = gateway.Bind(workbook, pivotTable, context);
                ModelMeasureWorkbookSnapshot before = gateway.Capture(target);
                PivotPlusWorkbookMetadata baseMetadata = ownership.ReadBase(
                    workbook,
                    setupId,
                    target.Identity,
                    out PivotPlusPendingSemanticApplyMetadata? existingPending);
                PendingApplyUndoSeed? applyUndoSeed = null;
                if (existingPending != null)
                {
                    TryGetPendingApplyUndoSeed(
                        workbook,
                        setupId,
                        existingPending.ApplyId,
                        existingPending.PlanFingerprint,
                        target.Identity,
                        out applyUndoSeed);
                }

                MeasureApplyPlan plan = BuildPlan(
                    baseMetadata,
                    existingPending,
                    before,
                    definitions,
                    definitionsById,
                    placement,
                    pendingPreview: applyUndoSeed?.Before);
                PivotModelMeasurePreparedMutation prepared = CreatePreparedMutation(
                    workbook,
                    setupId,
                    target,
                    before,
                    plan,
                    definitions,
                    definitionsById,
                    placement,
                    baseMetadata.Artifacts.ToList(),
                    allowNewUndoSeed: existingPending == null);
                if (prepared.IsNoChange && existingPending == null)
                {
                    if (TryPromoteRecoveredUndo(
                            workbook,
                            setupId,
                            target.Identity,
                            prepared.Pending.PlanFingerprint,
                            before,
                            baseMetadata,
                            out MeasureUndoState? recoveredUndo) &&
                        recoveredUndo != null)
                    {
                        return new PivotModelMeasureApplyResult(
                            recoveredUndo.ApplyId,
                            PivotModelMeasureApplyStatus.NoChange,
                            0,
                            0,
                            0,
                            undoAvailable: true);
                    }

                    bool existingUndoAvailable = TryGetUndo(
                            workbook,
                            setupId,
                            out MeasureUndoState? existingUndo) &&
                        existingUndo != null &&
                        UndoIsUsable(existingUndo, target.Identity, before);
                    return new PivotModelMeasureApplyResult(
                        existingUndoAvailable ? existingUndo!.ApplyId : string.Empty,
                        PivotModelMeasureApplyStatus.NoChange,
                        0,
                        0,
                        0,
                        undoAvailable: existingUndoAvailable);
                }

                prepared.PrimeUndoContribution(
                    prepared.Pending.ApplyId,
                    prepared.Pending.PlanFingerprint);
                workbookIdentity.Persist(workbook, target.Identity.WorkbookId);
                PivotModelMeasureOwnershipSession session;
                try
                {
                    session = ownership.Journal(
                        workbook,
                        setupId,
                        target.Identity,
                        prepared.Pending);
                }
                catch (Exception journalFailure)
                {
                    throw new PivotModelMeasureMutationException(
                        "The semantic ownership journal could not be confirmed; retry the identical Apply before changing the model.",
                        rollbackCompleted: true,
                        recoveryRequired: true,
                        journalFailure);
                }

                ModelMeasureWorkbookSnapshot? after = null;
                try
                {
                    coordinator.Execute(
                        prepared.Target.PivotTable,
                        prepared.Steps,
                        prepared.Refresh,
                        () => after = prepared.Verify());
                }
                catch (PivotMutationException mutationFailure)
                {
                    // A retry starts from a durable pending checkpoint, which
                    // may already contain a committed create/update/delete.
                    // Coordinator rollback can restore only the retry-start
                    // snapshot, not the pre-checkpoint active truth. Keep that
                    // journal until an identical retry reaches verification.
                    Exception failure = mutationFailure;
                    bool exactRollback = false;
                    if (mutationFailure.RollbackCompleted)
                    {
                        try
                        {
                            DemandExactWorkbookSnapshot(
                                before,
                                gateway.Capture(target),
                                "Apply rollback");
                            exactRollback = true;
                        }
                        catch (Exception verificationFailure)
                        {
                            failure = new AggregateException(
                                mutationFailure,
                                verificationFailure);
                        }
                    }

                    bool canClearJournal = exactRollback && existingPending == null;
                    if (canClearJournal)
                    {
                        try
                        {
                            ownership.RestoreBase(workbook, session);
                            ForgetPendingApplyUndoSeed(workbook, setupId);
                        }
                        catch (Exception ownershipFailure)
                        {
                            throw new PivotModelMeasureMutationException(
                                "The measure Apply failed and Excel was restored, but the pending ownership journal could not be cleared.",
                                rollbackCompleted: true,
                                recoveryRequired: true,
                                new AggregateException(failure, ownershipFailure));
                        }
                    }

                    throw new PivotModelMeasureMutationException(
                        canClearJournal
                            ? "The measure Apply failed and the prior model and PivotTable layout were restored."
                            : "The measure Apply failed and the pending semantic change requires an identical retry.",
                        exactRollback,
                        recoveryRequired: !canClearJournal,
                        failure);
                }
                catch (Exception executionFailure)
                {
                    Exception failure = executionFailure;
                    bool exactRollback = false;
                    try
                    {
                        DemandExactWorkbookSnapshot(
                            before,
                            gateway.Capture(target),
                            "Apply setup failure");
                        exactRollback = true;
                    }
                    catch (Exception verificationFailure)
                    {
                        failure = new AggregateException(executionFailure, verificationFailure);
                    }

                    bool canClearJournal = exactRollback && existingPending == null;
                    if (canClearJournal)
                    {
                        try
                        {
                            ownership.RestoreBase(workbook, session);
                            ForgetPendingApplyUndoSeed(workbook, setupId);
                        }
                        catch (Exception ownershipFailure)
                        {
                            throw new PivotModelMeasureMutationException(
                                "Apply did not start, but the pending ownership journal could not be cleared.",
                                rollbackCompleted: true,
                                recoveryRequired: true,
                                new AggregateException(failure, ownershipFailure));
                        }
                    }

                    throw new PivotModelMeasureMutationException(
                        canClearJournal
                            ? "The measure Apply could not start; the exact prior state was retained."
                            : "The measure Apply state is ambiguous and requires an identical retry.",
                        exactRollback,
                        recoveryRequired: !canClearJournal,
                        failure);
                }

                if (after == null)
                {
                    throw new InvalidOperationException(
                        "The measure Apply completed without a verified workbook snapshot.");
                }

                IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts =
                    prepared.BuildArtifacts(after);
                PivotModelMeasureUndoContribution? contribution =
                    prepared.BuildUndoContribution(
                        prepared.Pending.ApplyId,
                        prepared.Pending.PlanFingerprint,
                        after);
                MeasureUndoState? undoState = contribution == null
                    ? null
                    : new MeasureUndoState(contribution);

                if (undoState != null)
                {
                    RememberPendingApplyUndo(workbook, setupId, undoState);
                }
                try
                {
                    ownership.Commit(
                        workbook,
                        session,
                        finalArtifacts,
                        undoState?.WorkbookUndo);
                }
                catch (Exception commitFailure)
                {
                    throw new PivotModelMeasureMutationException(
                        "Excel contains the verified measure change, but durable ownership could not be finalized. Retry the identical Apply to recover.",
                        rollbackCompleted: false,
                        recoveryRequired: true,
                        commitFailure);
                }

                if (undoState != null)
                {
                    RememberUndo(workbook, setupId, undoState);
                    ForgetPendingApplyUndo(workbook, setupId);
                }
                else
                {
                    ForgetUndo(workbook, setupId);
                }
                ForgetPendingApplyUndoSeed(workbook, setupId);

                return new PivotModelMeasureApplyResult(
                    prepared.Pending.ApplyId,
                    PivotModelMeasureApplyStatus.Applied,
                    prepared.CreateCount,
                    prepared.UpdateCount,
                    prepared.DeleteCount,
                    undoAvailable: undoState != null);
            }
            finally
            {
                Exit();
            }
        }

        public void Undo(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");

            Enter();
            try
            {
                if (!TryGetUndo(workbook, setupId, out MeasureUndoState? undo) || undo == null)
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "Measure Undo is available only in the add-in session that performed the Apply; prior DAX is never stored in workbook metadata.");
                }

                BoundModelMeasureTarget target = gateway.Bind(workbook, pivotTable, context);
                DemandSameTarget(undo.Target, target.Identity);
                ModelMeasureWorkbookSnapshot current = gateway.Capture(target);
                PivotPlusWorkbookMetadata baseMetadata = ownership.ReadBase(
                    workbook,
                    setupId,
                    target.Identity,
                    out PivotPlusPendingSemanticApplyMetadata? existingPending);
                PivotPlusPendingSemanticApplyMetadata pending = BuildUndoPending(undo);

                if (IsExactUndoFinal(undo, current))
                {
                    if (existingPending == null)
                    {
                        DemandExactActiveMeasureOwnership(
                            baseMetadata,
                            undo.BeforeOwnedArtifacts,
                            "finalized Undo");
                        ForgetUndo(workbook, setupId);
                        return;
                    }

                    DemandSamePendingPlan(existingPending, pending);
                    workbookIdentity.Persist(workbook, target.Identity.WorkbookId);
                    PivotModelMeasureOwnershipSession resumed = JournalUndo(
                        workbook,
                        setupId,
                        target.Identity,
                        pending);
                    CommitUndo(workbook, setupId, resumed, undo);
                    return;
                }

                DemandExactActiveMeasureOwnership(
                    baseMetadata,
                    undo.AfterOwnedArtifacts,
                    "Undo start");
                if (existingPending != null)
                {
                    DemandSamePendingPlan(existingPending, pending);
                    try
                    {
                        DemandUndoAfterState(undo, current);
                    }
                    catch (PivotModelMeasureUndoUnavailableException)
                    {
                        DemandUndoIntermediateState(undo, current);
                    }
                }
                else
                {
                    DemandUndoAfterState(undo, current);
                    undo = undo.WithUndoStart(current);
                    RememberUndo(workbook, setupId, undo);
                }

                workbookIdentity.Persist(workbook, target.Identity.WorkbookId);
                PivotModelMeasureOwnershipSession session = JournalUndo(
                    workbook,
                    setupId,
                    target.Identity,
                    pending);

                try
                {
                    ExecuteUndo(target, current, undo);
                }
                catch (PivotMutationException mutationFailure)
                {
                    Exception failure = mutationFailure;
                    bool exactRollback = false;
                    if (mutationFailure.RollbackCompleted)
                    {
                        try
                        {
                            DemandExactWorkbookSnapshot(
                                current,
                                gateway.Capture(target),
                                "Undo rollback");
                            exactRollback = true;
                        }
                        catch (Exception verificationFailure)
                        {
                            failure = new AggregateException(
                                mutationFailure,
                                verificationFailure);
                        }
                    }

                    bool canClearJournal = exactRollback && existingPending == null;
                    if (canClearJournal)
                    {
                        try
                        {
                            ownership.RestoreBase(workbook, session);
                        }
                        catch (Exception ownershipFailure)
                        {
                            RememberUndo(workbook, setupId, undo);
                            throw new PivotModelMeasureMutationException(
                                "Undo failed and Excel was restored, but its pending ownership journal could not be cleared.",
                                rollbackCompleted: true,
                                recoveryRequired: true,
                                new AggregateException(failure, ownershipFailure));
                        }
                    }

                    RememberUndo(workbook, setupId, undo);
                    throw new PivotModelMeasureMutationException(
                        canClearJournal
                            ? "Undo failed and the post-Apply state was restored."
                            : exactRollback
                                ? "Undo retry failed; its durable recovery journal remains pending."
                                : "Undo failed and its exact recovery journal must be retained.",
                        exactRollback,
                        recoveryRequired: !canClearJournal,
                        failure);
                }
                catch (Exception executionFailure)
                {
                    Exception failure = executionFailure;
                    bool exactRollback = false;
                    try
                    {
                        DemandExactWorkbookSnapshot(
                            current,
                            gateway.Capture(target),
                            "Undo setup failure");
                        exactRollback = true;
                    }
                    catch (Exception verificationFailure)
                    {
                        failure = new AggregateException(
                            executionFailure,
                            verificationFailure);
                    }

                    bool canClearJournal = exactRollback && existingPending == null;
                    if (canClearJournal)
                    {
                        try
                        {
                            ownership.RestoreBase(workbook, session);
                        }
                        catch (Exception ownershipFailure)
                        {
                            RememberUndo(workbook, setupId, undo);
                            throw new PivotModelMeasureMutationException(
                                "Undo could not start and its pending ownership journal could not be cleared.",
                                rollbackCompleted: true,
                                recoveryRequired: true,
                                new AggregateException(failure, ownershipFailure));
                        }
                    }

                    RememberUndo(workbook, setupId, undo);
                    throw new PivotModelMeasureMutationException(
                        canClearJournal
                            ? "Undo could not start; the exact post-Apply state was retained."
                            : exactRollback
                                ? "Undo retry could not start; its durable recovery journal remains pending."
                                : "Undo setup failed and its exact recovery journal must be retained.",
                        exactRollback,
                        recoveryRequired: !canClearJournal,
                        failure);
                }

                CommitUndo(workbook, setupId, session, undo);
            }
            finally
            {
                Exit();
            }
        }

        internal PivotModelMeasurePreparedMutation PrepareParticipant(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId,
            PivotDaxCompilation compilation,
            PivotMeasurePlacementPlan placement,
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            PivotModelMeasureParticipantRetryBinding? retryBinding = null)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
            if (placement == null) throw new ArgumentNullException(nameof(placement));
            if (baseMetadata == null) throw new ArgumentNullException(nameof(baseMetadata));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");

            IReadOnlyList<DesiredModelMeasure> definitions = CompileDefinitions(
                setupId,
                compilation);
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById =
                definitions.ToDictionary(
                    definition => definition.DefinitionId,
                    StringComparer.OrdinalIgnoreCase);
            BoundModelMeasureTarget target = gateway.Bind(workbook, pivotTable, context);
            ModelMeasureWorkbookSnapshot before = gateway.Capture(target);
            DemandParticipantBase(baseMetadata, setupId, target.Identity);
            DemandParticipantPending(baseMetadata, existingPending);
            if (existingPending != null)
            {
                if (retryBinding == null || !string.Equals(
                        existingPending.PlanFingerprint,
                        retryBinding.CombinedPlanFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The combined semantic retry was not bound to the exact durable combined plan.");
                }
            }

            PendingApplyUndoSeed? applyUndoSeed = null;
            if (existingPending != null)
            {
                TryGetPendingApplyUndoSeed(
                    workbook,
                    setupId,
                    existingPending.ApplyId,
                    existingPending.PlanFingerprint,
                    target.Identity,
                    out applyUndoSeed);
            }
            PivotPlusPendingSemanticApplyMetadata? measurePending =
                CreateParticipantPendingSlice(existingPending);
            MeasureApplyPlan plan = BuildPlan(
                baseMetadata,
                measurePending,
                before,
                definitions,
                definitionsById,
                placement,
                compositeParticipant: existingPending != null,
                pendingPreview: applyUndoSeed?.Before);
            if (existingPending != null && !string.Equals(
                    plan.Pending.PlanFingerprint,
                    retryBinding!.MeasurePlanFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The recomputed measure placement does not match the exact pending participant plan.");
            }

            return CreatePreparedMutation(
                workbook,
                setupId,
                target,
                before,
                plan,
                definitions,
                definitionsById,
                placement,
                baseMetadata.Artifacts.ToList(),
                allowNewUndoSeed: existingPending == null);
        }

        internal PivotModelMeasurePreparedUndo PrepareUndoParticipant(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            PivotModelMeasureUndoContribution contribution)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (contribution == null) throw new ArgumentNullException(nameof(contribution));

            BoundModelMeasureTarget target = gateway.Bind(workbook, pivotTable, context);
            DemandSameTarget(contribution.Target, target.Identity);
            ModelMeasureWorkbookSnapshot current = gateway.Capture(target);
            var undo = new MeasureUndoState(contribution);
            try
            {
                DemandUndoAfterState(undo, current);
            }
            catch (PivotModelMeasureUndoUnavailableException)
            {
                DemandUndoIntermediateState(undo, current);
            }

            return CreatePreparedUndo(target, current, undo);
        }

        private PivotModelMeasureOwnershipSession JournalUndo(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotPlusPendingSemanticApplyMetadata pending)
        {
            try
            {
                return ownership.Journal(workbook, setupId, target, pending);
            }
            catch (Exception journalFailure)
            {
                throw new PivotModelMeasureMutationException(
                    "The Undo ownership journal could not be confirmed; retry Undo without changing the model.",
                    rollbackCompleted: true,
                    recoveryRequired: true,
                    journalFailure);
            }
        }

        private void CommitUndo(
            object workbook,
            string setupId,
            PivotModelMeasureOwnershipSession session,
            MeasureUndoState undo)
        {
            try
            {
                ownership.Commit(
                    workbook,
                    session,
                    undo.BeforeOwnedArtifacts,
                    undo: null);
                ForgetUndo(workbook, setupId);
            }
            catch (Exception commitFailure)
            {
                RememberUndo(workbook, setupId, undo);
                throw new PivotModelMeasureMutationException(
                    "Undo reached the verified prior state, but ownership finalization is ambiguous. Retry Undo to finalize it.",
                    rollbackCompleted: false,
                    recoveryRequired: true,
                    commitFailure);
            }
        }

        private PivotModelMeasurePreparedMutation CreatePreparedMutation(
            object workbook,
            string setupId,
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            MeasureApplyPlan plan,
            IReadOnlyList<DesiredModelMeasure> definitions,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            PivotMeasurePlacementPlan placement,
            IReadOnlyList<PivotPlusOwnedArtifact> baseArtifacts,
            bool allowNewUndoSeed)
        {
            var upsertSteps = new List<PivotMutationStep>();
            var placementSteps = new List<PivotMutationStep>();
            var deleteSteps = new List<PivotMutationStep>();
            var upsertReceipts = new Dictionary<string, LiveModelMeasureSnapshot>(
                StringComparer.OrdinalIgnoreCase);
            var createFailures = new Dictionary<string, Exception>(
                StringComparer.OrdinalIgnoreCase);

            foreach (MeasureUpsert upsert in plan.Creates)
            {
                MeasureUpsert captured = upsert;
                if (captured.Before == null)
                {
                    upsertSteps.Add(new PivotMutationStep(
                        "create model measure " + captured.Definition.DefinitionId,
                        () =>
                        {
                            try
                            {
                                upsertReceipts[captured.Definition.Name] = gateway.CreateMeasure(
                                    target,
                                    captured.Definition);
                            }
                            catch (Exception exception)
                            {
                                createFailures[captured.Definition.Name] = exception;
                                throw;
                            }
                        },
                        () =>
                        {
                            if (upsertReceipts.TryGetValue(
                                    captured.Definition.Name,
                                    out LiveModelMeasureSnapshot? value))
                            {
                                gateway.DeleteMeasure(target, value);
                                return;
                            }

                            if (createFailures.TryGetValue(
                                    captured.Definition.Name,
                                    out Exception? failure))
                            {
                                throw new InvalidOperationException(
                                    "Measure creation failed without an exact host receipt; cleanup cannot be claimed.",
                                    failure);
                            }
                        }));
                }
                else
                {
                    // Exact replay of a pending Create whose host Add already
                    // committed. It remains a logical create relative to the
                    // active metadata, so rollback must remove it rather than
                    // adopt it as a user's prior measure.
                    upsertSteps.Add(new PivotMutationStep(
                        "repair pending model measure " + captured.Definition.DefinitionId,
                        () =>
                        {
                            try
                            {
                                upsertReceipts[captured.Definition.Name] = gateway.UpdateMeasure(
                                    target,
                                    captured.Before,
                                    captured.Definition);
                            }
                            catch (Exception exception)
                            {
                                createFailures[captured.Definition.Name] = exception;
                                throw;
                            }
                        },
                        () =>
                        {
                            if (upsertReceipts.TryGetValue(
                                    captured.Definition.Name,
                                    out LiveModelMeasureSnapshot? value))
                            {
                                gateway.DeleteMeasure(target, value);
                                return;
                            }

                            if (createFailures.TryGetValue(
                                    captured.Definition.Name,
                                    out Exception? failure))
                            {
                                throw new InvalidOperationException(
                                    "Pending measure repair failed without an exact host receipt; cleanup cannot be claimed.",
                                    failure);
                            }
                        }));
                }
            }

            foreach (MeasureUpsert upsert in plan.Updates)
            {
                MeasureUpsert captured = upsert;
                if (captured.Before == null)
                {
                    throw new InvalidOperationException(
                        "An owned measure update is missing its exact prior snapshot.");
                }

                upsertSteps.Add(new PivotMutationStep(
                    "update model measure " + captured.Definition.DefinitionId,
                    () => upsertReceipts[captured.Definition.Name] = gateway.UpdateMeasure(
                        target,
                        captured.Before,
                        captured.Definition),
                    () => gateway.RestoreMeasure(target, captured.Before)));
            }

            if (plan.PlacementNeedsPreviewRepair)
            {
                placementSteps.Add(new PivotMutationStep(
                    "repair exact preview Values sequence",
                    () => gateway.RestorePlacement(target, plan.PlacementPreview.SelectedPivot),
                    () => gateway.RestorePlacement(target, before.SelectedPivot)));
            }

            if (!plan.PlacementAlreadyFinal)
            {
                placementSteps.Add(new PivotMutationStep(
                    "apply exact Values sequence",
                    () => gateway.ApplyPlacement(
                        target,
                        placement,
                        definitionsById,
                        plan.PlacementPreview),
                    () => gateway.RestorePlacement(
                        target,
                        plan.PlacementPreview.SelectedPivot)));
            }

            foreach (LiveModelMeasureSnapshot deletion in plan.Deletes)
            {
                LiveModelMeasureSnapshot captured = deletion;
                deleteSteps.Add(new PivotMutationStep(
                    "delete model measure " + captured.Name,
                    () => gateway.DeleteMeasure(target, captured),
                    () => gateway.RestoreMeasure(target, captured)));
            }

            return new PivotModelMeasurePreparedMutation(
                target,
                before,
                plan.Pending,
                new ReadOnlyCollection<PivotMutationStep>(upsertSteps),
                new ReadOnlyCollection<PivotMutationStep>(placementSteps),
                new ReadOnlyCollection<PivotMutationStep>(deleteSteps),
                plan.IsNoChange,
                plan.Creates.Count,
                plan.Updates.Count,
                plan.DeleteCount,
                () => gateway.Refresh(target),
                () =>
                {
                    ModelMeasureWorkbookSnapshot after = gateway.Capture(target);
                    VerifyFinal(
                        before,
                        after,
                        plan,
                        definitionsById,
                        placement,
                        upsertReceipts);
                    return after;
                },
                () => DemandExactWorkbookSnapshot(
                    before,
                    gateway.Capture(target),
                    "combined measure Apply rollback"),
                verified => BuildFinalArtifacts(definitions, verified),
                (combinedApplyId, combinedPlanFingerprint) =>
                {
                    if (!allowNewUndoSeed)
                    {
                        // A resumed durable Apply may expose only its current
                        // host state. Never relabel that state as the original
                        // session-only Undo baseline after a restart.
                        return;
                    }

                    RememberPendingApplyUndoSeed(
                        workbook,
                        setupId,
                        new PendingApplyUndoSeed(
                            combinedApplyId,
                            combinedPlanFingerprint,
                            target.Identity,
                            before,
                            baseArtifacts
                                .Where(artifact => artifact.Kind == PivotPlusArtifactKind.Measure)
                                .Select(CloneArtifact)
                                .ToList()));
                },
                (combinedApplyId, combinedPlanFingerprint, verified) =>
                {
                    if (!TryGetPendingApplyUndoSeed(
                            workbook,
                            setupId,
                            combinedApplyId,
                            combinedPlanFingerprint,
                            target.Identity,
                            out PendingApplyUndoSeed? seed) ||
                        seed == null)
                    {
                        // Exact prior DAX/layout is intentionally session-only.
                        // A restart can still converge the durable Apply, but
                        // cannot manufacture a trustworthy Undo snapshot.
                        return null;
                    }

                    IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts =
                        BuildFinalArtifacts(definitions, verified);
                    return new PivotModelMeasureUndoContribution(
                        combinedApplyId,
                        combinedPlanFingerprint,
                        target.Identity,
                        seed.Before,
                        verified,
                        seed.BeforeOwnedArtifacts.Select(CloneArtifact).ToList(),
                        finalArtifacts.Select(CloneArtifact).ToList(),
                        BuildUndoMetadata(
                            combinedApplyId,
                            seed.Before,
                            verified,
                            plan,
                            finalArtifacts));
                });
        }

        private void ExecuteUndo(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot current,
            MeasureUndoState undo)
        {
            PivotModelMeasurePreparedUndo prepared = CreatePreparedUndo(target, current, undo);
            coordinator.Execute(
                target.PivotTable,
                prepared.Steps,
                prepared.Refresh,
                prepared.Verify);
        }

        private PivotModelMeasurePreparedUndo CreatePreparedUndo(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot current,
            MeasureUndoState undo)
        {
            var upsertSteps = new List<PivotMutationStep>();
            var placementSteps = new List<PivotMutationStep>();
            var deleteSteps = new List<PivotMutationStep>();
            var currentByName = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var priorByName = undo.Before.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var recreated = new Dictionary<string, LiveModelMeasureSnapshot>(
                StringComparer.OrdinalIgnoreCase);

            IReadOnlyList<LiveModelMeasureSnapshot> restoreCandidates =
                OrderRestores(undo.Before.Measures.Where(prior =>
                    undo.BeforeOwnedArtifacts.Any(artifact => string.Equals(
                        artifact.ArtifactId,
                        prior.Name,
                        StringComparison.OrdinalIgnoreCase)) &&
                    (!currentByName.TryGetValue(
                         prior.Name,
                         out LiveModelMeasureSnapshot? currentMeasure) ||
                     !string.Equals(
                         currentMeasure.LiveFingerprint,
                         prior.LiveFingerprint,
                         StringComparison.Ordinal))).ToList());
            foreach (LiveModelMeasureSnapshot prior in restoreCandidates)
            {
                LiveModelMeasureSnapshot capturedPrior = prior;
                currentByName.TryGetValue(prior.Name, out LiveModelMeasureSnapshot? capturedCurrent);
                upsertSteps.Add(new PivotMutationStep(
                    "restore prior model measure " + prior.Name,
                    () => recreated[prior.Name] = gateway.RestoreMeasure(target, capturedPrior),
                    () =>
                    {
                        if (capturedCurrent != null)
                        {
                            gateway.RestoreMeasure(target, capturedCurrent);
                        }
                        else if (recreated.TryGetValue(prior.Name, out LiveModelMeasureSnapshot? value))
                        {
                            gateway.DeleteMeasure(target, value);
                        }
                    }));
            }

            placementSteps.Add(new PivotMutationStep(
                "restore prior Values sequence",
                () => gateway.RestorePlacement(target, undo.Before.SelectedPivot),
                () => gateway.RestorePlacement(target, current.SelectedPivot)));

            IReadOnlyList<LiveModelMeasureSnapshot> createdMeasures = OrderDeletes(
                current.Measures.Where(measure =>
                    undo.AfterOwnedArtifacts.Any(artifact => string.Equals(
                        artifact.ArtifactId,
                        measure.Name,
                        StringComparison.OrdinalIgnoreCase)) &&
                    !priorByName.ContainsKey(measure.Name)).ToList());
            foreach (LiveModelMeasureSnapshot measure in createdMeasures)
            {
                LiveModelMeasureSnapshot captured = measure;
                deleteSteps.Add(new PivotMutationStep(
                    "remove created model measure " + measure.Name,
                    () => gateway.DeleteMeasure(target, captured),
                    () => gateway.RestoreMeasure(target, captured)));
            }

            return new PivotModelMeasurePreparedUndo(
                target,
                current,
                new ReadOnlyCollection<PivotMutationStep>(upsertSteps),
                new ReadOnlyCollection<PivotMutationStep>(placementSteps),
                new ReadOnlyCollection<PivotMutationStep>(deleteSteps),
                () => gateway.Refresh(target),
                () =>
                {
                    ModelMeasureWorkbookSnapshot restored = gateway.Capture(target);
                    DemandUndoFinal(undo, undo.UndoStart, restored);
                },
                () => DemandExactWorkbookSnapshot(
                    current,
                    gateway.Capture(target),
                    "combined measure Undo rollback"));
        }

        private static MeasureApplyPlan BuildPlan(
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            ModelMeasureWorkbookSnapshot before,
            IReadOnlyList<DesiredModelMeasure> definitions,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            PivotMeasurePlacementPlan placement,
            bool compositeParticipant = false,
            ModelMeasureWorkbookSnapshot? pendingPreview = null)
        {
            ModelMeasureWorkbookSnapshot placementPreview = pendingPreview ?? before;
            IReadOnlyList<PivotPlusOwnedArtifact> active = baseMetadata.Artifacts
                .Where(artifact => artifact.Kind == PivotPlusArtifactKind.Measure)
                .ToList();
            DemandUniqueArtifacts(active);
            var activeByName = active.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var liveByName = before.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            DemandLiveOwnership(active, existingPending, liveByName, definitionsById);
            IEnumerable<string> placementOwnedNames = activeByName.Keys.Concat(
                existingPending == null
                    ? Enumerable.Empty<string>()
                    : existingPending.Transitions
                        .Where(transition =>
                            transition.Kind == PivotPlusArtifactKind.Measure &&
                            transition.Operation == PivotPlusSemanticArtifactOperation.Create)
                        .Select(transition => transition.ArtifactId));
            bool placementAlreadyFinal = false;
            bool placementNeedsPreviewRepair = false;
            string observedExpectedPivot = string.Empty;
            if (existingPending != null &&
                PivotModelMeasureCanonical.TryCreateObservedExpectedPivotFingerprint(
                    placement,
                    definitionsById,
                    before.SelectedPivot,
                    out observedExpectedPivot))
            {
                placementAlreadyFinal = compositeParticipant || string.Equals(
                    observedExpectedPivot,
                    existingPending.ExpectedPivotFingerprint,
                    StringComparison.Ordinal);
            }

            bool hasSessionPreview = existingPending == null || pendingPreview != null;
            if (existingPending != null && pendingPreview == null &&
                placement.Values
                    .Where(value => value.IsGeneratedMeasure)
                    .Select(value => definitionsById[value.DefinitionId!].Name)
                    .Any(name => activeByName.ContainsKey(name) &&
                                 before.SelectedPivot.DataFields.Any(field =>
                                     string.Equals(
                                         field.ModelMeasureName,
                                         name,
                                         StringComparison.OrdinalIgnoreCase))))
            {
                throw new InvalidOperationException(
                    "The pending Apply includes a previously placed owned measure, but its exact session preview display state is unavailable.");
            }

            DemandPlacementPlan(
                placement,
                definitionsById,
                placementPreview.SelectedPivot,
                placementOwnedNames,
                placementAlreadyFinal);
            if (existingPending != null && !placementAlreadyFinal &&
                !string.Equals(
                    before.SelectedPivotFingerprint,
                    existingPending.BeforePivotFingerprint,
                    StringComparison.Ordinal))
            {
                if (pendingPreview == null ||
                    !string.Equals(
                        placementPreview.SelectedPivotFingerprint,
                        existingPending.BeforePivotFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The selected PivotTable is in a partial pending Values state, but its exact session preview is unavailable.");
                }

                DemandPendingApplyIntermediateState(
                    placementPreview,
                    before,
                    existingPending,
                    definitionsById,
                    placementOwnedNames);
                placementNeedsPreviewRepair = true;
            }

            var creates = new List<MeasureUpsert>();
            var updates = new List<MeasureUpsert>();
            foreach (DesiredModelMeasure definition in definitions
                         .OrderBy(item => item.CreationOrder))
            {
                if (!activeByName.TryGetValue(definition.Name, out PivotPlusOwnedArtifact? owned))
                {
                    if (liveByName.TryGetValue(definition.Name, out LiveModelMeasureSnapshot? collision))
                    {
                        if (existingPending == null ||
                            !string.Equals(collision.Description, definition.DescriptionMarker, StringComparison.Ordinal))
                        {
                            throw new InvalidOperationException(
                                "An unowned model measure already uses generated name '" + definition.Name + "'.");
                        }

                        creates.Add(new MeasureUpsert(definition, collision));
                    }
                    else
                    {
                        creates.Add(new MeasureUpsert(definition, before: null));
                    }

                    continue;
                }

                LiveModelMeasureSnapshot live = liveByName[definition.Name];
                if (!string.Equals(live.LiveFingerprint, owned.Fingerprint, StringComparison.Ordinal) &&
                    existingPending == null)
                {
                    throw new InvalidOperationException(
                        "An owned model measure changed outside PivotTable+.");
                }

                if (!string.Equals(
                        live.Description,
                        definition.DescriptionMarker,
                        StringComparison.Ordinal) ||
                    !string.Equals(live.LiveFingerprint, owned.Fingerprint, StringComparison.Ordinal))
                {
                    updates.Add(new MeasureUpsert(definition, live));
                }
            }

            HashSet<string> desiredNames = new HashSet<string>(
                definitions.Select(definition => definition.Name),
                StringComparer.OrdinalIgnoreCase);
            List<PivotPlusOwnedArtifact> deletionArtifacts = active
                .Where(artifact => !desiredNames.Contains(artifact.ArtifactId))
                .ToList();
            var deletes = deletionArtifacts
                .Select(artifact => liveByName.TryGetValue(
                    artifact.ArtifactId,
                    out LiveModelMeasureSnapshot? live)
                    ? live
                    : null)
                .Where(live => live != null)
                .Cast<LiveModelMeasureSnapshot>()
                .ToList();

            DemandNoUnsafeDependenciesAndSharedUse(
                before,
                activeByName.Keys,
                creates.Concat(updates).ToList(),
                deletes,
                deletionArtifacts.Select(artifact => artifact.ArtifactId).ToList(),
                desiredNames);
            deletes = OrderDeletes(deletes);

            var transitions = new List<PivotPlusSemanticArtifactTransition>();
            transitions.AddRange(creates.Select(item => new PivotPlusSemanticArtifactTransition
            {
                Kind = PivotPlusArtifactKind.Measure,
                ArtifactId = item.Definition.Name,
                Operation = PivotPlusSemanticArtifactOperation.Create,
                BeforeLiveFingerprint = string.Empty,
                PlannedDefinitionFingerprint = item.Definition.DefinitionFingerprint
            }));
            transitions.AddRange(updates.Select(item => new PivotPlusSemanticArtifactTransition
            {
                Kind = PivotPlusArtifactKind.Measure,
                ArtifactId = item.Definition.Name,
                Operation = PivotPlusSemanticArtifactOperation.Update,
                BeforeLiveFingerprint = activeByName[item.Definition.Name].Fingerprint,
                PlannedDefinitionFingerprint = item.Definition.DefinitionFingerprint
            }));
            transitions.AddRange(deletionArtifacts.Select(item => new PivotPlusSemanticArtifactTransition
            {
                Kind = PivotPlusArtifactKind.Measure,
                ArtifactId = item.ArtifactId,
                Operation = PivotPlusSemanticArtifactOperation.Delete,
                BeforeLiveFingerprint = item.Fingerprint,
                PlannedDefinitionFingerprint = PivotModelMeasureCanonical.CreateDeleteDefinitionFingerprint(
                    item.ArtifactId,
                    item.Fingerprint)
            }));

            string planFingerprint = PivotModelMeasureCanonical.CreatePlanFingerprint(
                definitions,
                placement);
            string expectedPivot = placementAlreadyFinal
                ? (compositeParticipant
                    ? observedExpectedPivot
                    : existingPending!.ExpectedPivotFingerprint)
                : PivotModelMeasureCanonical.CreateExpectedPivotFingerprint(
                    placement,
                    definitionsById,
                    placementPreview.SelectedPivot);
            var pending = new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = existingPending == null
                    ? "apply_" + Guid.NewGuid().ToString("N")
                    : existingPending.ApplyId,
                PlanFingerprint = planFingerprint,
                BeforePivotFingerprint = existingPending == null
                    ? before.SelectedPivotFingerprint
                    : existingPending.BeforePivotFingerprint,
                ExpectedPivotFingerprint = expectedPivot,
                Transitions = transitions
            };
            if (existingPending != null)
            {
                if (compositeParticipant)
                {
                    DemandSameMeasurePendingSlice(existingPending, pending);
                }
                else
                {
                    DemandSamePendingPlan(existingPending, pending);
                }
            }

            bool placementMatches = placementAlreadyFinal || PlacementMatches(
                placement,
                definitionsById,
                placementPreview.SelectedPivot,
                before.SelectedPivot);
            return new MeasureApplyPlan(
                pending,
                creates,
                updates,
                deletes,
                deletionArtifacts.Select(artifact => artifact.ArtifactId).ToList(),
                deletionArtifacts.Count,
                isNoChange: transitions.Count == 0 && placementMatches,
                placementAlreadyFinal,
                placementNeedsPreviewRepair,
                placementPreview,
                hasSessionPreview);
        }

        private static IReadOnlyList<DesiredModelMeasure> CompileDefinitions(
            string setupId,
            PivotDaxCompilation compilation)
        {
            if (compilation.Measures == null || compilation.CreationSequence == null ||
                compilation.Measures.Count > MaximumMeasures ||
                compilation.CreationSequence.Count != compilation.Measures.Count)
            {
                throw new ArgumentException(
                    "The compiled measure set is incomplete or exceeds the measure limit.",
                    nameof(compilation));
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var creationOrders = new HashSet<int>();
            var displayOrders = new HashSet<int>();
            var result = new List<DesiredModelMeasure>();
            foreach (OwnedPivotMeasureDefinition measure in compilation.Measures)
            {
                if (measure == null)
                {
                    throw new ArgumentException("A compiled measure cannot be null.", nameof(compilation));
                }

                PivotPlusMetadataValidator.ValidateId(
                    measure.DefinitionId,
                    "measure definition identifier");
                PivotPlusMetadataValidator.ValidateArtifactName(measure.GeneratedMeasureName);
                PivotPlusMetadataValidator.ValidateFingerprint(
                    measure.DefinitionFingerprint,
                    "measure definition fingerprint");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    measure.FormulaFingerprint,
                    "measure formula fingerprint");
                if (!ids.Add(measure.DefinitionId) || !names.Add(measure.GeneratedMeasureName) ||
                    !creationOrders.Add(measure.CreationOrder) ||
                    !displayOrders.Add(measure.DisplayOrder) ||
                    measure.CreationOrder < 1 || measure.CreationOrder > compilation.Measures.Count ||
                    measure.DisplayOrder < 1 || measure.DisplayOrder > compilation.Measures.Count ||
                    string.IsNullOrWhiteSpace(measure.HomeTableName) ||
                    string.IsNullOrWhiteSpace(measure.DaxFormula))
                {
                    throw new ArgumentException(
                        "The compiled measure identity, order, table, or formula is invalid.",
                        nameof(compilation));
                }

                string marker = PivotModelMeasureCanonical.CreateDescriptionMarker(
                    setupId,
                    measure.DefinitionId,
                    measure.DefinitionFingerprint);
                result.Add(new DesiredModelMeasure(
                    measure.DefinitionId,
                    measure.DisplayOrder,
                    measure.CreationOrder,
                    measure.HomeTableName,
                    measure.GeneratedMeasureName,
                    measure.DaxFormula,
                    measure.Format,
                    measure.DirectDependencyDefinitionIds,
                    measure.DefinitionFingerprint,
                    marker));
            }

            foreach (DesiredModelMeasure definition in result)
            {
                if (definition.DirectDependencyDefinitionIds.Any(id => !ids.Contains(id)))
                {
                    throw new ArgumentException(
                        "A compiled measure references an unknown dependency.",
                        nameof(compilation));
                }
            }

            string[] expectedCreationSequence = result
                .OrderBy(definition => definition.CreationOrder)
                .Select(definition => definition.DefinitionId)
                .ToArray();
            string[] actualCreationSequence = compilation.CreationSequence
                .Select(definition => definition.DefinitionId)
                .ToArray();
            if (!expectedCreationSequence.SequenceEqual(
                    actualCreationSequence,
                    StringComparer.Ordinal))
            {
                throw new ArgumentException(
                    "The compiler creation sequence does not match the declared dependency order.",
                    nameof(compilation));
            }

            return new ReadOnlyCollection<DesiredModelMeasure>(result);
        }

        private static void DemandPlacementPlan(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelPivotUsageSnapshot current,
            IEnumerable<string> ownedMeasureNames,
            bool placementAlreadyFinal)
        {
            if (placement.Values == null || placement.Values.Count > MaximumValuePlacements ||
                placement.Values.Any(value => value == null) ||
                !Enum.IsDefined(typeof(PivotValuesAxis), placement.ValuesAxis) ||
                placement.ValuesPosition < 1)
            {
                throw new ArgumentException("The measure placement plan is invalid.", nameof(placement));
            }

            int[] positions = placement.Values.Select(value => value.Position).OrderBy(value => value).ToArray();
            if (!positions.SequenceEqual(Enumerable.Range(1, positions.Length)))
            {
                throw new ArgumentException(
                    "The final Values sequence must use contiguous one-based positions.",
                    nameof(placement));
            }

            if (placement.Values.Count <= 1 &&
                (placement.ValuesAxis != PivotValuesAxis.Automatic ||
                 placement.ValuesPosition != 1))
            {
                throw new ArgumentException(
                    "Zero or one Value requires the automatic Values axis sentinel position.",
                    nameof(placement));
            }

            if (placement.Values.Count >= 2 &&
                (placement.ValuesAxis != PivotValuesAxis.Rows &&
                 placement.ValuesAxis != PivotValuesAxis.Columns))
            {
                throw new ArgumentException(
                    "Two or more Values require an explicit row or column Values axis.",
                    nameof(placement));
            }

            var generated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var owned = new HashSet<string>(ownedMeasureNames, StringComparer.OrdinalIgnoreCase);
            foreach (PivotMeasureValuePlacement value in placement.Values)
            {
                if (value.IsGeneratedMeasure)
                {
                    string? definitionId = value.DefinitionId;
                    if (string.IsNullOrWhiteSpace(definitionId) ||
                        !definitionsById.ContainsKey(definitionId!) ||
                        !generated.Add(definitionId!))
                    {
                        throw new ArgumentException(
                            "A generated Values entry is missing, unknown, or duplicated.",
                            nameof(placement));
                    }

                    continue;
                }

                PivotExistingDataFieldIdentity identity = value.ExistingDataField!;
                if (string.IsNullOrWhiteSpace(identity.UniqueName) ||
                    identity.UniqueName.Length > 255 ||
                    identity.UniqueName.Any(char.IsControl))
                {
                    throw new ArgumentException(
                        "An existing Values entry has an invalid unique field identity.",
                        nameof(placement));
                }
                PivotPlusMetadataValidator.ValidateFingerprint(
                    identity.CurrentCaptionFingerprint,
                    "current value caption fingerprint");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    identity.CurrentNumberFormatFingerprint,
                    "current value number-format fingerprint");
                if (identity.CurrentPosition < 1)
                {
                    throw new ArgumentException(
                        "An existing Values entry has an invalid preview position.",
                        nameof(placement));
                }
                string key = ExistingKey(
                    identity.UniqueName,
                    identity.CurrentCaptionFingerprint,
                    identity.CurrentNumberFormatFingerprint,
                    identity.CurrentPosition);
                if (!existing.Add(key))
                {
                    throw new ArgumentException(
                        "An existing Values entry is duplicated.",
                        nameof(placement));
                }
            }

            if (generated.Count != definitionsById.Count)
            {
                throw new ArgumentException(
                    "Every generated measure must appear exactly once in the final Values sequence.",
                    nameof(placement));
            }

            var requiredExisting = new HashSet<string>(
                current.DataFields
                    .Where(field => string.IsNullOrWhiteSpace(field.ModelMeasureName) ||
                                    !owned.Contains(field.ModelMeasureName!))
                    .Select(field => ExistingKey(
                        field.UniqueName,
                        field.CaptionFingerprint,
                        PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                            field.NumberFormat),
                        field.Position)),
                StringComparer.OrdinalIgnoreCase);
            if (!placementAlreadyFinal && !requiredExisting.SetEquals(existing))
            {
                throw new ArgumentException(
                    "Every current unowned Values field must appear exactly once with its preview caption fingerprint.",
                    nameof(placement));
            }
        }

        private static void DemandPendingApplyIntermediateState(
            ModelMeasureWorkbookSnapshot preview,
            ModelMeasureWorkbookSnapshot current,
            PivotPlusPendingSemanticApplyMetadata pending,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            IEnumerable<string> ownedMeasureNames)
        {
            var involvedNames = new HashSet<string>(
                pending.Transitions
                    .Where(transition => transition.Kind == PivotPlusArtifactKind.Measure)
                    .Select(transition => transition.ArtifactId)
                    .Concat(definitionsById.Values.Select(definition => definition.Name)),
                StringComparer.OrdinalIgnoreCase);
            var previewMeasures = preview.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var currentMeasures = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);

            foreach (LiveModelMeasureSnapshot prior in preview.Measures.Where(
                         measure => !involvedNames.Contains(measure.Name)))
            {
                if (!currentMeasures.TryGetValue(
                        prior.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        prior.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "An unrelated model measure changed while the pending Apply was incomplete.");
                }
            }

            foreach (LiveModelMeasureSnapshot live in current.Measures)
            {
                if (involvedNames.Contains(live.Name)) continue;
                if (!previewMeasures.TryGetValue(
                        live.Name,
                        out LiveModelMeasureSnapshot? prior) ||
                    !string.Equals(
                        prior.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A model measure appeared or changed outside the pending Apply.");
                }
            }

            DemandExactOtherPivotUsages(
                preview,
                current,
                "A different Data Model PivotTable changed while the pending Apply was incomplete.");

            var mutableNames = new HashSet<string>(
                involvedNames.Concat(ownedMeasureNames),
                StringComparer.OrdinalIgnoreCase);
            Dictionary<string, int> requiredUnowned = CreateDataFieldOccurrenceCounts(
                preview.SelectedPivot.DataFields.Where(field =>
                    string.IsNullOrWhiteSpace(field.ModelMeasureName) ||
                    !mutableNames.Contains(field.ModelMeasureName!)));
            Dictionary<string, int> currentUnowned = CreateDataFieldOccurrenceCounts(
                current.SelectedPivot.DataFields.Where(field =>
                    string.IsNullOrWhiteSpace(field.ModelMeasureName) ||
                    !mutableNames.Contains(field.ModelMeasureName!)));
            if (!OccurrenceCountsEqual(requiredUnowned, currentUnowned))
            {
                throw new InvalidOperationException(
                    "An unowned Values field changed while the pending Apply was incomplete.");
            }
        }

        private static Dictionary<string, int> CreateDataFieldOccurrenceCounts(
            IEnumerable<ModelDataFieldSnapshot> fields)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (ModelDataFieldSnapshot field in fields)
            {
                string key = DataFieldOccurrenceKey(field);
                if (result.ContainsKey(key))
                {
                    result[key]++;
                }
                else
                {
                    result.Add(key, 1);
                }
            }

            return result;
        }

        private static bool OccurrenceCountsEqual(
            IReadOnlyDictionary<string, int> first,
            IReadOnlyDictionary<string, int> second)
        {
            return first.Count == second.Count && first.All(pair =>
                second.TryGetValue(pair.Key, out int count) && count == pair.Value);
        }

        private static string DataFieldOccurrenceKey(ModelDataFieldSnapshot field)
        {
            return field.UniqueName + "\u001f" + field.CaptionFingerprint + "\u001f" +
                   PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                       field.NumberFormat) + "\u001f" +
                   (field.ModelMeasureName ?? string.Empty);
        }

        private static void DemandExactOtherPivotUsages(
            ModelMeasureWorkbookSnapshot expected,
            ModelMeasureWorkbookSnapshot actual,
            string message)
        {
            IReadOnlyList<ModelPivotUsageSnapshot> expectedOthers = expected.PivotUsages
                .Where(usage => !usage.IsSelectedTarget)
                .ToList();
            IReadOnlyList<ModelPivotUsageSnapshot> actualOthers = actual.PivotUsages
                .Where(usage => !usage.IsSelectedTarget)
                .ToList();
            if (expectedOthers.Count != actualOthers.Count)
            {
                throw new InvalidOperationException(message);
            }

            foreach (ModelPivotUsageSnapshot prior in expectedOthers)
            {
                ModelPivotUsageSnapshot? live = actualOthers.SingleOrDefault(usage =>
                    string.Equals(
                        usage.WorksheetName,
                        prior.WorksheetName,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        usage.PivotTableName,
                        prior.PivotTableName,
                        StringComparison.Ordinal));
                if (live == null || !string.Equals(
                        PivotModelMeasureCanonical.CreatePivotFingerprint(prior),
                        PivotModelMeasureCanonical.CreatePivotFingerprint(live),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(message);
                }
            }
        }

        private static void DemandLiveOwnership(
            IReadOnlyList<PivotPlusOwnedArtifact> active,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            IReadOnlyDictionary<string, LiveModelMeasureSnapshot> liveByName,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById)
        {
            var desiredByName = definitionsById.Values.ToDictionary(
                definition => definition.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact owned in active)
            {
                if (!liveByName.TryGetValue(owned.ArtifactId, out LiveModelMeasureSnapshot? live))
                {
                    bool pendingDelete = existingPending != null &&
                        existingPending.Transitions.Any(transition =>
                            transition.Kind == PivotPlusArtifactKind.Measure &&
                            transition.Operation == PivotPlusSemanticArtifactOperation.Delete &&
                            string.Equals(
                                transition.ArtifactId,
                                owned.ArtifactId,
                                StringComparison.OrdinalIgnoreCase));
                    if (!pendingDelete)
                    {
                        throw new InvalidOperationException(
                            "An owned model measure is missing from the workbook model.");
                    }

                    continue;
                }

                if (string.Equals(live.LiveFingerprint, owned.Fingerprint, StringComparison.Ordinal))
                {
                    continue;
                }

                bool repairable = existingPending != null &&
                    desiredByName.TryGetValue(owned.ArtifactId, out DesiredModelMeasure? desired) &&
                    string.Equals(live.Description, desired.DescriptionMarker, StringComparison.Ordinal);
                if (!repairable)
                {
                    throw new InvalidOperationException(
                        "An owned model measure no longer matches its exact ownership fingerprint.");
                }
            }
        }

        private static void DemandNoUnsafeDependenciesAndSharedUse(
            ModelMeasureWorkbookSnapshot snapshot,
            IEnumerable<string> activeOwnedNames,
            IReadOnlyList<MeasureUpsert> upserts,
            IReadOnlyList<LiveModelMeasureSnapshot> deletes,
            IReadOnlyList<string> deletionNames,
            ISet<string> desiredNames)
        {
            var owned = new HashSet<string>(activeOwnedNames, StringComparer.OrdinalIgnoreCase);
            var changed = new HashSet<string>(
                upserts.Select(item => item.Definition.Name)
                    .Concat(deletionNames),
                StringComparer.OrdinalIgnoreCase);
            var affected = new HashSet<string>(changed, StringComparer.OrdinalIgnoreCase);
            bool expanded;
            do
            {
                expanded = false;
                foreach (LiveModelMeasureSnapshot measure in snapshot.Measures)
                {
                    if (affected.Contains(measure.Name)) continue;
                    IReadOnlyCollection<string> references =
                        DaxMeasureReferenceScanner.ReadPossibleReferences(measure.Formula);
                    if (references.Any(affected.Contains))
                    {
                        if (!owned.Contains(measure.Name) || !desiredNames.Contains(measure.Name))
                        {
                            throw new InvalidOperationException(
                                "A user or out-of-scope model measure may depend on the requested measure change.");
                        }

                        affected.Add(measure.Name);
                        expanded = true;
                    }
                }
            }
            while (expanded);

            foreach (ModelPivotUsageSnapshot usage in snapshot.PivotUsages.Where(item => !item.IsSelectedTarget))
            {
                if (usage.DataFields.Any(field =>
                        !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                        affected.Contains(field.ModelMeasureName!)))
                {
                    throw new InvalidOperationException(
                        "A requested measure change would alter another PivotTable in the workbook.");
                }
            }

            var deletionNameSet = new HashSet<string>(
                deletionNames,
                StringComparer.OrdinalIgnoreCase);
            var upsertsByName = upserts.ToDictionary(
                item => item.Definition.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (string deletionName in deletionNameSet)
            {
                foreach (LiveModelMeasureSnapshot dependent in snapshot.Measures)
                {
                    if (string.Equals(deletionName, dependent.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    bool currentlyReferences = DaxMeasureReferenceScanner
                        .ReadPossibleReferences(dependent.Formula)
                        .Contains(deletionName, StringComparer.OrdinalIgnoreCase);
                    if (!currentlyReferences || deletionNameSet.Contains(dependent.Name))
                    {
                        continue;
                    }

                    // An owned dependent may be updated first to remove the
                    // reference. This is the only supported non-delete escape;
                    // user measures and unchanged owned measures remain a hard
                    // block.
                    bool updateRemovesReference = upsertsByName.TryGetValue(
                            dependent.Name,
                            out MeasureUpsert? dependentUpdate) &&
                        !DaxMeasureReferenceScanner
                            .ReadPossibleReferences(dependentUpdate.Definition.Formula)
                            .Contains(deletionName, StringComparer.OrdinalIgnoreCase);
                    if (!updateRemovesReference)
                    {
                        throw new InvalidOperationException(
                            "A model measure cannot be deleted while another measure still references it.");
                    }
                }
            }
        }

        private static List<LiveModelMeasureSnapshot> OrderDeletes(
            IReadOnlyList<LiveModelMeasureSnapshot> deletes)
        {
            var byName = deletes.ToDictionary(item => item.Name, StringComparer.OrdinalIgnoreCase);
            var result = new List<LiveModelMeasureSnapshot>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot measure in deletes)
            {
                VisitDelete(measure, byName, visited, result);
            }

            result.Reverse();
            return result;
        }

        private static List<LiveModelMeasureSnapshot> OrderRestores(
            IReadOnlyList<LiveModelMeasureSnapshot> restores)
        {
            var byName = restores.ToDictionary(item => item.Name, StringComparer.OrdinalIgnoreCase);
            var result = new List<LiveModelMeasureSnapshot>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot measure in restores)
            {
                VisitRestore(measure, byName, visited, result);
            }

            return result;
        }

        private static void VisitRestore(
            LiveModelMeasureSnapshot measure,
            IReadOnlyDictionary<string, LiveModelMeasureSnapshot> byName,
            ISet<string> visited,
            ICollection<LiveModelMeasureSnapshot> result)
        {
            if (!visited.Add(measure.Name)) return;
            foreach (string reference in DaxMeasureReferenceScanner.ReadPossibleReferences(measure.Formula))
            {
                if (byName.TryGetValue(reference, out LiveModelMeasureSnapshot? dependency))
                {
                    VisitRestore(dependency, byName, visited, result);
                }
            }

            result.Add(measure);
        }

        private static void VisitDelete(
            LiveModelMeasureSnapshot measure,
            IReadOnlyDictionary<string, LiveModelMeasureSnapshot> byName,
            ISet<string> visited,
            ICollection<LiveModelMeasureSnapshot> result)
        {
            if (!visited.Add(measure.Name)) return;
            foreach (string reference in DaxMeasureReferenceScanner.ReadPossibleReferences(measure.Formula))
            {
                if (byName.TryGetValue(reference, out LiveModelMeasureSnapshot? dependency))
                {
                    VisitDelete(dependency, byName, visited, result);
                }
            }

            result.Add(measure);
        }

        private static void VerifyFinal(
            ModelMeasureWorkbookSnapshot before,
            ModelMeasureWorkbookSnapshot after,
            MeasureApplyPlan plan,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, LiveModelMeasureSnapshot> upsertReceipts)
        {
            var afterByName = after.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (DesiredModelMeasure definition in definitionsById.Values)
            {
                if (!afterByName.TryGetValue(definition.Name, out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(live.Description, definition.DescriptionMarker, StringComparison.Ordinal) ||
                    !string.Equals(live.AssociatedTableName, definition.HomeTableName, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Excel did not retain the exact generated measure definition.");
                }

                LiveModelMeasureSnapshot expected;
                if (upsertReceipts.TryGetValue(
                        definition.Name,
                        out LiveModelMeasureSnapshot? capturedExpected))
                {
                    expected = capturedExpected;
                }
                else
                {
                    expected = before.Measures.Single(measure => string.Equals(
                        measure.Name,
                        definition.Name,
                        StringComparison.OrdinalIgnoreCase));
                }

                if (!string.Equals(
                        live.LiveFingerprint,
                        expected.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "A generated measure changed after its exact host receipt was captured.");
                }
            }

            if (plan.DeletedNames.Any(afterByName.ContainsKey))
            {
                throw new InvalidOperationException("Excel retained a model measure scheduled for deletion.");
            }

            var expectedMeasureNames = new HashSet<string>(
                before.Measures
                    .Where(measure => !plan.DeletedNames.Contains(
                        measure.Name,
                        StringComparer.OrdinalIgnoreCase))
                    .Select(measure => measure.Name)
                    .Concat(definitionsById.Values.Select(definition => definition.Name)),
                StringComparer.OrdinalIgnoreCase);
            if (after.Measures.Count != expectedMeasureNames.Count ||
                after.Measures.Any(measure => !expectedMeasureNames.Contains(measure.Name)))
            {
                throw new InvalidOperationException(
                    "The workbook model-measure inventory changed during the PivotTable+ Apply.");
            }

            HashSet<string> changed = new HashSet<string>(
                plan.Upserts.Select(item => item.Definition.Name)
                    .Concat(plan.DeletedNames),
                StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot prior in before.Measures.Where(item => !changed.Contains(item.Name)))
            {
                if (!afterByName.TryGetValue(prior.Name, out LiveModelMeasureSnapshot? current) ||
                    !string.Equals(prior.LiveFingerprint, current.LiveFingerprint, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "An unrelated model measure changed during the PivotTable+ Apply.");
                }
            }


            DemandNoPostRefreshUnsafeRelationships(
                after,
                changed,
                new HashSet<string>(
                    definitionsById.Values.Select(definition => definition.Name),
                    StringComparer.OrdinalIgnoreCase),
                "Apply");

            bool placementMatches = plan.HasSessionPreview
                ? PlacementMatches(
                    placement,
                    definitionsById,
                    plan.PlacementPreview.SelectedPivot,
                    after.SelectedPivot)
                : PivotModelMeasureCanonical.MatchesObservedExpectedPivotFingerprint(
                    placement,
                    definitionsById,
                    after.SelectedPivot,
                    plan.Pending.ExpectedPivotFingerprint);
            if (!placementMatches)
            {
                throw new InvalidOperationException(
                    "Excel did not apply the exact requested Values sequence.");
            }

            foreach (ModelPivotUsageSnapshot prior in before.PivotUsages.Where(item => !item.IsSelectedTarget))
            {
                ModelPivotUsageSnapshot current = after.PivotUsages.SingleOrDefault(item =>
                    string.Equals(item.WorksheetName, prior.WorksheetName, StringComparison.Ordinal) &&
                    string.Equals(item.PivotTableName, prior.PivotTableName, StringComparison.Ordinal)) ??
                    throw new InvalidOperationException("Another model PivotTable disappeared during Apply.");
                if (!string.Equals(
                        PivotModelMeasureCanonical.CreatePivotFingerprint(prior),
                        PivotModelMeasureCanonical.CreatePivotFingerprint(current),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Another PivotTable changed during the selected measure Apply.");
                }
            }


            if (after.PivotUsages.Count != before.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model PivotTable inventory changed during Apply.");
            }
        }

        private static void DemandNoPostRefreshUnsafeRelationships(
            ModelMeasureWorkbookSnapshot snapshot,
            ISet<string> changedMeasureNames,
            ISet<string> allowedDependentMeasureNames,
            string operation)
        {
            if (changedMeasureNames.Count == 0) return;
            foreach (LiveModelMeasureSnapshot measure in snapshot.Measures.Where(measure =>
                         !allowedDependentMeasureNames.Contains(measure.Name)))
            {
                if (DaxMeasureReferenceScanner.ReadPossibleReferences(measure.Formula)
                    .Any(changedMeasureNames.Contains))
                {
                    throw new InvalidOperationException(
                        operation + " introduced or retained an out-of-scope measure dependency on a changed measure.");
                }
            }

            if (snapshot.PivotUsages.Where(usage => !usage.IsSelectedTarget).Any(usage =>
                    usage.DataFields.Any(field =>
                        !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                        changedMeasureNames.Contains(field.ModelMeasureName!))))
            {
                throw new InvalidOperationException(
                    operation + " introduced or retained another PivotTable use of a changed measure.");
            }
        }

        private static bool PlacementMatches(
            PivotMeasurePlacementPlan placement,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            ModelPivotUsageSnapshot before,
            ModelPivotUsageSnapshot actual)
        {
            if (placement.ValuesAxis != actual.ValuesAxis ||
                placement.ValuesPosition != actual.ValuesPosition ||
                placement.Values.Count != actual.DataFields.Count)
            {
                return false;
            }

            var existing = before.DataFields.ToDictionary(
                field => ExistingKey(
                    field.UniqueName,
                    field.CaptionFingerprint,
                    PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                        field.NumberFormat),
                    field.Position),
                StringComparer.OrdinalIgnoreCase);
            foreach (PivotMeasureValuePlacement requested in placement.Values)
            {
                ModelDataFieldSnapshot current = actual.DataFields.SingleOrDefault(field =>
                    field.Position == requested.Position)!;
                if (current == null) return false;
                if (requested.IsGeneratedMeasure)
                {
                    DesiredModelMeasure definition = definitionsById[requested.DefinitionId!];
                    if (!string.Equals(
                            current.ModelMeasureName,
                            definition.Name,
                            StringComparison.OrdinalIgnoreCase) ||
                        !current.IsModelMeasure)
                    {
                        return false;
                    }

                    ModelDataFieldSnapshot? priorGenerated = before.DataFields
                        .SingleOrDefault(field => string.Equals(
                            field.ModelMeasureName,
                            definition.Name,
                            StringComparison.OrdinalIgnoreCase));
                    if (priorGenerated != null &&
                        (!string.Equals(
                             current.UniqueName,
                             priorGenerated.UniqueName,
                             StringComparison.OrdinalIgnoreCase) ||
                         !string.Equals(
                             current.CaptionFingerprint,
                             priorGenerated.CaptionFingerprint,
                             StringComparison.Ordinal) ||
                         !string.Equals(
                             current.NumberFormat,
                             priorGenerated.NumberFormat,
                             StringComparison.Ordinal)))
                    {
                        return false;
                    }
                }
                else
                {
                    PivotExistingDataFieldIdentity identity = requested.ExistingDataField!;
                    string key = ExistingKey(
                        identity.UniqueName,
                        identity.CurrentCaptionFingerprint,
                        identity.CurrentNumberFormatFingerprint,
                        identity.CurrentPosition);
                    if (!existing.TryGetValue(key, out ModelDataFieldSnapshot? prior) ||
                        !string.Equals(current.UniqueName, prior.UniqueName, StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(current.CaptionFingerprint, prior.CaptionFingerprint, StringComparison.Ordinal) ||
                        !string.Equals(current.NumberFormat, prior.NumberFormat, StringComparison.Ordinal))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static IReadOnlyList<PivotPlusOwnedArtifact> BuildFinalArtifacts(
            IReadOnlyList<DesiredModelMeasure> definitions,
            ModelMeasureWorkbookSnapshot after)
        {
            var live = after.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            return definitions
                .OrderBy(definition => definition.DefinitionId, StringComparer.Ordinal)
                .Select(definition =>
                {
                    LiveModelMeasureSnapshot measure = live[definition.Name];
                    return new PivotPlusOwnedArtifact
                    {
                        Kind = PivotPlusArtifactKind.Measure,
                        ArtifactId = definition.Name,
                        Fingerprint = measure.LiveFingerprint
                    };
                })
                .ToList();
        }

        private static PivotPlusUndoMetadata BuildUndoMetadata(
            string applyId,
            ModelMeasureWorkbookSnapshot before,
            ModelMeasureWorkbookSnapshot after,
            MeasureApplyPlan plan,
            IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts)
        {
            HashSet<string> createdNames = new HashSet<string>(
                plan.Creates.Select(item => item.Definition.Name),
                StringComparer.OrdinalIgnoreCase);
            return new PivotPlusUndoMetadata
            {
                ApplyId = applyId,
                BeforePivotFingerprint = before.SelectedPivotFingerprint,
                AfterPivotFingerprint = after.SelectedPivotFingerprint,
                CreatedArtifacts = finalArtifacts
                    .Where(artifact => createdNames.Contains(artifact.ArtifactId))
                    .Select(CloneArtifact)
                    .ToList(),
                PreviousFieldPlacements = before.SelectedPivot.DataFields.Select(field =>
                {
                    if (field.Position < 1 || field.Position > MaximumValuePlacements)
                    {
                        throw new InvalidOperationException(
                            "A native Values position is outside the supported Undo range.");
                    }

                    return new PivotPlusUndoFieldPlacement
                    {
                        FieldFingerprint = PivotPlusFingerprint.Create(
                            "measure.value.v1",
                            field.UniqueName + "\u001f" + field.CaptionFingerprint + "\u001f" +
                            field.NumberFormat),
                        Area = PivotPlusFieldArea.Data,
                        // Native Pivot data-field positions are one-based. The
                        // persistence contract is deliberately zero-based.
                        Position = field.Position - 1
                    };
                }).ToList()
            };
        }

        private static PivotPlusPendingSemanticApplyMetadata BuildUndoPending(
            MeasureUndoState undo)
        {
            var beforeNames = new HashSet<string>(
                undo.BeforeOwnedArtifacts.Select(item => item.ArtifactId),
                StringComparer.OrdinalIgnoreCase);
            var afterNames = new HashSet<string>(
                undo.AfterOwnedArtifacts.Select(item => item.ArtifactId),
                StringComparer.OrdinalIgnoreCase);
            var transitions = new List<PivotPlusSemanticArtifactTransition>();
            foreach (PivotPlusOwnedArtifact after in undo.AfterOwnedArtifacts)
            {
                if (!beforeNames.Contains(after.ArtifactId))
                {
                    transitions.Add(new PivotPlusSemanticArtifactTransition
                    {
                        Kind = PivotPlusArtifactKind.Measure,
                        ArtifactId = after.ArtifactId,
                        Operation = PivotPlusSemanticArtifactOperation.Delete,
                        BeforeLiveFingerprint = after.Fingerprint,
                        PlannedDefinitionFingerprint =
                            PivotModelMeasureCanonical.CreateDeleteDefinitionFingerprint(
                                after.ArtifactId,
                                after.Fingerprint)
                    });
                }
                else
                {
                    PivotPlusOwnedArtifact prior = undo.BeforeOwnedArtifacts.Single(item =>
                        string.Equals(item.ArtifactId, after.ArtifactId, StringComparison.OrdinalIgnoreCase));
                    if (!string.Equals(prior.Fingerprint, after.Fingerprint, StringComparison.Ordinal))
                    {
                        transitions.Add(new PivotPlusSemanticArtifactTransition
                        {
                            Kind = PivotPlusArtifactKind.Measure,
                            ArtifactId = after.ArtifactId,
                            Operation = PivotPlusSemanticArtifactOperation.Update,
                            BeforeLiveFingerprint = after.Fingerprint,
                            PlannedDefinitionFingerprint = PivotPlusFingerprint.Create(
                                "measure.undo-definition.v1",
                                prior.Fingerprint)
                        });
                    }
                }
            }

            foreach (PivotPlusOwnedArtifact prior in undo.BeforeOwnedArtifacts.Where(item =>
                         !afterNames.Contains(item.ArtifactId)))
            {
                transitions.Add(new PivotPlusSemanticArtifactTransition
                {
                    Kind = PivotPlusArtifactKind.Measure,
                    ArtifactId = prior.ArtifactId,
                    Operation = PivotPlusSemanticArtifactOperation.Create,
                    BeforeLiveFingerprint = string.Empty,
                    PlannedDefinitionFingerprint = PivotPlusFingerprint.Create(
                        "measure.undo-definition.v1",
                        prior.Fingerprint)
                });
            }

            string undoIdFingerprint = PivotPlusFingerprint.Create(
                "measure.undo-id.v1",
                undo.ApplyId);
            string digest = undoIdFingerprint.Substring(
                undoIdFingerprint.LastIndexOf(':') + 1);
            return new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = "undo_" + digest.Substring(0, 32),
                PlanFingerprint = PivotPlusFingerprint.Create(
                    "measure.undo-plan.v1",
                    undo.ApplyId),
                BeforePivotFingerprint = undo.After.SelectedPivotFingerprint,
                ExpectedPivotFingerprint = undo.Before.SelectedPivotFingerprint,
                Transitions = transitions
            };
        }

        private static void DemandUndoAfterState(
            MeasureUndoState undo,
            ModelMeasureWorkbookSnapshot current)
        {
            if (!string.Equals(
                    undo.After.SelectedPivotFingerprint,
                    current.SelectedPivotFingerprint,
                    StringComparison.Ordinal))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "The selected PivotTable changed after Apply, so bounded Undo cannot overwrite it.");
            }

            var currentByName = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact artifact in undo.AfterOwnedArtifacts)
            {
                if (!currentByName.TryGetValue(artifact.ArtifactId, out LiveModelMeasureSnapshot? measure) ||
                    !string.Equals(measure.LiveFingerprint, artifact.Fingerprint, StringComparison.Ordinal))
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "An owned measure changed after Apply, so bounded Undo cannot overwrite it.");
                }
            }

            var beforeByName = undo.BeforeOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var afterByName = undo.AfterOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var changingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact after in afterByName.Values)
            {
                if (!beforeByName.TryGetValue(after.ArtifactId, out PivotPlusOwnedArtifact? prior) ||
                    !string.Equals(prior.Fingerprint, after.Fingerprint, StringComparison.Ordinal))
                {
                    changingNames.Add(after.ArtifactId);
                }
            }

            foreach (PivotPlusOwnedArtifact prior in beforeByName.Values)
            {
                if (!afterByName.ContainsKey(prior.ArtifactId))
                {
                    changingNames.Add(prior.ArtifactId);
                    if (currentByName.ContainsKey(prior.ArtifactId))
                    {
                        throw new PivotModelMeasureUndoUnavailableException(
                            "A model measure now collides with the exact measure Undo would recreate.");
                    }
                }
            }

            var affected = new HashSet<string>(changingNames, StringComparer.OrdinalIgnoreCase);
            bool expanded;
            do
            {
                expanded = false;
                foreach (LiveModelMeasureSnapshot measure in current.Measures)
                {
                    if (affected.Contains(measure.Name) ||
                        !DaxMeasureReferenceScanner.ReadPossibleReferences(measure.Formula)
                            .Any(affected.Contains))
                    {
                        continue;
                    }

                    if (!afterByName.ContainsKey(measure.Name))
                    {
                        throw new PivotModelMeasureUndoUnavailableException(
                            "A user or out-of-scope measure now depends on a measure Undo would change.");
                    }

                    affected.Add(measure.Name);
                    expanded = true;
                }
            }
            while (expanded);

            IReadOnlyList<ModelPivotUsageSnapshot> expectedOthers = undo.After.PivotUsages
                .Where(usage => !usage.IsSelectedTarget)
                .ToList();
            IReadOnlyList<ModelPivotUsageSnapshot> currentOthers = current.PivotUsages
                .Where(usage => !usage.IsSelectedTarget)
                .ToList();
            if (expectedOthers.Count != currentOthers.Count)
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "The workbook Data Model PivotTable set changed after Apply.");
            }

            foreach (ModelPivotUsageSnapshot expected in expectedOthers)
            {
                ModelPivotUsageSnapshot? live = currentOthers.SingleOrDefault(usage =>
                    string.Equals(
                        usage.WorksheetName,
                        expected.WorksheetName,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        usage.PivotTableName,
                        expected.PivotTableName,
                        StringComparison.Ordinal));
                if (live == null || !string.Equals(
                        PivotModelMeasureCanonical.CreatePivotFingerprint(expected),
                        PivotModelMeasureCanonical.CreatePivotFingerprint(live),
                        StringComparison.Ordinal))
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "Another workbook Data Model PivotTable changed after Apply.");
                }
            }

            if (currentOthers.Any(usage => usage.DataFields.Any(field =>
                    !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                    affected.Contains(field.ModelMeasureName!))))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "Another PivotTable uses a measure whose result Undo would change.");
            }
        }

        private static void DemandUndoIntermediateState(
            MeasureUndoState undo,
            ModelMeasureWorkbookSnapshot current)
        {
            var beforeOwned = undo.BeforeOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var afterOwned = undo.AfterOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var involved = new HashSet<string>(
                beforeOwned.Keys.Concat(afterOwned.Keys),
                StringComparer.OrdinalIgnoreCase);
            var startByName = undo.UndoStart.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var priorByName = undo.Before.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var currentByName = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);

            foreach (LiveModelMeasureSnapshot start in undo.UndoStart.Measures.Where(
                         measure => !involved.Contains(measure.Name)))
            {
                if (!currentByName.TryGetValue(
                        start.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        start.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "An unrelated model measure changed during the incomplete Undo.");
                }
            }

            foreach (LiveModelMeasureSnapshot live in current.Measures)
            {
                if (!involved.Contains(live.Name))
                {
                    if (!startByName.TryGetValue(
                            live.Name,
                            out LiveModelMeasureSnapshot? start) ||
                        !string.Equals(
                            start.LiveFingerprint,
                            live.LiveFingerprint,
                            StringComparison.Ordinal))
                    {
                        throw new PivotModelMeasureUndoUnavailableException(
                            "A model measure appeared or changed outside the incomplete Undo.");
                    }

                    continue;
                }

                bool matchesStart = startByName.TryGetValue(
                        live.Name,
                        out LiveModelMeasureSnapshot? startInvolved) &&
                    string.Equals(
                        startInvolved.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal);
                bool matchesPrior = priorByName.TryGetValue(
                        live.Name,
                        out LiveModelMeasureSnapshot? priorInvolved) &&
                    string.Equals(
                        priorInvolved.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal);
                if (!matchesStart && !matchesPrior)
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "An owned measure is neither its exact pre-Undo nor prior definition.");
                }
            }

            try
            {
                DemandExactOtherPivotUsages(
                    undo.UndoStart,
                    current,
                    "Another PivotTable changed during the incomplete Undo.");
            }
            catch (InvalidOperationException exception)
            {
                throw new PivotModelMeasureUndoUnavailableException(exception.Message);
            }

            Dictionary<string, int> requiredUnowned = CreateDataFieldOccurrenceCounts(
                undo.UndoStart.SelectedPivot.DataFields.Where(field =>
                    string.IsNullOrWhiteSpace(field.ModelMeasureName) ||
                    !involved.Contains(field.ModelMeasureName!)));
            Dictionary<string, int> currentUnowned = CreateDataFieldOccurrenceCounts(
                current.SelectedPivot.DataFields.Where(field =>
                    string.IsNullOrWhiteSpace(field.ModelMeasureName) ||
                    !involved.Contains(field.ModelMeasureName!)));
            if (!OccurrenceCountsEqual(requiredUnowned, currentUnowned))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "An unowned Values field changed during the incomplete Undo.");
            }

            Dictionary<string, int> allowedOwned = CreateDataFieldOccurrenceCounts(
                undo.UndoStart.SelectedPivot.DataFields.Where(field =>
                    !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                    involved.Contains(field.ModelMeasureName!)));
            Dictionary<string, int> priorOwned = CreateDataFieldOccurrenceCounts(
                undo.Before.SelectedPivot.DataFields.Where(field =>
                    !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                    involved.Contains(field.ModelMeasureName!)));
            foreach (KeyValuePair<string, int> pair in priorOwned)
            {
                if (!allowedOwned.TryGetValue(pair.Key, out int count) || pair.Value > count)
                {
                    allowedOwned[pair.Key] = pair.Value;
                }
            }
            Dictionary<string, int> currentOwned = CreateDataFieldOccurrenceCounts(
                current.SelectedPivot.DataFields.Where(field =>
                    !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                    involved.Contains(field.ModelMeasureName!)));
            foreach (KeyValuePair<string, int> pair in currentOwned)
            {
                if (!allowedOwned.TryGetValue(pair.Key, out int allowed) || pair.Value > allowed)
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "The incomplete Undo contains an unexpected owned Values field.");
                }
            }

            if (current.PivotUsages.Where(usage => !usage.IsSelectedTarget).Any(usage =>
                    usage.DataFields.Any(field =>
                        !string.IsNullOrWhiteSpace(field.ModelMeasureName) &&
                        involved.Contains(field.ModelMeasureName!))))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "Another PivotTable uses a measure involved in the incomplete Undo.");
            }
        }

        private static void DemandUndoFinal(
            MeasureUndoState undo,
            ModelMeasureWorkbookSnapshot current,
            ModelMeasureWorkbookSnapshot actual)
        {
            if (!string.Equals(
                    undo.Before.SelectedPivotFingerprint,
                    actual.SelectedPivotFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Undo did not restore the prior Values layout.");
            }

            var actualByName = actual.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var beforeOwned = undo.BeforeOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var afterOwned = undo.AfterOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var involved = new HashSet<string>(
                beforeOwned.Keys.Concat(afterOwned.Keys),
                StringComparer.OrdinalIgnoreCase);
            HashSet<string> changingNames = CreateUndoChangingNames(
                beforeOwned,
                afterOwned);

            foreach (PivotPlusOwnedArtifact artifact in beforeOwned.Values)
            {
                if (!actualByName.TryGetValue(artifact.ArtifactId, out LiveModelMeasureSnapshot? restored) ||
                    !string.Equals(artifact.Fingerprint, restored.LiveFingerprint, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Undo did not restore an exact prior model measure.");
                }
            }

            foreach (string createdName in afterOwned.Keys.Where(name => !beforeOwned.ContainsKey(name)))
            {
                if (actualByName.ContainsKey(createdName))
                {
                    throw new InvalidOperationException("Undo retained a model measure created by Apply.");
                }
            }

            foreach (LiveModelMeasureSnapshot prior in current.Measures.Where(
                         measure => !involved.Contains(measure.Name)))
            {
                if (!actualByName.TryGetValue(prior.Name, out LiveModelMeasureSnapshot? unchanged) ||
                    !string.Equals(prior.LiveFingerprint, unchanged.LiveFingerprint, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Undo changed an unrelated model measure.");
                }
            }


            var expectedMeasureNames = new HashSet<string>(
                current.Measures
                    .Where(measure => !involved.Contains(measure.Name))
                    .Select(measure => measure.Name)
                    .Concat(beforeOwned.Keys),
                StringComparer.OrdinalIgnoreCase);
            if (actual.Measures.Count != expectedMeasureNames.Count ||
                actual.Measures.Any(measure => !expectedMeasureNames.Contains(measure.Name)))
            {
                throw new InvalidOperationException(
                    "The workbook model-measure inventory changed during Undo.");
            }

            DemandNoPostRefreshUnsafeRelationships(
                actual,
                changingNames,
                new HashSet<string>(beforeOwned.Keys, StringComparer.OrdinalIgnoreCase),
                "Undo");

            foreach (ModelPivotUsageSnapshot prior in current.PivotUsages.Where(
                         usage => !usage.IsSelectedTarget))
            {
                ModelPivotUsageSnapshot unchanged = actual.PivotUsages.SingleOrDefault(usage =>
                    !usage.IsSelectedTarget &&
                    string.Equals(usage.WorksheetName, prior.WorksheetName, StringComparison.Ordinal) &&
                    string.Equals(usage.PivotTableName, prior.PivotTableName, StringComparison.Ordinal)) ??
                    throw new InvalidOperationException("Undo removed another model PivotTable.");
                if (!string.Equals(
                        PivotModelMeasureCanonical.CreatePivotFingerprint(prior),
                        PivotModelMeasureCanonical.CreatePivotFingerprint(unchanged),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Undo changed another model PivotTable.");
                }
            }


            if (actual.PivotUsages.Count != current.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model PivotTable inventory changed during Undo.");
            }
        }

        private static HashSet<string> CreateUndoChangingNames(
            IReadOnlyDictionary<string, PivotPlusOwnedArtifact> beforeOwned,
            IReadOnlyDictionary<string, PivotPlusOwnedArtifact> afterOwned)
        {
            var changed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact after in afterOwned.Values)
            {
                if (!beforeOwned.TryGetValue(
                        after.ArtifactId,
                        out PivotPlusOwnedArtifact? prior) ||
                    !string.Equals(
                        prior.Fingerprint,
                        after.Fingerprint,
                        StringComparison.Ordinal))
                {
                    changed.Add(after.ArtifactId);
                }
            }

            foreach (PivotPlusOwnedArtifact prior in beforeOwned.Values)
            {
                if (!afterOwned.ContainsKey(prior.ArtifactId))
                {
                    changed.Add(prior.ArtifactId);
                }
            }

            return changed;
        }

        private static bool IsExactUndoFinal(
            MeasureUndoState undo,
            ModelMeasureWorkbookSnapshot current)
        {
            try
            {
                DemandUndoFinal(undo, undo.UndoStart, current);
                return true;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private static void DemandExactActiveMeasureOwnership(
            PivotPlusWorkbookMetadata metadata,
            IReadOnlyList<PivotPlusOwnedArtifact> expected,
            string operation)
        {
            if (!OwnedMeasureArtifactsEqual(metadata.Artifacts, expected))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    operation + " does not match the exact active measure ownership receipts.");
            }
        }

        private static void DemandExactWorkbookSnapshot(
            ModelMeasureWorkbookSnapshot expected,
            ModelMeasureWorkbookSnapshot actual,
            string operation)
        {
            if (!string.Equals(
                    expected.SelectedPivotFingerprint,
                    actual.SelectedPivotFingerprint,
                    StringComparison.Ordinal) ||
                expected.Measures.Count != actual.Measures.Count ||
                expected.PivotUsages.Count != actual.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    operation + " did not restore the exact workbook Data Model snapshot.");
            }

            var actualMeasures = actual.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot measure in expected.Measures)
            {
                if (!actualMeasures.TryGetValue(measure.Name, out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        measure.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        operation + " did not restore an exact model measure.");
                }
            }

            foreach (ModelPivotUsageSnapshot pivot in expected.PivotUsages)
            {
                ModelPivotUsageSnapshot live = actual.PivotUsages.SingleOrDefault(candidate =>
                    string.Equals(
                        candidate.WorksheetName,
                        pivot.WorksheetName,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        candidate.PivotTableName,
                        pivot.PivotTableName,
                        StringComparison.Ordinal)) ??
                    throw new InvalidOperationException(
                        operation + " did not restore an exact workbook PivotTable set.");
                if (live.IsSelectedTarget != pivot.IsSelectedTarget ||
                    !string.Equals(
                        PivotModelMeasureCanonical.CreatePivotFingerprint(pivot),
                        PivotModelMeasureCanonical.CreatePivotFingerprint(live),
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        operation + " did not restore an exact workbook PivotTable layout.");
                }
            }
        }

        private static void DemandSamePendingPlan(
            PivotPlusPendingSemanticApplyMetadata existing,
            PivotPlusPendingSemanticApplyMetadata proposed)
        {
            if (!string.Equals(existing.ApplyId, proposed.ApplyId, StringComparison.Ordinal) ||
                !string.Equals(existing.PlanFingerprint, proposed.PlanFingerprint, StringComparison.Ordinal) ||
                !string.Equals(existing.BeforePivotFingerprint, proposed.BeforePivotFingerprint, StringComparison.Ordinal) ||
                !string.Equals(existing.ExpectedPivotFingerprint, proposed.ExpectedPivotFingerprint, StringComparison.Ordinal) ||
                !TransitionListsEqual(existing.Transitions, proposed.Transitions))
            {
                throw new InvalidOperationException(
                    "The retry does not match the exact pending semantic Apply.");
            }
        }

        private static void DemandSameMeasurePendingSlice(
            PivotPlusPendingSemanticApplyMetadata existing,
            PivotPlusPendingSemanticApplyMetadata proposed)
        {
            if (!string.Equals(existing.ApplyId, proposed.ApplyId, StringComparison.Ordinal) ||
                !string.Equals(
                    existing.BeforePivotFingerprint,
                    proposed.BeforePivotFingerprint,
                    StringComparison.Ordinal) ||
                !TransitionListsEqual(existing.Transitions, proposed.Transitions))
            {
                throw new InvalidOperationException(
                    "The measure participant does not match the measure portion of the pending combined semantic Apply.");
            }
        }

        private static bool TransitionListsEqual(
            IEnumerable<PivotPlusSemanticArtifactTransition> first,
            IEnumerable<PivotPlusSemanticArtifactTransition> second)
        {
            List<PivotPlusSemanticArtifactTransition> left = first
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ArtifactId, StringComparer.Ordinal)
                .ToList();
            List<PivotPlusSemanticArtifactTransition> right = second
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ArtifactId, StringComparer.Ordinal)
                .ToList();
            if (left.Count != right.Count) return false;
            for (int index = 0; index < left.Count; index++)
            {
                PivotPlusSemanticArtifactTransition a = left[index];
                PivotPlusSemanticArtifactTransition b = right[index];
                if (a.Kind != b.Kind ||
                    a.Operation != b.Operation ||
                    !string.Equals(a.ArtifactId, b.ArtifactId, StringComparison.Ordinal) ||
                    !string.Equals(
                        a.BeforeLiveFingerprint,
                        b.BeforeLiveFingerprint,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        a.PlannedDefinitionFingerprint,
                        b.PlannedDefinitionFingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static PivotPlusPendingSemanticApplyMetadata? CreateParticipantPendingSlice(
            PivotPlusPendingSemanticApplyMetadata? combinedPending)
        {
            if (combinedPending == null) return null;
            return new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = combinedPending.ApplyId,
                PlanFingerprint = combinedPending.PlanFingerprint,
                BeforePivotFingerprint = combinedPending.BeforePivotFingerprint,
                ExpectedPivotFingerprint = combinedPending.ExpectedPivotFingerprint,
                Transitions = combinedPending.Transitions
                    .Where(transition => transition.Kind == PivotPlusArtifactKind.Measure)
                    .Select(transition => new PivotPlusSemanticArtifactTransition
                    {
                        Kind = transition.Kind,
                        ArtifactId = transition.ArtifactId,
                        Operation = transition.Operation,
                        BeforeLiveFingerprint = transition.BeforeLiveFingerprint,
                        PlannedDefinitionFingerprint = transition.PlannedDefinitionFingerprint
                    })
                    .ToList()
            };
        }

        private static void DemandParticipantBase(
            PivotPlusWorkbookMetadata baseMetadata,
            string setupId,
            PivotTargetIdentity target)
        {
            PivotPlusMetadataValidator.Validate(baseMetadata);
            if (!string.Equals(baseMetadata.SetupId, setupId, StringComparison.Ordinal) ||
                !string.Equals(
                    baseMetadata.TargetWorksheetName,
                    target.WorksheetName,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    baseMetadata.TargetPivotTableName,
                    target.PivotTableName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The measure participant base metadata belongs to a different setup or PivotTable target.");
            }

            if (baseMetadata.RecoveryPhase != PivotPlusRecoveryPhase.None ||
                baseMetadata.PendingSemanticApply != null)
            {
                throw new InvalidOperationException(
                    "The measure participant requires active base metadata with conversion and semantic journals supplied separately.");
            }
        }

        private static void DemandParticipantPending(
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? combinedPending)
        {
            if (combinedPending == null) return;
            var candidate = new PivotPlusWorkbookMetadata
            {
                SchemaVersion = baseMetadata.SchemaVersion,
                SetupId = baseMetadata.SetupId,
                TargetWorksheetName = baseMetadata.TargetWorksheetName,
                TargetPivotTableName = baseMetadata.TargetPivotTableName,
                RecoveryPhase = baseMetadata.RecoveryPhase,
                TargetAnchorAddress = baseMetadata.TargetAnchorAddress,
                StagingStateFingerprint = baseMetadata.StagingStateFingerprint,
                Artifacts = baseMetadata.Artifacts.Select(CloneArtifact).ToList(),
                PendingSemanticApply = combinedPending,
                Undo = baseMetadata.Undo
            };
            PivotPlusMetadataValidator.Validate(candidate);
        }

        private static void DemandUniqueArtifacts(IReadOnlyList<PivotPlusOwnedArtifact> artifacts)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PivotPlusOwnedArtifact artifact in artifacts)
            {
                if (!names.Add(artifact.ArtifactId))
                {
                    throw new InvalidOperationException("Duplicate measure ownership was found.");
                }
            }
        }

        private static void DemandSameTarget(
            PivotTargetIdentity expected,
            PivotTargetIdentity actual)
        {
            if (!string.Equals(expected.WorkbookId, actual.WorkbookId, StringComparison.Ordinal) ||
                !string.Equals(expected.WorksheetName, actual.WorksheetName, StringComparison.Ordinal) ||
                !string.Equals(expected.PivotTableName, actual.PivotTableName, StringComparison.Ordinal))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "Undo was requested for a different workbook or PivotTable target.");
            }
        }

        private static string ExistingKey(
            string uniqueName,
            string captionFingerprint,
            string numberFormatFingerprint,
            int position)
        {
            return uniqueName + "\u001f" + captionFingerprint + "\u001f" +
                   numberFormatFingerprint + "\u001f" +
                   position.ToString(CultureInfo.InvariantCulture);
        }

        private bool TryPromoteRecoveredUndo(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            string planFingerprint,
            ModelMeasureWorkbookSnapshot current,
            PivotPlusWorkbookMetadata activeMetadata,
            out MeasureUndoState? recovered)
        {
            recovered = null;
            if (!TryGetPendingApplyUndo(workbook, setupId, out MeasureUndoState? pending) ||
                pending == null ||
                !string.Equals(
                    pending.PlanFingerprint,
                    planFingerprint,
                    StringComparison.Ordinal) ||
                !UndoIsUsable(pending, target, current) ||
                !OwnedMeasureArtifactsEqual(
                    activeMetadata.Artifacts,
                    pending.AfterOwnedArtifacts))
            {
                return false;
            }

            RememberUndo(workbook, setupId, pending);
            ForgetPendingApplyUndo(workbook, setupId);
            ForgetPendingApplyUndoSeed(workbook, setupId);
            recovered = pending;
            return true;
        }

        private static bool UndoIsUsable(
            MeasureUndoState undo,
            PivotTargetIdentity target,
            ModelMeasureWorkbookSnapshot current)
        {
            try
            {
                DemandSameTarget(undo.Target, target);
                DemandUndoAfterState(undo, current);
                return true;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private static bool OwnedMeasureArtifactsEqual(
            IEnumerable<PivotPlusOwnedArtifact> actual,
            IEnumerable<PivotPlusOwnedArtifact> expected)
        {
            List<PivotPlusOwnedArtifact> left = actual
                .Where(artifact => artifact.Kind == PivotPlusArtifactKind.Measure)
                .OrderBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
                .ToList();
            List<PivotPlusOwnedArtifact> right = expected
                .OrderBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
                .ToList();
            return left.Count == right.Count && left.Zip(right, (first, second) =>
                    string.Equals(
                        first.ArtifactId,
                        second.ArtifactId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        first.Fingerprint,
                        second.Fingerprint,
                        StringComparison.Ordinal))
                .All(value => value);
        }

        private void RememberUndo(object workbook, string setupId, MeasureUndoState state)
        {
            UndoLedger ledger = undoLedgers.GetOrCreateValue(workbook);
            lock (ledger.States)
            {
                ledger.States[setupId] = state;
            }
        }

        private bool TryGetUndo(object workbook, string setupId, out MeasureUndoState? state)
        {
            state = null;
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return false;
            lock (ledger.States)
            {
                return ledger.States.TryGetValue(setupId, out state);
            }
        }

        private void ForgetUndo(object workbook, string setupId)
        {
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return;
            lock (ledger.States)
            {
                ledger.States.Remove(setupId);
            }
        }

        private void RememberPendingApplyUndo(
            object workbook,
            string setupId,
            MeasureUndoState state)
        {
            UndoLedger ledger = undoLedgers.GetOrCreateValue(workbook);
            lock (ledger.States)
            {
                ledger.PendingApplyStates[setupId] = state;
            }
        }

        private bool TryGetPendingApplyUndo(
            object workbook,
            string setupId,
            out MeasureUndoState? state)
        {
            state = null;
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return false;
            lock (ledger.States)
            {
                return ledger.PendingApplyStates.TryGetValue(setupId, out state);
            }
        }

        private void ForgetPendingApplyUndo(object workbook, string setupId)
        {
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return;
            lock (ledger.States)
            {
                ledger.PendingApplyStates.Remove(setupId);
            }
        }

        private void RememberPendingApplyUndoSeed(
            object workbook,
            string setupId,
            PendingApplyUndoSeed seed)
        {
            UndoLedger ledger = undoLedgers.GetOrCreateValue(workbook);
            lock (ledger.States)
            {
                ledger.PendingApplySeeds[setupId] = seed;
            }
        }

        private bool TryGetPendingApplyUndoSeed(
            object workbook,
            string setupId,
            string applyId,
            string planFingerprint,
            PivotTargetIdentity target,
            out PendingApplyUndoSeed? seed)
        {
            seed = null;
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return false;
            lock (ledger.States)
            {
                if (!ledger.PendingApplySeeds.TryGetValue(setupId, out PendingApplyUndoSeed? candidate) ||
                    !string.Equals(candidate.ApplyId, applyId, StringComparison.Ordinal) ||
                    !string.Equals(
                        candidate.PlanFingerprint,
                        planFingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }

                try
                {
                    DemandSameTarget(candidate.Target, target);
                }
                catch (PivotModelMeasureUndoUnavailableException)
                {
                    return false;
                }

                seed = candidate;
                return true;
            }
        }

        private void ForgetPendingApplyUndoSeed(object workbook, string setupId)
        {
            if (!undoLedgers.TryGetValue(workbook, out UndoLedger? ledger)) return;
            lock (ledger.States)
            {
                ledger.PendingApplySeeds.Remove(setupId);
            }
        }

        private void Enter()
        {
            lock (synchronization)
            {
                if (applyActive)
                {
                    throw new InvalidOperationException(
                        "A PivotTable+ model-measure mutation is already active.");
                }

                applyActive = true;
            }
        }

        private void Exit()
        {
            lock (synchronization)
            {
                applyActive = false;
            }
        }

        private static PivotPlusOwnedArtifact CloneArtifact(PivotPlusOwnedArtifact artifact)
        {
            return new PivotPlusOwnedArtifact
            {
                Kind = artifact.Kind,
                ArtifactId = artifact.ArtifactId,
                Fingerprint = artifact.Fingerprint
            };
        }

        private sealed class MeasureUpsert
        {
            public MeasureUpsert(
                DesiredModelMeasure definition,
                LiveModelMeasureSnapshot? before)
            {
                Definition = definition;
                Before = before;
            }

            public DesiredModelMeasure Definition { get; }

            public LiveModelMeasureSnapshot? Before { get; }
        }

        private sealed class MeasureApplyPlan
        {
            public MeasureApplyPlan(
                PivotPlusPendingSemanticApplyMetadata pending,
                IReadOnlyList<MeasureUpsert> creates,
                IReadOnlyList<MeasureUpsert> updates,
                IReadOnlyList<LiveModelMeasureSnapshot> deletes,
                IReadOnlyList<string> deletedNames,
                int deleteCount,
                bool isNoChange,
                bool placementAlreadyFinal,
                bool placementNeedsPreviewRepair,
                ModelMeasureWorkbookSnapshot placementPreview,
                bool hasSessionPreview)
            {
                Pending = pending;
                Creates = creates;
                Updates = updates;
                Deletes = deletes;
                DeletedNames = deletedNames;
                DeleteCount = deleteCount;
                Upserts = creates.Concat(updates)
                    .OrderBy(item => item.Definition.CreationOrder)
                    .ToList();
                IsNoChange = isNoChange;
                PlacementAlreadyFinal = placementAlreadyFinal;
                PlacementNeedsPreviewRepair = placementNeedsPreviewRepair;
                PlacementPreview = placementPreview;
                HasSessionPreview = hasSessionPreview;
            }

            public PivotPlusPendingSemanticApplyMetadata Pending { get; }

            public IReadOnlyList<MeasureUpsert> Creates { get; }

            public IReadOnlyList<MeasureUpsert> Updates { get; }

            public IReadOnlyList<MeasureUpsert> Upserts { get; }

            public IReadOnlyList<LiveModelMeasureSnapshot> Deletes { get; }

            public IReadOnlyList<string> DeletedNames { get; }

            public int DeleteCount { get; }

            public bool IsNoChange { get; }

            public bool PlacementAlreadyFinal { get; }

            public bool PlacementNeedsPreviewRepair { get; }

            public ModelMeasureWorkbookSnapshot PlacementPreview { get; }

            public bool HasSessionPreview { get; }
        }

        private sealed class MeasureUndoState
        {
            public MeasureUndoState(PivotModelMeasureUndoContribution contribution)
                : this(
                    contribution.ApplyId,
                    contribution.PlanFingerprint,
                    contribution.Target,
                    contribution.Before,
                    contribution.After,
                    contribution.BeforeOwnedArtifacts,
                    contribution.AfterOwnedArtifacts,
                    contribution.WorkbookUndo)
            {
            }

            public MeasureUndoState(
                string applyId,
                string planFingerprint,
                PivotTargetIdentity target,
                ModelMeasureWorkbookSnapshot before,
                ModelMeasureWorkbookSnapshot after,
                IReadOnlyList<PivotPlusOwnedArtifact> beforeOwnedArtifacts,
                IReadOnlyList<PivotPlusOwnedArtifact> afterOwnedArtifacts,
                PivotPlusUndoMetadata workbookUndo,
                ModelMeasureWorkbookSnapshot? undoStart = null)
            {
                ApplyId = applyId;
                PlanFingerprint = planFingerprint;
                Target = target;
                Before = before;
                After = after;
                BeforeOwnedArtifacts = beforeOwnedArtifacts;
                AfterOwnedArtifacts = afterOwnedArtifacts;
                WorkbookUndo = workbookUndo;
                UndoStart = undoStart ?? after;
            }

            public string ApplyId { get; }

            public string PlanFingerprint { get; }

            public PivotTargetIdentity Target { get; }

            public ModelMeasureWorkbookSnapshot Before { get; }

            public ModelMeasureWorkbookSnapshot After { get; }

            public IReadOnlyList<PivotPlusOwnedArtifact> BeforeOwnedArtifacts { get; }

            public IReadOnlyList<PivotPlusOwnedArtifact> AfterOwnedArtifacts { get; }

            public PivotPlusUndoMetadata WorkbookUndo { get; }

            /// <summary>
            /// Exact post-Apply snapshot captured before the first Undo
            /// mutation. It remains the authority across partial Undo retries.
            /// </summary>
            public ModelMeasureWorkbookSnapshot UndoStart { get; }

            public MeasureUndoState WithAfter(
                ModelMeasureWorkbookSnapshot after,
                IReadOnlyList<PivotPlusOwnedArtifact> afterOwnedArtifacts)
            {
                return new MeasureUndoState(
                    ApplyId,
                    PlanFingerprint,
                    Target,
                    Before,
                    after,
                    BeforeOwnedArtifacts,
                    afterOwnedArtifacts,
                    WorkbookUndo,
                    undoStart: after);
            }

            public MeasureUndoState WithUndoStart(ModelMeasureWorkbookSnapshot undoStart)
            {
                return new MeasureUndoState(
                    ApplyId,
                    PlanFingerprint,
                    Target,
                    Before,
                    After,
                    BeforeOwnedArtifacts,
                    AfterOwnedArtifacts,
                    WorkbookUndo,
                    undoStart);
            }
        }

        private sealed class PendingApplyUndoSeed
        {
            public PendingApplyUndoSeed(
                string applyId,
                string planFingerprint,
                PivotTargetIdentity target,
                ModelMeasureWorkbookSnapshot before,
                IReadOnlyList<PivotPlusOwnedArtifact> beforeOwnedArtifacts)
            {
                ApplyId = applyId;
                PlanFingerprint = planFingerprint;
                Target = target;
                Before = before;
                BeforeOwnedArtifacts = beforeOwnedArtifacts;
            }

            public string ApplyId { get; }

            public string PlanFingerprint { get; }

            public PivotTargetIdentity Target { get; }

            public ModelMeasureWorkbookSnapshot Before { get; }

            public IReadOnlyList<PivotPlusOwnedArtifact> BeforeOwnedArtifacts { get; }
        }

        private sealed class UndoLedger
        {
            public Dictionary<string, MeasureUndoState> States { get; } =
                new Dictionary<string, MeasureUndoState>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, MeasureUndoState> PendingApplyStates { get; } =
                new Dictionary<string, MeasureUndoState>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, PendingApplyUndoSeed> PendingApplySeeds { get; } =
                new Dictionary<string, PendingApplyUndoSeed>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
