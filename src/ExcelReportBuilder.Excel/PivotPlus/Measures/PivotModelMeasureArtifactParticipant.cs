using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    /// <summary>
    /// Trusted, compiler-produced binding exposed to a combined semantic
    /// coordinator. It contains no DAX and does not derive identity from a
    /// PivotTable caption.
    /// </summary>
    internal sealed class PivotModelMeasureArtifactBinding
    {
        public PivotModelMeasureArtifactBinding(
            string definitionId,
            string hostMeasureName,
            string definitionFingerprint)
        {
            DefinitionId = definitionId;
            HostMeasureName = hostMeasureName;
            DefinitionFingerprint = definitionFingerprint;
        }

        public string DefinitionId { get; }

        public string HostMeasureName { get; }

        public string DefinitionFingerprint { get; }
    }

    /// <summary>
    /// Measure-artifact-only participant for one combined semantic Apply.
    /// Values placement and combined Pivot fingerprints deliberately remain
    /// owned by the higher-level layout participant.
    /// </summary>
    internal sealed class PivotModelMeasureArtifactPreparedMutation
    {
        private readonly Action refresh;
        private readonly Func<ModelMeasureWorkbookSnapshot> verify;
        private readonly Action verifyRollback;
        private readonly Func<ModelMeasureWorkbookSnapshot, IReadOnlyList<PivotPlusOwnedArtifact>>
            buildArtifacts;
        private readonly Action<string, string> primeUndoContribution;
        private readonly Func<string, string, ModelMeasureWorkbookSnapshot,
            PivotModelMeasureArtifactUndoContribution?> buildUndoContribution;

        public PivotModelMeasureArtifactPreparedMutation(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            string applyId,
            string participantPlanFingerprint,
            IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions,
            IReadOnlyDictionary<string, PivotModelMeasureArtifactBinding> definitionBindings,
            IReadOnlyList<PivotMutationStep> upsertSteps,
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
                PivotModelMeasureArtifactUndoContribution?> buildUndoContribution)
        {
            Target = target;
            Before = before;
            ApplyId = applyId;
            ParticipantPlanFingerprint = participantPlanFingerprint;
            Transitions = transitions;
            DefinitionBindings = definitionBindings;
            UpsertSteps = upsertSteps;
            DeleteSteps = deleteSteps;
            Steps = new ReadOnlyCollection<PivotMutationStep>(
                upsertSteps.Concat(deleteSteps).ToList());
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

        public string ApplyId { get; }

        public string ParticipantPlanFingerprint { get; }

        public IReadOnlyList<PivotPlusSemanticArtifactTransition> Transitions { get; }

        public IReadOnlyDictionary<string, PivotModelMeasureArtifactBinding> DefinitionBindings { get; }

        /// <summary>Run before named-set upserts.</summary>
        public IReadOnlyList<PivotMutationStep> UpsertSteps { get; }

        /// <summary>Run after named-set deletes and all layout work.</summary>
        public IReadOnlyList<PivotMutationStep> DeleteSteps { get; }

        public IReadOnlyList<PivotMutationStep> Steps { get; }

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

        public void VerifyRollback()
        {
            verifyRollback();
        }

        public IReadOnlyList<PivotPlusOwnedArtifact> BuildArtifacts(
            ModelMeasureWorkbookSnapshot verified)
        {
            return buildArtifacts(verified);
        }

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

        public PivotModelMeasureArtifactUndoContribution? BuildUndoContribution(
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
    /// Session-only measure artifact state for a combined semantic Undo. The
    /// exact formulas remain only in memory and no Values layout is owned here.
    /// </summary>
    internal sealed class PivotModelMeasureArtifactUndoContribution
    {
        public PivotModelMeasureArtifactUndoContribution(
            string applyId,
            string combinedPlanFingerprint,
            string participantPlanFingerprint,
            PivotTargetIdentity target,
            ModelMeasureWorkbookSnapshot before,
            ModelMeasureWorkbookSnapshot after,
            IReadOnlyList<PivotPlusOwnedArtifact> beforeOwnedArtifacts,
            IReadOnlyList<PivotPlusOwnedArtifact> afterOwnedArtifacts)
        {
            ApplyId = applyId;
            CombinedPlanFingerprint = combinedPlanFingerprint;
            ParticipantPlanFingerprint = participantPlanFingerprint;
            Target = target;
            Before = before;
            After = after;
            BeforeOwnedArtifacts = beforeOwnedArtifacts;
            AfterOwnedArtifacts = afterOwnedArtifacts;
        }

        public string ApplyId { get; }

        public string CombinedPlanFingerprint { get; }

        public string ParticipantPlanFingerprint { get; }

        public PivotTargetIdentity Target { get; }

        public ModelMeasureWorkbookSnapshot Before { get; }

        public ModelMeasureWorkbookSnapshot After { get; }

        public IReadOnlyList<PivotPlusOwnedArtifact> BeforeOwnedArtifacts { get; }

        public IReadOnlyList<PivotPlusOwnedArtifact> AfterOwnedArtifacts { get; }
    }

    /// <summary>
    /// Executable artifact-only contribution to a combined Undo. It never
    /// mutates Values placement and leaves the one refresh to its coordinator.
    /// </summary>
    internal sealed class PivotModelMeasureArtifactPreparedUndo
    {
        private readonly Action refresh;
        private readonly Action verify;
        private readonly Action verifyRollback;

        public PivotModelMeasureArtifactPreparedUndo(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            IReadOnlyList<PivotMutationStep> upsertSteps,
            IReadOnlyList<PivotMutationStep> deleteSteps,
            Action refresh,
            Action verify,
            Action verifyRollback)
        {
            Target = target;
            Before = before;
            UpsertSteps = upsertSteps;
            DeleteSteps = deleteSteps;
            Steps = new ReadOnlyCollection<PivotMutationStep>(
                upsertSteps.Concat(deleteSteps).ToList());
            this.refresh = refresh;
            this.verify = verify;
            this.verifyRollback = verifyRollback;
        }

        public BoundModelMeasureTarget Target { get; }

        public ModelMeasureWorkbookSnapshot Before { get; }

        public IReadOnlyList<PivotMutationStep> UpsertSteps { get; }

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

    public sealed partial class PivotModelMeasureMutationService
    {
        internal string ComputeArtifactParticipantPlanFingerprint(
            string setupId,
            PivotDaxCompilation compilation,
            PivotPlusPendingSemanticApplyMetadata pending)
        {
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
            if (pending == null) throw new ArgumentNullException(nameof(pending));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            IReadOnlyList<DesiredModelMeasure> definitions = CompileDefinitions(
                setupId,
                compilation);
            IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions =
                pending.Transitions
                    .Where(item => item.Kind == PivotPlusArtifactKind.Measure)
                    .ToList();
            return PivotModelMeasureCanonical.CreateArtifactPlanFingerprint(
                definitions,
                transitions);
        }

        internal PivotModelMeasureArtifactPreparedMutation PrepareArtifactParticipant(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId,
            PivotDaxCompilation compilation,
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            PivotModelMeasureParticipantRetryBinding? retryBinding = null)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
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

            if (existingPending == null && retryBinding != null)
            {
                throw new InvalidOperationException(
                    "A participant retry binding cannot be supplied without a pending combined Apply.");
            }

            if (existingPending != null &&
                (retryBinding == null || !string.Equals(
                    existingPending.PlanFingerprint,
                    retryBinding.CombinedPlanFingerprint,
                    StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The measure artifact retry was not bound to the exact durable combined plan.");
            }

            ArtifactMeasureApplyPlan plan = BuildArtifactPlan(
                baseMetadata,
                existingPending,
                before,
                definitions,
                definitionsById);
            if (existingPending != null && !string.Equals(
                    plan.PlanFingerprint,
                    retryBinding!.MeasurePlanFingerprint,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The recomputed measure artifact plan does not match the exact pending participant plan.");
            }

            return CreateArtifactPreparedMutation(
                workbook,
                setupId,
                target,
                before,
                plan,
                definitions,
                definitionsById,
                baseMetadata.Artifacts.ToList(),
                allowNewUndoSeed: existingPending == null);
        }

        internal PivotModelMeasureArtifactPreparedUndo PrepareArtifactUndoParticipant(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            PivotModelMeasureArtifactUndoContribution contribution)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (contribution == null) throw new ArgumentNullException(nameof(contribution));

            BoundModelMeasureTarget target = gateway.Bind(workbook, pivotTable, context);
            DemandSameTarget(contribution.Target, target.Identity);
            ModelMeasureWorkbookSnapshot current = gateway.Capture(target);
            DemandArtifactUndoStartOrIntermediate(contribution, current);
            return CreateArtifactPreparedUndo(target, current, contribution);
        }

        private PivotModelMeasureArtifactPreparedMutation CreateArtifactPreparedMutation(
            object workbook,
            string setupId,
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot before,
            ArtifactMeasureApplyPlan plan,
            IReadOnlyList<DesiredModelMeasure> definitions,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            IReadOnlyList<PivotPlusOwnedArtifact> baseArtifacts,
            bool allowNewUndoSeed)
        {
            var upsertSteps = new List<PivotMutationStep>();
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

            foreach (LiveModelMeasureSnapshot deletion in plan.Deletes)
            {
                LiveModelMeasureSnapshot captured = deletion;
                deleteSteps.Add(new PivotMutationStep(
                    "delete model measure " + captured.Name,
                    () => gateway.DeleteMeasure(target, captured),
                    () => gateway.RestoreMeasure(target, captured)));
            }

            var bindings = new ReadOnlyDictionary<string, PivotModelMeasureArtifactBinding>(
                definitions.ToDictionary(
                    definition => definition.DefinitionId,
                    definition => new PivotModelMeasureArtifactBinding(
                        definition.DefinitionId,
                        definition.Name,
                        definition.DefinitionFingerprint),
                    StringComparer.OrdinalIgnoreCase));

            return new PivotModelMeasureArtifactPreparedMutation(
                target,
                before,
                plan.ApplyId,
                plan.PlanFingerprint,
                new ReadOnlyCollection<PivotPlusSemanticArtifactTransition>(
                    plan.Transitions.Select(CloneTransition).ToList()),
                bindings,
                new ReadOnlyCollection<PivotMutationStep>(upsertSteps),
                new ReadOnlyCollection<PivotMutationStep>(deleteSteps),
                plan.IsNoChange,
                plan.Creates.Count,
                plan.Updates.Count,
                plan.DeleteCount,
                () => gateway.Refresh(target),
                () =>
                {
                    ModelMeasureWorkbookSnapshot after = gateway.Capture(target);
                    VerifyArtifactFinal(
                        before,
                        after,
                        plan,
                        definitionsById,
                        upsertReceipts);
                    return after;
                },
                () => DemandExactArtifactSnapshot(
                    before,
                    gateway.Capture(target),
                    "combined measure artifact Apply rollback"),
                verified => BuildFinalArtifacts(definitions, verified),
                (combinedApplyId, combinedPlanFingerprint) =>
                {
                    if (!allowNewUndoSeed) return;
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
                        return null;
                    }

                    IReadOnlyList<PivotPlusOwnedArtifact> finalArtifacts =
                        BuildFinalArtifacts(definitions, verified);
                    return new PivotModelMeasureArtifactUndoContribution(
                        combinedApplyId,
                        combinedPlanFingerprint,
                        plan.PlanFingerprint,
                        target.Identity,
                        seed.Before,
                        verified,
                        seed.BeforeOwnedArtifacts.Select(CloneArtifact).ToList(),
                        finalArtifacts.Select(CloneArtifact).ToList());
                });
        }

        private static ArtifactMeasureApplyPlan BuildArtifactPlan(
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata? existingPending,
            ModelMeasureWorkbookSnapshot before,
            IReadOnlyList<DesiredModelMeasure> definitions,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById)
        {
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
            PivotPlusPendingSemanticApplyMetadata? measurePending =
                CreateParticipantPendingSlice(existingPending);
            DemandLiveOwnership(active, measurePending, liveByName, definitionsById);

            var creates = new List<MeasureUpsert>();
            var updates = new List<MeasureUpsert>();
            foreach (DesiredModelMeasure definition in definitions
                         .OrderBy(item => item.CreationOrder))
            {
                if (!activeByName.TryGetValue(definition.Name, out PivotPlusOwnedArtifact? owned))
                {
                    if (liveByName.TryGetValue(
                            definition.Name,
                            out LiveModelMeasureSnapshot? collision))
                    {
                        if (existingPending == null || !string.Equals(
                                collision.Description,
                                definition.DescriptionMarker,
                                StringComparison.Ordinal))
                        {
                            throw new InvalidOperationException(
                                "An unowned model measure already uses generated name '" +
                                definition.Name + "'.");
                        }

                        creates.Add(new MeasureUpsert(definition, collision));
                    }
                    else
                    {
                        creates.Add(new MeasureUpsert(definition, before: null));
                    }

                    continue;
                }

                if (!liveByName.TryGetValue(
                        definition.Name,
                        out LiveModelMeasureSnapshot? live))
                {
                    throw new InvalidOperationException(
                        "A desired owned model measure is missing during semantic replay.");
                }

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

            var desiredNames = new HashSet<string>(
                definitions.Select(definition => definition.Name),
                StringComparer.OrdinalIgnoreCase);
            List<PivotPlusOwnedArtifact> deletionArtifacts = active
                .Where(artifact => !desiredNames.Contains(artifact.ArtifactId))
                .ToList();
            List<LiveModelMeasureSnapshot> deletes = deletionArtifacts
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
            transitions.AddRange(deletionArtifacts.Select(item =>
                new PivotPlusSemanticArtifactTransition
                {
                    Kind = PivotPlusArtifactKind.Measure,
                    ArtifactId = item.ArtifactId,
                    Operation = PivotPlusSemanticArtifactOperation.Delete,
                    BeforeLiveFingerprint = item.Fingerprint,
                    PlannedDefinitionFingerprint =
                        PivotModelMeasureCanonical.CreateDeleteDefinitionFingerprint(
                            item.ArtifactId,
                            item.Fingerprint)
                }));

            string planFingerprint = PivotModelMeasureCanonical.CreateArtifactPlanFingerprint(
                definitions,
                transitions);
            if (existingPending != null)
            {
                DemandSameArtifactTransitionSlice(existingPending, transitions);
            }

            return new ArtifactMeasureApplyPlan(
                existingPending == null
                    ? "apply_" + Guid.NewGuid().ToString("N")
                    : existingPending.ApplyId,
                planFingerprint,
                transitions,
                creates,
                updates,
                deletes,
                deletionArtifacts.Select(artifact => artifact.ArtifactId).ToList(),
                deletionArtifacts.Count);
        }

        private PivotModelMeasureArtifactPreparedUndo CreateArtifactPreparedUndo(
            BoundModelMeasureTarget target,
            ModelMeasureWorkbookSnapshot current,
            PivotModelMeasureArtifactUndoContribution contribution)
        {
            var upsertSteps = new List<PivotMutationStep>();
            var deleteSteps = new List<PivotMutationStep>();
            var currentByName = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var priorByName = contribution.Before.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var recreated = new Dictionary<string, LiveModelMeasureSnapshot>(
                StringComparer.OrdinalIgnoreCase);

            IReadOnlyList<LiveModelMeasureSnapshot> restoreCandidates = OrderRestores(
                contribution.Before.Measures.Where(prior =>
                    contribution.BeforeOwnedArtifacts.Any(artifact => string.Equals(
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
                        else if (recreated.TryGetValue(
                                     prior.Name,
                                     out LiveModelMeasureSnapshot? value))
                        {
                            gateway.DeleteMeasure(target, value);
                        }
                    }));
            }

            IReadOnlyList<LiveModelMeasureSnapshot> createdMeasures = OrderDeletes(
                current.Measures.Where(measure =>
                    contribution.AfterOwnedArtifacts.Any(artifact => string.Equals(
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

            return new PivotModelMeasureArtifactPreparedUndo(
                target,
                current,
                new ReadOnlyCollection<PivotMutationStep>(upsertSteps),
                new ReadOnlyCollection<PivotMutationStep>(deleteSteps),
                () => gateway.Refresh(target),
                () => DemandArtifactUndoFinal(
                    contribution,
                    current,
                    gateway.Capture(target)),
                () => DemandExactArtifactSnapshot(
                    current,
                    gateway.Capture(target),
                    "combined measure artifact Undo rollback"));
        }

        private static void VerifyArtifactFinal(
            ModelMeasureWorkbookSnapshot before,
            ModelMeasureWorkbookSnapshot after,
            ArtifactMeasureApplyPlan plan,
            IReadOnlyDictionary<string, DesiredModelMeasure> definitionsById,
            IReadOnlyDictionary<string, LiveModelMeasureSnapshot> upsertReceipts)
        {
            var afterByName = after.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (DesiredModelMeasure definition in definitionsById.Values)
            {
                if (!afterByName.TryGetValue(
                        definition.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        live.Description,
                        definition.DescriptionMarker,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        live.AssociatedTableName,
                        definition.HomeTableName,
                        StringComparison.Ordinal))
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
                throw new InvalidOperationException(
                    "Excel retained a model measure scheduled for deletion.");
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
                    "The workbook model-measure inventory changed during the combined Apply.");
            }

            var changed = new HashSet<string>(
                plan.Upserts.Select(item => item.Definition.Name)
                    .Concat(plan.DeletedNames),
                StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot prior in before.Measures.Where(
                         measure => !changed.Contains(measure.Name)))
            {
                if (!afterByName.TryGetValue(
                        prior.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        prior.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "An unrelated model measure changed during the combined Apply.");
                }
            }

            DemandNoPostRefreshUnsafeRelationships(
                after,
                changed,
                new HashSet<string>(
                    definitionsById.Values.Select(definition => definition.Name),
                    StringComparer.OrdinalIgnoreCase),
                "combined Apply");
            DemandExactOtherPivotUsages(
                before,
                after,
                "Another PivotTable changed during the combined measure artifact Apply.");
            if (after.PivotUsages.Count != before.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model PivotTable inventory changed during the combined Apply.");
            }
        }

        private static void DemandExactArtifactSnapshot(
            ModelMeasureWorkbookSnapshot expected,
            ModelMeasureWorkbookSnapshot actual,
            string operation)
        {
            if (expected.Measures.Count != actual.Measures.Count ||
                expected.PivotUsages.Count != actual.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    operation + " did not restore the exact measure artifact inventory.");
            }

            var actualByName = actual.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            foreach (LiveModelMeasureSnapshot prior in expected.Measures)
            {
                if (!actualByName.TryGetValue(
                        prior.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        prior.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        operation + " did not restore an exact model measure.");
                }
            }

            DemandExactOtherPivotUsages(
                expected,
                actual,
                operation + " changed a sibling Data Model PivotTable.");
        }

        private static void DemandArtifactUndoStartOrIntermediate(
            PivotModelMeasureArtifactUndoContribution contribution,
            ModelMeasureWorkbookSnapshot current)
        {
            var beforeByName = contribution.Before.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var afterByName = contribution.After.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var currentByName = current.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            var involved = new HashSet<string>(
                contribution.BeforeOwnedArtifacts.Select(artifact => artifact.ArtifactId)
                    .Concat(contribution.AfterOwnedArtifacts.Select(artifact => artifact.ArtifactId)),
                StringComparer.OrdinalIgnoreCase);

            foreach (LiveModelMeasureSnapshot expected in contribution.After.Measures.Where(
                         measure => !involved.Contains(measure.Name)))
            {
                if (!currentByName.TryGetValue(
                        expected.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        expected.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "An unrelated model measure changed after the combined Apply.");
                }
            }

            if (current.Measures.Any(measure =>
                    !involved.Contains(measure.Name) && !afterByName.ContainsKey(measure.Name)))
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "A model measure appeared outside the combined Apply.");
            }

            foreach (string name in involved)
            {
                beforeByName.TryGetValue(name, out LiveModelMeasureSnapshot? before);
                afterByName.TryGetValue(name, out LiveModelMeasureSnapshot? after);
                if (!currentByName.TryGetValue(name, out LiveModelMeasureSnapshot? live))
                {
                    if (before != null && after != null)
                    {
                        throw new PivotModelMeasureUndoUnavailableException(
                            "A changed owned measure is missing from the model.");
                    }

                    continue;
                }

                bool matchesBefore = before != null && string.Equals(
                    live.LiveFingerprint,
                    before.LiveFingerprint,
                    StringComparison.Ordinal);
                bool matchesAfter = after != null && string.Equals(
                    live.LiveFingerprint,
                    after.LiveFingerprint,
                    StringComparison.Ordinal);
                if (!matchesBefore && !matchesAfter)
                {
                    throw new PivotModelMeasureUndoUnavailableException(
                        "An owned measure no longer matches an exact combined Apply state.");
                }
            }

            try
            {
                DemandExactOtherPivotUsages(
                    contribution.After,
                    current,
                    "Another PivotTable changed after the combined Apply.");
            }
            catch (InvalidOperationException exception)
            {
                throw new PivotModelMeasureUndoUnavailableException(exception.Message);
            }

            if (current.PivotUsages.Count != contribution.After.PivotUsages.Count)
            {
                throw new PivotModelMeasureUndoUnavailableException(
                    "The workbook Data Model PivotTable inventory changed after the combined Apply.");
            }
        }

        private static void DemandArtifactUndoFinal(
            PivotModelMeasureArtifactUndoContribution contribution,
            ModelMeasureWorkbookSnapshot undoStart,
            ModelMeasureWorkbookSnapshot actual)
        {
            var actualByName = actual.Measures.ToDictionary(
                measure => measure.Name,
                StringComparer.OrdinalIgnoreCase);
            if (actual.Measures.Count != contribution.Before.Measures.Count)
            {
                throw new InvalidOperationException(
                    "Combined Undo did not restore the exact prior measure inventory.");
            }

            foreach (LiveModelMeasureSnapshot prior in contribution.Before.Measures)
            {
                if (!actualByName.TryGetValue(
                        prior.Name,
                        out LiveModelMeasureSnapshot? live) ||
                    !string.Equals(
                        prior.LiveFingerprint,
                        live.LiveFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Combined Undo did not restore an exact prior model measure.");
                }
            }

            var beforeOwned = contribution.BeforeOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            var afterOwned = contribution.AfterOwnedArtifacts.ToDictionary(
                artifact => artifact.ArtifactId,
                StringComparer.OrdinalIgnoreCase);
            DemandNoPostRefreshUnsafeRelationships(
                actual,
                CreateUndoChangingNames(beforeOwned, afterOwned),
                new HashSet<string>(beforeOwned.Keys, StringComparer.OrdinalIgnoreCase),
                "combined Undo");
            DemandExactOtherPivotUsages(
                undoStart,
                actual,
                "Combined Undo changed a sibling Data Model PivotTable.");
            if (actual.PivotUsages.Count != undoStart.PivotUsages.Count)
            {
                throw new InvalidOperationException(
                    "The workbook Data Model PivotTable inventory changed during combined Undo.");
            }
        }

        private static void DemandSameArtifactTransitionSlice(
            PivotPlusPendingSemanticApplyMetadata existing,
            IEnumerable<PivotPlusSemanticArtifactTransition> proposed)
        {
            IReadOnlyList<PivotPlusSemanticArtifactTransition> existingMeasures = existing.Transitions
                .Where(transition => transition.Kind == PivotPlusArtifactKind.Measure)
                .ToList();
            if (!TransitionListsEqual(existingMeasures, proposed))
            {
                throw new InvalidOperationException(
                    "The measure artifacts do not match the exact measure portion of the pending combined Apply.");
            }
        }

        private static PivotPlusSemanticArtifactTransition CloneTransition(
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

        private sealed class ArtifactMeasureApplyPlan
        {
            public ArtifactMeasureApplyPlan(
                string applyId,
                string planFingerprint,
                IReadOnlyList<PivotPlusSemanticArtifactTransition> transitions,
                IReadOnlyList<MeasureUpsert> creates,
                IReadOnlyList<MeasureUpsert> updates,
                IReadOnlyList<LiveModelMeasureSnapshot> deletes,
                IReadOnlyList<string> deletedNames,
                int deleteCount)
            {
                ApplyId = applyId;
                PlanFingerprint = planFingerprint;
                Transitions = transitions;
                Creates = creates;
                Updates = updates;
                Upserts = creates.Concat(updates)
                    .OrderBy(item => item.Definition.CreationOrder)
                    .ToList();
                Deletes = deletes;
                DeletedNames = deletedNames;
                DeleteCount = deleteCount;
            }

            public string ApplyId { get; }

            public string PlanFingerprint { get; }

            public IReadOnlyList<PivotPlusSemanticArtifactTransition> Transitions { get; }

            public IReadOnlyList<MeasureUpsert> Creates { get; }

            public IReadOnlyList<MeasureUpsert> Updates { get; }

            public IReadOnlyList<MeasureUpsert> Upserts { get; }

            public IReadOnlyList<LiveModelMeasureSnapshot> Deletes { get; }

            public IReadOnlyList<string> DeletedNames { get; }

            public int DeleteCount { get; }

            public bool IsNoChange => Transitions.Count == 0;
        }
    }
}
