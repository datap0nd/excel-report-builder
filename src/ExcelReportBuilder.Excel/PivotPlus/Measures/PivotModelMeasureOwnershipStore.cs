using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.Measures
{
    internal sealed class PivotModelMeasureOwnershipSession
    {
        public PivotModelMeasureOwnershipSession(
            PivotPlusWorkbookMetadata baseMetadata,
            PivotPlusPendingSemanticApplyMetadata pending,
            bool resumed)
        {
            BaseMetadata = baseMetadata;
            Pending = pending;
            Resumed = resumed;
        }

        public PivotPlusWorkbookMetadata BaseMetadata { get; }

        public PivotPlusPendingSemanticApplyMetadata Pending { get; }

        public bool Resumed { get; }
    }

    internal interface IPivotModelMeasureOwnershipStore
    {
        PivotPlusWorkbookMetadata ReadBase(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            out PivotPlusPendingSemanticApplyMetadata? existingPending);

        PivotModelMeasureOwnershipSession Journal(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotPlusPendingSemanticApplyMetadata pending);

        void Commit(
            object workbook,
            PivotModelMeasureOwnershipSession session,
            IReadOnlyList<PivotPlusOwnedArtifact> measures,
            PivotPlusUndoMetadata? undo);

        void RestoreBase(object workbook, PivotModelMeasureOwnershipSession session);
    }

    internal sealed class PivotModelMeasureOwnershipStore : IPivotModelMeasureOwnershipStore
    {
        private readonly PivotPlusWorkbookMetadataStore store;

        public PivotModelMeasureOwnershipStore()
            : this(new PivotPlusWorkbookMetadataStore())
        {
        }

        internal PivotModelMeasureOwnershipStore(PivotPlusWorkbookMetadataStore store)
        {
            this.store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public PivotPlusWorkbookMetadata ReadBase(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            out PivotPlusPendingSemanticApplyMetadata? existingPending)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (target == null) throw new ArgumentNullException(nameof(target));

            PivotPlusWorkbookMetadata? byId = store.Load((dynamic)workbook, setupId);
            PivotPlusWorkbookMetadata? byTarget = store.LoadForTarget(
                (dynamic)workbook,
                target.WorksheetName,
                target.PivotTableName);
            if (byTarget != null &&
                !string.Equals(byTarget.SetupId, setupId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Another PivotTable+ setup already owns the selected target metadata.");
            }

            PivotPlusWorkbookMetadata current = byId ?? new PivotPlusWorkbookMetadata
            {
                SetupId = setupId,
                TargetWorksheetName = target.WorksheetName,
                TargetPivotTableName = target.PivotTableName
            };
            DemandTarget(current, target);
            if (current.RecoveryPhase != PivotPlusRecoveryPhase.None)
            {
                throw new InvalidOperationException(
                    "The selected PivotTable has an unfinished Data Model conversion recovery.");
            }

            existingPending = current.PendingSemanticApply == null
                ? null
                : Clone(current.PendingSemanticApply);
            PivotPlusWorkbookMetadata result = Clone(current);
            result.PendingSemanticApply = null;
            return result;
        }

        public PivotModelMeasureOwnershipSession Journal(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotPlusPendingSemanticApplyMetadata pending)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (pending == null) throw new ArgumentNullException(nameof(pending));

            PivotPlusWorkbookMetadata baseMetadata = ReadBase(
                workbook,
                setupId,
                target,
                out PivotPlusPendingSemanticApplyMetadata? existingPending);
            if (existingPending != null)
            {
                if (!PendingEquals(existingPending, pending))
                {
                    throw new InvalidOperationException(
                        "A different PivotTable+ semantic Apply is already pending for this setup.");
                }

                return new PivotModelMeasureOwnershipSession(
                    baseMetadata,
                    Clone(pending),
                    resumed: true);
            }

            PivotPlusWorkbookMetadata journaled = Clone(baseMetadata);
            journaled.PendingSemanticApply = Clone(pending);
            store.Save((dynamic)workbook, journaled);
            return new PivotModelMeasureOwnershipSession(
                baseMetadata,
                Clone(pending),
                resumed: false);
        }

        public void Commit(
            object workbook,
            PivotModelMeasureOwnershipSession session,
            IReadOnlyList<PivotPlusOwnedArtifact> measures,
            PivotPlusUndoMetadata? undo)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (measures == null) throw new ArgumentNullException(nameof(measures));

            DemandExactJournal(workbook, session);

            PivotPlusWorkbookMetadata committed = Clone(session.BaseMetadata);
            var artifacts = committed.Artifacts
                .Where(artifact => artifact.Kind != PivotPlusArtifactKind.Measure)
                .Select(Clone)
                .Concat(measures.Select(Clone))
                .ToList();
            committed.Artifacts = artifacts;
            committed.PendingSemanticApply = null;
            committed.Undo = undo == null ? null : Clone(undo);
            store.Save((dynamic)workbook, committed);
        }

        internal void CommitSemantic(
            object workbook,
            PivotModelMeasureOwnershipSession session,
            IReadOnlyList<PivotPlusOwnedArtifact> semanticArtifacts,
            PivotPlusUndoMetadata? undo)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (semanticArtifacts == null)
            {
                throw new ArgumentNullException(nameof(semanticArtifacts));
            }

            if (semanticArtifacts.Any(artifact =>
                    artifact.Kind != PivotPlusArtifactKind.Measure &&
                    artifact.Kind != PivotPlusArtifactKind.NamedSet))
            {
                throw new ArgumentException(
                    "A combined semantic commit can contain only Measure and NamedSet artifacts.",
                    nameof(semanticArtifacts));
            }

            DemandExactJournal(workbook, session);
            PivotPlusWorkbookMetadata committed = Clone(session.BaseMetadata);
            committed.Artifacts = committed.Artifacts
                .Where(artifact =>
                    artifact.Kind != PivotPlusArtifactKind.Measure &&
                    artifact.Kind != PivotPlusArtifactKind.NamedSet)
                .Select(Clone)
                .Concat(semanticArtifacts.Select(Clone))
                .ToList();
            committed.PendingSemanticApply = null;
            committed.Undo = undo == null ? null : Clone(undo);
            store.Save((dynamic)workbook, committed);
        }

        internal void CommitSemanticState(
            object workbook,
            PivotModelMeasureOwnershipSession session,
            PivotPlusWorkbookMetadata desiredState)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (desiredState == null)
            {
                throw new ArgumentNullException(nameof(desiredState));
            }

            DemandExactJournal(workbook, session);
            PivotPlusWorkbookMetadata committed = Clone(desiredState);
            committed.PendingSemanticApply = null;
            committed.Undo = null;
            store.Save((dynamic)workbook, committed);
        }

        public void RestoreBase(object workbook, PivotModelMeasureOwnershipSession session)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (session == null) throw new ArgumentNullException(nameof(session));
            DemandExactJournal(workbook, session);
            PivotPlusWorkbookMetadata restored = Clone(session.BaseMetadata);
            restored.PendingSemanticApply = null;
            store.Save((dynamic)workbook, restored);
        }

        private void DemandExactJournal(
            object workbook,
            PivotModelMeasureOwnershipSession session)
        {
            PivotPlusWorkbookMetadata current = store.Load(
                (dynamic)workbook,
                session.BaseMetadata.SetupId) ??
                throw new InvalidOperationException(
                    "The pending semantic ownership journal is missing.");
            if (current.PendingSemanticApply == null ||
                !PendingEquals(current.PendingSemanticApply, session.Pending))
            {
                throw new InvalidOperationException(
                    "The pending semantic ownership journal changed before finalization.");
            }

            PivotPlusWorkbookMetadata currentBase = Clone(current);
            currentBase.PendingSemanticApply = null;
            if (!BaseEquals(currentBase, session.BaseMetadata))
            {
                throw new InvalidOperationException(
                    "Active PivotTable+ ownership changed while the semantic Apply was pending.");
            }
        }

        private static void DemandTarget(
            PivotPlusWorkbookMetadata metadata,
            PivotTargetIdentity target)
        {
            if (!string.Equals(
                    metadata.TargetWorksheetName,
                    target.WorksheetName,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    metadata.TargetPivotTableName,
                    target.PivotTableName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The setup metadata belongs to a different PivotTable target.");
            }
        }

        private static bool PendingEquals(
            PivotPlusPendingSemanticApplyMetadata left,
            PivotPlusPendingSemanticApplyMetadata right)
        {
            if (!string.Equals(left.ApplyId, right.ApplyId, StringComparison.Ordinal) ||
                !string.Equals(left.PlanFingerprint, right.PlanFingerprint, StringComparison.Ordinal) ||
                !string.Equals(left.BeforePivotFingerprint, right.BeforePivotFingerprint, StringComparison.Ordinal) ||
                !string.Equals(left.ExpectedPivotFingerprint, right.ExpectedPivotFingerprint, StringComparison.Ordinal) ||
                left.Transitions.Count != right.Transitions.Count)
            {
                return false;
            }

            List<PivotPlusSemanticArtifactTransition> orderedLeft = Order(left.Transitions);
            List<PivotPlusSemanticArtifactTransition> orderedRight = Order(right.Transitions);
            for (int index = 0; index < orderedLeft.Count; index++)
            {
                PivotPlusSemanticArtifactTransition first = orderedLeft[index];
                PivotPlusSemanticArtifactTransition second = orderedRight[index];
                if (first.Kind != second.Kind ||
                    first.Operation != second.Operation ||
                    !string.Equals(first.ArtifactId, second.ArtifactId, StringComparison.Ordinal) ||
                    !string.Equals(first.BeforeLiveFingerprint, second.BeforeLiveFingerprint, StringComparison.Ordinal) ||
                    !string.Equals(first.PlannedDefinitionFingerprint, second.PlannedDefinitionFingerprint, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool BaseEquals(
            PivotPlusWorkbookMetadata left,
            PivotPlusWorkbookMetadata right)
        {
            return string.Equals(left.SetupId, right.SetupId, StringComparison.Ordinal) &&
                   string.Equals(
                       left.TargetWorksheetName,
                       right.TargetWorksheetName,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       left.TargetPivotTableName,
                       right.TargetPivotTableName,
                       StringComparison.Ordinal) &&
                   left.RecoveryPhase == right.RecoveryPhase &&
                   string.Equals(
                       left.TargetAnchorAddress,
                       right.TargetAnchorAddress,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       left.StagingStateFingerprint,
                       right.StagingStateFingerprint,
                       StringComparison.Ordinal) &&
                   ArtifactListsEqual(left.Artifacts, right.Artifacts) &&
                   UndoEquals(left.Undo, right.Undo);
        }

        private static bool ArtifactListsEqual(
            IEnumerable<PivotPlusOwnedArtifact> left,
            IEnumerable<PivotPlusOwnedArtifact> right)
        {
            List<PivotPlusOwnedArtifact> orderedLeft = left
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ArtifactId, StringComparer.Ordinal)
                .ToList();
            List<PivotPlusOwnedArtifact> orderedRight = right
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ArtifactId, StringComparer.Ordinal)
                .ToList();
            if (orderedLeft.Count != orderedRight.Count) return false;
            for (int index = 0; index < orderedLeft.Count; index++)
            {
                if (orderedLeft[index].Kind != orderedRight[index].Kind ||
                    !string.Equals(
                        orderedLeft[index].ArtifactId,
                        orderedRight[index].ArtifactId,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        orderedLeft[index].Fingerprint,
                        orderedRight[index].Fingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool UndoEquals(
            PivotPlusUndoMetadata? left,
            PivotPlusUndoMetadata? right)
        {
            if (left == null || right == null) return left == null && right == null;
            if (!string.Equals(left.ApplyId, right.ApplyId, StringComparison.Ordinal) ||
                !string.Equals(
                    left.BeforePivotFingerprint,
                    right.BeforePivotFingerprint,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    left.AfterPivotFingerprint,
                    right.AfterPivotFingerprint,
                    StringComparison.Ordinal) ||
                !ArtifactListsEqual(left.CreatedArtifacts, right.CreatedArtifacts) ||
                left.PreviousFieldPlacements.Count != right.PreviousFieldPlacements.Count)
            {
                return false;
            }

            List<PivotPlusUndoFieldPlacement> orderedLeft = left.PreviousFieldPlacements
                .OrderBy(value => value.Area)
                .ThenBy(value => value.Position)
                .ThenBy(value => value.FieldFingerprint, StringComparer.Ordinal)
                .ToList();
            List<PivotPlusUndoFieldPlacement> orderedRight = right.PreviousFieldPlacements
                .OrderBy(value => value.Area)
                .ThenBy(value => value.Position)
                .ThenBy(value => value.FieldFingerprint, StringComparer.Ordinal)
                .ToList();
            for (int index = 0; index < orderedLeft.Count; index++)
            {
                if (orderedLeft[index].Area != orderedRight[index].Area ||
                    orderedLeft[index].Position != orderedRight[index].Position ||
                    !string.Equals(
                        orderedLeft[index].FieldFingerprint,
                        orderedRight[index].FieldFingerprint,
                        StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static List<PivotPlusSemanticArtifactTransition> Order(
            IEnumerable<PivotPlusSemanticArtifactTransition> values)
        {
            return values
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ArtifactId, StringComparer.Ordinal)
                .ToList();
        }

        private static PivotPlusWorkbookMetadata Clone(PivotPlusWorkbookMetadata source)
        {
            return new PivotPlusWorkbookMetadata
            {
                SchemaVersion = PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                SetupId = source.SetupId,
                TargetWorksheetName = source.TargetWorksheetName,
                TargetPivotTableName = source.TargetPivotTableName,
                RecoveryPhase = source.RecoveryPhase,
                TargetAnchorAddress = source.TargetAnchorAddress,
                StagingStateFingerprint = source.StagingStateFingerprint,
                Artifacts = source.Artifacts.Select(Clone).ToList(),
                PendingSemanticApply = source.PendingSemanticApply == null
                    ? null
                    : Clone(source.PendingSemanticApply),
                Undo = source.Undo == null ? null : Clone(source.Undo)
            };
        }

        private static PivotPlusPendingSemanticApplyMetadata Clone(
            PivotPlusPendingSemanticApplyMetadata source)
        {
            return new PivotPlusPendingSemanticApplyMetadata
            {
                ApplyId = source.ApplyId,
                PlanFingerprint = source.PlanFingerprint,
                BeforePivotFingerprint = source.BeforePivotFingerprint,
                ExpectedPivotFingerprint = source.ExpectedPivotFingerprint,
                Transitions = source.Transitions.Select(Clone).ToList()
            };
        }

        private static PivotPlusSemanticArtifactTransition Clone(
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

        private static PivotPlusOwnedArtifact Clone(PivotPlusOwnedArtifact source)
        {
            return new PivotPlusOwnedArtifact
            {
                Kind = source.Kind,
                ArtifactId = source.ArtifactId,
                Fingerprint = source.Fingerprint
            };
        }

        private static PivotPlusUndoMetadata Clone(PivotPlusUndoMetadata source)
        {
            return new PivotPlusUndoMetadata
            {
                ApplyId = source.ApplyId,
                BeforePivotFingerprint = source.BeforePivotFingerprint,
                AfterPivotFingerprint = source.AfterPivotFingerprint,
                CreatedArtifacts = source.CreatedArtifacts.Select(Clone).ToList(),
                PreviousFieldPlacements = source.PreviousFieldPlacements.Select(value =>
                    new PivotPlusUndoFieldPlacement
                    {
                        FieldFingerprint = value.FieldFingerprint,
                        Area = value.Area,
                        Position = value.Position
                    }).ToList()
            };
        }
    }
}
