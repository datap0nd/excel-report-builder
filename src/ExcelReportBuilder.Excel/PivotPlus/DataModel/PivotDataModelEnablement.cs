using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.DataModel
{
    public enum PivotDataModelEnablementStage
    {
        BindTarget,
        PersistWorkbookIdentity,
        PreflightOwnership,
        InspectSource,
        CaptureState,
        CreateModelArtifacts,
        CreateStagingPivot,
        RestoreStagingState,
        VerifyStagingPivot,
        PrepareReplacement,
        ReplaceOriginal,
        VerifyReplacement,
        CleanupStaging,
        PersistOwnership,
        Complete
    }

    public sealed class ClassicPivotSourceDescriptor
    {
        public ClassicPivotSourceDescriptor(string workbookObjectName, PivotPlusWorkbookObjectKind objectKind)
        {
            if (string.IsNullOrWhiteSpace(workbookObjectName))
            {
                throw new ArgumentException("A workbook source object name is required.", nameof(workbookObjectName));
            }

            if (!Enum.IsDefined(typeof(PivotPlusWorkbookObjectKind), objectKind))
            {
                throw new ArgumentOutOfRangeException(nameof(objectKind));
            }

            WorkbookObjectName = workbookObjectName;
            ObjectKind = objectKind;
        }

        internal ClassicPivotSourceDescriptor(object nativeRange, string canonicalReference)
        {
            NativeRange = nativeRange ?? throw new ArgumentNullException(nameof(nativeRange));
            CanonicalReference = canonicalReference ??
                throw new ArgumentNullException(nameof(canonicalReference));
            WorkbookObjectName = string.Empty;
            ObjectKind = PivotPlusWorkbookObjectKind.NamedRange;
            RequiresOwnedWorkbookName = true;
        }

        public string WorkbookObjectName { get; }

        public PivotPlusWorkbookObjectKind ObjectKind { get; }

        /// <summary>
        /// True only for a bounded same-workbook raw range that must be exposed
        /// to Excel.CurrentWorkbook through a generated workbook-scoped name.
        /// </summary>
        public bool RequiresOwnedWorkbookName { get; }

        internal object? NativeRange { get; }

        internal string? CanonicalReference { get; }
    }

    public sealed class PivotOwnedWorkbookNameArtifact
    {
        internal PivotOwnedWorkbookNameArtifact(
            string name,
            string referenceFingerprint,
            string canonicalReference,
            object nativeName)
        {
            Name = name;
            ReferenceFingerprint = referenceFingerprint;
            CanonicalReference = canonicalReference;
            NativeName = nativeName;
        }

        public PivotPlusArtifactKind Kind => PivotPlusArtifactKind.WorkbookName;

        public string Name { get; }

        public string ReferenceFingerprint { get; }

        internal string CanonicalReference { get; }

        internal object NativeName { get; }
    }

    public sealed class PivotTemporaryWorksheetArtifact
    {
        internal PivotTemporaryWorksheetArtifact(
            string name,
            string purpose,
            string fingerprint,
            string targetAnchorAddress = "")
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Purpose = purpose ?? throw new ArgumentNullException(nameof(purpose));
            Fingerprint = fingerprint ?? throw new ArgumentNullException(nameof(fingerprint));
            TargetAnchorAddress = targetAnchorAddress ??
                throw new ArgumentNullException(nameof(targetAnchorAddress));
        }

        public PivotPlusArtifactKind Kind => PivotPlusArtifactKind.TemporaryWorksheet;

        public string Name { get; }

        public string Fingerprint { get; }

        internal string Purpose { get; }

        internal string TargetAnchorAddress { get; }
    }

    public sealed class PivotTemporaryPivotTableArtifact
    {
        internal PivotTemporaryPivotTableArtifact(
            string setupId,
            string name,
            string fingerprint,
            string targetWorksheetName,
            string targetPivotTableName,
            string targetAnchorAddress,
            string connectionName,
            string modelTableName)
        {
            SetupId = setupId ?? throw new ArgumentNullException(nameof(setupId));
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Fingerprint = fingerprint ?? throw new ArgumentNullException(nameof(fingerprint));
            TargetWorksheetName = targetWorksheetName ??
                throw new ArgumentNullException(nameof(targetWorksheetName));
            TargetPivotTableName = targetPivotTableName ??
                throw new ArgumentNullException(nameof(targetPivotTableName));
            TargetAnchorAddress = targetAnchorAddress ??
                throw new ArgumentNullException(nameof(targetAnchorAddress));
            ConnectionName = connectionName ?? throw new ArgumentNullException(nameof(connectionName));
            ModelTableName = modelTableName ?? throw new ArgumentNullException(nameof(modelTableName));
        }

        public PivotPlusArtifactKind Kind => PivotPlusArtifactKind.TemporaryPivotTable;

        internal string SetupId { get; }

        public string Name { get; }

        public string Fingerprint { get; }

        internal string TargetWorksheetName { get; }

        internal string TargetPivotTableName { get; }

        internal string TargetAnchorAddress { get; }

        internal string ConnectionName { get; }

        internal string ModelTableName { get; }
    }

    /// <summary>
    /// An opaque, bounded snapshot captured by the Excel gateway before any
    /// generated query, connection, worksheet, or PivotTable is created.
    /// </summary>
    public sealed class PivotNativeStateSnapshot
    {
        internal PivotNativeStateSnapshot(
            string worksheetName,
            string pivotTableName,
            string anchorAddress,
            string fingerprint,
            object nativeState)
        {
            WorksheetName = worksheetName ?? throw new ArgumentNullException(nameof(worksheetName));
            PivotTableName = pivotTableName ?? throw new ArgumentNullException(nameof(pivotTableName));
            AnchorAddress = anchorAddress ?? throw new ArgumentNullException(nameof(anchorAddress));
            Fingerprint = fingerprint ?? throw new ArgumentNullException(nameof(fingerprint));
            NativeState = nativeState ?? throw new ArgumentNullException(nameof(nativeState));
        }

        public string WorksheetName { get; }

        public string PivotTableName { get; }

        public string AnchorAddress { get; }

        public string Fingerprint { get; }

        internal object NativeState { get; }
    }

    /// <summary>
    /// Exact receipt for the query and model connection generated by one
    /// conversion. Cleanup may delete only objects matching this receipt.
    /// </summary>
    public sealed class PivotDataModelArtifacts
    {
        internal PivotDataModelArtifacts(
            string queryName,
            string connectionName,
            string modelTableName,
            string queryFormula,
            string queryFingerprint,
            string connectionFingerprint,
            object nativeConnection,
            PivotOwnedWorkbookNameArtifact? ownedWorkbookName = null,
            IReadOnlyList<PivotTemporaryWorksheetArtifact>? temporaryWorksheets = null,
            object? nativeDataModelConnection = null,
            PivotTemporaryPivotTableArtifact? temporaryPivotTable = null)
        {
            QueryName = queryName;
            ConnectionName = connectionName;
            ModelTableName = modelTableName;
            QueryFormula = queryFormula;
            QueryFingerprint = queryFingerprint;
            ConnectionFingerprint = connectionFingerprint;
            NativeConnection = nativeConnection;
            OwnedWorkbookName = ownedWorkbookName;
            TemporaryWorksheets = temporaryWorksheets ??
                Array.Empty<PivotTemporaryWorksheetArtifact>();
            NativeDataModelConnection = nativeDataModelConnection ?? nativeConnection;
            TemporaryPivotTable = temporaryPivotTable;
        }

        public string QueryName { get; }

        public string ConnectionName { get; }

        public string ModelTableName { get; }

        internal string QueryFormula { get; }

        public string QueryFingerprint { get; }

        public string ConnectionFingerprint { get; }

        public PivotOwnedWorkbookNameArtifact? OwnedWorkbookName { get; }

        public IReadOnlyList<PivotTemporaryWorksheetArtifact> TemporaryWorksheets { get; }

        public PivotTemporaryPivotTableArtifact? TemporaryPivotTable { get; }

        internal object NativeConnection { get; }

        internal object NativeDataModelConnection { get; }
    }

    public sealed class PivotDataModelArtifactPlan
    {
        internal PivotDataModelArtifactPlan(
            string queryName,
            string connectionName,
            string modelTableName,
            string queryFormula,
            string queryFingerprint,
            string connectionFingerprint,
            string? workbookName,
            string? workbookNameFingerprint,
            string? requestedWorkbookNameReference,
            IReadOnlyList<PivotTemporaryWorksheetArtifact> temporaryWorksheets,
            PivotTemporaryPivotTableArtifact? temporaryPivotTable = null)
        {
            QueryName = queryName;
            ConnectionName = connectionName;
            ModelTableName = modelTableName;
            QueryFormula = queryFormula;
            QueryFingerprint = queryFingerprint;
            ConnectionFingerprint = connectionFingerprint;
            WorkbookName = workbookName;
            WorkbookNameFingerprint = workbookNameFingerprint;
            RequestedWorkbookNameReference = requestedWorkbookNameReference;
            TemporaryWorksheets = temporaryWorksheets;
            TemporaryPivotTable = temporaryPivotTable;
        }

        public string QueryName { get; }
        public string ConnectionName { get; }
        public string ModelTableName { get; }
        internal string QueryFormula { get; }
        public string QueryFingerprint { get; }
        public string ConnectionFingerprint { get; }
        public string? WorkbookName { get; }
        public string? WorkbookNameFingerprint { get; }
        internal string? RequestedWorkbookNameReference { get; }
        public IReadOnlyList<PivotTemporaryWorksheetArtifact> TemporaryWorksheets { get; }
        public PivotTemporaryPivotTableArtifact? TemporaryPivotTable { get; }
    }

    public sealed class PivotStagedDataModelPivot
    {
        internal PivotStagedDataModelPivot(
            string worksheetName,
            string pivotTableName,
            object nativeWorksheet,
            object nativePivotTable,
            object nativePivotCache,
            PivotTemporaryWorksheetArtifact? stagingWorksheet = null,
            PivotTemporaryWorksheetArtifact? formatBackupWorksheet = null,
            PivotTemporaryPivotTableArtifact? temporaryTargetPivotTable = null)
        {
            WorksheetName = worksheetName;
            PivotTableName = pivotTableName;
            NativeWorksheet = nativeWorksheet;
            NativePivotTable = nativePivotTable;
            NativePivotCache = nativePivotCache;
            StagingWorksheet = stagingWorksheet;
            FormatBackupWorksheet = formatBackupWorksheet;
            TemporaryTargetPivotTable = temporaryTargetPivotTable;
        }

        public string WorksheetName { get; }

        public string PivotTableName { get; }

        internal object NativeWorksheet { get; }

        internal object NativePivotTable { get; }

        internal object NativePivotCache { get; }

        internal PivotTemporaryWorksheetArtifact? StagingWorksheet { get; }

        internal PivotTemporaryWorksheetArtifact? FormatBackupWorksheet { get; }

        internal PivotTemporaryPivotTableArtifact? TemporaryTargetPivotTable { get; }
    }

    public sealed class PivotDataModelEnablementResult
    {
        internal PivotDataModelEnablementResult(
            PivotTargetIdentity target,
            ClassicPivotSourceDescriptor source,
            PivotDataModelArtifacts artifacts)
        {
            Target = target;
            Source = source;
            Artifacts = artifacts;
        }

        public PivotTargetIdentity Target { get; }

        public ClassicPivotSourceDescriptor Source { get; }

        public PivotDataModelArtifacts Artifacts { get; }
    }

    public sealed class PivotPendingDataModelRecovery
    {
        internal PivotPendingDataModelRecovery(
            PivotTargetIdentity target,
            PivotDataModelArtifacts artifacts)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
            Artifacts = artifacts ?? throw new ArgumentNullException(nameof(artifacts));
        }

        public PivotTargetIdentity Target { get; }

        public PivotDataModelArtifacts Artifacts { get; }
    }

    public sealed class PivotDataModelEnablementException : Exception
    {
        internal PivotDataModelEnablementException(
            PivotDataModelEnablementStage failedStage,
            bool rollbackCompleted,
            bool recoveryRequired,
            Exception innerException)
            : base(
                (recoveryRequired
                    ? "PivotTable+ could not enable the Data Model; the original PivotTable was preserved or restored, and generated cache dependencies were retained under durable recovery ownership."
                    : rollbackCompleted
                        ? "PivotTable+ could not enable the Data Model; the original PivotTable was preserved or restored."
                        : "PivotTable+ could not enable the Data Model and cleanup did not complete.") +
                " Stage: " + failedStage + ". " + innerException.Message,
                innerException)
        {
            FailedStage = failedStage;
            RollbackCompleted = rollbackCompleted;
            RecoveryRequired = recoveryRequired;
        }

        public PivotDataModelEnablementStage FailedStage { get; }

        public bool RollbackCompleted { get; }

        public bool RecoveryRequired { get; }
    }

    public interface IPivotReplacementTransaction : IDisposable
    {
        bool ReplacementAttempted { get; }

        bool IsCommitted { get; }

        object? ReplacementPivotTable { get; }

        void ReplaceAtOriginalLocation();

        void VerifyReplacement();

        void RollBack();

        void Commit();
    }

    /// <summary>
    /// Excel boundary for the reversible conversion. Implementations must make
    /// Artifact planning must precede durable Pending ownership. Ensure is
    /// idempotent against that exact plan; staging mutations retain recovery
    /// ownership when Excel creates a non-deletable PivotCache dependency.
    /// </summary>
    public interface IPivotDataModelEnablementGateway
    {
        void VerifyBoundTarget(
            object workbook,
            object pivotTable,
            PivotTargetIdentity expectedTarget);

        void PersistBoundWorkbookIdentity(
            object workbook,
            PivotTargetIdentity expectedTarget);

        ClassicPivotSourceDescriptor InspectSupportedSource(
            object workbook,
            object pivotTable,
            PivotSourceDescriptor expectedSource);

        PivotNativeStateSnapshot CaptureReversibleState(object pivotTable);

        PivotDataModelArtifactPlan PlanOwnedModelArtifacts(
            string setupId,
            ClassicPivotSourceDescriptor source,
            PivotTargetIdentity target,
            PivotNativeStateSnapshot originalState);

        void PreflightOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata? recoveryOwnership);

        PivotDataModelArtifacts EnsureOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata ownership);

        PivotDataModelArtifacts ValidatePendingDataModelFinalization(
            object workbook,
            object pivotTable,
            string setupId,
            PivotPlusWorkbookMetadata ownership);

        PivotStagedDataModelPivot CreateStagedDataModelPivot(
            object workbook,
            string setupId,
            PivotDataModelArtifacts artifacts);

        void RestoreState(
            object pivotTable,
            PivotNativeStateSnapshot snapshot,
            string modelTableName);

        void RefreshPivotTable(object pivotTable);

        string VerifyDataModelState(object pivotTable, PivotNativeStateSnapshot expectedState);

        void MarkStagingVerified(
            PivotStagedDataModelPivot stagedPivot,
            string stagingStateFingerprint);

        IPivotReplacementTransaction PrepareReplacement(
            object workbook,
            object originalPivotTable,
            PivotStagedDataModelPivot stagedPivot,
            PivotNativeStateSnapshot originalState,
            string modelTableName);

        void DeleteStagingPivot(object workbook, PivotStagedDataModelPivot stagedPivot);

        void DeleteOwnedModelArtifacts(object workbook, PivotDataModelArtifacts artifacts);

        PivotPendingDataModelRecovery RecoverPending(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership);

        void VerifyActiveDataModelOwnership(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership);
    }

    internal interface IPivotDataModelOwnershipStore
    {
        PivotPlusWorkbookMetadata? DemandAvailableOrExactRecovery(
            object workbook,
            string setupId,
            PivotTargetIdentity target);

        PivotPlusWorkbookMetadata SavePending(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifactPlan plan);

        PivotPlusWorkbookMetadata MarkStagingVerified(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifacts artifacts,
            string stagingStateFingerprint);

        PivotPlusWorkbookMetadata DemandPendingBySetupId(
            object workbook,
            string setupId);

        void MarkActive(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifacts artifacts);
    }

    internal sealed class PivotDataModelOwnershipStore : IPivotDataModelOwnershipStore
    {
        private readonly PivotPlusWorkbookMetadataStore store =
            new PivotPlusWorkbookMetadataStore();

        public PivotPlusWorkbookMetadata? DemandAvailableOrExactRecovery(
            object workbook,
            string setupId,
            PivotTargetIdentity target)
        {
            IReadOnlyList<PivotPlusWorkbookMetadata> existing =
                store.LoadAll((dynamic)workbook);
            PivotPlusWorkbookMetadata? setup = existing.SingleOrDefault(item =>
                string.Equals(item.SetupId, setupId, StringComparison.OrdinalIgnoreCase));
            PivotPlusWorkbookMetadata? targetSetup = existing.SingleOrDefault(item =>
                string.Equals(
                    item.TargetWorksheetName,
                    target.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    item.TargetPivotTableName,
                    target.PivotTableName,
                    StringComparison.OrdinalIgnoreCase));
            if (setup != null &&
                (!string.Equals(
                     setup.TargetWorksheetName,
                     target.WorksheetName,
                     StringComparison.OrdinalIgnoreCase) ||
                 !string.Equals(
                     setup.TargetPivotTableName,
                     target.PivotTableName,
                     StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "A PivotTable+ setup already uses this setup identifier for a different target.");
            }

            if (targetSetup != null && setup == null)
            {
                throw new InvalidOperationException(
                    "A different PivotTable+ setup already references the selected PivotTable.");
            }

            if (setup != null &&
                setup.RecoveryPhase == PivotPlusRecoveryPhase.None &&
                !setup.Artifacts.Any(item =>
                    item.Kind == PivotPlusArtifactKind.TemporaryWorksheet))
            {
                throw new InvalidOperationException(
                    "The PivotTable+ setup is already active and is not a pending conversion recovery.");
            }

            return setup;
        }

        public PivotPlusWorkbookMetadata SavePending(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifactPlan plan)
        {
            var owned = new List<PivotPlusOwnedArtifact>
            {
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Query,
                    ArtifactId = plan.QueryName,
                    Fingerprint = plan.QueryFingerprint
                },
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Connection,
                    ArtifactId = plan.ConnectionName,
                    Fingerprint = plan.ConnectionFingerprint
                }
            };
            if (plan.WorkbookName != null &&
                plan.WorkbookNameFingerprint != null)
            {
                owned.Add(new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.WorkbookName,
                    ArtifactId = plan.WorkbookName,
                    Fingerprint = plan.WorkbookNameFingerprint
                });
            }

            foreach (PivotTemporaryWorksheetArtifact temporary in plan.TemporaryWorksheets)
            {
                owned.Add(new PivotPlusOwnedArtifact
                {
                    Kind = temporary.Kind,
                    ArtifactId = temporary.Name,
                    Fingerprint = temporary.Fingerprint
                });
            }

            PivotTemporaryPivotTableArtifact temporaryPivot =
                plan.TemporaryPivotTable ??
                throw new InvalidOperationException(
                    "A pending conversion requires a deterministic temporary target PivotTable receipt.");
            owned.Add(new PivotPlusOwnedArtifact
            {
                Kind = temporaryPivot.Kind,
                ArtifactId = temporaryPivot.Name,
                Fingerprint = temporaryPivot.Fingerprint
            });

            var metadata = new PivotPlusWorkbookMetadata
            {
                SetupId = setupId,
                TargetWorksheetName = target.WorksheetName,
                TargetPivotTableName = target.PivotTableName,
                RecoveryPhase = PivotPlusRecoveryPhase.Planned,
                TargetAnchorAddress = temporaryPivot.TargetAnchorAddress,
                Artifacts = owned
            };
            store.Save((dynamic)workbook, metadata);
            return metadata;
        }

        public PivotPlusWorkbookMetadata MarkStagingVerified(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifacts artifacts,
            string stagingStateFingerprint)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));
            PivotPlusMetadataValidator.ValidateFingerprint(
                stagingStateFingerprint,
                "staging state fingerprint");
            PivotPlusWorkbookMetadata metadata = store.Load((dynamic)workbook, setupId) ??
                throw new InvalidOperationException(
                    "PivotTable+ planned recovery ownership disappeared before the staging checkpoint.");
            DemandSameTarget(metadata, target);
            if (!string.Equals(
                    metadata.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal) ||
                (metadata.RecoveryPhase != PivotPlusRecoveryPhase.Planned &&
                 metadata.RecoveryPhase != PivotPlusRecoveryPhase.StagingVerified))
            {
                throw new InvalidOperationException(
                    "PivotTable+ ownership is not in a checkpointable recovery phase.");
            }

            DemandExactArtifactSet(metadata, artifacts);
            PivotTemporaryPivotTableArtifact temporary =
                artifacts.TemporaryPivotTable ??
                throw new InvalidOperationException(
                    "The staged conversion has no temporary target PivotTable receipt.");
            if (!string.Equals(
                    metadata.TargetAnchorAddress,
                    temporary.TargetAnchorAddress,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The persisted target anchor changed before the staging checkpoint.");
            }

            if (metadata.RecoveryPhase == PivotPlusRecoveryPhase.StagingVerified)
            {
                if (!string.Equals(
                        metadata.StagingStateFingerprint,
                        stagingStateFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The persisted staging checkpoint does not match the verified live state.");
                }

                return metadata;
            }

            metadata.RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified;
            metadata.StagingStateFingerprint = stagingStateFingerprint;
            store.Save((dynamic)workbook, metadata);
            return metadata;
        }

        public PivotPlusWorkbookMetadata DemandPendingBySetupId(
            object workbook,
            string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            PivotPlusWorkbookMetadata metadata = store.Load((dynamic)workbook, setupId) ??
                throw new InvalidOperationException(
                    "No PivotTable+ ownership exists for this setup identifier.");
            if (!string.Equals(
                    metadata.SchemaVersion,
                    PivotPlusWorkbookMetadata.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The PivotTable+ setup has no current pending recovery checkpoint.");
            }

            // Returning a fully validated Active record makes the public
            // recovery entry point idempotent after a CustomXML commit whose
            // COM acknowledgement was ambiguous. The service recognizes this
            // terminal state and does not replay gateway mutations.
            return metadata;
        }

        public void MarkActive(
            object workbook,
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifacts artifacts)
        {
            PivotPlusWorkbookMetadata metadata = store.Load((dynamic)workbook, setupId) ??
                throw new InvalidOperationException(
                    "PivotTable+ recovery ownership disappeared before activation.");
            DemandSameTarget(metadata, target);
            if (metadata.RecoveryPhase != PivotPlusRecoveryPhase.StagingVerified)
            {
                throw new InvalidOperationException(
                    "PivotTable+ recovery ownership was not staging-verified before activation.");
            }

            DemandExactArtifactSet(metadata, artifacts);

            metadata.SchemaVersion = PivotPlusWorkbookMetadata.CurrentSchemaVersion;
            metadata.Artifacts = metadata.Artifacts
                .Where(item =>
                    item.Kind != PivotPlusArtifactKind.TemporaryWorksheet &&
                    item.Kind != PivotPlusArtifactKind.TemporaryPivotTable)
                .ToList();
            metadata.RecoveryPhase = PivotPlusRecoveryPhase.None;
            metadata.TargetAnchorAddress = string.Empty;
            metadata.StagingStateFingerprint = string.Empty;
            store.Save((dynamic)workbook, metadata);
        }

        private static void DemandSameTarget(
            PivotPlusWorkbookMetadata metadata,
            PivotTargetIdentity target)
        {
            if (!string.Equals(
                    metadata.TargetWorksheetName,
                    target.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    metadata.TargetPivotTableName,
                    target.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "PivotTable+ recovery ownership changed target.");
            }
        }

        private static void DemandExactArtifactSet(
            PivotPlusWorkbookMetadata metadata,
            PivotDataModelArtifacts artifacts)
        {
            var expected = new List<PivotPlusOwnedArtifact>
            {
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Query,
                    ArtifactId = artifacts.QueryName,
                    Fingerprint = artifacts.QueryFingerprint
                },
                new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.Connection,
                    ArtifactId = artifacts.ConnectionName,
                    Fingerprint = artifacts.ConnectionFingerprint
                }
            };
            if (artifacts.OwnedWorkbookName != null)
            {
                expected.Add(new PivotPlusOwnedArtifact
                {
                    Kind = PivotPlusArtifactKind.WorkbookName,
                    ArtifactId = artifacts.OwnedWorkbookName.Name,
                    Fingerprint = artifacts.OwnedWorkbookName.ReferenceFingerprint
                });
            }

            expected.AddRange(artifacts.TemporaryWorksheets.Select(item =>
                new PivotPlusOwnedArtifact
                {
                    Kind = item.Kind,
                    ArtifactId = item.Name,
                    Fingerprint = item.Fingerprint
                }));
            PivotTemporaryPivotTableArtifact temporary =
                artifacts.TemporaryPivotTable ??
                throw new InvalidOperationException(
                    "The recovery artifact set has no temporary target PivotTable receipt.");
            expected.Add(new PivotPlusOwnedArtifact
            {
                Kind = temporary.Kind,
                ArtifactId = temporary.Name,
                Fingerprint = temporary.Fingerprint
            });
            if (metadata.Artifacts.Count != expected.Count ||
                expected.Any(receipt => metadata.Artifacts.Count(recorded =>
                    recorded.Kind == receipt.Kind &&
                    string.Equals(
                        recorded.ArtifactId,
                        receipt.ArtifactId,
                        StringComparison.Ordinal) &&
                    string.Equals(
                        recorded.Fingerprint,
                        receipt.Fingerprint,
                        StringComparison.Ordinal)) != 1))
            {
                throw new InvalidOperationException(
                    "PivotTable+ recovery ownership changed before activation.");
            }
        }
    }

    /// <summary>
    /// Converts only a selected classic PivotTable backed by a supported
    /// workbook table or named range. The original is not touched until a
    /// staged Data Model PivotTable has restored and verified the explicitly
    /// supported bounded native snapshot.
    /// </summary>
    public sealed class PivotDataModelEnablementService
    {
        private static readonly Regex SetupIdPattern = new Regex(
            "^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$",
            RegexOptions.CultureInvariant);

        private readonly IPivotDataModelEnablementGateway gateway;
        private readonly IPivotDataModelOwnershipStore ownershipStore;

        public PivotDataModelEnablementService()
            : this(
                new LateBoundPivotDataModelEnablementGateway(),
                new PivotDataModelOwnershipStore())
        {
        }

        public PivotDataModelEnablementService(IPivotDataModelEnablementGateway gateway)
            : this(gateway, new PivotDataModelOwnershipStore())
        {
        }

        internal PivotDataModelEnablementService(
            IPivotDataModelEnablementGateway gateway,
            IPivotDataModelOwnershipStore ownershipStore)
        {
            this.gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
            this.ownershipStore = ownershipStore ??
                throw new ArgumentNullException(nameof(ownershipStore));
        }

        public PivotDataModelEnablementResult Enable(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            string validatedSetupId = setupId ?? string.Empty;
            if (!SetupIdPattern.IsMatch(validatedSetupId))
            {
                throw new ArgumentException("A bounded path-free setup identifier is required.", nameof(setupId));
            }

            if (context.Definition.Source.Kind != PivotSourceKind.WorksheetRange &&
                context.Definition.Source.Kind != PivotSourceKind.WorksheetTable)
            {
                throw new InvalidOperationException(
                    "Only a classic worksheet-backed PivotTable can be explicitly upgraded to the Data Model.");
            }

            if ((context.Definition.Source.Capabilities & PivotCapability.UpgradeToDataModel) == 0)
            {
                throw new NotSupportedException(
                    "The selected PivotTable source was not discovered as safely upgradeable to the Data Model.");
            }

            ClassicPivotSourceDescriptor? source = null;
            PivotNativeStateSnapshot? snapshot = null;
            PivotDataModelArtifactPlan? artifactPlan = null;
            PivotDataModelArtifacts? artifacts = null;
            PivotStagedDataModelPivot? stagedPivot = null;
            IPivotReplacementTransaction? replacement = null;
            PivotPlusWorkbookMetadata? recoveryOwnership = null;
            bool recoveryOwnershipPersisted = false;
            PivotDataModelEnablementStage stage = PivotDataModelEnablementStage.BindTarget;

            try
            {
                gateway.VerifyBoundTarget(
                    workbook,
                    pivotTable,
                    context.Definition.Target);

                stage = PivotDataModelEnablementStage.PreflightOwnership;
                recoveryOwnership = ownershipStore.DemandAvailableOrExactRecovery(
                    workbook,
                    validatedSetupId,
                    context.Definition.Target);
                recoveryOwnershipPersisted = recoveryOwnership != null;

                stage = PivotDataModelEnablementStage.InspectSource;
                source = gateway.InspectSupportedSource(
                    workbook,
                    pivotTable,
                    context.Definition.Source);

                stage = PivotDataModelEnablementStage.CaptureState;
                snapshot = gateway.CaptureReversibleState(pivotTable);
                DemandSameTarget(context.Definition.Target, snapshot);

                stage = PivotDataModelEnablementStage.PersistWorkbookIdentity;
                gateway.PersistBoundWorkbookIdentity(
                    workbook,
                    context.Definition.Target);

                stage = PivotDataModelEnablementStage.CreateModelArtifacts;
                artifactPlan = gateway.PlanOwnedModelArtifacts(
                    validatedSetupId,
                    source,
                    context.Definition.Target,
                    snapshot);
                gateway.PreflightOwnedModelArtifacts(
                    workbook,
                    artifactPlan,
                    recoveryOwnership);
                if (recoveryOwnership == null)
                {
                    // Write-ahead ownership precedes every generated Excel
                    // object. Any partial COM commit is therefore recoverable
                    // by exact plan/fingerprint without adopting a collision.
                    stage = PivotDataModelEnablementStage.PersistOwnership;
                    recoveryOwnership = ownershipStore.SavePending(
                        workbook,
                        validatedSetupId,
                        context.Definition.Target,
                        artifactPlan);
                    recoveryOwnershipPersisted = true;
                }

                stage = PivotDataModelEnablementStage.CreateModelArtifacts;
                artifacts = gateway.EnsureOwnedModelArtifacts(
                    workbook,
                    artifactPlan,
                    recoveryOwnership ?? throw new InvalidOperationException(
                        "PivotTable+ pending ownership was not persisted."));

                stage = PivotDataModelEnablementStage.CreateStagingPivot;
                stagedPivot = gateway.CreateStagedDataModelPivot(
                    workbook,
                    validatedSetupId,
                    artifacts);

                stage = PivotDataModelEnablementStage.RestoreStagingState;
                gateway.RestoreState(
                    stagedPivot.NativePivotTable,
                    snapshot,
                    artifacts.ModelTableName);
                gateway.RefreshPivotTable(stagedPivot.NativePivotTable);

                stage = PivotDataModelEnablementStage.VerifyStagingPivot;
                string stagingStateFingerprint = gateway.VerifyDataModelState(
                    stagedPivot.NativePivotTable,
                    snapshot);
                gateway.MarkStagingVerified(
                    stagedPivot,
                    stagingStateFingerprint);
                recoveryOwnership = ownershipStore.MarkStagingVerified(
                    workbook,
                    validatedSetupId,
                    context.Definition.Target,
                    artifacts,
                    stagingStateFingerprint);

                stage = PivotDataModelEnablementStage.PrepareReplacement;
                replacement = gateway.PrepareReplacement(
                    workbook,
                    pivotTable,
                    stagedPivot,
                    snapshot,
                    artifacts.ModelTableName);

                stage = PivotDataModelEnablementStage.ReplaceOriginal;
                replacement.ReplaceAtOriginalLocation();

                stage = PivotDataModelEnablementStage.VerifyReplacement;
                replacement.VerifyReplacement();

                stage = PivotDataModelEnablementStage.CleanupStaging;
                // Format backup deletion is the forward-commit boundary. The
                // verified staging PivotTable remains the durable semantic
                // source until that succeeds.
                replacement.Commit();

                gateway.DeleteStagingPivot(workbook, stagedPivot);
                stagedPivot = null;

                stage = PivotDataModelEnablementStage.PersistOwnership;
                ownershipStore.MarkActive(
                    workbook,
                    validatedSetupId,
                    context.Definition.Target,
                    artifacts);
                stage = PivotDataModelEnablementStage.Complete;
                return new PivotDataModelEnablementResult(
                    context.Definition.Target,
                    source,
                    artifacts);
            }
            catch (Exception failure) when (!(failure is PivotDataModelEnablementException))
            {
                bool ownershipOutcomeAmbiguous =
                    failure is PivotPlusOwnershipAmbiguousException;
                bool preserveOwnedDependencies =
                    recoveryOwnershipPersisted || ownershipOutcomeAmbiguous;
                bool rollbackCompleted = TryRollBack(
                    workbook,
                    replacement,
                    stagedPivot,
                    artifacts,
                    preserveCacheDependencies: preserveOwnedDependencies,
                    out IReadOnlyList<Exception> rollbackFailures);
                Exception cause = rollbackFailures.Count == 0
                    ? failure
                    : new AggregateException(new[] { failure }.Concat(rollbackFailures));
                throw new PivotDataModelEnablementException(
                    stage,
                    rollbackCompleted,
                    preserveOwnedDependencies,
                    cause);
            }
            finally
            {
                replacement?.Dispose();
            }
        }

        /// <summary>
        /// Completes only the metadata transition left after a replacement was
        /// already committed but the final Pending-to-Active save failed. It
        /// never creates or clears a PivotTable, PivotCache, query, connection,
        /// name, or worksheet.
        /// </summary>
        public void FinalizePending(
            object workbook,
            object pivotTable,
            PivotTableContext context,
            string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            string validatedSetupId = setupId ?? string.Empty;
            if (!SetupIdPattern.IsMatch(validatedSetupId))
            {
                throw new ArgumentException(
                    "A bounded path-free setup identifier is required.",
                    nameof(setupId));
            }

            if (context.Definition.Source.Kind != PivotSourceKind.DataModel)
            {
                throw new InvalidOperationException(
                    "Pending ownership finalization requires the already-committed native Data Model PivotTable.");
            }

            gateway.VerifyBoundTarget(
                workbook,
                pivotTable,
                context.Definition.Target);
            // Route the legacy selection-based entry point through the same
            // persisted, hash-verified recovery state machine. Selection and
            // context are used only as an additional target binding guard.
            RecoverPending(workbook, validatedSetupId);
        }

        /// <summary>
        /// Resumes a durably checkpointed conversion without relying on the
        /// active selection or an in-memory classic PivotTable snapshot. The
        /// gateway validates every persisted receipt and live recovery object
        /// before it mutates the target.
        /// </summary>
        public void RecoverPending(object workbook, string setupId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            string validatedSetupId = setupId ?? string.Empty;
            if (!SetupIdPattern.IsMatch(validatedSetupId))
            {
                throw new ArgumentException(
                    "A bounded path-free setup identifier is required.",
                    nameof(setupId));
            }

            PivotPlusWorkbookMetadata ownership =
                ownershipStore.DemandPendingBySetupId(
                    workbook,
                    validatedSetupId);
            if (ownership.RecoveryPhase == PivotPlusRecoveryPhase.None)
            {
                // MarkActive may have committed even when Excel surfaced a
                // late COM failure. Verify the exact live target and owned
                // artifacts before accepting that terminal receipt; no
                // temporary recovery object may be touched again.
                gateway.VerifyActiveDataModelOwnership(
                    workbook,
                    validatedSetupId,
                    ownership);
                return;
            }

            PivotPendingDataModelRecovery recovery = gateway.RecoverPending(
                workbook,
                validatedSetupId,
                ownership);
            ownershipStore.MarkActive(
                workbook,
                validatedSetupId,
                recovery.Target,
                recovery.Artifacts);
        }

        private static void DemandSameTarget(
            PivotTargetIdentity expected,
            PivotNativeStateSnapshot actual)
        {
            if (!string.Equals(expected.WorksheetName, actual.WorksheetName, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.PivotTableName, actual.PivotTableName, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable changed after discovery. Re-select it before enabling the Data Model.");
            }
        }

        private bool TryRollBack(
            object workbook,
            IPivotReplacementTransaction? replacement,
            PivotStagedDataModelPivot? stagedPivot,
            PivotDataModelArtifacts? artifacts,
            bool preserveCacheDependencies,
            out IReadOnlyList<Exception> failures)
        {
            var rollbackFailures = new List<Exception>();
            bool originalRestored = replacement == null ||
                                    !replacement.ReplacementAttempted ||
                                    replacement.IsCommitted;
            if (replacement != null &&
                replacement.ReplacementAttempted &&
                !replacement.IsCommitted)
            {
                originalRestored = TryCleanup(
                    "restore the original PivotTable",
                    replacement.RollBack,
                    rollbackFailures);
            }

            bool mayDeleteStaging = replacement == null ||
                                    !replacement.ReplacementAttempted ||
                                    (originalRestored && !replacement.IsCommitted);
            if (stagedPivot != null && mayDeleteStaging)
            {
                TryCleanup(
                    "remove the staging PivotTable",
                    () => gateway.DeleteStagingPivot(workbook, stagedPivot),
                    rollbackFailures);
            }

            // Do not remove a connection that may still back the replacement.
            // A failed rollback deliberately leaves the exact generated model
            // artifacts in place for a subsequent explicit recovery attempt.
            if (artifacts != null && originalRestored && !preserveCacheDependencies)
            {
                TryCleanup(
                    "remove the generated model artifacts",
                    () => gateway.DeleteOwnedModelArtifacts(workbook, artifacts),
                    rollbackFailures);
            }

            failures = rollbackFailures;
            return rollbackFailures.Count == 0;
        }

        private static bool TryCleanup(
            string operation,
            Action cleanup,
            ICollection<Exception> failures)
        {
            try
            {
                cleanup();
                return true;
            }
            catch (Exception exception)
            {
                failures.Add(new InvalidOperationException(
                    "PivotTable+ could not " + operation + ".",
                    exception));
                return false;
            }
        }
    }
}
