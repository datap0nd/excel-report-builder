using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.DataModel;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotDataModelEnablementTests
{
    [Fact]
    public void Enable_VerifiesStagingBeforeTouchingOriginal_AndCommitsNativeReplacement()
    {
        var gateway = new RecordingGateway();
        var service = Service(gateway);

        PivotDataModelEnablementResult result = service.Enable(
            new object(),
            new object(),
            ClassicContext(),
            "setup-1");

        Assert.Equal("Sheet1", result.Target.WorksheetName);
        Assert.Equal("SalesTable", result.Source.WorkbookObjectName);
        Assert.Equal(
            new[]
            {
                "bind-target",
                "ownership-preflight",
                "inspect",
                "capture",
                "persist-identity",
                "plan-artifacts",
                "artifact-preflight",
                "ownership-save",
                "ensure-artifacts",
                "create-stage",
                "restore-stage",
                "refresh-stage",
                "verify-stage",
                "mark-stage-verified",
                "ownership-stage-verified",
                "prepare-replacement",
                "replace-original",
                "verify-replacement",
                "commit",
                "delete-stage",
                "ownership-active",
                "dispose"
            },
            gateway.Events);
        Assert.True(
            gateway.Events.IndexOf("verify-stage") <
            gateway.Events.IndexOf("replace-original"));
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
        Assert.Contains("Excel.CurrentWorkbook", gateway.QueryFormula, StringComparison.Ordinal);
    }

    [Fact]
    public void Enable_WhenStagingVerificationFails_LeavesOriginalUntouchedAndCleansGeneratedObjects()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnStagingVerification = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.VerifyStagingPivot, exception.FailedStage);
        Assert.True(exception.RollbackCompleted);
        Assert.True(exception.RecoveryRequired);
        Assert.DoesNotContain("prepare-replacement", gateway.Events);
        Assert.DoesNotContain("replace-original", gateway.Events);
        Assert.Equal("delete-stage", gateway.Events.Last());
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
    }

    [Fact]
    public void Enable_WhenReplacementFailsAfterPartialMutation_RestoresOriginalThenCleansArtifacts()
    {
        var gateway = new RecordingGateway
        {
            ThrowDuringReplacement = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.ReplaceOriginal, exception.FailedStage);
        Assert.True(exception.RollbackCompleted);
        Assert.True(exception.RecoveryRequired);
        Assert.Equal(
            new[] { "replace-original", "rollback-original", "delete-stage", "dispose" },
            gateway.Events.TakeLast(4));
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
    }

    [Fact]
    public void Enable_WhenOriginalCannotBeRestored_PreservesModelArtifactsForRecovery()
    {
        var gateway = new RecordingGateway
        {
            ThrowDuringReplacement = true,
            ThrowDuringRollback = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.False(exception.RollbackCompleted);
        Assert.Equal(1, gateway.Events.Count(item => item == "rollback-original"));
        Assert.DoesNotContain("delete-stage", gateway.Events);
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
        Assert.True(gateway.StagingPresent);
    }

    [Fact]
    public void Enable_WhenCommitThrowsAfterForwardCommit_PreservesVerifiedTargetAndStagingForRecovery()
    {
        var gateway = new RecordingGateway
        {
            ThrowAfterReplacementCommit = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.CleanupStaging, exception.FailedStage);
        Assert.True(exception.RollbackCompleted);
        Assert.True(exception.RecoveryRequired);
        Assert.True(gateway.ReplacementCommitted);
        Assert.True(gateway.StagingPresent);
        Assert.Equal(1, gateway.Events.Count(item => item == "commit"));
        Assert.DoesNotContain("rollback-original", gateway.Events);
        Assert.DoesNotContain("delete-stage", gateway.Events);
        Assert.DoesNotContain("ownership-active", gateway.Events);
    }

    [Fact]
    public void Enable_WhenStagingDeletionFailsAfterCommit_DoesNotRetryDeletionOrRollBack()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnStageDeletion = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.CleanupStaging, exception.FailedStage);
        Assert.True(exception.RollbackCompleted);
        Assert.True(exception.RecoveryRequired);
        Assert.True(gateway.ReplacementCommitted);
        Assert.True(gateway.StagingPresent);
        Assert.Equal(1, gateway.StageDeleteCalls);
        Assert.Equal(
            new[] { "commit", "delete-stage", "dispose" },
            gateway.Events.TakeLast(3));
        Assert.DoesNotContain("rollback-original", gateway.Events);
        Assert.DoesNotContain("ownership-active", gateway.Events);
    }

    [Fact]
    public void Enable_RejectsNonClassicContextBeforeCreatingAnything()
    {
        var gateway = new RecordingGateway();
        var service = Service(gateway);

        Assert.Throws<InvalidOperationException>(
            () => service.Enable(
                new object(),
                new object(),
                Context(PivotSourceKind.DataModel, PivotCapability.DataModel),
                "setup-1"));

        Assert.Empty(gateway.Events);
    }

    [Fact]
    public void Enable_RejectsChangedTargetBeforeCreatingOwnedArtifacts()
    {
        var gateway = new RecordingGateway
        {
            CapturedWorksheetName = "OtherSheet"
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.CaptureState, exception.FailedStage);
        Assert.Equal(
            new[] { "bind-target", "ownership-preflight", "inspect", "capture" },
            gateway.Events);
    }

    [Fact]
    public void Enable_WhenCapturePreflightRejects_PersistsNoIdentityOrOwnership()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnCapture = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.CaptureState, exception.FailedStage);
        Assert.Equal(
            new[] { "bind-target", "ownership-preflight", "inspect", "capture" },
            gateway.Events);
        Assert.DoesNotContain("persist-identity", gateway.Events);
        Assert.DoesNotContain("ownership-save", gateway.Events);
        Assert.DoesNotContain("plan-artifacts", gateway.Events);
    }

    [Fact]
    public void Enable_WhenRecoveryOwnershipPersistenceFails_CreatesNoArtifactsOrCache()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnOwnershipSave = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.PersistOwnership, exception.FailedStage);
        Assert.True(exception.RollbackCompleted);
        Assert.False(exception.RecoveryRequired);
        Assert.Equal("ownership-save", gateway.Events.Last());
        Assert.DoesNotContain("ensure-artifacts", gateway.Events);
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
        Assert.DoesNotContain("create-stage", gateway.Events);
    }

    [Fact]
    public void Enable_WhenStagingCacheCreationFails_PreservesDurablyOwnedDependencies()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnStageCreation = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.CreateStagingPivot, exception.FailedStage);
        Assert.True(exception.RecoveryRequired);
        Assert.Contains("ownership-save", gateway.Events);
        Assert.DoesNotContain("delete-artifacts", gateway.Events);
    }

    [Fact]
    public void Enable_RetryWithExactPendingOwnership_ReusesArtifactsWithoutSavingAgain()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnStagingVerification = true
        };
        var service = Service(gateway);

        Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));
        Assert.Equal(1, gateway.ArtifactCreateCalls);
        Assert.Equal(1, gateway.OwnershipSaveCalls);

        gateway.ThrowOnStagingVerification = false;
        gateway.Events.Clear();
        service.Enable(new object(), new object(), ClassicContext(), "setup-1");

        Assert.Contains("recover-artifacts", gateway.Events);
        Assert.Contains("plan-artifacts", gateway.Events);
        Assert.Contains("artifact-preflight", gateway.Events);
        Assert.DoesNotContain("ownership-save", gateway.Events);
        Assert.Equal(1, gateway.ArtifactCreateCalls);
        Assert.Equal(1, gateway.OwnershipSaveCalls);
        Assert.Equal(1, gateway.RecoveryCalls);
    }

    [Fact]
    public void Enable_WhenActivationFails_DoesNotRollBackCommittedReplacement()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnOwnershipActivation = true
        };
        var service = Service(gateway);

        PivotDataModelEnablementException exception = Assert.Throws<PivotDataModelEnablementException>(
            () => service.Enable(new object(), new object(), ClassicContext(), "setup-1"));

        Assert.Equal(PivotDataModelEnablementStage.PersistOwnership, exception.FailedStage);
        Assert.True(exception.RecoveryRequired);
        Assert.Contains("commit", gateway.Events);
        Assert.True(gateway.ReplacementCommitted);
        Assert.False(gateway.StagingPresent);
        Assert.Equal(1, gateway.StageDeleteCalls);
        Assert.DoesNotContain("rollback-original", gateway.Events);

        gateway.ThrowOnOwnershipActivation = false;
        gateway.Events.Clear();
        service.RecoverPending(new object(), "setup-1");

        Assert.Equal(
            new[]
            {
                "ownership-load-pending",
                "recover-pending",
                "ownership-active"
            },
            gateway.Events);
        Assert.DoesNotContain("create-artifacts", gateway.Events);
        Assert.DoesNotContain("create-stage", gateway.Events);
        Assert.DoesNotContain("replace-original", gateway.Events);
    }

    [Fact]
    public void RecoverPending_RecoversGatewayStateBeforeActivatingOwnership_InExactOrder()
    {
        var gateway = new RecordingGateway();
        gateway.SeedStagingVerifiedRecovery();
        var service = Service(gateway);

        service.RecoverPending(new object(), "setup-1");

        Assert.Equal(
            new[]
            {
                "ownership-load-pending",
                "recover-pending",
                "ownership-active"
            },
            gateway.Events);
        Assert.Equal(1, gateway.PendingRecoveryCalls);
        Assert.Equal(1, gateway.OwnershipActivationCalls);
        Assert.True(gateway.OwnershipIsActive);
    }

    [Fact]
    public void RecoverPending_WhenGatewayRecoveryFails_DoesNotActivateOwnership()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnPendingRecovery = true
        };
        gateway.SeedStagingVerifiedRecovery();
        var service = Service(gateway);

        Assert.Throws<InvalidOperationException>(
            () => service.RecoverPending(new object(), "setup-1"));

        Assert.Equal(
            new[] { "ownership-load-pending", "recover-pending" },
            gateway.Events);
        Assert.Equal(1, gateway.PendingRecoveryCalls);
        Assert.Equal(0, gateway.OwnershipActivationCalls);
        Assert.False(gateway.OwnershipIsActive);
    }

    [Fact]
    public void RecoverPending_WhenActivationFailsBeforeCommit_SecondCallConverges()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnOwnershipActivation = true
        };
        gateway.SeedStagingVerifiedRecovery();
        var service = Service(gateway);

        Assert.Throws<InvalidOperationException>(
            () => service.RecoverPending(new object(), "setup-1"));

        Assert.Equal(
            new[]
            {
                "ownership-load-pending",
                "recover-pending",
                "ownership-active"
            },
            gateway.Events);
        Assert.False(gateway.OwnershipIsActive);

        gateway.ThrowOnOwnershipActivation = false;
        gateway.Events.Clear();
        service.RecoverPending(new object(), "setup-1");

        Assert.Equal(
            new[]
            {
                "ownership-load-pending",
                "recover-pending",
                "ownership-active"
            },
            gateway.Events);
        Assert.Equal(2, gateway.PendingRecoveryCalls);
        Assert.Equal(2, gateway.OwnershipActivationCalls);
        Assert.True(gateway.OwnershipIsActive);
    }

    [Fact]
    public void RecoverPending_WhenActivationCommitsThenReportsAmbiguity_ActiveReceiptMakesRetryANoOp()
    {
        var gateway = new RecordingGateway
        {
            ThrowOnOwnershipActivationAfterCommit = true
        };
        gateway.SeedStagingVerifiedRecovery();
        var service = Service(gateway);

        Assert.Throws<PivotPlusOwnershipAmbiguousException>(
            () => service.RecoverPending(new object(), "setup-1"));

        Assert.True(gateway.OwnershipIsActive);
        Assert.Equal(1, gateway.OwnershipActivationCommits);

        gateway.Events.Clear();
        service.RecoverPending(new object(), "setup-1");

        Assert.Equal(
            new[] { "ownership-load-pending", "verify-active" },
            gateway.Events);
        Assert.Equal(1, gateway.PendingRecoveryCalls);
        Assert.Equal(1, gateway.OwnershipActivationCalls);
        Assert.Equal(1, gateway.OwnershipActivationCommits);
        Assert.True(gateway.OwnershipIsActive);
    }

    [Fact]
    public void RecoverPending_RealStoreDeleteCommitsThenInspectionFails_RepeatedPublicRetryVerifiesActiveOnly()
    {
        var gateway = new RecordingGateway();
        gateway.SeedStagingVerifiedRecovery();
        PivotPlusWorkbookMetadata pending = gateway.RecoveryOwnership ??
            throw new InvalidOperationException();
        var workbook = new PivotPlusPersistenceTests.FaultingWorkbook();
        workbook.CustomXMLParts.ThrowOnceOnSelectAfterCommittedDelete = true;
        workbook.CustomXMLParts.Seed(
            new PivotPlusWorkbookMetadataStore().Serialize(pending),
            throwOnDelete: true,
            removeBeforeThrow: true);
        var service = new PivotDataModelEnablementService(
            gateway,
            new PivotDataModelOwnershipStore());

        Assert.Throws<InvalidOperationException>(() =>
            service.RecoverPending(workbook, "setup-1"));
        Assert.Equal(1, gateway.PendingRecoveryCalls);

        gateway.Events.Clear();
        service.RecoverPending(workbook, "setup-1");
        service.RecoverPending(workbook, "setup-1");

        Assert.Equal(new[] { "verify-active", "verify-active" }, gateway.Events);
        Assert.Equal(1, gateway.PendingRecoveryCalls);
        Assert.Single(workbook.CustomXMLParts.AllXml);
        PivotPlusWorkbookMetadata active =
            new PivotPlusWorkbookMetadataStore().Load(workbook, "setup-1") ??
            throw new InvalidOperationException();
        Assert.Equal(PivotPlusRecoveryPhase.None, active.RecoveryPhase);
    }

    private static PivotDataModelEnablementService Service(RecordingGateway gateway)
    {
        return new PivotDataModelEnablementService(
            gateway,
            new RecordingGateway.RecordingOwnershipStore(gateway));
    }

    private static PivotTableContext ClassicContext()
    {
        return Context(
            PivotSourceKind.WorksheetTable,
            PivotCapability.NativeFieldPlacement |
            PivotCapability.LayoutFormatting |
            PivotCapability.Refresh |
            PivotCapability.UpgradeToDataModel);
    }

    private static PivotTableContext Context(PivotSourceKind kind, PivotCapability capabilities)
    {
        var definition = new PivotLayoutDefinition(
            new PivotTargetIdentity("workbook-token", "Sheet1", "Pivot1"),
            new PivotSourceDescriptor(kind, "SalesTable", capabilities),
            Array.Empty<PivotFieldDescriptor>(),
            Array.Empty<PivotFieldPlacement>());
        return new PivotTableContext(definition, isConnected: true, sourceFieldsComplete: true);
    }

    private sealed class RecordingGateway : IPivotDataModelEnablementGateway
    {
        private readonly object stagedNativePivot = new object();

        public List<string> Events { get; } = new List<string>();

        public bool ThrowOnStagingVerification { get; set; }

        public bool ThrowDuringReplacement { get; set; }

        public bool ThrowDuringRollback { get; set; }

        public bool ThrowOnOwnershipSave { get; set; }

        public bool ThrowOnStageCreation { get; set; }

        public bool ThrowOnOwnershipActivation { get; set; }

        public bool ThrowOnOwnershipActivationAfterCommit { get; set; }

        public bool ThrowOnPendingRecovery { get; set; }

        public bool ThrowOnStageDeletion { get; set; }

        public bool ThrowAfterReplacementCommit { get; set; }

        public bool ThrowOnCapture { get; set; }

        public PivotPlusWorkbookMetadata? RecoveryOwnership { get; set; }

        public string CapturedWorksheetName { get; set; } = "Sheet1";

        public string QueryFormula { get; private set; } = string.Empty;

        public int ArtifactCreateCalls { get; private set; }

        public int RecoveryCalls { get; private set; }

        public int PendingRecoveryCalls { get; private set; }

        public int OwnershipSaveCalls { get; private set; }

        public int OwnershipActivationCalls { get; private set; }

        public int OwnershipActivationCommits { get; private set; }

        public int StageDeleteCalls { get; private set; }

        public bool OwnershipIsActive { get; private set; }

        public bool ReplacementCommitted { get; private set; }

        public bool StagingPresent { get; private set; }

        private bool nativeArtifactsExist;

        public void SeedStagingVerifiedRecovery(string setupId = "setup-1")
        {
            var source = new ClassicPivotSourceDescriptor(
                "SalesTable",
                PivotPlusWorkbookObjectKind.Table);
            var target = new PivotTargetIdentity(
                "workbook-token",
                "Sheet1",
                "Pivot1");
            var snapshot = new PivotNativeStateSnapshot(
                "Sheet1",
                "Pivot1",
                "A1",
                "snapshot",
                new object());
            PivotDataModelArtifactPlan plan = BuildArtifactPlan(
                setupId,
                source,
                target,
                snapshot);
            RecoveryOwnership = CreatePendingOwnership(
                setupId,
                target,
                plan,
                PivotPlusRecoveryPhase.StagingVerified,
                PivotPlusFingerprint.Create("pivotplus.staging-state.v1", "verified"));
            nativeArtifactsExist = true;
            OwnershipIsActive = false;
        }

        public void VerifyBoundTarget(
            object workbook,
            object pivotTable,
            PivotTargetIdentity expectedTarget)
        {
            Events.Add("bind-target");
        }

        public void PersistBoundWorkbookIdentity(
            object workbook,
            PivotTargetIdentity expectedTarget)
        {
            Events.Add("persist-identity");
        }

        public ClassicPivotSourceDescriptor InspectSupportedSource(
            object workbook,
            object pivotTable,
            PivotSourceDescriptor expectedSource)
        {
            Events.Add("inspect");
            return new ClassicPivotSourceDescriptor(
                "SalesTable",
                PivotPlusWorkbookObjectKind.Table);
        }

        public PivotNativeStateSnapshot CaptureReversibleState(object pivotTable)
        {
            Events.Add("capture");
            if (ThrowOnCapture)
            {
                throw new NotSupportedException("capture preflight refused conversion");
            }

            return new PivotNativeStateSnapshot(
                CapturedWorksheetName,
                "Pivot1",
                "A1",
                "snapshot",
                new object());
        }

        public PivotDataModelArtifactPlan PlanOwnedModelArtifacts(
            string setupId,
            ClassicPivotSourceDescriptor source,
            PivotTargetIdentity target,
            PivotNativeStateSnapshot originalState)
        {
            Events.Add("plan-artifacts");
            return BuildArtifactPlan(setupId, source, target, originalState);
        }

        private PivotDataModelArtifactPlan BuildArtifactPlan(
            string setupId,
            ClassicPivotSourceDescriptor source,
            PivotTargetIdentity target,
            PivotNativeStateSnapshot originalState)
        {
            string queryFormula = PivotPlusSourceQueryCompiler.Compile(
                source.WorkbookObjectName,
                source.ObjectKind);
            QueryFormula = queryFormula;
            return new PivotDataModelArtifactPlan(
                "query",
                "connection",
                "model",
                queryFormula,
                PivotPlusFingerprint.Create("pivotplus.query.v1", queryFormula),
                PivotPlusFingerprint.Create("pivotplus.connection.v1", "connection-contract"),
                null,
                null,
                null,
                new[]
                {
                    new PivotTemporaryWorksheetArtifact(
                        "_stage",
                        "staging",
                        PivotPlusFingerprint.Create("pivotplus.temporary-worksheet.v2", "staging\n_stage\nA1"),
                        "A1"),
                    new PivotTemporaryWorksheetArtifact(
                        "_format",
                        "format-backup",
                        PivotPlusFingerprint.Create("pivotplus.temporary-worksheet.v2", "format-backup\n_format\nA1"),
                        "A1")
                },
                new PivotTemporaryPivotTableArtifact(
                    setupId,
                    "_target",
                    PivotPlusFingerprint.Create(
                        "pivotplus.temporary-pivot-table.v1",
                        setupId + "\n_target\nSheet1\nPivot1\nA1\nconnection\nmodel"),
                    "Sheet1",
                    "Pivot1",
                    "A1",
                    "connection",
                    "model"));
        }

        private static PivotPlusWorkbookMetadata CreatePendingOwnership(
            string setupId,
            PivotTargetIdentity target,
            PivotDataModelArtifactPlan plan,
            PivotPlusRecoveryPhase recoveryPhase,
            string stagingStateFingerprint)
        {
            var artifacts = new List<PivotPlusOwnedArtifact>
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
            artifacts.AddRange(plan.TemporaryWorksheets.Select(item =>
                new PivotPlusOwnedArtifact
                {
                    Kind = item.Kind,
                    ArtifactId = item.Name,
                    Fingerprint = item.Fingerprint
                }));
            PivotTemporaryPivotTableArtifact temporaryPivot =
                plan.TemporaryPivotTable ?? throw new InvalidOperationException();
            artifacts.Add(new PivotPlusOwnedArtifact
            {
                Kind = temporaryPivot.Kind,
                ArtifactId = temporaryPivot.Name,
                Fingerprint = temporaryPivot.Fingerprint
            });
            return new PivotPlusWorkbookMetadata
            {
                SetupId = setupId,
                TargetWorksheetName = target.WorksheetName,
                TargetPivotTableName = target.PivotTableName,
                RecoveryPhase = recoveryPhase,
                TargetAnchorAddress = temporaryPivot.TargetAnchorAddress,
                StagingStateFingerprint = stagingStateFingerprint,
                Artifacts = artifacts
            };
        }

        private void CommitOwnershipActivation(PivotPlusWorkbookMetadata pending)
        {
            RecoveryOwnership = new PivotPlusWorkbookMetadata
            {
                SetupId = pending.SetupId,
                TargetWorksheetName = pending.TargetWorksheetName,
                TargetPivotTableName = pending.TargetPivotTableName,
                RecoveryPhase = PivotPlusRecoveryPhase.None,
                Artifacts = pending.Artifacts
                    .Where(item =>
                        item.Kind != PivotPlusArtifactKind.TemporaryWorksheet &&
                        item.Kind != PivotPlusArtifactKind.TemporaryPivotTable)
                    .ToList()
            };
            OwnershipIsActive = true;
            OwnershipActivationCommits++;
        }

        public void PreflightOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata? recoveryOwnership)
        {
            Events.Add("artifact-preflight");
        }

        public PivotDataModelArtifacts EnsureOwnedModelArtifacts(
            object workbook,
            PivotDataModelArtifactPlan plan,
            PivotPlusWorkbookMetadata ownership)
        {
            if (nativeArtifactsExist)
            {
                Events.Add("recover-artifacts");
                RecoveryCalls++;
            }
            else
            {
                Events.Add("ensure-artifacts");
                ArtifactCreateCalls++;
                nativeArtifactsExist = true;
            }

            return new PivotDataModelArtifacts(
                plan.QueryName,
                plan.ConnectionName,
                plan.ModelTableName,
                QueryFormula,
                plan.QueryFingerprint,
                plan.ConnectionFingerprint,
                new object(),
                temporaryWorksheets: plan.TemporaryWorksheets,
                temporaryPivotTable: plan.TemporaryPivotTable);
        }

        public PivotDataModelArtifacts ValidatePendingDataModelFinalization(
            object workbook,
            object pivotTable,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            Events.Add("finalize-pending");
            return new PivotDataModelArtifacts(
                "query",
                "connection",
                "model",
                QueryFormula,
                "query-fingerprint",
                "connection-fingerprint",
                new object(),
                temporaryPivotTable: new PivotTemporaryPivotTableArtifact(
                    setupId,
                    "_target",
                    PivotPlusFingerprint.Create(
                        "pivotplus.temporary-pivot-table.v1",
                        setupId + "\n_target\nSheet1\nPivot1\nA1\nconnection\nmodel"),
                    "Sheet1",
                    "Pivot1",
                    "A1",
                    "connection",
                    "model"));
        }

        public sealed class RecordingOwnershipStore : IPivotDataModelOwnershipStore
        {
            private readonly RecordingGateway owner;

            public RecordingOwnershipStore(RecordingGateway owner)
            {
                this.owner = owner;
            }

            public PivotPlusWorkbookMetadata? DemandAvailableOrExactRecovery(
                object workbook,
                string setupId,
                PivotTargetIdentity target)
            {
                owner.Events.Add("ownership-preflight");
                return owner.RecoveryOwnership;
            }

            public PivotPlusWorkbookMetadata SavePending(
                object workbook,
                string setupId,
                PivotTargetIdentity target,
                PivotDataModelArtifactPlan plan)
            {
                owner.Events.Add("ownership-save");
                owner.OwnershipSaveCalls++;
                if (owner.ThrowOnOwnershipSave)
                {
                    throw new InvalidOperationException("metadata save failed");
                }

                owner.RecoveryOwnership = CreatePendingOwnership(
                    setupId,
                    target,
                    plan,
                    PivotPlusRecoveryPhase.Planned,
                    string.Empty);
                owner.OwnershipIsActive = false;
                return owner.RecoveryOwnership;
            }

            public PivotPlusWorkbookMetadata MarkStagingVerified(
                object workbook,
                string setupId,
                PivotTargetIdentity target,
                PivotDataModelArtifacts artifacts,
                string stagingStateFingerprint)
            {
                owner.Events.Add("ownership-stage-verified");
                PivotPlusWorkbookMetadata metadata = owner.RecoveryOwnership ??
                    throw new InvalidOperationException();
                metadata.RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified;
                metadata.StagingStateFingerprint = stagingStateFingerprint;
                return metadata;
            }

            public PivotPlusWorkbookMetadata DemandPendingBySetupId(
                object workbook,
                string setupId)
            {
                owner.Events.Add("ownership-load-pending");
                PivotPlusWorkbookMetadata metadata = owner.RecoveryOwnership ??
                    throw new InvalidOperationException();
                return metadata;
            }

            public void MarkActive(
                object workbook,
                string setupId,
                PivotTargetIdentity target,
                PivotDataModelArtifacts artifacts)
            {
                owner.Events.Add("ownership-active");
                owner.OwnershipActivationCalls++;
                if (owner.OwnershipIsActive)
                {
                    throw new InvalidOperationException(
                        "The fake setup is already active.");
                }

                if (owner.ThrowOnOwnershipActivation)
                {
                    throw new InvalidOperationException("metadata activation failed");
                }

                PivotPlusWorkbookMetadata pending = owner.RecoveryOwnership ??
                    throw new InvalidOperationException();
                owner.CommitOwnershipActivation(pending);
                if (owner.ThrowOnOwnershipActivationAfterCommit)
                {
                    throw new PivotPlusOwnershipAmbiguousException(
                        "metadata activation committed but its result was ambiguous",
                        new InvalidOperationException("simulated read-back failure"));
                }
            }
        }

        public PivotStagedDataModelPivot CreateStagedDataModelPivot(
            object workbook,
            string setupId,
            PivotDataModelArtifacts artifacts)
        {
            Events.Add("create-stage");
            if (ThrowOnStageCreation)
            {
                throw new InvalidOperationException("staging cache creation failed");
            }
            StagingPresent = true;
            return new PivotStagedDataModelPivot(
                "_stage",
                "stage-pivot",
                new object(),
                stagedNativePivot,
                new object(),
                artifacts.TemporaryWorksheets.Single(item => item.Purpose == "staging"),
                artifacts.TemporaryWorksheets.Single(item => item.Purpose == "format-backup"),
                artifacts.TemporaryPivotTable);
        }

        public void RestoreState(
            object pivotTable,
            PivotNativeStateSnapshot snapshot,
            string modelTableName)
        {
            Assert.Same(stagedNativePivot, pivotTable);
            Events.Add("restore-stage");
        }

        public void RefreshPivotTable(object pivotTable)
        {
            Events.Add("refresh-stage");
        }

        public string VerifyDataModelState(object pivotTable, PivotNativeStateSnapshot expectedState)
        {
            Events.Add("verify-stage");
            if (ThrowOnStagingVerification)
            {
                throw new InvalidOperationException("staging verification failed");
            }

            return PivotPlusFingerprint.Create("pivotplus.staging-state.v1", "verified");
        }

        public void MarkStagingVerified(
            PivotStagedDataModelPivot stagedPivot,
            string stagingStateFingerprint)
        {
            Events.Add("mark-stage-verified");
        }

        public PivotPendingDataModelRecovery RecoverPending(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            Events.Add("recover-pending");
            PendingRecoveryCalls++;
            if (ThrowOnPendingRecovery)
            {
                throw new InvalidOperationException("pending recovery failed validation");
            }

            var temporary = new PivotTemporaryPivotTableArtifact(
                setupId,
                "_target",
                PivotPlusFingerprint.Create(
                    "pivotplus.temporary-pivot-table.v1",
                    setupId + "\n_target\nSheet1\nPivot1\nA1\nconnection\nmodel"),
                "Sheet1",
                "Pivot1",
                "A1",
                "connection",
                "model");
            return new PivotPendingDataModelRecovery(
                new PivotTargetIdentity("workbook-token", "Sheet1", "Pivot1"),
                new PivotDataModelArtifacts(
                    "query",
                    "connection",
                    "model",
                    QueryFormula,
                    PivotPlusFingerprint.Create("pivotplus.query.v1", QueryFormula),
                    PivotPlusFingerprint.Create("pivotplus.connection.v1", "connection-contract"),
                    new object(),
                    temporaryWorksheets: ownership.Artifacts
                        .Where(item => item.Kind == PivotPlusArtifactKind.TemporaryWorksheet)
                        .Select(item => new PivotTemporaryWorksheetArtifact(
                            item.ArtifactId,
                            item.ArtifactId == "_stage" ? "staging" : "format-backup",
                            item.Fingerprint,
                            "A1"))
                        .ToList(),
                    temporaryPivotTable: temporary));
        }

        public void VerifyActiveDataModelOwnership(
            object workbook,
            string setupId,
            PivotPlusWorkbookMetadata ownership)
        {
            Events.Add("verify-active");
            if (ownership.RecoveryPhase != PivotPlusRecoveryPhase.None)
            {
                throw new InvalidOperationException(
                    "The fake Active ownership is not exact.");
            }
        }

        public IPivotReplacementTransaction PrepareReplacement(
            object workbook,
            object originalPivotTable,
            PivotStagedDataModelPivot stagedPivot,
            PivotNativeStateSnapshot originalState,
            string modelTableName)
        {
            Events.Add("prepare-replacement");
            return new RecordingReplacement(this);
        }

        public void DeleteStagingPivot(object workbook, PivotStagedDataModelPivot stagedPivot)
        {
            Events.Add("delete-stage");
            StageDeleteCalls++;
            if (ThrowOnStageDeletion)
            {
                throw new InvalidOperationException("staging deletion failed");
            }

            StagingPresent = false;
        }

        public void DeleteOwnedModelArtifacts(object workbook, PivotDataModelArtifacts artifacts)
        {
            Events.Add("delete-artifacts");
        }

        private sealed class RecordingReplacement : IPivotReplacementTransaction
        {
            private readonly RecordingGateway owner;

            public RecordingReplacement(RecordingGateway owner)
            {
                this.owner = owner;
            }

            public bool ReplacementAttempted { get; private set; }

            public bool IsCommitted { get; private set; }

            public object? ReplacementPivotTable => ReplacementAttempted ? this : null;

            public void ReplaceAtOriginalLocation()
            {
                ReplacementAttempted = true;
                owner.Events.Add("replace-original");
                if (owner.ThrowDuringReplacement)
                {
                    throw new InvalidOperationException("replacement failed after mutation");
                }
            }

            public void VerifyReplacement()
            {
                owner.Events.Add("verify-replacement");
            }

            public void RollBack()
            {
                owner.Events.Add("rollback-original");
                if (owner.ThrowDuringRollback)
                {
                    throw new InvalidOperationException("rollback failed");
                }
            }

            public void Commit()
            {
                owner.Events.Add("commit");
                IsCommitted = true;
                owner.ReplacementCommitted = true;
                if (owner.ThrowAfterReplacementCommit)
                {
                    throw new InvalidOperationException(
                        "replacement committed but cleanup acknowledgement failed");
                }
            }

            public void Dispose()
            {
                owner.Events.Add("dispose");
            }
        }
    }
}
