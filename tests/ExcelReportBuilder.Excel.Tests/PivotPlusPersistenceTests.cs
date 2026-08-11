using System.Xml.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.DataModel;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;
using FakeWorkbook = ExcelReportBuilder.Excel.Tests.WorkbookSpecStoreTests.FakeWorkbook;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotPlusPersistenceTests
{
    [Fact]
    public void DataModelOwnershipStore_ReturnsExactActiveReceiptForIdempotentRecoveryAcknowledgement()
    {
        var workbook = new FaultingWorkbook();
        var metadata = CreateMetadata("setup_active", "Sheet1", "PivotTable1");
        var metadataStore = new PivotPlusWorkbookMetadataStore();
        workbook.CustomXMLParts.Seed(metadataStore.Serialize(metadata));
        var ownershipStore = new PivotDataModelOwnershipStore();

        PivotPlusWorkbookMetadata loaded =
            ownershipStore.DemandPendingBySetupId(workbook, "setup_active");

        Assert.Equal(PivotPlusRecoveryPhase.None, loaded.RecoveryPhase);
        Assert.Equal("setup_active", loaded.SetupId);
        Assert.DoesNotContain(
            loaded.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.TemporaryWorksheet ||
                    item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
    }

    [Fact]
    public void DataModelOwnershipStore_DeleteCommitsThenInspectionFails_RetryLoadsExactActiveReceipt()
    {
        var workbook = new FaultingWorkbook();
        var metadataStore = new PivotPlusWorkbookMetadataStore();
        const string setupId = "setup_recovery";
        var target = new PivotTargetIdentity(
            "workbook-token",
            "Sheet1",
            "Pivot1");
        var staging = new PivotTemporaryWorksheetArtifact(
            "_stage",
            "staging",
            Fingerprint(
                "pivotplus.temporary-worksheet.v2",
                "staging\n_stage\nA1"),
            "A1");
        var format = new PivotTemporaryWorksheetArtifact(
            "_format",
            "format-backup",
            Fingerprint(
                "pivotplus.temporary-worksheet.v2",
                "format-backup\n_format\nA1"),
            "A1");
        var temporaryPivot = new PivotTemporaryPivotTableArtifact(
            setupId,
            "_target",
            Fingerprint(
                "pivotplus.temporary-pivot-table.v1",
                setupId + "\n_target\nSheet1\nPivot1\nA1\nconnection\nmodel"),
            "Sheet1",
            "Pivot1",
            "A1",
            "connection",
            "model");
        string queryFingerprint = Fingerprint("pivotplus.query.v1", "query");
        string connectionFingerprint = Fingerprint(
            "pivotplus.connection.v1",
            "connection");
        var artifacts = new PivotDataModelArtifacts(
            "query",
            "connection",
            "model",
            "query",
            queryFingerprint,
            connectionFingerprint,
            new object(),
            temporaryWorksheets: new[] { staging, format },
            temporaryPivotTable: temporaryPivot);
        var pending = new PivotPlusWorkbookMetadata
        {
            SetupId = setupId,
            TargetWorksheetName = "Sheet1",
            TargetPivotTableName = "Pivot1",
            RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified,
            TargetAnchorAddress = "A1",
            StagingStateFingerprint = Fingerprint(
                "pivotplus.staging-state.v1",
                "verified"),
            Artifacts = new List<PivotPlusOwnedArtifact>
            {
                new()
                {
                    Kind = PivotPlusArtifactKind.Query,
                    ArtifactId = "query",
                    Fingerprint = queryFingerprint
                },
                new()
                {
                    Kind = PivotPlusArtifactKind.Connection,
                    ArtifactId = "connection",
                    Fingerprint = connectionFingerprint
                },
                new()
                {
                    Kind = staging.Kind,
                    ArtifactId = staging.Name,
                    Fingerprint = staging.Fingerprint
                },
                new()
                {
                    Kind = format.Kind,
                    ArtifactId = format.Name,
                    Fingerprint = format.Fingerprint
                },
                new()
                {
                    Kind = temporaryPivot.Kind,
                    ArtifactId = temporaryPivot.Name,
                    Fingerprint = temporaryPivot.Fingerprint
                }
            }
        };
        workbook.CustomXMLParts.ThrowOnceOnSelectAfterCommittedDelete = true;
        workbook.CustomXMLParts.Seed(
            metadataStore.Serialize(pending),
            throwOnDelete: true,
            removeBeforeThrow: true);
        var ownershipStore = new PivotDataModelOwnershipStore();

        Assert.Throws<InvalidOperationException>(() =>
            ownershipStore.MarkActive(
                workbook,
                setupId,
                target,
                artifacts));

        Assert.Single(workbook.CustomXMLParts.AllXml);
        PivotPlusWorkbookMetadata recovered =
            ownershipStore.DemandPendingBySetupId(workbook, setupId);
        Assert.Equal(PivotPlusRecoveryPhase.None, recovered.RecoveryPhase);
        Assert.Equal(2, recovered.Artifacts.Count);
        Assert.DoesNotContain(
            recovered.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.TemporaryWorksheet ||
                    item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
    }

    [Fact]
    public void Store_round_trips_all_owned_kinds_target_reference_and_bounded_undo()
    {
        var workbook = new FakeWorkbook();
        var metadata = CreateMetadata("setup_1", "Report Sheet", "PivotTable1");
        var createdMeasure = metadata.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.Measure);
        var createdSourceName = metadata.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.WorkbookName);
        metadata.Undo = new PivotPlusUndoMetadata
        {
            ApplyId = "apply_1",
            BeforePivotFingerprint = Fingerprint("pivot-layout", "before"),
            AfterPivotFingerprint = Fingerprint("pivot-layout", "after"),
            CreatedArtifacts = new List<PivotPlusOwnedArtifact>
            {
                Copy(createdMeasure),
                Copy(createdSourceName)
            },
            PreviousFieldPlacements = new List<PivotPlusUndoFieldPlacement>
            {
                new()
                {
                    FieldFingerprint = Fingerprint("pivot-field", "Region"),
                    Area = PivotPlusFieldArea.Row,
                    Position = 0
                },
                new()
                {
                    FieldFingerprint = Fingerprint("pivot-field", "Month"),
                    Area = PivotPlusFieldArea.Column,
                    Position = 0
                }
            }
        };

        var store = new PivotPlusWorkbookMetadataStore();
        store.Save(workbook, metadata);

        var reopened = new FakeWorkbook();
        reopened.CustomXMLParts.Add(Assert.Single(workbook.CustomXMLParts.AllXml));
        var loaded = Assert.Single(store.LoadAll(reopened));

        Assert.Equal("setup_1", loaded.SetupId);
        Assert.Equal("Report Sheet", loaded.TargetWorksheetName);
        Assert.Equal("PivotTable1", loaded.TargetPivotTableName);
        Assert.Equal(
            new[]
            {
                PivotPlusArtifactKind.Measure,
                PivotPlusArtifactKind.NamedSet,
                PivotPlusArtifactKind.Query,
                PivotPlusArtifactKind.Connection,
                PivotPlusArtifactKind.WorkbookName
            },
            loaded.Artifacts.Select(item => item.Kind));
        Assert.NotNull(loaded.Undo);
        Assert.Equal(2, loaded.Undo!.CreatedArtifacts.Count);
        Assert.Contains(
            loaded.Undo.CreatedArtifacts,
            item => item.Kind == PivotPlusArtifactKind.WorkbookName);
        Assert.Equal(2, loaded.Undo.PreviousFieldPlacements.Count);

        // The native target is only a location reference; ownership is limited
        // to one of the explicitly generated artifact kinds.
        Assert.True(store.IsOwnedArtifact(
            reopened,
            loaded.SetupId,
            createdMeasure.Kind,
            createdMeasure.ArtifactId,
            createdMeasure.Fingerprint));
        Assert.DoesNotContain(
            loaded.Artifacts,
            item => string.Equals(item.ArtifactId, "PivotTable1", StringComparison.Ordinal));
    }

    [Fact]
    public void Serialization_is_deterministic_for_equivalent_input_orders()
    {
        var first = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        first.Artifacts = first.Artifacts.Reverse().ToList();
        first.Undo = CreateUndo(first, reverse: true);

        var second = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        second.Undo = CreateUndo(second, reverse: false);

        var store = new PivotPlusWorkbookMetadataStore();

        Assert.Equal(store.Serialize(first), store.Serialize(second));
    }

    [Fact]
    public void Save_replaces_only_the_same_setup_and_leaves_foreign_xml_untouched()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add("<foreign xmlns=\"urn:another-product\" />");
        var store = new PivotPlusWorkbookMetadataStore();
        var first = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        var second = CreateMetadata("setup_2", "Sheet2", "PivotTable2");

        store.Save(workbook, first);
        store.Save(workbook, second);
        first.TargetPivotTableName = "PivotTable3";
        store.Save(workbook, first);

        Assert.Equal(3, workbook.CustomXMLParts.TotalCount);
        Assert.Equal(2, store.LoadAll(workbook).Count);
        Assert.Equal("PivotTable3", store.Load(workbook, "setup_1")!.TargetPivotTableName);
        Assert.Contains(
            workbook.CustomXMLParts.AllXml,
            xml => xml.Contains("urn:another-product", StringComparison.Ordinal));
    }

    [Fact]
    public void Save_preserves_prior_metadata_when_replacement_add_fails()
    {
        var workbook = new FaultingWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        string original = store.Serialize(
            CreateMetadata("setup_1", "Sheet1", "PivotTable1"));
        workbook.CustomXMLParts.Seed(original);
        workbook.CustomXMLParts.ThrowOnAdd = true;

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.Save(
                workbook,
                CreateMetadata("setup_1", "Sheet1", "PivotTable2")));

        Assert.Contains("prior metadata was not removed", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new[] { original }, workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Save_accepts_exact_part_when_add_inserts_then_throws()
    {
        var workbook = new FaultingWorkbook();
        workbook.CustomXMLParts.InsertThenThrow = true;
        var store = new PivotPlusWorkbookMetadataStore();

        store.Save(workbook, CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        Assert.Equal(
            "setup_1",
            Assert.Single(store.LoadAll(workbook)).SetupId);
        Assert.Single(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Save_accepts_exact_part_when_add_inserts_then_returns_null()
    {
        var workbook = new FaultingWorkbook();
        workbook.CustomXMLParts.InsertThenReturnNull = true;
        var store = new PivotPlusWorkbookMetadataStore();

        store.Save(workbook, CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        Assert.Equal(
            "setup_1",
            Assert.Single(store.LoadAll(workbook)).SetupId);
        Assert.Single(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Save_removes_new_part_when_prior_part_delete_fails()
    {
        var workbook = new FaultingWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        string original = store.Serialize(
            CreateMetadata("setup_1", "Sheet1", "PivotTable1"));
        workbook.CustomXMLParts.Seed(original, throwOnDelete: true);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.Save(
                workbook,
                CreateMetadata("setup_1", "Sheet1", "PivotTable2")));

        Assert.Contains("transactionally", exception.Message, StringComparison.Ordinal);
        Assert.Equal(new[] { original }, workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Save_keeps_committed_replacement_when_excel_reports_delete_failure_after_commit()
    {
        var workbook = new FaultingWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        string original = store.Serialize(
            CreateMetadata("setup_1", "Sheet1", "PivotTable1"));
        workbook.CustomXMLParts.Seed(
            original,
            throwOnDelete: true,
            removeBeforeThrow: true);

        var replacement = CreateMetadata("setup_1", "Sheet1", "PivotTable2");
        store.Save(workbook, replacement);

        Assert.Equal("PivotTable2", store.Load(workbook, "setup_1")!.TargetPivotTableName);
        Assert.Single(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Load_and_save_fail_closed_when_metadata_namespace_cannot_be_read()
    {
        var workbook = new FaultingWorkbook();
        workbook.CustomXMLParts.ThrowOnSelect = true;
        var store = new PivotPlusWorkbookMetadataStore();

        Assert.Throws<InvalidOperationException>(() => store.LoadAll(workbook));
        Assert.Throws<InvalidOperationException>(() =>
            store.Save(
                workbook,
                CreateMetadata("setup_1", "Sheet1", "PivotTable1")));
        Assert.Empty(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Store_rejects_case_insensitive_target_collisions_before_changing_workbook()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        store.Save(workbook, CreateMetadata("setup_1", "Report", "PivotTable1"));
        string originalXml = Assert.Single(workbook.CustomXMLParts.AllXml);

        var collision = CreateMetadata("setup_2", "report", "pivottable1");
        var exception = Assert.Throws<InvalidOperationException>(
            () => store.Save(workbook, collision));

        Assert.Contains("same target PivotTable", exception.Message, StringComparison.Ordinal);
        Assert.Equal(originalXml, Assert.Single(workbook.CustomXMLParts.AllXml));
    }

    [Fact]
    public void Store_rejects_generated_artifact_collisions_across_setups()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var first = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        store.Save(workbook, first);

        var collision = CreateMetadata("setup_2", "Sheet2", "PivotTable2");
        collision.Artifacts[0].ArtifactId = first.Artifacts[0].ArtifactId.ToUpperInvariant();
        collision.Artifacts[0].Fingerprint = first.Artifacts[0].Fingerprint;

        var exception = Assert.Throws<InvalidOperationException>(
            () => store.Save(workbook, collision));

        Assert.Contains("same generated artifact", exception.Message, StringComparison.Ordinal);
        Assert.Single(store.LoadAll(workbook));
    }

    [Fact]
    public void Store_rejects_workbook_name_ownership_collisions_across_setups()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var first = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        store.Save(workbook, first);
        PivotPlusOwnedArtifact firstName = first.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.WorkbookName);

        var collision = CreateMetadata("setup_2", "Sheet2", "PivotTable2");
        PivotPlusOwnedArtifact collidingName = collision.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.WorkbookName);
        collidingName.ArtifactId = firstName.ArtifactId.ToUpperInvariant();
        collidingName.Fingerprint = Fingerprint("workbook-name", "different source identity");

        var exception = Assert.Throws<InvalidOperationException>(
            () => store.Save(workbook, collision));

        Assert.Contains("same generated artifact", exception.Message, StringComparison.Ordinal);
        Assert.Single(store.LoadAll(workbook));
    }

    [Fact]
    public void Ownership_requires_exact_kind_id_and_fingerprint()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        store.Save(workbook, metadata);
        var measure = metadata.Artifacts.Single(item => item.Kind == PivotPlusArtifactKind.Measure);

        Assert.True(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            measure.Kind,
            measure.ArtifactId,
            measure.Fingerprint));
        Assert.False(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            PivotPlusArtifactKind.NamedSet,
            measure.ArtifactId,
            measure.Fingerprint));
        Assert.False(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            measure.Kind,
            measure.ArtifactId.ToUpperInvariant(),
            measure.Fingerprint));
        Assert.False(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            measure.Kind,
            measure.ArtifactId,
            Fingerprint("dax-measure", "modified definition")));
    }

    [Fact]
    public void Artifact_ownership_accepts_exact_excel_names_with_spaces_but_rejects_paths()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        PivotPlusOwnedArtifact connection = metadata.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.Connection);
        connection.ArtifactId = "PivotTable Plus Sales Model";
        connection.Fingerprint = Fingerprint("connection", connection.ArtifactId);

        store.Save(workbook, metadata);

        Assert.True(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            connection.Kind,
            connection.ArtifactId,
            connection.Fingerprint));

        connection.ArtifactId = @"C:\Finance\Sales Model";
        Assert.Throws<ArgumentException>(() => store.Serialize(metadata));
    }

    [Fact]
    public void Workbook_name_ownership_requires_exact_id_kind_and_fingerprint()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        store.Save(workbook, metadata);
        PivotPlusOwnedArtifact sourceName = metadata.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.WorkbookName);

        Assert.True(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            PivotPlusArtifactKind.WorkbookName,
            sourceName.ArtifactId,
            sourceName.Fingerprint));
        Assert.False(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            PivotPlusArtifactKind.WorkbookName,
            sourceName.ArtifactId,
            Fingerprint("workbook-name", "changed RefersTo identity")));
        Assert.False(store.IsOwnedArtifact(
            workbook,
            metadata.SetupId,
            PivotPlusArtifactKind.Connection,
            sourceName.ArtifactId,
            sourceName.Fingerprint));
    }

    [Fact]
    public void Store_rejects_duplicate_setup_parts_instead_of_silently_repairing_ownership()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        string xml = store.Serialize(CreateMetadata("setup_1", "Sheet1", "PivotTable1"));
        workbook.CustomXMLParts.Add(xml);
        workbook.CustomXMLParts.Add(xml);

        var exception = Assert.Throws<InvalidOperationException>(() => store.LoadAll(workbook));

        Assert.Contains("same setup identifier", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Store_rejects_unknown_versions_and_unknown_payload_fields()
    {
        var store = new PivotPlusWorkbookMetadataStore();
        string valid = store.Serialize(CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        var unknownVersion = new FakeWorkbook();
        unknownVersion.CustomXMLParts.Add(valid.Replace(
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.CurrentSchemaVersion + "\"",
            "schemaVersion=\"9.0\"",
            StringComparison.Ordinal));
        var versionException = Assert.Throws<NotSupportedException>(
            () => store.LoadAll(unknownVersion));
        Assert.Equal("Unknown PivotTable+ metadata version.", versionException.Message);

        var injectedPayload = new FakeWorkbook();
        var document = XDocument.Parse(valid);
        document.Root!.Add(new XElement(
            XNamespace.Get(PivotPlusWorkbookMetadataStore.NamespaceUri) + "payload",
            "C:\\Finance\\secret.xlsx"));
        injectedPayload.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));

        Assert.Throws<InvalidOperationException>(() => store.LoadAll(injectedPayload));
    }

    [Fact]
    public void Store_reads_version_1_0_and_replaces_it_with_current_metadata()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var legacy = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        legacy.Artifacts = legacy.Artifacts
            .Where(item => item.Kind != PivotPlusArtifactKind.WorkbookName)
            .ToList();
        string legacyXml = store.Serialize(legacy).Replace(
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.CurrentSchemaVersion + "\"",
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.Version1_0 + "\"",
            StringComparison.Ordinal);
        workbook.CustomXMLParts.Add(legacyXml);

        PivotPlusWorkbookMetadata loadedLegacy = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.Version1_0, loadedLegacy.SchemaVersion);
        Assert.DoesNotContain(
            loadedLegacy.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.WorkbookName);

        store.Save(workbook, CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        PivotPlusWorkbookMetadata migrated = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.CurrentSchemaVersion, migrated.SchemaVersion);
        Assert.Contains(
            migrated.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.WorkbookName);
        Assert.Contains(
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.CurrentSchemaVersion + "\"",
            Assert.Single(workbook.CustomXMLParts.AllXml),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Store_reads_version_1_1_and_migrates_it_to_current_1_3()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var legacy = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        string legacyXml = store.Serialize(legacy).Replace(
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.CurrentSchemaVersion + "\"",
            "schemaVersion=\"" + PivotPlusWorkbookMetadata.Version1_1 + "\"",
            StringComparison.Ordinal);
        workbook.CustomXMLParts.Add(legacyXml);

        PivotPlusWorkbookMetadata loaded = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.Version1_1, loaded.SchemaVersion);
        Assert.Contains(
            loaded.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.WorkbookName);

        store.Save(workbook, CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        Assert.Equal(
            PivotPlusWorkbookMetadata.CurrentSchemaVersion,
            Assert.Single(store.LoadAll(workbook)).SchemaVersion);
    }

    [Fact]
    public void Planned_recovery_round_trips_exact_receipts_and_anchor()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var metadata = CreatePendingMetadata(
            "setup_1",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned,
            "C7");

        store.Save(workbook, metadata);

        PivotPlusWorkbookMetadata loaded = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.CurrentSchemaVersion, loaded.SchemaVersion);
        Assert.Equal(PivotPlusRecoveryPhase.Planned, loaded.RecoveryPhase);
        Assert.Equal("C7", loaded.TargetAnchorAddress);
        Assert.Equal(string.Empty, loaded.StagingStateFingerprint);
        Assert.Equal(
            2,
            loaded.Artifacts.Count(item =>
                item.Kind == PivotPlusArtifactKind.TemporaryWorksheet));
        Assert.Single(
            loaded.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
    }

    [Fact]
    public void Staging_verified_recovery_round_trips_checkpoint_hash()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        var metadata = CreatePendingMetadata(
            "setup_1",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.StagingVerified,
            "XFD1048576");

        store.Save(workbook, metadata);

        PivotPlusWorkbookMetadata loaded = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusRecoveryPhase.StagingVerified, loaded.RecoveryPhase);
        Assert.Equal("XFD1048576", loaded.TargetAnchorAddress);
        Assert.Equal(metadata.StagingStateFingerprint, loaded.StagingStateFingerprint);
        Assert.Equal(
            Fingerprint("pivotplus.staging-state.v1", "verified staged state"),
            loaded.StagingStateFingerprint);
        Assert.Contains("phase=\"stagingVerified\"", Assert.Single(workbook.CustomXMLParts.AllXml));
    }

    [Fact]
    public void Pending_recovery_requires_exactly_two_temporary_worksheets_and_one_temporary_pivot()
    {
        var invalidMetadata = new List<PivotPlusWorkbookMetadata>();

        PivotPlusWorkbookMetadata oneWorksheet = CreatePendingMetadata(
            "setup_one_sheet",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        oneWorksheet.Artifacts.Remove(oneWorksheet.Artifacts.First(item =>
            item.Kind == PivotPlusArtifactKind.TemporaryWorksheet));
        invalidMetadata.Add(oneWorksheet);

        PivotPlusWorkbookMetadata threeWorksheets = CreatePendingMetadata(
            "setup_three_sheets",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        threeWorksheets.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryWorksheet,
            "_PP_extra_sheet",
            "pivotplus.temporary-worksheet.v2"));
        invalidMetadata.Add(threeWorksheets);

        PivotPlusWorkbookMetadata noTemporaryPivot = CreatePendingMetadata(
            "setup_no_pivot",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        noTemporaryPivot.Artifacts.Remove(noTemporaryPivot.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.TemporaryPivotTable));
        invalidMetadata.Add(noTemporaryPivot);

        PivotPlusWorkbookMetadata twoTemporaryPivots = CreatePendingMetadata(
            "setup_two_pivots",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        twoTemporaryPivots.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryPivotTable,
            "PP_Target_extra",
            "pivotplus.temporary-pivot-table.v1"));
        invalidMetadata.Add(twoTemporaryPivots);

        var store = new PivotPlusWorkbookMetadataStore();
        foreach (PivotPlusWorkbookMetadata metadata in invalidMetadata)
        {
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                store.Serialize(metadata));
            Assert.Contains(
                "exactly two temporary worksheets and one temporary PivotTable",
                exception.Message,
                StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Recovery_anchor_is_a_strict_bounded_local_A1_address()
    {
        var store = new PivotPlusWorkbookMetadataStore();
        string[] invalidAddresses =
        {
            string.Empty,
            "a1",
            "$A$1",
            "Sheet1!A1",
            "A0",
            "A01",
            "XFE1",
            "A1048577",
            "A1:B2",
            " A1",
            "A1 "
        };

        foreach (string invalidAddress in invalidAddresses)
        {
            PivotPlusWorkbookMetadata metadata = CreatePendingMetadata(
                "setup_1",
                "Sheet1",
                "PivotTable1",
                PivotPlusRecoveryPhase.Planned);
            metadata.TargetAnchorAddress = invalidAddress;

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                store.Serialize(metadata));
            Assert.Contains("strict bounded local A1", exception.Message, StringComparison.Ordinal);
        }

        Assert.Contains(
            "targetAnchor=\"XFD1048576\"",
            store.Serialize(CreatePendingMetadata(
                "setup_1",
                "Sheet1",
                "PivotTable1",
                PivotPlusRecoveryPhase.Planned,
                "XFD1048576")),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Recovery_phase_enforces_the_staging_checkpoint_hash_contract()
    {
        var store = new PivotPlusWorkbookMetadataStore();

        PivotPlusWorkbookMetadata planned = CreatePendingMetadata(
            "setup_planned",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        planned.StagingStateFingerprint = Fingerprint(
            "pivotplus.staging-state.v1",
            "unverified state");
        Assert.Throws<ArgumentException>(() => store.Serialize(planned));

        PivotPlusWorkbookMetadata missingHash = CreatePendingMetadata(
            "setup_missing_hash",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.StagingVerified);
        missingHash.StagingStateFingerprint = string.Empty;
        Assert.Throws<ArgumentException>(() => store.Serialize(missingHash));

        PivotPlusWorkbookMetadata malformedHash = CreatePendingMetadata(
            "setup_malformed_hash",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.StagingVerified);
        malformedHash.StagingStateFingerprint = "not-a-canonical-fingerprint";
        Assert.Throws<ArgumentException>(() => store.Serialize(malformedHash));

        PivotPlusWorkbookMetadata unknownPhase = CreatePendingMetadata(
            "setup_unknown_phase",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        unknownPhase.RecoveryPhase = (PivotPlusRecoveryPhase)999;
        Assert.Throws<ArgumentException>(() => store.Serialize(unknownPhase));
    }

    [Fact]
    public void Active_metadata_forbids_all_recovery_fields_and_temporary_receipts()
    {
        var invalidMetadata = new List<PivotPlusWorkbookMetadata>();

        PivotPlusWorkbookMetadata withAnchor = CreateMetadata(
            "setup_anchor",
            "Sheet1",
            "PivotTable1");
        withAnchor.TargetAnchorAddress = "A1";
        invalidMetadata.Add(withAnchor);

        PivotPlusWorkbookMetadata withHash = CreateMetadata(
            "setup_hash",
            "Sheet1",
            "PivotTable1");
        withHash.StagingStateFingerprint = Fingerprint(
            "pivotplus.staging-state.v1",
            "stale state");
        invalidMetadata.Add(withHash);

        PivotPlusWorkbookMetadata withTemporaryWorksheet = CreateMetadata(
            "setup_sheet",
            "Sheet1",
            "PivotTable1");
        withTemporaryWorksheet.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryWorksheet,
            "_PP_stale_sheet",
            "pivotplus.temporary-worksheet.v2"));
        invalidMetadata.Add(withTemporaryWorksheet);

        PivotPlusWorkbookMetadata withTemporaryPivot = CreateMetadata(
            "setup_pivot",
            "Sheet1",
            "PivotTable1");
        withTemporaryPivot.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryPivotTable,
            "PP_Target_stale",
            "pivotplus.temporary-pivot-table.v1"));
        invalidMetadata.Add(withTemporaryPivot);

        var store = new PivotPlusWorkbookMetadataStore();
        foreach (PivotPlusWorkbookMetadata metadata in invalidMetadata)
        {
            Assert.Equal(PivotPlusRecoveryPhase.None, metadata.RecoveryPhase);
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                store.Serialize(metadata));
            Assert.Contains("Active PivotTable+ metadata", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void Version_1_2_reads_temporary_worksheet_receipts_without_recovery_and_migrates_to_1_3()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        XDocument legacyDocument = XDocument.Parse(store.Serialize(CreatePendingMetadata(
            "setup_1",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned)));
        XNamespace ns = PivotPlusWorkbookMetadataStore.NamespaceUri;
        legacyDocument.Root!.SetAttributeValue(
            "schemaVersion",
            PivotPlusWorkbookMetadata.Version1_2);
        legacyDocument.Root.Element(ns + "recovery")!.Remove();
        legacyDocument.Root.Element(ns + "artifacts")!
            .Elements(ns + "artifact")
            .Single(element =>
                string.Equals(
                    (string?)element.Attribute("kind"),
                    "temporaryPivotTable",
                    StringComparison.Ordinal))
            .Remove();
        workbook.CustomXMLParts.Add(
            legacyDocument.ToString(SaveOptions.DisableFormatting));

        PivotPlusWorkbookMetadata loaded = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.Version1_2, loaded.SchemaVersion);
        Assert.Equal(PivotPlusRecoveryPhase.None, loaded.RecoveryPhase);
        Assert.Equal(string.Empty, loaded.TargetAnchorAddress);
        Assert.Equal(string.Empty, loaded.StagingStateFingerprint);
        Assert.Equal(
            2,
            loaded.Artifacts.Count(item =>
                item.Kind == PivotPlusArtifactKind.TemporaryWorksheet));
        Assert.DoesNotContain(
            loaded.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);

        store.Save(workbook, CreateMetadata("setup_1", "Sheet1", "PivotTable1"));

        PivotPlusWorkbookMetadata migrated = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.CurrentSchemaVersion, migrated.SchemaVersion);
        Assert.Equal(PivotPlusRecoveryPhase.None, migrated.RecoveryPhase);
        Assert.DoesNotContain(
            migrated.Artifacts,
            item => item.Kind == PivotPlusArtifactKind.TemporaryWorksheet ||
                    item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
    }

    [Theory]
    [InlineData(PivotPlusWorkbookMetadata.Version1_0)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_1)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_2)]
    public void Save_migrates_a_loaded_active_legacy_record_to_current_schema(string version)
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        workbook.CustomXMLParts.Add(CreateLegacyActiveXml(store, version));
        PivotPlusWorkbookMetadata legacy = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(version, legacy.SchemaVersion);

        store.Save(workbook, legacy);

        PivotPlusWorkbookMetadata migrated = Assert.Single(store.LoadAll(workbook));
        Assert.Equal(PivotPlusWorkbookMetadata.CurrentSchemaVersion, migrated.SchemaVersion);
        Assert.Equal(PivotPlusRecoveryPhase.None, migrated.RecoveryPhase);
        Assert.Equal(string.Empty, migrated.TargetAnchorAddress);
        Assert.Equal(string.Empty, migrated.StagingStateFingerprint);
    }

    [Theory]
    [InlineData(PivotPlusWorkbookMetadata.Version1_0)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_1)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_2)]
    public void Legacy_versions_reject_recovery_checkpoints(string version)
    {
        var store = new PivotPlusWorkbookMetadataStore();
        XDocument document = XDocument.Parse(CreateLegacyActiveXml(store, version));
        XNamespace ns = PivotPlusWorkbookMetadataStore.NamespaceUri;
        document.Root!.Add(new XElement(
            ns + "recovery",
            new XAttribute("phase", "planned"),
            new XAttribute("targetAnchor", "A1")));
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.LoadAll(workbook));
        Assert.Contains(
            "Managed PivotTable+ metadata could not be read",
            exception.Message,
            StringComparison.Ordinal);
        Assert.Contains(
            "not valid before metadata version 1.3",
            exception.InnerException!.Message,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PivotPlusWorkbookMetadata.Version1_0)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_1)]
    [InlineData(PivotPlusWorkbookMetadata.Version1_2)]
    public void Legacy_versions_reject_temporary_pivot_receipts(string version)
    {
        var store = new PivotPlusWorkbookMetadataStore();
        XDocument document = XDocument.Parse(CreateLegacyActiveXml(store, version));
        XNamespace ns = PivotPlusWorkbookMetadataStore.NamespaceUri;
        document.Root!.Element(ns + "artifacts")!.Add(new XElement(
            ns + "artifact",
            new XAttribute("kind", "temporaryPivotTable"),
            new XAttribute("id", "PP_Target_legacy"),
            new XAttribute(
                "fingerprint",
                Fingerprint("pivotplus.temporary-pivot-table.v1", "legacy receipt"))));
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.LoadAll(workbook));
        Assert.Contains(
            "not valid before metadata version 1.3",
            exception.InnerException!.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Store_rejects_unknown_recovery_phase_and_attributes()
    {
        var store = new PivotPlusWorkbookMetadataStore();
        string valid = store.Serialize(CreatePendingMetadata(
            "setup_1",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned));
        XNamespace ns = PivotPlusWorkbookMetadataStore.NamespaceUri;

        XDocument unknownPhase = XDocument.Parse(valid);
        unknownPhase.Root!.Element(ns + "recovery")!
            .SetAttributeValue("phase", "futurePhase");
        var unknownPhaseWorkbook = new FakeWorkbook();
        unknownPhaseWorkbook.CustomXMLParts.Add(
            unknownPhase.ToString(SaveOptions.DisableFormatting));
        InvalidOperationException phaseException = Assert.Throws<InvalidOperationException>(() =>
            store.LoadAll(unknownPhaseWorkbook));
        Assert.Contains(
            "recovery phase is invalid",
            phaseException.InnerException!.Message,
            StringComparison.Ordinal);

        XDocument unknownAttribute = XDocument.Parse(valid);
        unknownAttribute.Root!.Element(ns + "recovery")!
            .SetAttributeValue("sourceRange", "Secret!A1:B2");
        var unknownAttributeWorkbook = new FakeWorkbook();
        unknownAttributeWorkbook.CustomXMLParts.Add(
            unknownAttribute.ToString(SaveOptions.DisableFormatting));
        InvalidOperationException attributeException = Assert.Throws<InvalidOperationException>(() =>
            store.LoadAll(unknownAttributeWorkbook));
        Assert.Contains(
            "unknown attribute",
            attributeException.InnerException!.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Store_rejects_temporary_pivot_ownership_collisions_across_setups()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        PivotPlusWorkbookMetadata first = CreatePendingMetadata(
            "setup_1",
            "Sheet1",
            "PivotTable1",
            PivotPlusRecoveryPhase.Planned);
        store.Save(workbook, first);

        PivotPlusOwnedArtifact firstTemporaryPivot = first.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
        PivotPlusWorkbookMetadata collision = CreatePendingMetadata(
            "setup_2",
            "Sheet2",
            "PivotTable2",
            PivotPlusRecoveryPhase.Planned);
        PivotPlusOwnedArtifact collidingTemporaryPivot = collision.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
        collidingTemporaryPivot.ArtifactId = firstTemporaryPivot.ArtifactId.ToUpperInvariant();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.Save(workbook, collision));
        Assert.Contains("same generated artifact", exception.Message, StringComparison.Ordinal);
        Assert.Single(store.LoadAll(workbook));
    }

    [Fact]
    public void Version_1_0_cannot_smuggle_workbook_name_ownership()
    {
        var workbook = new FakeWorkbook();
        var store = new PivotPlusWorkbookMetadataStore();
        string invalidLegacyXml = store.Serialize(
                CreateMetadata("setup_1", "Sheet1", "PivotTable1"))
            .Replace(
                "schemaVersion=\"" + PivotPlusWorkbookMetadata.CurrentSchemaVersion + "\"",
                "schemaVersion=\"" + PivotPlusWorkbookMetadata.Version1_0 + "\"",
                StringComparison.Ordinal);
        workbook.CustomXMLParts.Add(invalidLegacyXml);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
            () => store.LoadAll(workbook));

        Assert.Contains(
            "Managed PivotTable+ metadata could not be read",
            exception.Message,
            StringComparison.Ordinal);
        Assert.Contains(
            "not valid in metadata version 1.0",
            exception.InnerException!.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Schema_is_path_free_and_does_not_persist_generated_definitions_or_values()
    {
        const string secretDefinition = "SUM('Sensitive Table'[Salary])";
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        metadata.Artifacts[0].Fingerprint = PivotPlusFingerprint.Create(
            "dax-measure",
            secretDefinition);
        const string sourceReference = "='Finance Data'!$A$1:$F$500";
        PivotPlusOwnedArtifact sourceName = metadata.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.WorkbookName);
        sourceName.Fingerprint = PivotPlusFingerprint.Create(
            "workbook-name",
            sourceReference);

        string xml = new PivotPlusWorkbookMetadataStore().Serialize(metadata);

        Assert.DoesNotContain(secretDefinition, xml, StringComparison.Ordinal);
        Assert.DoesNotContain(sourceReference, xml, StringComparison.Ordinal);
        Assert.DoesNotContain("formula", xml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("refersTo", xml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("workbookPath", xml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("cellValue", xml, StringComparison.OrdinalIgnoreCase);

        metadata.TargetWorksheetName = "C:\\Finance\\Report";
        Assert.Throws<ArgumentException>(
            () => new PivotPlusWorkbookMetadataStore().Serialize(metadata));
    }

    [Fact]
    public void Undo_is_one_level_hash_only_and_bounded()
    {
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        metadata.Undo = CreateUndo(metadata, reverse: false);
        metadata.Undo.PreviousFieldPlacements = Enumerable.Range(0, 257)
            .Select(index => new PivotPlusUndoFieldPlacement
            {
                FieldFingerprint = Fingerprint("pivot-field", index.ToString()),
                Area = PivotPlusFieldArea.Row,
                Position = index
            })
            .ToList();

        var exception = Assert.Throws<ArgumentException>(
            () => new PivotPlusWorkbookMetadataStore().Serialize(metadata));

        Assert.Contains("field-placement limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Undo_can_reference_only_an_exact_current_owned_artifact()
    {
        var metadata = CreateMetadata("setup_1", "Sheet1", "PivotTable1");
        var measure = Copy(metadata.Artifacts[0]);
        measure.Fingerprint = Fingerprint("dax-measure", "not the owned definition");
        metadata.Undo = new PivotPlusUndoMetadata
        {
            ApplyId = "apply_1",
            BeforePivotFingerprint = Fingerprint("pivot-layout", "before"),
            AfterPivotFingerprint = Fingerprint("pivot-layout", "after"),
            CreatedArtifacts = new List<PivotPlusOwnedArtifact> { measure }
        };

        var exception = Assert.Throws<InvalidOperationException>(
            () => new PivotPlusWorkbookMetadataStore().Serialize(metadata));

        Assert.Contains("exactly matching owned artifact", exception.Message, StringComparison.Ordinal);
    }

    private static PivotPlusWorkbookMetadata CreateMetadata(
        string setupId,
        string worksheet,
        string pivotTable)
    {
        return new PivotPlusWorkbookMetadata
        {
            SetupId = setupId,
            TargetWorksheetName = worksheet,
            TargetPivotTableName = pivotTable,
            Artifacts = new List<PivotPlusOwnedArtifact>
            {
                Artifact(PivotPlusArtifactKind.Measure, "measure_portion_" + setupId, "dax-measure"),
                Artifact(PivotPlusArtifactKind.NamedSet, "set_periods_" + setupId, "mdx-set"),
                Artifact(PivotPlusArtifactKind.Query, "query_model_" + setupId, "power-query"),
                Artifact(PivotPlusArtifactKind.Connection, "connection_model_" + setupId, "connection"),
                Artifact(PivotPlusArtifactKind.WorkbookName, "source_name_" + setupId, "workbook-name")
            }
        };
    }

    private static PivotPlusWorkbookMetadata CreatePendingMetadata(
        string setupId,
        string worksheet,
        string pivotTable,
        PivotPlusRecoveryPhase recoveryPhase,
        string targetAnchor = "A1")
    {
        PivotPlusWorkbookMetadata metadata = CreateMetadata(
            setupId,
            worksheet,
            pivotTable);
        metadata.RecoveryPhase = recoveryPhase;
        metadata.TargetAnchorAddress = targetAnchor;
        metadata.StagingStateFingerprint =
            recoveryPhase == PivotPlusRecoveryPhase.StagingVerified
                ? Fingerprint(
                    "pivotplus.staging-state.v1",
                    "verified staged state")
                : string.Empty;
        metadata.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryWorksheet,
            "_PP_stage_" + setupId,
            "pivotplus.temporary-worksheet.v2"));
        metadata.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryWorksheet,
            "_PP_format_" + setupId,
            "pivotplus.temporary-worksheet.v2"));
        metadata.Artifacts.Add(Artifact(
            PivotPlusArtifactKind.TemporaryPivotTable,
            "PP_Target_" + setupId,
            "pivotplus.temporary-pivot-table.v1"));
        return metadata;
    }

    private static string CreateLegacyActiveXml(
        PivotPlusWorkbookMetadataStore store,
        string version)
    {
        PivotPlusWorkbookMetadata metadata = CreateMetadata(
            "setup_legacy",
            "Sheet1",
            "PivotTable1");
        if (string.Equals(
                version,
                PivotPlusWorkbookMetadata.Version1_0,
                StringComparison.Ordinal))
        {
            metadata.Artifacts = metadata.Artifacts
                .Where(item => item.Kind != PivotPlusArtifactKind.WorkbookName)
                .ToList();
        }

        XDocument document = XDocument.Parse(store.Serialize(metadata));
        document.Root!.SetAttributeValue("schemaVersion", version);
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static PivotPlusUndoMetadata CreateUndo(
        PivotPlusWorkbookMetadata metadata,
        bool reverse)
    {
        var created = metadata.Artifacts
            .Where(item =>
                item.Kind == PivotPlusArtifactKind.Measure ||
                item.Kind == PivotPlusArtifactKind.NamedSet)
            .Select(Copy)
            .ToList();
        var placements = new List<PivotPlusUndoFieldPlacement>
        {
            new()
            {
                FieldFingerprint = Fingerprint("pivot-field", "Department"),
                Area = PivotPlusFieldArea.Row,
                Position = 1
            },
            new()
            {
                FieldFingerprint = Fingerprint("pivot-field", "Region"),
                Area = PivotPlusFieldArea.Row,
                Position = 0
            }
        };

        if (reverse)
        {
            created.Reverse();
            placements.Reverse();
        }

        return new PivotPlusUndoMetadata
        {
            ApplyId = "apply_1",
            BeforePivotFingerprint = Fingerprint("pivot-layout", "before"),
            AfterPivotFingerprint = Fingerprint("pivot-layout", "after"),
            CreatedArtifacts = created,
            PreviousFieldPlacements = placements
        };
    }

    private static PivotPlusOwnedArtifact Artifact(
        PivotPlusArtifactKind kind,
        string artifactId,
        string contractKind)
    {
        return new PivotPlusOwnedArtifact
        {
            Kind = kind,
            ArtifactId = artifactId,
            Fingerprint = Fingerprint(contractKind, artifactId + " definition")
        };
    }

    private static PivotPlusOwnedArtifact Copy(PivotPlusOwnedArtifact artifact)
    {
        return new PivotPlusOwnedArtifact
        {
            Kind = artifact.Kind,
            ArtifactId = artifact.ArtifactId,
            Fingerprint = artifact.Fingerprint
        };
    }

    private static string Fingerprint(string contractKind, string value)
    {
        return PivotPlusFingerprint.Create(contractKind, value);
    }

    public sealed class FaultingWorkbook
    {
        public FaultingCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FaultingCustomXmlParts
    {
        private readonly List<FaultingCustomXmlPart> parts = new();

        public bool ThrowOnAdd { get; set; }

        public bool InsertThenThrow { get; set; }

        public bool InsertThenReturnNull { get; set; }

        public bool ThrowOnSelect { get; set; }

        public bool ThrowOnceOnSelectAfterCommittedDelete { get; set; }

        private bool throwOnNextSelect;

        public IReadOnlyList<string> AllXml => parts.Select(item => item.XML).ToList();

        public FaultingCustomXmlPart Seed(
            string xml,
            bool throwOnDelete = false,
            bool removeBeforeThrow = false)
        {
            FaultingCustomXmlPart? holder = null;
            holder = new FaultingCustomXmlPart(
                xml,
                () =>
                {
                    if (removeBeforeThrow)
                    {
                        parts.Remove(holder!);
                        if (ThrowOnceOnSelectAfterCommittedDelete)
                        {
                            throwOnNextSelect = true;
                        }
                    }

                    if (throwOnDelete)
                    {
                        throw new InvalidOperationException("delete failed");
                    }

                    if (!removeBeforeThrow)
                    {
                        parts.Remove(holder!);
                    }
                });
            parts.Add(holder);
            return holder;
        }

        public FaultingCustomXmlPart Add(string xml)
        {
            if (ThrowOnAdd)
            {
                throw new InvalidOperationException("add failed");
            }

            if (InsertThenThrow)
            {
                Seed(xml);
                throw new InvalidOperationException("add failed after insert");
            }

            if (InsertThenReturnNull)
            {
                Seed(xml);
                return null!;
            }

            return Seed(xml);
        }

        public FaultingCustomXmlPartSelection SelectByNamespace(string namespaceUri)
        {
            if (throwOnNextSelect)
            {
                throwOnNextSelect = false;
                throw new InvalidOperationException("selection failed after committed delete");
            }

            if (ThrowOnSelect)
            {
                throw new InvalidOperationException("selection failed");
            }

            return new FaultingCustomXmlPartSelection(parts.Where(part =>
            {
                var document = XDocument.Parse(part.XML);
                return document.Root?.Name.NamespaceName == namespaceUri;
            }).ToList());
        }
    }

    public sealed class FaultingCustomXmlPartSelection
    {
        private readonly IReadOnlyList<FaultingCustomXmlPart> parts;

        public FaultingCustomXmlPartSelection(IReadOnlyList<FaultingCustomXmlPart> parts)
        {
            this.parts = parts;
        }

        public int Count => parts.Count;

        public FaultingCustomXmlPart Item(int index) => parts[index - 1];
    }

    public sealed class FaultingCustomXmlPart
    {
        private readonly Action delete;

        public FaultingCustomXmlPart(string xml, Action delete)
        {
            XML = xml;
            this.delete = delete;
        }

        public string XML { get; }

        public void Delete() => delete();
    }
}
