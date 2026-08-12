using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.PivotPlus.DataModel;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class LateBoundPivotDataModelEnablementGatewayTests
{
    [Fact]
    public void InspectSupportedSource_ResolvesWorkbookTableWithoutUsingAPath()
    {
        var workbook = new FakeWorkbook(
            new[]
            {
                new FakeWorksheet(new[] { new FakeTable("SalesTable") })
            },
            Array.Empty<FakeName>());
        var pivot = new FakePivot(new FakePivotCache(false, "'Data'!SalesTable"));

        ClassicPivotSourceDescriptor source =
            new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot);

        Assert.Equal("SalesTable", source.WorkbookObjectName);
        Assert.Equal(PivotPlusWorkbookObjectKind.Table, source.ObjectKind);
    }

    [Fact]
    public void InspectSupportedSource_ReadsOneBasedSafeArraySourceData()
    {
        Array sourceData = Array.CreateInstance(
            typeof(string),
            new[] { 1 },
            new[] { 1 });
        sourceData.SetValue("'Data'!SalesTable", 1);
        var workbook = new FakeWorkbook(
            new[]
            {
                new FakeWorksheet(new[] { new FakeTable("SalesTable") })
            },
            Array.Empty<FakeName>());
        var pivot = new FakePivot(new FakePivotCache(false, sourceData));

        ClassicPivotSourceDescriptor source =
            new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot);

        Assert.Equal("SalesTable", source.WorkbookObjectName);
    }

    [Fact]
    public void InspectSupportedSource_ResolvesWorkbookScopedNamedRange()
    {
        var worksheet = new FakeWorksheet(Array.Empty<FakeTable>());
        var name = new FakeName("Book.xlsx!SalesRange", new object());
        var workbook = new FakeWorkbook(
            new[] { worksheet },
            new[] { name });
        name.RefersToRange = new FakeNamedRange(worksheet, 10, 5);
        var pivot = new FakePivot(new FakePivotCache(false, "Book.xlsx!SalesRange"));

        ClassicPivotSourceDescriptor source =
            new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot);

        Assert.Equal("SalesRange", source.WorkbookObjectName);
        Assert.Equal(PivotPlusWorkbookObjectKind.NamedRange, source.ObjectKind);
    }

    [Fact]
    public void InspectSupportedSource_RejectsRawCellRange()
    {
        var workbook = new FakeWorkbook(
            new[] { new FakeWorksheet(Array.Empty<FakeTable>()) },
            Array.Empty<FakeName>());
        var pivot = new FakePivot(new FakePivotCache(false, "Data!R1C1:R20C5"));

        Assert.Throws<NotSupportedException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot));
    }

    [Fact]
    public void InspectSupportedSource_ResolvesBoundedR1C1WorksheetRangeWithoutChangingCells()
    {
        var workbook = new FakeRawWorkbook("Data", "$A$1:$E$10", 10, 5);
        var pivot = new FakePivot(new FakePivotCache(false, "Data!R1C1:R10C5"));

        ClassicPivotSourceDescriptor source =
            new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot);

        Assert.True(source.RequiresOwnedWorkbookName);
        Assert.Equal("='Data'!$A$1:$E$10", source.CanonicalReference);
        Assert.Same(workbook.Worksheet.ResolvedRange, source.NativeRange);
        Assert.False(workbook.Worksheet.ResolvedRange.WasWritten);
    }

    [Theory]
    [InlineData("[Other.xlsx]Data!R1C1:R10C5")]
    [InlineData("Data:Other!R1C1:R10C5")]
    [InlineData("Data!R1C1:R10C5,Data!R20C1:R30C5")]
    [InlineData("Data!A:C")]
    [InlineData("Data!$A$1:$XFD$1048576")]
    public void InspectSupportedSource_RejectsUnsafeRawRangeForms(string sourceData)
    {
        var workbook = new FakeRawWorkbook("Data", "$A$1:$E$10", 10, 5);
        var pivot = new FakePivot(new FakePivotCache(false, sourceData));

        Assert.Throws<NotSupportedException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot));
    }

    [Fact]
    public void InspectSupportedSource_RejectsSameAddressOnDifferentWorksheetAfterDiscovery()
    {
        var workbook = new FakeRawWorkbook("Sheet B", "$A$1:$D$20", 20, 4);
        var pivot = new FakePivot(
            new FakePivotCache(false, "'Sheet B'!$A$1:$D$20"));
        var discoveredSource = new PivotSourceDescriptor(
            PivotSourceKind.WorksheetRange,
            "'Sheet A'!$A$1:$D$20",
            PivotCapability.UpgradeToDataModel);

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot, discoveredSource));
    }

    [Fact]
    public void InspectSupportedSource_DoesNotConflateSignificantWorksheetSpaces()
    {
        var workbook = new FakeRawWorkbook("SalesQ1", "$A$1:$D$20", 20, 4);
        var pivot = new FakePivot(
            new FakePivotCache(false, "SalesQ1!$A$1:$D$20"));
        var discoveredSource = new PivotSourceDescriptor(
            PivotSourceKind.WorksheetRange,
            "'Sales Q1'!$A$1:$D$20",
            PivotCapability.UpgradeToDataModel);

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot, discoveredSource));
    }

    [Fact]
    public void InspectSupportedSource_DoesNotConflateDollarInWorksheetNameWithAbsoluteMarker()
    {
        var workbook = new FakeRawWorkbook("SalesQ1", "$A$1:$D$20", 20, 4);
        var pivot = new FakePivot(
            new FakePivotCache(false, "SalesQ1!$A$1:$D$20"));
        var discoveredSource = new PivotSourceDescriptor(
            PivotSourceKind.WorksheetRange,
            "'Sales$Q1'!$A$1:$D$20",
            PivotCapability.UpgradeToDataModel);

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot, discoveredSource));
    }

    [Fact]
    public void RawSourcePlanFingerprint_PreservesDollarAndApostropheInWorksheetIdentity()
    {
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan dollar = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(new object(), "='A$B'!$A$1:$D$9"));
        PivotDataModelArtifactPlan noDollar = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(new object(), "='AB'!$A$1:$D$9"));
        PivotDataModelArtifactPlan apostrophe = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(new object(), "='O''Brien'!$A$1:$D$9"));
        PivotDataModelArtifactPlan noApostrophe = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(new object(), "='OBrien'!$A$1:$D$9"));

        Assert.NotEqual(dollar.WorkbookNameFingerprint, noDollar.WorkbookNameFingerprint);
        Assert.NotEqual(apostrophe.WorkbookNameFingerprint, noApostrophe.WorkbookNameFingerprint);
    }

    [Fact]
    public void InspectSupportedSource_RejectsNonDatabaseClassicCache()
    {
        var workbook = new FakeWorkbook(
            new[] { new FakeWorksheet(Array.Empty<FakeTable>()) },
            Array.Empty<FakeName>());
        var pivot = new FakePivot(new FakePivotCache(false, "Data!A1:D20", 4));

        Assert.Throws<NotSupportedException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot));
    }

    [Fact]
    public void InspectSupportedSource_RejectsExistingOlapPivot()
    {
        var workbook = new FakeWorkbook(
            Array.Empty<FakeWorksheet>(),
            Array.Empty<FakeName>());
        var pivot = new FakePivot(new FakePivotCache(true, "SalesTable"));

        Assert.Throws<NotSupportedException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .InspectSupportedSource(workbook, pivot));
    }

    [Fact]
    public void GeneratedNames_AreDeterministicBoundedAndSetupSpecific()
    {
        GeneratedNames first =
            LateBoundPivotDataModelEnablementGateway.CompileGeneratedNames("setup-1");
        GeneratedNames repeat =
            LateBoundPivotDataModelEnablementGateway.CompileGeneratedNames("setup-1");
        GeneratedNames second =
            LateBoundPivotDataModelEnablementGateway.CompileGeneratedNames("setup-2");

        Assert.Equal(first.QueryName, repeat.QueryName);
        Assert.NotEqual(first.QueryName, second.QueryName);
        Assert.True(first.StagingWorksheetName.Length <= 31);
        Assert.DoesNotContain("\\", first.QueryName, StringComparison.Ordinal);
        Assert.DoesNotContain("/", first.QueryName, StringComparison.Ordinal);
    }

    [Fact]
    public void DemandNoConnectedSlicersOrTimelines_BlocksConnectedPivot()
    {
        var workbook = new FakeSlicerWorkbook(
            new[]
            {
                new FakeSlicerCache(
                    new[] { new FakeSlicerPivotReference("Pivot1") })
            });
        var pivot = new FakeSlicerPivot(
            "Pivot1",
            new FakeSlicerWorksheet(workbook));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoConnectedSlicersOrTimelines(pivot));

        Assert.Contains("slicer or timeline", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DemandNoConnectedSlicersOrTimelines_AllowsUnrelatedSlicer()
    {
        var workbook = new FakeSlicerWorkbook(
            new[]
            {
                new FakeSlicerCache(
                    new[] { new FakeSlicerPivotReference("OtherPivot") })
            });
        var pivot = new FakeSlicerPivot(
            "Pivot1",
            new FakeSlicerWorksheet(workbook));

        LateBoundPivotDataModelEnablementGateway
            .DemandNoConnectedSlicersOrTimelines(pivot);
    }

    [Fact]
    public void DemandNoConnectedSlicersOrTimelines_WhenCollectionUnreadable_FailsClosed()
    {
        var workbook = new FakeUnreadableSlicerWorkbook();
        var pivot = new FakeUnreadableSlicerPivot(
            new FakeUnreadableSlicerWorksheet(workbook));

        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoConnectedSlicersOrTimelines(pivot));
    }

    [Fact]
    public void ReplacementRollback_WhenClearThrowsAndClassicPivotSurvives_RestoresInPlace()
    {
        var cache = new FakeClassicCache();
        var pivot = new FakeRollbackPivot(cache);
        var worksheet = new FakeRollbackWorksheet(pivot);
        var workbook = new FakeRollbackWorkbook(worksheet);
        var state = new LateBoundPivotState(
            cache,
            Array.Empty<LateBoundFieldState>(),
            new LateBoundStyleState(0, true, true, true, string.Empty, true, false, false),
            LateBoundPivotDataModelEnablementGateway.ReadResultSignature(pivot.TableRange2));
        var snapshot = new PivotNativeStateSnapshot(
            "Sheet1",
            "Pivot1",
            "A1",
            "snapshot",
            state);
        PivotTemporaryWorksheetArtifact stagingReceipt =
            TemporaryReceipt("_stage", "staging");
        string stagingFingerprint = PivotPlusFingerprint.Create(
            "pivotplus.staging-state.v1",
            "clear-before-removal");
        var staged = new PivotStagedDataModelPivot(
            "_stage",
            "stage-pivot",
            VerifiedStagingWorksheet(stagingReceipt, stagingFingerprint),
            new object(),
            new FakeReplacementCache(),
            stagingReceipt,
            TemporaryReceipt("_format", "format-backup"),
            TemporaryPivotReceipt("setup-1", "ModelTable"));
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        using IPivotReplacementTransaction replacement = gateway.PrepareReplacement(
            workbook,
            pivot,
            staged,
            snapshot,
            "ModelTable");

        Assert.Throws<InvalidOperationException>(replacement.ReplaceAtOriginalLocation);
        replacement.RollBack();

        Assert.Equal(0, cache.CreateCalls);
        Assert.True(pivot.Refreshed);
        Assert.Equal(new[] { -4122, 8 }, pivot.TableRange2.PasteTypes);
        Assert.True(workbook.Worksheets.BackupWorksheet.Deleted);
    }

    [Fact]
    public void DemandNoAttachedPivotCharts_BlocksChartBoundToTarget()
    {
        var chartPivot = new FakeNamedPivot("Pivot1");
        var chart = new FakeChart(chartPivot);
        var worksheet = new FakeChartWorksheet(
            new[] { new FakeChartObject(chart) });
        var workbook = new FakeChartWorkbook(
            Array.Empty<FakeChart>(),
            new[] { worksheet });
        worksheet.Parent = workbook;
        var pivot = new FakeChartTargetPivot("Pivot1", worksheet);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoAttachedPivotCharts(pivot));

        Assert.Contains("PivotChart", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DemandNoAttachedPivotCharts_BlocksChartOnDifferentWorksheet()
    {
        var targetSheet = new FakeChartWorksheet(Array.Empty<FakeChartObject>());
        var otherSheet = new FakeChartWorksheet(
            new[]
            {
                new FakeChartObject(new FakeChart(new FakeNamedPivot("Pivot1")))
            });
        var workbook = new FakeChartWorkbook(
            Array.Empty<FakeChart>(),
            new[] { targetSheet, otherSheet });
        targetSheet.Parent = workbook;
        otherSheet.Parent = workbook;
        var pivot = new FakeChartTargetPivot("Pivot1", targetSheet);

        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoAttachedPivotCharts(pivot));
    }

    [Fact]
    public void StateVerification_DetectsMemberOrderLoss()
    {
        LateBoundPivotState expected = StateWithField(
            new LateBoundFieldState(
                PivotNativeFieldArea.Row,
                "Region",
                "Region",
                "Region",
                1,
                null,
                string.Empty,
                false,
                Array.Empty<bool>(),
                new[]
                {
                    new LateBoundMemberState("North", "North", true, 1),
                    new LateBoundMemberState("South", "South", true, 2)
                },
                string.Empty,
                false));
        LateBoundPivotState reordered = StateWithField(
            new LateBoundFieldState(
                PivotNativeFieldArea.Row,
                "[Model].[Region]",
                "Region",
                "Region",
                1,
                null,
                string.Empty,
                false,
                Array.Empty<bool>(),
                new[]
                {
                    new LateBoundMemberState("South", "South", true, 1),
                    new LateBoundMemberState("North", "North", true, 2)
                },
                string.Empty,
                false));

        Assert.False(LateBoundPivotState.SemanticallyEquals(expected, reordered));
    }

    [Fact]
    public void StateVerification_DetectsValueBoundToWrongSourceField()
    {
        LateBoundPivotState expected = StateWithField(ValueField("Cost"));
        LateBoundPivotState wrong = StateWithField(ValueField("Revenue"));

        Assert.False(LateBoundPivotState.SemanticallyEquals(expected, wrong));
    }

    [Fact]
    public void CaptureGuard_RejectsShowValuesAsBeforeConversion()
    {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandPlainValueCalculation(new FakeValueCalculation(3)));

        Assert.Contains("Show Values As", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CaptureGuard_RejectsCalculatedFieldBeforeConversion()
    {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoCalculatedOrGroupedField(new FakeCalculatedField()));

        Assert.Contains("Calculated PivotFields", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CaptureGuard_AcceptsOrdinaryExcel2021FieldWithoutIsCalculated()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandNoCalculatedOrGroupedField(new FakeOrdinaryExcel2021Field());
    }

    [Fact]
    public void CaptureGuard_RejectsUnplacedCalculatedFieldInPivotCache()
    {
        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedClassicDefinitions(
                    new FakePivotWithCalculatedDefinition()));

        Assert.Contains("Calculated PivotFields", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_UsesExactWorkbookOnlyModelConnectionShape()
    {
        var workbook = new FakeModelWorkbook();
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        const string formula = "let\n    Source = Excel.CurrentWorkbook(){[Name=\"SalesTable\"]}[Content]\nin\n    Source";

        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor("SalesTable", PivotPlusWorkbookObjectKind.Table));
        PivotDataModelArtifacts artifacts = gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            MetadataFor("setup-1", plan));

        Assert.Equal(artifacts.QueryName, workbook.Queries.AddedName);
        Assert.Equal(formula, workbook.Queries.AddedFormula);
        Assert.Equal(artifacts.ConnectionName, workbook.Connections.AddedName);
        Assert.Equal(2, workbook.Connections.CommandType);
        Assert.True(workbook.Connections.CreateModelConnection);
        Assert.False(workbook.Connections.ImportRelationships);
        Assert.Equal(
            CanonicalConnectionContract.ConnectionString(artifacts.QueryName),
            workbook.Connections.ConnectionString);
        Assert.Equal(
            CanonicalConnectionContract.CommandText(artifacts.QueryName),
            workbook.Connections.CommandText);
        Assert.True(workbook.Connections.Connection.Refreshed);
        Assert.True(workbook.Connections.Connection.OLEDBConnection.BackgroundQuery);
        Assert.Same(
            workbook.Model.DataModelConnection,
            artifacts.NativeDataModelConnection);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_ReusesOnlyExactPendingArtifactsWithoutAdds()
    {
        var workbook = new FakeModelWorkbook();
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        var source = new ClassicPivotSourceDescriptor(
            "SalesTable",
            PivotPlusWorkbookObjectKind.Table);
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            source);
        PivotPlusWorkbookMetadata ownership = MetadataFor(
            "setup-1",
            plan);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);

        PivotDataModelArtifacts recovered = gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            ownership);

        Assert.Same(workbook.Connections.Connection, recovered.NativeConnection);
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
        Assert.Equal(0, workbook.Names.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_WhenQueryChanged_FailsBeforeAnyAdd()
    {
        var workbook = new FakeModelWorkbook();
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        var source = new ClassicPivotSourceDescriptor(
            "SalesTable",
            PivotPlusWorkbookObjectKind.Table);
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            source);
        PivotPlusWorkbookMetadata ownership = MetadataFor(
            "setup-1",
            plan);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);
        workbook.Queries.Query.Formula = "let Source = 7 in Source";

        Assert.Throws<InvalidOperationException>(() =>
            gateway.EnsureOwnedModelArtifacts(
                workbook,
                plan,
                ownership));

        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_ForRawRange_AddsExactHiddenWorkbookAliasBeforeQuery()
    {
        var workbook = new FakeModelWorkbook();
        var source = new ClassicPivotSourceDescriptor(
            new object(),
            "='Data'!$A$1:$E$10");

        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            source);
        PivotDataModelArtifacts artifacts = gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            MetadataFor("setup-1", plan));

        Assert.NotNull(artifacts.OwnedWorkbookName);
        Assert.Equal(
            PivotPlusArtifactKind.WorkbookName,
            artifacts.OwnedWorkbookName!.Kind);
        Assert.Equal(
            artifacts.OwnedWorkbookName.Name,
            workbook.Names.AddedName);
        Assert.Equal("='Data'!$A$1:$E$10", workbook.Names.AddedReference);
        Assert.False(workbook.Names.AddedVisible);
        Assert.Contains(
            "Name=\"" + artifacts.OwnedWorkbookName.Name + "\"",
            workbook.Queries.AddedFormula,
            StringComparison.Ordinal);
        Assert.DoesNotContain("$A$1", workbook.Queries.AddedFormula, StringComparison.Ordinal);
    }

    [Fact]
    public void ValidatePendingFinalization_ForRawRangeAcceptsCanonicalV2NameReceipt()
    {
        var workbook = new FakeModelWorkbook();
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(
                new object(),
                "='O''Brien$'!$A$1:$D$9"));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);

        PivotDataModelArtifacts finalized =
            gateway.ValidatePendingDataModelFinalization(
                workbook,
                new FakeFinalizationPivot(workbook.Model.DataModelConnection),
                "setup-1",
                ownership);

        Assert.NotNull(finalized.OwnedWorkbookName);
        Assert.Equal(
            plan.WorkbookNameFingerprint,
            finalized.OwnedWorkbookName!.ReferenceFingerprint);
        Assert.Equal(1, workbook.Names.AddCalls);
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_WhenConnectionFails_LeavesWriteAheadOwnedPiecesForRetry()
    {
        var workbook = new FakeFailingRawModelWorkbook();
        var source = new ClassicPivotSourceDescriptor(
            new object(),
            "='Data'!$A$1:$E$10");

        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            source);
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);

        Assert.Throws<InvalidOperationException>(() =>
            gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership));

        Assert.False(workbook.Queries.Query.Deleted);
        Assert.False(workbook.Names.Name.Deleted);
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Names.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_QueryInsertThenThrow_RetryCreatesNoDuplicateQuery()
    {
        var workbook = new FakeModelWorkbook();
        workbook.Queries.ThrowAfterCommitOnce = true;
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor("SalesTable", PivotPlusWorkbookObjectKind.Table));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);

        Assert.Throws<InvalidOperationException>(() =>
            gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership));
        gateway.PreflightOwnedModelArtifacts(workbook, plan, ownership);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);

        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_ConnectionInsertThenThrow_RetryCreatesNoDuplicateConnection()
    {
        var workbook = new FakeModelWorkbook();
        workbook.Connections.ThrowAfterCommitOnce = true;
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor("SalesTable", PivotPlusWorkbookObjectKind.Table));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);

        Assert.Throws<InvalidOperationException>(() =>
            gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership));
        gateway.PreflightOwnedModelArtifacts(workbook, plan, ownership);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);

        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void EnsureOwnedModelArtifacts_NameInsertThenThrow_RetryCreatesNoDuplicateName()
    {
        var workbook = new FakeModelWorkbook();
        workbook.Names.ThrowAfterCommitOnce = true;
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(
                new object(),
                "='O''Brien$'!$A$1:$D$9"));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);

        Assert.Throws<InvalidOperationException>(() =>
            gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership));
        gateway.PreflightOwnedModelArtifacts(workbook, plan, ownership);
        gateway.EnsureOwnedModelArtifacts(workbook, plan, ownership);

        Assert.Equal(1, workbook.Names.AddCalls);
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void PreflightOwnedModelArtifacts_NamedCollisionWithoutMetadataIsNeverAdopted()
    {
        var workbook = new FakeModelWorkbook();
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor("SalesTable", PivotPlusWorkbookObjectKind.Table));
        gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            MetadataFor("setup-1", plan));

        Assert.Throws<InvalidOperationException>(() =>
            gateway.PreflightOwnedModelArtifacts(workbook, plan, null));
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
    }

    [Fact]
    public void CreateStagedDataModelPivot_UsesExternalCacheAndHiddenGeneratedSheet()
    {
        var workbook = new FakeStagingWorkbook();
        var sourceConnection = new object();
        var modelConnection = new FakeDataModelConnection();
        var artifacts = new PivotDataModelArtifacts(
            "query",
            "connection",
            "model",
            "formula",
            "query-fingerprint",
            "connection-fingerprint",
            sourceConnection,
            temporaryWorksheets: TemporaryReceipts("setup-1"),
            nativeDataModelConnection: modelConnection,
            temporaryPivotTable: TemporaryPivotReceipt("setup-1", "model"));

        PivotStagedDataModelPivot staged =
            new LateBoundPivotDataModelEnablementGateway()
                .CreateStagedDataModelPivot(workbook, "setup-1", artifacts);

        Assert.Equal(staged.WorksheetName, workbook.Worksheets.AddedWorksheet.Name);
        Assert.Equal(2, workbook.Worksheets.AddedWorksheet.Visible);
        Assert.Equal(2, workbook.PivotCachesValue.SourceType);
        Assert.Same(modelConnection, workbook.PivotCachesValue.SourceData);
        Assert.NotSame(sourceConnection, workbook.PivotCachesValue.SourceData);
        Assert.Equal(6, workbook.PivotCachesValue.Version);
        Assert.Same(
            workbook.Worksheets.AddedWorksheet.Range.Destination,
            workbook.PivotCachesValue.Cache.Destination);
        Assert.Equal(staged.PivotTableName, workbook.PivotCachesValue.Cache.PivotName);
    }

    [Fact]
    public void CreateStagedDataModelPivot_ReusesSingleExactOwnedCache()
    {
        var connection = new object();
        var existingCache = new FakeStagingCache
        {
            WorkbookConnection = connection
        };
        var caches = new FakePivotCaches(new[] { existingCache });
        var workbook = new FakeStagingWorkbook(caches);
        var artifacts = new PivotDataModelArtifacts(
            "query",
            "connection",
            "model",
            "formula",
            "query-fingerprint",
            "connection-fingerprint",
            connection,
            temporaryWorksheets: TemporaryReceipts("setup-1"),
            temporaryPivotTable: TemporaryPivotReceipt("setup-1", "model"));

        PivotStagedDataModelPivot staged =
            new LateBoundPivotDataModelEnablementGateway()
                .CreateStagedDataModelPivot(workbook, "setup-1", artifacts);

        Assert.Same(existingCache, staged.NativePivotCache);
        Assert.Equal(0, caches.CreateCalls);
    }

    [Fact]
    public void CreateStagedDataModelPivot_WhenMultipleModelCachesExist_ReusesFirstWithoutGrowing()
    {
        var connection = new object();
        var caches = new FakePivotCaches(new[]
        {
            new FakeStagingCache { WorkbookConnection = connection },
            new FakeStagingCache { WorkbookConnection = connection }
        });
        var workbook = new FakeStagingWorkbook(caches);
        var artifacts = new PivotDataModelArtifacts(
            "query",
            "connection",
            "model",
            "formula",
            "query-fingerprint",
            "connection-fingerprint",
            connection,
            temporaryWorksheets: TemporaryReceipts("setup-1"),
            temporaryPivotTable: TemporaryPivotReceipt("setup-1", "model"));

        PivotStagedDataModelPivot staged =
            new LateBoundPivotDataModelEnablementGateway()
                .CreateStagedDataModelPivot(workbook, "setup-1", artifacts);

        Assert.Same(caches.Item(1), staged.NativePivotCache);
        Assert.Equal(0, caches.CreateCalls);
    }

    [Fact]
    public void StaleIncompleteStagingWorksheet_WithOrdinaryPayloadIsNeverAdoptedOrDeleted()
    {
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandTemporaryWorksheetStructure(
                new FakeTemporaryPayloadWorksheet("user data"),
                "PP_Stage",
                isFormatBackup: false,
                allowIncomplete: true,
                expectedModelConnection: null));
    }

    [Fact]
    public void DeleteOwnedModelArtifacts_WhenQueryChanged_DeletesNeitherArtifact()
    {
        const string queryName = "PivotPlus_Test_Source";
        const string connectionName = "PivotPlus Test Model";
        const string originalFormula = "let Source = 1 in Source";
        string connectionString = CanonicalConnectionContract.ConnectionString(queryName);
        string commandText = CanonicalConnectionContract.CommandText(queryName);
        var connection = new FakeExactConnection(
            connectionName,
            connectionString,
            commandText);
        var query = new FakeExistingQuery("let Source = 2 in Source");
        var workbook = new FakeCleanupWorkbook(connection, query);
        var artifacts = new PivotDataModelArtifacts(
            queryName,
            connectionName,
            queryName,
            originalFormula,
            PivotPlusFingerprint.Create("pivotplus.query.v1", originalFormula),
            PivotPlusFingerprint.Create(
                "pivotplus.connection.v1",
                connectionString + "\n" + commandText),
            connection);

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .DeleteOwnedModelArtifacts(workbook, artifacts));

        Assert.False(connection.Deleted);
        Assert.False(query.Deleted);
    }

    [Fact]
    public void DeleteOwnedModelArtifacts_WhenRawAliasChanged_DeletesNothing()
    {
        const string queryName = "PivotPlus_Test_Source";
        const string connectionName = "PivotPlus_Test_Model";
        const string queryFormula = "let Source = 1 in Source";
        const string originalReference = "='Data'!$A$1:$E$10";
        string connectionString = CanonicalConnectionContract.ConnectionString(queryName);
        string commandText = CanonicalConnectionContract.CommandText(queryName);
        var connection = new FakeExactConnection(connectionName, connectionString, commandText);
        var query = new FakeExistingQuery(queryFormula);
        var name = new FakeOwnedName("PivotPlus_Test_Range", "='Data'!$A$1:$F$10");
        var workbook = new FakeCleanupWorkbook(connection, query, name);
        var nameReceipt = new PivotOwnedWorkbookNameArtifact(
            "PivotPlus_Test_Range",
            PivotPlusFingerprint.Create("pivotplus.workbook-name.v2", "DATA!A1:E10"),
            originalReference,
            name);
        var artifacts = new PivotDataModelArtifacts(
            queryName,
            connectionName,
            queryName,
            queryFormula,
            PivotPlusFingerprint.Create("pivotplus.query.v1", queryFormula),
            PivotPlusFingerprint.Create(
                "pivotplus.connection.v1",
                connectionString + "\n" + commandText),
            connection,
            nameReceipt);

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway()
                .DeleteOwnedModelArtifacts(workbook, artifacts));

        Assert.False(connection.Deleted);
        Assert.False(query.Deleted);
        Assert.False(name.Deleted);
    }

    [Fact]
    public void VerifyBoundTarget_RejectsWorkbookObjectThatDoesNotOwnPivot()
    {
        var actualWorkbook = new FakeBoundWorkbook("workbook_11111111111111111111111111111111");
        var passedWorkbook = new FakeBoundWorkbook("workbook_11111111111111111111111111111111");
        var pivot = new FakeBoundPivot(
            "Pivot1",
            new FakeBoundWorksheet("Sheet1", actualWorkbook));

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway().VerifyBoundTarget(
                passedWorkbook,
                pivot,
                new PivotTargetIdentity(
                    "workbook_11111111111111111111111111111111",
                    "Sheet1",
                    "Pivot1")));
    }

    [Fact]
    public void VerifyBoundTarget_RejectsChangedPivotNameBeforeArtifacts()
    {
        var workbook = new FakeBoundWorkbook("workbook_11111111111111111111111111111111");
        var pivot = new FakeBoundPivot(
            "OtherPivot",
            new FakeBoundWorksheet("Sheet1", workbook));

        Assert.Throws<InvalidOperationException>(
            () => new LateBoundPivotDataModelEnablementGateway().VerifyBoundTarget(
                workbook,
                pivot,
                new PivotTargetIdentity(
                    "workbook_11111111111111111111111111111111",
                    "Sheet1",
                    "Pivot1")));
    }

    [Fact]
    public void VerifyBoundTarget_RejectsForgedIdentityForUnpersistedWorkbook()
    {
        var workbook = new FakeUnpersistedBoundWorkbook();
        var pivot = new FakeUnpersistedBoundPivot(
            new FakeUnpersistedBoundWorksheet(workbook));

        Assert.Throws<InvalidOperationException>(() =>
            new LateBoundPivotDataModelEnablementGateway().VerifyBoundTarget(
                workbook,
                pivot,
                new PivotTargetIdentity(
                    "workbook_22222222222222222222222222222222",
                    "Sheet1",
                    "Pivot1")));
        Assert.Equal(0, workbook.CustomXMLParts.AddCalls);
    }

    [Fact]
    public void ReadFieldStates_WhenRequiredAxisCollectionUnreadable_FailsClosed()
    {
        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway.ReadFieldStates(
                new FakeUnreadableAxisPivot()));
    }

    [Fact]
    public void StateVerification_DetectsResultDimensionMismatch()
    {
        LateBoundPivotState expected = StateWithResult(2, 3, "same");
        LateBoundPivotState actual = StateWithResult(3, 2, "same");

        Assert.False(LateBoundPivotState.SemanticallyEquals(expected, actual));
    }

    [Fact]
    public void StateVerification_DetectsResultValueMismatch()
    {
        LateBoundPivotState expected = StateWithResult(2, 3, "first");
        LateBoundPivotState actual = StateWithResult(2, 3, "second");

        Assert.False(LateBoundPivotState.SemanticallyEquals(expected, actual));
    }

    [Fact]
    public void CaptureGuard_RejectsAggregationUnsupportedByGetMeasure()
    {
        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandSupportedModelAggregation(new FakeValueFunction(-4149)));
    }

    [Fact]
    public void RestorePageFilter_UsesCurrentPageNameForModelAndPivotItemForClassic()
    {
        var modelTarget = new FakePageField();
        var classicTarget = new FakePageField();
        LateBoundFieldState state = PageFieldState();

        LateBoundPivotDataModelEnablementGateway.ApplyRegularFieldState(
            modelTarget,
            state,
            isDataModel: true);
        LateBoundPivotDataModelEnablementGateway.ApplyRegularFieldState(
            classicTarget,
            state,
            isDataModel: false);

        Assert.Equal("[Model].[Region].&[North]", modelTarget.CurrentPageName);
        Assert.True(modelTarget.CubeField.AllItemsVisible);
        Assert.Null(modelTarget.CurrentPage);
        Assert.Same(classicTarget.Item, classicTarget.CurrentPage);
        Assert.Equal(string.Empty, classicTarget.CurrentPageName);
    }

    [Fact]
    public void RestoreFilteredOlapMembers_UsesVisibleItemsListAndNeverWritesPivotItemVisible()
    {
        var target = new FakeOlapFilterField();
        var state = new LateBoundFieldState(
            PivotNativeFieldArea.Row,
            "Region",
            "Region",
            "Region",
            1,
            null,
            string.Empty,
            false,
            Array.Empty<bool>(),
            new[]
            {
                new LateBoundMemberState("North", "North", true, 1),
                new LateBoundMemberState("South", "South", false, 2)
            },
            string.Empty,
            false);

        LateBoundPivotDataModelEnablementGateway.ApplyRegularFieldState(
            target,
            state,
            isDataModel: true);

        Assert.False(target.DatabaseSort);
        Assert.Equal(
            new[] { "[Model].[Region].&[North]" },
            target.VisibleItemsList);
        Assert.Equal(0, target.North.VisibleSetCalls);
        Assert.Equal(0, target.South.VisibleSetCalls);
    }

    [Fact]
    public void CaptureGuard_RejectsUnreadablePivotFilters()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandNoNativePivotFilters(
                new FakeUnreadablePivotFiltersField()));
    }

    [Fact]
    public void CaptureGuard_AcceptsExcel2021NoPivotFiltersComResult()
    {
        LateBoundPivotDataModelEnablementGateway.DemandNoNativePivotFilters(
            new FakeNoPivotFiltersField());
    }

    [Fact]
    public void CaptureGuard_RejectsAutomaticallySortedClassicField()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandClassicManualSort(
                new FakeAutoSortedField()));
    }

    [Fact]
    public void CaptureGuard_RejectsClassicShowAllItemsBecauseOlapForcesFalse()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandClassicNoShowAllItems(
                new FakeShowAllItemsField(true)));
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandClassicNoShowAllItems(
                new FakeUnreadableShowAllItemsField()));
    }

    [Fact]
    public void CaptureGuard_RejectsIncludeNewItemsFilterPolicyAndUnreadableState()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandClassicDefaultIncludeNewItemsInFilter(
                new FakeIncludeNewItemsField(false));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandClassicDefaultIncludeNewItemsInFilter(
                    new FakeIncludeNewItemsField(true)));
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandClassicDefaultIncludeNewItemsInFilter(
                    new FakeUnreadableIncludeNewItemsField()));
    }

    [Fact]
    public void CaptureGuard_RejectsNondefaultOrUnreadableClassicCachePolicy()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandCompatibleClassicCachePolicy(
                new FakeCachePolicyPivot(new FakeCachePolicy()));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandCompatibleClassicCachePolicy(
                    new FakeCachePolicyPivot(
                        new FakeCachePolicy(refreshOnFileOpen: true))));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandCompatibleClassicCachePolicy(
                    new FakeCachePolicyPivot(
                        new FakeCachePolicy(enableRefresh: false))));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandCompatibleClassicCachePolicy(
                    new FakeCachePolicyPivot(
                        new FakeCachePolicy(missingItemsLimit: 0))));
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandCompatibleClassicCachePolicy(
                    new FakeCachePolicyPivot(
                        new FakeUnreadableCachePolicy())));
    }

    [Fact]
    public void CaptureGuard_NormalizesReadableClassicSaveDataCategoryInvariant()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandCompatibleClassicSaveData(
                new FakeSaveDataPivot(false));
        LateBoundPivotDataModelEnablementGateway
            .DemandCompatibleClassicSaveData(
                new FakeSaveDataPivot(true));
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandCompatibleClassicSaveData(
                    new FakeUnreadableSaveDataPivot()));
    }

    [Fact]
    public void OlapInvariantGuard_RejectsDisabledDrilldownAndPageSubtotalMismatch()
    {
        var disabledDrilldown = new LateBoundStyleState(
            0,
            true,
            true,
            true,
            string.Empty,
            true,
            false,
            false,
            enableDrilldown: false);
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandCompatibleOlapInvariants(
                Array.Empty<LateBoundFieldState>(),
                disabledDrilldown));

        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandCompatibleOlapInvariants(
                new[] { PageFieldState() },
                new LateBoundStyleState(
                    0,
                    true,
                    true,
                    true,
                    string.Empty,
                    true,
                    false,
                    false,
                    subtotalHiddenPageItems: false)));
    }

    [Fact]
    public void OlapInvariantGuard_RejectsCustomSubtotalSlotsBeforeConversion()
    {
        var subtotals = new bool[12];
        subtotals[1] = true;
        var field = new LateBoundFieldState(
            PivotNativeFieldArea.Row,
            "Region",
            "Region",
            "Region",
            1,
            null,
            string.Empty,
            false,
            subtotals,
            Array.Empty<LateBoundMemberState>(),
            string.Empty,
            false);

        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway.DemandCompatibleOlapInvariants(
                new[] { field },
                new LateBoundStyleState(
                    0,
                    true,
                    true,
                    true,
                    string.Empty,
                    true,
                    false,
                    false)));
    }

    [Fact]
    public void StateVerification_NormalizesDocumentedOlapInvariantValues()
    {
        var expected = new LateBoundPivotState(
            new object(),
            Array.Empty<LateBoundFieldState>(),
            new LateBoundStyleState(
                0,
                true,
                true,
                true,
                string.Empty,
                true,
                false,
                false,
                enableDrilldown: true,
                subtotalHiddenPageItems: false));
        var actual = new LateBoundPivotState(
            new object(),
            Array.Empty<LateBoundFieldState>(),
            new LateBoundStyleState(
                0,
                true,
                true,
                true,
                string.Empty,
                true,
                false,
                false,
                enableDrilldown: true,
                subtotalHiddenPageItems: true));

        Assert.False(LateBoundPivotState.SemanticallyEquals(expected, actual));
        Assert.True(LateBoundPivotState.SemanticallyEquals(
            expected,
            actual,
            normalizeDataModelInvariants: true));
    }

    [Fact]
    public void FormattingGuard_AllowsNormalPreserveFormattingSetting()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandNoUnsupportedCustomFormatting(
                new FakeFormattingPivot(true),
                new FakeFormattingRange(0));
    }

    [Fact]
    public void FormattingGuard_RejectsConditionalFormatsBeforeArtifacts()
    {
        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCustomFormatting(
                    new FakeFormattingPivot(false),
                    new FakeFormattingRange(1)));
    }

    [Fact]
    public void CellMetadataGuard_AllowsEmptyReadableState()
    {
        LateBoundPivotDataModelEnablementGateway
            .DemandNoUnsupportedCellMetadata(new FakeCellMetadataRange());
    }

    [Fact]
    public void CellMetadataGuard_RejectsNotesAndThreadedCommentsInPivotResult()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(legacyComment: new FakeCellComment(2, 3))));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(threadedComment: new FakeCellComment(2, 3))));
    }

    [Fact]
    public void CellMetadataGuard_RejectsDataValidationBeforeArtifacts()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(hasValidation: true)));
    }

    [Fact]
    public void CellMetadataGuard_RejectsHyperlinksAndUnreadableHyperlinkCollection()
    {
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(hyperlinkCount: 1)));
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(unreadableHyperlinks: true)));
    }

    [Fact]
    public void CellMetadataGuard_FailsClosedOnUnreadableCommentOrValidationState()
    {
        Assert.Throws<InvalidOperationException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(unreadableComments: true)));
        Assert.Throws<NotSupportedException>(() =>
            LateBoundPivotDataModelEnablementGateway
                .DemandNoUnsupportedCellMetadata(
                    new FakeCellMetadataRange(unreadableValidation: true)));
    }

    [Fact]
    public void ApplyStyleState_ExplicitlyClearsDefaultStyleWhenSourceHasNoStyle()
    {
        var target = new FakeStyleTarget
        {
            TableStyle2 = "PivotStyleMedium2"
        };

        LateBoundPivotDataModelEnablementGateway.ApplyStyleState(
            target,
            new LateBoundStyleState(
                0,
                true,
                true,
                true,
                string.Empty,
                true,
                false,
                false));

        Assert.Equal(string.Empty, target.TableStyle2);
    }

    [Fact]
    public void DuplicateImplicitValueGuard_RejectsSameSourceAndFunctionTwice()
    {
        LateBoundFieldState first = ValueField("Cost");
        LateBoundFieldState second = new LateBoundFieldState(
            PivotNativeFieldArea.Values,
            "Cost",
            "Cost Again",
            "Cost Again",
            2,
            -4157,
            "#,##0",
            false,
            Array.Empty<bool>(),
            Array.Empty<LateBoundMemberState>(),
            string.Empty,
            false);

        Assert.Throws<NotSupportedException>(
            () => LateBoundPivotDataModelEnablementGateway
                .DemandNoDuplicateImplicitValues(new[] { first, second }));
    }

    [Fact]
    public void ReplacementTransaction_ClearsCreatesAndRollsBackInSafeOrder()
    {
        var events = new List<string>();
        var lookup = new FakeTransactionalPivotLookup();
        var classicCache = new FakeTransactionalCache(false, lookup, events, "classic");
        var modelCache = new FakeTransactionalCache(true, lookup, events, "model")
        {
            FailPromotion = true
        };
        var original = new FakeTransactionalPivot(classicCache, lookup, events, "original");
        lookup.Current = original;
        var worksheet = new FakeTransactionalWorksheet(lookup);
        var workbook = new FakeTransactionalWorkbook(worksheet);
        var snapshot = new PivotNativeStateSnapshot(
            "Sheet1",
            "Pivot1",
            "A1",
            "snapshot",
            new LateBoundPivotState(
                classicCache,
                Array.Empty<LateBoundFieldState>(),
                new LateBoundStyleState(0, true, true, true, string.Empty, true, false, false),
                LateBoundPivotDataModelEnablementGateway.ReadResultSignature(original.TableRange2)));
        PivotTemporaryWorksheetArtifact stagingReceipt =
            TemporaryReceipt("_stage", "staging");
        string stagingFingerprint = DataModelFingerprint(
            (LateBoundPivotState)snapshot.NativeState);
        var staged = new PivotStagedDataModelPivot(
            "_stage",
            "stage-pivot",
            VerifiedStagingWorksheet(stagingReceipt, stagingFingerprint),
            new object(),
            modelCache,
            stagingReceipt,
            TemporaryReceipt("_format", "format-backup"),
            TemporaryPivotReceipt(
                "setup-1",
                "ModelTable",
                name: "PP_Target_test"));
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        using IPivotReplacementTransaction transaction = gateway.PrepareReplacement(
            workbook,
            original,
            staged,
            snapshot,
            "ModelTable");

        Assert.Throws<InvalidOperationException>(transaction.ReplaceAtOriginalLocation);
        transaction.RollBack();

        Assert.Equal(
            new[]
            {
                "clear-original",
                "create-model",
                "refresh-model",
                "clear-model",
                "create-classic",
                "refresh-classic"
            },
            events);
        Assert.NotNull(lookup.Current);
        Assert.False(lookup.Current!.PivotCache().OLAP);
        Assert.Equal(new[] { -4122, 8 }, lookup.Current.TableRange2.PasteTypes);
        Assert.Equal(2, workbook.Worksheets.BackupWorksheet.Range.Anchor.CopyCalls);
        Assert.True(workbook.Worksheets.BackupWorksheet.Deleted);
    }

    [Fact]
    public void ReplacementRollback_WhenClassicRestoreFails_RetainsOwnedFormatBackup()
    {
        var events = new List<string>();
        var lookup = new FakeTransactionalPivotLookup();
        var classicCache = new FakeTransactionalCache(false, lookup, events, "classic")
        {
            FailCreatedRestore = true
        };
        var modelCache = new FakeTransactionalCache(true, lookup, events, "model")
        {
            FailPromotion = true
        };
        var original = new FakeTransactionalPivot(classicCache, lookup, events, "original");
        lookup.Current = original;
        var workbook = new FakeTransactionalWorkbook(
            new FakeTransactionalWorksheet(lookup));
        var snapshot = new PivotNativeStateSnapshot(
            "Sheet1",
            "Pivot1",
            "A1",
            "snapshot",
            new LateBoundPivotState(
                classicCache,
                Array.Empty<LateBoundFieldState>(),
                new LateBoundStyleState(0, true, true, true, string.Empty, true, false, false),
                LateBoundPivotDataModelEnablementGateway.ReadResultSignature(original.TableRange2)));
        PivotTemporaryWorksheetArtifact stagingReceipt =
            TemporaryReceipt("_stage", "staging");
        string stagingFingerprint = DataModelFingerprint(
            (LateBoundPivotState)snapshot.NativeState);
        var staged = new PivotStagedDataModelPivot(
            "_stage",
            "stage-pivot",
            VerifiedStagingWorksheet(stagingReceipt, stagingFingerprint),
            new object(),
            modelCache,
            stagingReceipt,
            TemporaryReceipt("_format", "format-backup"),
            TemporaryPivotReceipt(
                "setup-1",
                "ModelTable",
                name: "PP_Target_test"));
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        using IPivotReplacementTransaction transaction = gateway.PrepareReplacement(
            workbook,
            original,
            staged,
            snapshot,
            "ModelTable");

        Assert.Throws<InvalidOperationException>(transaction.ReplaceAtOriginalLocation);
        Assert.Throws<InvalidOperationException>(() => transaction.RollBack());

        Assert.False(workbook.Worksheets.BackupWorksheet.Deleted);
    }

    [Fact]
    public void RecoverPending_FinalVerifiedTargetWithNoTemporarySheets_IsIdempotentAndDoesNotRecreateObjects()
    {
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        var workbook = new FakeModelWorkbook();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(
                "SalesTable",
                PivotPlusWorkbookObjectKind.Table));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);
        PivotDataModelArtifacts artifacts = gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            ownership);

        var events = new List<string>();
        var lookup = new FakeTransactionalPivotLookup();
        var cache = new FakeTransactionalCache(
            true,
            lookup,
            events,
            "model",
            artifacts.NativeDataModelConnection);
        var target = new FakeTransactionalPivot(cache, lookup, events, "model");
        lookup.Current = target;
        var worksheet = new FakeTransactionalWorksheet(lookup);
        workbook.Worksheets = new FakeCollection<object>(new object[] { worksheet });
        var liveState = new LateBoundPivotState(
            cache,
            Array.Empty<LateBoundFieldState>(),
            new LateBoundStyleState(
                0,
                true,
                true,
                true,
                string.Empty,
                true,
                false,
                false,
                subtotalHiddenPageItems: true),
            LateBoundPivotDataModelEnablementGateway.ReadResultSignature(
                target.TableRange2));
        ownership.RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified;
        ownership.StagingStateFingerprint = DataModelFingerprint(liveState);

        PivotPendingDataModelRecovery first = gateway.RecoverPending(
            workbook,
            "setup-1",
            ownership);
        PivotPendingDataModelRecovery second = gateway.RecoverPending(
            workbook,
            "setup-1",
            ownership);

        Assert.Equal("Sheet1", first.Target.WorksheetName);
        Assert.Equal("Pivot1", second.Target.PivotTableName);
        Assert.Equal(1, workbook.Queries.AddCalls);
        Assert.Equal(1, workbook.Connections.AddCalls);
        Assert.DoesNotContain("clear-model", events);
        Assert.DoesNotContain("create-model", events);
    }

    [Fact]
    public void RecoverPending_OriginalNameModelTargetWithChangedState_FailsWithoutClearingIt()
    {
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        var workbook = new FakeModelWorkbook();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(
                "SalesTable",
                PivotPlusWorkbookObjectKind.Table));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);
        PivotDataModelArtifacts artifacts = gateway.EnsureOwnedModelArtifacts(
            workbook,
            plan,
            ownership);
        var events = new List<string>();
        var lookup = new FakeTransactionalPivotLookup();
        var cache = new FakeTransactionalCache(
            true,
            lookup,
            events,
            "model",
            artifacts.NativeDataModelConnection);
        lookup.Current = new FakeTransactionalPivot(cache, lookup, events, "model");
        workbook.Worksheets = new FakeCollection<object>(new object[]
        {
            new FakeTransactionalWorksheet(lookup)
        });
        ownership.RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified;
        ownership.StagingStateFingerprint = PivotPlusFingerprint.Create(
            "pivotplus.staging-state.v1",
            "changed-state");

        Assert.Throws<InvalidOperationException>(() => gateway.RecoverPending(
            workbook,
            "setup-1",
            ownership));

        Assert.NotNull(lookup.Current);
        Assert.Equal("Pivot1", lookup.Current!.Name);
        Assert.DoesNotContain("clear-model", events);
    }

    [Fact]
    public void RecoverPending_PlannedCheckpointWithExactClassicTarget_RequiresOrdinaryRetryWithoutMutation()
    {
        var gateway = new LateBoundPivotDataModelEnablementGateway();
        var workbook = new FakeModelWorkbook();
        PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
            "setup-1",
            new ClassicPivotSourceDescriptor(
                "SalesTable",
                PivotPlusWorkbookObjectKind.Table));
        PivotPlusWorkbookMetadata ownership = MetadataFor("setup-1", plan);
        var events = new List<string>();
        var lookup = new FakeTransactionalPivotLookup();
        var cache = new FakeTransactionalCache(false, lookup, events, "classic");
        lookup.Current = new FakeTransactionalPivot(cache, lookup, events, "classic");
        workbook.Worksheets = new FakeCollection<object>(new object[]
        {
            new FakeTransactionalWorksheet(lookup)
        });

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(
            () => gateway.RecoverPending(workbook, "setup-1", ownership));

        Assert.Contains("Re-run Data Model enablement", failure.Message, StringComparison.Ordinal);
        Assert.NotNull(lookup.Current);
        Assert.DoesNotContain("clear-classic", events);
    }

    [Fact]
    public void RecoverPending_AbsentTarget_RehydratesFormatAndPromotesThenConvergesOnSecondCall()
    {
        var fixture = new RecoveryFixture(RecoveryTargetMode.Absent);

        PivotPendingDataModelRecovery first = fixture.Recover();
        int createsAfterFirst = fixture.Controller.Events.Count(item => item == "create-target");
        PivotPendingDataModelRecovery second = fixture.Recover();

        Assert.Equal("Pivot1", first.Target.PivotTableName);
        Assert.Equal("Pivot1", second.Target.PivotTableName);
        Assert.Equal(1, createsAfterFirst);
        Assert.Equal(createsAfterFirst, fixture.Controller.Events.Count(item => item == "create-target"));
        Assert.True(fixture.TargetPivotIsFinalAndExact());
        Assert.False(fixture.Worksheets.Contains(fixture.Names.FormatBackupWorksheetName));
        Assert.False(fixture.Worksheets.Contains(fixture.Names.StagingWorksheetName));
        Assert.True(
            fixture.Controller.Events.IndexOf("delete-format") <
            fixture.Controller.Events.IndexOf("delete-stage"));
        Assert.Contains("paste-formats", fixture.Controller.Events);
        Assert.Contains("restore-dimension", fixture.Controller.Events);
    }

    [Fact]
    public void RecoverPending_ExactGeneratedPartialTarget_IsClearedAndRecreatedFromCheckpoint()
    {
        var fixture = new RecoveryFixture(RecoveryTargetMode.GeneratedPartial);

        fixture.Recover();

        Assert.True(fixture.TargetPivotIsFinalAndExact());
        Assert.Equal(1, fixture.Controller.Events.Count(item => item == "clear-target"));
        Assert.Equal(1, fixture.Controller.Events.Count(item => item == "create-target"));
    }

    public static IEnumerable<object[]> RecoveryCrashPoints()
    {
        foreach (RecoveryFailurePoint point in Enum.GetValues(typeof(RecoveryFailurePoint)))
        {
            if (point != RecoveryFailurePoint.None)
            {
                yield return new object[] { point };
            }
        }
    }

    [Theory]
    [MemberData(nameof(RecoveryCrashPoints))]
    public void RecoverPending_ComBoundaryFailure_RetryConvergesWithoutGrowingOwnedObjects(
        RecoveryFailurePoint point)
    {
        RecoveryTargetMode mode =
            point == RecoveryFailurePoint.BeforeClear ||
            point == RecoveryFailurePoint.AfterClear
                ? RecoveryTargetMode.GeneratedPartial
                : RecoveryTargetMode.Absent;
        var fixture = new RecoveryFixture(mode);
        fixture.Controller.FailurePoint = point;

        _ = Record.Exception(fixture.Recover);
        PivotPendingDataModelRecovery recovered = fixture.Recover();
        fixture.Recover();

        Assert.Equal("Pivot1", recovered.Target.PivotTableName);
        Assert.True(fixture.TargetPivotIsFinalAndExact());
        Assert.False(fixture.Worksheets.Contains(fixture.Names.FormatBackupWorksheetName));
        Assert.False(fixture.Worksheets.Contains(fixture.Names.StagingWorksheetName));
        Assert.Equal(1, fixture.Workbook.Queries.AddCalls);
        Assert.Equal(1, fixture.Workbook.Connections.AddCalls);
        Assert.True(fixture.Controller.Events.Count(item => item == "create-target") <= 2);
        Assert.DoesNotContain("clear-final-target", fixture.Controller.Events);
    }

    [Theory]
    [InlineData(RecoveryContamination.Values)]
    [InlineData(RecoveryContamination.Formulas)]
    [InlineData(RecoveryContamination.Note)]
    [InlineData(RecoveryContamination.Validation)]
    [InlineData(RecoveryContamination.Hyperlink)]
    [InlineData(RecoveryContamination.Merge)]
    [InlineData(RecoveryContamination.Table)]
    [InlineData(RecoveryContamination.NeighborPivot)]
    public void RecoverPending_ContaminatedAbsentDestination_FailsBeforeCreate(
        RecoveryContamination contamination)
    {
        var fixture = new RecoveryFixture(RecoveryTargetMode.Absent);
        fixture.Contaminate(contamination);

        Assert.ThrowsAny<Exception>(fixture.Recover);

        Assert.False(fixture.TargetPivotIsFinalAndExact());
        Assert.DoesNotContain("create-target", fixture.Controller.Events);
        Assert.True(fixture.Worksheets.Contains(fixture.Names.FormatBackupWorksheetName));
        Assert.True(fixture.Worksheets.Contains(fixture.Names.StagingWorksheetName));
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public void RecoverPending_FinalTarget_CleansOnlyExactRemainingTempsAndIsIdempotent(
        bool keepFormat,
        bool keepStage)
    {
        var fixture = new RecoveryFixture(RecoveryTargetMode.Final);
        if (!keepFormat) fixture.RemoveFormatSheet();
        if (!keepStage) fixture.RemoveStageSheet();

        fixture.Recover();
        fixture.Recover();

        Assert.True(fixture.TargetPivotIsFinalAndExact());
        Assert.False(fixture.Worksheets.Contains(fixture.Names.FormatBackupWorksheetName));
        Assert.False(fixture.Worksheets.Contains(fixture.Names.StagingWorksheetName));
        if (keepFormat && keepStage)
        {
            Assert.True(
                fixture.Controller.Events.IndexOf("delete-format") <
                fixture.Controller.Events.IndexOf("delete-stage"));
        }
    }

    [Fact]
    public void RecoverPending_ReceiptAnchorAndModelLineageMismatches_FailWithoutTargetMutation()
    {
        var receiptFixture = new RecoveryFixture(RecoveryTargetMode.Absent);
        PivotPlusOwnedArtifact receipt = receiptFixture.Ownership.Artifacts.Single(item =>
            item.Kind == PivotPlusArtifactKind.TemporaryPivotTable);
        receipt.Fingerprint = PivotPlusFingerprint.Create(
            "pivotplus.temporary-pivot-table.v1",
            "changed");
        Assert.Throws<InvalidOperationException>(receiptFixture.Recover);
        Assert.DoesNotContain("create-target", receiptFixture.Controller.Events);

        var anchorFixture = new RecoveryFixture(RecoveryTargetMode.Absent);
        anchorFixture.Ownership.TargetAnchorAddress = "B2";
        Assert.Throws<InvalidOperationException>(anchorFixture.Recover);
        Assert.DoesNotContain("create-target", anchorFixture.Controller.Events);

        var lineageFixture = new RecoveryFixture(RecoveryTargetMode.Absent);
        lineageFixture.Workbook.Model.ModelTables =
            new FakeCollection<FakeModelTable>(new[]
            {
                new FakeModelTable(
                    lineageFixture.Names.QueryName,
                    new object())
            });
        Assert.Throws<InvalidOperationException>(lineageFixture.Recover);
        Assert.DoesNotContain("create-target", lineageFixture.Controller.Events);

        var targetLineageFixture = new RecoveryFixture(
            RecoveryTargetMode.GeneratedPartial,
            wrongTargetConnection: true);
        Assert.Throws<InvalidOperationException>(targetLineageFixture.Recover);
        Assert.DoesNotContain("clear-target", targetLineageFixture.Controller.Events);
    }

    [Fact]
    public void VerifyActiveDataModelOwnership_ExactFinalTargetAndArtifacts_IsReadOnlyAndRepeatable()
    {
        var fixture = new RecoveryFixture(RecoveryTargetMode.Absent);
        fixture.Recover();
        fixture.MarkOwnershipActive();
        int eventCount = fixture.Controller.Events.Count;

        fixture.VerifyActive();
        fixture.VerifyActive();

        Assert.True(fixture.TargetPivotIsFinalAndExact());
        Assert.Equal(eventCount, fixture.Controller.Events.Count);
        Assert.Equal(1, fixture.Workbook.Queries.AddCalls);
        Assert.Equal(1, fixture.Workbook.Connections.AddCalls);
    }

    [Fact]
    public void VerifyActiveDataModelOwnership_QueryDriftOrTemporaryCollision_FailsClosed()
    {
        var drift = new RecoveryFixture(RecoveryTargetMode.Absent);
        drift.Recover();
        drift.MarkOwnershipActive();
        drift.Workbook.Queries.Query.Formula = "let Source = changed in Source";
        Assert.Throws<InvalidOperationException>(drift.VerifyActive);

        var collision = new RecoveryFixture(RecoveryTargetMode.Absent);
        collision.Recover();
        collision.MarkOwnershipActive();
        collision.AddUnownedTemporarySheetCollision();
        Assert.Throws<InvalidOperationException>(collision.VerifyActive);
        Assert.True(collision.TargetPivotIsFinalAndExact());
    }

    public enum RecoveryTargetMode
    {
        Absent,
        GeneratedPartial,
        Final
    }

    public enum RecoveryContamination
    {
        Values,
        Formulas,
        Note,
        Validation,
        Hyperlink,
        Merge,
        Table,
        NeighborPivot
    }

    public enum RecoveryFailurePoint
    {
        None,
        BeforeClear,
        AfterClear,
        BeforeCreate,
        AfterCreate,
        BeforeRestore,
        AfterRestore,
        BeforeRefresh,
        AfterRefresh,
        BeforePaste,
        AfterPaste,
        BeforeDimension,
        AfterDimension,
        BeforeRename,
        AfterRename,
        BeforeFormatDelete,
        AfterFormatDelete,
        BeforeStageDelete,
        AfterStageDelete
    }

    private sealed class RecoveryFixture
    {
        private readonly LateBoundPivotDataModelEnablementGateway gateway =
            new LateBoundPivotDataModelEnablementGateway();
        private readonly FakeRecoveryWorksheet targetSheet;
        private readonly FakeRecoveryWorksheet stageSheet;
        private readonly FakeRecoveryWorksheet formatSheet;
        private readonly FakeRecoveryPivotCache stageCache;
        private readonly PivotDataModelArtifacts artifacts;
        private readonly string stagingCheckpoint;

        public RecoveryFixture(
            RecoveryTargetMode targetMode,
            bool wrongTargetConnection = false)
        {
            Controller = new RecoveryCrashController();
            Workbook = new FakeRecoveryWorkbook();
            Names = LateBoundPivotDataModelEnablementGateway
                .CompileGeneratedNames("setup-1");
            PivotDataModelArtifactPlan plan = gateway.PlanOwnedModelArtifacts(
                "setup-1",
                new ClassicPivotSourceDescriptor(
                    "SalesTable",
                    PivotPlusWorkbookObjectKind.Table));
            Ownership = MetadataFor("setup-1", plan);
            artifacts = gateway.EnsureOwnedModelArtifacts(
                Workbook,
                plan,
                Ownership);

            targetSheet = new FakeRecoveryWorksheet(
                "Sheet1",
                Controller,
                Workbook.Worksheets,
                RecoverySheetPurpose.Target);
            stageSheet = new FakeRecoveryWorksheet(
                Names.StagingWorksheetName,
                Controller,
                Workbook.Worksheets,
                RecoverySheetPurpose.Stage);
            formatSheet = new FakeRecoveryWorksheet(
                Names.FormatBackupWorksheetName,
                Controller,
                Workbook.Worksheets,
                RecoverySheetPurpose.Format);
            Workbook.Worksheets.AddExisting(targetSheet);
            Workbook.Worksheets.AddExisting(stageSheet);
            Workbook.Worksheets.AddExisting(formatSheet);

            PivotTemporaryWorksheetArtifact stageReceipt =
                artifacts.TemporaryWorksheets.Single(item =>
                    item.Purpose == "staging");
            PivotTemporaryWorksheetArtifact formatReceipt =
                artifacts.TemporaryWorksheets.Single(item =>
                    item.Purpose == "format-backup");
            AddWorksheetMarker(stageSheet, stageReceipt);
            AddWorksheetMarker(formatSheet, formatReceipt);

            stageCache = new FakeRecoveryPivotCache(
                artifacts.NativeDataModelConnection,
                targetSheet,
                Controller);
            var stageRange = new FakeRecoveryRange(
                stageSheet,
                Controller,
                value: 42d,
                isTarget: false);
            var stagePivot = new FakeRecoveryPivot(
                Names.StagingPivotTableName,
                stageSheet,
                stageCache,
                stageRange,
                Controller,
                isTarget: false);
            stageSheet.PivotTables.Add(stagePivot);
            stageSheet.UsedRange = stageRange;
            LateBoundPivotState stageState = stagePivot.ReadState();
            string checkpoint = PivotPlusFingerprint.Create(
                "pivotplus.staging-state.v1",
                stageState.CanonicalValue());
            stagingCheckpoint = checkpoint;
            stageSheet.CustomProperties.Add(
                "PivotTablePlusStagingStateFingerprint",
                checkpoint);
            Ownership.RecoveryPhase = PivotPlusRecoveryPhase.StagingVerified;
            Ownership.StagingStateFingerprint = checkpoint;

            var formatRange = new FakeRecoveryRange(
                formatSheet,
                Controller,
                value: null,
                isTarget: false);
            formatSheet.UsedRange = formatRange;
            formatSheet.Range.Anchor = formatRange;

            if (targetMode != RecoveryTargetMode.Absent)
            {
                string name = targetMode == RecoveryTargetMode.Final
                    ? "Pivot1"
                    : artifacts.TemporaryPivotTable!.Name;
                object connection = wrongTargetConnection
                    ? new object()
                    : artifacts.NativeDataModelConnection;
                var cache = new FakeRecoveryPivotCache(
                    connection,
                    targetSheet,
                    Controller);
                targetSheet.Destination.Value2 = 42d;
                var pivot = new FakeRecoveryPivot(
                    name,
                    targetSheet,
                    cache,
                    targetSheet.Destination,
                    Controller,
                    isTarget: true);
                if (targetMode == RecoveryTargetMode.GeneratedPartial)
                {
                    pivot.RowGrand = false;
                }
                targetSheet.PivotTables.Add(pivot);
            }
        }

        public RecoveryCrashController Controller { get; }

        public FakeRecoveryWorkbook Workbook { get; }

        public FakeRecoveryWorksheets Worksheets => Workbook.Worksheets;

        public GeneratedNames Names { get; }

        public PivotPlusWorkbookMetadata Ownership { get; }

        public PivotPendingDataModelRecovery Recover()
        {
            return gateway.RecoverPending(
                Workbook,
                "setup-1",
                Ownership);
        }

        public bool TargetPivotIsPresent()
        {
            return targetSheet.PivotTables.Count != 0;
        }

        public bool TargetPivotIsFinalAndExact()
        {
            if (targetSheet.PivotTables.Count != 1) return false;
            FakeRecoveryPivot pivot = targetSheet.PivotTables.Item(1);
            return string.Equals(pivot.Name, "Pivot1", StringComparison.Ordinal) &&
                   ReferenceEquals(
                       pivot.PivotCache().WorkbookConnection,
                       artifacts.NativeDataModelConnection) &&
                   string.Equals(
                       PivotPlusFingerprint.Create(
                           "pivotplus.staging-state.v1",
                           pivot.ReadState().CanonicalValue()),
                       stagingCheckpoint,
                       StringComparison.Ordinal);
        }

        public void RemoveFormatSheet()
        {
            Workbook.Worksheets.Remove(formatSheet);
        }

        public void RemoveStageSheet()
        {
            Workbook.Worksheets.Remove(stageSheet);
        }

        public void MarkOwnershipActive()
        {
            Ownership.Artifacts = Ownership.Artifacts
                .Where(item =>
                    item.Kind != PivotPlusArtifactKind.TemporaryWorksheet &&
                    item.Kind != PivotPlusArtifactKind.TemporaryPivotTable)
                .ToList();
            Ownership.RecoveryPhase = PivotPlusRecoveryPhase.None;
            Ownership.TargetAnchorAddress = string.Empty;
            Ownership.StagingStateFingerprint = string.Empty;
        }

        public void VerifyActive()
        {
            gateway.VerifyActiveDataModelOwnership(
                Workbook,
                "setup-1",
                Ownership);
        }

        public void AddUnownedTemporarySheetCollision()
        {
            Workbook.Worksheets.AddExisting(new FakeRecoveryWorksheet(
                Names.StagingWorksheetName,
                Controller,
                Workbook.Worksheets,
                RecoverySheetPurpose.Stage));
        }

        public void Contaminate(RecoveryContamination contamination)
        {
            switch (contamination)
            {
                case RecoveryContamination.Values:
                    targetSheet.Destination.Value2 = "occupied";
                    break;
                case RecoveryContamination.Formulas:
                    targetSheet.Destination.Formula = "=1+1";
                    break;
                case RecoveryContamination.Note:
                    targetSheet.Comments = new FakeCollection<FakeCellComment>(
                        new[] { new FakeCellComment(1, 1) });
                    break;
                case RecoveryContamination.Validation:
                    targetSheet.Destination.HasValidation = true;
                    break;
                case RecoveryContamination.Hyperlink:
                    targetSheet.Destination.Hyperlinks = new FakeCountCollection(1);
                    break;
                case RecoveryContamination.Merge:
                    targetSheet.Destination.MergeCells = true;
                    break;
                case RecoveryContamination.Table:
                    targetSheet.ListObjects = new FakeCollection<object>(
                        new object[] { new FakeRecoveryListObject(targetSheet.Destination) });
                    break;
                case RecoveryContamination.NeighborPivot:
                    targetSheet.PivotTables.Add(new FakeRecoveryPivot(
                        "NeighborPivot",
                        targetSheet,
                        new FakeRecoveryPivotCache(
                            artifacts.NativeDataModelConnection,
                            targetSheet,
                            Controller),
                        targetSheet.Destination,
                        Controller,
                        isTarget: false));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(contamination));
            }
        }

        private static void AddWorksheetMarker(
            FakeRecoveryWorksheet worksheet,
            PivotTemporaryWorksheetArtifact receipt)
        {
            worksheet.CustomProperties.Add(
                "PivotTablePlusPurpose",
                receipt.Purpose);
            worksheet.CustomProperties.Add(
                "PivotTablePlusFingerprint",
                receipt.Fingerprint);
            worksheet.CustomProperties.Add(
                "PivotTablePlusTargetAnchor",
                receipt.TargetAnchorAddress);
        }
    }

    public enum RecoverySheetPurpose
    {
        Target,
        Stage,
        Format
    }

    public sealed class RecoveryCrashController
    {
        private bool triggered;
        private RecoveryFailurePoint failurePoint;

        public List<string> Events { get; } = new List<string>();

        public RecoveryFailurePoint FailurePoint
        {
            get => failurePoint;
            set
            {
                failurePoint = value;
                triggered = false;
            }
        }

        public void Before(RecoveryFailurePoint point)
        {
            ThrowOnce(point);
        }

        public void After(RecoveryFailurePoint point)
        {
            ThrowOnce(point);
        }

        private void ThrowOnce(RecoveryFailurePoint point)
        {
            if (triggered || failurePoint != point) return;
            triggered = true;
            throw new InvalidOperationException(
                "Injected recovery crash at " + point + ".");
        }
    }

    public sealed class FakeRecoveryWorkbook
    {
        public FakeRecoveryWorkbook()
        {
            Queries = new FakeModelQueries();
            Connections = new FakeModelConnections();
            Names = new FakeModelNames();
            Model = new FakeWorkbookModel(Connections.Connection);
            Worksheets = new FakeRecoveryWorksheets();
            CustomXMLParts = new FakeEmptyCustomXmlParts();
        }

        public FakeModelQueries Queries { get; }

        public FakeModelConnections Connections { get; }

        public FakeModelNames Names { get; }

        public FakeWorkbookModel Model { get; }

        public FakeRecoveryWorksheets Worksheets { get; }

        public FakeEmptyCustomXmlParts CustomXMLParts { get; }
    }

    public sealed class FakeRecoveryWorksheets
    {
        private readonly List<FakeRecoveryWorksheet> items =
            new List<FakeRecoveryWorksheet>();

        public int Count => items.Count;

        public FakeRecoveryWorksheet Item(int index)
        {
            return items[index - 1];
        }

        public void AddExisting(FakeRecoveryWorksheet worksheet)
        {
            items.Add(worksheet);
        }

        public void Remove(FakeRecoveryWorksheet worksheet)
        {
            items.Remove(worksheet);
        }

        public bool Contains(string name)
        {
            return items.Any(item => string.Equals(
                item.Name,
                name,
                StringComparison.OrdinalIgnoreCase));
        }
    }

    public sealed class FakeRecoveryWorksheet
    {
        private readonly RecoveryCrashController controller;
        private readonly FakeRecoveryWorksheets owner;
        private readonly RecoverySheetPurpose purpose;

        public FakeRecoveryWorksheet(
            string name,
            RecoveryCrashController controller,
            FakeRecoveryWorksheets owner,
            RecoverySheetPurpose purpose)
        {
            Name = name;
            this.controller = controller;
            this.owner = owner;
            this.purpose = purpose;
            Visible = purpose == RecoverySheetPurpose.Target ? -1 : 2;
            CustomProperties = new FakeWorksheetCustomProperties();
            PivotTables = new FakeRecoveryPivotCollection();
            ListObjects = new FakeCollection<object>(Array.Empty<object>());
            Shapes = new FakeCollection<object>(Array.Empty<object>());
            Comments = new FakeCollection<FakeCellComment>(
                Array.Empty<FakeCellComment>());
            CommentsThreaded = new FakeCollection<FakeCellComment>(
                Array.Empty<FakeCellComment>());
            var anchor = new FakeRecoveryRange(
                this,
                controller,
                value: null,
                isTarget: purpose == RecoverySheetPurpose.Target);
            Range = new FakeRecoveryRangeLookup(anchor);
            UsedRange = anchor;
        }

        public string Name { get; set; }

        public int Visible { get; set; }

        public FakeWorksheetCustomProperties CustomProperties { get; }

        public FakeRecoveryPivotCollection PivotTables { get; }

        public FakeCollection<object> ListObjects { get; set; }

        public FakeCollection<object> Shapes { get; set; }

        public FakeCollection<FakeCellComment> Comments { get; set; }

        public FakeCollection<FakeCellComment> CommentsThreaded { get; set; }

        public FakeRecoveryRangeLookup Range { get; }

        public FakeRecoveryRange UsedRange { get; set; }

        public FakeRecoveryRange Destination => Range.Anchor;

        public void Delete()
        {
            RecoveryFailurePoint before;
            RecoveryFailurePoint after;
            string eventName;
            if (purpose == RecoverySheetPurpose.Format)
            {
                before = RecoveryFailurePoint.BeforeFormatDelete;
                after = RecoveryFailurePoint.AfterFormatDelete;
                eventName = "delete-format";
            }
            else if (purpose == RecoverySheetPurpose.Stage)
            {
                before = RecoveryFailurePoint.BeforeStageDelete;
                after = RecoveryFailurePoint.AfterStageDelete;
                eventName = "delete-stage";
            }
            else
            {
                throw new InvalidOperationException(
                    "The target worksheet is not an owned recovery artifact.");
            }

            controller.Before(before);
            controller.Events.Add(eventName);
            owner.Remove(this);
            controller.After(after);
        }
    }

    public sealed class FakeRecoveryRangeLookup
    {
        public FakeRecoveryRangeLookup(FakeRecoveryRange anchor)
        {
            Anchor = anchor;
        }

        public FakeRecoveryRange Anchor { get; set; }

        public FakeRecoveryRange this[string address]
        {
            get
            {
                Assert.Equal("A1", address);
                return Anchor;
            }
        }
    }

    public sealed class FakeRecoveryResizeLookup
    {
        private readonly FakeRecoveryRange range;

        public FakeRecoveryResizeLookup(FakeRecoveryRange range)
        {
            this.range = range;
        }

        public FakeRecoveryRange this[int rows, int columns]
        {
            get
            {
                Assert.Equal(1, rows);
                Assert.Equal(1, columns);
                return range;
            }
        }
    }

    public sealed class FakeRecoveryRange
    {
        private readonly RecoveryCrashController controller;
        private readonly bool isTarget;

        public FakeRecoveryRange(
            FakeRecoveryWorksheet parent,
            RecoveryCrashController controller,
            object? value,
            bool isTarget)
        {
            Parent = parent;
            this.controller = controller;
            this.isTarget = isTarget;
            Value2 = value;
            Formula = string.Empty;
            Hyperlinks = new FakeCountCollection(0);
            Rows = new FakeRecoveryDimensionCollection(
                new FakeRecoveryDimension(controller, isTarget));
            Columns = new FakeRecoveryDimensionCollection(
                new FakeRecoveryDimension(controller, isTarget));
            Resize = new FakeRecoveryResizeLookup(this);
        }

        public FakeRecoveryWorksheet Parent { get; }

        public int Row => 1;

        public int Column => 1;

        public FakeRecoveryRange Cells => this;

        public long CountLarge => 1;

        public FakeRecoveryDimensionCollection Rows { get; }

        public FakeRecoveryDimensionCollection Columns { get; }

        public FakeRecoveryResizeLookup Resize { get; }

        public object? Value2 { get; set; }

        public object? Formula { get; set; }

        public object Hyperlinks { get; set; }

        public object MergeCells { get; set; } = false;

        public bool HasValidation { get; set; }

        public FakeRecoveryRange this[int row, int column]
        {
            get
            {
                Assert.Equal(1, row);
                Assert.Equal(1, column);
                return this;
            }
        }

        public string Address(bool rowAbsolute, bool columnAbsolute)
        {
            Assert.False(rowAbsolute);
            Assert.False(columnAbsolute);
            return "A1";
        }

        public object SpecialCells(int cellType)
        {
            if (HasValidation) return this;
            throw new COMException("No cells were found.", unchecked((int)0x800A03EC));
        }

        public void Copy()
        {
            controller.Events.Add(isTarget ? "copy-target" : "copy-format");
        }

        public void PasteSpecial(int pasteType)
        {
            if (!isTarget) return;
            if (pasteType == -4122)
            {
                controller.Before(RecoveryFailurePoint.BeforePaste);
                controller.Events.Add("paste-formats");
                controller.After(RecoveryFailurePoint.AfterPaste);
            }
            else
            {
                controller.Events.Add("paste-column-widths");
            }
        }

        public void Clear()
        {
            if (!isTarget)
            {
                Value2 = null;
                Formula = string.Empty;
                return;
            }

            controller.Before(RecoveryFailurePoint.BeforeClear);
            controller.Events.Add("clear-target");
            Parent.PivotTables.RemoveByRange(this);
            Value2 = null;
            Formula = string.Empty;
            controller.After(RecoveryFailurePoint.AfterClear);
        }
    }

    public sealed class FakeRecoveryDimensionCollection
    {
        private readonly FakeRecoveryDimension item;

        public FakeRecoveryDimensionCollection(FakeRecoveryDimension item)
        {
            this.item = item;
        }

        public int Count => 1;

        public FakeRecoveryDimension Item(int index)
        {
            Assert.Equal(1, index);
            return item;
        }
    }

    public sealed class FakeRecoveryDimension
    {
        private readonly RecoveryCrashController controller;
        private readonly bool isTarget;
        private double rowHeight = 15d;
        private double columnWidth = 8.43d;

        public FakeRecoveryDimension(
            RecoveryCrashController controller,
            bool isTarget)
        {
            this.controller = controller;
            this.isTarget = isTarget;
        }

        public double RowHeight
        {
            get => rowHeight;
            set
            {
                if (isTarget)
                {
                    controller.Before(RecoveryFailurePoint.BeforeDimension);
                    controller.Events.Add("restore-dimension");
                }

                rowHeight = value;
                if (isTarget)
                {
                    controller.After(RecoveryFailurePoint.AfterDimension);
                }
            }
        }

        public double ColumnWidth
        {
            get => columnWidth;
            set
            {
                if (isTarget)
                {
                    controller.Before(RecoveryFailurePoint.BeforeDimension);
                    controller.Events.Add("restore-dimension");
                }

                columnWidth = value;
                if (isTarget)
                {
                    controller.After(RecoveryFailurePoint.AfterDimension);
                }
            }
        }
    }

    public sealed class FakeRecoveryPivotCollection
    {
        private readonly List<FakeRecoveryPivot> items =
            new List<FakeRecoveryPivot>();

        public int Count => items.Count;

        public FakeRecoveryPivot Item(int index)
        {
            return items[index - 1];
        }

        public void Add(FakeRecoveryPivot pivot)
        {
            items.Add(pivot);
        }

        public void RemoveByRange(FakeRecoveryRange range)
        {
            items.RemoveAll(item => ReferenceEquals(item.TableRange2, range));
        }
    }

    public sealed class FakeRecoveryPivotCache
    {
        private readonly FakeRecoveryWorksheet targetWorksheet;
        private readonly RecoveryCrashController controller;

        public FakeRecoveryPivotCache(
            object workbookConnection,
            FakeRecoveryWorksheet targetWorksheet,
            RecoveryCrashController controller)
        {
            WorkbookConnection = workbookConnection;
            this.targetWorksheet = targetWorksheet;
            this.controller = controller;
        }

        public bool OLAP => true;

        public object WorkbookConnection { get; }

        public FakeRecoveryPivot CreatePivotTable(object destination, string name)
        {
            controller.Before(RecoveryFailurePoint.BeforeCreate);
            var range = (FakeRecoveryRange)destination;
            range.Value2 = 42d;
            range.Formula = string.Empty;
            var pivot = new FakeRecoveryPivot(
                name,
                targetWorksheet,
                this,
                range,
                controller,
                isTarget: true);
            targetWorksheet.PivotTables.Add(pivot);
            controller.Events.Add("create-target");
            controller.After(RecoveryFailurePoint.AfterCreate);
            return pivot;
        }
    }

    public sealed class FakeRecoveryPivot
    {
        private readonly FakeRecoveryPivotCache cache;
        private readonly RecoveryCrashController controller;
        private readonly bool isTarget;
        private string name;

        public FakeRecoveryPivot(
            string name,
            FakeRecoveryWorksheet parent,
            FakeRecoveryPivotCache cache,
            FakeRecoveryRange range,
            RecoveryCrashController controller,
            bool isTarget)
        {
            this.name = name;
            Parent = parent;
            this.cache = cache;
            TableRange2 = range;
            this.controller = controller;
            this.isTarget = isTarget;
            PivotFields = new FakeCollection<object>(Array.Empty<object>());
            CubeFields = new FakeCollection<object>(Array.Empty<object>());
            RowFields = new FakeCollection<object>(Array.Empty<object>());
            ColumnFields = new FakeCollection<object>(Array.Empty<object>());
            PageFields = new FakeCollection<object>(Array.Empty<object>());
            DataFields = new FakeCollection<object>(Array.Empty<object>());
        }

        public string Name
        {
            get => name;
            set
            {
                if (isTarget &&
                    !string.Equals(name, value, StringComparison.Ordinal))
                {
                    controller.Before(RecoveryFailurePoint.BeforeRename);
                    name = value;
                    controller.Events.Add("rename-target");
                    controller.After(RecoveryFailurePoint.AfterRename);
                    return;
                }

                name = value;
            }
        }

        public FakeRecoveryWorksheet Parent { get; }

        public bool ManualUpdate { get; set; }

        public bool RowGrand { get; set; } = true;

        public bool ColumnGrand { get; set; } = true;

        public bool DisplayFieldCaptions { get; set; } = true;

        public bool PreserveFormatting { get; set; } = true;

        public bool ShowTableStyleRowStripes { get; set; }

        public bool ShowTableStyleColumnStripes { get; set; }

        public bool DisplayNullString { get; set; } = true;

        public string NullString { get; set; } = string.Empty;

        public bool DisplayErrorString { get; set; }

        public string ErrorString { get; set; } = string.Empty;

        public bool ShowDrillIndicators { get; set; } = true;

        public bool EnableDrilldown { get; set; } = true;

        public bool VisualTotals { get; set; } = true;

        public bool SubtotalHiddenPageItems { get; set; } = true;

        public int PageFieldOrder { get; set; } = 1;

        public int PageFieldWrapCount { get; set; }

        public int CompactRowIndent { get; set; } = 1;

        public bool MergeLabels { get; set; }

        public int LayoutRowDefault { get; private set; }

        public string TableStyle2 { get; set; } = string.Empty;

        public FakeRecoveryRange TableRange2 { get; }

        public FakeCollection<object> PivotFields { get; }

        public FakeCollection<object> CubeFields { get; }

        public FakeCollection<object> RowFields { get; }

        public FakeCollection<object> ColumnFields { get; }

        public FakeCollection<object> PageFields { get; }

        public FakeCollection<object> DataFields { get; }

        public FakeRecoveryPivotCache PivotCache()
        {
            return cache;
        }

        public void ClearAllFilters()
        {
        }

        public void RowAxisLayout(int layout)
        {
            if (isTarget)
            {
                controller.Before(RecoveryFailurePoint.BeforeRestore);
            }

            LayoutRowDefault = layout;
            if (isTarget)
            {
                controller.Events.Add("restore-target");
                controller.After(RecoveryFailurePoint.AfterRestore);
            }
        }

        public void RefreshTable()
        {
            if (isTarget)
            {
                controller.Before(RecoveryFailurePoint.BeforeRefresh);
                controller.Events.Add("refresh-target");
                controller.After(RecoveryFailurePoint.AfterRefresh);
            }
        }

        internal LateBoundPivotState ReadState()
        {
            return new LateBoundPivotState(
                cache,
                Array.Empty<LateBoundFieldState>(),
                new LateBoundStyleState(
                    LayoutRowDefault,
                    RowGrand,
                    ColumnGrand,
                    DisplayFieldCaptions,
                    TableStyle2,
                    PreserveFormatting,
                    ShowTableStyleRowStripes,
                    ShowTableStyleColumnStripes,
                    DisplayNullString,
                    NullString,
                    DisplayErrorString,
                    ErrorString,
                    ShowDrillIndicators,
                    EnableDrilldown,
                    VisualTotals,
                    SubtotalHiddenPageItems,
                    PageFieldOrder,
                    PageFieldWrapCount,
                    CompactRowIndent,
                    MergeLabels),
                LateBoundPivotDataModelEnablementGateway.ReadResultSignature(
                    TableRange2),
                LateBoundDataAxisState.Hidden);
        }
    }

    public sealed class FakeRecoveryListObject
    {
        public FakeRecoveryListObject(FakeRecoveryRange range)
        {
            Range = range;
        }

        public FakeRecoveryRange Range { get; }
    }

    private static LateBoundFieldState ValueField(string sourceName)
    {
        return new LateBoundFieldState(
            PivotNativeFieldArea.Values,
            sourceName,
            "Sum of Cost",
            "Sum of Cost",
            1,
            -4157,
            "#,##0",
            false,
            Array.Empty<bool>(),
            Array.Empty<LateBoundMemberState>(),
            string.Empty,
            false);
    }

    private static IReadOnlyList<PivotTemporaryWorksheetArtifact>
        TemporaryReceipts(string setupId)
    {
        GeneratedNames names =
            LateBoundPivotDataModelEnablementGateway.CompileGeneratedNames(setupId);
        return new[]
        {
            TemporaryReceipt(names.StagingWorksheetName, "staging"),
            TemporaryReceipt(names.FormatBackupWorksheetName, "format-backup")
        };
    }

    private static PivotPlusWorkbookMetadata MetadataFor(
        string setupId,
        PivotDataModelArtifacts artifacts)
    {
        var owned = new List<PivotPlusOwnedArtifact>
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
            owned.Add(new PivotPlusOwnedArtifact
            {
                Kind = PivotPlusArtifactKind.WorkbookName,
                ArtifactId = artifacts.OwnedWorkbookName.Name,
                Fingerprint = artifacts.OwnedWorkbookName.ReferenceFingerprint
            });
        }

        owned.AddRange(artifacts.TemporaryWorksheets.Select(item =>
            new PivotPlusOwnedArtifact
            {
                Kind = item.Kind,
                ArtifactId = item.Name,
                Fingerprint = item.Fingerprint
            }));
        if (artifacts.TemporaryPivotTable != null)
        {
            owned.Add(new PivotPlusOwnedArtifact
            {
                Kind = artifacts.TemporaryPivotTable.Kind,
                ArtifactId = artifacts.TemporaryPivotTable.Name,
                Fingerprint = artifacts.TemporaryPivotTable.Fingerprint
            });
        }
        return new PivotPlusWorkbookMetadata
        {
            SetupId = setupId,
            TargetWorksheetName = "Sheet1",
            TargetPivotTableName = "Pivot1",
            RecoveryPhase = artifacts.TemporaryPivotTable == null
                ? PivotPlusRecoveryPhase.None
                : PivotPlusRecoveryPhase.Planned,
            TargetAnchorAddress = artifacts.TemporaryPivotTable?.TargetAnchorAddress ?? string.Empty,
            Artifacts = owned
        };
    }

    private static PivotPlusWorkbookMetadata MetadataFor(
        string setupId,
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
        if (plan.WorkbookName != null && plan.WorkbookNameFingerprint != null)
        {
            owned.Add(new PivotPlusOwnedArtifact
            {
                Kind = PivotPlusArtifactKind.WorkbookName,
                ArtifactId = plan.WorkbookName,
                Fingerprint = plan.WorkbookNameFingerprint
            });
        }

        owned.AddRange(plan.TemporaryWorksheets.Select(item =>
            new PivotPlusOwnedArtifact
            {
                Kind = item.Kind,
                ArtifactId = item.Name,
                Fingerprint = item.Fingerprint
            }));
        if (plan.TemporaryPivotTable != null)
        {
            owned.Add(new PivotPlusOwnedArtifact
            {
                Kind = plan.TemporaryPivotTable.Kind,
                ArtifactId = plan.TemporaryPivotTable.Name,
                Fingerprint = plan.TemporaryPivotTable.Fingerprint
            });
        }
        return new PivotPlusWorkbookMetadata
        {
            SetupId = setupId,
            TargetWorksheetName = "Sheet1",
            TargetPivotTableName = "Pivot1",
            RecoveryPhase = plan.TemporaryPivotTable == null
                ? PivotPlusRecoveryPhase.None
                : PivotPlusRecoveryPhase.Planned,
            TargetAnchorAddress = plan.TemporaryPivotTable?.TargetAnchorAddress ?? string.Empty,
            Artifacts = owned
        };
    }

    private static PivotTemporaryWorksheetArtifact TemporaryReceipt(
        string name,
        string purpose)
    {
        return new PivotTemporaryWorksheetArtifact(
            name,
            purpose,
            PivotPlusFingerprint.Create(
                "pivotplus.temporary-worksheet.v2",
                purpose + "\n" + name + "\nA1"),
            "A1");
    }

    private static PivotTemporaryPivotTableArtifact TemporaryPivotReceipt(
        string setupId,
        string modelTableName = "ModelTable",
        string connectionName = "connection",
        string? name = null)
    {
        string receiptName = name ??
            LateBoundPivotDataModelEnablementGateway
                .CompileGeneratedNames(setupId)
                .ReplacementPivotTableName;
        string canonical = setupId + "\n" + receiptName +
                           "\nSheet1\nPivot1\nA1\n" + connectionName +
                           "\n" + modelTableName;
        return new PivotTemporaryPivotTableArtifact(
            setupId,
            receiptName,
            PivotPlusFingerprint.Create(
                "pivotplus.temporary-pivot-table.v1",
                canonical),
            "Sheet1",
            "Pivot1",
            "A1",
            connectionName,
            modelTableName);
    }

    private static FakeFormatBackupWorksheet VerifiedStagingWorksheet(
        PivotTemporaryWorksheetArtifact receipt,
        string stagingStateFingerprint)
    {
        var worksheet = new FakeFormatBackupWorksheet
        {
            Name = receipt.Name,
            Visible = 2
        };
        worksheet.CustomProperties.Add("PivotTablePlusPurpose", receipt.Purpose);
        worksheet.CustomProperties.Add("PivotTablePlusFingerprint", receipt.Fingerprint);
        worksheet.CustomProperties.Add(
            "PivotTablePlusTargetAnchor",
            receipt.TargetAnchorAddress);
        worksheet.CustomProperties.Add(
            "PivotTablePlusStagingStateFingerprint",
            stagingStateFingerprint);
        return worksheet;
    }

    private static string DataModelFingerprint(LateBoundPivotState classicState)
    {
        LateBoundStyleState style = classicState.Style;
        var modelStyle = new LateBoundStyleState(
            style.RowAxisLayout,
            style.RowGrand,
            style.ColumnGrand,
            style.DisplayFieldCaptions,
            style.TableStyleName,
            style.PreserveFormatting,
            style.ShowRowStripes,
            style.ShowColumnStripes,
            style.DisplayNullString,
            style.NullString,
            style.DisplayErrorString,
            style.ErrorString,
            style.ShowDrillIndicators,
            enableDrilldown: true,
            style.VisualTotals,
            subtotalHiddenPageItems: true,
            style.PageFieldOrder,
            style.PageFieldWrapCount,
            style.CompactRowIndent,
            style.MergeLabels);
        var modelState = new LateBoundPivotState(
            classicState.OriginalCache,
            classicState.Fields,
            modelStyle,
            classicState.Result,
            classicState.DataAxis);
        return PivotPlusFingerprint.Create(
            "pivotplus.staging-state.v1",
            modelState.CanonicalValue());
    }

    private static LateBoundPivotState StateWithField(LateBoundFieldState field)
    {
        return new LateBoundPivotState(
            new object(),
            new[] { field },
            new LateBoundStyleState(0, true, true, true, string.Empty, true, false, false));
    }

    private static LateBoundPivotState StateWithResult(
        int rows,
        int columns,
        string fingerprint)
    {
        return new LateBoundPivotState(
            new object(),
            Array.Empty<LateBoundFieldState>(),
            new LateBoundStyleState(0, true, true, true, string.Empty, true, false, false),
            new PivotResultSignature(rows, columns, fingerprint));
    }

    private static LateBoundFieldState PageFieldState()
    {
        return new LateBoundFieldState(
            PivotNativeFieldArea.Filter,
            "Region",
            "Region",
            "Region",
            1,
            null,
            string.Empty,
            false,
            Array.Empty<bool>(),
            new[]
            {
                new LateBoundMemberState("North", "North", true, 1)
            },
            "North",
            false);
    }

    public sealed class FakePivot
    {
        private readonly FakePivotCache cache;

        public FakePivot(FakePivotCache cache)
        {
            this.cache = cache;
        }

        public FakePivotCache PivotCache()
        {
            return cache;
        }
    }

    public sealed class FakePivotCache
    {
        public FakePivotCache(bool olap, object sourceData, int sourceType = 1)
        {
            OLAP = olap;
            SourceData = sourceData;
            SourceType = sourceType;
        }

        public bool OLAP { get; }

        public int SourceType { get; }

        public object SourceData { get; }
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(
            IEnumerable<FakeWorksheet> worksheets,
            IEnumerable<FakeName> names)
        {
            List<FakeWorksheet> worksheetList = worksheets.ToList();
            foreach (FakeWorksheet worksheet in worksheetList)
            {
                worksheet.Parent = this;
            }

            Worksheets = new FakeCollection<FakeWorksheet>(worksheetList);
            Names = new FakeCollection<FakeName>(names);
        }

        public FakeCollection<FakeWorksheet> Worksheets { get; }

        public FakeCollection<FakeName> Names { get; }
    }

    public sealed class FakeRawWorkbook
    {
        public FakeRawWorkbook(
            string worksheetName,
            string expectedAddress,
            int rows,
            int columns)
        {
            Worksheet = new FakeRawWorksheet(
                worksheetName,
                expectedAddress,
                rows,
                columns);
            Worksheets = new FakeRawWorksheetLookup(Worksheet);
            Names = new FakeCollection<FakeName>(Array.Empty<FakeName>());
        }

        public FakeRawWorksheet Worksheet { get; }

        public FakeRawWorksheetLookup Worksheets { get; }

        public FakeCollection<FakeName> Names { get; }
    }

    public sealed class FakeRawWorksheetLookup
    {
        private readonly FakeRawWorksheet worksheet;

        public FakeRawWorksheetLookup(FakeRawWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public int Count => 1;

        public FakeRawWorksheet Item(int index)
        {
            Assert.Equal(1, index);
            return worksheet;
        }

        public FakeRawWorksheet Item(string name)
        {
            if (!string.Equals(name, worksheet.Name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("missing", nameof(name));
            }

            return worksheet;
        }
    }

    public sealed class FakeRawWorksheet
    {
        public FakeRawWorksheet(
            string name,
            string expectedAddress,
            int rows,
            int columns)
        {
            Name = name;
            ListObjects = new FakeCollection<FakeTable>(Array.Empty<FakeTable>());
            ResolvedRange = new FakeRawRange(this, rows, columns);
            Range = new FakeRawRangeLookup(expectedAddress, ResolvedRange);
        }

        public string Name { get; }

        public FakeCollection<FakeTable> ListObjects { get; }

        public FakeRawRange ResolvedRange { get; }

        public FakeRawRangeLookup Range { get; }
    }

    public sealed class FakeRawRangeLookup
    {
        private readonly string expectedAddress;
        private readonly FakeRawRange range;

        public FakeRawRangeLookup(string expectedAddress, FakeRawRange range)
        {
            this.expectedAddress = expectedAddress;
            this.range = range;
        }

        public FakeRawRange this[string address]
        {
            get
            {
                Assert.Equal(expectedAddress, address);
                return range;
            }
        }
    }

    public sealed class FakeRawRange
    {
        public FakeRawRange(FakeRawWorksheet parent, int rows, int columns)
        {
            Parent = parent;
            Areas = new FakeCollection<object>(new[] { new object() });
            Rows = new FakeCountCollection(rows);
            Columns = new FakeCountCollection(columns);
            Cells = new FakeRawCells((long)rows * columns);
        }

        public FakeRawWorksheet Parent { get; }

        public FakeCollection<object> Areas { get; }

        public FakeCountCollection Rows { get; }

        public FakeCountCollection Columns { get; }

        public FakeRawCells Cells { get; }

        public bool WasWritten { get; }
    }

    public sealed class FakeCountCollection
    {
        public FakeCountCollection(int count)
        {
            Count = count;
        }

        public int Count { get; }
    }

    public sealed class FakeRawCells
    {
        public FakeRawCells(long count)
        {
            CountLarge = count;
        }

        public long CountLarge { get; }
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(IEnumerable<FakeTable> tables, string name = "Data")
        {
            Name = name;
            ListObjects = new FakeCollection<FakeTable>(tables);
        }

        public string Name { get; }

        public FakeWorkbook? Parent { get; internal set; }

        public FakeCollection<FakeTable> ListObjects { get; }
    }

    public sealed class FakeTable
    {
        public FakeTable(string name)
        {
            Name = name;
            DisplayName = name;
        }

        public string Name { get; }

        public string DisplayName { get; }
    }

    public sealed class FakeName
    {
        public FakeName(string name, object refersToRange)
        {
            Name = name;
            RefersToRange = refersToRange;
        }

        public string Name { get; }

        public object RefersToRange { get; set; }
    }

    public sealed class FakeNamedRange
    {
        public FakeNamedRange(FakeWorksheet parent, int rows, int columns)
        {
            Parent = parent;
            Areas = new FakeCollection<object>(new[] { new object() });
            Rows = new FakeCountCollection(rows);
            Columns = new FakeCountCollection(columns);
            Cells = new FakeRawCells((long)rows * columns);
        }

        public FakeWorksheet Parent { get; }

        public FakeCollection<object> Areas { get; }

        public FakeCountCollection Rows { get; }

        public FakeCountCollection Columns { get; }

        public FakeRawCells Cells { get; }
    }

    public sealed class FakeCollection<T>
    {
        private readonly List<T> items;

        public FakeCollection(IEnumerable<T> items)
        {
            this.items = new List<T>(items);
        }

        public int Count => items.Count;

        public T Item(int index)
        {
            return items[index - 1];
        }

    }

    public sealed class FakeSlicerWorkbook
    {
        public FakeSlicerWorkbook(IEnumerable<FakeSlicerCache> caches)
        {
            SlicerCaches = new FakeCollection<FakeSlicerCache>(caches);
        }

        public FakeCollection<FakeSlicerCache> SlicerCaches { get; }
    }

    public sealed class FakeSlicerWorksheet
    {
        public FakeSlicerWorksheet(FakeSlicerWorkbook parent)
        {
            Parent = parent;
        }

        public FakeSlicerWorkbook Parent { get; }
    }

    public sealed class FakeSlicerPivot
    {
        public FakeSlicerPivot(string name, FakeSlicerWorksheet parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }

        public FakeSlicerWorksheet Parent { get; }
    }

    public sealed class FakeSlicerCache
    {
        public FakeSlicerCache(IEnumerable<FakeSlicerPivotReference> pivots)
        {
            PivotTables = new FakeCollection<FakeSlicerPivotReference>(pivots);
        }

        public FakeCollection<FakeSlicerPivotReference> PivotTables { get; }
    }

    public sealed class FakeSlicerPivotReference
    {
        public FakeSlicerPivotReference(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }

    public sealed class FakeUnreadableSlicerWorkbook
    {
        public object SlicerCaches =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeUnreadableSlicerWorksheet
    {
        public FakeUnreadableSlicerWorksheet(FakeUnreadableSlicerWorkbook parent)
        {
            Parent = parent;
        }

        public FakeUnreadableSlicerWorkbook Parent { get; }
    }

    public sealed class FakeUnreadableSlicerPivot
    {
        public FakeUnreadableSlicerPivot(FakeUnreadableSlicerWorksheet parent)
        {
            Parent = parent;
        }

        public string Name => "Pivot1";

        public FakeUnreadableSlicerWorksheet Parent { get; }
    }

    public sealed class FakeRollbackWorkbook
    {
        public FakeRollbackWorkbook(FakeRollbackWorksheet worksheet)
        {
            Worksheets = new FakeWorksheetLookup(worksheet);
        }

        public FakeWorksheetLookup Worksheets { get; }
    }

    public sealed class FakeWorksheetLookup
    {
        private readonly FakeRollbackWorksheet worksheet;

        public FakeWorksheetLookup(FakeRollbackWorksheet worksheet)
        {
            this.worksheet = worksheet;
            BackupWorksheet = new FakeFormatBackupWorksheet();
        }

        public int Count => 1;

        public FakeFormatBackupWorksheet BackupWorksheet { get; }

        public FakeRollbackWorksheet Item(int index)
        {
            Assert.Equal(1, index);
            return worksheet;
        }

        public FakeRollbackWorksheet Item(string name)
        {
            Assert.Equal("Sheet1", name);
            return worksheet;
        }

        public FakeFormatBackupWorksheet Add()
        {
            return BackupWorksheet;
        }
    }

    public sealed class FakeRollbackWorksheet
    {
        public FakeRollbackWorksheet(FakeRollbackPivot pivot)
        {
            PivotTables = new FakePivotLookup(pivot);
            Range = new FakeAddressRangeLookup();
        }

        public string Name => "Sheet1";

        public FakePivotLookup PivotTables { get; }

        public FakeAddressRangeLookup Range { get; }
    }

    public sealed class FakePivotLookup
    {
        private readonly FakeRollbackPivot pivot;

        public FakePivotLookup(FakeRollbackPivot pivot)
        {
            this.pivot = pivot;
        }

        public FakeRollbackPivot Item(string name)
        {
            if (string.Equals(name, "Pivot1", StringComparison.Ordinal))
            {
                return pivot;
            }

            throw new ArgumentException("missing", nameof(name));
        }
    }

    public sealed class FakeAddressRangeLookup
    {
        public object this[string address]
        {
            get
            {
                Assert.Equal("A1", address);
                return new object();
            }
        }
    }

    public sealed class FakeRollbackPivot
    {
        private readonly FakeClassicCache cache;

        public FakeRollbackPivot(FakeClassicCache cache)
        {
            this.cache = cache;
            TableRange2 = new FakeThrowingClearRange();
            PivotFields = new FakeCollection<object>(Array.Empty<object>());
            RowFields = new FakeCollection<object>(Array.Empty<object>());
            ColumnFields = new FakeCollection<object>(Array.Empty<object>());
            PageFields = new FakeCollection<object>(Array.Empty<object>());
            DataFields = new FakeCollection<object>(Array.Empty<object>());
        }

        public bool ManualUpdate { get; set; }

        public bool RowGrand { get; set; } = true;

        public bool ColumnGrand { get; set; } = true;

        public bool DisplayFieldCaptions { get; set; } = true;

        public bool PreserveFormatting { get; set; } = true;

        public bool ShowTableStyleRowStripes { get; set; }

        public bool ShowTableStyleColumnStripes { get; set; }

        public bool DisplayNullString { get; set; } = true;

        public string NullString { get; set; } = string.Empty;

        public bool DisplayErrorString { get; set; }

        public string ErrorString { get; set; } = string.Empty;

        public bool ShowDrillIndicators { get; set; } = true;

        public bool EnableDrilldown { get; set; } = true;

        public bool VisualTotals { get; set; } = true;

        public bool SubtotalHiddenPageItems { get; set; }

        public int PageFieldOrder { get; set; } = 1;

        public int PageFieldWrapCount { get; set; }

        public int CompactRowIndent { get; set; } = 1;

        public bool MergeLabels { get; set; }

        public int LayoutRowDefault { get; private set; }

        public string TableStyle2 { get; set; } = string.Empty;

        public FakeThrowingClearRange TableRange2 { get; }

        public FakeCollection<object> PivotFields { get; }

        public FakeCollection<object> RowFields { get; }

        public FakeCollection<object> ColumnFields { get; }

        public FakeCollection<object> PageFields { get; }

        public FakeCollection<object> DataFields { get; }

        public bool Refreshed { get; private set; }

        public FakeClassicCache PivotCache()
        {
            return cache;
        }

        public void ClearAllFilters()
        {
        }

        public void RowAxisLayout(int value)
        {
            LayoutRowDefault = value;
        }

        public void RefreshTable()
        {
            Refreshed = true;
        }
    }

    public sealed class FakeThrowingClearRange
    {
        public FakeDimensionCollection Rows { get; } =
            new FakeDimensionCollection(1);

        public FakeDimensionCollection Columns { get; } =
            new FakeDimensionCollection(1);

        public int CopyCalls { get; private set; }

        public List<int> PasteTypes { get; } = new List<int>();

        public double Value2 => 42d;

        public void Copy()
        {
            CopyCalls++;
        }

        public void PasteSpecial(int pasteType)
        {
            PasteTypes.Add(pasteType);
        }

        public void Clear()
        {
            throw new InvalidOperationException("clear failed before removal");
        }
    }

    public sealed class FakeClassicCache
    {
        public bool OLAP => false;

        public int CreateCalls { get; private set; }

        public object CreatePivotTable(object destination, string name)
        {
            CreateCalls++;
            return new object();
        }
    }

    public sealed class FakeReplacementCache
    {
        public object CreatePivotTable(object destination, string name)
        {
            throw new InvalidOperationException("replacement should not be reached");
        }
    }

    public sealed class FakeChartWorkbook
    {
        public FakeChartWorkbook(
            IEnumerable<FakeChart> charts,
            IEnumerable<FakeChartWorksheet> worksheets)
        {
            Charts = new FakeCollection<FakeChart>(charts);
            Worksheets = new FakeCollection<FakeChartWorksheet>(worksheets);
        }

        public FakeCollection<FakeChart> Charts { get; }

        public FakeCollection<FakeChartWorksheet> Worksheets { get; }
    }

    public sealed class FakeChartWorksheet
    {
        public FakeChartWorksheet(IEnumerable<FakeChartObject> charts)
        {
            ChartObjects = new FakeCollection<FakeChartObject>(charts);
        }

        public FakeChartWorkbook? Parent { get; set; }

        public FakeCollection<FakeChartObject> ChartObjects { get; }
    }

    public sealed class FakeChartTargetPivot
    {
        public FakeChartTargetPivot(string name, FakeChartWorksheet parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }

        public FakeChartWorksheet Parent { get; }
    }

    public sealed class FakeChartObject
    {
        public FakeChartObject(FakeChart chart)
        {
            Chart = chart;
        }

        public FakeChart Chart { get; }
    }

    public sealed class FakeChart
    {
        public FakeChart(FakeNamedPivot pivot)
        {
            PivotLayout = new FakePivotLayout(pivot);
        }

        public FakePivotLayout PivotLayout { get; }
    }

    public sealed class FakePivotLayout
    {
        public FakePivotLayout(FakeNamedPivot pivot)
        {
            PivotTable = pivot;
        }

        public FakeNamedPivot PivotTable { get; }
    }

    public sealed class FakeNamedPivot
    {
        public FakeNamedPivot(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }

    public sealed class FakeValueCalculation
    {
        public FakeValueCalculation(int calculation)
        {
            Calculation = calculation;
        }

        public int Calculation { get; }
    }

    public sealed class FakeCalculatedField
    {
        public bool IsCalculated => true;
    }

    public sealed class FakeOrdinaryExcel2021Field
    {
        public object IsCalculated =>
            throw new COMException("Member not found", unchecked((int)0x80020003));

        public FakeCollection<object> CalculatedItems() =>
            new FakeCollection<object>(Array.Empty<object>());

        public object? ParentField => null;

        public object? ChildField => null;

        public string Name => "Region";

        public string SourceName => "Region";
    }

    public sealed class FakePivotWithCalculatedDefinition
    {
        public FakeCollection<object> PivotFields { get; } =
            new FakeCollection<object>(Array.Empty<object>());

        public FakeCollection<object> CalculatedFields()
        {
            return new FakeCollection<object>(new[] { new object() });
        }
    }

    public sealed class FakeModelWorkbook
    {
        public FakeModelWorkbook()
        {
            Queries = new FakeModelQueries();
            Connections = new FakeModelConnections();
            Names = new FakeModelNames();
            Worksheets = new FakeCollection<object>(Array.Empty<object>());
            Model = new FakeWorkbookModel(Connections.Connection);
            CustomXMLParts = new FakeEmptyCustomXmlParts();
        }

        public FakeModelQueries Queries { get; }

        public FakeModelConnections Connections { get; }

        public FakeModelNames Names { get; }

        public FakeCollection<object> Worksheets { get; set; }

        public FakeWorkbookModel Model { get; }

        public FakeEmptyCustomXmlParts CustomXMLParts { get; }
    }

    public sealed class FakeModelQueries
    {
        private bool threwAfterCommit;

        public int Count => string.IsNullOrEmpty(AddedName) ? 0 : 1;

        public int AddCalls { get; private set; }

        public string AddedName { get; private set; } = string.Empty;

        public string AddedFormula { get; private set; } = string.Empty;

        public FakeExistingQuery Query { get; private set; } =
            new FakeExistingQuery(string.Empty);

        public bool ThrowAfterCommitOnce { get; set; }

        public object Item(int index)
        {
            Assert.Equal(1, index);
            return Query;
        }

        public FakeExistingQuery Add(string name, string formula)
        {
            AddedName = name;
            AddedFormula = formula;
            AddCalls++;
            Query = new FakeExistingQuery(formula, name);
            if (ThrowAfterCommitOnce && !threwAfterCommit)
            {
                threwAfterCommit = true;
                throw new InvalidOperationException("query inserted before COM failure");
            }
            return Query;
        }
    }

    public sealed class FakeModelNames
    {
        private bool threwAfterCommit;

        public int Count => string.IsNullOrEmpty(AddedName) ? 0 : 1;

        public int AddCalls { get; private set; }

        public string AddedName { get; private set; } = string.Empty;

        public string AddedReference { get; private set; } = string.Empty;

        public bool AddedVisible { get; private set; }

        public FakeOwnedName Name { get; private set; } =
            new FakeOwnedName(string.Empty, string.Empty);

        public bool ThrowAfterCommitOnce { get; set; }

        public object Item(int index)
        {
            Assert.Equal(1, index);
            return Name;
        }

        public FakeOwnedName Add(string name, string refersTo, bool visible)
        {
            AddedName = name;
            AddedReference = refersTo;
            AddedVisible = visible;
            AddCalls++;
            Name = new FakeOwnedName(name, refersTo);
            if (ThrowAfterCommitOnce && !threwAfterCommit)
            {
                threwAfterCommit = true;
                throw new InvalidOperationException("name inserted before COM failure");
            }
            return Name;
        }
    }

    public sealed class FakeOwnedName
    {
        public FakeOwnedName(string name, string refersTo)
        {
            Name = name;
            RefersTo = refersTo;
        }

        public string Name { get; }

        public string RefersTo { get; set; }

        public bool Deleted { get; private set; }

        public void Delete()
        {
            Deleted = true;
        }
    }

    public sealed class FakeModelConnections
    {
        private bool threwAfterCommit;

        public FakeModelConnections()
        {
            Connection = new FakeCreatedConnection();
        }

        public string AddedName { get; private set; } = string.Empty;

        public string ConnectionString { get; private set; } = string.Empty;

        public string CommandText { get; private set; } = string.Empty;

        public int CommandType { get; private set; }

        public bool CreateModelConnection { get; private set; }

        public bool ImportRelationships { get; private set; }

        public FakeCreatedConnection Connection { get; }

        public int Count => string.IsNullOrEmpty(AddedName) ? 0 : 1;

        public int AddCalls { get; private set; }

        public bool ThrowAfterCommitOnce { get; set; }

        public object Item(int index)
        {
            Assert.Equal(1, index);
            return Connection;
        }

        public FakeCreatedConnection Add2(
            string name,
            string description,
            string connectionString,
            string commandText,
            int commandType,
            bool createModelConnection,
            bool importRelationships)
        {
            AddedName = name;
            ConnectionString = connectionString;
            CommandText = commandText;
            CommandType = commandType;
            CreateModelConnection = createModelConnection;
            ImportRelationships = importRelationships;
            AddCalls++;
            Connection.Name = name;
            Connection.InModel = createModelConnection;
            Connection.OLEDBConnection.Connection = connectionString;
            Connection.OLEDBConnection.CommandText = commandText;
            if (ThrowAfterCommitOnce && !threwAfterCommit)
            {
                threwAfterCommit = true;
                throw new InvalidOperationException("connection inserted before COM failure");
            }
            return Connection;
        }
    }

    public sealed class FakeCreatedConnection
    {
        public FakeCreatedConnection()
        {
            OLEDBConnection = new FakeCreatedOleDbConnection();
        }

        public FakeCreatedOleDbConnection OLEDBConnection { get; }

        public string Name { get; set; } = string.Empty;

        public bool InModel { get; set; }

        public bool Refreshed { get; private set; }

        public void Refresh()
        {
            Refreshed = true;
        }

        public void Delete()
        {
        }
    }

    public sealed class FakeCreatedOleDbConnection
    {
        public bool BackgroundQuery { get; set; } = true;

        public bool Refreshing => false;

        public string Connection { get; set; } = string.Empty;

        public string CommandText { get; set; } = string.Empty;
    }

    public sealed class FakeWorkbookModel
    {
        public FakeWorkbookModel(object sourceConnection)
        {
            DataModelConnection = new FakeDataModelConnection();
            ModelTables = new FakeCollection<FakeModelTable>(
                new[]
                {
                    new FakeModelTable(
                        LateBoundPivotDataModelEnablementGateway
                            .CompileGeneratedNames("setup-1")
                            .QueryName,
                        sourceConnection)
                });
        }

        public FakeDataModelConnection DataModelConnection { get; }

        public FakeCollection<FakeModelTable> ModelTables { get; set; }
    }

    public sealed class FakeDataModelConnection
    {
        public int Type => 7;
    }

    public sealed class FakeFinalizationPivot
    {
        private readonly FakeFinalizationCache cache;

        public FakeFinalizationPivot(object connection)
        {
            cache = new FakeFinalizationCache(connection);
        }

        public FakeFinalizationCache PivotCache()
        {
            return cache;
        }
    }

    public sealed class FakeFinalizationCache
    {
        public FakeFinalizationCache(object connection)
        {
            WorkbookConnection = connection;
        }

        public bool OLAP => true;

        public object WorkbookConnection { get; }
    }

    public sealed class FakeModelTable
    {
        public FakeModelTable(string name, object sourceWorkbookConnection)
        {
            Name = name;
            SourceWorkbookConnection = sourceWorkbookConnection;
        }

        public string Name { get; }

        public object SourceWorkbookConnection { get; }
    }

    public sealed class FakeFailingRawModelWorkbook
    {
        public FakeFailingRawModelWorkbook()
        {
            Queries = new FakeModelQueries();
            Names = new FakeModelNames();
            Connections = new FakeFailingConnections();
        }

        public FakeModelQueries Queries { get; }

        public FakeModelNames Names { get; }

        public FakeFailingConnections Connections { get; }
    }

    public sealed class FakeFailingConnections
    {
        public int Count => 0;

        public object Item(int index)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        public object Add2(
            string name,
            string description,
            string connectionString,
            string commandText,
            int commandType,
            bool createModelConnection,
            bool importRelationships)
        {
            throw new InvalidOperationException("connection creation failed");
        }
    }

    public sealed class FakeStagingWorkbook
    {
        public FakeStagingWorkbook(FakePivotCaches? pivotCaches = null)
        {
            Worksheets = new FakeStagingWorksheets();
            PivotCachesValue = pivotCaches ?? new FakePivotCaches();
        }

        public FakeStagingWorksheets Worksheets { get; }

        public FakePivotCaches PivotCachesValue { get; }

        public FakePivotCaches PivotCaches()
        {
            return PivotCachesValue;
        }
    }

    public sealed class FakeStagingWorksheets
    {
        public FakeStagingWorksheets()
        {
            AddedWorksheet = new FakeStagingWorksheet();
        }

        public FakeStagingWorksheet AddedWorksheet { get; }

        public int Count => 0;

        public object Item(int index)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        public FakeStagingWorksheet Add()
        {
            return AddedWorksheet;
        }
    }

    public sealed class FakeStagingWorksheet
    {
        public FakeStagingWorksheet()
        {
            Range = new FakeStagingRangeLookup();
            CustomProperties = new FakeWorksheetCustomProperties();
        }

        public string Name { get; set; } = string.Empty;

        public int Visible { get; set; }

        public FakeStagingRangeLookup Range { get; }

        public FakeWorksheetCustomProperties CustomProperties { get; }

        public bool Deleted { get; private set; }

        public void Delete()
        {
            Deleted = true;
        }
    }

    public sealed class FakeTemporaryPayloadWorksheet
    {
        public FakeTemporaryPayloadWorksheet(object? value)
        {
            UsedRange = new FakeTemporaryPayloadRange(value);
            PivotTables = new FakeCollection<object>(Array.Empty<object>());
            ListObjects = new FakeCollection<object>(Array.Empty<object>());
            Shapes = new FakeCollection<object>(Array.Empty<object>());
        }

        public int Visible => 2;

        public FakeTemporaryPayloadRange UsedRange { get; }

        public FakeCollection<object> PivotTables { get; }

        public FakeCollection<object> ListObjects { get; }

        public FakeCollection<object> Shapes { get; }
    }

    public sealed class FakeTemporaryPayloadRange
    {
        public FakeTemporaryPayloadRange(object? value)
        {
            Value2 = value;
            Formula = string.Empty;
            Cells = new FakeRawCells(1);
        }

        public object? Value2 { get; }

        public string Formula { get; }

        public FakeRawCells Cells { get; }
    }

    public sealed class FakeStagingRangeLookup
    {
        public object Destination { get; } = new object();

        public object this[string address]
        {
            get
            {
                Assert.Equal("A1", address);
                return Destination;
            }
        }
    }

    public sealed class FakePivotCaches
    {
        private readonly List<FakeStagingCache> caches;

        public FakePivotCaches(IEnumerable<FakeStagingCache>? existing = null)
        {
            caches = new List<FakeStagingCache>(
                existing ?? Array.Empty<FakeStagingCache>());
            Cache = caches.FirstOrDefault() ?? new FakeStagingCache();
        }

        public int SourceType { get; private set; }

        public object? SourceData { get; private set; }

        public int Version { get; private set; }

        public FakeStagingCache Cache { get; private set; }

        public int Count => caches.Count;

        public int CreateCalls { get; private set; }

        public FakeStagingCache Item(int index)
        {
            return caches[index - 1];
        }

        public FakeStagingCache Create(int sourceType, object sourceData, int version)
        {
            SourceType = sourceType;
            SourceData = sourceData;
            Version = version;
            CreateCalls++;
            Cache = new FakeStagingCache
            {
                WorkbookConnection = sourceData
            };
            caches.Add(Cache);
            return Cache;
        }
    }

    public sealed class FakeStagingCache
    {
        public object? Destination { get; private set; }

        public string PivotName { get; private set; } = string.Empty;

        public bool OLAP => true;

        public object? WorkbookConnection { get; set; }

        public object CreatePivotTable(object destination, string pivotName)
        {
            Destination = destination;
            PivotName = pivotName;
            return new object();
        }
    }

    public sealed class FakeCleanupWorkbook
    {
        public FakeCleanupWorkbook(
            FakeExactConnection connection,
            FakeExistingQuery query,
            FakeOwnedName? name = null)
        {
            Connections = new FakeExistingConnectionLookup(connection);
            Queries = new FakeExistingQueryLookup(query);
            Names = new FakeExistingNameLookup(name);
        }

        public FakeExistingConnectionLookup Connections { get; }

        public FakeExistingQueryLookup Queries { get; }

        public FakeExistingNameLookup Names { get; }
    }

    public sealed class FakeExistingConnectionLookup
    {
        private readonly FakeExactConnection connection;

        public FakeExistingConnectionLookup(FakeExactConnection connection)
        {
            this.connection = connection;
        }

        public int Count => 1;

        public FakeExactConnection Item(int index)
        {
            Assert.Equal(1, index);
            return connection;
        }
    }

    public sealed class FakeExistingQueryLookup
    {
        private readonly FakeExistingQuery query;

        public FakeExistingQueryLookup(FakeExistingQuery query)
        {
            this.query = query;
        }

        public int Count => 1;

        public FakeExistingQuery Item(int index)
        {
            Assert.Equal(1, index);
            return query;
        }
    }

    public sealed class FakeExistingNameLookup
    {
        private readonly FakeOwnedName? name;

        public FakeExistingNameLookup(FakeOwnedName? name)
        {
            this.name = name;
        }

        public int Count => name == null ? 0 : 1;

        public FakeOwnedName Item(int index)
        {
            Assert.Equal(1, index);
            if (name == null)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return name;
        }
    }

    public sealed class FakeExactConnection
    {
        public FakeExactConnection(string name, string connection, string commandText)
        {
            Name = name;
            OLEDBConnection = new FakeExactOleDb(connection, commandText);
        }

        public string Name { get; }

        public FakeExactOleDb OLEDBConnection { get; }

        public bool Deleted { get; private set; }

        public void Delete()
        {
            Deleted = true;
        }
    }

    public sealed class FakeExactOleDb
    {
        public FakeExactOleDb(string connection, string commandText)
        {
            Connection = connection;
            CommandText = commandText;
        }

        public string Connection { get; }

        public string CommandText { get; }
    }

    public sealed class FakeExistingQuery
    {
        public FakeExistingQuery(string formula, string name = "PivotPlus_Test_Source")
        {
            Formula = formula;
            Name = name;
        }

        public string Name { get; }

        public string Formula { get; set; }

        public bool Deleted { get; private set; }

        public void Delete()
        {
            Deleted = true;
        }
    }

    public sealed class FakeTransactionalWorkbook
    {
        public FakeTransactionalWorkbook(FakeTransactionalWorksheet worksheet)
        {
            Worksheets = new FakeTransactionalWorksheetLookup(worksheet);
        }

        public FakeTransactionalWorksheetLookup Worksheets { get; }
    }

    public sealed class FakeTransactionalWorksheetLookup
    {
        private readonly FakeTransactionalWorksheet worksheet;

        public FakeTransactionalWorksheetLookup(FakeTransactionalWorksheet worksheet)
        {
            this.worksheet = worksheet;
            BackupWorksheet = new FakeFormatBackupWorksheet();
        }

        public int Count => 1;

        public FakeFormatBackupWorksheet BackupWorksheet { get; }

        public FakeTransactionalWorksheet Item(int index)
        {
            Assert.Equal(1, index);
            return worksheet;
        }

        public FakeTransactionalWorksheet Item(string name)
        {
            Assert.Equal("Sheet1", name);
            return worksheet;
        }

        public FakeFormatBackupWorksheet Add()
        {
            return BackupWorksheet;
        }
    }

    public sealed class FakeTransactionalWorksheet
    {
        public FakeTransactionalWorksheet(FakeTransactionalPivotLookup lookup)
        {
            PivotTables = lookup;
            Range = new FakeAddressRangeLookup();
            lookup.Worksheet = this;
        }

        public string Name => "Sheet1";

        public FakeTransactionalPivotLookup PivotTables { get; }

        public FakeAddressRangeLookup Range { get; }
    }

    public sealed class FakeTransactionalPivotLookup
    {
        public FakeTransactionalPivot? Current { get; set; }

        public FakeTransactionalWorksheet? Worksheet { get; set; }

        public int Count => Current == null ? 0 : 1;

        public FakeTransactionalPivot Item(int index)
        {
            if (index == 1 && Current != null) return Current;
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        public FakeTransactionalPivot Item(string name)
        {
            if (Current != null &&
                string.Equals(Current.Name, name, StringComparison.Ordinal))
            {
                return Current;
            }

            throw new ArgumentException("missing", nameof(name));
        }
    }

    public sealed class FakeTransactionalCache
    {
        private readonly FakeTransactionalPivotLookup lookup;
        private readonly List<string> events;
        private readonly string kind;

        public FakeTransactionalCache(
            bool olap,
            FakeTransactionalPivotLookup lookup,
            List<string> events,
            string kind,
            object? workbookConnection = null)
        {
            OLAP = olap;
            this.lookup = lookup;
            this.events = events;
            this.kind = kind;
            WorkbookConnection = workbookConnection ?? new object();
        }

        public bool OLAP { get; }

        public object WorkbookConnection { get; }

        public bool FailCreatedRestore { get; set; }

        public bool FailPromotion { get; set; }

        public FakeTransactionalPivot CreatePivotTable(object destination, string name)
        {
            if (string.Equals(kind, "classic", StringComparison.Ordinal))
            {
                Assert.Equal("Pivot1", name);
            }
            else
            {
                Assert.Equal("PP_Target_test", name);
            }
            events.Add("create-" + kind);
            var created = new FakeTransactionalPivot(
                this,
                lookup,
                events,
                kind,
                name);
            lookup.Current = created;
            return created;
        }
    }

    public sealed class FakeTransactionalPivot
    {
        private readonly FakeTransactionalCache cache;
        private readonly List<string> events;
        private readonly string kind;
        private string name;

        public FakeTransactionalPivot(
            FakeTransactionalCache cache,
            FakeTransactionalPivotLookup lookup,
            List<string> events,
            string kind,
            string name = "Pivot1")
        {
            this.cache = cache;
            this.lookup = lookup;
            this.events = events;
            this.kind = kind;
            this.name = name;
            SubtotalHiddenPageItems = cache.OLAP;
            TableRange2 = new FakeTransactionalRange(lookup, events, kind);
            PivotFields = new FakeCollection<object>(Array.Empty<object>());
            CubeFields = new FakeCollection<object>(Array.Empty<object>());
            RowFields = new FakeCollection<object>(Array.Empty<object>());
            ColumnFields = new FakeCollection<object>(Array.Empty<object>());
            PageFields = new FakeCollection<object>(Array.Empty<object>());
            DataFields = new FakeCollection<object>(Array.Empty<object>());
        }

        private readonly FakeTransactionalPivotLookup lookup;

        public string Name
        {
            get => name;
            set
            {
                if (cache.FailPromotion &&
                    string.Equals(kind, "model", StringComparison.Ordinal) &&
                    string.Equals(value, "Pivot1", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("promotion failed before rename");
                }

                name = value;
            }
        }

        public FakeTransactionalWorksheet Parent => lookup.Worksheet ??
            throw new InvalidOperationException("worksheet not bound");

        public bool ManualUpdate { get; set; }

        public bool RowGrand { get; set; } = true;

        public bool ColumnGrand { get; set; } = true;

        public bool DisplayFieldCaptions { get; set; } = true;

        public bool PreserveFormatting { get; set; } = true;

        public bool ShowTableStyleRowStripes { get; set; }

        public bool ShowTableStyleColumnStripes { get; set; }

        public bool DisplayNullString { get; set; } = true;

        public string NullString { get; set; } = string.Empty;

        public bool DisplayErrorString { get; set; }

        public string ErrorString { get; set; } = string.Empty;

        public bool ShowDrillIndicators { get; set; } = true;

        public bool EnableDrilldown { get; set; } = true;

        public bool VisualTotals { get; set; } = true;

        public bool SubtotalHiddenPageItems { get; set; }

        public int PageFieldOrder { get; set; } = 1;

        public int PageFieldWrapCount { get; set; }

        public int CompactRowIndent { get; set; } = 1;

        public bool MergeLabels { get; set; }

        public int LayoutRowDefault { get; private set; }

        public string TableStyle2 { get; set; } = string.Empty;

        public FakeTransactionalRange TableRange2 { get; }

        public FakeCollection<object> PivotFields { get; }

        public FakeCollection<object> CubeFields { get; }

        public FakeCollection<object> RowFields { get; }

        public FakeCollection<object> ColumnFields { get; }

        public FakeCollection<object> PageFields { get; }

        public FakeCollection<object> DataFields { get; }

        public FakeTransactionalCache PivotCache()
        {
            return cache;
        }

        public void ClearAllFilters()
        {
        }

        public void RowAxisLayout(int value)
        {
            if (string.Equals(kind, "classic", StringComparison.Ordinal) &&
                cache.FailCreatedRestore)
            {
                throw new InvalidOperationException("classic restore failed");
            }

            LayoutRowDefault = value;
        }

        public void RefreshTable()
        {
            events.Add("refresh-" + kind);
        }
    }

    public sealed class FakeTransactionalRange
    {
        private readonly FakeTransactionalPivotLookup lookup;
        private readonly List<string> events;
        private readonly string kind;

        public FakeTransactionalRange(
            FakeTransactionalPivotLookup lookup,
            List<string> events,
            string kind)
        {
            this.lookup = lookup;
            this.events = events;
            this.kind = kind;
        }

        public FakeDimensionCollection Rows { get; } =
            new FakeDimensionCollection(1);

        public FakeDimensionCollection Columns { get; } =
            new FakeDimensionCollection(1);

        public int CopyCalls { get; private set; }

        public List<int> PasteTypes { get; } = new List<int>();

        public double Value2 => 42d;

        public int Row => 1;

        public int Column => 1;

        public FakeTransactionalRange Cells => this;

        public FakeTransactionalRange this[int row, int column]
        {
            get
            {
                Assert.Equal(1, row);
                Assert.Equal(1, column);
                return this;
            }
        }

        public string Address(bool rowAbsolute, bool columnAbsolute)
        {
            Assert.False(rowAbsolute);
            Assert.False(columnAbsolute);
            return "A1";
        }

        public void Copy()
        {
            CopyCalls++;
        }

        public void PasteSpecial(int pasteType)
        {
            PasteTypes.Add(pasteType);
        }

        public void Clear()
        {
            events.Add("clear-" + kind);
            lookup.Current = null;
        }
    }

    public sealed class FakeBoundWorkbook
    {
        public FakeBoundWorkbook(string workbookId)
        {
            string xml =
                "<workbookIdentity xmlns=\"urn:excel-report-builder:workbook-identity\" " +
                "schemaVersion=\"1.0\" id=\"" + workbookId + "\" />";
            CustomXMLParts = new FakeCustomXmlParts(xml);
        }

        public FakeCustomXmlParts CustomXMLParts { get; }
    }

    public sealed class FakeCustomXmlParts
    {
        private readonly FakeCustomXmlPartCollection parts;

        public FakeCustomXmlParts(string xml)
        {
            parts = new FakeCustomXmlPartCollection(xml);
        }

        public FakeCustomXmlPartCollection SelectByNamespace(string namespaceUri)
        {
            return parts;
        }
    }

    public sealed class FakeCustomXmlPartCollection
    {
        private readonly FakeCustomXmlPart part;

        public FakeCustomXmlPartCollection(string xml)
        {
            part = new FakeCustomXmlPart(xml);
        }

        public int Count => 1;

        public FakeCustomXmlPart Item(int index)
        {
            Assert.Equal(1, index);
            return part;
        }
    }

    public sealed class FakeCustomXmlPart
    {
        public FakeCustomXmlPart(string xml)
        {
            XML = xml;
        }

        public string XML { get; }
    }

    public sealed class FakeBoundWorksheet
    {
        public FakeBoundWorksheet(string name, FakeBoundWorkbook parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }

        public FakeBoundWorkbook Parent { get; }
    }

    public sealed class FakeBoundPivot
    {
        public FakeBoundPivot(string name, FakeBoundWorksheet parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }

        public FakeBoundWorksheet Parent { get; }
    }

    public sealed class FakeUnpersistedBoundWorkbook
    {
        public FakeUnpersistedBoundWorkbook()
        {
            CustomXMLParts = new FakeEmptyCustomXmlParts();
        }

        public FakeEmptyCustomXmlParts CustomXMLParts { get; }
    }

    public sealed class FakeEmptyCustomXmlParts
    {
        public int AddCalls { get; private set; }

        public FakeEmptyCustomXmlPartCollection SelectByNamespace(string namespaceUri)
        {
            return new FakeEmptyCustomXmlPartCollection();
        }

        public object Add(string xml)
        {
            AddCalls++;
            return new FakeCustomXmlPart(xml);
        }
    }

    public sealed class FakeEmptyCustomXmlPartCollection
    {
        public int Count => 0;

        public object Item(int index)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }
    }

    public sealed class FakeUnpersistedBoundWorksheet
    {
        public FakeUnpersistedBoundWorksheet(FakeUnpersistedBoundWorkbook parent)
        {
            Parent = parent;
        }

        public string Name => "Sheet1";

        public FakeUnpersistedBoundWorkbook Parent { get; }
    }

    public sealed class FakeUnpersistedBoundPivot
    {
        public FakeUnpersistedBoundPivot(FakeUnpersistedBoundWorksheet parent)
        {
            Parent = parent;
        }

        public string Name => "Pivot1";

        public FakeUnpersistedBoundWorksheet Parent { get; }
    }

    public sealed class FakeUnreadableAxisPivot
    {
        public FakeClassicCache PivotCache()
        {
            return new FakeClassicCache();
        }
    }

    public sealed class FakeValueFunction
    {
        public FakeValueFunction(int function)
        {
            Function = function;
        }

        public int Function { get; }
    }

    public sealed class FakePageField
    {
        public FakePageField()
        {
            Item = new FakePageItem();
            PivotItems = new FakeCollection<FakePageItem>(new[] { Item });
            CubeField = new FakePageCubeField();
        }

        public string Caption { get; set; } = string.Empty;

        public bool RepeatLabels { get; set; }

        public bool LayoutBlankLine { get; set; }

        public bool LayoutPageBreak { get; set; }

        public bool EnableMultiplePageItems { get; set; }

        public bool DatabaseSort { get; set; } = true;

        public string[] VisibleItemsList { get; set; } = Array.Empty<string>();

        public FakePageCubeField CubeField { get; }

        public object? CurrentPage { get; set; }

        public string CurrentPageName { get; set; } = string.Empty;

        public FakePageItem Item { get; }

        public FakeCollection<FakePageItem> PivotItems { get; }

    }

    public sealed class FakePageCubeField
    {
        public bool AllItemsVisible { get; private set; }

        public void ClearManualFilter()
        {
            AllItemsVisible = true;
        }
    }

    public sealed class FakeOlapFilterField
    {
        private string[] visibleItemsList = Array.Empty<string>();

        public FakeOlapFilterField()
        {
            North = new FakeOlapFilterItem("North", "[Model].[Region].&[North]");
            South = new FakeOlapFilterItem("South", "[Model].[Region].&[South]");
            PivotItems = new FakeCollection<FakeOlapFilterItem>(new[] { North, South });
            CubeField = new FakePageCubeField();
        }

        public string Caption { get; set; } = string.Empty;

        public bool RepeatLabels { get; set; }

        public bool LayoutBlankLine { get; set; }

        public bool LayoutPageBreak { get; set; }

        public bool EnableMultiplePageItems { get; set; }

        public bool DatabaseSort { get; set; } = true;

        public string[] VisibleItemsList
        {
            get => visibleItemsList;
            set => visibleItemsList = value;
        }

        public FakePageCubeField CubeField { get; }

        public FakeOlapFilterItem North { get; }

        public FakeOlapFilterItem South { get; }

        public FakeCollection<FakeOlapFilterItem> PivotItems { get; }
    }

    public sealed class FakeOlapFilterItem
    {
        private bool visible = true;

        public FakeOlapFilterItem(string caption, string uniqueName)
        {
            Name = caption;
            Caption = caption;
            SourceName = uniqueName;
            SourceNameStandard = uniqueName;
        }

        public string Name { get; }

        public string Caption { get; }

        public string SourceName { get; }

        public string SourceNameStandard { get; }

        public int Position { get; set; }

        public int VisibleSetCalls { get; private set; }

        public bool Visible
        {
            get => visible;
            set
            {
                visible = value;
                VisibleSetCalls++;
            }
        }
    }

    public sealed class FakeUnreadablePivotFiltersField
    {
        public object PivotFilters =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeNoPivotFiltersField
    {
        public object PivotFilters =>
            throw new COMException("Application-defined or object-defined error", unchecked((int)0x800A03EC));
    }

    public sealed class FakeAutoSortedField
    {
        public int AutoSortOrder => 1;
    }

    public sealed class FakeShowAllItemsField
    {
        public FakeShowAllItemsField(bool showAllItems)
        {
            ShowAllItems = showAllItems;
        }

        public bool ShowAllItems { get; }
    }

    public sealed class FakeUnreadableShowAllItemsField
    {
        public bool ShowAllItems =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeIncludeNewItemsField
    {
        public FakeIncludeNewItemsField(bool includeNewItemsInFilter)
        {
            IncludeNewItemsInFilter = includeNewItemsInFilter;
        }

        public bool IncludeNewItemsInFilter { get; }
    }

    public sealed class FakeUnreadableIncludeNewItemsField
    {
        public bool IncludeNewItemsInFilter =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeCachePolicyPivot
    {
        private readonly object cache;

        public FakeCachePolicyPivot(object cache)
        {
            this.cache = cache;
        }

        public object PivotCache()
        {
            return cache;
        }
    }

    public class FakeCachePolicy
    {
        public FakeCachePolicy(
            bool refreshOnFileOpen = false,
            bool enableRefresh = true,
            int missingItemsLimit = -1)
        {
            RefreshOnFileOpen = refreshOnFileOpen;
            EnableRefresh = enableRefresh;
            MissingItemsLimit = missingItemsLimit;
        }

        public virtual bool RefreshOnFileOpen { get; }

        public virtual bool EnableRefresh { get; }

        public virtual int MissingItemsLimit { get; }
    }

    public sealed class FakeUnreadableCachePolicy : FakeCachePolicy
    {
        public override bool RefreshOnFileOpen =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeSaveDataPivot
    {
        public FakeSaveDataPivot(bool saveData)
        {
            SaveData = saveData;
        }

        public bool SaveData { get; }
    }

    public sealed class FakeUnreadableSaveDataPivot
    {
        public bool SaveData =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakePageItem
    {
        public string Name => "North";

        public string SourceName => "[Model].[Region].&[North]";

        public string SourceNameStandard => SourceName;

        public string Caption => "North";

        public bool Visible { get; set; } = true;

        public int Position { get; set; } = 1;
    }

    public sealed class FakeFormattingPivot
    {
        public FakeFormattingPivot(bool preserveFormatting)
        {
            PreserveFormatting = preserveFormatting;
        }

        public bool PreserveFormatting { get; }
    }

    public sealed class FakeStyleTarget
    {
        public bool RowGrand { get; set; }

        public bool ColumnGrand { get; set; }

        public bool DisplayFieldCaptions { get; set; }

        public bool PreserveFormatting { get; set; }

        public bool ShowTableStyleRowStripes { get; set; }

        public bool ShowTableStyleColumnStripes { get; set; }

        public bool DisplayNullString { get; set; }

        public string NullString { get; set; } = string.Empty;

        public bool DisplayErrorString { get; set; }

        public string ErrorString { get; set; } = string.Empty;

        public bool ShowDrillIndicators { get; set; }

        public bool EnableDrilldown { get; set; }

        public bool VisualTotals { get; set; }

        public bool SubtotalHiddenPageItems { get; set; }

        public int PageFieldOrder { get; set; }

        public int PageFieldWrapCount { get; set; }

        public int CompactRowIndent { get; set; }

        public bool MergeLabels { get; set; }

        public string TableStyle2 { get; set; } = string.Empty;

        public int Layout { get; private set; }

        public void RowAxisLayout(int layout)
        {
            Layout = layout;
        }
    }

    public sealed class FakeFormattingRange
    {
        public FakeFormattingRange(int conditionCount)
        {
            FormatConditions = new FakeCountCollection(conditionCount);
        }

        public FakeCountCollection FormatConditions { get; }
    }

    public sealed class FakeCellMetadataRange
    {
        private readonly bool hasValidation;
        private readonly bool unreadableValidation;

        public FakeCellMetadataRange(
            FakeCellComment? legacyComment = null,
            FakeCellComment? threadedComment = null,
            bool hasValidation = false,
            bool unreadableComments = false,
            bool unreadableValidation = false,
            int hyperlinkCount = 0,
            bool unreadableHyperlinks = false)
        {
            this.hasValidation = hasValidation;
            this.unreadableValidation = unreadableValidation;
            Hyperlinks = unreadableHyperlinks
                ? new FakeUnreadableCountCollection()
                : new FakeCountCollection(hyperlinkCount);
            Parent = new FakeCellMetadataWorksheet(
                legacyComment,
                threadedComment,
                unreadableComments);
        }

        public int Row => 2;

        public int Column => 3;

        public FakeDimensionCollection Rows { get; } =
            new FakeDimensionCollection(4);

        public FakeDimensionCollection Columns { get; } =
            new FakeDimensionCollection(5);

        public object Hyperlinks { get; }

        public FakeCellMetadataWorksheet Parent { get; }

        public object SpecialCells(int cellType)
        {
            Assert.Equal(-4174, cellType);
            if (unreadableValidation)
            {
                throw new COMException(
                    "Validation dispatch failed.",
                    unchecked((int)0x800A03EC));
            }

            if (hasValidation)
            {
                return new object();
            }

            throw new COMException(
                "No cells were found.",
                unchecked((int)0x800A03EC));
        }
    }

    public sealed class FakeCellMetadataWorksheet
    {
        private readonly bool unreadableComments;
        private readonly FakeCollection<FakeCellComment> threadedComments;

        public FakeCellMetadataWorksheet(
            FakeCellComment? legacyComment,
            FakeCellComment? threadedComment,
            bool unreadableComments)
        {
            this.unreadableComments = unreadableComments;
            Comments = new FakeCollection<FakeCellComment>(
                legacyComment == null
                    ? Array.Empty<FakeCellComment>()
                    : new[] { legacyComment });
            threadedComments = new FakeCollection<FakeCellComment>(
                threadedComment == null
                    ? Array.Empty<FakeCellComment>()
                    : new[] { threadedComment });
        }

        public FakeCollection<FakeCellComment> Comments { get; }

        public FakeCollection<FakeCellComment> CommentsThreaded =>
            unreadableComments
                ? throw new InvalidOperationException("RPC read failed")
                : threadedComments;
    }

    public sealed class FakeUnreadableCountCollection
    {
        public int Count =>
            throw new InvalidOperationException("RPC read failed");
    }

    public sealed class FakeCellComment
    {
        public FakeCellComment(int row, int column)
        {
            Parent = new FakeCommentCell(row, column);
        }

        public FakeCommentCell Parent { get; }
    }

    public sealed class FakeCommentCell
    {
        public FakeCommentCell(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }
    }

    public sealed class FakeFormatBackupWorksheet
    {
        public FakeFormatBackupWorksheet()
        {
            Range = new FakeFormatBackupRangeLookup();
            CustomProperties = new FakeWorksheetCustomProperties();
            PivotTables = new FakeCollection<object>(Array.Empty<object>());
            ListObjects = new FakeCollection<object>(Array.Empty<object>());
            Shapes = new FakeCollection<object>(Array.Empty<object>());
            UsedRange = new FakeEmptyUsedRange();
        }

        public string Name { get; set; } = string.Empty;

        public int Visible { get; set; }

        public bool Deleted { get; private set; }

        public FakeFormatBackupRangeLookup Range { get; }

        public FakeWorksheetCustomProperties CustomProperties { get; }

        public FakeCollection<object> PivotTables { get; }

        public FakeCollection<object> ListObjects { get; }

        public FakeCollection<object> Shapes { get; }

        public FakeEmptyUsedRange UsedRange { get; }

        public void Delete()
        {
            Deleted = true;
        }
    }

    public sealed class FakeFormatBackupRangeLookup
    {
        public FakeFormatBackupRange Anchor { get; } = new FakeFormatBackupRange();

        public FakeFormatBackupRange this[string address]
        {
            get
            {
                Assert.Equal("A1", address);
                return Anchor;
            }
        }
    }

    public sealed class FakeFormatBackupRange
    {
        public FakeFormatBackupRange()
        {
            Resize = new FakeFormatResizeLookup(this);
        }

        public FakeFormatResizeLookup Resize { get; }

        public FakeDimensionCollection Rows { get; } =
            new FakeDimensionCollection(1);

        public FakeDimensionCollection Columns { get; } =
            new FakeDimensionCollection(1);

        public int CopyCalls { get; private set; }

        public List<int> PasteTypes { get; } = new List<int>();

        public void Copy()
        {
            CopyCalls++;
        }

        public void PasteSpecial(int pasteType)
        {
            PasteTypes.Add(pasteType);
        }
    }

    public sealed class FakeFormatResizeLookup
    {
        private readonly FakeFormatBackupRange range;

        public FakeFormatResizeLookup(FakeFormatBackupRange range)
        {
            this.range = range;
        }

        public FakeFormatBackupRange this[int rows, int columns]
        {
            get
            {
                Assert.Equal(1, rows);
                Assert.Equal(1, columns);
                return range;
            }
        }
    }

    public sealed class FakeDimensionCollection
    {
        private readonly List<FakeDimension> items;

        public FakeDimensionCollection(int count)
        {
            items = Enumerable.Range(0, count)
                .Select(_ => new FakeDimension())
                .ToList();
        }

        public int Count => items.Count;

        public FakeDimension Item(int index)
        {
            return items[index - 1];
        }
    }

    public sealed class FakeDimension
    {
        public double RowHeight { get; set; } = 15d;

        public double ColumnWidth { get; set; } = 8.43d;
    }

    public sealed class FakeWorksheetCustomProperties
    {
        private readonly List<FakeWorksheetCustomProperty> items =
            new List<FakeWorksheetCustomProperty>();

        public int Count => items.Count;

        public FakeWorksheetCustomProperty Item(int index)
        {
            return items[index - 1];
        }

        public void Add(string name, string value)
        {
            items.Add(new FakeWorksheetCustomProperty(name, value));
        }
    }

    public sealed class FakeWorksheetCustomProperty
    {
        public FakeWorksheetCustomProperty(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public string Name { get; }

        public string Value { get; }
    }

    public sealed class FakeEmptyUsedRange
    {
        public FakeRawCells Cells { get; } = new FakeRawCells(1);

        public object? Value2 => null;

        public object? Formula => null;
    }
}
