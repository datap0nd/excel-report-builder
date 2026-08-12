using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Measures;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class LateBoundPivotModelMeasureGatewayTests
{
    [Fact]
    public void Bind_requires_the_exact_workbook_data_model_connection()
    {
        var fixture = new HostFixture();
        BoundModelMeasureTarget target = fixture.Bind();

        Assert.Same(fixture.Workbook, target.Workbook);
        Assert.Same(fixture.Model, target.Model);
        Assert.Same(fixture.ModelConnection, target.DataModelConnection);

        fixture.Pivot.Cache.WorkbookConnection = new FakeConnection("External", 7);
        Assert.Throws<NotSupportedException>(() => fixture.Bind());

        fixture.Pivot.Cache.WorkbookConnection = fixture.ModelConnection;
        fixture.Pivot.Cache.OLAP = false;
        Assert.Throws<NotSupportedException>(() => fixture.Bind());
    }

    [Fact]
    public void Bind_rejects_non_model_context_and_bounded_collection_overflow()
    {
        var fixture = new HostFixture();
        Assert.Throws<NotSupportedException>(() => fixture.Gateway.Bind(
            fixture.Workbook,
            fixture.Pivot,
            fixture.Context(PivotSourceKind.ExternalOlap)));

        fixture.Model.ModelTables.CountOverride = 513;
        Assert.Throws<NotSupportedException>(() => fixture.Bind());
    }

    [Fact]
    public void Capture_uses_strict_one_based_collections_and_exact_formula_readback()
    {
        var fixture = new HostFixture();
        FakeModelMeasure measure = fixture.AddMeasure(
            "Gross Margin",
            "DIVIDE([Gross Profit], [Revenue])\r\n",
            FakeFormat.Percentage(2, false),
            "marker");

        ModelMeasureWorkbookSnapshot snapshot = fixture.Gateway.Capture(fixture.Bind());

        LiveModelMeasureSnapshot live = Assert.Single(snapshot.Measures);
        Assert.Equal(measure.Formula, live.Formula);
        Assert.Equal("Sales", live.AssociatedTableName);
        Assert.False(string.IsNullOrWhiteSpace(live.AssociatedTableLineageFingerprint));
        Assert.False(string.IsNullOrWhiteSpace(live.LiveFingerprint));
        Assert.All(fixture.Model.ModelTables.RequestedIndices, index => Assert.True(index > 0));
        Assert.All(fixture.Model.ModelMeasures.RequestedIndices, index => Assert.True(index > 0));
        Assert.All(fixture.Workbook.Worksheets.RequestedIndices, index => Assert.True(index > 0));
        Assert.All(fixture.Sheet.PivotTableCollection.RequestedIndices, index => Assert.True(index > 0));
    }

    [Fact]
    public void Capture_reads_all_eight_documented_model_format_types()
    {
        var fixture = new HostFixture();
        fixture.AddMeasure("General", "1", FakeFormat.General());
        fixture.AddMeasure("Boolean", "TRUE()", FakeFormat.Boolean());
        fixture.AddMeasure("Whole", "1", FakeFormat.Whole(true));
        fixture.AddMeasure("Decimal", "1.25", FakeFormat.Decimal(3, true));
        fixture.AddMeasure("Percent", "0.25", FakeFormat.Percentage(1, false));
        fixture.AddMeasure("Scientific", "100000", FakeFormat.Scientific(4));
        fixture.AddMeasure("Currency", "12.50", FakeFormat.Currency(2, "EUR"));
        fixture.AddMeasure("Date", "DATE(2026,1,1)", FakeFormat.Date("yyyy-mm-dd"));

        IReadOnlyDictionary<string, ModelMeasureFormatSnapshot> formats = fixture.Gateway
            .Capture(fixture.Bind())
            .Measures
            .ToDictionary(item => item.Name, item => item.Format, StringComparer.Ordinal);

        Assert.Equal(ExcelModelMeasureFormatKind.General, formats["General"].Kind);
        Assert.Equal(ExcelModelMeasureFormatKind.Boolean, formats["Boolean"].Kind);
        Assert.Equal(ExcelModelMeasureFormatKind.WholeNumber, formats["Whole"].Kind);
        Assert.True(formats["Whole"].UseThousandsSeparator);
        Assert.Equal(ExcelModelMeasureFormatKind.DecimalNumber, formats["Decimal"].Kind);
        Assert.Equal(3, formats["Decimal"].DecimalPlaces);
        Assert.Equal(ExcelModelMeasureFormatKind.PercentageNumber, formats["Percent"].Kind);
        Assert.Equal(1, formats["Percent"].DecimalPlaces);
        Assert.Equal(ExcelModelMeasureFormatKind.ScientificNumber, formats["Scientific"].Kind);
        Assert.Equal(4, formats["Scientific"].DecimalPlaces);
        Assert.Equal(ExcelModelMeasureFormatKind.Currency, formats["Currency"].Kind);
        Assert.Equal("EUR", formats["Currency"].CurrencySymbol);
        Assert.Equal(ExcelModelMeasureFormatKind.Date, formats["Date"].Kind);
        Assert.Equal("yyyy-mm-dd", formats["Date"].DateFormatString);
    }

    [Fact]
    public void Capture_fails_closed_for_unknown_or_unreadable_format_types()
    {
        var fixture = new HostFixture();
        fixture.AddMeasure("Unknown", "1", new FakeFormat("UnknownFormat"));

        Assert.Throws<NotSupportedException>(() =>
            fixture.Gateway.Capture(fixture.Bind()));
    }

    [Fact]
    public void Capture_scans_model_measure_usage_across_worksheets()
    {
        var fixture = new HostFixture();
        fixture.AddMeasure("Owned", "1", FakeFormat.Decimal(2, true));
        fixture.Pivot.AddCubeField("Owned", 5, visible: true, "First", "0.00");
        FakeWorksheet second = fixture.AddWorksheet("Other");
        FakePivot secondPivot = fixture.AddPivot(second, "PivotOther");
        secondPivot.AddCubeField("Owned", 5, visible: true, "Second", "0.0");

        ModelMeasureWorkbookSnapshot snapshot = fixture.Gateway.Capture(fixture.Bind());

        Assert.Equal(2, snapshot.PivotUsages.Count);
        Assert.Single(snapshot.PivotUsages.Single(item => item.PivotTableName == "PivotOther").DataFields);
        Assert.Equal(
            "Owned",
            snapshot.PivotUsages.Single(item => item.PivotTableName == "PivotOther")
                .DataFields[0].ModelMeasureName);
    }

    [Fact]
    public void Create_uses_exact_five_argument_add_and_does_not_mutate_template()
    {
        var fixture = new HostFixture();
        FakeFormat template = fixture.Model.ModelFormatCurrency;
        DesiredModelMeasure desired = Desired(
            "currency",
            "Owned Currency",
            "SUM('Sales'[Amount])",
            new PivotMeasureFormat(PivotMeasureFormatKind.Currency, 2, true, "EUR"));

        LiveModelMeasureSnapshot live = fixture.Gateway.CreateMeasure(
            fixture.Bind(),
            desired);

        FakeAddCall call = Assert.Single(fixture.Model.ModelMeasures.AddCalls);
        Assert.Equal("Owned Currency", call.Name);
        Assert.Same(fixture.Table, call.Table);
        Assert.Equal(desired.Formula, call.Formula);
        Assert.Same(template, call.Format);
        Assert.Equal(desired.DescriptionMarker, call.Description);
        Assert.Equal("EUR", live.Format.CurrencySymbol);
        Assert.Equal(string.Empty, template.Symbol);
        Assert.Equal(0, fixture.Model.ModelMeasures.Items[0].NameSetCount);
    }

    [Fact]
    public void Create_reconciles_add_exception_after_commit_and_rejects_before_commit()
    {
        var after = new HostFixture();
        after.Model.ModelMeasures.AddFailure = FakeCommitFailure.AfterCommit;
        DesiredModelMeasure desired = Desired(
            "ratio",
            "Owned Ratio",
            "DIVIDE([A],[B])",
            new PivotMeasureFormat(PivotMeasureFormatKind.Percentage, 1, false));

        LiveModelMeasureSnapshot created = after.Gateway.CreateMeasure(after.Bind(), desired);
        Assert.Equal("Owned Ratio", created.Name);
        Assert.Single(after.Model.ModelMeasures.Items);

        var before = new HostFixture();
        before.Model.ModelMeasures.AddFailure = FakeCommitFailure.BeforeCommit;
        Assert.Throws<InvalidOperationException>(() =>
            before.Gateway.CreateMeasure(before.Bind(), desired));
        Assert.Empty(before.Model.ModelMeasures.Items);
    }

    [Fact]
    public void Create_accepts_canonicalized_formula_readback_and_fingerprints_host_text()
    {
        var fixture = new HostFixture();
        fixture.Model.ModelMeasures.FormulaTransform = value => "/* Excel */ " + value;
        DesiredModelMeasure desired = Desired(
            "canonical-create",
            "Canonical Create",
            "DIVIDE([A],[B])",
            new PivotMeasureFormat(PivotMeasureFormatKind.Percentage, 2, false));

        LiveModelMeasureSnapshot created = fixture.Gateway.CreateMeasure(
            fixture.Bind(),
            desired);

        Assert.Equal("/* Excel */ DIVIDE([A],[B])", created.Formula);
        Assert.NotEqual(desired.DefinitionFingerprint, created.LiveFingerprint);
        Assert.Equal(desired.DescriptionMarker, created.Description);
    }

    [Fact]
    public void Update_is_in_place_never_sets_name_and_reconciles_last_setter_exception()
    {
        var fixture = new HostFixture();
        FakeModelMeasure native = fixture.AddMeasure(
            "Owned",
            "1",
            FakeFormat.Decimal(0, false),
            "old");
        LiveModelMeasureSnapshot before = Assert.Single(
            fixture.Gateway.Capture(fixture.Bind()).Measures);
        native.ThrowAfterDescriptionSet = true;
        DesiredModelMeasure desired = Desired(
            "owned",
            "Owned",
            "2",
            new PivotMeasureFormat(PivotMeasureFormatKind.DecimalNumber, 2, true));

        LiveModelMeasureSnapshot updated = fixture.Gateway.UpdateMeasure(
            fixture.Bind(),
            before,
            desired);

        Assert.Same(native, fixture.Model.ModelMeasures.Items[0]);
        Assert.Equal("2", updated.Formula);
        Assert.Equal(2, updated.Format.DecimalPlaces);
        Assert.True(updated.Format.UseThousandsSeparator);
        Assert.Equal(0, native.NameSetCount);
    }

    [Fact]
    public void Update_accepts_canonicalized_formula_readback_with_exact_definition_marker()
    {
        var fixture = new HostFixture();
        fixture.AddMeasure("Owned", "1", FakeFormat.Decimal(0, false), "old");
        LiveModelMeasureSnapshot before = Assert.Single(
            fixture.Gateway.Capture(fixture.Bind()).Measures);
        fixture.Model.ModelMeasures.FormulaTransform = value => "/* Excel */ " + value;
        DesiredModelMeasure desired = Desired(
            "canonical-update",
            "Owned",
            "2",
            new PivotMeasureFormat(PivotMeasureFormatKind.DecimalNumber, 1, false));

        LiveModelMeasureSnapshot updated = fixture.Gateway.UpdateMeasure(
            fixture.Bind(),
            before,
            desired);

        Assert.Equal("/* Excel */ 2", updated.Formula);
        Assert.Equal(desired.DescriptionMarker, updated.Description);
    }

    [Fact]
    public void Update_restores_exact_prior_definition_after_partial_failure()
    {
        var fixture = new HostFixture();
        FakeModelMeasure native = fixture.AddMeasure(
            "Owned",
            "1",
            FakeFormat.Currency(2, "USD"),
            "old");
        LiveModelMeasureSnapshot before = Assert.Single(
            fixture.Gateway.Capture(fixture.Bind()).Measures);
        native.ThrowAfterFormulaSetOnce = true;
        DesiredModelMeasure desired = Desired(
            "owned",
            "Owned",
            "2",
            new PivotMeasureFormat(PivotMeasureFormatKind.Currency, 0, false, "EUR"));

        Assert.Throws<InvalidOperationException>(() =>
            fixture.Gateway.UpdateMeasure(fixture.Bind(), before, desired));

        LiveModelMeasureSnapshot restored = Assert.Single(
            fixture.Gateway.Capture(fixture.Bind()).Measures);
        Assert.Equal(before.LiveFingerprint, restored.LiveFingerprint);
        Assert.Equal("1", native.Formula);
        Assert.Equal("USD", ((FakeFormat)native.FormatInformation).Symbol);
    }

    [Fact]
    public void Delete_reconciles_before_and_after_commit_failures()
    {
        var after = new HostFixture();
        FakeModelMeasure afterNative = after.AddMeasure("Owned", "1", FakeFormat.General());
        LiveModelMeasureSnapshot afterSnapshot = Assert.Single(
            after.Gateway.Capture(after.Bind()).Measures);
        afterNative.DeleteFailure = FakeCommitFailure.AfterCommit;

        after.Gateway.DeleteMeasure(after.Bind(), afterSnapshot);
        Assert.Empty(after.Model.ModelMeasures.Items);

        var before = new HostFixture();
        FakeModelMeasure beforeNative = before.AddMeasure("Owned", "1", FakeFormat.General());
        LiveModelMeasureSnapshot beforeSnapshot = Assert.Single(
            before.Gateway.Capture(before.Bind()).Measures);
        beforeNative.DeleteFailure = FakeCommitFailure.BeforeCommit;

        Assert.Throws<InvalidOperationException>(() =>
            before.Gateway.DeleteMeasure(before.Bind(), beforeSnapshot));
        Assert.Single(before.Model.ModelMeasures.Items);
    }

    [Fact]
    public void Restore_recreates_deleted_measure_with_exact_eight_type_snapshot()
    {
        var fixture = new HostFixture();
        FakeModelMeasure native = fixture.AddMeasure(
            "Owned Date",
            "DATE(2026,1,1)",
            FakeFormat.Date("yyyy-mm-dd"),
            "old");
        LiveModelMeasureSnapshot before = Assert.Single(
            fixture.Gateway.Capture(fixture.Bind()).Measures);
        native.Delete();

        LiveModelMeasureSnapshot restored = fixture.Gateway.RestoreMeasure(
            fixture.Bind(),
            before);

        Assert.Equal(before.LiveFingerprint, restored.LiveFingerprint);
        Assert.Equal(ExcelModelMeasureFormatKind.Date, restored.Format.Kind);
        Assert.Equal("yyyy-mm-dd", restored.Format.DateFormatString);
    }

    [Fact]
    public void ApplyPlacement_supports_arbitrary_interleave_and_values_pseudo_axis()
    {
        var fixture = new HostFixture();
        fixture.Pivot.AddImplicitCubeField("[Measures].[Actual]", "Actual caption", "#,##0");
        fixture.Pivot.AddImplicitCubeField("[Measures].[Plan]", "Plan caption", "0.0");
        fixture.Pivot.DataPivotField.Orientation = 2;
        fixture.Pivot.DataPivotField.Position = 1;
        fixture.AddMeasure("Owned", "1", FakeFormat.Decimal(2, true), "marker");
        fixture.Pivot.AddCubeField("Owned", 5, visible: false, "Owned", "0.00");
        BoundModelMeasureTarget target = fixture.Bind();
        ModelMeasureWorkbookSnapshot before = fixture.Gateway.Capture(target);
        ModelDataFieldSnapshot actual = before.SelectedPivot.DataFields[0];
        ModelDataFieldSnapshot plan = before.SelectedPivot.DataFields[1];
        DesiredModelMeasure desired = Desired(
            "owned",
            "Owned",
            "1",
            new PivotMeasureFormat(PivotMeasureFormatKind.DecimalNumber, 2, true));
        var definitions = new Dictionary<string, DesiredModelMeasure>(StringComparer.Ordinal)
        {
            [desired.DefinitionId] = desired
        };
        var placement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, Existing(plan)),
                new PivotMeasureValuePlacement(2, desired.DefinitionId),
                new PivotMeasureValuePlacement(3, Existing(actual))
            },
            PivotValuesAxis.Rows,
            2);

        fixture.Gateway.ApplyPlacement(target, placement, definitions, before);

        Assert.Equal(
            new[] { "[Measures].[Plan]", "[Measures].[Owned]", "[Measures].[Actual]" },
            fixture.Pivot.DataFieldItems.Select(item => item.CubeField.Name));
        Assert.Equal("Plan caption", fixture.Pivot.DataFieldItems[0].Caption);
        Assert.Equal("0.0", fixture.Pivot.DataFieldItems[0].NumberFormat);
        Assert.Equal("Actual caption", fixture.Pivot.DataFieldItems[2].Caption);
        Assert.Equal("#,##0", fixture.Pivot.DataFieldItems[2].NumberFormat);
        Assert.Equal(1, fixture.Pivot.DataPivotField.Orientation);
        Assert.Equal(2, fixture.Pivot.DataPivotField.Position);
        Assert.Equal(0, fixture.Pivot.CubeFields.GetMeasureCalls);
    }

    [Fact]
    public void ApplyAndRestorePlacement_preserve_repeated_existing_value_occurrences()
    {
        var fixture = new HostFixture();
        FakeCubeField cube = fixture.Pivot.AddImplicitCubeField(
            "[Measures].[Actual]",
            "Actual",
            "#,##0");
        FakeDataField firstNative = fixture.Pivot.DataFieldItems.Single();
        FakeDataField secondNative = fixture.Pivot.AddDuplicateDataField(
            cube,
            "Actual",
            "#,##0");
        ModelMeasureWorkbookSnapshot before = fixture.Gateway.Capture(fixture.Bind());
        ModelDataFieldSnapshot first = before.SelectedPivot.DataFields[0];
        ModelDataFieldSnapshot second = before.SelectedPivot.DataFields[1];
        var placement = new PivotMeasurePlacementPlan(
            new[]
            {
                new PivotMeasureValuePlacement(1, Existing(second)),
                new PivotMeasureValuePlacement(2, Existing(first))
            },
            PivotValuesAxis.Columns,
            1);

        fixture.Gateway.ApplyPlacement(
            fixture.Bind(),
            placement,
            new Dictionary<string, DesiredModelMeasure>(),
            before);

        Assert.Same(secondNative, fixture.Pivot.DataFieldItems[0]);
        Assert.Same(firstNative, fixture.Pivot.DataFieldItems[1]);

        fixture.Gateway.RestorePlacement(fixture.Bind(), before.SelectedPivot);

        Assert.Equal(2, fixture.Pivot.DataFieldItems.Count);
        Assert.All(fixture.Pivot.DataFieldItems, field =>
        {
            Assert.Equal("Actual", field.Caption);
            Assert.Equal("#,##0", field.NumberFormat);
        });
    }

    [Fact]
    public void ApplyPlacement_removes_only_plan_omissions_after_complete_preview()
    {
        var fixture = new HostFixture();
        fixture.Pivot.AddImplicitCubeField("[Measures].[Keep]", "Keep", "0");
        fixture.AddMeasure("Old Owned", "1", FakeFormat.General(), "old");
        fixture.Pivot.AddCubeField("Old Owned", 5, visible: true, "Old", "0");
        ModelMeasureWorkbookSnapshot before = fixture.Gateway.Capture(fixture.Bind());
        ModelDataFieldSnapshot keep = before.SelectedPivot.DataFields.Single(item =>
            item.UniqueName == "[Measures].[Keep]");
        var placement = new PivotMeasurePlacementPlan(
            new[] { new PivotMeasureValuePlacement(1, Existing(keep)) },
            PivotValuesAxis.Automatic,
            1);

        fixture.Gateway.ApplyPlacement(
            fixture.Bind(),
            placement,
            new Dictionary<string, DesiredModelMeasure>(),
            before);

        Assert.Equal("[Measures].[Keep]", Assert.Single(fixture.Pivot.DataFieldItems).CubeField.Name);
    }

    [Fact]
    public void ApplyPlacement_rejects_implicit_cube_subtype_for_authored_measure()
    {
        var fixture = new HostFixture();
        fixture.AddMeasure("Owned", "1", FakeFormat.General(), "marker");
        fixture.Pivot.AddCubeField("Owned", 11, visible: false, "Owned", "0");
        BoundModelMeasureTarget target = fixture.Bind();
        ModelMeasureWorkbookSnapshot before = fixture.Gateway.Capture(target);
        DesiredModelMeasure desired = Desired(
            "owned",
            "Owned",
            "1",
            new PivotMeasureFormat(PivotMeasureFormatKind.WholeNumber, 0, false));
        var placement = new PivotMeasurePlacementPlan(
            new[] { new PivotMeasureValuePlacement(1, desired.DefinitionId) },
            PivotValuesAxis.Automatic,
            1);

        Assert.Throws<NotSupportedException>(() => fixture.Gateway.ApplyPlacement(
            target,
            placement,
            new Dictionary<string, DesiredModelMeasure> { [desired.DefinitionId] = desired },
            before));
        Assert.Equal(0, fixture.Pivot.CubeFields.GetMeasureCalls);
    }

    [Fact]
    public void RestorePlacement_restores_exact_values_order_styles_and_axis()
    {
        var fixture = new HostFixture();
        fixture.Pivot.AddImplicitCubeField("[Measures].[A]", "Alpha", "0.0");
        fixture.Pivot.AddImplicitCubeField("[Measures].[B]", "Beta", "#,##0");
        fixture.Pivot.DataPivotField.Orientation = 2;
        fixture.Pivot.DataPivotField.Position = 3;
        ModelMeasureWorkbookSnapshot captured = fixture.Gateway.Capture(fixture.Bind());
        fixture.Pivot.DataFieldItems[1].Position = 1;
        fixture.Pivot.DataFieldItems[0].Caption = "Changed";
        fixture.Pivot.DataFieldItems[0].NumberFormat = "General";
        fixture.Pivot.DataPivotField.Orientation = 1;
        fixture.Pivot.DataPivotField.Position = 1;

        fixture.Gateway.RestorePlacement(fixture.Bind(), captured.SelectedPivot);

        Assert.Equal(new[] { "Alpha", "Beta" }, fixture.Pivot.DataFieldItems.Select(item => item.Caption));
        Assert.Equal(new[] { "0.0", "#,##0" }, fixture.Pivot.DataFieldItems.Select(item => item.NumberFormat));
        Assert.Equal(2, fixture.Pivot.DataPivotField.Orientation);
        Assert.Equal(3, fixture.Pivot.DataPivotField.Position);
    }

    [Fact]
    public void Refresh_calls_refresh_table_exactly_once_and_requires_true()
    {
        var fixture = new HostFixture();
        BoundModelMeasureTarget target = fixture.Bind();

        fixture.Gateway.Refresh(target);
        Assert.Equal(1, fixture.Pivot.RefreshCalls);

        fixture.Pivot.RefreshResult = false;
        Assert.Throws<InvalidOperationException>(() => fixture.Gateway.Refresh(target));
        Assert.Equal(2, fixture.Pivot.RefreshCalls);
    }

    private static DesiredModelMeasure Desired(
        string id,
        string name,
        string formula,
        PivotMeasureFormat format)
    {
        return new DesiredModelMeasure(
            id,
            displayOrder: 1,
            creationOrder: 1,
            homeTableName: "Sales",
            name,
            formula,
            format,
            directDependencyDefinitionIds: Array.Empty<string>(),
            definitionFingerprint: "fingerprint-" + id,
            descriptionMarker: "PivotTable+|setup|" + id + "|fingerprint");
    }

    private static PivotExistingDataFieldIdentity Existing(ModelDataFieldSnapshot field)
    {
        return new PivotExistingDataFieldIdentity(
            field.UniqueName,
            field.CaptionFingerprint,
            PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                field.NumberFormat),
            field.Position);
    }

    public enum FakeCommitFailure
    {
        None,
        BeforeCommit,
        AfterCommit
    }

    public sealed class HostFixture
    {
        public HostFixture()
        {
            ModelConnection = new FakeConnection("ThisWorkbookDataModel", 7);
            SourceConnection = new FakeConnection("Connection - Sales", 1);
            Table = new FakeModelTable("Sales", SourceConnection);
            Model = new FakeModel(ModelConnection, Table);
            Workbook = new FakeWorkbook(Model);
            Sheet = AddWorksheet("Sheet1");
            Pivot = AddPivot(Sheet, "PivotTable1");
            Gateway = new LateBoundPivotModelMeasureGateway(value =>
                ((FakeFormat)value).TypeName);
        }

        public FakeConnection ModelConnection { get; }
        public FakeConnection SourceConnection { get; }
        public FakeModelTable Table { get; }
        public FakeModel Model { get; }
        public FakeWorkbook Workbook { get; }
        public FakeWorksheet Sheet { get; }
        public FakePivot Pivot { get; }
        internal LateBoundPivotModelMeasureGateway Gateway { get; }

        internal BoundModelMeasureTarget Bind()
        {
            return Gateway.Bind(Workbook, Pivot, Context(PivotSourceKind.DataModel));
        }

        public PivotTableContext Context(PivotSourceKind sourceKind)
        {
            string workbookId = new StoredWorkbookIdentityResolver().Resolve(Workbook);
            PivotCapability capabilities = PivotCapability.NativeFieldPlacement |
                                           PivotCapability.DataModel |
                                           PivotCapability.ModelMeasures;
            return new PivotTableContext(
                new PivotLayoutDefinition(
                    new PivotTargetIdentity(workbookId, Sheet.Name, Pivot.Name),
                    new PivotSourceDescriptor(
                        sourceKind,
                        "ThisWorkbookDataModel",
                        capabilities,
                        "Sales"),
                    fields: Array.Empty<PivotFieldDescriptor>(),
                    placements: Array.Empty<PivotFieldPlacement>(),
                    clearAll: true),
                isConnected: true,
                sourceFieldsComplete: true);
        }

        public FakeModelMeasure AddMeasure(
            string name,
            string formula,
            FakeFormat format,
            string description = "")
        {
            return Model.ModelMeasures.AddDirect(
                name,
                Table,
                formula,
                format,
                description);
        }

        public FakeWorksheet AddWorksheet(string name)
        {
            var sheet = new FakeWorksheet(name, Workbook);
            Workbook.WorksheetItems.Add(sheet);
            return sheet;
        }

        public FakePivot AddPivot(FakeWorksheet worksheet, string name)
        {
            var pivot = new FakePivot(name, worksheet, ModelConnection);
            worksheet.PivotItems.Add(pivot);
            return pivot;
        }
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(FakeModel model)
        {
            Model = model;
            Worksheets = new FakeCollection<FakeWorksheet>(() => WorksheetItems);
        }

        public FakeModel Model { get; }
        public List<FakeWorksheet> WorksheetItems { get; } = new();
        public FakeCollection<FakeWorksheet> Worksheets { get; }
        public FakeCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FakeCustomXmlParts
    {
        private readonly FakeCollection<object> empty =
            new FakeCollection<object>(() => Array.Empty<object>());

        public FakeCollection<object> SelectByNamespace(string namespaceUri)
        {
            _ = namespaceUri;
            return empty;
        }
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(string name, FakeWorkbook parent)
        {
            Name = name;
            Parent = parent;
            PivotTableCollection = new FakeCollection<FakePivot>(() => PivotItems);
        }

        public string Name { get; }
        public FakeWorkbook Parent { get; }
        public List<FakePivot> PivotItems { get; } = new();
        public FakeCollection<FakePivot> PivotTableCollection { get; }

        public FakeCollection<FakePivot> PivotTables() => PivotTableCollection;
    }

    public sealed class FakePivot
    {
        public FakePivot(string name, FakeWorksheet parent, FakeConnection connection)
        {
            Name = name;
            Parent = parent;
            Cache = new FakePivotCache(true, connection);
            CubeFields = new FakeCubeFieldCollection(this);
            DataFields = new FakeCollection<FakeDataField>(() => DataFieldItems);
        }

        public string Name { get; }
        public FakeWorksheet Parent { get; }
        public FakePivotCache Cache { get; }
        public List<FakeCubeField> CubeFieldItems { get; } = new();
        public List<FakeDataField> DataFieldItems { get; } = new();
        public FakeCubeFieldCollection CubeFields { get; }
        public FakeCollection<FakeDataField> DataFields { get; }
        public FakeDataPivotField DataPivotField { get; } = new();
        public bool RefreshResult { get; set; } = true;
        public int RefreshCalls { get; private set; }

        public FakePivotCache PivotCache() => Cache;

        public bool RefreshTable()
        {
            RefreshCalls++;
            return RefreshResult;
        }

        public FakeCubeField AddCubeField(
            string measureName,
            int subType,
            bool visible,
            string caption,
            string numberFormat)
        {
            var cube = new FakeCubeField(
                this,
                "[Measures].[" + measureName.Replace("]", "]]" ) + "]",
                2,
                subType,
                caption,
                numberFormat);
            CubeFieldItems.Add(cube);
            if (visible) cube.Orientation = 4;
            return cube;
        }

        public FakeCubeField AddImplicitCubeField(
            string uniqueName,
            string caption,
            string numberFormat)
        {
            var cube = new FakeCubeField(
                this,
                uniqueName,
                2,
                11,
                caption,
                numberFormat);
            CubeFieldItems.Add(cube);
            cube.Orientation = 4;
            return cube;
        }

        public FakeDataField AddDuplicateDataField(
            FakeCubeField cube,
            string caption,
            string numberFormat)
        {
            var field = new FakeDataField(this, cube, caption, numberFormat);
            DataFieldItems.Add(field);
            return field;
        }

        public void SetVisible(FakeCubeField cube, bool visible)
        {
            List<FakeDataField> existing = DataFieldItems.Where(item =>
                ReferenceEquals(item.CubeField, cube)).ToList();
            if (visible && existing.Count == 0)
            {
                DataFieldItems.Add(new FakeDataField(
                    this,
                    cube,
                    cube.DefaultCaption,
                    cube.DefaultNumberFormat));
            }
            else if (!visible && existing.Count > 0)
            {
                foreach (FakeDataField field in existing)
                {
                    DataFieldItems.Remove(field);
                }
            }
        }

        public void Move(FakeDataField field, int position)
        {
            if (position <= 0 || position > DataFieldItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(position));
            }

            DataFieldItems.Remove(field);
            DataFieldItems.Insert(position - 1, field);
        }
    }

    public sealed class FakePivotCache
    {
        public FakePivotCache(bool olap, FakeConnection workbookConnection)
        {
            OLAP = olap;
            WorkbookConnection = workbookConnection;
        }

        public bool OLAP { get; set; }
        public FakeConnection WorkbookConnection { get; set; }
    }

    public sealed class FakeDataPivotField
    {
        public int Orientation { get; set; } = 2;
        public int Position { get; set; } = 1;
    }

    public sealed class FakeCubeFieldCollection
    {
        private readonly FakePivot owner;

        public FakeCubeFieldCollection(FakePivot owner)
        {
            this.owner = owner;
        }

        public int Count => owner.CubeFieldItems.Count;
        public int GetMeasureCalls { get; private set; }
        public List<int> RequestedIndices { get; } = new();

        public FakeCubeField Item(int index)
        {
            RequestedIndices.Add(index);
            if (index <= 0) throw new ArgumentOutOfRangeException(nameof(index));
            return owner.CubeFieldItems[index - 1];
        }

        public object GetMeasure(object hierarchy, int function, string caption)
        {
            _ = hierarchy;
            _ = function;
            _ = caption;
            GetMeasureCalls++;
            throw new InvalidOperationException("GetMeasure must never be called.");
        }
    }

    public sealed class FakeCubeField
    {
        private readonly FakePivot owner;
        private int orientation;

        public FakeCubeField(
            FakePivot owner,
            string name,
            int cubeFieldType,
            int cubeFieldSubType,
            string defaultCaption,
            string defaultNumberFormat)
        {
            this.owner = owner;
            Name = name;
            CubeFieldType = cubeFieldType;
            CubeFieldSubType = cubeFieldSubType;
            DefaultCaption = defaultCaption;
            DefaultNumberFormat = defaultNumberFormat;
        }

        public string Name { get; }
        public int CubeFieldType { get; }
        public int CubeFieldSubType { get; }
        public string DefaultCaption { get; }
        public string DefaultNumberFormat { get; }

        public int Orientation
        {
            get => orientation;
            set
            {
                orientation = value;
                owner.SetVisible(this, value == 4);
            }
        }
    }

    public sealed class FakeDataField
    {
        private readonly FakePivot owner;

        public FakeDataField(
            FakePivot owner,
            FakeCubeField cubeField,
            string caption,
            string numberFormat)
        {
            this.owner = owner;
            CubeField = cubeField;
            Caption = caption;
            NumberFormat = numberFormat;
        }

        public FakeCubeField CubeField { get; }
        public string Caption { get; set; }
        public string NumberFormat { get; set; }

        public int Position
        {
            get => owner.DataFieldItems.IndexOf(this) + 1;
            set => owner.Move(this, value);
        }
    }

    public sealed class FakeConnection
    {
        public FakeConnection(string name, int type)
        {
            Name = name;
            Type = type;
        }

        public string Name { get; }
        public int Type { get; }
    }

    public sealed class FakeModelTable
    {
        public FakeModelTable(string name, FakeConnection sourceWorkbookConnection)
        {
            Name = name;
            SourceWorkbookConnection = sourceWorkbookConnection;
        }

        public string Name { get; }
        public FakeConnection SourceWorkbookConnection { get; }
    }

    public sealed class FakeModel
    {
        public FakeModel(FakeConnection dataModelConnection, FakeModelTable table)
        {
            DataModelConnection = dataModelConnection;
            ModelTableItems.Add(table);
            ModelTables = new FakeCollection<FakeModelTable>(() => ModelTableItems);
            ModelMeasures = new FakeModelMeasures();
        }

        public FakeConnection DataModelConnection { get; }
        public List<FakeModelTable> ModelTableItems { get; } = new();
        public FakeCollection<FakeModelTable> ModelTables { get; }
        public FakeModelMeasures ModelMeasures { get; }
        public FakeFormat ModelFormatGeneral { get; } = FakeFormat.General();
        public FakeFormat ModelFormatBoolean { get; } = FakeFormat.Boolean();
        public FakeFormat ModelFormatWholeNumber { get; } = FakeFormat.Whole(false);
        public FakeFormat ModelFormatDecimalNumber { get; } = FakeFormat.Decimal(0, false);
        public FakeFormat ModelFormatPercentageNumber { get; } = FakeFormat.Percentage(0, false);
        public FakeFormat ModelFormatScientificNumber { get; } = FakeFormat.Scientific(0);
        public FakeFormat ModelFormatCurrency { get; } = FakeFormat.Currency(0, string.Empty);
        public FakeFormat ModelFormatDate { get; } = FakeFormat.Date(string.Empty);
    }

    public sealed class FakeModelMeasures
    {
        public List<FakeModelMeasure> Items { get; } = new();
        public List<int> RequestedIndices { get; } = new();
        public List<FakeAddCall> AddCalls { get; } = new();
        public int? CountOverride { get; set; }
        public FakeCommitFailure AddFailure { get; set; }
        public Func<string, string> FormulaTransform { get; set; } = value => value;

        public int Count => CountOverride ?? Items.Count;

        public FakeModelMeasure Item(int index)
        {
            RequestedIndices.Add(index);
            if (index <= 0) throw new ArgumentOutOfRangeException(nameof(index));
            return Items[index - 1];
        }

        public object? Add(
            string name,
            object table,
            string formula,
            object format,
            string description)
        {
            AddCalls.Add(new FakeAddCall(name, table, formula, format, description));
            if (AddFailure == FakeCommitFailure.BeforeCommit)
            {
                throw new InvalidOperationException("before Add commit");
            }

            FakeModelMeasure created = AddDirect(
                name,
                (FakeModelTable)table,
                formula,
                ((FakeFormat)format).Clone(),
                description);
            if (AddFailure == FakeCommitFailure.AfterCommit)
            {
                throw new InvalidOperationException("after Add commit");
            }

            return created;
        }

        public FakeModelMeasure AddDirect(
            string name,
            FakeModelTable table,
            string formula,
            FakeFormat format,
            string description)
        {
            var measure = new FakeModelMeasure(
                this,
                name,
                table,
                FormulaTransform(formula),
                format.Clone(),
                description);
            Items.Add(measure);
            return measure;
        }
    }

    public sealed class FakeAddCall
    {
        public FakeAddCall(
            string name,
            object table,
            string formula,
            object format,
            string description)
        {
            Name = name;
            Table = table;
            Formula = formula;
            Format = format;
            Description = description;
        }

        public string Name { get; }
        public object Table { get; }
        public string Formula { get; }
        public object Format { get; }
        public string Description { get; }
    }

    public sealed class FakeModelMeasure
    {
        private readonly FakeModelMeasures owner;
        private string name;
        private object associatedTable;
        private string formula;
        private FakeFormat formatInformation;
        private string description;

        public FakeModelMeasure(
            FakeModelMeasures owner,
            string name,
            FakeModelTable associatedTable,
            string formula,
            FakeFormat formatInformation,
            string description)
        {
            this.owner = owner;
            this.name = name;
            this.associatedTable = associatedTable;
            this.formula = formula;
            this.formatInformation = formatInformation;
            this.description = description;
        }

        public int NameSetCount { get; private set; }
        public bool ThrowAfterDescriptionSet { get; set; }
        public bool ThrowAfterFormulaSetOnce { get; set; }
        public FakeCommitFailure DeleteFailure { get; set; }

        public string Name
        {
            get => name;
            set
            {
                NameSetCount++;
                name = value;
            }
        }

        public object AssociatedTable
        {
            get => associatedTable;
            set => associatedTable = value;
        }

        public string Formula
        {
            get => formula;
            set
            {
                formula = owner.FormulaTransform(value);
                if (ThrowAfterFormulaSetOnce)
                {
                    ThrowAfterFormulaSetOnce = false;
                    throw new InvalidOperationException("after Formula commit");
                }
            }
        }

        public object FormatInformation
        {
            get => formatInformation;
            set => formatInformation = ((FakeFormat)value).Clone();
        }

        public string Description
        {
            get => description;
            set
            {
                description = value;
                if (ThrowAfterDescriptionSet)
                {
                    ThrowAfterDescriptionSet = false;
                    throw new InvalidOperationException("after Description commit");
                }
            }
        }

        public void Delete()
        {
            if (DeleteFailure == FakeCommitFailure.BeforeCommit)
            {
                throw new InvalidOperationException("before Delete commit");
            }

            owner.Items.Remove(this);
            if (DeleteFailure == FakeCommitFailure.AfterCommit)
            {
                throw new InvalidOperationException("after Delete commit");
            }
        }
    }

    public sealed class FakeFormat
    {
        public FakeFormat(string typeName)
        {
            TypeName = typeName;
        }

        public string TypeName { get; }
        public int DecimalPlaces { get; set; }
        public bool UseThousandSeparator { get; set; }
        public string Symbol { get; set; } = string.Empty;
        public string FormatString { get; set; } = string.Empty;

        public FakeFormat Clone()
        {
            return new FakeFormat(TypeName)
            {
                DecimalPlaces = DecimalPlaces,
                UseThousandSeparator = UseThousandSeparator,
                Symbol = Symbol,
                FormatString = FormatString
            };
        }

        public static FakeFormat General() => new("ModelFormatGeneral");
        public static FakeFormat Boolean() => new("ModelFormatBoolean");
        public static FakeFormat Whole(bool separator) =>
            new("ModelFormatWholeNumber") { UseThousandSeparator = separator };
        public static FakeFormat Decimal(int places, bool separator) =>
            new("ModelFormatDecimalNumber")
            {
                DecimalPlaces = places,
                UseThousandSeparator = separator
            };
        public static FakeFormat Percentage(int places, bool separator) =>
            new("ModelFormatPercentageNumber")
            {
                DecimalPlaces = places,
                UseThousandSeparator = separator
            };
        public static FakeFormat Scientific(int places) =>
            new("ModelFormatScientificNumber") { DecimalPlaces = places };
        public static FakeFormat Currency(int places, string symbol) =>
            new("ModelFormatCurrency") { DecimalPlaces = places, Symbol = symbol };
        public static FakeFormat Date(string format) =>
            new("ModelFormatDate") { FormatString = format };
    }

    public sealed class FakeCollection<T>
    {
        private readonly Func<IReadOnlyList<T>> values;

        public FakeCollection(Func<IReadOnlyList<T>> values)
        {
            this.values = values;
        }

        public int? CountOverride { get; set; }
        public int Count => CountOverride ?? values().Count;
        public List<int> RequestedIndices { get; } = new();

        public T Item(int index)
        {
            RequestedIndices.Add(index);
            if (index <= 0) throw new ArgumentOutOfRangeException(nameof(index));
            return values()[index - 1];
        }
    }
}
