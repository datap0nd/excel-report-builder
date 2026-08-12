using System.Runtime.InteropServices;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotPlusContextDiscoveryTests
{
    [Fact]
    public void Requires_the_active_cell_to_be_inside_a_pivot_table()
    {
        var application = new FakeExcelApplication(new FakeCell(null, throwOnRead: true));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            new PivotTableContextDiscovery().Discover(application));

        Assert.Contains("Select a cell inside a PivotTable", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Rejects_a_null_excel_application()
    {
        Assert.Throws<ArgumentNullException>(() =>
            new PivotTableContextDiscovery().Discover(null!));
    }

    [Fact]
    public void Discovery_resolves_a_session_stable_identity_without_writing_workbook_xml()
    {
        FakePivotTable pivot = CreatePivot(new FakePivotCache(isOlap: false));
        var application = new FakeExcelApplication(new FakeCell(pivot));

        PivotTableContext first = new PivotTableContextDiscovery().Discover(application);
        PivotTableContext second = new PivotTableContextDiscovery().Discover(application);

        Assert.Equal(first.Definition.Target.WorkbookId, second.Definition.Target.WorkbookId);
        Assert.StartsWith("workbook_", first.Definition.Target.WorkbookId, StringComparison.Ordinal);
        Assert.Equal(0, pivot.Parent.Parent.CustomXMLParts.TotalCount);
    }

    [Fact]
    public void Classifies_a_classic_pivot_and_exposes_only_classic_capabilities()
    {
        FakePivotTable pivot = CreatePivot(new FakePivotCache(isOlap: false));

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.WorksheetRange, context.Definition.Source.Kind);
        Assert.True(context.IsConnected);
        Assert.True(HasCapability(context, PivotCapability.NativeFieldPlacement));
        Assert.True(HasCapability(context, PivotCapability.UpgradeToDataModel));
        Assert.False(HasCapability(context, PivotCapability.ModelMeasures));
        Assert.False(HasCapability(context, PivotCapability.CalculatedMembers));
        Assert.False(HasCapability(context, PivotCapability.NamedSets));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(-4148)]
    public void Rejects_non_worksheet_non_olap_pivot_cache_sources(int sourceType)
    {
        FakePivotTable pivot = CreatePivot(
            new FakePivotCache(isOlap: false, sourceType: sourceType));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            Discover(pivot));

        Assert.Contains("worksheet table and range", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Classifies_a_classic_source_as_a_worksheet_table_when_the_table_exists()
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Data", workbook);
        worksheet.AddTable("SalesTable");
        var pivot = new FakePivotTable(
            "TablePivot",
            worksheet,
            new FakePivotCache(isOlap: false, sourceData: "SalesTable"));

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.WorksheetTable, context.Definition.Source.Kind);
        Assert.Equal("SalesTable", context.Definition.Source.SourceName);
        Assert.True(PivotPlusValidator.Validate(context.Definition).IsValid);
    }

    [Fact]
    public void Classifies_a_data_model_pivot_from_connection_type_seven()
    {
        FakePivotTable pivot = CreatePivot(
            new FakePivotCache(isOlap: true, new FakeWorkbookConnection(7)));

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.DataModel, context.Definition.Source.Kind);
        Assert.True(context.IsConnected);
        Assert.True(HasCapability(context, PivotCapability.NativeFieldPlacement));
        Assert.True(HasCapability(context, PivotCapability.DataModel));
        Assert.True(HasCapability(context, PivotCapability.ModelMeasures));
        Assert.True(HasCapability(context, PivotCapability.CalculatedMembers));
        Assert.True(HasCapability(context, PivotCapability.NamedSets));
    }

    [Fact]
    public void Classifies_a_connected_non_model_olap_pivot_as_external_olap()
    {
        FakePivotTable pivot = CreatePivot(
            new FakePivotCache(isOlap: true, new FakeWorkbookConnection(1)));

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.ExternalOlap, context.Definition.Source.Kind);
        Assert.True(context.IsConnected);
        Assert.True(HasCapability(context, PivotCapability.NativeFieldPlacement));
        Assert.False(HasCapability(context, PivotCapability.ModelMeasures));
        Assert.True(HasCapability(context, PivotCapability.CalculatedMembers));
        Assert.True(HasCapability(context, PivotCapability.NamedSets));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Classifies_missing_or_unreadable_olap_connection_as_disconnected(bool throwOnRead)
    {
        FakePivotTable pivot = CreatePivot(
            new FakePivotCache(isOlap: true, connection: null, throwOnConnectionRead: throwOnRead));

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.ExternalOlap, context.Definition.Source.Kind);
        Assert.False(context.IsConnected);
        Assert.Equal(PivotCapability.None, context.Definition.Source.Capabilities);
    }

    [Fact]
    public void Reads_path_free_identity_all_source_fields_and_current_classic_layout()
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Analysis", workbook);
        var region = new FakePivotField("Region", "Region", "Region", orientation: 1, position: 1);
        var department = new FakePivotField("Department", "Department", "Department", orientation: 1, position: 2);
        var period = new FakePivotField("Period", "Period", "Period", orientation: 2, position: 1);
        var cost = new FakePivotField("Cost", "Cost", "Cost", orientation: 4, position: 1);
        var scenario = new FakePivotField("Scenario", "Scenario", "Scenario", orientation: 3, position: 1);
        var dataField = new FakePivotField("Sum of Cost", "Sum of Cost", "Cost", orientation: 4, position: 1);
        var pivot = new FakePivotTable(
            "PivotTable1",
            worksheet,
            new FakePivotCache(isOlap: false),
            sourceFields: new[] { region, department, period, cost, scenario },
            rowFields: new[] { region, department },
            columnFields: new[] { period },
            dataFields: new[] { dataField },
            pageFields: new[] { scenario });

        PivotTableContext context = new PivotTableContextDiscovery().Discover(
            new FakeExcelApplication(new FakeCell(pivot)));

        Assert.StartsWith("workbook_", context.Definition.Target.WorkbookId, StringComparison.Ordinal);
        Assert.True(Guid.TryParseExact(
            context.Definition.Target.WorkbookId.Substring("workbook_".Length),
            "N",
            out _));
        Assert.Equal("Analysis", context.Definition.Target.WorksheetName);
        Assert.Equal("PivotTable1", context.Definition.Target.PivotTableName);
        Assert.Equal(0, workbook.PathReadCount);
        Assert.True(context.SourceFieldsComplete);
        Assert.Equal("sales_long!R1C1:R13C7", context.Definition.Source.SourceName);
        Assert.Equal(
            new[] { "Region", "Department", "Period", "Cost", "Scenario" },
            context.Definition.Fields.Select(field => field.Name));
        Assert.All(context.Definition.Fields, field =>
        {
            Assert.False(field.IsMeasure);
            Assert.Equal(PivotFieldAreaSupport.All, field.SupportedAreas);
        });
        PivotFieldPlacement[] rows = Placements(context, PivotFieldArea.Row);
        Assert.Equal(new[] { "Region", "Department" }, rows.Select(field => field.FieldName));
        Assert.Equal(new[] { 1, 2 }, rows.Select(field => field.Position));
        Assert.Equal("Period", Assert.Single(Placements(context, PivotFieldArea.Column)).FieldName);
        PivotFieldPlacement value = Assert.Single(Placements(context, PivotFieldArea.Values));
        Assert.Equal("Sum of Cost", value.Caption);
        Assert.Equal("Cost", value.FieldName);
        Assert.Equal(PivotAggregationFunction.Sum, value.Aggregation);
        Assert.Equal(PivotFieldArea.Values, value.Area);
        Assert.Equal("Scenario", Assert.Single(Placements(context, PivotFieldArea.Filter)).FieldName);
        Assert.True(PivotPlusValidator.Validate(context.Definition).IsValid);
    }

    [Fact]
    public void Reads_data_model_cube_field_kinds_and_layout_without_interop_types()
    {
        var workbook = new FakeWorkbook("Model.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        var regionCube = new FakePivotField(
            "[Data].[Region]",
            "Region",
            "[Data].[Region]",
            orientation: 1,
            position: 1,
            cubeFieldType: 1);
        var costCube = new FakePivotField(
            "[Measures].[Total Cost]",
            "Total Cost",
            "[Measures].[Total Cost]",
            orientation: 4,
            position: 1,
            cubeFieldType: 2);
        var customSetCube = new FakePivotField(
            "[PivotPlus Columns]",
            "PivotPlus Columns",
            "[PivotPlus Columns]",
            orientation: 0,
            position: 0,
            cubeFieldType: 3);
        var rowField = new FakePivotField(
            "Region",
            "Region",
            "[Data].[Region]",
            orientation: 1,
            position: 1,
            cubeField: regionCube);
        var valueField = new FakePivotField(
            "Total Cost",
            "Total Cost",
            "[Measures].[Total Cost]",
            orientation: 4,
            position: 1,
            cubeField: costCube);
        var pivot = new FakePivotTable(
            "ModelPivot",
            worksheet,
            new FakePivotCache(isOlap: true, new FakeWorkbookConnection(7)),
            cubeFields: new[] { regionCube, costCube, customSetCube },
            rowFields: new[] { rowField },
            dataFields: new[] { valueField });

        PivotTableContext context = Discover(pivot);

        Assert.True(context.SourceFieldsComplete);
        Assert.Equal("Synthetic Connection", context.Definition.Source.SourceName);
        Assert.Collection(
            context.Definition.Fields,
            field =>
            {
                Assert.False(field.IsMeasure);
                Assert.Equal("Data", field.TableName);
                Assert.Equal(
                    PivotFieldAreaSupport.Row |
                    PivotFieldAreaSupport.Column |
                    PivotFieldAreaSupport.Filter,
                    field.SupportedAreas);
            },
            field =>
            {
                Assert.True(field.IsMeasure);
                Assert.Equal(PivotFieldAreaSupport.Values, field.SupportedAreas);
            },
            field =>
            {
                Assert.True(field.IsCalculated);
                Assert.Equal(
                    PivotFieldAreaSupport.Row | PivotFieldAreaSupport.Column,
                    field.SupportedAreas);
            });
        Assert.Equal("[Data].[Region]", Assert.Single(Placements(context, PivotFieldArea.Row)).FieldName);
        PivotFieldPlacement valuePlacement = Assert.Single(Placements(context, PivotFieldArea.Values));
        Assert.Equal("[Measures].[Total Cost]", valuePlacement.FieldName);
        Assert.Null(valuePlacement.Aggregation);
        Assert.True(PivotPlusValidator.Validate(context.Definition).IsValid);
    }

    [Fact]
    public void Preserves_repeated_value_instances_by_source_field_caption_and_position()
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        var cost = new FakePivotField("Cost", "Cost", "Cost", orientation: 0, position: 0);
        var sum = new FakePivotField(
            "Sum of Cost",
            "Sum of Cost",
            "Cost",
            orientation: 4,
            position: 1,
            function: -4157);
        var average = new FakePivotField(
            "Average of Cost",
            "Average of Cost",
            "Cost",
            orientation: 4,
            position: 2,
            function: -4106);
        var valuesField = new FakePivotField(
            "Values",
            "Values",
            "Values",
            orientation: 2,
            position: 1);
        var pivot = new FakePivotTable(
            "RepeatedValuesPivot",
            worksheet,
            new FakePivotCache(isOlap: false),
            sourceFields: new[] { cost },
            columnFields: new[] { valuesField },
            dataFields: new[] { sum, average },
            dataPivotField: valuesField);

        PivotTableContext context = Discover(pivot);
        PivotFieldPlacement[] values = Placements(context, PivotFieldArea.Values);

        Assert.Equal(new[] { "Cost", "Cost" }, values.Select(value => value.FieldName));
        Assert.Equal(new[] { "Sum of Cost", "Average of Cost" }, values.Select(value => value.Caption));
        Assert.Equal(new[] { 1, 2 }, values.Select(value => value.Position));
        Assert.Equal(
            new PivotAggregationFunction?[]
            {
                PivotAggregationFunction.Sum,
                PivotAggregationFunction.Average
            },
            values.Select(value => value.Aggregation));
        Assert.Empty(Placements(context, PivotFieldArea.Column));
        Assert.Equal(PivotValuesAxis.Columns, context.Definition.Layout.ValuesAxis);
        Assert.Equal(1, context.Definition.Layout.ValuesPosition);
        Assert.True(PivotPlusValidator.Validate(context.Definition).IsValid);
    }

    [Fact]
    public void Reconstructs_a_partial_inventory_from_visible_layout_when_disconnected_cube_fields_fail()
    {
        var workbook = new FakeWorkbook("Disconnected.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        var region = new FakePivotField(
            "Region",
            "Region",
            "[Data].[Region]",
            orientation: 1,
            position: 1);
        var value = new FakePivotField(
            "Total Cost",
            "Total Cost",
            "[Measures].[Total Cost]",
            orientation: 4,
            position: 1);
        var pivot = new FakePivotTable(
            "DisconnectedPivot",
            worksheet,
            new FakePivotCache(isOlap: true),
            rowFields: new[] { region },
            dataFields: new[] { value },
            throwOnCubeFieldsRead: true);

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotSourceKind.ExternalOlap, context.Definition.Source.Kind);
        Assert.False(context.IsConnected);
        Assert.False(context.SourceFieldsComplete);
        Assert.Equal(
            new[] { "[Data].[Region]", "[Measures].[Total Cost]" },
            context.Definition.Fields.Select(field => field.Name));
        Assert.True(Assert.Single(
            context.Definition.Fields,
            field => field.Name == "[Measures].[Total Cost]").IsMeasure);
        Assert.Contains(
            PivotPlusValidator.Validate(context.Definition).Issues,
            issue => issue.Code == "PIVOT_OPERATION_CAPABILITY_REQUIRED");
    }

    [Fact]
    public void Maps_native_layout_and_format_metadata_into_the_core_definition()
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        var region = new FakePivotField(
            "Region",
            "Region",
            "Region",
            orientation: 1,
            position: 1,
            repeatLabels: true);
        var pivot = new FakePivotTable(
            "StyledPivot",
            worksheet,
            new FakePivotCache(isOlap: false),
            sourceFields: new[] { region },
            rowFields: new[] { region },
            layoutRowDefault: 1,
            rowGrand: false,
            columnGrand: true,
            displayFieldCaptions: false,
            tableStyle2: "PivotStyleMedium9",
            preserveFormatting: false,
            showRowStripes: true,
            showColumnStripes: true);

        PivotTableContext context = Discover(pivot);

        Assert.Equal(PivotLayoutForm.Tabular, context.Definition.Layout.Form);
        Assert.True(context.Definition.Layout.RepeatItemLabels);
        Assert.False(context.Definition.Layout.ShowRowGrandTotals);
        Assert.True(context.Definition.Layout.ShowColumnGrandTotals);
        Assert.False(context.Definition.Layout.ShowFieldHeaders);
        Assert.Equal("PivotStyleMedium9", context.Definition.Format.PivotTableStyleName);
        Assert.False(context.Definition.Format.PreserveFormatting);
        Assert.True(context.Definition.Format.ShowRowStripes);
        Assert.True(context.Definition.Format.ShowColumnStripes);
    }

    [Fact]
    public void Does_not_copy_a_file_path_from_classic_source_data_into_the_core_contract()
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        var pivot = new FakePivotTable(
            "PathFreePivot",
            worksheet,
            new FakePivotCache(
                isOlap: false,
                sourceData: @"C:\Sensitive\Source.xlsx!Data!R1C1:R10C4"));

        PivotTableContext context = Discover(pivot);

        Assert.Equal("PathFreePivot", context.Definition.Source.SourceName);
        Assert.DoesNotContain("Sensitive", context.Definition.Source.SourceName, StringComparison.Ordinal);
        Assert.True(PivotPlusValidator.Validate(context.Definition).IsValid);
    }

    private static bool HasCapability(PivotTableContext context, PivotCapability capability)
    {
        return (context.Definition.Source.Capabilities & capability) == capability;
    }

    private static PivotFieldPlacement[] Placements(
        PivotTableContext context,
        PivotFieldArea area)
    {
        return context.Definition.Placements
            .Where(placement => placement.Area == area)
            .OrderBy(placement => placement.Position)
            .ToArray();
    }

    private static PivotTableContext Discover(FakePivotTable pivot)
    {
        return new PivotTableContextDiscovery(
            new LateBoundActivePivotTableAccessor(),
            new FakeWorkbookIdentityResolver("workbook-synthetic")).Discover(
            new FakeExcelApplication(new FakeCell(pivot)));
    }

    private sealed class FakeWorkbookIdentityResolver : IWorkbookIdentityResolver
    {
        private readonly string workbookId;

        public FakeWorkbookIdentityResolver(string workbookId)
        {
            this.workbookId = workbookId;
        }

        public string Resolve(object workbook)
        {
            return workbookId;
        }

        public void Persist(object workbook, string expectedWorkbookId)
        {
            throw new InvalidOperationException("Discovery must not persist a workbook identity.");
        }
    }

    private static FakePivotTable CreatePivot(FakePivotCache cache)
    {
        var workbook = new FakeWorkbook("Synthetic.xlsx");
        var worksheet = new FakeWorksheet("Pivot", workbook);
        return new FakePivotTable("PivotTable1", worksheet, cache);
    }

    public sealed class FakeExcelApplication
    {
        public FakeExcelApplication(object activeCell)
        {
            ActiveCell = activeCell;
        }

        public object ActiveCell { get; }
    }

    public sealed class FakeCell
    {
        private readonly object? pivotTable;
        private readonly bool throwOnRead;

        public FakeCell(object? pivotTable, bool throwOnRead = false)
        {
            this.pivotTable = pivotTable;
            this.throwOnRead = throwOnRead;
        }

        public object? PivotTable
        {
            get
            {
                if (throwOnRead)
                {
                    throw new COMException("The active cell is not in a PivotTable.");
                }

                return pivotTable;
            }
        }
    }

    public sealed class FakeWorkbook
    {
        private readonly List<FakeWorksheet> worksheets = new();

        public FakeWorkbook(string name)
        {
            Name = name;
            Worksheets = new FakeWorksheetCollection(worksheets);
        }

        public string Name { get; }

        public WorkbookSpecStoreTests.FakeCustomXmlParts CustomXMLParts { get; } = new();

        public FakeWorksheetCollection Worksheets { get; }

        public int PathReadCount { get; private set; }

        public string FullName
        {
            get
            {
                PathReadCount++;
                throw new InvalidOperationException("The discovery layer must not read workbook paths.");
            }
        }

        public string Path
        {
            get
            {
                PathReadCount++;
                throw new InvalidOperationException("The discovery layer must not read workbook paths.");
            }
        }

        public void Register(FakeWorksheet worksheet)
        {
            worksheets.Add(worksheet);
        }
    }

    public sealed class FakeWorksheet
    {
        private readonly List<FakeListObject> tables = new();

        public FakeWorksheet(string name, FakeWorkbook workbook)
        {
            Name = name;
            Parent = workbook;
            ListObjects = new FakeListObjectCollection(tables);
            workbook.Register(this);
        }

        public string Name { get; }

        public FakeWorkbook Parent { get; }

        public FakeListObjectCollection ListObjects { get; }

        public void AddTable(string name)
        {
            tables.Add(new FakeListObject(name));
        }
    }

    public sealed class FakeWorksheetCollection
    {
        private readonly IReadOnlyList<FakeWorksheet> worksheets;

        public FakeWorksheetCollection(IReadOnlyList<FakeWorksheet> worksheets)
        {
            this.worksheets = worksheets;
        }

        public int Count => worksheets.Count;

        public FakeWorksheet Item(int index)
        {
            return worksheets[index - 1];
        }
    }

    public sealed class FakeListObjectCollection
    {
        private readonly IReadOnlyList<FakeListObject> tables;

        public FakeListObjectCollection(IReadOnlyList<FakeListObject> tables)
        {
            this.tables = tables;
        }

        public int Count => tables.Count;

        public FakeListObject Item(int index)
        {
            return tables[index - 1];
        }
    }

    public sealed class FakeListObject
    {
        public FakeListObject(string name)
        {
            Name = name;
        }

        public string Name { get; }

        public string DisplayName => Name;
    }

    public sealed class FakePivotTable
    {
        private readonly FakePivotCache cache;
        private readonly FakePivotFieldCollection cubeFields;
        private readonly bool throwOnCubeFieldsRead;

        public FakePivotTable(
            string name,
            FakeWorksheet worksheet,
            FakePivotCache cache,
            IReadOnlyList<FakePivotField>? sourceFields = null,
            IReadOnlyList<FakePivotField>? cubeFields = null,
            IReadOnlyList<FakePivotField>? rowFields = null,
            IReadOnlyList<FakePivotField>? columnFields = null,
            IReadOnlyList<FakePivotField>? dataFields = null,
            IReadOnlyList<FakePivotField>? pageFields = null,
            bool throwOnCubeFieldsRead = false,
            int layoutRowDefault = 0,
            bool rowGrand = true,
            bool columnGrand = true,
            bool displayFieldCaptions = true,
            string tableStyle2 = "",
            bool preserveFormatting = true,
            bool showRowStripes = false,
            bool showColumnStripes = false,
            FakePivotField? dataPivotField = null)
        {
            Name = name;
            Parent = worksheet;
            this.cache = cache;
            PivotFields = new FakePivotFieldCollection(sourceFields);
            this.cubeFields = new FakePivotFieldCollection(cubeFields);
            RowFields = new FakePivotFieldCollection(rowFields);
            ColumnFields = new FakePivotFieldCollection(columnFields);
            DataFields = new FakePivotFieldCollection(dataFields);
            PageFields = new FakePivotFieldCollection(pageFields);
            this.throwOnCubeFieldsRead = throwOnCubeFieldsRead;
            LayoutRowDefault = layoutRowDefault;
            RowGrand = rowGrand;
            ColumnGrand = columnGrand;
            DisplayFieldCaptions = displayFieldCaptions;
            TableStyle2 = tableStyle2;
            PreserveFormatting = preserveFormatting;
            ShowTableStyleRowStripes = showRowStripes;
            ShowTableStyleColumnStripes = showColumnStripes;
            DataPivotField = dataPivotField;
        }

        public string Name { get; }

        public FakeWorksheet Parent { get; }

        public FakePivotFieldCollection PivotFields { get; }

        public FakePivotFieldCollection CubeFields
        {
            get
            {
                if (throwOnCubeFieldsRead)
                {
                    throw new COMException("The disconnected cube cannot expose CubeFields.");
                }

                return cubeFields;
            }
        }

        public FakePivotFieldCollection RowFields { get; }

        public FakePivotFieldCollection ColumnFields { get; }

        public FakePivotFieldCollection DataFields { get; }

        public FakePivotFieldCollection PageFields { get; }

        public FakePivotField? DataPivotField { get; }

        public int LayoutRowDefault { get; }

        public bool RowGrand { get; }

        public bool ColumnGrand { get; }

        public bool DisplayFieldCaptions { get; }

        public string TableStyle2 { get; }

        public bool PreserveFormatting { get; }

        public bool ShowTableStyleRowStripes { get; }

        public bool ShowTableStyleColumnStripes { get; }

        public FakePivotCache PivotCache()
        {
            return cache;
        }
    }

    public sealed class FakePivotCache
    {
        private readonly object? connection;
        private readonly bool throwOnConnectionRead;

        public FakePivotCache(
            bool isOlap,
            object? connection = null,
            bool throwOnConnectionRead = false,
            string sourceData = "sales_long!R1C1:R13C7",
            int sourceType = 1)
        {
            OLAP = isOlap;
            this.connection = connection;
            this.throwOnConnectionRead = throwOnConnectionRead;
            SourceData = sourceData;
            SourceType = sourceType;
        }

        public bool OLAP { get; }

        public string SourceData { get; }

        public int SourceType { get; }

        public object? WorkbookConnection
        {
            get
            {
                if (throwOnConnectionRead)
                {
                    throw new COMException("The workbook connection was removed.");
                }

                return connection;
            }
        }
    }

    public sealed class FakeWorkbookConnection
    {
        public FakeWorkbookConnection(int type, string name = "Synthetic Connection")
        {
            Type = type;
            Name = name;
        }

        public int Type { get; }

        public string Name { get; }
    }

    public sealed class FakePivotFieldCollection
    {
        private readonly IReadOnlyList<FakePivotField> fields;

        public FakePivotFieldCollection(IReadOnlyList<FakePivotField>? fields)
        {
            this.fields = fields ?? Array.Empty<FakePivotField>();
        }

        public int Count => fields.Count;

        public FakePivotField Item(int index)
        {
            return fields[index - 1];
        }
    }

    public sealed class FakePivotField
    {
        public FakePivotField(
            string name,
            string caption,
            string sourceName,
            int orientation,
            int position,
            int? cubeFieldType = null,
            object? cubeField = null,
            int function = -4157,
            string numberFormat = "#,##0",
            bool isCalculated = false,
            bool? repeatLabels = null)
        {
            Name = name;
            Caption = caption;
            SourceName = sourceName;
            Orientation = orientation;
            Position = position;
            CubeFieldType = cubeFieldType;
            CubeField = cubeField;
            Function = function;
            NumberFormat = numberFormat;
            IsCalculated = isCalculated;
            RepeatLabels = repeatLabels;
        }

        public string Name { get; }

        public string Caption { get; }

        public string SourceName { get; }

        public int Orientation { get; }

        public int Position { get; }

        public int? CubeFieldType { get; }

        public object? CubeField { get; }

        public int Function { get; }

        public string NumberFormat { get; }

        public bool IsCalculated { get; }

        public bool? RepeatLabels { get; }
    }
}
