using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Semantics;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class LateBoundPivotSemanticLayoutGatewayTests
{
    [Fact]
    public void Binds_only_the_exact_selected_data_model_pivot()
    {
        var fixture = new HostFixture();

        BoundPivotSemanticLayoutTarget target = fixture.Bind();

        Assert.Same(fixture.Workbook, target.Workbook);
        Assert.Same(fixture.Pivot, target.PivotTable);
        Assert.Same(fixture.Model, target.Model);
        Assert.Same(fixture.Connection, target.DataModelConnection);
        Assert.Throws<NotSupportedException>(() => fixture.Gateway.Bind(
            fixture.Workbook,
            fixture.Pivot,
            fixture.Context(PivotSourceKind.ExternalOlap)));
        fixture.Pivot.Cache.WorkbookConnection = new FakeConnection(7);
        Assert.Throws<NotSupportedException>(() => fixture.Bind());
    }

    [Fact]
    public void Capture_removes_DataPivotField_by_COM_identity_and_normalizes_regular_positions()
    {
        var fixture = new HostFixture(valuesOnRows: true);

        PivotSemanticLayoutSnapshot snapshot = fixture.Capture();

        Assert.Equal(new[] { "[Region]", "[Department]" },
            snapshot.Rows.Select(field => field.UniqueName));
        Assert.Equal(new[] { 1, 2 }, snapshot.Rows.Select(field => field.Position));
        Assert.Equal(PivotValuesAxis.Rows, snapshot.ValuesAxis);
        Assert.Equal(2, snapshot.ValuesPosition);
        Assert.Equal(new[] { "[Measures].[Cost]", "[Measures].[Revenue]" },
            snapshot.Values.Select(field => field.UniqueName));
        Assert.Single(snapshot.Filters);
        Assert.StartsWith("semantic.filters.v1:sha256:", snapshot.FilterFingerprint);
        Assert.StartsWith("semantic.layout.v1:sha256:", snapshot.LayoutFingerprint);
    }

    [Fact]
    public void Applies_exact_set_replacement_interleaving_value_order_and_unchanged_filters()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        PivotSemanticPreparedPlacement prepared = fixture.Prepare(before);

        prepared.Apply();

        PivotSemanticLayoutSnapshot after = fixture.Capture();
        Assert.Equal(new[] { "[Region]", "[PivotTablePlus_Set]" },
            after.Rows.Select(field => field.UniqueName));
        Assert.Equal(new[] { 1, 2 }, after.Rows.Select(field => field.Position));
        Assert.Equal(new[] { "[Month]" },
            after.Columns.Select(field => field.UniqueName));
        Assert.Equal(new[] { "[Measures].[Variance]", "[Measures].[Cost]" },
            after.Values.Select(field => field.UniqueName));
        Assert.Equal(PivotValuesAxis.Rows, after.ValuesAxis);
        Assert.Equal(2, after.ValuesPosition);
        Assert.Equal(
            new object[] { fixture.Region.AxisField, fixture.Pivot.DataPivotField, fixture.Set.AxisField },
            fixture.Pivot.RowFields.Items);
        Assert.Equal(before.FilterFingerprint, after.FilterFingerprint);
        Assert.Equal("North", fixture.Filter.AxisField.CurrentPageName);
        Assert.False(fixture.Filter.AxisField.PivotItemItems.Single(item => item.Name == "South").Visible);
        Assert.Equal(0, fixture.Pivot.RefreshTableCalls);
    }

    [Fact]
    public void Recovered_final_layout_is_verified_as_a_no_op_against_the_original_filter_snapshot()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        PivotSemanticLayoutPlan plan = fixture.Plan(before);
        PivotSemanticPreparedPlacement initial = fixture.Prepare(before);
        initial.Apply();
        int mutations = fixture.Pivot.MutationCalls;

        PivotSemanticPreparedPlacement recovered = fixture.Gateway.PrepareRecoveredFinal(
            fixture.Bind(),
            plan,
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["set_1"] = fixture.Set.Name
            },
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["measure_1"] = fixture.Variance.Name
            },
            before);
        recovered.Apply();

        Assert.Equal(mutations, fixture.Pivot.MutationCalls);
        recovered.VerifyDesired();
    }

    [Fact]
    public void Prepare_rejects_layout_drift_and_verify_rejects_filter_drift()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        fixture.Region.AxisField.Caption = "Changed region";

        InvalidOperationException layoutFailure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Prepare(before));

        Assert.Contains("changed", layoutFailure.Message, StringComparison.OrdinalIgnoreCase);

        fixture.Region.AxisField.Caption = "Region";
        before = fixture.Capture();
        PivotSemanticPreparedPlacement prepared = fixture.Prepare(before);
        fixture.Filter.AxisField.CurrentPageName = "South";
        fixture.Filter.AxisField.PivotItemItems.Single(item => item.Name == "North").Visible = false;
        fixture.Filter.AxisField.PivotItemItems.Single(item => item.Name == "South").Visible = true;

        InvalidOperationException filterFailure = Assert.Throws<InvalidOperationException>(() =>
            prepared.VerifyDesired());

        Assert.Contains("Filters", filterFailure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, fixture.Pivot.MutationCalls);
    }

    [Fact]
    public void Capture_fails_closed_when_filter_state_read_is_unavailable()
    {
        var fixture = new HostFixture();
        fixture.Filter.AxisField.ThrowOnVisibleItemsListRead = true;

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            fixture.Capture());

        Assert.Contains("VisibleItemsList", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, fixture.Pivot.MutationCalls);
    }

    [Fact]
    public void Capture_uses_OLAP_CubeField_flags_and_detects_IncludeNewItems_drift()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        PivotSemanticPreparedPlacement prepared = fixture.Prepare(before);
        fixture.Filter.IncludeNewItemsInFilter = false;

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() =>
            prepared.VerifyDesired());

        Assert.Contains("Filters", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, fixture.Pivot.MutationCalls);
    }

    [Fact]
    public void Partial_apply_restores_exact_prior_placements_without_refresh_or_deletion()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        PivotSemanticPreparedPlacement prepared = fixture.Prepare(before);
        fixture.Pivot.FailNextPositionForUniqueName = fixture.Set.Name;

        PivotSemanticPlacementException failure = Assert.Throws<PivotSemanticPlacementException>(() =>
            prepared.Apply());

        Assert.True(failure.RollbackCompleted);
        PivotSemanticLayoutSnapshot restored = fixture.Capture();
        Assert.Equal(before.LayoutFingerprint, restored.LayoutFingerprint);
        Assert.True(fixture.Pivot.MutationCalls > 3);
        Assert.Equal(0, fixture.Pivot.RefreshTableCalls);
        Assert.Equal(0, fixture.Pivot.CubeFieldDeleteCalls);
        Assert.Equal(0, fixture.Pivot.ModelMeasureDeleteCalls);
        Assert.Equal(0, fixture.Pivot.CalculatedMemberDeleteCalls);
        Assert.Equal(0, fixture.Pivot.GetMeasureCalls);
    }

    [Fact]
    public void Rejects_multiple_visible_levels_from_one_hierarchy_CubeField()
    {
        var fixture = new HostFixture();
        fixture.Pivot.ExtraAxisFields.Add(new FakeAxisField(fixture.Pivot, fixture.Region)
        {
            Caption = "Region duplicate"
        });

        NotSupportedException failure = Assert.Throws<NotSupportedException>(() =>
            fixture.Capture());

        Assert.Contains("hierarchy", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Complete_plan_uses_one_global_256_placement_bound()
    {
        var fixture = new HostFixture();
        PivotSemanticLayoutSnapshot before = fixture.Capture();
        var values = Enumerable.Range(1, 255)
            .Select(index => new PivotSemanticValuePlacement(
                index,
                "measure_" + index))
            .ToArray();
        var measures = values.ToDictionary(
            item => item.DefinitionId!,
            item => "[Measures].[" + item.DefinitionId + "]",
            StringComparer.Ordinal);
        var plan = new PivotSemanticLayoutPlan(
            new[]
            {
                new PivotSemanticAxisPlacement(
                    1,
                    fixture.AxisIdentity(before.Rows[0]))
            },
            new[]
            {
                new PivotSemanticAxisPlacement(
                    1,
                    fixture.AxisIdentity(before.Columns[0]))
            },
            values,
            PivotValuesAxis.Rows,
            valuesPosition: 2);

        NotSupportedException failure = Assert.Throws<NotSupportedException>(() =>
            fixture.Gateway.Prepare(
                fixture.Bind(),
                plan,
                new Dictionary<string, string>(),
                measures,
                before));

        Assert.Contains("bounded", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class HostFixture
    {
        public HostFixture(bool valuesOnRows = false)
        {
            Connection = new FakeConnection(7);
            Model = new FakeModel(Connection);
            Workbook = new FakeWorkbook(Model);
            Sheet = new FakeWorksheet("Sheet1", Workbook);
            Pivot = new FakePivot("PivotTable1", Sheet, Connection);
            Sheet.Pivots.Add(Pivot);
            Region = Pivot.AddCube("[Region]", "Region", 1, 1, 1);
            Department = Pivot.AddCube("[Department]", "Department", 1, 1,
                valuesOnRows ? 3 : 2);
            Month = Pivot.AddCube("[Month]", "Month", 1, 2, 1);
            Set = Pivot.AddCube("[PivotTablePlus_Set]", "Management Rows", 3, 0, 0);
            Cost = Pivot.AddCube("[Measures].[Cost]", "Cost", 2, 0, 0);
            Revenue = Pivot.AddCube("[Measures].[Revenue]", "Revenue", 2, 0, 0);
            Variance = Pivot.AddCube("[Measures].[Variance]", "Variance", 2, 0, 0);
            Pivot.AddValue(Cost, "Sum of Cost", "#,##0", 1);
            Pivot.AddValue(Revenue, "Sum of Revenue", "$#,##0", 2);
            Pivot.DataPivotField.SetInitial(valuesOnRows ? 1 : 2, 2);
            Filter = Pivot.AddCube("[Category]", "Category", 1, 3, 1);
            Filter.AxisField.CurrentPageName = "North";
            Filter.AxisField.VisibleItemsList = new[] { "North" };
            Filter.AxisField.HiddenItemsList = new[] { "South" };
            Filter.AllItemsVisible = false;
            Filter.IncludeNewItemsInFilter = true;
            Filter.AxisField.PivotItemItems.Add(new FakePivotItem("North", "North", true));
            Filter.AxisField.PivotItemItems.Add(new FakePivotItem("South", "South", false));
            Gateway = new LateBoundPivotSemanticLayoutGateway();
            Pivot.MutationCalls = 0;
        }

        public FakeConnection Connection { get; }
        public FakeModel Model { get; }
        public FakeWorkbook Workbook { get; }
        public FakeWorksheet Sheet { get; }
        public FakePivot Pivot { get; }
        public FakeCubeField Region { get; }
        public FakeCubeField Department { get; }
        public FakeCubeField Month { get; }
        public FakeCubeField Set { get; }
        public FakeCubeField Cost { get; }
        public FakeCubeField Revenue { get; }
        public FakeCubeField Variance { get; }
        public FakeCubeField Filter { get; }
        public LateBoundPivotSemanticLayoutGateway Gateway { get; }

        public BoundPivotSemanticLayoutTarget Bind()
        {
            return Gateway.Bind(Workbook, Pivot, Context(PivotSourceKind.DataModel));
        }

        public PivotSemanticLayoutSnapshot Capture()
        {
            return Gateway.Capture(Bind());
        }

        public PivotSemanticPreparedPlacement Prepare(PivotSemanticLayoutSnapshot before)
        {
            return Gateway.Prepare(
                Bind(),
                Plan(before),
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["set_1"] = Set.Name
                },
                new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["measure_1"] = Variance.Name
                },
                before);
        }

        public PivotSemanticLayoutPlan Plan(PivotSemanticLayoutSnapshot before)
        {
            PivotSemanticAxisFieldSnapshot region = before.Rows.Single(field =>
                field.UniqueName == Region.Name);
            PivotSemanticAxisFieldSnapshot month = before.Columns.Single(field =>
                field.UniqueName == Month.Name);
            PivotSemanticValueFieldSnapshot cost = before.Values.Single(field =>
                field.UniqueName == Cost.Name);
            return new PivotSemanticLayoutPlan(
                new PivotSemanticAxisPlacement[]
                {
                    new(1, AxisIdentity(region)),
                    new(2, "set_1")
                },
                new PivotSemanticAxisPlacement[]
                {
                    new(1, AxisIdentity(month))
                },
                new PivotSemanticValuePlacement[]
                {
                    new(1, "measure_1"),
                    new(2, ValueIdentity(cost))
                },
                PivotValuesAxis.Rows,
                valuesPosition: 2);
        }

        public PivotExistingAxisFieldIdentity AxisIdentity(
            PivotSemanticAxisFieldSnapshot field)
        {
            return new PivotExistingAxisFieldIdentity(
                field.UniqueName,
                field.CaptionFingerprint,
                field.Area,
                field.Position);
        }

        public PivotExistingSemanticValueIdentity ValueIdentity(
            PivotSemanticValueFieldSnapshot field)
        {
            return new PivotExistingSemanticValueIdentity(
                field.UniqueName,
                field.CaptionFingerprint,
                field.NumberFormatFingerprint,
                field.Position);
        }

        public PivotTableContext Context(PivotSourceKind kind)
        {
            string workbookId = new StoredWorkbookIdentityResolver().Resolve(Workbook);
            PivotCapability capabilities = PivotCapability.NativeFieldPlacement |
                                           PivotCapability.DataModel |
                                           PivotCapability.ModelMeasures |
                                           PivotCapability.CalculatedMembers |
                                           PivotCapability.NamedSets |
                                           PivotCapability.AsymmetricAxes |
                                           PivotCapability.Refresh;
            return new PivotTableContext(
                new PivotLayoutDefinition(
                    new PivotTargetIdentity(workbookId, Sheet.Name, Pivot.Name),
                    new PivotSourceDescriptor(
                        kind,
                        "ThisWorkbookDataModel",
                        capabilities,
                        "Sales"),
                    fields: Array.Empty<PivotFieldDescriptor>(),
                    placements: Array.Empty<PivotFieldPlacement>(),
                    clearAll: true),
                isConnected: true,
                sourceFieldsComplete: true);
        }
    }

    public sealed class FakeConnection
    {
        public FakeConnection(int type) => Type = type;
        public int Type { get; }
    }

    public sealed class FakeModel
    {
        public FakeModel(FakeConnection connection) => DataModelConnection = connection;
        public FakeConnection DataModelConnection { get; }
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(FakeModel model) => Model = model;
        public FakeModel Model { get; }
        public FakeCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FakeCustomXmlParts
    {
        public FakeCollection<object> SelectByNamespace(string namespaceUri)
        {
            _ = namespaceUri;
            return new FakeCollection<object>(() => Array.Empty<object>());
        }
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(string name, FakeWorkbook parent)
        {
            Name = name;
            Parent = parent;
        }

        public string Name { get; }
        public FakeWorkbook Parent { get; }
        public List<FakePivot> Pivots { get; } = new();
    }

    public sealed class FakePivotCache
    {
        public FakePivotCache(FakeConnection connection)
        {
            WorkbookConnection = connection;
        }

        public bool OLAP { get; set; } = true;
        public FakeConnection WorkbookConnection { get; set; }
    }

    public sealed class FakePivot
    {
        public FakePivot(
            string name,
            FakeWorksheet parent,
            FakeConnection connection)
        {
            Name = name;
            Parent = parent;
            Cache = new FakePivotCache(connection);
            DataPivotField = new FakeDataPivotField(this);
        }

        public string Name { get; }
        public FakeWorksheet Parent { get; }
        public FakePivotCache Cache { get; }
        public List<FakeCubeField> CubeItems { get; } = new();
        public List<FakeDataField> DataItems { get; } = new();
        public List<FakeAxisField> ExtraAxisFields { get; } = new();
        public FakeDataPivotField DataPivotField { get; }
        public int MutationCalls { get; set; }
        public int RefreshTableCalls { get; private set; }
        public int CubeFieldDeleteCalls { get; set; }
        public int ModelMeasureDeleteCalls { get; set; }
        public int CalculatedMemberDeleteCalls { get; set; }
        public int GetMeasureCalls { get; set; }
        public string? FailNextPositionForUniqueName { get; set; }

        public FakePivotCache PivotCache() => Cache;

        public FakeCollection<FakeCubeField> CubeFields =>
            new(() => CubeItems);

        public FakeCollection<object> RowFields => new(() => AxisItems(1));

        public FakeCollection<object> ColumnFields => new(() => AxisItems(2));

        public FakeCollection<FakeAxisField> PageFields => new(() =>
            CubeItems.Where(cube => cube.Orientation == 3)
                .Select(cube => cube.AxisField)
                .OrderBy(field => field.Position));

        public FakeCollection<FakeDataField> DataFields => new(() =>
            DataItems.Where(field => field.Orientation == 4)
                .OrderBy(field => field.Position));

        public FakeCubeField AddCube(
            string name,
            string caption,
            int type,
            int orientation,
            int position)
        {
            var cube = new FakeCubeField(this, name, caption, type);
            CubeItems.Add(cube);
            cube.SetInitial(orientation, position);
            return cube;
        }

        public FakeDataField AddValue(
            FakeCubeField cube,
            string caption,
            string numberFormat,
            int position)
        {
            cube.SetInitial(4, 0);
            var field = new FakeDataField(this, cube, caption, numberFormat, position);
            DataItems.Add(field);
            return field;
        }

        public bool RefreshTable()
        {
            RefreshTableCalls++;
            return true;
        }

        public void RecordMutation()
        {
            MutationCalls++;
        }

        public int NextAxisPosition(int orientation)
        {
            int maximum = CubeItems
                .Where(cube => cube.Orientation == orientation)
                .Select(cube => cube.Position)
                .DefaultIfEmpty(0)
                .Max();
            if (DataPivotField.Orientation == orientation)
            {
                maximum = Math.Max(maximum, DataPivotField.Position);
            }

            return maximum + 1;
        }

        public int NextDataPosition()
        {
            return DataItems.Where(field => field.Orientation == 4)
                       .Select(field => field.Position)
                       .DefaultIfEmpty(0)
                       .Max() + 1;
        }

        private IEnumerable<object> AxisItems(int orientation)
        {
            IEnumerable<object> regular = CubeItems
                .Where(cube => cube.Orientation == orientation && cube.CubeFieldType != 2)
                .Select(cube => (object)cube.AxisField)
                .Concat(ExtraAxisFields.Where(field => field.Orientation == orientation));
            if (DataPivotField.Orientation == orientation && DataFields.Count > 1)
            {
                regular = regular.Concat(new object[] { DataPivotField });
            }

            return regular.OrderBy(item => ((dynamic)item).Position).ToArray();
        }
    }

    public sealed class FakeCubeField
    {
        private readonly FakePivot owner;
        private int orientation;
        private int position;

        public FakeCubeField(
            FakePivot owner,
            string name,
            string caption,
            int type)
        {
            this.owner = owner;
            Name = name;
            Caption = caption;
            CubeFieldType = type;
            AxisField = new FakeAxisField(owner, this) { Caption = caption };
        }

        public string Name { get; }
        public string Caption { get; }
        public int CubeFieldType { get; }
        public FakeAxisField AxisField { get; }
        public bool EnableMultiplePageItems { get; set; }
        public bool AllItemsVisible { get; set; } = true;
        public bool IncludeNewItemsInFilter { get; set; }

        public int Orientation
        {
            get => orientation;
            set
            {
                owner.RecordMutation();
                if (CubeFieldType == 2 && value == 4 &&
                    !owner.DataItems.Any(field =>
                        field.Orientation == 4 && ReferenceEquals(field.CubeField, this)))
                {
                    owner.DataItems.Add(new FakeDataField(
                        owner,
                        this,
                        Caption,
                        "General",
                        owner.NextDataPosition()));
                }

                if (value != 0 && value != orientation && position <= 0)
                {
                    position = value == 4
                        ? owner.NextDataPosition()
                        : owner.NextAxisPosition(value);
                }

                orientation = value;
            }
        }

        public int Position
        {
            get => position;
            set
            {
                owner.RecordMutation();
                if (string.Equals(
                        owner.FailNextPositionForUniqueName,
                        Name,
                        StringComparison.Ordinal))
                {
                    owner.FailNextPositionForUniqueName = null;
                    throw new InvalidOperationException("synthetic position failure");
                }

                position = value;
            }
        }

        public void SetInitial(int initialOrientation, int initialPosition)
        {
            orientation = initialOrientation;
            position = initialPosition;
        }

        public void Delete()
        {
            owner.CubeFieldDeleteCalls++;
        }
    }

    public sealed class FakeAxisField
    {
        private object visibleItemsList = Array.Empty<string>();

        public FakeAxisField(FakePivot owner, FakeCubeField cubeField)
        {
            _ = owner;
            CubeField = cubeField;
            PivotItems = new FakeCollection<FakePivotItem>(() => PivotItemItems);
            PivotFilters = new FakeCollection<object>(() => Array.Empty<object>());
        }

        public FakeCubeField CubeField { get; }
        public string Caption { get; set; } = string.Empty;
        public bool AllItemsVisible =>
            throw new InvalidOperationException(
                "PivotField.AllItemsVisible is unavailable for OLAP.");
        public string CurrentPageName { get; set; } = string.Empty;
        public object CurrentPageList { get; set; } = Array.Empty<string>();

        public object VisibleItemsList
        {
            get
            {
                if (ThrowOnVisibleItemsListRead)
                {
                    throw new InvalidOperationException("VisibleItemsList unavailable");
                }

                return visibleItemsList;
            }
            set => visibleItemsList = value;
        }

        public object HiddenItemsList { get; set; } = Array.Empty<string>();
        public bool ThrowOnVisibleItemsListRead { get; set; }
        public List<FakePivotItem> PivotItemItems { get; } = new();
        public FakeCollection<FakePivotItem> PivotItems { get; }
        public FakeCollection<object> PivotFilters { get; }

        public int Orientation
        {
            get => CubeField.Orientation;
            set => CubeField.Orientation = value;
        }

        public int Position
        {
            get => CubeField.Position;
            set => CubeField.Position = value;
        }
    }

    public sealed class FakeDataField
    {
        private readonly FakePivot owner;
        private int orientation = 4;
        private int position;

        public FakeDataField(
            FakePivot owner,
            FakeCubeField cubeField,
            string caption,
            string numberFormat,
            int position)
        {
            this.owner = owner;
            CubeField = cubeField;
            Caption = caption;
            NumberFormat = numberFormat;
            this.position = position;
        }

        public FakeCubeField CubeField { get; }
        public string Caption { get; set; }
        public string NumberFormat { get; set; }

        public int Orientation
        {
            get => orientation;
            set
            {
                owner.RecordMutation();
                orientation = value;
            }
        }

        public int Position
        {
            get => position;
            set
            {
                owner.RecordMutation();
                position = value;
            }
        }
    }

    public sealed class FakeDataPivotField
    {
        private readonly FakePivot owner;
        private int orientation;
        private int position = 1;

        public FakeDataPivotField(FakePivot owner)
        {
            this.owner = owner;
        }

        public int Orientation
        {
            get => orientation;
            set
            {
                owner.RecordMutation();
                orientation = value;
            }
        }

        public int Position
        {
            get => position;
            set
            {
                owner.RecordMutation();
                position = value;
            }
        }

        public void SetInitial(int initialOrientation, int initialPosition)
        {
            orientation = initialOrientation;
            position = initialPosition;
        }
    }

    public sealed class FakePivotItem
    {
        public FakePivotItem(string name, string caption, bool visible)
        {
            Name = name;
            Caption = caption;
            Visible = visible;
        }

        public string Name { get; }
        public string Caption { get; }
        public bool Visible { get; set; }
    }

    public sealed class FakeCollection<T>
    {
        private readonly Func<IEnumerable<T>> values;

        public FakeCollection(Func<IEnumerable<T>> values)
        {
            this.values = values;
        }

        public IReadOnlyList<T> Items => values().ToArray();
        public int Count => Items.Count;
        public T Item(int index) => Items[index - 1];
    }
}
