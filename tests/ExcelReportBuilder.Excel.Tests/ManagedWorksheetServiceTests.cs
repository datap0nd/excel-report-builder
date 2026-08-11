using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ManagedWorksheetServiceTests
{
    [Fact]
    public void Reuses_owned_sheet_even_when_an_unmanaged_sheet_holds_the_preferred_name()
    {
        var identity = DraftIdentity();
        var preferredName = ManagedName.Worksheet("Report draft", identity.ObjectId);
        var workbook = new FakeWorkbook();
        var unmanaged = workbook.Worksheets.AddExisting(preferredName);
        var owned = workbook.Worksheets.AddExisting(preferredName + " 2");
        new ManagedOwnershipGuard().MarkOwned(owned, identity);
        var additionsBeforeLookup = workbook.Worksheets.AddCount;

        dynamic actual = new ManagedWorksheetService().GetOrCreateDraft(
            workbook,
            identity,
            "Report draft");

        Assert.Same(owned, actual);
        Assert.Equal(additionsBeforeLookup, workbook.Worksheets.AddCount);
        Assert.Equal(preferredName, unmanaged.Name);
        Assert.Empty(unmanaged.Log);
    }

    [Fact]
    public void Allocates_a_numbered_sheet_without_changing_an_unmanaged_name_collision()
    {
        var identity = DraftIdentity();
        var preferredName = ManagedName.Worksheet("Report draft", identity.ObjectId);
        var workbook = new FakeWorkbook();
        var unmanaged = workbook.Worksheets.AddExisting(preferredName);

        dynamic actual = new ManagedWorksheetService().GetOrCreateDraft(
            workbook,
            identity,
            "Report draft");

        Assert.Equal(preferredName + " 2", actual.Name);
        Assert.Equal(preferredName, unmanaged.Name);
        Assert.Empty(unmanaged.Log);
        Assert.True(new ManagedOwnershipGuard().IsOwned(actual, identity));
    }

    [Fact]
    public void Rejects_duplicate_worksheets_with_the_same_ownership_identity()
    {
        var identity = DraftIdentity();
        var workbook = new FakeWorkbook();
        var first = workbook.Worksheets.AddExisting("Managed 1");
        var second = workbook.Worksheets.AddExisting("Managed 2");
        var guard = new ManagedOwnershipGuard();
        guard.MarkOwned(first, identity);
        guard.MarkOwned(second, identity);

        var error = Assert.Throws<InvalidOperationException>(() =>
            new ManagedWorksheetService().GetOrCreateDraft(workbook, identity, "Report draft"));

        Assert.Contains("same managed-object ownership marker", error.Message);
        Assert.Equal(2, workbook.Worksheets.AddCount);
    }

    [Fact]
    public void Clears_structured_objects_before_clearing_cells_on_an_owned_sheet()
    {
        var identity = DraftIdentity();
        var worksheet = new FakeWorksheet("Managed");
        worksheet.AddPivot("Pivot 1");
        worksheet.AddPivot("Pivot 2");
        worksheet.AddListObject("Table 1", isRefreshing: true);
        worksheet.AddListObject("Table 2", isRefreshing: false);
        new ManagedOwnershipGuard().MarkOwned(worksheet, identity);

        new ManagedWorksheetService().ClearOwned(worksheet, identity);

        Assert.Equal(0, worksheet.PivotTables().Count);
        Assert.Equal(0, worksheet.ListObjects.Count);
        Assert.Equal("pivot:Pivot 2:clear", worksheet.Log[0]);
        Assert.Equal("pivot:Pivot 1:clear", worksheet.Log[1]);
        Assert.Contains("query:Table 1:cancel", worksheet.Log);
        Assert.Equal("cells:clear", worksheet.Log[^1]);
        Assert.All(
            worksheet.Log.Take(worksheet.Log.Count - 1),
            entry => Assert.NotEqual("cells:clear", entry));
    }

    [Fact]
    public void Refuses_to_clear_an_unmanaged_sheet_before_touching_its_contents()
    {
        var worksheet = new FakeWorksheet("Unmanaged");
        worksheet.AddPivot("Existing Pivot");
        worksheet.AddListObject("Existing Table", isRefreshing: true);

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedWorksheetService().ClearOwned(worksheet, DraftIdentity()));

        Assert.Empty(worksheet.Log);
        Assert.Equal(1, worksheet.PivotTables().Count);
        Assert.Equal(1, worksheet.ListObjects.Count);
    }

    private static ManagedObjectIdentity DraftIdentity()
    {
        return new ManagedObjectIdentity("report", "draft-object", ManagedObjectKind.DraftWorksheet);
    }

    public sealed class FakeWorkbook
    {
        public FakeWorksheetCollection Worksheets { get; } = new();
    }

    public sealed class FakeWorksheetCollection
    {
        private readonly List<FakeWorksheet> worksheets = new();

        public int Count => worksheets.Count;

        public int AddCount { get; private set; }

        public FakeWorksheet Add()
        {
            AddCount++;
            var worksheet = new FakeWorksheet("Sheet" + AddCount);
            worksheets.Add(worksheet);
            return worksheet;
        }

        public FakeWorksheet AddExisting(string name)
        {
            return AddWithName(name);
        }

        public FakeWorksheet Item(int index)
        {
            return worksheets[index - 1];
        }

        public FakeWorksheet Item(string name)
        {
            return worksheets.Single(worksheet =>
                string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private FakeWorksheet AddWithName(string name)
        {
            var worksheet = Add();
            worksheet.Name = name;
            return worksheet;
        }
    }

    public sealed class FakeWorksheet
    {
        private readonly FakePivotTableCollection pivotTables;

        public FakeWorksheet(string name)
        {
            Name = name;
            Log = new List<string>();
            pivotTables = new FakePivotTableCollection(Log);
            ListObjects = new FakeListObjectCollection(Log);
            Cells = new FakeCells(Log);
        }

        public string Name { get; set; }

        public int Visible { get; set; }

        public List<string> Log { get; }

        public FakeCustomProperties CustomProperties { get; } = new();

        public FakeListObjectCollection ListObjects { get; }

        public FakeCells Cells { get; }

        public FakePivotTableCollection PivotTables()
        {
            return pivotTables;
        }

        public void AddPivot(string name)
        {
            pivotTables.Add(name);
        }

        public void AddListObject(string name, bool isRefreshing)
        {
            ListObjects.Add(name, isRefreshing);
        }
    }

    public sealed class FakePivotTableCollection
    {
        private readonly List<FakePivotTable> values = new();
        private readonly List<string> log;

        public FakePivotTableCollection(List<string> log)
        {
            this.log = log;
        }

        public int Count => values.Count;

        public FakePivotTable Item(int index)
        {
            return values[index - 1];
        }

        public void Add(string name)
        {
            FakePivotTable? pivot = null;
            pivot = new FakePivotTable(name, log, () => values.Remove(pivot!));
            values.Add(pivot);
        }
    }

    public sealed class FakePivotTable
    {
        public FakePivotTable(string name, List<string> log, Action remove)
        {
            TableRange2 = new FakePivotRange(name, log, remove);
        }

        public FakePivotRange TableRange2 { get; }
    }

    public sealed class FakePivotRange
    {
        private readonly string name;
        private readonly List<string> log;
        private readonly Action remove;

        public FakePivotRange(string name, List<string> log, Action remove)
        {
            this.name = name;
            this.log = log;
            this.remove = remove;
        }

        public void Clear()
        {
            log.Add("pivot:" + name + ":clear");
            remove();
        }
    }

    public sealed class FakeListObjectCollection
    {
        private readonly List<FakeListObject> values = new();
        private readonly List<string> log;

        public FakeListObjectCollection(List<string> log)
        {
            this.log = log;
        }

        public int Count => values.Count;

        public FakeListObject Item(int index)
        {
            return values[index - 1];
        }

        public void Add(string name, bool isRefreshing)
        {
            FakeListObject? table = null;
            table = new FakeListObject(name, isRefreshing, log, () => values.Remove(table!));
            values.Add(table);
        }
    }

    public sealed class FakeListObject
    {
        private readonly string name;
        private readonly List<string> log;
        private readonly Action remove;

        public FakeListObject(string name, bool isRefreshing, List<string> log, Action remove)
        {
            this.name = name;
            this.log = log;
            this.remove = remove;
            QueryTable = new FakeQueryTable(name, isRefreshing, log);
        }

        public FakeQueryTable QueryTable { get; }

        public void Delete()
        {
            log.Add("table:" + name + ":delete");
            remove();
        }
    }

    public sealed class FakeQueryTable
    {
        private readonly string name;
        private readonly List<string> log;

        public FakeQueryTable(string name, bool isRefreshing, List<string> log)
        {
            this.name = name;
            Refreshing = isRefreshing;
            this.log = log;
        }

        public bool Refreshing { get; }

        public void CancelRefresh()
        {
            log.Add("query:" + name + ":cancel");
        }
    }

    public sealed class FakeCells
    {
        private readonly List<string> log;

        public FakeCells(List<string> log)
        {
            this.log = log;
        }

        public void Clear()
        {
            log.Add("cells:clear");
        }
    }

    public sealed class FakeCustomProperties
    {
        private readonly Dictionary<string, FakeCustomProperty> values = new(StringComparer.Ordinal);

        public FakeCustomProperty Item(string name)
        {
            return values.TryGetValue(name, out var value)
                ? value
                : throw new InvalidOperationException("Property not found.");
        }

        public void Add(string name, string value)
        {
            values[name] = new FakeCustomProperty { Value = value };
        }
    }

    public sealed class FakeCustomProperty
    {
        public string Value { get; set; } = string.Empty;
    }
}
