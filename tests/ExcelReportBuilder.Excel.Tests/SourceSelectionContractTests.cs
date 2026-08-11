using ExcelReportBuilder.Excel.Source;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class SourceSelectionContractTests
{
    [Fact]
    public void Exact_table_selection_profiles_only_table_data_rows_and_reuses_table_name()
    {
        FakeRange selection = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$C$4",
            new object?[] { "Region", "Amount", "Period" },
            new object?[] { "North", 10m, "Jan-26" },
            new object?[] { "South", 20m, "Feb-26" },
            new object?[] { "Grand Total", 30m, null });
        FakeRange tableRange = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$C$4",
            new object?[] { "Region", "Amount", "Period" },
            new object?[] { "North", 10m, "Jan-26" },
            new object?[] { "South", 20m, "Feb-26" },
            new object?[] { "Grand Total", 30m, null });
        selection.ListObject = new FakeListObject(
            "SourceTable",
            tableRange,
            dataRowCount: 2,
            columnCount: 3);

        SourceSelectionSnapshot snapshot = new SourceSelectionInspector().Inspect(
            new FakeExcelApplication(selection));
        string workbookObjectName = new ManagedSourceNameService().EnsureWorkbookObject(
            new object(),
            selection,
            "report",
            "source");

        Assert.Equal("SourceTable", snapshot.WorkbookObjectName);
        Assert.Equal("SourceTable", workbookObjectName);
        Assert.Equal(2, snapshot.RowCount);
        Assert.Equal(3, snapshot.ColumnCount);
        Assert.Equal(new[] { "Region", "Amount", "Period" }, snapshot.Headers);
        Assert.Equal(2, snapshot.SampleRows.Count);
        Assert.Equal("South", snapshot.SampleRows[1][0]);
    }

    [Fact]
    public void Table_contained_subrange_is_rejected_before_profiling_or_naming()
    {
        FakeRange tableRange = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$C$3",
            new object?[] { "Region", "Amount", "Period" },
            new object?[] { "North", 10m, "Jan-26" },
            new object?[] { "South", 20m, "Feb-26" });
        var table = new FakeListObject(
            "SourceTable",
            tableRange,
            dataRowCount: 2,
            columnCount: 3);
        FakeRange partialSelection = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$B$3",
            new object?[] { "Region", "Amount" },
            new object?[] { "North", 10m },
            new object?[] { "South", 20m });
        partialSelection.ListObject = table;

        InvalidOperationException inspectException = Assert.Throws<InvalidOperationException>(() =>
            new SourceSelectionInspector().Inspect(new FakeExcelApplication(partialSelection)));
        InvalidOperationException namingException = Assert.Throws<InvalidOperationException>(() =>
            new ManagedSourceNameService().EnsureWorkbookObject(
                new object(),
                partialSelection,
                "report",
                "source"));

        Assert.Contains("not the complete table", inspectException.Message, StringComparison.Ordinal);
        Assert.Contains("not the complete table", namingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Exact_table_with_hidden_headers_is_rejected_before_first_data_row_can_be_misread()
    {
        FakeRange tableRange = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$B$3",
            new object?[] { "North", 10m },
            new object?[] { "South", 20m },
            new object?[] { "West", 30m });
        var table = new FakeListObject(
            "SourceTable",
            tableRange,
            dataRowCount: 3,
            columnCount: 2)
        {
            ShowHeaders = false
        };
        tableRange.ListObject = table;

        InvalidOperationException inspectException = Assert.Throws<InvalidOperationException>(() =>
            new SourceSelectionInspector().Inspect(new FakeExcelApplication(tableRange)));
        InvalidOperationException namingException = Assert.Throws<InvalidOperationException>(() =>
            new ManagedSourceNameService().EnsureWorkbookObject(
                new object(),
                tableRange,
                "report",
                "source"));

        Assert.Contains("must show its header row", inspectException.Message, StringComparison.Ordinal);
        Assert.Contains("must show its header row", namingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Ordinary_range_profiles_the_same_header_and_data_extent_that_will_be_named()
    {
        FakeRange selection = FakeRange.Create(
            "[Synthetic.xlsx]Data!$D$5:$E$7",
            new object?[] { " Region ", "Amount" },
            new object?[] { "North", 10m },
            new object?[] { "South", 20m });

        SourceSelectionSnapshot snapshot = new SourceSelectionInspector().Inspect(
            new FakeExcelApplication(selection));

        Assert.Equal(2, snapshot.RowCount);
        Assert.Equal(2, snapshot.ColumnCount);
        Assert.Equal(new[] { " Region ", "Amount" }, snapshot.Headers);
        Assert.Equal(2, snapshot.SampleRows.Count);
        Assert.Contains("[Synthetic.xlsx]Data!$D$5:$E$7", snapshot.WorkbookObjectName, StringComparison.Ordinal);
    }

    [Fact]
    public void Existing_managed_range_name_is_repointed_through_writable_refers_to_formula()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedSourceNameService();
        FakeRange first = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$B$2",
            new object?[] { "Region", "Amount" },
            new object?[] { "North", 10m });
        FakeRange second = FakeRange.Create(
            "[Synthetic.xlsx]Data!$D$5:$E$6",
            new object?[] { "Region", "Amount" },
            new object?[] { "South", 20m });

        string name = service.EnsureWorkbookObject(workbook, first, "report", "source");
        Assert.Equal(
            "=[Synthetic.xlsx]Data!$A$1:$B$2",
            workbook.Names.Item(name).RefersTo);

        service.EnsureWorkbookObject(workbook, second, "report", "source");

        FakeName managedName = workbook.Names.Item(name);
        Assert.Equal("=[Synthetic.xlsx]Data!$D$5:$E$6", managedName.RefersTo);
    }

    [Fact]
    public void Same_named_replacement_workbook_name_is_rejected_without_refers_to_assignment()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedSourceNameService();
        FakeRange first = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$B$2",
            new object?[] { "Region", "Amount" },
            new object?[] { "North", 10m });
        FakeRange second = FakeRange.Create(
            "[Synthetic.xlsx]Data!$D$5:$E$6",
            new object?[] { "Region", "Amount" },
            new object?[] { "South", 20m });

        string name = service.EnsureWorkbookObject(workbook, first, "report", "source");
        workbook.Names.Remove(name);
        FakeName replacement = workbook.Names.Add(
            name,
            "='[Other.xlsx]Sheet1'!$A$1:$B$2");

        Assert.Throws<InvalidOperationException>(() =>
            service.EnsureWorkbookObject(workbook, second, "report", "source"));

        Assert.Equal(0, replacement.RefersToAssignmentCount);
        Assert.Equal("='[Other.xlsx]Sheet1'!$A$1:$B$2", replacement.RefersTo);
    }

    [Fact]
    public void Registry_only_source_name_claim_is_removed_before_recreation()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedSourceNameService();
        FakeRange first = FakeRange.Create(
            "[Synthetic.xlsx]Data!$A$1:$B$2",
            new object?[] { "Region", "Amount" },
            new object?[] { "North", 10m });
        FakeRange second = FakeRange.Create(
            "[Synthetic.xlsx]Data!$D$5:$E$6",
            new object?[] { "Region", "Amount" },
            new object?[] { "South", 20m });

        string name = service.EnsureWorkbookObject(workbook, first, "report", "source");
        workbook.Names.Remove(name);

        service.EnsureWorkbookObject(workbook, second, "report", "source");

        Assert.Equal("=[Synthetic.xlsx]Data!$D$5:$E$6", workbook.Names.Item(name).RefersTo);
        ManagedObjectRecord record = Assert.Single(new WorkbookOwnershipRegistry().Load(workbook));
        Assert.False(string.IsNullOrWhiteSpace(record.SourceContract));
    }

    public sealed class FakeExcelApplication
    {
        public FakeExcelApplication(FakeRange selection)
        {
            Selection = selection;
        }

        public object Selection { get; }
    }

    public sealed class FakeRange
    {
        private readonly Array values;

        private FakeRange(string externalAddress, Array values)
        {
            Address = new FakeAddress(externalAddress);
            this.values = values;
            Areas = new FakeCount(1);
            Rows = new FakeCount(values.GetLength(0));
            Columns = new FakeCount(values.GetLength(1));
            Worksheet = new FakeWorksheet("Data");
            Resize = new FakeResize(this);
        }

        public FakeCount Areas { get; }

        public FakeCount Rows { get; }

        public FakeCount Columns { get; }

        public FakeResize Resize { get; }

        public FakeAddress Address { get; }

        public FakeWorksheet Worksheet { get; }

        public FakeListObject? ListObject { get; set; }

        public object Value2 => values;

        public static FakeRange Create(string externalAddress, params object?[][] rows)
        {
            if (rows.Length == 0)
            {
                throw new ArgumentException("At least one row is required.", nameof(rows));
            }

            int columnCount = rows[0].Length;
            Array values = Array.CreateInstance(
                typeof(object),
                new[] { rows.Length, columnCount },
                new[] { 1, 1 });
            for (int row = 0; row < rows.Length; row++)
            {
                if (rows[row].Length != columnCount)
                {
                    throw new ArgumentException("Fake rows must be rectangular.", nameof(rows));
                }

                for (int column = 0; column < columnCount; column++)
                {
                    values.SetValue(rows[row][column], row + 1, column + 1);
                }
            }

            return new FakeRange(externalAddress, values);
        }

        public FakeRange Slice(int rowCount, int columnCount)
        {
            Array slice = Array.CreateInstance(
                typeof(object),
                new[] { rowCount, columnCount },
                new[] { 1, 1 });
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    slice.SetValue(values.GetValue(row, column), row, column);
                }
            }

            return new FakeRange(Address.ExternalAddress, slice);
        }
    }

    public sealed class FakeResize
    {
        private readonly FakeRange range;

        public FakeResize(FakeRange range)
        {
            this.range = range;
        }

        public FakeRange this[int rowCount, int columnCount] => range.Slice(rowCount, columnCount);
    }

    public sealed class FakeAddress
    {
        public FakeAddress(string externalAddress)
        {
            ExternalAddress = externalAddress;
        }

        public string ExternalAddress { get; }

        public string this[bool rowAbsolute, bool columnAbsolute, int referenceStyle, bool external] =>
            ExternalAddress;
    }

    public sealed class FakeListObject
    {
        public FakeListObject(
            string name,
            FakeRange range,
            int dataRowCount,
            int columnCount)
        {
            Name = name;
            Range = range;
            ListRows = new FakeCount(dataRowCount);
            ListColumns = new FakeCount(columnCount);
        }

        public string Name { get; }

        public FakeRange Range { get; }

        public FakeCount ListRows { get; }

        public FakeCount ListColumns { get; }

        public bool ShowHeaders { get; set; } = true;
    }

    public sealed class FakeWorkbook
    {
        public FakeNames Names { get; } = new();

        public WorkbookSpecStoreTests.FakeCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FakeNames
    {
        private readonly Dictionary<string, FakeName> names = new(StringComparer.Ordinal);

        public FakeName Item(string name)
        {
            return names.TryGetValue(name, out FakeName? value)
                ? value
                : throw new InvalidOperationException("Name not found.");
        }

        public FakeName Add(string name, string refersTo)
        {
            var created = new FakeName(refersTo);
            names.Add(name, created);
            return created;
        }

        public void Remove(string name)
        {
            names.Remove(name);
        }
    }

    public sealed class FakeName
    {
        private string refersTo;

        public FakeName(string refersTo)
        {
            this.refersTo = refersTo;
        }

        public string RefersTo
        {
            get => refersTo;
            set
            {
                RefersToAssignmentCount++;
                refersTo = value;
            }
        }

        public int RefersToAssignmentCount { get; private set; }
    }

    public sealed class FakeCount
    {
        public FakeCount(int count)
        {
            Count = count;
        }

        public int Count { get; }
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }
}
