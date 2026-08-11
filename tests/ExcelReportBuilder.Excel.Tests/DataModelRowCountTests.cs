using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class DataModelRowCountTests
{
    [Fact]
    public void Reads_record_count_only_from_the_exact_managed_source_connection()
    {
        var workbook = new FakeWorkbook(
            new FakeModelTable("Unrelated connection", 999),
            new FakeModelTable("Managed connection", 42));

        long result = ExcelReportExecutor.CountDataModelRows(
            workbook,
            "Managed connection");

        Assert.Equal(42, result);
    }

    [Fact]
    public void Rejects_ambiguous_or_missing_managed_model_tables()
    {
        var duplicate = new FakeWorkbook(
            new FakeModelTable("Managed connection", 20),
            new FakeModelTable("managed connection", 20));
        var missing = new FakeWorkbook(new FakeModelTable("Other", 20));

        Assert.Throws<InvalidOperationException>(() =>
            ExcelReportExecutor.CountDataModelRows(duplicate, "Managed connection"));
        Assert.Throws<InvalidOperationException>(() =>
            ExcelReportExecutor.CountDataModelRows(missing, "Managed connection"));
    }

    public sealed class FakeWorkbook
    {
        public FakeWorkbook(params FakeModelTable[] tables)
        {
            Model = new FakeModel(tables);
        }

        public FakeModel Model { get; }
    }

    public sealed class FakeModel
    {
        public FakeModel(FakeModelTable[] tables)
        {
            ModelTables = new FakeModelTables(tables);
        }

        public FakeModelTables ModelTables { get; }
    }

    public sealed class FakeModelTables
    {
        private readonly IReadOnlyList<FakeModelTable> _tables;

        public FakeModelTables(IReadOnlyList<FakeModelTable> tables)
        {
            _tables = tables;
        }

        public int Count => _tables.Count;

        public FakeModelTable Item(int index)
        {
            return _tables[index - 1];
        }
    }

    public sealed class FakeModelTable
    {
        public FakeModelTable(string connectionName, long recordCount)
        {
            SourceWorkbookConnection = new FakeConnection(connectionName);
            RecordCount = recordCount;
        }

        public FakeConnection SourceWorkbookConnection { get; }

        public long RecordCount { get; }
    }

    public sealed class FakeConnection
    {
        public FakeConnection(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }
}
