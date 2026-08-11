using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class CanonicalDataAuditTests
{
    private const string CanonicalFormula =
        "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

    [Fact]
    public void Audit_query_counts_every_canonical_row_and_sums_only_additive_fields()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[]
            {
                Aggregate("net_value", "Net value", AggregateFunction.Sum),
                Aggregate("units", "Units", AggregateFunction.Sum),
                Aggregate("average_value", "Net value", AggregateFunction.Average)
            });

        Assert.Contains("Table.RowCount(_)", plan.Formula, StringComparison.Ordinal);
        Assert.Contains("List.Sum(Values)", plan.Formula, StringComparison.Ordinal);
        Assert.Contains("Table.Column(_, \"Net value\")", plan.Formula, StringComparison.Ordinal);
        Assert.Contains("Table.Column(_, \"Units\")", plan.Formula, StringComparison.Ordinal);
        Assert.DoesNotContain("Table.First", plan.Formula, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Pivot", plan.Formula, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, plan.Fields.Count);
        Assert.Equal("net_value", Assert.Single(plan.Fields[0].MeasureIds));
        Assert.Equal("units", Assert.Single(plan.Fields[1].MeasureIds));
    }

    [Fact]
    public void Audit_query_reuses_one_total_for_measures_over_the_same_field()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[]
            {
                Aggregate("amount", "Amount", AggregateFunction.Sum),
                Aggregate("amount_copy", "amount", AggregateFunction.Sum)
            });

        var field = Assert.Single(plan.Fields);
        Assert.Equal(new[] { "amount", "amount_copy" }, field.MeasureIds);
        Assert.Equal(1, CountOccurrences(plan.Formula, "List.Sum(Values)"));
    }

    [Fact]
    public void Audit_query_escapes_quotes_in_canonical_field_names()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[] { Aggregate("quoted", "Value \"net\"", AggregateFunction.Sum) });

        Assert.Contains(
            "Table.Column(_, \"Value \"\"net\"\"\")",
            plan.Formula,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Audit_query_rejects_control_characters_in_fields()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            CanonicalAuditQueryCompiler.Compile(
                CanonicalFormula,
                new[] { Aggregate("unsafe", "Value\nInjected", AggregateFunction.Sum) }));

        Assert.Contains("control characters", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Audit_formula_remains_current_workbook_only()
    {
        var policy = new RestrictedQueryFormulaPolicy();
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[] { Aggregate("amount", "Amount", AggregateFunction.Sum) });

        policy.DemandWorkbookOnly(plan.Formula);
        Assert.DoesNotContain("File.Contents", plan.Formula, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Web.Contents", plan.Formula, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Sql.Database", plan.Formula, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Empty_canonical_results_still_produce_one_zero_audit_row()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[] { Aggregate("amount", "Amount", AggregateFunction.Sum) });

        Assert.Contains("EmptyResult = #table", plan.Formula, StringComparison.Ordinal);
        Assert.Contains("{{0, 0}}", plan.Formula, StringComparison.Ordinal);
        Assert.Contains(
            "if Table.IsEmpty(WithoutKey) then EmptyResult else WithoutKey",
            plan.Formula,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Row_count_only_audit_does_not_depend_on_zero_column_table_behavior()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[] { Aggregate("average", "Amount", AggregateFunction.Average) });

        Assert.Empty(plan.Fields);
        Assert.Contains("{{Table.RowCount(Canonical)}}", plan.Formula, StringComparison.Ordinal);
        Assert.DoesNotContain("Table.SelectColumns", plan.Formula, StringComparison.Ordinal);
        Assert.DoesNotContain("Table.Group", plan.Formula, StringComparison.Ordinal);
    }

    [Fact]
    public void Audit_result_maps_independent_field_totals_back_to_every_measure()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[]
            {
                Aggregate("amount", "Amount", AggregateFunction.Sum),
                Aggregate("amount_copy", "amount", AggregateFunction.Sum)
            });
        var table = new AuditResultTable(new Dictionary<string, object?>
        {
            [CanonicalAuditQueryPlan.RowCountColumnName] = 12m,
            [plan.Fields[0].ResultColumnName] = 45.5m
        });

        var result = CanonicalDataAuditor.ReadResult(table, plan);

        Assert.Equal(12, result.ActualRows);
        Assert.Equal(45.5m, result.Totals["amount"]);
        Assert.Equal(45.5m, result.Totals["amount_copy"]);
    }

    [Fact]
    public void Audit_result_fails_closed_on_a_non_numeric_total()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(
            CanonicalFormula,
            new[] { Aggregate("amount", "Amount", AggregateFunction.Sum) });
        var table = new AuditResultTable(new Dictionary<string, object?>
        {
            [CanonicalAuditQueryPlan.RowCountColumnName] = 12m,
            [plan.Fields[0].ResultColumnName] = "#VALUE!"
        });

        var exception = Assert.Throws<InvalidOperationException>(() =>
            CanonicalDataAuditor.ReadResult(table, plan));

        Assert.Contains("non-numeric", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Audit_result_fails_closed_unless_exactly_one_summary_row_is_returned()
    {
        var plan = CanonicalAuditQueryCompiler.Compile(CanonicalFormula, Array.Empty<MeasureDefinition>());
        var table = new AuditResultTable(
            new Dictionary<string, object?>
            {
                [CanonicalAuditQueryPlan.RowCountColumnName] = 0m
            },
            rowCount: 0);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            CanonicalDataAuditor.ReadResult(table, plan));

        Assert.Contains("exactly one", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Same_named_replacement_audit_connection_is_not_deleted()
    {
        var workbook = new CanonicalDataLoaderIdempotencyTests.FakeWorkbook();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_connection",
            ManagedObjectKind.DataModelConnection);
        const string connectionName = "Connection - ERB audit";
        string queryName = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_query",
            ManagedObjectKind.CanonicalQuery).ExcelName;
        var replacement = workbook.Connections.AddExisting(
            connectionName,
            "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\example\\external.xlsx",
            "SELECT * FROM [ExternalQuery]",
            2,
            false);
        new WorkbookOwnershipRegistry().Register(workbook, identity, connectionName);

        Assert.Throws<InvalidOperationException>(() =>
            new CanonicalDataAuditor().DeletePriorManagedConnection(
                workbook,
                identity,
                CanonicalConnectionContract.ConnectionString(queryName),
                CanonicalConnectionContract.CommandText(queryName)));

        Assert.Equal(0, replacement.DeleteCount);
        Assert.Same(replacement, Assert.Single(workbook.Connections.All));
    }

    [Fact]
    public void Registry_only_audit_connection_claim_is_removed_without_deleting_anything()
    {
        var workbook = new CanonicalDataLoaderIdempotencyTests.FakeWorkbook();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_connection",
            ManagedObjectKind.DataModelConnection);
        const string connectionName = "Connection - ERB audit";
        string queryName = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_query",
            ManagedObjectKind.CanonicalQuery).ExcelName;
        var registry = new WorkbookOwnershipRegistry();
        registry.Register(workbook, identity, connectionName);

        new CanonicalDataAuditor().DeletePriorManagedConnection(
            workbook,
            identity,
            CanonicalConnectionContract.ConnectionString(queryName),
            CanonicalConnectionContract.CommandText(queryName));

        Assert.Empty(registry.Load(workbook));
        Assert.Empty(workbook.Connections.All);
    }

    [Fact]
    public void Exact_owned_audit_connection_is_deleted_after_live_contract_validation()
    {
        var workbook = new CanonicalDataLoaderIdempotencyTests.FakeWorkbook();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_connection",
            ManagedObjectKind.DataModelConnection);
        const string connectionName = "Connection - ERB audit";
        string queryName = new ManagedObjectIdentity(
            "report",
            "source_canonical_audit_query",
            ManagedObjectKind.CanonicalQuery).ExcelName;
        string expectedConnection = CanonicalConnectionContract.ConnectionString(queryName);
        string expectedCommand = CanonicalConnectionContract.CommandText(queryName);
        var connection = workbook.Connections.AddExisting(
            connectionName,
            expectedConnection,
            expectedCommand,
            2,
            false);
        var registry = new WorkbookOwnershipRegistry();
        registry.Register(workbook, identity, connectionName);

        new CanonicalDataAuditor().DeletePriorManagedConnection(
            workbook,
            identity,
            expectedConnection,
            expectedCommand);

        Assert.Equal(1, connection.DeleteCount);
        Assert.Empty(workbook.Connections.All);
        Assert.Empty(registry.Load(workbook));
    }

    private static MeasureDefinition Aggregate(
        string id,
        string field,
        AggregateFunction function)
    {
        return new MeasureDefinition
        {
            Id = id,
            Label = id,
            Expression = new AggregateMeasureExpression
            {
                Field = field,
                Function = function
            }
        };
    }

    private static int CountOccurrences(string value, string expected)
    {
        var count = 0;
        var offset = 0;
        while ((offset = value.IndexOf(expected, offset, StringComparison.Ordinal)) >= 0)
        {
            count++;
            offset += expected.Length;
        }

        return count;
    }

    public sealed class AuditResultTable
    {
        public AuditResultTable(IReadOnlyDictionary<string, object?> values, int rowCount = 1)
        {
            ListRows = new AuditRows(rowCount);
            ListColumns = new AuditColumns(values);
        }

        public AuditRows ListRows { get; }

        public AuditColumns ListColumns { get; }
    }

    public sealed class AuditRows
    {
        public AuditRows(int count)
        {
            Count = count;
        }

        public int Count { get; }
    }

    public sealed class AuditColumns
    {
        private readonly IReadOnlyDictionary<string, object?> values;

        public AuditColumns(IReadOnlyDictionary<string, object?> values)
        {
            this.values = values;
        }

        public int Count => values.Count;

        public AuditColumn Item(string name)
        {
            return new AuditColumn(values[name]);
        }
    }

    public sealed class AuditColumn
    {
        public AuditColumn(object? value)
        {
            DataBodyRange = new AuditRange(value);
        }

        public AuditRange DataBodyRange { get; }
    }

    public sealed class AuditRange
    {
        public AuditRange(object? value)
        {
            Cells = new AuditCells(value);
        }

        public AuditCells Cells { get; }
    }

    public sealed class AuditCells
    {
        private readonly object? value;

        public AuditCells(object? value)
        {
            this.value = value;
        }

        public AuditCell this[int row, int column] => new AuditCell(value);
    }

    public sealed class AuditCell
    {
        public AuditCell(object? value)
        {
            Value2 = value;
        }

        public object? Value2 { get; }
    }
}
