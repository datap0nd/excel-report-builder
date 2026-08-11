using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class CanonicalDataLoaderIdempotencyTests
{
    [Fact]
    public void Same_data_model_route_updates_query_and_refreshes_connection_in_place()
    {
        var workbook = new FakeWorkbook();
        var loader = new CanonicalDataLoader();
        const string firstFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";
        const string secondFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content], Renamed = Table.RenameColumns(Source, {{\"Amount\", \"Value\"}}) in Renamed";

        CanonicalLoadPlan first = loader.Load(
            workbook,
            "report",
            "source",
            firstFormula,
            12,
            CanonicalBackend.DataModel);
        FakeQuery query = Assert.Single(workbook.Queries.All);
        FakeConnection connection = Assert.Single(workbook.Connections.All);

        CanonicalLoadPlan second = loader.Load(
            workbook,
            "report",
            "source",
            secondFormula,
            12,
            CanonicalBackend.DataModel);

        Assert.Same(query, Assert.Single(workbook.Queries.All));
        Assert.Same(connection, Assert.Single(workbook.Connections.All));
        Assert.Equal(secondFormula, query.Formula);
        Assert.Equal(1, workbook.Queries.AddCount);
        Assert.Equal(1, workbook.Connections.AddCount);
        Assert.Equal(0, query.DeleteCount);
        Assert.Equal(0, connection.DeleteCount);
        Assert.Equal(2, connection.RefreshCount);
        Assert.Equal(first.TableOrConnectionName, second.TableOrConnectionName);
        Assert.Equal(CanonicalBackend.DataModel, second.Backend);
    }

    [Fact]
    public void Edited_managed_query_formula_is_rejected_without_formula_assignment()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedQueryService();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_query",
            ManagedObjectKind.CanonicalQuery);
        const string queryName = "ERB_CanonicalQuery_report_source";
        const string firstFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";
        const string secondFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Table.Buffer(Source)";

        FakeQuery query = service.ReplaceQuery(workbook, identity, queryName, firstFormula);
        query.ReplaceFormulaExternally(
            "let Source = Web.Contents(\"https://example.invalid/data\") in Source");

        Assert.Throws<InvalidOperationException>(() =>
            service.ReplaceQuery(workbook, identity, queryName, secondFormula));

        Assert.Equal(0, query.FormulaAssignmentCount);
        Assert.Contains("Web.Contents", query.Formula, StringComparison.Ordinal);
    }

    [Fact]
    public void Same_named_replacement_query_is_rejected_without_formula_assignment()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedQueryService();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_query",
            ManagedObjectKind.CanonicalQuery);
        const string queryName = "ERB_CanonicalQuery_report_source";
        const string managedFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

        FakeQuery original = service.ReplaceQuery(
            workbook,
            identity,
            queryName,
            managedFormula);
        original.Delete();
        FakeQuery replacement = workbook.Queries.Add(
            queryName,
            "let Source = Web.Contents(\"https://example.invalid/data\") in Source");

        Assert.Throws<InvalidOperationException>(() =>
            service.ReplaceQuery(workbook, identity, queryName, managedFormula));

        Assert.Equal(0, replacement.FormulaAssignmentCount);
        Assert.Contains("Web.Contents", replacement.Formula, StringComparison.Ordinal);
    }

    [Fact]
    public void Registry_only_query_claim_is_removed_before_recreation()
    {
        var workbook = new FakeWorkbook();
        var service = new ManagedQueryService();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_query",
            ManagedObjectKind.CanonicalQuery);
        const string queryName = "ERB_CanonicalQuery_report_source";
        const string firstFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";
        const string secondFormula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Table.Buffer(Source)";

        FakeQuery original = service.ReplaceQuery(
            workbook,
            identity,
            queryName,
            firstFormula);
        original.Delete();

        FakeQuery recreated = service.ReplaceQuery(
            workbook,
            identity,
            queryName,
            secondFormula);

        Assert.Equal(secondFormula, recreated.Formula);
        Assert.Equal(2, workbook.Queries.AddCount);
        ManagedObjectRecord record = Assert.Single(new WorkbookOwnershipRegistry().Load(workbook));
        Assert.False(string.IsNullOrWhiteSpace(record.SourceContract));
    }

    [Fact]
    public void Same_named_unmanaged_data_model_connection_is_rejected_without_refresh_or_delete()
    {
        var workbook = new FakeWorkbook();
        var connectionName = ManagedName.Create("Model", "report", "source");
        FakeConnection unmanaged = workbook.Connections.AddExisting(connectionName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new CanonicalDataLoader().Load(
                workbook,
                "report",
                "source",
                "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source",
                12,
                CanonicalBackend.DataModel));

        Assert.Contains("unmanaged", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, unmanaged.RefreshCount);
        Assert.Equal(0, unmanaged.DeleteCount);
    }

    [Fact]
    public void Registered_but_missing_data_model_connection_fails_closed_without_creating_a_duplicate()
    {
        var workbook = new FakeWorkbook();
        var identity = new ManagedObjectIdentity(
            "report",
            "source_model",
            ManagedObjectKind.DataModelConnection);
        var connectionName = ManagedName.Create("Model", "report", "source");
        new WorkbookOwnershipRegistry().Register(workbook, identity, connectionName);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new CanonicalDataLoader().Load(
                workbook,
                "report",
                "source",
                "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source",
                12,
                CanonicalBackend.DataModel));

        Assert.Contains("missing or renamed", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(workbook.Connections.All);
        Assert.Equal(0, workbook.Connections.AddCount);
    }

    [Theory]
    [InlineData("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\example\\external.xlsx;Extended Properties=\"Excel 12.0\"")]
    [InlineData("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=https://example.invalid/data;Location=ManagedQuery;Extended Properties=\"\"")]
    [InlineData("OLEDB;Provider=MSDASQL.1;Data Source=$Workbook$;Location=ManagedQuery;Extended Properties=\"\"")]
    public void Edited_data_model_connection_target_is_rejected_before_refresh(string editedConnection)
    {
        var workbook = new FakeWorkbook();
        var loader = new CanonicalDataLoader();
        const string formula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

        loader.Load(
            workbook,
            "report",
            "source",
            formula,
            12,
            CanonicalBackend.DataModel);
        FakeConnection connection = Assert.Single(workbook.Connections.All);
        connection.OLEDBConnection.Connection = editedConnection;

        var exception = Assert.Throws<InvalidOperationException>(() =>
            loader.Load(
                workbook,
                "report",
                "source",
                formula,
                12,
                CanonicalBackend.DataModel));

        Assert.Contains("workbook-only connection contract", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, connection.RefreshCount);
        Assert.Equal(0, connection.DeleteCount);
    }

    [Fact]
    public void Edited_data_model_command_is_rejected_before_refresh()
    {
        var workbook = new FakeWorkbook();
        var loader = new CanonicalDataLoader();
        const string formula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

        loader.Load(
            workbook,
            "report",
            "source",
            formula,
            12,
            CanonicalBackend.DataModel);
        FakeConnection connection = Assert.Single(workbook.Connections.All);
        connection.OLEDBConnection.CommandText = "SELECT * FROM [ExternalQuery]";

        Assert.Throws<InvalidOperationException>(() =>
            loader.Load(
                workbook,
                "report",
                "source",
                formula,
                12,
                CanonicalBackend.DataModel));
        Assert.Equal(1, connection.RefreshCount);
    }

    [Fact]
    public void Edited_data_model_command_type_is_rejected_before_refresh()
    {
        var workbook = new FakeWorkbook();
        var loader = new CanonicalDataLoader();
        const string formula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

        loader.Load(
            workbook,
            "report",
            "source",
            formula,
            12,
            CanonicalBackend.DataModel);
        FakeConnection connection = Assert.Single(workbook.Connections.All);
        connection.OLEDBConnection.CommandType = 1;

        Assert.Throws<InvalidOperationException>(() =>
            loader.Load(
                workbook,
                "report",
                "source",
                formula,
                12,
                CanonicalBackend.DataModel));
        Assert.Equal(1, connection.RefreshCount);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 2)]
    public void Edited_data_model_membership_or_connection_type_is_rejected_before_refresh(
        bool inModel,
        int connectionType)
    {
        var workbook = new FakeWorkbook();
        var loader = new CanonicalDataLoader();
        const string formula =
            "let Source = Excel.CurrentWorkbook(){[Name=\"RawData\"]}[Content] in Source";

        loader.Load(
            workbook,
            "report",
            "source",
            formula,
            12,
            CanonicalBackend.DataModel);
        FakeConnection connection = Assert.Single(workbook.Connections.All);
        connection.InModel = inModel;
        connection.Type = connectionType;

        Assert.Throws<InvalidOperationException>(() =>
            loader.Load(
                workbook,
                "report",
                "source",
                formula,
                12,
                CanonicalBackend.DataModel));
        Assert.Equal(1, connection.RefreshCount);
    }

    [Fact]
    public void Exact_worksheet_query_contract_accepts_array_command_and_non_model_membership()
    {
        const string queryName = "ERB_CanonicalQuery_report_source";
        var queryTable = FakeQueryTable.Valid(queryName);

        CanonicalDataLoader.DemandExactWorksheetConnection(
            queryTable,
            CanonicalConnectionContract.ConnectionString(queryName),
            CanonicalConnectionContract.CommandText(queryName));
    }

    [Fact]
    public void Worksheet_query_contract_does_not_read_query_table_command_text()
    {
        const string queryName = "ERB_CanonicalQuery_report_source";
        var queryTable = FakeQueryTable.Valid(queryName);
        CanonicalDataLoader.DemandExactWorksheetConnection(
            queryTable,
            CanonicalConnectionContract.ConnectionString(queryName),
            CanonicalConnectionContract.CommandText(queryName));

    }

    [Fact]
    public void Worksheet_query_configuration_uses_only_workbook_connection_command_text()
    {
        const string queryName = "ERB_CanonicalQuery_report_source";
        var queryTable = FakeQueryTable.Valid(queryName);
        queryTable.WorkbookConnection.OLEDBConnection.CommandType = 0;
        queryTable.WorkbookConnection.OLEDBConnection.CommandText = string.Empty;

        CanonicalDataLoader.ConfigureWorksheetConnection(
            queryTable,
            CanonicalConnectionContract.CommandText(queryName));

        Assert.Equal(2, queryTable.CommandType);
        Assert.Equal(2, queryTable.WorkbookConnection.OLEDBConnection.CommandType);
        Assert.Equal(
            CanonicalConnectionContract.CommandText(queryName),
            queryTable.WorkbookConnection.OLEDBConnection.CommandText);
    }

    [Theory]
    [InlineData("query-connection")]
    [InlineData("query-command-type")]
    [InlineData("workbook-connection")]
    [InlineData("workbook-command")]
    [InlineData("workbook-command-type")]
    [InlineData("workbook-connection-type")]
    [InlineData("model-membership")]
    public void Edited_worksheet_query_contract_is_rejected(string edit)
    {
        const string queryName = "ERB_CanonicalQuery_report_source";
        var queryTable = FakeQueryTable.Valid(queryName);
        switch (edit)
        {
            case "query-connection":
                queryTable.Connection =
                    "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\example\\external.xlsx";
                break;
            case "query-command-type":
                queryTable.CommandType = 1;
                break;
            case "workbook-connection":
                queryTable.WorkbookConnection.OLEDBConnection.Connection =
                    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=https://example.invalid/data";
                break;
            case "workbook-command":
                queryTable.WorkbookConnection.OLEDBConnection.CommandText =
                    "SELECT * FROM [ExternalQuery]";
                break;
            case "workbook-command-type":
                queryTable.WorkbookConnection.OLEDBConnection.CommandType = 1;
                break;
            case "workbook-connection-type":
                queryTable.WorkbookConnection.Type = 2;
                break;
            case "model-membership":
                queryTable.WorkbookConnection.InModel = true;
                break;
        }

        Assert.Throws<InvalidOperationException>(() =>
            CanonicalDataLoader.DemandExactWorksheetConnection(
                queryTable,
                CanonicalConnectionContract.ConnectionString(queryName),
                CanonicalConnectionContract.CommandText(queryName)));
    }

    public sealed class FakeWorkbook
    {
        public FakeQueryCollection Queries { get; } = new();

        public FakeConnectionCollection Connections { get; } = new();

        public FakeWorksheetCollection Worksheets { get; } = new();

        public WorkbookSpecStoreTests.FakeCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FakeQueryCollection
    {
        private readonly List<FakeQuery> values = new();

        public IReadOnlyList<FakeQuery> All => values;

        public int AddCount { get; private set; }

        public FakeQuery Item(string name)
        {
            return values.Single(value =>
                string.Equals(value.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public FakeQuery Add(string name, string formula)
        {
            AddCount++;
            FakeQuery? query = null;
            query = new FakeQuery(name, formula, () => values.Remove(query!));
            values.Add(query);
            return query;
        }
    }

    public sealed class FakeQuery
    {
        private readonly Action delete;
        private string formula;

        public FakeQuery(string name, string formula, Action delete)
        {
            Name = name;
            this.formula = formula;
            this.delete = delete;
        }

        public string Name { get; }

        public string Formula
        {
            get => formula;
            set
            {
                FormulaAssignmentCount++;
                formula = value;
            }
        }

        public int FormulaAssignmentCount { get; private set; }

        public int DeleteCount { get; private set; }

        public void Delete()
        {
            DeleteCount++;
            delete();
        }

        public void ReplaceFormulaExternally(string value)
        {
            formula = value;
        }
    }

    public sealed class FakeConnectionCollection
    {
        private readonly List<FakeConnection> values = new();

        public IReadOnlyList<FakeConnection> All => values;

        public int AddCount { get; private set; }

        public FakeConnection Item(string name)
        {
            return values.Single(value =>
                string.Equals(value.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public FakeConnection Add2(
            string name,
            string description,
            string connectionString,
            object commandText,
            int commandType,
            bool createModelConnection,
            bool importRelationships)
        {
            AddCount++;
            return AddExisting(
                name,
                connectionString,
                commandText,
                commandType,
                createModelConnection);
        }

        public FakeConnection AddExisting(string name)
        {
            return AddExisting(
                name,
                string.Empty,
                string.Empty,
                0,
                false);
        }

        public FakeConnection AddExisting(
            string name,
            string connectionString,
            object commandText,
            int commandType,
            bool inModel)
        {
            FakeConnection? connection = null;
            connection = new FakeConnection(
                name,
                connectionString,
                commandText,
                commandType,
                inModel,
                () => values.Remove(connection!));
            values.Add(connection);
            return connection;
        }
    }

    public sealed class FakeConnection
    {
        private readonly Action delete;

        public FakeConnection(
            string name,
            string connectionString,
            object commandText,
            int commandType,
            bool inModel,
            Action delete)
        {
            Name = name;
            InModel = inModel;
            OLEDBConnection = new FakeOleDbConnection
            {
                Connection = connectionString,
                CommandText = commandText,
                CommandType = commandType
            };
            this.delete = delete;
        }

        public string Name { get; }

        public int Type { get; set; } = 1;

        public bool InModel { get; set; }

        public FakeOleDbConnection OLEDBConnection { get; }

        public int RefreshCount { get; private set; }

        public int DeleteCount { get; private set; }

        public void Refresh()
        {
            RefreshCount++;
        }

        public void Delete()
        {
            DeleteCount++;
            delete();
        }
    }

    public sealed class FakeOleDbConnection
    {
        public string Connection { get; set; } = string.Empty;

        public int CommandType { get; set; }

        public object CommandText { get; set; } = string.Empty;

        public bool BackgroundQuery { get; set; }

        public bool Refreshing => false;

        public void CancelRefresh()
        {
        }
    }

    public sealed class FakeQueryTable
    {
        public string Connection { get; set; } = string.Empty;

        public int CommandType { get; set; }

        public object CommandText
        {
            get => throw new InvalidOperationException(
                "QueryTable.CommandText cannot be read while its worksheet is not active.");
            set => throw new InvalidOperationException(
                "QueryTable.CommandText cannot be written while its worksheet is not active.");
        }

        public FakeConnection WorkbookConnection { get; private set; } = null!;

        public static FakeQueryTable Valid(string queryName)
        {
            string connectionString = CanonicalConnectionContract.ConnectionString(queryName);
            string commandText = CanonicalConnectionContract.CommandText(queryName);
            var workbookConnection = new FakeConnection(
                "Worksheet query connection",
                connectionString,
                commandText,
                2,
                false,
                () => { });
            return new FakeQueryTable
            {
                Connection = connectionString,
                CommandType = 2,
                WorkbookConnection = workbookConnection
            };
        }
    }

    public sealed class FakeWorksheetCollection
    {
        public int Count => 0;

        public object Item(int index)
        {
            throw new InvalidOperationException("No synthetic worksheet exists.");
        }
    }
}
