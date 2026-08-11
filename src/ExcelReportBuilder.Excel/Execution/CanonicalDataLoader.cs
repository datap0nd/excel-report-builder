using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public enum CanonicalBackend
    {
        Worksheet,
        DataModel
    }

    public sealed class CanonicalLoadPlan
    {
        public CanonicalBackend Backend { get; set; }

        public long ProjectedRows { get; set; }

        public string QueryName { get; set; } = string.Empty;

        public string TableOrConnectionName { get; set; } = string.Empty;
    }

    public sealed class CanonicalDestinationRouter
    {
        public const long ExcelWorksheetRows = 1_048_576;

        public CanonicalBackend Choose(long projectedDataRows, int headerRows = 1)
        {
            if (projectedDataRows < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(projectedDataRows));
            }

            if (headerRows < 1 || headerRows >= ExcelWorksheetRows)
            {
                throw new ArgumentOutOfRangeException(nameof(headerRows));
            }

            return projectedDataRows <= ExcelWorksheetRows - headerRows
                ? CanonicalBackend.Worksheet
                : CanonicalBackend.DataModel;
        }

        public CanonicalBackend ResolveRequired(
            long projectedDataRows,
            CanonicalBackend requiredBackend,
            int headerRows = 1)
        {
            CanonicalBackend sizeRoute = Choose(projectedDataRows, headerRows);
            if (requiredBackend == CanonicalBackend.Worksheet &&
                sizeRoute == CanonicalBackend.DataModel)
            {
                throw new InvalidOperationException(
                    "The validated plan cannot load an oversized normalized result to a worksheet.");
            }

            if (requiredBackend != CanonicalBackend.Worksheet &&
                requiredBackend != CanonicalBackend.DataModel)
            {
                throw new ArgumentOutOfRangeException(nameof(requiredBackend));
            }

            return requiredBackend;
        }
    }

    public sealed class RestrictedQueryFormulaPolicy
    {
        private static readonly string[] ForbiddenAccessors =
        {
            "File.Contents",
            "Folder.Files",
            "Web.Contents",
            "Odbc.",
            "OleDb.",
            "Sql.Database",
            "Access.Database",
            "SharePoint.",
            "AzureStorage.",
            "Exchange.Contents",
            "AnalysisServices."
        };

        public void DemandWorkbookOnly(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new ArgumentException("A compiled query formula is required.", nameof(formula));
            }

            if (formula.IndexOf("Excel.CurrentWorkbook", StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("The query must read from the selected object in the current workbook.");
            }

            foreach (var accessor in ForbiddenAccessors)
            {
                if (formula.IndexOf(accessor, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    throw new InvalidOperationException("The compiled query contains a forbidden external data accessor.");
                }
            }
        }
    }

    /// <summary>
    /// Creates the complete canonical result in a hidden worksheet or the Data
    /// Model. It never falls back to a truncated worksheet result.
    /// </summary>
    public sealed class CanonicalDataLoader
    {
        private const int SourceExternal = 0;
        private const int HasHeadersYes = 1;
        private const int CommandTypeSql = 2;
        private const int ConnectionTypeOleDb = 1;
        private readonly CanonicalDestinationRouter router;
        private readonly RestrictedQueryFormulaPolicy formulaPolicy;
        private readonly ManagedQueryService queryService;
        private readonly ManagedWorksheetService worksheetService;
        private readonly WorkbookOwnershipRegistry registry;

        public CanonicalDataLoader(
            CanonicalDestinationRouter? router = null,
            RestrictedQueryFormulaPolicy? formulaPolicy = null,
            ManagedQueryService? queryService = null,
            ManagedWorksheetService? worksheetService = null,
            WorkbookOwnershipRegistry? registry = null)
        {
            this.router = router ?? new CanonicalDestinationRouter();
            this.formulaPolicy = formulaPolicy ?? new RestrictedQueryFormulaPolicy();
            this.registry = registry ?? new WorkbookOwnershipRegistry();
            this.queryService = queryService ?? new ManagedQueryService(this.registry);
            this.worksheetService = worksheetService ?? new ManagedWorksheetService();
        }

        public CanonicalLoadPlan Load(
            dynamic workbook,
            string reportId,
            string objectId,
            string restrictedFormula,
            long projectedRows,
            CanonicalBackend requiredBackend,
            IExcelProgressSink? progressSink = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            formulaPolicy.DemandWorkbookOnly(restrictedFormula);
            var backend = router.ResolveRequired(projectedRows, requiredBackend);
            var queryIdentity = new ManagedObjectIdentity(reportId, objectId + "_query", ManagedObjectKind.CanonicalQuery);
            var queryName = queryIdentity.ExcelName;

            // Validate the opposite managed backend before updating the stable
            // query. Excel PivotCaches have no supported delete API, so an old
            // cache can retain a dependency on that backend after a route
            // change. Keeping at most one exact owned object per backend avoids
            // breaking that dependency. The pivot-cache source contract still
            // forces a new cache for the newly selected route.
            if (backend == CanonicalBackend.Worksheet)
            {
                DemandManagedDataModelBackendOrMissing(workbook, reportId, objectId);
            }
            else
            {
                DemandManagedWorksheetBackendOrMissing(workbook, reportId, objectId);
            }

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Normalizing,
                Operation = "Creating the managed canonical query.",
                ManagedObject = queryName,
                ProjectedRows = projectedRows
            });
            queryService.ReplaceQuery(workbook, queryIdentity, queryName, restrictedFormula);

            if (backend == CanonicalBackend.Worksheet)
            {
                return LoadWorksheet(workbook, reportId, objectId, queryName, projectedRows, progressSink, cancellationToken);
            }

            return LoadDataModel(workbook, reportId, objectId, queryName, projectedRows, progressSink, cancellationToken);
        }

        private void DemandManagedDataModelBackendOrMissing(
            dynamic workbook,
            string reportId,
            string objectId)
        {
            var identity = new ManagedObjectIdentity(
                reportId,
                objectId + "_model",
                ManagedObjectKind.DataModelConnection);
            var connectionName = ManagedName.Create("Model", reportId, objectId);
            dynamic? existing = TryGetConnection(workbook, connectionName);
            DemandExactConnectionRegistryState(
                registry.Load((object)workbook),
                identity,
                connectionName,
                existing != null);
        }

        private void DemandManagedWorksheetBackendOrMissing(
            dynamic workbook,
            string reportId,
            string objectId)
        {
            var identity = new ManagedObjectIdentity(
                reportId,
                objectId + "_canonical",
                ManagedObjectKind.CanonicalTable);
            dynamic? existing = worksheetService.TryFindOwnedWorksheet(workbook, identity);
            if (existing != null)
            {
                return;
            }

            var records = registry.Load((object)workbook).Where(record =>
                string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == identity.Kind).ToList();
            if (records.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the prior worksheet canonical source.");
            }

            // A registry-only record is stale and cannot represent a live
            // dependency. Remove only that exact claim.
            if (records.Count == 1)
            {
                registry.Remove((object)workbook, new[] { identity });
            }
        }

        private CanonicalLoadPlan LoadWorksheet(
            dynamic workbook,
            string reportId,
            string objectId,
            string queryName,
            long projectedRows,
            IExcelProgressSink progressSink,
            CancellationToken cancellationToken)
        {
            var sheetIdentity = new ManagedObjectIdentity(reportId, objectId + "_canonical", ManagedObjectKind.CanonicalTable);
            dynamic sheet = worksheetService.GetOrCreateHidden(workbook, sheetIdentity);
            var tableName = ManagedName.Create("Canonical", reportId, objectId);
            var connectionString = CanonicalConnectionContract.ConnectionString(queryName);
            var commandText = CanonicalConnectionContract.CommandText(queryName);
            dynamic? existing = TryGetListObject(sheet, tableName);
            dynamic listObject;
            if (existing != null)
            {
                if (Convert.ToInt32(sheet.ListObjects.Count) != 1)
                {
                    throw new InvalidOperationException(
                        "The managed canonical worksheet contains unexpected tables and cannot be refreshed safely.");
                }

                listObject = existing;
                DemandExactWorksheetConnection(listObject.QueryTable, connectionString, commandText);
            }
            else
            {
                worksheetService.ClearOwned(sheet, sheetIdentity);
                listObject = sheet.ListObjects.Add(
                    SourceExternal,
                    connectionString,
                    Type.Missing,
                    HasHeadersYes,
                    sheet.Range["A1"]);
                listObject.Name = tableName;
            }
            dynamic queryTable = listObject.QueryTable;
            ConfigureWorksheetConnection(queryTable, commandText);
            queryTable.RefreshStyle = 1;
            queryTable.AdjustColumnWidth = false;
            queryTable.PreserveFormatting = true;
            DemandExactWorksheetConnection(queryTable, connectionString, commandText);
            var backgroundRefresh = TryEnableBackgroundRefresh(queryTable);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Normalizing,
                Operation = backgroundRefresh
                    ? (existing == null
                        ? "Loading all normalized rows to a managed hidden table in the background."
                        : "Refreshing the existing managed hidden table in place in the background.")
                    : "Excel cannot refresh this query in the background. The next operation may temporarily block Excel.",
                ManagedObject = tableName,
                ProjectedRows = projectedRows
            });
            queryTable.Refresh(backgroundRefresh);
            if (backgroundRefresh)
            {
                PollRefresh(
                    () => Convert.ToBoolean(queryTable.Refreshing),
                    () => queryTable.CancelRefresh(),
                    "Still loading normalized rows to the managed hidden table.",
                    tableName,
                    projectedRows,
                    progressSink,
                    cancellationToken);
            }
            registry.Register(workbook, sheetIdentity, tableName);
            return new CanonicalLoadPlan
            {
                Backend = CanonicalBackend.Worksheet,
                ProjectedRows = projectedRows,
                QueryName = queryName,
                TableOrConnectionName = tableName
            };
        }

        private CanonicalLoadPlan LoadDataModel(
            dynamic workbook,
            string reportId,
            string objectId,
            string queryName,
            long projectedRows,
            IExcelProgressSink progressSink,
            CancellationToken cancellationToken)
        {
            var identity = new ManagedObjectIdentity(reportId, objectId + "_model", ManagedObjectKind.DataModelConnection);
            var connectionName = ManagedName.Create("Model", reportId, objectId);
            dynamic? existing = TryGetConnection(workbook, connectionName);
            DemandExactConnectionRegistryState(
                registry.Load((object)workbook),
                identity,
                connectionName,
                existing != null);
            dynamic connection;
            if (existing != null)
            {
                connection = existing;
                DemandExactDataModelConnection(connection, connectionName, queryName);
            }
            else
            {
                var connectionString = CanonicalConnectionContract.ConnectionString(queryName);
                var commandText = CanonicalConnectionContract.CommandText(queryName);
                connection = workbook.Connections.Add2(
                    connectionName,
                    "Managed canonical data for a dense management report.",
                    connectionString,
                    commandText,
                    CommandTypeSql,
                    true,
                    false);
                DemandExactDataModelConnection(connection, connectionName, queryName);
            }
            registry.Register(workbook, identity, connectionName);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Normalizing,
                Operation = existing == null
                    ? "Loading the complete normalized result to Excel's large-data engine. This operation may temporarily block Excel."
                    : "Refreshing the existing managed large-data connection in place. This operation may temporarily block Excel.",
                ManagedObject = connectionName,
                ProjectedRows = projectedRows
            });
            dynamic oleDbConnection = connection.OLEDBConnection;
            var backgroundRefresh = TryEnableBackgroundRefresh(oleDbConnection);
            connection.Refresh();
            if (backgroundRefresh)
            {
                PollRefresh(
                    () => Convert.ToBoolean(oleDbConnection.Refreshing),
                    () => oleDbConnection.CancelRefresh(),
                    "Still loading normalized rows to Excel's large-data engine.",
                    connectionName,
                    projectedRows,
                    progressSink,
                    cancellationToken);
            }

            return new CanonicalLoadPlan
            {
                Backend = CanonicalBackend.DataModel,
                ProjectedRows = projectedRows,
                QueryName = queryName,
                TableOrConnectionName = connectionName
            };
        }

        private static void PollRefresh(
            Func<bool> isRefreshing,
            Action cancelRefresh,
            string heartbeatMessage,
            string managedObject,
            long projectedRows,
            IExcelProgressSink progressSink,
            CancellationToken cancellationToken)
        {
            var nextHeartbeat = DateTime.UtcNow.AddSeconds(5);
            while (isRefreshing())
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    cancelRefresh();
                    cancellationToken.ThrowIfCancellationRequested();
                }

                if (DateTime.UtcNow >= nextHeartbeat)
                {
                    progressSink.Report(new ExcelProgress
                    {
                        Stage = ExcelBuildStage.Normalizing,
                        Operation = heartbeatMessage,
                        ManagedObject = managedObject,
                        ProjectedRows = projectedRows
                    });
                    nextHeartbeat = DateTime.UtcNow.AddSeconds(5);
                }

                Thread.Sleep(200);
            }

            cancellationToken.ThrowIfCancellationRequested();
        }

        private static bool TryEnableBackgroundRefresh(dynamic refreshable)
        {
            try
            {
                refreshable.BackgroundQuery = true;
                return Convert.ToBoolean(refreshable.BackgroundQuery);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static dynamic? TryGetConnection(dynamic workbook, string name)
        {
            try
            {
                return workbook.Connections.Item(name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static dynamic? TryGetListObject(dynamic worksheet, string name)
        {
            try
            {
                return worksheet.ListObjects.Item(name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static void DemandExactWorksheetConnection(
            dynamic queryTable,
            string expectedConnectionString,
            string expectedCommandText)
        {
            if (queryTable == null)
            {
                throw new InvalidOperationException(
                    "The managed canonical table has no query connection and cannot be refreshed safely.");
            }

            try
            {
                CanonicalConnectionContract.DemandExactConnectionAndCommandType(
                    queryTable.Connection,
                    queryTable.CommandType,
                    expectedConnectionString,
                    "worksheet query");

                DemandExactWorksheetWorkbookConnection(
                    queryTable.WorkbookConnection,
                    expectedConnectionString,
                    expectedCommandText);
            }
            catch (Exception exception) when (!(exception is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    "The managed canonical worksheet query connection could not be verified before refresh.",
                    exception);
            }
        }

        internal static void ConfigureWorksheetConnection(
            dynamic queryTable,
            string commandText)
        {
            if (queryTable == null)
            {
                throw new ArgumentNullException(nameof(queryTable));
            }

            queryTable.CommandType = CommandTypeSql;
            dynamic workbookConnection = queryTable.WorkbookConnection;
            dynamic oleDbConnection = workbookConnection.OLEDBConnection;
            oleDbConnection.CommandType = CommandTypeSql;
            oleDbConnection.CommandText = commandText;
        }

        internal static void DemandExactWorksheetWorkbookConnection(
            dynamic workbookConnection,
            string expectedConnectionString,
            string expectedCommandText)
        {
            object? workbookConnectionObject = workbookConnection as object;
            if (workbookConnectionObject == null)
            {
                throw new InvalidOperationException(
                    "The managed canonical worksheet query no longer has its expected workbook-only connection.");
            }

            dynamic verifiedWorkbookConnection = workbookConnectionObject;
            if (Convert.ToInt32(
                    verifiedWorkbookConnection.Type,
                    CultureInfo.InvariantCulture) != ConnectionTypeOleDb ||
                Convert.ToBoolean(
                    verifiedWorkbookConnection.InModel,
                    CultureInfo.InvariantCulture))
            {
                throw new InvalidOperationException(
                    "The managed canonical worksheet query no longer has its expected workbook-only connection.");
            }

            object? oleDbConnectionObject = verifiedWorkbookConnection.OLEDBConnection as object;
            if (oleDbConnectionObject == null)
            {
                throw new InvalidOperationException(
                    "The managed canonical worksheet query no longer has its expected OLE DB connection.");
            }

            dynamic oleDbConnection = oleDbConnectionObject;
            CanonicalConnectionContract.DemandExactOleDbContract(
                oleDbConnection.Connection,
                oleDbConnection.CommandType,
                oleDbConnection.CommandText,
                expectedConnectionString,
                expectedCommandText,
                "worksheet workbook connection");
        }

        internal static void DemandExactDataModelConnection(
            dynamic connection,
            string expectedConnectionName,
            string queryName)
        {
            if (connection == null)
            {
                throw new InvalidOperationException(
                    "The managed Data Model connection is missing and cannot be refreshed safely.");
            }

            try
            {
                var actualName = Convert.ToString(
                    connection.Name,
                    CultureInfo.InvariantCulture) ?? string.Empty;
                if (!string.Equals(actualName, expectedConnectionName, StringComparison.Ordinal) ||
                    Convert.ToInt32(connection.Type, CultureInfo.InvariantCulture) != ConnectionTypeOleDb ||
                    !Convert.ToBoolean(connection.InModel, CultureInfo.InvariantCulture))
                {
                    throw new InvalidOperationException(
                        "The managed Data Model connection no longer has its expected identity or model membership.");
                }

                dynamic oleDbConnection = connection.OLEDBConnection;
                CanonicalConnectionContract.DemandExactOleDbContract(
                    oleDbConnection.Connection,
                    oleDbConnection.CommandType,
                    oleDbConnection.CommandText,
                    CanonicalConnectionContract.ConnectionString(queryName),
                    CanonicalConnectionContract.CommandText(queryName),
                    "Data Model connection");
            }
            catch (Exception exception) when (!(exception is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    "The managed Data Model connection could not be verified before refresh.",
                    exception);
            }
        }

        internal static void DemandExactConnectionRegistryState(
            System.Collections.Generic.IReadOnlyList<ManagedObjectRecord> records,
            ManagedObjectIdentity identity,
            string connectionName,
            bool connectionExists)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (identity == null) throw new ArgumentNullException(nameof(identity));
            if (identity.Kind != ManagedObjectKind.DataModelConnection)
            {
                throw new ArgumentException(
                    "A Data Model connection ownership identity is required.",
                    nameof(identity));
            }

            if (string.IsNullOrWhiteSpace(connectionName))
            {
                throw new ArgumentException("A connection name is required.", nameof(connectionName));
            }

            var exact = records.Where(record =>
                string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == identity.Kind).ToList();
            if (exact.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the managed Data Model connection.");
            }

            if (records.Any(record =>
                    record.Kind == ManagedObjectKind.DataModelConnection &&
                    !exact.Contains(record) &&
                    string.Equals(
                        record.ExcelName,
                        connectionName,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "The requested Data Model connection name is owned by another object.");
            }

            ManagedObjectRecord? registration = exact.SingleOrDefault();
            if (connectionExists && registration == null)
            {
                throw new InvalidOperationException(
                    "A connection with the requested name exists but is unmanaged.");
            }

            if (!connectionExists && registration != null)
            {
                throw new InvalidOperationException(
                    "The exact managed Data Model connection registration exists but its workbook connection is missing or renamed.");
            }

            if (registration != null &&
                !string.Equals(
                    registration.ExcelName,
                    connectionName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The managed Data Model connection is registered under a different name.");
            }
        }

    }

    internal static class CanonicalConnectionContract
    {
        private const int CommandTypeSql = 2;

        public static string ConnectionString(string queryName)
        {
            if (string.IsNullOrWhiteSpace(queryName))
            {
                throw new ArgumentException("A managed query name is required.", nameof(queryName));
            }

            return "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" +
                   queryName + ";Extended Properties=\"\"";
        }

        public static string CommandText(string queryName)
        {
            if (string.IsNullOrWhiteSpace(queryName))
            {
                throw new ArgumentException("A managed query name is required.", nameof(queryName));
            }

            return "SELECT * FROM [" + queryName.Replace("]", "]]" ) + "]";
        }

        public static void DemandExactOleDbContract(
            object? actualConnection,
            object? actualCommandType,
            object? actualCommandText,
            string expectedConnection,
            string expectedCommandText,
            string label)
        {
            DemandExactConnectionAndCommandType(
                actualConnection,
                actualCommandType,
                expectedConnection,
                label);

            var commandText = ReadSingleCommandText(actualCommandText, label);
            if (!string.Equals(commandText, expectedCommandText, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The managed " + label +
                    " no longer matches the exact workbook-only connection contract.");
            }
        }

        public static void DemandExactConnectionAndCommandType(
            object? actualConnection,
            object? actualCommandType,
            string expectedConnection,
            string label)
        {
            var connection = Convert.ToString(
                actualConnection,
                CultureInfo.InvariantCulture) ?? string.Empty;
            int commandType;
            try
            {
                commandType = Convert.ToInt32(actualCommandType, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "The managed " + label + " has an invalid command type.",
                    exception);
            }

            if (!string.Equals(connection, expectedConnection, StringComparison.Ordinal) ||
                commandType != CommandTypeSql)
            {
                throw new InvalidOperationException(
                    "The managed " + label +
                    " no longer matches the exact workbook-only connection contract.");
            }
        }

        private static string ReadSingleCommandText(object? value, string label)
        {
            if (value is string text)
            {
                return text;
            }

            if (value is Array array && array.Rank == 1 && array.Length == 1)
            {
                object? item = array.GetValue(array.GetLowerBound(0));
                if (item is string arrayText)
                {
                    return arrayText;
                }
            }

            throw new InvalidOperationException(
                "The managed " + label + " does not expose exactly one SQL command.");
        }
    }
}
