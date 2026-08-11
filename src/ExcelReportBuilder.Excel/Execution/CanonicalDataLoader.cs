using System;
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
            IExcelProgressSink? progressSink = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            formulaPolicy.DemandWorkbookOnly(restrictedFormula);
            var backend = router.Choose(projectedRows);
            var queryIdentity = new ManagedObjectIdentity(reportId, objectId + "_query", ManagedObjectKind.CanonicalQuery);
            var queryName = queryIdentity.ExcelName;

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
            worksheetService.ClearOwned(sheet, sheetIdentity);

            var tableName = ManagedName.Create("Canonical", reportId, objectId);
            var connectionString =
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" +
                queryName + ";Extended Properties=\"\"";
            dynamic listObject = sheet.ListObjects.Add(
                SourceExternal,
                connectionString,
                Type.Missing,
                HasHeadersYes,
                sheet.Range["A1"]);
            listObject.Name = tableName;
            dynamic queryTable = listObject.QueryTable;
            queryTable.CommandType = CommandTypeSql;
            queryTable.CommandText = new[] { "SELECT * FROM [" + queryName.Replace("]", "]]" ) + "]" };
            queryTable.RefreshStyle = 1;
            queryTable.AdjustColumnWidth = false;
            queryTable.PreserveFormatting = true;
            var backgroundRefresh = TryEnableBackgroundRefresh(queryTable);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Normalizing,
                Operation = backgroundRefresh
                    ? "Loading all normalized rows to a managed hidden table in the background."
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
            if (existing != null)
            {
                if (!registry.IsOwned(workbook, identity, connectionName))
                {
                    throw new InvalidOperationException("A connection with the requested name exists but is unmanaged.");
                }

                existing.Delete();
            }

            var connectionString =
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" +
                queryName + ";Extended Properties=\"\"";
            var commandText = "SELECT * FROM [" + queryName.Replace("]", "]]" ) + "]";
            dynamic connection = workbook.Connections.Add2(
                connectionName,
                "Managed canonical data for a dense management report.",
                connectionString,
                commandText,
                CommandTypeSql,
                true,
                false);
            registry.Register(workbook, identity, connectionName);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Normalizing,
                Operation = "Loading the complete normalized result to Excel's large-data engine. This operation may temporarily block Excel.",
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
    }
}
