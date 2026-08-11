using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class CanonicalAuditResult
    {
        public long ActualRows { get; set; }

        public IReadOnlyDictionary<string, decimal> Totals { get; set; } =
            new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class CanonicalAuditField
    {
        public string SourceField { get; set; } = string.Empty;

        public string ResultColumnName { get; set; } = string.Empty;

        public IReadOnlyList<string> MeasureIds { get; set; } = Array.Empty<string>();
    }

    internal sealed class CanonicalAuditQueryPlan
    {
        public const string RowCountColumnName = "ERB_AuditRows";

        public string Formula { get; set; } = string.Empty;

        public IReadOnlyList<CanonicalAuditField> Fields { get; set; } = Array.Empty<CanonicalAuditField>();
    }

    /// <summary>
    /// Compiles a bounded Power Query result that evaluates the exact canonical
    /// transformation independently of every PivotTable. The result always has
    /// one row: the complete canonical row count and one total per unique source
    /// field used by an additive Sum measure.
    /// </summary>
    internal static class CanonicalAuditQueryCompiler
    {
        private const int MaximumAuditFields = 128;

        public static CanonicalAuditQueryPlan Compile(
            string restrictedCanonicalFormula,
            IEnumerable<MeasureDefinition> measures)
        {
            if (string.IsNullOrWhiteSpace(restrictedCanonicalFormula))
            {
                throw new ArgumentException("A canonical query formula is required.", nameof(restrictedCanonicalFormula));
            }

            if (measures == null)
            {
                throw new ArgumentNullException(nameof(measures));
            }

            var measureIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var fieldsByName = new Dictionary<string, MutableAuditField>(StringComparer.OrdinalIgnoreCase);
            var orderedFields = new List<MutableAuditField>();
            foreach (var measure in measures)
            {
                if (measure == null)
                {
                    throw new InvalidOperationException("Audit measures cannot contain null entries.");
                }

                if (string.IsNullOrWhiteSpace(measure.Id) || !measureIds.Add(measure.Id))
                {
                    throw new InvalidOperationException("Audit measure identifiers must be non-blank and unique.");
                }

                if (!(measure.Expression is AggregateMeasureExpression aggregate) ||
                    aggregate.Function != AggregateFunction.Sum)
                {
                    continue;
                }

                DemandSafeFieldName(aggregate.Field);
                if (!fieldsByName.TryGetValue(aggregate.Field, out var auditField))
                {
                    if (orderedFields.Count >= MaximumAuditFields)
                    {
                        throw new InvalidOperationException("The normalized-data audit exceeds the supported field limit.");
                    }

                    auditField = new MutableAuditField
                    {
                        SourceField = aggregate.Field,
                        ResultColumnName = "ERB_AuditTotal_" +
                            (orderedFields.Count + 1).ToString(CultureInfo.InvariantCulture)
                    };
                    fieldsByName.Add(aggregate.Field, auditField);
                    orderedFields.Add(auditField);
                }

                auditField.MeasureIds.Add(measure.Id);
            }

            var selectedFields = orderedFields.Select(field => field.SourceField).ToList();
            var groupingColumn = AllocateInternalColumnName(selectedFields);
            var resultColumns = new List<string> { CanonicalAuditQueryPlan.RowCountColumnName };
            resultColumns.AddRange(orderedFields.Select(field => field.ResultColumnName));

            var builder = new StringBuilder();
            builder.AppendLine("let");
            builder.AppendLine("    Canonical = (");
            AppendIndented(builder, restrictedCanonicalFormula, 8);
            builder.AppendLine();
            builder.AppendLine("    ),");
            if (orderedFields.Count == 0)
            {
                builder.Append("    Result = #table({")
                    .Append(MString(CanonicalAuditQueryPlan.RowCountColumnName))
                    .AppendLine("}, {{Table.RowCount(Canonical)}})");
                builder.AppendLine("in");
                builder.Append("    Result");
                return new CanonicalAuditQueryPlan
                {
                    Formula = builder.ToString(),
                    Fields = Array.Empty<CanonicalAuditField>()
                };
            }

            builder.Append("    Selected = Table.SelectColumns(Canonical, ")
                .Append(MStringList(selectedFields))
                .AppendLine(", MissingField.Error),");
            builder.Append("    Keyed = Table.AddColumn(Selected, ")
                .Append(MString(groupingColumn))
                .AppendLine(", each 1, Int64.Type),");
            builder.Append("    Grouped = Table.Group(Keyed, {")
                .Append(MString(groupingColumn))
                .AppendLine("}, {");
            builder.Append("        {")
                .Append(MString(CanonicalAuditQueryPlan.RowCountColumnName))
                .AppendLine(", each Table.RowCount(_), Int64.Type}");
            foreach (var field in orderedFields)
            {
                builder.Append("        ,{")
                    .Append(MString(field.ResultColumnName))
                    .Append(", each let Values = List.RemoveNulls(Table.Column(_, ")
                    .Append(MString(field.SourceField))
                    .AppendLine(")) in if List.IsEmpty(Values) then 0 else List.Sum(Values), type nullable number}");
            }

            builder.AppendLine("    }),");
            builder.Append("    WithoutKey = Table.RemoveColumns(Grouped, {")
                .Append(MString(groupingColumn))
                .AppendLine("}),");
            builder.Append("    EmptyResult = #table(")
                .Append(MStringList(resultColumns))
                .Append(", {{0");
            for (var index = 0; index < orderedFields.Count; index++)
            {
                builder.Append(", 0");
            }

            builder.AppendLine("}}),");
            builder.AppendLine("    Result = if Table.IsEmpty(WithoutKey) then EmptyResult else WithoutKey");
            builder.AppendLine("in");
            builder.Append("    Result");

            return new CanonicalAuditQueryPlan
            {
                Formula = builder.ToString(),
                Fields = orderedFields.Select(field => new CanonicalAuditField
                {
                    SourceField = field.SourceField,
                    ResultColumnName = field.ResultColumnName,
                    MeasureIds = field.MeasureIds.ToArray()
                }).ToArray()
            };
        }

        private static void DemandSafeFieldName(string field)
        {
            if (string.IsNullOrWhiteSpace(field) || field.Length > 255 || field.Any(char.IsControl))
            {
                throw new InvalidOperationException(
                    "Additive audit fields must have bounded names without control characters.");
            }
        }

        private static string AllocateInternalColumnName(IReadOnlyCollection<string> sourceFields)
        {
            var result = "ERB_AuditGroup";
            while (sourceFields.Contains(result, StringComparer.OrdinalIgnoreCase))
            {
                result += "_";
            }

            return result;
        }

        private static void AppendIndented(StringBuilder builder, string value, int spaces)
        {
            var indentation = new string(' ', spaces);
            var normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');
            for (var index = 0; index < lines.Length; index++)
            {
                builder.Append(indentation).Append(lines[index]);
                if (index < lines.Length - 1)
                {
                    builder.AppendLine();
                }
            }
        }

        private static string MStringList(IEnumerable<string> values)
        {
            return "{" + string.Join(", ", values.Select(MString)) + "}";
        }

        private static string MString(string value)
        {
            if (value.Any(char.IsControl))
            {
                throw new InvalidOperationException("Power Query audit literals cannot contain control characters.");
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private sealed class MutableAuditField
        {
            public string SourceField { get; set; } = string.Empty;

            public string ResultColumnName { get; set; } = string.Empty;

            public List<string> MeasureIds { get; } = new List<string>();
        }
    }

    /// <summary>
    /// Loads the bounded canonical audit query to a managed very-hidden table.
    /// This path is deliberately separate from PivotCaches and PivotTables.
    /// </summary>
    public sealed class CanonicalDataAuditor
    {
        private const int SourceExternal = 0;
        private const int HasHeadersYes = 1;
        private const int CommandTypeSql = 2;
        private readonly RestrictedQueryFormulaPolicy formulaPolicy;
        private readonly ManagedQueryService queryService;
        private readonly ManagedWorksheetService worksheetService;
        private readonly WorkbookOwnershipRegistry registry;

        public CanonicalDataAuditor(
            RestrictedQueryFormulaPolicy? formulaPolicy = null,
            ManagedQueryService? queryService = null,
            ManagedWorksheetService? worksheetService = null,
            WorkbookOwnershipRegistry? registry = null)
        {
            this.formulaPolicy = formulaPolicy ?? new RestrictedQueryFormulaPolicy();
            this.registry = registry ?? new WorkbookOwnershipRegistry();
            this.queryService = queryService ?? new ManagedQueryService(this.registry);
            this.worksheetService = worksheetService ?? new ManagedWorksheetService();
        }

        public CanonicalAuditResult AuditDataModel(
            dynamic workbook,
            string reportId,
            string objectId,
            string restrictedCanonicalFormula,
            IEnumerable<MeasureDefinition> measures,
            long projectedRows,
            IExcelProgressSink? progressSink = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (projectedRows < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(projectedRows));
            }

            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            formulaPolicy.DemandWorkbookOnly(restrictedCanonicalFormula);
            var auditPlan = CanonicalAuditQueryCompiler.Compile(restrictedCanonicalFormula, measures);
            formulaPolicy.DemandWorkbookOnly(auditPlan.Formula);

            cancellationToken.ThrowIfCancellationRequested();
            var sheetIdentity = new ManagedObjectIdentity(
                reportId,
                objectId + "_canonical_audit_sheet",
                ManagedObjectKind.ChecksWorksheet);
            dynamic sheet = worksheetService.GetOrCreateHidden(workbook, sheetIdentity);
            worksheetService.ClearOwned(sheet, sheetIdentity);

            var connectionIdentity = new ManagedObjectIdentity(
                reportId,
                objectId + "_canonical_audit_connection",
                ManagedObjectKind.DataModelConnection);

            var queryIdentity = new ManagedObjectIdentity(
                reportId,
                objectId + "_canonical_audit_query",
                ManagedObjectKind.CanonicalQuery);
            var queryName = queryIdentity.ExcelName;
            var connectionString = CanonicalConnectionContract.ConnectionString(queryName);
            var commandText = CanonicalConnectionContract.CommandText(queryName);
            DeletePriorManagedConnection(
                workbook,
                connectionIdentity,
                connectionString,
                commandText);

            var tableIdentity = new ManagedObjectIdentity(
                reportId,
                objectId + "_canonical_audit_result",
                ManagedObjectKind.CanonicalTable);
            var tableName = ManagedName.Create("CanonicalAudit", reportId, objectId);
            if (FindTable(workbook, tableName) != null)
            {
                throw new InvalidOperationException(
                    "An unmanaged table conflicts with the normalized-data audit result name.");
            }

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Checking,
                Operation = "Preparing an independent normalized-data audit.",
                ManagedObject = queryName,
                ProjectedRows = projectedRows
            });
            queryService.ReplaceQuery(workbook, queryIdentity, queryName, auditPlan.Formula);
            cancellationToken.ThrowIfCancellationRequested();

            dynamic listObject = sheet.ListObjects.Add(
                SourceExternal,
                connectionString,
                Type.Missing,
                HasHeadersYes,
                sheet.Range["A1"]);
            listObject.Name = tableName;
            dynamic queryTable = listObject.QueryTable;
            CanonicalDataLoader.ConfigureWorksheetConnection(queryTable, commandText);
            queryTable.RefreshStyle = 1;
            queryTable.AdjustColumnWidth = false;
            queryTable.PreserveFormatting = true;

            dynamic workbookConnection;
            try
            {
                workbookConnection = queryTable.WorkbookConnection;
                var connectionName = Convert.ToString(workbookConnection.Name, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(connectionName))
                {
                    throw new InvalidOperationException("Excel did not expose the audit connection name.");
                }

                CanonicalDataLoader.DemandExactWorksheetConnection(
                    queryTable,
                    connectionString,
                    commandText);
                registry.Register(workbook, connectionIdentity, connectionName!);
            }
            catch (Exception exception) when (!(exception is InvalidOperationException))
            {
                throw new InvalidOperationException(
                    "Excel did not expose the managed connection for the normalized-data audit.",
                    exception);
            }

            registry.Register(workbook, tableIdentity, tableName);
            cancellationToken.ThrowIfCancellationRequested();
            var backgroundRefresh = TryEnableBackgroundRefresh(queryTable);
            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Checking,
                Operation = backgroundRefresh
                    ? "Refreshing an independent audit across every normalized row."
                    : "Excel cannot refresh the independent audit in the background. The next operation may temporarily block Excel.",
                ManagedObject = tableName,
                ProjectedRows = projectedRows
            });
            var refreshAccepted = Convert.ToBoolean(
                queryTable.Refresh(backgroundRefresh),
                CultureInfo.InvariantCulture);
            if (!refreshAccepted)
            {
                throw new InvalidOperationException(
                    "Excel did not accept the independent normalized-data audit refresh.");
            }
            if (backgroundRefresh)
            {
                PollRefresh(
                    queryTable,
                    tableName,
                    projectedRows,
                    progressSink,
                    cancellationToken);
            }

            cancellationToken.ThrowIfCancellationRequested();
            var result = ReadResult(listObject, auditPlan);
            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Checking,
                Operation = "Independent normalized-data row count and totals are ready.",
                ManagedObject = tableName,
                ProjectedRows = projectedRows
            });
            return result;
        }

        internal void DeletePriorManagedConnection(
            dynamic workbook,
            ManagedObjectIdentity identity,
            string expectedConnectionString,
            string expectedCommandText)
        {
            var records = registry.Load((object)workbook).Where(item =>
                string.Equals(item.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(item.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                item.Kind == identity.Kind).ToList();
            if (records.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the prior normalized-data audit connection.");
            }

            var record = records.SingleOrDefault();
            if (record == null)
            {
                return;
            }

            dynamic? connection = TryGetConnection(workbook, record.ExcelName);
            if (connection == null)
            {
                registry.Remove((object)workbook, new[] { identity });
                return;
            }

            var actualName = Convert.ToString(
                connection.Name,
                CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.Equals(actualName, record.ExcelName, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The prior normalized-data audit connection no longer has its exact registered name.");
            }

            CanonicalDataLoader.DemandExactWorksheetWorkbookConnection(
                connection,
                expectedConnectionString,
                expectedCommandText);
            connection.Delete();
            registry.Remove((object)workbook, new[] { identity });
        }

        internal static CanonicalAuditResult ReadResult(
            dynamic listObject,
            CanonicalAuditQueryPlan plan)
        {
            var rowCount = Convert.ToInt32(listObject.ListRows.Count, CultureInfo.InvariantCulture);
            if (rowCount != 1)
            {
                throw new InvalidOperationException(
                    "The normalized-data audit did not return exactly one complete result row.");
            }

            var expectedColumnCount = plan.Fields.Count + 1;
            var actualColumnCount = Convert.ToInt32(listObject.ListColumns.Count, CultureInfo.InvariantCulture);
            if (actualColumnCount != expectedColumnCount)
            {
                throw new InvalidOperationException(
                    "The normalized-data audit result has an unexpected column count.");
            }

            dynamic rowColumn = listObject.ListColumns.Item(CanonicalAuditQueryPlan.RowCountColumnName);
            var rowsValue = ReadRequiredNumber(rowColumn.DataBodyRange.Cells[1, 1].Value2, "row count");
            if (rowsValue < 0m || rowsValue != decimal.Truncate(rowsValue) || rowsValue > long.MaxValue)
            {
                throw new InvalidOperationException("The normalized-data audit returned an invalid row count.");
            }

            var totals = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in plan.Fields)
            {
                dynamic resultColumn = listObject.ListColumns.Item(field.ResultColumnName);
                var total = ReadRequiredNumber(
                    resultColumn.DataBodyRange.Cells[1, 1].Value2,
                    "total for " + field.SourceField);
                foreach (var measureId in field.MeasureIds)
                {
                    totals.Add(measureId, total);
                }
            }

            return new CanonicalAuditResult
            {
                ActualRows = decimal.ToInt64(rowsValue),
                Totals = totals
            };
        }

        private static decimal ReadRequiredNumber(object? value, string label)
        {
            if (value == null || value is string || value is bool || value is DateTime)
            {
                throw new InvalidOperationException(
                    "The normalized-data audit returned a missing or non-numeric " + label + ".");
            }

            try
            {
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "The normalized-data audit returned a missing or non-numeric " + label + ".",
                    exception);
            }
        }

        private static void PollRefresh(
            dynamic queryTable,
            string managedObject,
            long projectedRows,
            IExcelProgressSink progressSink,
            CancellationToken cancellationToken)
        {
            var nextHeartbeat = DateTime.UtcNow.AddSeconds(5);
            while (Convert.ToBoolean(queryTable.Refreshing, CultureInfo.InvariantCulture))
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    queryTable.CancelRefresh();
                    cancellationToken.ThrowIfCancellationRequested();
                }

                if (DateTime.UtcNow >= nextHeartbeat)
                {
                    progressSink.Report(new ExcelProgress
                    {
                        Stage = ExcelBuildStage.Checking,
                        Operation = "Still auditing every normalized row independently of the report pivots.",
                        ManagedObject = managedObject,
                        ProjectedRows = projectedRows,
                        IsHeartbeat = true
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
                return Convert.ToBoolean(refreshable.BackgroundQuery, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static dynamic? FindTable(dynamic workbook, string tableName)
        {
            var sheetCount = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
            for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                dynamic sheet = workbook.Worksheets.Item(sheetIndex);
                try
                {
                    return sheet.ListObjects.Item(tableName);
                }
                catch (Exception)
                {
                }
            }

            return null;
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
