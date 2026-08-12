using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Excel.PivotPlus;
using ExcelReportBuilder.Excel.PivotPlus.Measures;
using ExcelReportBuilder.Excel.PivotPlus.Native;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;
using ExcelReportBuilder.Excel.PivotPlus.DataModel;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// UI-facing, bounded PivotTable+ host. It exposes native field placement
    /// and one typed Portion preset; it never accepts formulas, arbitrary COM,
    /// file paths, or workbook save operations from the pane.
    /// </summary>
    public sealed class ExcelPivotPlusHostService : IPivotPlusHostService
    {
        private readonly object application;
        private readonly PivotTableContextDiscovery discovery =
            new PivotTableContextDiscovery();
        private readonly PivotTableNativeLayoutMutationService nativeLayout =
            new PivotTableNativeLayoutMutationService();
        private readonly PivotModelMeasureMutationService measures =
            new PivotModelMeasureMutationService();
        private readonly PivotDataModelEnablementService enablement =
            new PivotDataModelEnablementService();
        private readonly PivotPlusWorkbookMetadataStore metadata =
            new PivotPlusWorkbookMetadataStore();

        public ExcelPivotPlusHostService(object excelApplication)
        {
            application = excelApplication ??
                throw new ArgumentNullException(nameof(excelApplication));
        }

        public Task<PivotPlusPaneSnapshot> InspectAsync(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            return Task.FromResult(ToSnapshot(context));
        }

        public Task<PivotPlusPaneSnapshot> ApplyLayoutAsync(
            IReadOnlyList<PivotPlusPlacementRequest> placements,
            CancellationToken cancellationToken)
        {
            if (placements == null) throw new ArgumentNullException(nameof(placements));
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            object pivot = ReadActivePivotTable();
            var requested = placements.Select(item => new PivotFieldPlacement(
                item.FieldName,
                item.Area,
                item.Position,
                string.IsNullOrWhiteSpace(item.Caption) ? null : item.Caption,
                item.Area == PivotFieldArea.Values
                    ? item.Aggregation ?? PivotAggregationFunction.Sum
                    : (PivotAggregationFunction?)null,
                item.Area == PivotFieldArea.Values &&
                !string.IsNullOrWhiteSpace(item.NumberFormatCode)
                    ? item.NumberFormatCode
                    : null,
                item.Area == PivotFieldArea.Row
                    ? PivotSubtotalMode.Automatic
                    : PivotSubtotalMode.None)).ToList();
            PivotValuesAxis valuesAxis = context.Definition.Layout.ValuesAxis;
            int valuesPosition = context.Definition.Layout.ValuesPosition;
            int valueCount = requested.Count(item => item.Area == PivotFieldArea.Values);
            if (valueCount <= 1)
            {
                valuesAxis = PivotValuesAxis.Automatic;
                valuesPosition = 1;
            }
            else if (valuesAxis == PivotValuesAxis.Automatic)
            {
                valuesAxis = PivotValuesAxis.Columns;
                valuesPosition = requested.Count(item =>
                    item.Area == PivotFieldArea.Column) + 1;
            }

            var definition = new PivotLayoutDefinition(
                context.Definition.Target,
                context.Definition.Source,
                context.Definition.Fields,
                requested,
                filters: Array.Empty<PivotFieldFilter>(),
                layout: new PivotLayoutMetadata(
                    context.Definition.Layout.Form,
                    context.Definition.Layout.RepeatItemLabels,
                    context.Definition.Layout.ShowRowGrandTotals,
                    context.Definition.Layout.ShowColumnGrandTotals,
                    context.Definition.Layout.ShowFieldHeaders,
                    valuesAxis,
                    valuesPosition),
                format: context.Definition.Format,
                capabilityRequirements: context.Definition.CapabilityRequirements,
                clearAll: requested.Count == 0);
            nativeLayout.Apply(pivot, context, definition);
            return Task.FromResult(ToSnapshot(discovery.Discover(application)));
        }

        public Task<PivotPlusPaneSnapshot> AddParentPortionAsync(
            string valueFieldName,
            string detailFieldName,
            string measureCaption,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            if (context.Definition.Source.Kind != PivotSourceKind.DataModel)
            {
                throw new NotSupportedException(
                    "Portion requires a Data Model PivotTable. Use Enable PivotTable+ first.");
            }

            PivotFieldDescriptor value = DemandField(
                context,
                valueFieldName,
                "value field");
            PivotFieldDescriptor detail = DemandField(
                context,
                detailFieldName,
                "detail row field");
            if (value.IsMeasure)
            {
                throw new NotSupportedException(
                    "Choose the numeric source column behind the current Value, not an existing measure.");
            }

            string caption = string.IsNullOrWhiteSpace(measureCaption)
                ? "Portion %"
                : measureCaption.Trim();
            PivotModelSchema schema = CreateModelSchema(
                context,
                value.Name,
                detail.Name,
                out string valueId,
                out string detailId,
                out string homeTableId);
            var definition = new PivotMeasureSetDefinition(
                schema,
                new[]
                {
                    new PivotMeasureDefinition(
                        "portion",
                        caption,
                        homeTableId,
                        new PivotMeasureFormat(
                            PivotMeasureFormatKind.Percentage,
                            decimalPlaces: 1,
                            useThousandsSeparator: false),
                        new PivotShareExpression(
                            new PivotAggregateExpression(
                                valueId,
                                PivotCalculationAggregateFunction.Sum),
                            new PivotParentShareDenominator(new[] { detailId }),
                            PivotDenominatorBehavior.Blank))
                });
            PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

            IReadOnlyList<PivotFieldPlacement> currentValues = context.Definition.Placements
                .Where(item => item.Area == PivotFieldArea.Values)
                .OrderBy(item => item.Position)
                .ToList();
            var finalValues = new List<PivotMeasureValuePlacement>();
            foreach (PivotFieldPlacement placement in currentValues)
            {
                PivotFieldDescriptor descriptor = context.Definition.Fields.First(field =>
                    string.Equals(
                        field.Name,
                        placement.FieldName,
                        StringComparison.OrdinalIgnoreCase));
                string currentCaption = placement.Caption ??
                    descriptor.Caption ?? descriptor.Name;
                string numberFormat = placement.NumberFormatCode ?? "General";
                finalValues.Add(new PivotMeasureValuePlacement(
                    finalValues.Count + 1,
                    new PivotExistingDataFieldIdentity(
                        placement.FieldName,
                        PivotMeasurePlacementFingerprint.CreateCaptionFingerprint(
                            currentCaption),
                        PivotMeasurePlacementFingerprint.CreateNumberFormatFingerprint(
                            numberFormat),
                        placement.Position)));
            }

            finalValues.Add(new PivotMeasureValuePlacement(
                finalValues.Count + 1,
                "portion"));
            PivotValuesAxis axis = finalValues.Count > 1
                ? PivotValuesAxis.Rows
                : PivotValuesAxis.Automatic;
            int axisPosition = finalValues.Count > 1
                ? context.Definition.Placements.Count(item =>
                    item.Area == PivotFieldArea.Row) + 1
                : 1;
            object pivot = ReadActivePivotTable();
            object workbook = ReadPivotWorkbook(pivot);
            string setupId = ResolveSetupId(workbook, context);
            measures.Apply(
                workbook,
                pivot,
                context,
                setupId,
                compilation,
                new PivotMeasurePlacementPlan(finalValues, axis, axisPosition));
            return Task.FromResult(ToSnapshot(discovery.Discover(application)));
        }

        public Task<PivotPlusPaneSnapshot> EnableDataModelAsync(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            object pivot = ReadActivePivotTable();
            object workbook = ReadPivotWorkbook(pivot);
            enablement.Enable(
                workbook,
                pivot,
                context,
                ResolveSetupId(workbook, context));
            return Task.FromResult(ToSnapshot(discovery.Discover(application)));
        }

        public Task<PivotPlusPaneSnapshot> UndoLastExtraAsync(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            object pivot = ReadActivePivotTable();
            object workbook = ReadPivotWorkbook(pivot);
            measures.Undo(
                workbook,
                pivot,
                context,
                ResolveSetupId(workbook, context));
            return Task.FromResult(ToSnapshot(discovery.Discover(application)));
        }

        public void OpenExcelFieldList()
        {
            dynamic app = application;
            app.CommandBars.ExecuteMso("PivotFieldListShowHide");
        }

        private object ReadActivePivotTable()
        {
            dynamic app = application;
            object? pivot = app.ActiveCell == null ? null : (object?)app.ActiveCell.PivotTable;
            return pivot ?? throw new InvalidOperationException(
                "Select a cell inside a PivotTable first.");
        }

        private static object ReadPivotWorkbook(object pivotTable)
        {
            dynamic pivot = pivotTable;
            object? worksheet = (object?)pivot.Parent;
            dynamic sheet = worksheet ?? throw new InvalidOperationException(
                "Excel did not expose the PivotTable worksheet.");
            return (object?)sheet.Parent ?? throw new InvalidOperationException(
                "Excel did not expose the PivotTable workbook.");
        }

        private string ResolveSetupId(object workbook, PivotTableContext context)
        {
            PivotPlusWorkbookMetadata? current = metadata.LoadForTarget(
                (dynamic)workbook,
                context.Definition.Target.WorksheetName,
                context.Definition.Target.PivotTableName);
            if (current != null) return current.SetupId;
            string hash = PivotPlusFingerprint.Create(
                "pane.setup.v1",
                context.Definition.Target.WorksheetName + "\u001f" +
                context.Definition.Target.PivotTableName);
            return "pane_" + hash.Substring(hash.Length - 16);
        }

        private static PivotFieldDescriptor DemandField(
            PivotTableContext context,
            string fieldName,
            string label)
        {
            PivotFieldDescriptor? field = context.Definition.Fields.SingleOrDefault(item =>
                string.Equals(item.Name, fieldName, StringComparison.OrdinalIgnoreCase));
            return field ?? throw new ArgumentException(
                "The selected " + label + " is no longer present in the PivotTable schema.");
        }

        private static PivotModelSchema CreateModelSchema(
            PivotTableContext context,
            string valueFieldName,
            string detailFieldName,
            out string valueId,
            out string detailId,
            out string homeTableId)
        {
            var ids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var tables = new List<PivotModelTableSchema>();
            int fieldIndex = 0;
            int tableIndex = 0;
            foreach (IGrouping<string, PivotFieldDescriptor> group in
                     context.Definition.Fields.Where(item => !item.IsMeasure)
                         .GroupBy(item =>
                             string.IsNullOrWhiteSpace(item.TableName)
                                 ? context.Definition.Source.ModelTableName ?? "Model"
                                 : item.TableName!,
                             StringComparer.OrdinalIgnoreCase))
            {
                string tableId = "table_" + (++tableIndex).ToString(
                    CultureInfo.InvariantCulture);
                var fields = new List<PivotModelFieldSchema>();
                foreach (PivotFieldDescriptor field in group)
                {
                    string id = "field_" + (++fieldIndex).ToString(
                        CultureInfo.InvariantCulture);
                    ids[field.Name] = id;
                    fields.Add(new PivotModelFieldSchema(
                        id,
                        LastUniqueNameSegment(field.Name),
                        string.Equals(
                            field.Name,
                            valueFieldName,
                            StringComparison.OrdinalIgnoreCase)
                            ? PivotModelDataType.DecimalNumber
                            : PivotModelDataType.Unknown));
                }

                tables.Add(new PivotModelTableSchema(
                    tableId,
                    NormalizeTableName(group.Key),
                    fields));
            }

            if (!ids.TryGetValue(valueFieldName, out valueId!) ||
                !ids.TryGetValue(detailFieldName, out detailId!))
            {
                throw new InvalidOperationException(
                    "The Portion fields could not be bound to Data Model columns.");
            }

            string selectedValueId = valueId;
            homeTableId = tables.Single(table => table.Fields.Any(field =>
                string.Equals(field.Id, selectedValueId, StringComparison.Ordinal))).Id;
            return new PivotModelSchema(tables);
        }

        private static string NormalizeTableName(string value)
        {
            string result = value.Trim();
            if (result.Length >= 2 && result[0] == '[' && result[result.Length - 1] == ']')
            {
                result = result.Substring(1, result.Length - 2).Replace("]]", "]");
            }

            return result;
        }

        private static string LastUniqueNameSegment(string value)
        {
            int start = value.LastIndexOf(".[", StringComparison.Ordinal);
            if (start >= 0 && value.EndsWith("]", StringComparison.Ordinal))
            {
                return value.Substring(start + 2, value.Length - start - 3)
                    .Replace("]]", "]");
            }

            return NormalizeTableName(value);
        }

        private static PivotPlusPaneSnapshot ToSnapshot(PivotTableContext context)
        {
            IReadOnlyList<PivotPlusFieldSnapshot> fields = context.Definition.Fields
                .Select(item => new PivotPlusFieldSnapshot(
                    item.Name,
                    item.Caption ?? item.Name,
                    item.IsMeasure,
                    item.IsCalculated,
                    item.SupportedAreas))
                .OrderBy(item => item.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            IReadOnlyList<PivotPlusPlacementSnapshot> placements =
                context.Definition.Placements.Select(item =>
                {
                    PivotFieldDescriptor? field = context.Definition.Fields.FirstOrDefault(
                        candidate => string.Equals(
                            candidate.Name,
                            item.FieldName,
                            StringComparison.OrdinalIgnoreCase));
                    return new PivotPlusPlacementSnapshot(
                        item.FieldName,
                        item.Caption ?? field?.Caption ?? item.FieldName,
                        item.Area,
                        item.Position,
                        item.Aggregation,
                        item.NumberFormatCode ?? string.Empty);
                }).ToList();
            return new PivotPlusPaneSnapshot(
                context.Definition.Target.WorksheetName,
                context.Definition.Target.PivotTableName,
                context.Definition.Source.Kind,
                fields,
                placements);
        }
    }
}
