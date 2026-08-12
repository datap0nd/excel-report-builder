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
        private ClassicPortionUndoState? classicPortionUndo;

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
                if (context.Definition.Source.Kind == PivotSourceKind.WorksheetTable ||
                    context.Definition.Source.Kind == PivotSourceKind.WorksheetRange)
                {
                    return Task.FromResult(AddClassicParentPortion(
                        context,
                        valueFieldName,
                        detailFieldName,
                        measureCaption));
                }

                throw new NotSupportedException(
                    "Portion % is supported for worksheet and Data Model PivotTables only.");
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

        public async Task<PivotPlusPaneSnapshot> EnableDataModelAsync(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // A Power Query -> Data Model refresh needs Excel's main STA to
            // keep pumping OLE messages. Running the bounded COM transaction
            // from the click callback deadlocks Excel 2021 even when the
            // connection itself is configured for background refresh. COM
            // marshals these calls back to Excel while the UI thread remains
            // free to service the refresh.
            await Task.Run(
                () =>
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
                },
                cancellationToken).ConfigureAwait(true);

            cancellationToken.ThrowIfCancellationRequested();
            return ToSnapshot(discovery.Discover(application));
        }

        public Task<PivotPlusPaneSnapshot> UndoLastExtraAsync(
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (classicPortionUndo != null)
            {
                object classicPivot = classicPortionUndo.PivotTable;
                UndoClassicPortion(classicPortionUndo);
                classicPortionUndo = null;
                SelectPivotAnchor(classicPivot);
                return Task.FromResult(ToSnapshot(discovery.Discover(application)));
            }

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

        public Task<PivotPlusPaneSnapshot> GroupDateAsync(
            string fieldName,
            PivotDateGrouping grouping,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PivotTableContext context = discovery.Discover(application);
            if (context.Definition.Source.Kind != PivotSourceKind.WorksheetTable &&
                context.Definition.Source.Kind != PivotSourceKind.WorksheetRange)
            {
                throw new NotSupportedException(
                    "Date grouping is currently available for worksheet PivotTables.");
            }

            PivotFieldDescriptor field = DemandField(context, fieldName, "date field");
            _ = context.Definition.Placements.FirstOrDefault(item =>
                (item.Area == PivotFieldArea.Row || item.Area == PivotFieldArea.Column) &&
                string.Equals(item.FieldName, field.Name, StringComparison.OrdinalIgnoreCase)) ??
                throw new NotSupportedException(
                    "Drag the date field into Rows or Columns before grouping it.");

            dynamic pivot = ReadActivePivotTable();
            dynamic nativeField = pivot.PivotFields.Item(field.Name);
            dynamic cell = nativeField.DataRange.Cells.Item(1, 1);
            if (grouping == PivotDateGrouping.Ungrouped)
            {
                cell.Ungroup();
            }
            else
            {
                var periods = new object[]
                {
                    false, false, false, false,
                    grouping == PivotDateGrouping.Months,
                    grouping == PivotDateGrouping.Quarters,
                    grouping == PivotDateGrouping.Years
                };
                cell.Group(Type.Missing, Type.Missing, Type.Missing, periods);
            }

            SelectPivotAnchor((object)pivot);
            return Task.FromResult(ToSnapshot(discovery.Discover(application)));
        }

        private PivotPlusPaneSnapshot AddClassicParentPortion(
            PivotTableContext context,
            string valueFieldName,
            string detailFieldName,
            string measureCaption)
        {
            PivotFieldDescriptor value = DemandField(context, valueFieldName, "value field");
            PivotFieldDescriptor detail = DemandField(context, detailFieldName, "detail row field");
            PivotFieldPlacement detailPlacement = context.Definition.Placements.FirstOrDefault(item =>
                item.Area == PivotFieldArea.Row &&
                string.Equals(item.FieldName, detail.Name, StringComparison.OrdinalIgnoreCase)) ??
                throw new NotSupportedException(
                    "Choose a detail field that is currently placed in Rows.");
            PivotFieldPlacement valuePlacement = context.Definition.Placements.FirstOrDefault(item =>
                item.Area == PivotFieldArea.Values &&
                string.Equals(item.FieldName, value.Name, StringComparison.OrdinalIgnoreCase)) ??
                throw new NotSupportedException(
                    "Choose the numeric source field currently placed in Values.");

            string caption = string.IsNullOrWhiteSpace(measureCaption)
                ? "Portion %"
                : measureCaption.Trim();
            if (context.Definition.Placements.Any(item =>
                    item.Area == PivotFieldArea.Values &&
                    string.Equals(item.Caption, caption, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "A Value named '" + caption + "' already exists in this PivotTable.");
            }

            dynamic pivot = ReadActivePivotTable();
            if (context.Definition.Placements.Count(item =>
                    item.Area == PivotFieldArea.Values &&
                    string.Equals(item.FieldName, value.Name, StringComparison.OrdinalIgnoreCase)) != 1)
            {
                throw new NotSupportedException(
                    "Portion % currently requires the selected source field to appear exactly once in Values.");
            }

            DemandBlankExpansionSpace(
                pivot,
                context.Definition.Placements.Count(item =>
                    item.Area == PivotFieldArea.Values));
            dynamic sourceField = pivot.PivotFields.Item(value.Name);
            int priorValueCount = context.Definition.Placements.Count(item =>
                item.Area == PivotFieldArea.Values);
            object? priorDataPivot = null;
            int priorDataPivotOrientation = 0;
            int priorDataPivotPosition = 1;
            try
            {
                priorDataPivot = (object?)pivot.DataPivotField;
                if (priorDataPivot != null)
                {
                    dynamic dataPivot = priorDataPivot;
                    priorDataPivotOrientation = Convert.ToInt32(
                        dataPivot.Orientation,
                        CultureInfo.InvariantCulture);
                    priorDataPivotPosition = Math.Max(
                        1,
                        Convert.ToInt32(dataPivot.Position, CultureInfo.InvariantCulture));
                }
            }
            catch
            {
                priorDataPivot = null;
            }

            dynamic? added = null;
            try
            {
                // xlSum = -4157; xlPercentOfParentRow = 10; xlRowField = 1.
                added = pivot.AddDataField(sourceField, caption, -4157);
                added.Calculation = 10;
                added.NumberFormat = "0.0%";
                dynamic dataPivot = pivot.DataPivotField;
                dataPivot.Orientation = 1;
                dataPivot.Position = detailPlacement.Position + 1;

                if (Convert.ToInt32(added.Calculation, CultureInfo.InvariantCulture) != 10)
                {
                    throw new InvalidOperationException(
                        "Excel did not retain the Portion % parent-row calculation.");
                }

                classicPortionUndo = new ClassicPortionUndoState(
                    (object)pivot,
                    caption,
                    priorValueCount,
                    priorDataPivot,
                    priorDataPivotOrientation,
                    priorDataPivotPosition,
                    valuePlacement.Position);
                return ToSnapshot(discovery.Discover(application));
            }
            catch
            {
                if (added != null)
                {
                    try { added.Orientation = 0; } catch { }
                }

                throw;
            }
        }

        private static void UndoClassicPortion(ClassicPortionUndoState state)
        {
            dynamic pivot = state.PivotTable;
            dynamic dataField;
            try
            {
                dataField = pivot.DataFields.Item(state.DataFieldCaption);
            }
            catch
            {
                dataField = pivot.DataFields.Item(pivot.DataFields.Count);
            }
            // XlPivotFieldOrientation.xlHidden is 0 (not -1).
            dataField.Orientation = 0;

            if (state.PriorValueCount > 1 && state.PriorDataPivotField != null)
            {
                dynamic dataPivot = pivot.DataPivotField;
                dataPivot.Orientation = state.PriorDataPivotOrientation;
                dataPivot.Position = state.PriorDataPivotPosition;
            }
        }

        private static void DemandBlankExpansionSpace(dynamic pivot, int priorValueCount)
        {
            dynamic range = pivot.TableRange2;
            int bottom = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture) +
                         Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture) - 1;
            int left = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
            int width = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            int currentRows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            int reserveRows = Math.Max(8, currentRows * Math.Max(2, priorValueCount + 1));
            dynamic sheet = pivot.Parent;
            dynamic reserve = sheet.Range(
                sheet.Cells.Item(bottom + 1, left),
                sheet.Cells.Item(bottom + reserveRows, left + width - 1));
            dynamic app = sheet.Application;
            double occupied = Convert.ToDouble(
                app.WorksheetFunction.CountA(reserve),
                CultureInfo.InvariantCulture);
            if (occupied > 0)
            {
                throw new InvalidOperationException(
                    "Portion % needs blank rows below the PivotTable because Values are shown vertically. Move or clear the content directly below the PivotTable, then try again.");
            }
        }

        private void SelectPivotAnchor(object pivotTable)
        {
            dynamic pivot = pivotTable;
            dynamic app = application;
            app.Goto(pivot.TableRange2.Cells.Item(1, 1), false);
        }

        private sealed class ClassicPortionUndoState
        {
            internal ClassicPortionUndoState(
                object pivotTable,
                string dataFieldCaption,
                int priorValueCount,
                object? priorDataPivotField,
                int priorDataPivotOrientation,
                int priorDataPivotPosition,
                int priorValuePosition)
            {
                PivotTable = pivotTable;
                DataFieldCaption = dataFieldCaption;
                PriorValueCount = priorValueCount;
                PriorDataPivotField = priorDataPivotField;
                PriorDataPivotOrientation = priorDataPivotOrientation;
                PriorDataPivotPosition = priorDataPivotPosition;
                PriorValuePosition = priorValuePosition;
            }

            internal object PivotTable { get; }
            internal string DataFieldCaption { get; }
            internal int PriorValueCount { get; }
            internal object? PriorDataPivotField { get; }
            internal int PriorDataPivotOrientation { get; }
            internal int PriorDataPivotPosition { get; }
            internal int PriorValuePosition { get; }
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
                    FriendlyCaption(item.Caption ?? item.Name),
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

        private static string FriendlyCaption(string value)
        {
            string candidate = value ?? string.Empty;
            if (candidate.Length > 5 &&
                candidate.EndsWith("Field", StringComparison.OrdinalIgnoreCase) &&
                candidate.IndexOfAny(new[] { ' ', '[', ']' }) < 0)
            {
                return candidate.Substring(0, candidate.Length - 5);
            }

            return candidate;
        }
    }
}
