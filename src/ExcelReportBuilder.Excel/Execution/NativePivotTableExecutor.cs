using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Ownership;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class PivotDataFieldDescriptor
    {
        public string MeasureId { get; set; } = string.Empty;

        public string ComponentId { get; set; } = string.Empty;

        public AggregateComponentRole Role { get; set; }

        public string SourceField { get; set; } = string.Empty;

        public AggregateFunction Function { get; set; }

        public string PivotCaption { get; set; } = string.Empty;

        public IReadOnlyList<MeasureFilterSpec> Filters { get; set; } = Array.Empty<MeasureFilterSpec>();

        public string? PeriodSliceId { get; set; }
    }

    public sealed class PivotBuildResult
    {
        public string PivotTableName { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;

        public string AnchorCell { get; set; } = string.Empty;

        public IReadOnlyList<PivotDataFieldDescriptor> DataFields { get; set; } = Array.Empty<PivotDataFieldDescriptor>();
    }

    /// <summary>
    /// Builds native PivotCaches and PivotTables through late-bound Excel. It
    /// accepts only the typed compiler plan and never a formula or VBA string.
    /// </summary>
    public sealed class NativePivotTableExecutor
    {
        private const int SourceDatabase = 1;
        private const int SourceExternal = 2;
        private const int OrientationRow = 1;
        private const int OrientationColumn = 2;
        private const int OrientationPage = 3;
        private const int OrientationData = 4;
        private const int SortAscending = 1;
        private const int SortDescending = 2;
        private const int PivotVersion15 = 6;
        private readonly WorkbookOwnershipRegistry registry;
        private readonly ManagedOwnershipGuard ownershipGuard;

        public NativePivotTableExecutor(
            WorkbookOwnershipRegistry? registry = null,
            ManagedOwnershipGuard? ownershipGuard = null)
        {
            this.registry = registry ?? new WorkbookOwnershipRegistry();
            this.ownershipGuard = ownershipGuard ?? new ManagedOwnershipGuard();
        }

        public PivotBuildResult Build(
            dynamic workbook,
            dynamic destinationWorksheet,
            string destinationCell,
            string reportId,
            DenseReportBlockPlan block,
            CanonicalLoadPlan source,
            IExcelProgressSink? progressSink = null)
        {
            if (block == null)
            {
                throw new ArgumentNullException(nameof(block));
            }

            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            var destinationIdentity = block.OutputMode == ReportOutputMode.DenseGrid
                ? new ManagedObjectIdentity(
                    reportId,
                    block.OwnershipId + "_pivot_sheet",
                    ManagedObjectKind.PivotTable)
                : ManagedOutputIdentity.Draft(reportId, block.WorksheetName);
            ownershipGuard.DemandOwned(destinationWorksheet, destinationIdentity);
            var identity = new ManagedObjectIdentity(reportId, block.OwnershipId, ManagedObjectKind.PivotTable);
            RemoveExistingOwnedPivot(workbook, destinationWorksheet, block.Pivot.ManagedPivotName, identity);
            ClearOwnedDestination(destinationWorksheet, destinationCell, block);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.BuildingPivots,
                Operation = "Creating native PivotTable " + block.Pivot.ManagedPivotName + ".",
                ManagedObject = block.Pivot.ManagedPivotName,
                ProjectedRows = source.ProjectedRows
            });

            dynamic sourceData = source.Backend == CanonicalBackend.Worksheet
                ? (dynamic)source.TableOrConnectionName
                : workbook.Connections.Item(source.TableOrConnectionName);
            var sourceType = source.Backend == CanonicalBackend.Worksheet ? SourceDatabase : SourceExternal;
            dynamic cache = workbook.PivotCaches().Create(sourceType, sourceData, PivotVersion15);
            dynamic destination = destinationWorksheet.Range[destinationCell];
            dynamic pivot = cache.CreatePivotTable(destination, block.Pivot.ManagedPivotName);
            pivot.ManualUpdate = true;

            try
            {
                var densePostProcessing = block.OutputMode == ReportOutputMode.DenseGrid;
                ConfigureFields(
                    pivot,
                    block.Pivot.Rows,
                    OrientationRow,
                    source.Backend,
                    source.QueryName,
                    densePostProcessing);
                ConfigureFields(
                    pivot,
                    block.Pivot.Columns,
                    OrientationColumn,
                    source.Backend,
                    source.QueryName,
                    densePostProcessing);
                ConfigureFilters(pivot, block.Pivot.Filters, source.Backend, source.QueryName);
                var dataFields = ConfigureValues(
                    pivot,
                    block.Pivot.Values.Concat(block.Pivot.SupportingValues).ToList(),
                    source.Backend,
                    source.QueryName,
                    block.OutputMode == ReportOutputMode.DenseGrid);
                ConfigureTopN(pivot, block.Pivot.Rows, dataFields, source.Backend, source.QueryName, densePostProcessing);
                ConfigureTopN(pivot, block.Pivot.Columns, dataFields, source.Backend, source.QueryName, densePostProcessing);
                pivot.RowGrand = block.Pivot.GrandTotals.ShowRows;
                pivot.ColumnGrand = block.Pivot.GrandTotals.ShowColumns;
                if (!string.IsNullOrWhiteSpace(block.Pivot.GrandTotals.RowLabel))
                {
                    pivot.GrandTotalName = block.Pivot.GrandTotals.RowLabel;
                }
                pivot.DisplayFieldCaptions = true;
                pivot.ShowTableStyleRowStripes = false;
                pivot.ShowTableStyleColumnStripes = false;
                pivot.InGridDropZones = false;
                pivot.RowAxisLayout(2);

                pivot.ManualUpdate = false;
                pivot.RefreshTable();
                dynamic renderedRange = pivot.TableRange2;
                var renderedRows = Convert.ToInt32(renderedRange.Rows.Count, CultureInfo.InvariantCulture);
                var renderedColumns = Convert.ToInt32(renderedRange.Columns.Count, CultureInfo.InvariantCulture);
                try
                {
                    DemandDimensionsWithinOwnedRange(block, renderedRows, renderedColumns);
                }
                catch (InvalidOperationException)
                {
                    renderedRange.Clear();
                    throw;
                }

                registry.Register(workbook, identity, block.Pivot.ManagedPivotName);
                return new PivotBuildResult
                {
                    PivotTableName = block.Pivot.ManagedPivotName,
                    WorksheetName = Convert.ToString(destinationWorksheet.Name, CultureInfo.InvariantCulture) ?? string.Empty,
                    AnchorCell = destinationCell,
                    DataFields = dataFields
                };
            }
            catch (Exception)
            {
                pivot.ManualUpdate = false;
                throw;
            }
        }

        internal static void DemandDimensionsWithinOwnedRange(
            DenseReportBlockPlan block,
            int renderedRows,
            int renderedColumns)
        {
            if (block == null) throw new ArgumentNullException(nameof(block));
            if (renderedRows < 1 || renderedColumns < 1)
            {
                throw new InvalidOperationException("The managed PivotTable produced an invalid extent.");
            }

            var maximumRows = block.OutputMode == ReportOutputMode.DenseGrid
                ? 1_048_574
                : block.OwnedRange.RowCount;
            var maximumColumns = block.OutputMode == ReportOutputMode.DenseGrid
                ? 16_384
                : block.OwnedRange.ColumnCount;
            if (renderedRows > maximumRows || renderedColumns > maximumColumns)
            {
                throw new InvalidOperationException(
                    block.OutputMode == ReportOutputMode.DenseGrid
                        ? "The managed hidden PivotTable exceeds Excel's worksheet capacity and was removed."
                        : "The managed PivotTable exceeds its validated owned extent and was removed.");
            }
        }

        private static void ClearOwnedDestination(
            dynamic worksheet,
            string destinationCell,
            DenseReportBlockPlan block)
        {
            var anchor = CellAddress.Parse(destinationCell);
            dynamic range = worksheet.Range[
                worksheet.Cells[anchor.Row, anchor.Column],
                worksheet.Cells[
                    anchor.Row + block.OwnedRange.RowCount - 1,
                    anchor.Column + block.OwnedRange.ColumnCount - 1]];
            range.Clear();
        }

        private static void ConfigureFields(
            dynamic pivot,
            IReadOnlyList<PivotFieldPlan> fields,
            int orientation,
            CanonicalBackend backend,
            string modelTableName,
            bool densePostProcessing)
        {
            for (var index = 0; index < fields.Count; index++)
            {
                var plan = fields[index];
                dynamic field;
                if (backend == CanonicalBackend.DataModel)
                {
                    dynamic cubeField = ResolveCubeField(pivot, modelTableName, plan.Field);
                    cubeField.Orientation = orientation;
                    cubeField.Position = index + 1;
                    field = ResolveVisiblePivotField(pivot, cubeField, plan.Field);
                }
                else
                {
                    field = pivot.PivotFields(plan.Field);
                    field.Orientation = orientation;
                    field.Position = index + 1;
                }
                ConfigureSubtotals(
                    field,
                    densePostProcessing
                        ? new SubtotalSpec { Mode = SubtotalMode.None }
                        : plan.Subtotals);
                if (!string.IsNullOrWhiteSpace(plan.Caption))
                {
                    field.Caption = plan.Caption;
                }

                if (orientation == OrientationRow)
                {
                    field.LayoutForm = 1;
                    field.LayoutSubtotalLocation = plan.Subtotals.Placement == TotalPlacement.BeforeMembers ? 1 : 2;
                    if (!string.IsNullOrWhiteSpace(plan.Subtotals.Label))
                    {
                        field.SubtotalName = plan.Subtotals.Label;
                    }
                }

                ApplyMemberStages(field, plan, densePostProcessing);
            }
        }

        private static void ApplyMemberStages(
            dynamic field,
            PivotFieldPlan plan,
            bool densePostProcessing)
        {
            foreach (var stage in plan.MemberStages)
            {
                switch (stage)
                {
                    case PivotMemberStageKind.ApplyMemberOrder:
                        for (var itemIndex = 0; itemIndex < plan.MemberOrder.Count; itemIndex++)
                        {
                            dynamic item = ResolvePivotItem(field, plan.MemberOrder[itemIndex]);
                            item.Position = itemIndex + 1;
                        }

                        break;
                    case PivotMemberStageKind.SortAscending:
                        field.AutoSort(SortAscending, plan.Field);
                        break;
                    case PivotMemberStageKind.SortDescending:
                        field.AutoSort(SortDescending, plan.Field);
                        break;
                    case PivotMemberStageKind.GroupMembers:
                        if (!densePostProcessing)
                        {
                            throw new NotSupportedException(
                                "Layout grouping buckets require the dense-layout member executor.");
                        }

                        break;
                    case PivotMemberStageKind.ApplyTopN:
                        // A Top N stage depends on a configured DataField and is
                        // applied after values are created.
                        break;
                    case PivotMemberStageKind.AggregateOthers:
                        if (!densePostProcessing)
                        {
                            throw new NotSupportedException(
                                "An explicit Others member requires the dense-layout remainder executor.");
                        }

                        break;
                    default:
                        throw new NotSupportedException("The PivotTable member stage is not supported.");
                }
            }
        }

        private static void ConfigureTopN(
            dynamic pivot,
            IReadOnlyList<PivotFieldPlan> fields,
            IReadOnlyList<PivotDataFieldDescriptor> dataFields,
            CanonicalBackend backend,
            string modelTableName,
            bool densePostProcessing)
        {
            if (densePostProcessing)
            {
                return;
            }

            foreach (var plan in fields)
            {
                if (plan.TopN == null || !plan.MemberStages.Contains(PivotMemberStageKind.ApplyTopN))
                {
                    continue;
                }

                if (plan.TopN.IncludeOthers)
                {
                    throw new NotSupportedException(
                        "An explicit Others row requires the dense-layout remainder executor and cannot be represented as a native PivotTable filter.");
                }

                var descriptor = dataFields.FirstOrDefault(candidate =>
                    string.Equals(candidate.MeasureId, plan.TopN.MeasureId, StringComparison.OrdinalIgnoreCase));
                if (descriptor == null)
                {
                    throw new InvalidOperationException("The Top N measure does not have a native aggregate component.");
                }

                dynamic field;
                if (backend == CanonicalBackend.DataModel)
                {
                    dynamic cubeField = ResolveCubeField(pivot, modelTableName, plan.Field);
                    field = ResolveVisiblePivotField(pivot, cubeField, plan.Field);
                }
                else
                {
                    field = pivot.PivotFields(plan.Field);
                }
                dynamic dataField = pivot.DataFields.Item(descriptor.PivotCaption);
                var filterType = plan.TopN.Direction == TopNDirection.Top ? 1 : 2;
                field.ClearValueFilters();
                field.PivotFilters.Add(filterType, dataField, plan.TopN.Count);
            }
        }

        private static void ConfigureSubtotals(dynamic field, SubtotalSpec subtotal)
        {
            var values = new object[12];
            for (var index = 0; index < values.Length; index++)
            {
                values[index] = false;
            }

            if (subtotal.Mode == SubtotalMode.Automatic)
            {
                values[0] = true;
            }

            field.Subtotals = values;
        }

        private static void ConfigureFilters(
            dynamic pivot,
            IReadOnlyList<PivotFilterPlan> filters,
            CanonicalBackend backend,
            string modelTableName)
        {
            for (var index = 0; index < filters.Count; index++)
            {
                var plan = filters[index];
                dynamic field;
                if (backend == CanonicalBackend.DataModel)
                {
                    dynamic cubeField = ResolveCubeField(pivot, modelTableName, plan.Field);
                    cubeField.Orientation = OrientationPage;
                    cubeField.Position = index + 1;
                    field = ResolveVisiblePivotField(pivot, cubeField, plan.Field);
                }
                else
                {
                    field = pivot.PivotFields(plan.Field);
                    field.Orientation = OrientationPage;
                    field.Position = index + 1;
                }
                var selectedValues = new List<ScalarValue>(plan.SelectedValues);
                if (plan.IncludeBlank && !selectedValues.Any(value => value.Kind == ScalarValueKind.Null))
                {
                    selectedValues.Add(ScalarValue.Null());
                }

                if (selectedValues.Count == 1)
                {
                    var itemName = ResolvePivotItemName(field, selectedValues[0]);
                    if (backend == CanonicalBackend.DataModel)
                    {
                        field.CurrentPageName = itemName;
                    }
                    else
                    {
                        field.CurrentPage = itemName;
                    }
                }
                else if (selectedValues.Count > 1)
                {
                    field.EnableMultiplePageItems = true;
                    var itemNames = selectedValues
                        .Select(value => ResolvePivotItemName(field, value))
                        .ToArray();
                    if (backend == CanonicalBackend.DataModel)
                    {
                        field.VisibleItemsList = itemNames;
                    }
                    else
                    {
                        field.CurrentPageList = itemNames;
                    }
                }
            }
        }

        private static dynamic ResolvePivotItem(dynamic field, ScalarValue value)
        {
            var caption = ScalarToPivotCaption(value);
            try
            {
                return field.PivotItems(caption);
            }
            catch (Exception)
            {
                dynamic items = field.PivotItems();
                var count = Convert.ToInt32(items.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = items.Item(index);
                    if (PivotItemMatches(item, value, caption))
                    {
                        return item;
                    }
                }

                throw new InvalidOperationException("A configured filter member was not found in the managed PivotTable.");
            }
        }

        private static string ResolvePivotItemName(dynamic field, ScalarValue value)
        {
            dynamic item = ResolvePivotItem(field, value);
            return Convert.ToString(item.Name, CultureInfo.InvariantCulture) ?? ScalarToPivotCaption(value);
        }

        private static bool PivotItemMatches(dynamic item, ScalarValue value, string fallbackCaption)
        {
            object? raw = null;
            try
            {
                raw = item.Value;
            }
            catch (Exception)
            {
            }

            switch (value.Kind)
            {
                case ScalarValueKind.Null:
                    return raw == null || string.Equals(
                        Convert.ToString(item.Caption, CultureInfo.InvariantCulture),
                        "(blank)",
                        StringComparison.OrdinalIgnoreCase);
                case ScalarValueKind.Number:
                    try
                    {
                        return Convert.ToDecimal(raw, CultureInfo.InvariantCulture) == value.Number;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                case ScalarValueKind.Boolean:
                    try
                    {
                        return Convert.ToBoolean(raw, CultureInfo.InvariantCulture) == value.Boolean;
                    }
                    catch (Exception)
                    {
                        break;
                    }
                case ScalarValueKind.Date:
                case ScalarValueKind.DateTime:
                    if (TryConvertPivotDate(raw, out var date) && value.Temporal.HasValue)
                    {
                        return value.Kind == ScalarValueKind.Date
                            ? date.Date == value.Temporal.Value.Date
                            : date == value.Temporal.Value;
                    }

                    break;
            }

            return string.Equals(
                Convert.ToString(item.Caption, CultureInfo.InvariantCulture),
                fallbackCaption,
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryConvertPivotDate(object? value, out DateTime result)
        {
            if (value is DateTime date)
            {
                result = date;
                return true;
            }

            if (value is double serial)
            {
                try
                {
                    result = DateTime.FromOADate(serial);
                    return true;
                }
                catch (ArgumentException)
                {
                }
            }

            return DateTime.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out result);
        }

        private static IReadOnlyList<PivotDataFieldDescriptor> ConfigureValues(
            dynamic pivot,
            IReadOnlyList<PivotValuePlan> values,
            CanonicalBackend backend,
            string modelTableName,
            bool deduplicateComponents)
        {
            var descriptors = new List<PivotDataFieldDescriptor>();
            var captions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in values)
            {
                for (var index = 0; index < value.AggregateComponents.Count; index++)
                {
                    var component = value.AggregateComponents[index];
                    var leaf = new PivotDataFieldDescriptor
                    {
                        MeasureId = value.MeasureId,
                        ComponentId = component.Id,
                        Role = component.Role,
                        SourceField = component.Field,
                        Function = component.Function,
                        Filters = component.Filters,
                        PeriodSliceId = component.PeriodSliceId
                    };
                    var key = (deduplicateComponents ? string.Empty : value.MeasureId + "|") +
                              leaf.SourceField + "|" + leaf.Function + "|" +
                              FilterKey(leaf.Filters) + "|" + leaf.PeriodSliceId;
                    if (captions.TryGetValue(key, out var existingCaption))
                    {
                        leaf.PivotCaption = existingCaption;
                        descriptors.Add(leaf);
                        continue;
                    }

                    if (leaf.Function == AggregateFunction.DistinctCount &&
                        backend != CanonicalBackend.DataModel)
                    {
                        throw new NotSupportedException(
                            "Distinct count requires a Data Model-backed managed PivotTable.");
                    }

                    var caption = value.AggregateComponents.Count == 1 && !value.RequiresPostAggregationCalculation
                        ? value.Label
                        : "ERB " + value.MeasureId + " " + (index + 1).ToString(CultureInfo.InvariantCulture);
                    dynamic dataField;
                    if (backend == CanonicalBackend.DataModel)
                    {
                        dynamic sourceField = ResolveCubeField(pivot, modelTableName, leaf.SourceField);
                        dynamic measure = pivot.CubeFields.GetMeasure(
                            sourceField,
                            ConsolidationFunction(leaf.Function),
                            caption);
                        measure.Orientation = OrientationData;
                        dataField = pivot.DataFields.Item(caption);
                    }
                    else
                    {
                        dynamic sourceField = pivot.PivotFields(leaf.SourceField);
                        dataField = pivot.AddDataField(sourceField, caption, ConsolidationFunction(leaf.Function));
                    }
                    if (!string.IsNullOrWhiteSpace(value.NumberFormat))
                    {
                        dataField.NumberFormat = value.NumberFormat;
                    }

                    leaf.PivotCaption = caption;
                    captions[key] = caption;
                    descriptors.Add(leaf);
                }
            }

            return descriptors;
        }

        private static dynamic ResolveCubeField(dynamic pivot, string modelTableName, string fieldName)
        {
            var escapedTable = modelTableName.Replace("]", "]]" );
            var escapedField = fieldName.Replace("]", "]]" );
            var candidates = new[]
            {
                "[" + escapedTable + "].[" + escapedField + "]",
                "[" + escapedTable + "].[" + escapedField + "].[" + escapedField + "]"
            };
            foreach (var candidate in candidates)
            {
                try
                {
                    return pivot.CubeFields.Item(candidate);
                }
                catch (Exception)
                {
                }
            }

            dynamic fields = pivot.CubeFields;
            var count = Convert.ToInt32(fields.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic field = fields.Item(index);
                var caption = Convert.ToString(field.Caption, CultureInfo.InvariantCulture);
                var name = Convert.ToString(field.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                if (string.Equals(caption, fieldName, StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith(".[" + escapedField + "]", StringComparison.OrdinalIgnoreCase))
                {
                    return field;
                }
            }

            throw new InvalidOperationException("A required Data Model field could not be found in the managed PivotTable.");
        }

        private static dynamic ResolveVisiblePivotField(dynamic pivot, dynamic cubeField, string fieldName)
        {
            var candidates = new[]
            {
                Convert.ToString(cubeField.Name, CultureInfo.InvariantCulture),
                Convert.ToString(cubeField.Caption, CultureInfo.InvariantCulture),
                fieldName
            };
            foreach (var candidate in candidates.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                try
                {
                    return pivot.PivotFields(candidate);
                }
                catch (Exception)
                {
                }
            }

            throw new InvalidOperationException("A required Data Model PivotField could not be activated.");
        }

        private void RemoveExistingOwnedPivot(
            dynamic workbook,
            dynamic worksheet,
            string pivotName,
            ManagedObjectIdentity identity)
        {
            dynamic? pivot = null;
            try
            {
                pivot = worksheet.PivotTables(pivotName);
            }
            catch (Exception)
            {
            }

            if (pivot == null)
            {
                return;
            }

            if (!registry.IsOwned(workbook, identity, pivotName))
            {
                throw new InvalidOperationException("A PivotTable with the requested name exists but is unmanaged.");
            }

            pivot.TableRange2.Clear();
        }

        internal static int ConsolidationFunction(AggregateFunction function)
        {
            switch (function)
            {
                case AggregateFunction.Sum: return -4157;
                case AggregateFunction.Count: return -4112;
                case AggregateFunction.Average: return -4106;
                case AggregateFunction.Minimum: return -4139;
                case AggregateFunction.Maximum: return -4136;
                case AggregateFunction.DistinctCount: return 11;
                default: throw new NotSupportedException("The aggregate function is not supported by Excel PivotTables.");
            }
        }

        private static string ScalarToPivotCaption(ScalarValue scalar)
        {
            switch (scalar.Kind)
            {
                case ScalarValueKind.Null: return "(blank)";
                case ScalarValueKind.Text: return scalar.Text ?? string.Empty;
                case ScalarValueKind.Number: return scalar.Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                case ScalarValueKind.Boolean: return scalar.Boolean == true ? "TRUE" : "FALSE";
                case ScalarValueKind.Date:
                    return scalar.Temporal?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? string.Empty;
                case ScalarValueKind.DateTime:
                    return scalar.Temporal?.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) ?? string.Empty;
                default: throw new NotSupportedException("The filter value kind is not supported.");
            }
        }

        private static string FilterKey(IReadOnlyList<MeasureFilterSpec> filters)
        {
            return string.Join(";", filters
                .Select(filter =>
                    (filter.Field ?? string.Empty).ToUpperInvariant() + ":" + filter.Operator + ":" +
                    string.Join(",", filter.Values
                        .Select(ScalarToPivotCaption)
                        .OrderBy(value => value, StringComparer.Ordinal)))
                .OrderBy(value => value, StringComparer.Ordinal));
        }
    }

}
