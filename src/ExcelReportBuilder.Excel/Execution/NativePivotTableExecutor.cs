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

        public string PivotCacheName { get; set; } = string.Empty;

        public int PivotCacheIndex { get; set; }

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
            var identity = new ManagedObjectIdentity(
                reportId,
                block.OwnershipId,
                ManagedObjectKind.PivotTable);
            var cacheSlot = ManagedPivotCacheSlot.For(
                reportId,
                block.OwnershipId,
                block.Pivot.ManagedCacheName,
                source.Backend);
            var cacheSlots = new[]
            {
                ManagedPivotCacheSlot.For(
                    reportId,
                    block.OwnershipId,
                    block.Pivot.ManagedCacheName,
                    CanonicalBackend.Worksheet),
                ManagedPivotCacheSlot.For(
                    reportId,
                    block.OwnershipId,
                    block.Pivot.ManagedCacheName,
                    CanonicalBackend.DataModel)
            };
            var cacheIdentity = cacheSlot.Identity;
            var records = registry.Load((object)workbook);
            var sourceContract = PivotCacheSourceContract.From(source);
            dynamic? existingPivot = FindExistingOwnedPivot(
                workbook,
                destinationWorksheet,
                block.Pivot.ManagedPivotName,
                identity,
                cacheSlots,
                records);
            records = registry.Load((object)workbook);
            dynamic? candidateCache = null;
            PivotCacheSnapshot? candidateSnapshot = null;
            var registrations = records.Where(record =>
                string.Equals(record.ReportId, cacheIdentity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, cacheIdentity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == cacheIdentity.Kind).ToList();
            var candidateIsExistingManagedPivotCache = false;
            if (registrations.Count == 1 &&
                int.TryParse(
                    registrations[0].Locator,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var registeredIndex) &&
                registeredIndex > 0)
            {
                candidateCache = TryGetPivotCache(workbook, registeredIndex);
                if (candidateCache != null)
                {
                    candidateSnapshot = ReadPivotCacheSnapshot(workbook, candidateCache);
                    candidateIsExistingManagedPivotCache = existingPivot != null &&
                        Convert.ToInt32(
                            ((dynamic)existingPivot!).CacheIndex,
                            CultureInfo.InvariantCulture) == candidateSnapshot.Index;
                }
            }
            else if (registrations.Count == 0 && existingPivot != null)
            {
                candidateCache = ((dynamic)existingPivot!).PivotCache();
                candidateSnapshot = ReadPivotCacheSnapshot(workbook, candidateCache);
                candidateIsExistingManagedPivotCache = true;
            }

            var cachePlan = ManagedPivotCachePolicy.Plan(
                records,
                cacheIdentity,
                cacheSlot.RegistryName,
                sourceContract,
                candidateIsExistingManagedPivotCache,
                candidateSnapshot);
            if (existingPivot != null)
            {
                existingPivot.TableRange2.Clear();
            }

            ClearOwnedDestination(destinationWorksheet, destinationCell, block);

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.BuildingPivots,
                Operation = (cachePlan.Action == ManagedPivotCacheAction.Reuse ||
                             cachePlan.Action == ManagedPivotCacheAction.ReuseAndRegister
                        ? "Reusing managed PivotCache " + block.Pivot.ManagedCacheName +
                          " and rebuilding PivotTable " + block.Pivot.ManagedPivotName + "."
                        : "Creating managed PivotCache " + block.Pivot.ManagedCacheName +
                          " and PivotTable " + block.Pivot.ManagedPivotName + "."),
                ManagedObject = block.Pivot.ManagedPivotName,
                ProjectedRows = source.ProjectedRows
            });

            dynamic cache;
            if (cachePlan.Action == ManagedPivotCacheAction.Reuse ||
                cachePlan.Action == ManagedPivotCacheAction.ReuseAndRegister)
            {
                cache = candidateCache ?? throw new InvalidOperationException(
                    "The validated managed PivotCache reuse plan has no cache object.");
            }
            else
            {
                if (cachePlan.Action == ManagedPivotCacheAction.RetireAndCreate)
                {
                    registry.Remove((object)workbook, new[] { cacheIdentity });
                }

                dynamic sourceData = source.Backend == CanonicalBackend.Worksheet
                    ? (dynamic)source.TableOrConnectionName
                    : workbook.Connections.Item(source.TableOrConnectionName);
                var sourceType = source.Backend == CanonicalBackend.Worksheet
                    ? SourceDatabase
                    : SourceExternal;
                cache = workbook.PivotCaches().Create(sourceType, sourceData, PivotVersion15);
            }

            if (cachePlan.Action == ManagedPivotCacheAction.Reuse ||
                cachePlan.Action == ManagedPivotCacheAction.ReuseAndRegister)
            {
                progressSink.Report(new ExcelProgress
                {
                    Stage = ExcelBuildStage.BuildingPivots,
                    Operation = "Refreshing the exact managed PivotCache before rebuilding its PivotTable.",
                    ManagedObject = block.Pivot.ManagedCacheName,
                    ProjectedRows = source.ProjectedRows
                });
                cache.Refresh();
            }

            var cacheSnapshot = ReadPivotCacheSnapshot(workbook, cache);
            if (!sourceContract.Matches(cacheSnapshot))
            {
                throw new InvalidOperationException(
                    "Excel created or returned a PivotCache that does not match the validated source contract.");
            }

            registry.Register(
                workbook,
                cacheIdentity,
                cacheSlot.RegistryName,
                cacheSnapshot.Index.ToString(CultureInfo.InvariantCulture),
                sourceContract.Serialized);
            dynamic destination = destinationWorksheet.Range[destinationCell];
            dynamic pivot = cache.CreatePivotTable(destination, block.Pivot.ManagedPivotName);
            registry.Register(
                workbook,
                identity,
                block.Pivot.ManagedPivotName,
                cacheSnapshot.Index.ToString(CultureInfo.InvariantCulture),
                sourceContract.Serialized);
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

                return new PivotBuildResult
                {
                    PivotTableName = block.Pivot.ManagedPivotName,
                    PivotCacheName = block.Pivot.ManagedCacheName,
                    PivotCacheIndex = cacheSnapshot.Index,
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

        private dynamic? FindExistingOwnedPivot(
            dynamic workbook,
            dynamic worksheet,
            string pivotName,
            ManagedObjectIdentity identity,
            IReadOnlyList<ManagedPivotCacheSlot> cacheSlots,
            IReadOnlyList<ManagedObjectRecord> records)
        {
            var registrations = records.Where(record =>
                SameIdentity(record, identity)).ToList();
            if (registrations.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the managed PivotTable identity.");
            }

            if (records.Any(record =>
                    record.Kind == ManagedObjectKind.PivotTable &&
                    !SameIdentity(record, identity) &&
                    string.Equals(record.ExcelName, pivotName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "The requested managed PivotTable name is already owned by another object.");
            }

            var registration = registrations.SingleOrDefault();
            dynamic? pivot = TryGetPivotTable(worksheet, pivotName);
            if (pivot == null)
            {
                if (registration != null)
                {
                    dynamic? differentlyNamed = TryGetPivotTable(
                        worksheet,
                        registration.ExcelName);
                    if (differentlyNamed != null)
                    {
                        throw new InvalidOperationException(
                            "The managed PivotTable identity is registered under a different live name.");
                    }

                    registry.Remove((object)workbook, new[] { identity });
                }

                return null;
            }

            if (registration == null)
            {
                throw new InvalidOperationException(
                    "A PivotTable with the requested name exists but is unmanaged.");
            }

            var livePivotCacheIndex = Convert.ToInt32(
                pivot.CacheIndex,
                CultureInfo.InvariantCulture);
            RegisteredPivotCacheBinding binding = ResolveExactLivePivotCacheBinding(
                records,
                cacheSlots,
                livePivotCacheIndex);

            dynamic? cache = TryGetPivotCache(workbook, livePivotCacheIndex);
            if (cache == null)
            {
                throw new InvalidOperationException(
                    "The live managed PivotTable's registered PivotCache cannot be located.");
            }

            var cacheSnapshot = ReadPivotCacheSnapshot(workbook, cache);
            DemandExactExistingPivotContract(
                registration,
                pivotName,
                binding.Registration,
                binding.Slot,
                livePivotCacheIndex,
                cacheSnapshot);
            return pivot;
        }

        internal static RegisteredPivotCacheBinding ResolveExactLivePivotCacheBinding(
            IReadOnlyList<ManagedObjectRecord> records,
            IReadOnlyList<ManagedPivotCacheSlot> cacheSlots,
            int livePivotCacheIndex)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (cacheSlots == null) throw new ArgumentNullException(nameof(cacheSlots));
            if (cacheSlots.Count != 2 || livePivotCacheIndex < 1)
            {
                throw new InvalidOperationException(
                    "A live managed PivotTable must be checked against its two backend cache slots and a valid cache index.");
            }

            var candidates = new List<RegisteredPivotCacheBinding>();
            foreach (ManagedPivotCacheSlot slot in cacheSlots)
            {
                var exact = records.Where(record =>
                    SameIdentity(record, slot.Identity)).ToList();
                if (exact.Count > 1)
                {
                    throw new InvalidOperationException(
                        "More than one ownership record claims a managed PivotCache backend slot.");
                }

                ManagedObjectRecord? registration = exact.SingleOrDefault();
                if (registration != null &&
                    int.TryParse(
                        registration.Locator,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out var registeredIndex) &&
                    registeredIndex == livePivotCacheIndex)
                {
                    candidates.Add(new RegisteredPivotCacheBinding
                    {
                        Slot = slot,
                        Registration = registration
                    });
                }
            }

            if (candidates.Count != 1)
            {
                throw new InvalidOperationException(
                    "The live managed PivotTable does not match exactly one registered backend PivotCache slot.");
            }

            return candidates[0];
        }

        internal static void DemandExactExistingPivotContract(
            ManagedObjectRecord pivotRegistration,
            string expectedPivotName,
            ManagedObjectRecord cacheRegistration,
            ManagedPivotCacheSlot cacheSlot,
            int livePivotCacheIndex,
            PivotCacheSnapshot liveCache)
        {
            if (pivotRegistration == null) throw new ArgumentNullException(nameof(pivotRegistration));
            if (cacheRegistration == null) throw new ArgumentNullException(nameof(cacheRegistration));
            if (cacheSlot == null) throw new ArgumentNullException(nameof(cacheSlot));
            if (liveCache == null) throw new ArgumentNullException(nameof(liveCache));

            if (!string.Equals(
                    pivotRegistration.ExcelName,
                    expectedPivotName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The managed PivotTable identity is registered under a different name.");
            }

            if (!SameIdentity(cacheRegistration, cacheSlot.Identity) ||
                !string.Equals(
                    cacheRegistration.ExcelName,
                    cacheSlot.RegistryName,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The managed PivotTable does not reference its exact registered PivotCache slot.");
            }


            if ((!string.IsNullOrWhiteSpace(pivotRegistration.Locator) &&
                 !string.Equals(
                     pivotRegistration.Locator,
                     cacheRegistration.Locator,
                     StringComparison.Ordinal)) ||
                (!string.IsNullOrWhiteSpace(pivotRegistration.SourceContract) &&
                 !string.Equals(
                     pivotRegistration.SourceContract,
                     cacheRegistration.SourceContract,
                     StringComparison.Ordinal)))
            {
                throw new InvalidOperationException(
                    "The managed PivotTable ownership record disagrees with its registered PivotCache contract.");
            }

            if (!int.TryParse(
                    cacheRegistration.Locator,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var registeredCacheIndex) ||
                registeredCacheIndex < 1 ||
                livePivotCacheIndex != registeredCacheIndex ||
                liveCache.Index != registeredCacheIndex)
            {
                throw new InvalidOperationException(
                    "The live PivotTable and registered PivotCache locator do not match.");
            }

            var registeredSource = PivotCacheSourceContract.Parse(
                cacheRegistration.SourceContract ?? string.Empty);
            if (!registeredSource.Matches(liveCache))
            {
                throw new InvalidOperationException(
                    "The live PivotTable's cache no longer matches its registered source contract.");
            }
        }

        private static dynamic? TryGetPivotTable(dynamic worksheet, string pivotName)
        {
            if (string.IsNullOrWhiteSpace(pivotName))
            {
                return null;
            }

            try
            {
                return worksheet.PivotTables(pivotName);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool SameIdentity(
            ManagedObjectRecord record,
            ManagedObjectIdentity identity)
        {
            return string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                   string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                   record.Kind == identity.Kind;
        }

        private static dynamic? TryGetPivotCache(dynamic workbook, int index)
        {
            try
            {
                return workbook.PivotCaches().Item(index);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static PivotCacheSnapshot ReadPivotCacheSnapshot(
            dynamic workbook,
            dynamic cache)
        {
            var index = Convert.ToInt32(cache.Index, CultureInfo.InvariantCulture);
            var sourceType = Convert.ToInt32(cache.SourceType, CultureInfo.InvariantCulture);
            var snapshot = new PivotCacheSnapshot
            {
                Index = index,
                SourceType = sourceType,
                PivotTableCount = CountPivotTablesUsingCache(workbook, index)
            };
            if (sourceType == SourceDatabase)
            {
                var source = Convert.ToString(cache.SourceData, CultureInfo.InvariantCulture);
                snapshot.WorksheetSource = ResolveWorksheetSourceName(workbook, source);
            }
            else if (sourceType == SourceExternal)
            {
                dynamic connection = cache.WorkbookConnection;
                snapshot.ConnectionName = Convert.ToString(
                    connection.Name,
                    CultureInfo.InvariantCulture);
            }

            return snapshot;
        }

        private static int CountPivotTablesUsingCache(dynamic workbook, int cacheIndex)
        {
            var result = 0;
            dynamic worksheets = workbook.Worksheets;
            var worksheetCount = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            for (var worksheetIndex = 1; worksheetIndex <= worksheetCount; worksheetIndex++)
            {
                dynamic worksheet = worksheets.Item(worksheetIndex);
                dynamic pivots = worksheet.PivotTables();
                var pivotCount = Convert.ToInt32(pivots.Count, CultureInfo.InvariantCulture);
                for (var pivotIndex = 1; pivotIndex <= pivotCount; pivotIndex++)
                {
                    dynamic pivot = pivots.Item(pivotIndex);
                    if (Convert.ToInt32(pivot.CacheIndex, CultureInfo.InvariantCulture) == cacheIndex)
                    {
                        result++;
                    }
                }
            }

            return result;
        }

        private static string? ResolveWorksheetSourceName(
            dynamic workbook,
            string? sourceReference)
        {
            if (string.IsNullOrWhiteSpace(sourceReference))
            {
                return sourceReference;
            }

            dynamic worksheets = workbook.Worksheets;
            var worksheetCount = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            for (var worksheetIndex = 1; worksheetIndex <= worksheetCount; worksheetIndex++)
            {
                dynamic worksheet = worksheets.Item(worksheetIndex);
                dynamic tables = worksheet.ListObjects;
                var tableCount = Convert.ToInt32(tables.Count, CultureInfo.InvariantCulture);
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    dynamic table = tables.Item(tableIndex);
                    var tableName = Convert.ToString(table.Name, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(tableName))
                    {
                        continue;
                    }

                    if (PivotCacheSourceContract.WorksheetSourceMatches(sourceReference, tableName!))
                    {
                        return tableName;
                    }

                    dynamic range = table.Range;
                    var localAddresses = new[]
                    {
                        Convert.ToString(range.Address[true, true, 1, false], CultureInfo.InvariantCulture),
                        Convert.ToString(range.Address[true, true, -4150, false], CultureInfo.InvariantCulture)
                    };
                    var worksheetName = Convert.ToString(
                        worksheet.Name,
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    var qualifiedAddresses = localAddresses
                        .Where(address => !string.IsNullOrWhiteSpace(address))
                        .SelectMany(address => new[]
                        {
                            worksheetName + "!" + address,
                            "'" + worksheetName.Replace("'", "''") + "'!" + address
                        });
                    var addresses = new[]
                    {
                        Convert.ToString(range.Address[true, true, 1, true], CultureInfo.InvariantCulture),
                        Convert.ToString(range.Address[true, true, -4150, true], CultureInfo.InvariantCulture)
                    }.Concat(localAddresses).Concat(qualifiedAddresses);
                    if (addresses.Any(address =>
                            SourceReferencesEqual(sourceReference, address) ||
                            SourceReferenceStartsEqual(sourceReference, address)))
                    {
                        return tableName;
                    }
                }
            }

            return sourceReference;
        }

        internal static bool SourceReferencesEqual(string? left, string? right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            {
                return false;
            }

            return string.Equals(
                NormalizeSourceReference(left!),
                NormalizeSourceReference(right!),
                StringComparison.OrdinalIgnoreCase);
        }

        internal static bool SourceReferenceStartsEqual(string? left, string? right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            {
                return false;
            }

            var leftStart = SourceReferenceStart(NormalizeSourceReference(left!));
            var rightStart = SourceReferenceStart(NormalizeSourceReference(right!));
            return leftStart.IndexOf("[", StringComparison.Ordinal) < 0 &&
                   rightStart.IndexOf("[", StringComparison.Ordinal) < 0 &&
                   string.Equals(leftStart, rightStart, StringComparison.OrdinalIgnoreCase);
        }

        private static string SourceReferenceStart(string value)
        {
            var rangeStart = value.LastIndexOf('!') + 1;
            var separator = value.IndexOf(':', rangeStart);
            return separator < 0 ? value : value.Substring(0, separator);
        }

        private static string NormalizeSourceReference(string value)
        {
            var normalized = value.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1);
            }

            return normalized
                .Replace("$", string.Empty)
                .Replace("'", string.Empty)
                .Replace("\"", string.Empty)
                .Replace(" ", string.Empty);
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
