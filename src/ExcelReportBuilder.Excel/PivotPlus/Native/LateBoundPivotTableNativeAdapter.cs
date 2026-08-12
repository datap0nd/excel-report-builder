using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using CorePivotFieldArea = ExcelReportBuilder.Core.PivotPlus.PivotFieldArea;

namespace ExcelReportBuilder.Excel.PivotPlus.Native
{
    /// <summary>
    /// The only PivotTable+ field/layout type that talks to the Excel object
    /// model. Constants are the stable values from Excel's native enums so the
    /// assembly remains Office-version-neutral and late bound.
    /// </summary>
    internal sealed class LateBoundPivotTableNativeAdapter : IPivotTableNativeAdapter
    {
        private const int OrientationHidden = 0;
        private const int OrientationRow = 1;
        private const int OrientationColumn = 2;
        private const int OrientationPage = 3;
        private const int OrientationData = 4;
        private const int DoNotRepeatLabels = 1;
        private const int RepeatLabels = 2;
        private const int NoAdditionalCalculation = -4143;
        private const int MaximumNativeFields = 16384;
        private const int PivotSourceTypeDatabase = 1;
        private const int ConnectionTypeModel = 7;

        public PivotTargetIdentity ReadTarget(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (workbookIdentityResolver == null)
            {
                throw new ArgumentNullException(nameof(workbookIdentityResolver));
            }

            dynamic pivot = pivotTable;
            string pivotName = ReadRequiredName(
                () => (object?)pivot.Name,
                "Excel did not expose the selected PivotTable's name.");
            object worksheetObject = ReadRequiredObject(
                () => (object?)pivot.Parent,
                "Excel did not expose the worksheet containing the selected PivotTable.");
            dynamic worksheet = worksheetObject;
            string worksheetName = ReadRequiredName(
                () => (object?)worksheet.Name,
                "Excel did not expose the selected PivotTable worksheet's name.");
            object workbookObject = ReadRequiredObject(
                () => (object?)worksheet.Parent,
                "Excel did not expose the workbook containing the selected PivotTable.");
            string workbookId = workbookIdentityResolver.Resolve(workbookObject);
            if (string.IsNullOrWhiteSpace(workbookId))
            {
                throw new InvalidOperationException(
                    "Excel did not expose a path-free identity for the selected PivotTable workbook.");
            }

            return new PivotTargetIdentity(workbookId, worksheetName, pivotName);
        }

        public void PersistWorkbookIdentity(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver,
            string expectedWorkbookId)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (workbookIdentityResolver == null)
            {
                throw new ArgumentNullException(nameof(workbookIdentityResolver));
            }

            dynamic pivot = pivotTable;
            object worksheetObject = ReadRequiredObject(
                () => (object?)pivot.Parent,
                "Excel did not expose the worksheet containing the selected PivotTable.");
            dynamic worksheet = worksheetObject;
            object workbookObject = ReadRequiredObject(
                () => (object?)worksheet.Parent,
                "Excel did not expose the workbook containing the selected PivotTable.");
            workbookIdentityResolver.Persist(workbookObject, expectedWorkbookId);
        }

        public NativePivotSourceIdentity ReadSource(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));

            dynamic pivot = pivotTable;
            object cacheObject = ReadPivotCache(pivot);
            dynamic cache = cacheObject;
            bool isOlap = ReadRequiredBoolean(
                () => (object?)cache.OLAP,
                "Excel did not expose whether the selected PivotCache uses an OLAP source.");
            if (!isOlap)
            {
                int sourceType = ReadRequiredInt(
                    () => (object?)cache.SourceType,
                    "Excel did not expose a valid non-OLAP PivotCache source type.");
                if (sourceType != PivotSourceTypeDatabase)
                {
                    throw new NotSupportedException(
                        "PivotTable+ native mutation requires an xlDatabase worksheet PivotCache for a classic source.");
                }

                string sourceName = ReadRequiredClassicSourceName(cacheObject);
                return new NativePivotSourceIdentity(
                    NativePivotCacheKind.ClassicDatabase,
                    sourceName);
            }

            object connectionObject = ReadRequiredObject(
                () => (object?)cache.WorkbookConnection,
                "Excel did not expose the selected OLAP PivotCache workbook connection.");
            dynamic connection = connectionObject;
            int connectionType = ReadRequiredInt(
                () => (object?)connection.Type,
                "Excel did not expose a valid selected PivotCache connection type.");
            string connectionName = ReadRequiredName(
                () => (object?)connection.Name,
                "Excel did not expose the selected PivotCache connection name.");
            DemandSafeSourceName(connectionName);
            return new NativePivotSourceIdentity(
                connectionType == ConnectionTypeModel
                    ? NativePivotCacheKind.DataModel
                    : NativePivotCacheKind.ExternalOlap,
                connectionName);
        }

        public object CaptureState(object pivotTable, PivotSourceKind sourceKind)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            DemandNoUnsupportedCalculatedObjects(pivot);
            DemandNoActiveNativeFilters(pivot, sourceKind);
            DemandNormalValueCalculations(pivot);
            DemandUniformRowRepeatLabels(pivot);
            var fields = new List<SnapshotField>();
            CaptureArea(pivot, sourceKind, CorePivotFieldArea.Row, fields);
            CaptureArea(pivot, sourceKind, CorePivotFieldArea.Column, fields);
            CaptureArea(pivot, sourceKind, CorePivotFieldArea.Filter, fields);
            CaptureArea(pivot, sourceKind, CorePivotFieldArea.Values, fields);
            ReadValuesAxis(
                pivot,
                fields.Count(field => field.Area == CorePivotFieldArea.Values),
                out PivotValuesAxis valuesAxis,
                out int valuesPosition);

            return new NativePivotSnapshot
            {
                SourceKind = sourceKind,
                Fields = fields,
                CubeFields = sourceKind == PivotSourceKind.DataModel
                    ? CaptureCubeFields((object)pivot)
                    : Array.Empty<SnapshotCubeField>(),
                Layout = new NativePivotLayoutCommand
                {
                    RowAxisLayout = ReadRequiredInt(
                        () => (object?)pivot.LayoutRowDefault,
                        "Excel did not expose the row-axis layout required for rollback."),
                    RepeatItemLabels = ReadRepeatLabels(pivot),
                    ShowRowGrandTotals = ReadRequiredBoolean(
                        () => (object?)pivot.RowGrand,
                        "Excel did not expose the row-grand-total state required for rollback."),
                    ShowColumnGrandTotals = ReadRequiredBoolean(
                        () => (object?)pivot.ColumnGrand,
                        "Excel did not expose the column-grand-total state required for rollback."),
                    ShowFieldHeaders = ReadRequiredBoolean(
                        () => (object?)pivot.DisplayFieldCaptions,
                        "Excel did not expose the field-header state required for rollback."),
                    ValuesAxis = valuesAxis,
                    ValuesPosition = valuesPosition,
                    PivotTableStyleName = ReadPivotTableStyleName(pivot),
                    SetPivotTableStyle = !string.IsNullOrWhiteSpace(
                        ReadPivotTableStyleName(pivot)),
                    PreserveFormatting = ReadRequiredBoolean(
                        () => (object?)pivot.PreserveFormatting,
                        "Excel did not expose the preserve-formatting state required for rollback."),
                    ShowRowStripes = ReadRequiredBoolean(
                        () => (object?)pivot.ShowTableStyleRowStripes,
                        "Excel did not expose the row-stripe state required for rollback."),
                    ShowColumnStripes = ReadRequiredBoolean(
                        () => (object?)pivot.ShowTableStyleColumnStripes,
                        "Excel did not expose the column-stripe state required for rollback.")
                }
            };
        }

        public void ClearLayout(object pivotTable, PivotSourceKind sourceKind)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;

            IReadOnlyList<object> dataFields = ReadCollection((object)pivot, "DataFields");
            object? dataPivotField = ReadDataPivotField(
                pivot,
                required: dataFields.Count > 1);
            var fields = dataFields
                .Concat(ReadCollection((object)pivot, "PageFields"))
                // Excel refuses to hide an outer compact-form hierarchy while
                // an inner field is still visible.  Work from the innermost
                // position back toward the outer field on both native axes.
                .Concat(ReadAreaCollection((object)pivot, "ColumnFields").Reverse())
                .Concat(ReadAreaCollection((object)pivot, "RowFields").Reverse())
                .ToList();

            try
            {
                foreach (object fieldObject in fields)
                {
                    dynamic field = fieldObject;
                    if (IsClassic(sourceKind))
                    {
                        field.Orientation = OrientationHidden;
                    }
                    else
                    {
                        object? cubeObject = ReadObject(field, "CubeField");
                        dynamic cube = cubeObject ?? fieldObject;
                        cube.Orientation = OrientationHidden;
                    }
                }

                if (dataPivotField != null)
                {
                    ((dynamic)dataPivotField).Orientation = OrientationHidden;
                }
            }
            catch (Exception compatibilityFailure)
            {
                // ClearTable is destructive to classic grouped fields on Excel
                // 2021: the generated Months/Quarters fields disappear and can
                // no longer be restored by name. Use it only when ordered field
                // removal is genuinely unavailable, never as the first choice.
                try
                {
                    pivot.ClearTable();
                    return;
                }
                catch (Exception clearTableFailure)
                {
                    throw new InvalidOperationException(
                        "Excel could not clear the PivotTable layout using either " +
                        "ordered field removal or ClearTable. Ordered removal: " +
                        compatibilityFailure.Message + " | ClearTable: " +
                        clearTableFailure.Message,
                        clearTableFailure);
                }
            }
        }

        public void RemoveFieldsNotInPlan(
            object pivotTable,
            PivotSourceKind sourceKind,
            IReadOnlyList<NativePivotFieldCommand> desiredFields)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (desiredFields == null) throw new ArgumentNullException(nameof(desiredFields));

            dynamic pivot = pivotTable;
            var desiredAxes = new HashSet<string>(
                desiredFields
                    .Where(item => item.Area != CorePivotFieldArea.Values)
                    .Select(item => item.FieldName),
                StringComparer.OrdinalIgnoreCase);

            var visibleAxes = ReadCollection((object)pivot, "PageFields")
                .Concat(ReadAreaCollection((object)pivot, "ColumnFields").Reverse())
                .Concat(ReadAreaCollection((object)pivot, "RowFields").Reverse())
                .ToList();
            foreach (object fieldObject in visibleAxes)
            {
                string liveName = ReadFieldIdentity(fieldObject, sourceKind);
                if (desiredAxes.Contains(liveName)) continue;

                dynamic field = fieldObject;
                if (IsClassic(sourceKind))
                {
                    field.Orientation = OrientationHidden;
                }
                else
                {
                    object? cubeObject = ReadObject(field, "CubeField");
                    ((dynamic)(cubeObject ?? fieldObject)).Orientation = OrientationHidden;
                }
            }

            IReadOnlyList<NativePivotFieldCommand> desiredValues = desiredFields
                .Where(item => item.Area == CorePivotFieldArea.Values)
                .ToList();
            foreach (object dataFieldObject in ReadCollection((object)pivot, "DataFields"))
            {
                if (desiredValues.Any(command =>
                        ExistingValueMatches(dataFieldObject, sourceKind, command)))
                {
                    continue;
                }

                dynamic dataField = dataFieldObject;
                dataField.Orientation = OrientationHidden;
            }
        }

        public void PlaceField(
            object pivotTable,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (command == null) throw new ArgumentNullException(nameof(command));

            dynamic pivot = pivotTable;
            if (command.Area == CorePivotFieldArea.Values)
            {
                if (IsClassic(sourceKind))
                {
                    PlaceClassicValue(pivot, command);
                }
                else
                {
                    PlaceCubeValue(pivot, sourceKind, command);
                }

                return;
            }

            int orientation = Orientation(command.Area);
            dynamic visibleField;
            if (IsClassic(sourceKind))
            {
                visibleField = ResolvePivotField(pivot, command.FieldName);
                visibleField.Orientation = orientation;
                visibleField.Position = command.Position;
            }
            else
            {
                dynamic cubeField = ResolveCubeField(pivot, command.FieldName);
                cubeField.Orientation = orientation;
                cubeField.Position = command.Position;
                visibleField = ResolveVisiblePivotField(pivot, cubeField, command.FieldName);
            }

            ApplyCaption(visibleField, command.Caption, command.SetCaption);
            if (command.Area == CorePivotFieldArea.Row)
            {
                ApplySubtotals(visibleField, command.SubtotalMode);
            }
        }

        public void ApplyLayout(object pivotTable, NativePivotLayoutCommand command)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (command == null) throw new ArgumentNullException(nameof(command));

            dynamic pivot = pivotTable;
            pivot.RowAxisLayout(command.RowAxisLayout);
            ApplyValuesAxis(pivot, command.ValuesAxis, command.ValuesPosition);
            pivot.RepeatAllLabels(command.RepeatItemLabels ? RepeatLabels : DoNotRepeatLabels);
            pivot.RowGrand = command.ShowRowGrandTotals;
            pivot.ColumnGrand = command.ShowColumnGrandTotals;
            pivot.DisplayFieldCaptions = command.ShowFieldHeaders;
            pivot.PreserveFormatting = command.PreserveFormatting;
            pivot.ShowTableStyleRowStripes = command.ShowRowStripes;
            pivot.ShowTableStyleColumnStripes = command.ShowColumnStripes;
            if (command.SetPivotTableStyle)
            {
                pivot.TableStyle2 = command.PivotTableStyleName ?? string.Empty;
            }
        }

        public void RestoreState(object pivotTable, object snapshot)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (!(snapshot is NativePivotSnapshot native))
            {
                throw new ArgumentException("The PivotTable snapshot is invalid.", nameof(snapshot));
            }

            if (native.SourceKind == PivotSourceKind.DataModel)
            {
                DeleteCubeFieldsCreatedAfterSnapshot(
                    pivotTable,
                    native.CubeFields);
                RestoreCubeFieldCaptions(pivotTable, native.CubeFields);
            }

            var restoreCommands = native.Fields
                         .OrderBy(item => AreaOrder(item.Area))
                         .ThenBy(item => item.Position)
                         .Select(field => new NativePivotFieldCommand
                {
                    InstanceId = "rollback:" + field.Area + ":" +
                                 field.Position.ToString(CultureInfo.InvariantCulture),
                    FieldName = field.FieldName,
                    Caption = field.Caption,
                    SetCaption = true,
                    Area = field.Area,
                    Position = field.Position,
                    IsMeasure = field.IsCubeMeasure,
                    ConsolidationFunction = field.ConsolidationFunction,
                    NumberFormatCode = field.NumberFormatCode,
                    SubtotalMode = field.Subtotals.Length > 0 && field.Subtotals[0]
                        ? PivotSubtotalMode.Automatic
                        : PivotSubtotalMode.None
                })
                .ToList();

            RemoveFieldsNotInPlan(pivotTable, native.SourceKind, restoreCommands);
            foreach (NativePivotFieldCommand command in restoreCommands)
            {
                PlaceField(pivotTable, native.SourceKind, command);
                if (command.Area == CorePivotFieldArea.Row)
                {
                    dynamic visible = ResolvePlacedField(
                        (dynamic)pivotTable,
                        native.SourceKind,
                        command);
                    SnapshotField field = native.Fields.Single(item =>
                        item.Area == command.Area && item.Position == command.Position);
                    WriteSubtotals(visible, field.Subtotals);
                }
            }

            ApplyLayout(pivotTable, native.Layout);
            if (native.SourceKind == PivotSourceKind.DataModel)
            {
                VerifyCubeFieldCaptions(pivotTable, native.CubeFields);
            }
        }

        public void Refresh(object pivotTable)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            dynamic pivot = pivotTable;
            pivot.RefreshTable();
        }

        public void Verify(object pivotTable, NativePivotMutationPlan plan)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            dynamic pivot = pivotTable;

            DemandExactAreaFieldCount(
                (object)pivot,
                "RowFields",
                plan.Fields.Count(command => command.Area == CorePivotFieldArea.Row));
            DemandExactAreaFieldCount(
                (object)pivot,
                "ColumnFields",
                plan.Fields.Count(command => command.Area == CorePivotFieldArea.Column));
            DemandExactAreaFieldCount(
                (object)pivot,
                "PageFields",
                plan.Fields.Count(command => command.Area == CorePivotFieldArea.Filter));
            DemandExactAreaFieldCount(
                (object)pivot,
                "DataFields",
                plan.Fields.Count(command => command.Area == CorePivotFieldArea.Values));

            foreach (NativePivotFieldCommand command in plan.Fields)
            {
                dynamic field = ResolvePlacedField(pivot, plan.SourceKind, command);
                int actualPosition = command.Area == CorePivotFieldArea.Row ||
                                     command.Area == CorePivotFieldArea.Column
                    ? ReadNormalizedAreaPosition((object)pivot, command.Area, (object)field)
                    : ReadInt(field, "Position", -1);
                DemandEqual(
                    command.Position,
                    actualPosition,
                    "Excel placed a PivotTable field at the wrong position.");
                DemandEqual(
                    Orientation(command.Area),
                    ReadInt(field, "Orientation", -1),
                    "Excel placed a PivotTable field in the wrong area.");

                if (command.SetCaption)
                {
                    DemandEqual(
                        command.Caption,
                        ReadOptionalString(field, "Caption") ?? string.Empty,
                        "Excel applied the wrong PivotTable field caption.");
                }

                if (command.Area == CorePivotFieldArea.Row)
                {
                    var expectedSubtotals = new bool[12];
                    expectedSubtotals[0] =
                        command.SubtotalMode == PivotSubtotalMode.Automatic;
                    if (!expectedSubtotals.SequenceEqual(ReadSubtotals((object)field)))
                    {
                        throw new InvalidOperationException(
                            "Excel applied the wrong 12-slot row subtotal state.");
                    }
                }

                if (command.Area == CorePivotFieldArea.Values &&
                    command.ConsolidationFunction.HasValue)
                {
                    DemandEqual(
                        command.ConsolidationFunction.Value,
                        ReadInt(field, "Function", int.MinValue),
                        "Excel applied the wrong Values aggregation.");
                }

                if (command.Area == CorePivotFieldArea.Values &&
                    !string.IsNullOrWhiteSpace(command.NumberFormatCode))
                {
                    DemandEqual(
                        command.NumberFormatCode!,
                        ReadOptionalString(field, "NumberFormat") ?? string.Empty,
                        "Excel applied the wrong Values number format.");
                }
            }

            DemandEqual(
                plan.Layout.RowAxisLayout,
                ReadInt(pivot, "LayoutRowDefault", -1),
                "Excel applied the wrong row-axis layout.");
            DemandEqual(
                plan.Layout.ShowRowGrandTotals,
                ReadBoolean(pivot, "RowGrand", !plan.Layout.ShowRowGrandTotals),
                "Excel applied the wrong row-grand-total setting.");
            DemandEqual(
                plan.Layout.ShowColumnGrandTotals,
                ReadBoolean(pivot, "ColumnGrand", !plan.Layout.ShowColumnGrandTotals),
                "Excel applied the wrong column-grand-total setting.");
            DemandEqual(
                plan.Layout.ShowFieldHeaders,
                ReadBoolean(pivot, "DisplayFieldCaptions", !plan.Layout.ShowFieldHeaders),
                "Excel applied the wrong field-header setting.");
            DemandEqual(
                plan.Layout.PreserveFormatting,
                ReadBoolean(pivot, "PreserveFormatting", !plan.Layout.PreserveFormatting),
                "Excel applied the wrong preserve-formatting setting.");
            DemandEqual(
                plan.Layout.ShowRowStripes,
                ReadBoolean(pivot, "ShowTableStyleRowStripes", !plan.Layout.ShowRowStripes),
                "Excel applied the wrong row-stripe setting.");
            DemandEqual(
                plan.Layout.ShowColumnStripes,
                ReadBoolean(pivot, "ShowTableStyleColumnStripes", !plan.Layout.ShowColumnStripes),
                "Excel applied the wrong column-stripe setting.");
            if (plan.Layout.SetPivotTableStyle)
            {
                string expectedStyle = plan.Layout.PivotTableStyleName ?? string.Empty;
                string actualStyle = ReadPivotTableStyleName(pivot) ?? string.Empty;
                if (!string.Equals(expectedStyle, actualStyle, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Excel applied the wrong PivotTable style. Expected '" +
                        expectedStyle + "' but read back '" + actualStyle + "'.");
                }
            }


            VerifyValuesAxis(
                pivot,
                plan.Layout.ValuesAxis,
                plan.Layout.ValuesPosition,
                plan.Fields.Count(command => command.Area == CorePivotFieldArea.Values));

            if (plan.Fields.Any(command => command.Area == CorePivotFieldArea.Row))
            {
                DemandEqual(
                    plan.Layout.RepeatItemLabels,
                    ReadRepeatLabels(pivot),
                    "Excel applied the wrong repeated-label setting.");
            }
        }

        private static void DemandExactAreaFieldCount(
            object pivotTable,
            string collectionName,
            int expected)
        {
            int actual = ReadAreaCollection(pivotTable, collectionName).Count;
            DemandEqual(
                expected,
                actual,
                "Excel exposed an unexpected number of fields in " + collectionName + ".");
        }

        private static void PlaceClassicValue(dynamic pivot, NativePivotFieldCommand command)
        {
            if (!command.ConsolidationFunction.HasValue)
            {
                throw new InvalidOperationException(
                    "A classic Values field requires a native aggregation.");
            }

            object? existing = ReadCollection((object)pivot, "DataFields")
                .FirstOrDefault(candidate =>
                    ExistingValueMatches(candidate, PivotSourceKind.WorksheetTable, command));
            dynamic dataField;
            if (existing != null)
            {
                dataField = existing;
                ApplyCaption(dataField, command.Caption, command.SetCaption);
            }
            else
            {
                dynamic sourceField = ResolvePivotField(pivot, command.FieldName);
                dataField = pivot.AddDataField(
                    sourceField,
                    command.Caption,
                    command.ConsolidationFunction.Value);
            }

            dataField.Position = command.Position;
            ApplyNumberFormat(dataField, command.NumberFormatCode);
        }

        private static string ReadFieldIdentity(object fieldObject, PivotSourceKind sourceKind)
        {
            dynamic field = fieldObject;
            if (IsClassic(sourceKind))
            {
                return ReadRequiredName(
                    () => (object?)field.SourceName,
                    "Excel did not expose a classic PivotField source name required for incremental mutation.");
            }

            object? cubeObject = ReadObject(field, "CubeField");
            if (cubeObject == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the CubeField identity required for incremental mutation.");
            }

            dynamic cube = cubeObject;
            return ReadRequiredName(
                () => (object?)cube.Name,
                "Excel did not expose the CubeField name required for incremental mutation.");
        }

        private static bool ExistingValueMatches(
            object dataFieldObject,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command)
        {
            dynamic dataField = dataFieldObject;
            if (IsClassic(sourceKind))
            {
                int? function = command.ConsolidationFunction;
                if (function.HasValue &&
                    (!TryRead(() => (object?)dataField.Function, out object? rawFunction) ||
                     rawFunction == null ||
                     Convert.ToInt32(rawFunction, CultureInfo.InvariantCulture) != function.Value))
                {
                    return false;
                }

                if (MatchesAnyName(dataFieldObject, command.FieldName)) return true;
                object? sourceField = ReadObject(dataField, "PivotField");
                return sourceField != null && MatchesAnyName(sourceField, command.FieldName);
            }

            object? cubeObject = ReadObject(dataField, "CubeField");
            return cubeObject != null && MatchesAnyName(cubeObject, command.FieldName);
        }

        private static void PlaceCubeValue(
            dynamic pivot,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command)
        {
            dynamic cubeField = ResolveCubeField(pivot, command.FieldName);
            if (command.IsMeasure)
            {
                cubeField.Orientation = OrientationData;
                cubeField.Position = command.Position;
            }
            else
            {
                if (sourceKind != PivotSourceKind.DataModel ||
                    !command.ConsolidationFunction.HasValue)
                {
                    throw new NotSupportedException(
                        "Only Data Model sources can create an implicit aggregate measure.");
                }

                dynamic measure = pivot.CubeFields.GetMeasure(
                    cubeField,
                    command.ConsolidationFunction.Value,
                    command.Caption);
                measure.Orientation = OrientationData;
                measure.Position = command.Position;
            }

            dynamic dataField = ResolveDataField(pivot, command.Caption, command.FieldName);
            ApplyCaption(dataField, command.Caption, command.SetCaption);
            dataField.Position = command.Position;
            ApplyNumberFormat(dataField, command.NumberFormatCode);
        }

        private static dynamic ResolvePlacedField(
            dynamic pivot,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command)
        {
            if (command.Area == CorePivotFieldArea.Values)
            {
                return ResolveDataField(pivot, command.Caption, command.FieldName);
            }

            if (IsClassic(sourceKind))
            {
                return ResolvePivotField(pivot, command.FieldName);
            }

            dynamic cube = ResolveCubeField(pivot, command.FieldName);
            return ResolveVisiblePivotField(pivot, cube, command.FieldName);
        }

        private static dynamic ResolvePivotField(dynamic pivot, string name)
        {
            if (TryRead(() => (object?)pivot.PivotFields.Item(name), out object? field) && field != null)
            {
                return field;
            }

            if (TryRead(() => (object?)pivot.PivotFields(name), out field) && field != null)
            {
                return field;
            }

            foreach (object candidate in ReadCollection((object)pivot, "PivotFields"))
            {
                if (MatchesAnyName(candidate, name)) return candidate;
            }

            throw new InvalidOperationException(
                "Excel could not resolve the native PivotField '" + name + "'.");
        }

        private static dynamic ResolveCubeField(dynamic pivot, string name)
        {
            if (TryRead(() => (object?)pivot.CubeFields.Item(name), out object? field) && field != null)
            {
                return field;
            }

            foreach (object candidate in ReadCollection((object)pivot, "CubeFields"))
            {
                if (MatchesAnyName(candidate, name)) return candidate;
            }

            throw new InvalidOperationException(
                "Excel could not resolve the native CubeField '" + name + "'.");
        }

        private static dynamic ResolveVisiblePivotField(
            dynamic pivot,
            dynamic cubeField,
            string fallbackName)
        {
            var candidates = new[]
            {
                ReadOptionalString(cubeField, "Name"),
                ReadOptionalString(cubeField, "Caption"),
                fallbackName
            };
            foreach (string? candidate in candidates.Where(item => !string.IsNullOrWhiteSpace(item)))
            {
                if (TryRead(
                        () => (object?)pivot.PivotFields.Item(candidate),
                        out object? field) &&
                    field != null)
                {
                    return field;
                }
            }

            foreach (object candidate in ReadCollection((object)pivot, "PivotFields"))
            {
                object? nestedCube = ReadObject((dynamic)candidate, "CubeField");
                if (nestedCube != null &&
                    string.Equals(
                        ReadOptionalString((dynamic)nestedCube, "Name"),
                        ReadOptionalString(cubeField, "Name"),
                        StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException(
                "Excel activated a CubeField but did not expose its visible PivotField.");
        }

        private static dynamic ResolveDataField(
            dynamic pivot,
            string caption,
            string fieldName)
        {
            if (TryRead(() => (object?)pivot.DataFields.Item(caption), out object? field) && field != null)
            {
                return field;
            }

            IReadOnlyList<object> dataFields = ReadCollection((object)pivot, "DataFields");
            List<object> matches = dataFields
                .Where(candidate =>
                    MatchesAnyName(candidate, caption) ||
                    MatchesAnyName(candidate, fieldName) ||
                    string.Equals(
                        ReadOptionalString((dynamic)ReadObject((dynamic)candidate, "CubeField"), "Name"),
                        fieldName,
                        StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 1)
            {
                return matches[0];
            }

            throw new InvalidOperationException(
                "Excel did not expose a unique native Values field for '" + caption + "'.");
        }

        private static void CaptureArea(
            dynamic pivot,
            PivotSourceKind sourceKind,
            CorePivotFieldArea area,
            ICollection<SnapshotField> result)
        {
            string collectionName;
            switch (area)
            {
                case CorePivotFieldArea.Row: collectionName = "RowFields"; break;
                case CorePivotFieldArea.Column: collectionName = "ColumnFields"; break;
                case CorePivotFieldArea.Filter: collectionName = "PageFields"; break;
                case CorePivotFieldArea.Values: collectionName = "DataFields"; break;
                default: throw new ArgumentOutOfRangeException(nameof(area));
            }

            var logicalPosition = 0;
            foreach (object fieldObject in ReadAreaCollection((object)pivot, collectionName))
            {
                logicalPosition++;
                dynamic field = fieldObject;
                object? cubeObject = ReadObject(field, "CubeField");
                string fieldName;
                if (IsClassic(sourceKind))
                {
                    fieldName = ReadRequiredName(
                        () => (object?)field.SourceName,
                        "Excel did not expose a classic PivotField source name required for rollback.");
                }
                else
                {
                    if (cubeObject == null)
                    {
                        throw new NotSupportedException(
                            "Excel did not expose the CubeField identity required for OLAP rollback.");
                    }

                    dynamic cube = cubeObject;
                    fieldName = ReadRequiredName(
                        () => (object?)cube.Name,
                        "Excel did not expose the CubeField name required for OLAP rollback.");
                }

                int nativePosition = ReadRequiredInt(
                    () => (object?)field.Position,
                    "Excel did not expose a PivotField position required for rollback.");
                if (nativePosition < 1)
                {
                    throw new NotSupportedException(
                        "Excel exposed an invalid PivotField position for rollback.");
                }

                result.Add(new SnapshotField
                {
                    FieldName = fieldName,
                    Caption = ReadRequiredName(
                        () => (object?)field.Caption,
                        "Excel did not expose a PivotField caption required for rollback."),
                    Area = area,
                    Position = logicalPosition,
                    IsCubeMeasure = !IsClassic(sourceKind) &&
                                    area == CorePivotFieldArea.Values,
                    ConsolidationFunction = area == CorePivotFieldArea.Values &&
                                              IsClassic(sourceKind)
                        ? ReadRequiredInt(
                            () => (object?)field.Function,
                            "Excel did not expose a Values aggregation required for rollback.")
                        : null,
                    NumberFormatCode = area == CorePivotFieldArea.Values
                        ? ReadRequiredName(
                            () => (object?)field.NumberFormat,
                            "Excel did not expose a Values number format required for rollback.")
                        : null,
                    Subtotals = area == CorePivotFieldArea.Row
                        ? ReadSubtotals((object)field)
                        : new bool[12]
                });
            }
        }

        private static IReadOnlyList<SnapshotCubeField> CaptureCubeFields(dynamic pivot)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<SnapshotCubeField>();
            foreach (object cubeObject in ReadCollection((object)pivot, "CubeFields"))
            {
                dynamic cube = cubeObject;
                string name = ReadRequiredName(
                    () => (object?)cube.Name,
                    "Excel did not expose a CubeField name required for rollback.");
                if (!names.Add(name))
                {
                    throw new NotSupportedException(
                        "Excel exposed duplicate CubeField identities, so implicit-measure rollback cannot be bounded safely.");
                }

                result.Add(new SnapshotCubeField
                {
                    Name = name,
                    Caption = ReadRequiredName(
                        () => (object?)cube.Caption,
                        "Excel did not expose a CubeField caption required for rollback.")
                });
            }

            return result
                .OrderBy(field => field.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void DeleteCubeFieldsCreatedAfterSnapshot(
            object pivotTable,
            IReadOnlyList<SnapshotCubeField> snapshotFields)
        {
            var prior = new HashSet<string>(
                (snapshotFields ?? throw new ArgumentNullException(nameof(snapshotFields)))
                    .Select(field => field.Name),
                StringComparer.OrdinalIgnoreCase);
            dynamic pivot = pivotTable;
            foreach (object cubeObject in ReadCollection((object)pivot, "CubeFields"))
            {
                dynamic cube = cubeObject;
                string name = FirstName(cube, "Name");
                if (prior.Contains(name))
                {
                    continue;
                }

                try
                {
                    cube.Delete();
                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException(
                        "Excel could not delete an implicit CubeField created by the failed PivotTable+ mutation.",
                        exception);
                }
            }

            IReadOnlyList<SnapshotCubeField> remaining = CaptureCubeFields((object)pivot);
            if (remaining.Any(field => !prior.Contains(field.Name)))
            {
                throw new InvalidOperationException(
                    "Excel retained an implicit CubeField created by the failed PivotTable+ mutation.");
            }
        }

        private static void RestoreCubeFieldCaptions(
            object pivotTable,
            IReadOnlyList<SnapshotCubeField> snapshotFields)
        {
            IReadOnlyDictionary<string, object> current = ReadCubeFieldsByName(pivotTable);
            foreach (SnapshotCubeField snapshot in snapshotFields)
            {
                if (!current.TryGetValue(snapshot.Name, out object? cubeObject))
                {
                    throw new InvalidOperationException(
                        "Excel removed a pre-existing CubeField during rollback.");
                }

                dynamic cube = cubeObject;
                cube.Caption = snapshot.Caption;
            }

            VerifyCubeFieldCaptions(pivotTable, snapshotFields);
        }

        private static void VerifyCubeFieldCaptions(
            object pivotTable,
            IReadOnlyList<SnapshotCubeField> snapshotFields)
        {
            IReadOnlyDictionary<string, object> current = ReadCubeFieldsByName(pivotTable);
            foreach (SnapshotCubeField snapshot in snapshotFields)
            {
                if (!current.TryGetValue(snapshot.Name, out object? cubeObject))
                {
                    throw new InvalidOperationException(
                        "Excel removed a pre-existing CubeField during rollback.");
                }

                dynamic cube = cubeObject;
                string caption = ReadRequiredName(
                    () => (object?)cube.Caption,
                    "Excel did not expose a CubeField caption while verifying rollback.");
                DemandEqual(
                    snapshot.Caption,
                    caption,
                    "Excel did not restore a pre-existing CubeField caption.");
            }
        }

        private static IReadOnlyDictionary<string, object> ReadCubeFieldsByName(
            object pivotTable)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (object cubeObject in ReadCollection(pivotTable, "CubeFields"))
            {
                dynamic cube = cubeObject;
                string name = ReadRequiredName(
                    () => (object?)cube.Name,
                    "Excel did not expose a CubeField name while verifying rollback.");
                if (result.ContainsKey(name))
                {
                    throw new InvalidOperationException(
                        "Excel exposed duplicate CubeField identities during rollback.");
                }

                result.Add(name, cubeObject);
            }

            return result;
        }

        private static void DemandNormalValueCalculations(dynamic pivot)
        {
            foreach (object fieldObject in ReadCollection((object)pivot, "DataFields"))
            {
                dynamic field = fieldObject;
                if (!TryRead(() => (object?)field.Calculation, out object? value) || value == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose Show Values As state for an existing Values field.");
                }

                int calculation;
                try
                {
                    calculation = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                }
                catch (Exception exception) when (
                    exception is FormatException ||
                    exception is InvalidCastException ||
                    exception is OverflowException)
                {
                    throw new NotSupportedException(
                        "Excel exposed an invalid Show Values As state.",
                        exception);
                }

                if (calculation != NoAdditionalCalculation)
                {
                    throw new NotSupportedException(
                        "This PivotTable uses Show Values As. Its typed native contract must be preserved before layout mutation is allowed.");
                }
            }
        }

        private static void DemandNoActiveNativeFilters(
            dynamic pivot,
            PivotSourceKind sourceKind)
        {
            IReadOnlyList<object> axisFields = ReadAreaCollection((object)pivot, "RowFields")
                .Concat(ReadAreaCollection((object)pivot, "ColumnFields"))
                .Concat(ReadCollection((object)pivot, "PageFields"))
                .ToList();
            foreach (object fieldObject in axisFields)
            {
                dynamic field = fieldObject;
                if (!TryRead(() => (object?)field.PivotFilters, out object? filters) ||
                    filters == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose native label, value, or date filter state for a visible PivotTable field.");
                }

                if (ReadCollectionCount(filters, "native PivotFilter") > 0)
                {
                    throw new NotSupportedException(
                        "This PivotTable has an active label, value, or date filter that is not yet in the mutation contract.");
                }

                object filterOwnerObject = fieldObject;
                if (!IsClassic(sourceKind))
                {
                    filterOwnerObject = ReadObject(field, "CubeField") ??
                        throw new NotSupportedException(
                            "Excel did not expose the CubeField filter state for a visible OLAP PivotTable field.");
                }

                dynamic filterOwner = filterOwnerObject;
                if (!TryRead(
                        () => (object?)filterOwner.AllItemsVisible,
                        out object? allVisible) ||
                    allVisible == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose native member-filter state for a visible PivotTable field.");
                }

                bool allItemsVisible;
                try
                {
                    allItemsVisible = Convert.ToBoolean(allVisible, CultureInfo.InvariantCulture);
                }
                catch (Exception exception) when (
                    exception is FormatException ||
                    exception is InvalidCastException)
                {
                    throw new NotSupportedException(
                        "Excel exposed invalid native member-filter state.",
                        exception);
                }

                if (!allItemsVisible)
                {
                    throw new NotSupportedException(
                        "This PivotTable has an active member filter that is not yet in the mutation contract.");
                }

                if (TryRead(
                        () => (object?)filterOwner.EnableMultiplePageItems,
                        out object? multiplePageItems) &&
                    multiplePageItems != null &&
                    Convert.ToBoolean(multiplePageItems, CultureInfo.InvariantCulture))
                {
                    throw new NotSupportedException(
                        "This PivotTable has an active multi-select page filter that is not yet in the mutation contract.");
                }
            }
        }

        private static void DemandNoUnsupportedCalculatedObjects(dynamic pivot)
        {
            object calculatedFields = ReadRequiredObject(
                () => (object?)pivot.CalculatedFields(),
                "Excel did not expose the calculated-field collection required for safe mutation.");
            if (ReadCollectionCount(calculatedFields, "calculated field") > 0)
            {
                throw new NotSupportedException(
                    "Classic calculated fields are not part of the safe native mutation contract.");
            }

            IReadOnlyList<object> axisFields = ReadAreaCollection((object)pivot, "RowFields")
                .Concat(ReadAreaCollection((object)pivot, "ColumnFields"))
                .Concat(ReadCollection((object)pivot, "PageFields"))
                .ToList();
            foreach (object fieldObject in axisFields)
            {
                dynamic field = fieldObject;
                object? calculatedItems = null;
                if (!TryRead(
                        () => (object?)field.CalculatedItems(),
                        out calculatedItems) ||
                    calculatedItems == null)
                {
                    TryRead(
                        () => (object?)field.CalculatedItems,
                        out calculatedItems);
                }

                // Excel 2021 and some Microsoft 365 builds raise 1004 instead of
                // returning an empty CalculatedItems collection for an ordinary
                // PivotField. A real calculated item makes the collection readable,
                // so keep blocking non-empty collections but accept the host's
                // documented "not available" shape for an ordinary field.
                if (calculatedItems != null &&
                    ReadCollectionCount(calculatedItems, "calculated item") > 0)
                {
                    throw new NotSupportedException(
                        "Classic calculated items are not part of the safe native mutation contract.");
                }
            }
        }

        private static void DemandUniformRowRepeatLabels(dynamic pivot)
        {
            IReadOnlyList<object> rows = ReadAreaCollection((object)pivot, "RowFields");
            bool? expected = null;
            foreach (object rowObject in rows)
            {
                dynamic row = rowObject;
                bool repeat = ReadRequiredBoolean(
                    () => (object?)row.RepeatLabels,
                    "Excel did not expose repeated-label state for a visible row field.");
                if (expected.HasValue && expected.Value != repeat)
                {
                    throw new NotSupportedException(
                        "This PivotTable has mixed per-row RepeatLabels settings that the global PivotTable+ layout contract cannot preserve yet.");
                }

                expected = repeat;
            }
        }

        private static int ReadCollectionCount(object collectionObject, string label)
        {
            dynamic collection = collectionObject;
            if (!TryRead(() => (object?)collection.Count, out object? value) || value == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the " + label + " collection count.");
            }

            try
            {
                int count = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                if (count < 0 || count > MaximumNativeFields)
                {
                    throw new NotSupportedException(
                        "Excel exposed an invalid " + label + " collection count.");
                }

                return count;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid " + label + " collection count.",
                    exception);
            }
        }

        private static IReadOnlyList<object> ReadCollection(dynamic owner, string memberName)
        {
            object? collectionObject = null;
            switch (memberName)
            {
                case "RowFields":
                    TryRead(() => (object?)owner.RowFields, out collectionObject);
                    if (collectionObject == null) TryRead(() => (object?)owner.RowFields(), out collectionObject);
                    break;
                case "ColumnFields":
                    TryRead(() => (object?)owner.ColumnFields, out collectionObject);
                    if (collectionObject == null) TryRead(() => (object?)owner.ColumnFields(), out collectionObject);
                    break;
                case "PageFields":
                    TryRead(() => (object?)owner.PageFields, out collectionObject);
                    if (collectionObject == null) TryRead(() => (object?)owner.PageFields(), out collectionObject);
                    break;
                case "DataFields":
                    TryRead(() => (object?)owner.DataFields, out collectionObject);
                    if (collectionObject == null) TryRead(() => (object?)owner.DataFields(), out collectionObject);
                    break;
                case "PivotFields":
                    TryRead(() => (object?)owner.PivotFields, out collectionObject);
                    if (collectionObject == null) TryRead(() => (object?)owner.PivotFields(), out collectionObject);
                    break;
                case "CubeFields":
                    TryRead(() => (object?)owner.CubeFields, out collectionObject);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(memberName));
            }

            if (collectionObject == null)
            {
                throw new NotSupportedException(
                    "Excel did not expose the required " + memberName + " collection.");
            }

            dynamic collection = collectionObject;
            int count = ReadCollectionCount(collectionObject, memberName);
            var result = new List<object>(count);
            for (var index = 1; index <= count; index++)
            {
                object? item;
                bool itemRead = TryRead(
                    () => (object?)collection.Item(index),
                    out item);
                if ((!itemRead || item == null) &&
                    !TryRead(() => (object?)collection[index], out item))
                {
                    item = null;
                }

                if (item == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose item " +
                        index.ToString(CultureInfo.InvariantCulture) +
                        " from the required " + memberName + " collection.");
                }

                result.Add(item);
            }

            return result;
        }

        private static IReadOnlyList<object> ReadAreaCollection(
            object pivotTable,
            string memberName)
        {
            IReadOnlyList<object> fields = ReadCollection(pivotTable, memberName);
            if (memberName != "RowFields" && memberName != "ColumnFields")
            {
                return fields;
            }

            dynamic pivot = pivotTable;
            if (!TryRead(() => (object?)pivot.DataPivotField, out object? dataPivotField) ||
                dataPivotField == null)
            {
                return fields;
            }

            return fields
                .Where(field => !ComObjectIdentity.AreSame(field, dataPivotField))
                .ToList();
        }

        private static object? ReadDataPivotField(dynamic pivot, bool required)
        {
            if (TryRead(() => (object?)pivot.DataPivotField, out object? dataPivotField) &&
                dataPivotField != null)
            {
                return dataPivotField;
            }

            if (required)
            {
                throw new NotSupportedException(
                    "Excel did not expose the Values pseudo-field required for safe native mutation.");
            }

            return null;
        }

        private static void ReadValuesAxis(
            dynamic pivot,
            int valueCount,
            out PivotValuesAxis valuesAxis,
            out int valuesPosition)
        {
            valuesAxis = PivotValuesAxis.Automatic;
            valuesPosition = 1;
            object? dataPivotField = ReadDataPivotField(pivot, required: valueCount > 1);
            if (dataPivotField == null) return;

            dynamic field = dataPivotField;
            int orientation = ReadRequiredInt(
                () => (object?)field.Orientation,
                "Excel did not expose the Values pseudo-field orientation required for rollback.");
            switch (orientation)
            {
                case OrientationHidden:
                    if (valueCount > 1)
                    {
                        throw new NotSupportedException(
                            "Excel exposed a hidden Values pseudo-field for a multi-value PivotTable.");
                    }

                    return;
                case OrientationRow:
                    valuesAxis = PivotValuesAxis.Rows;
                    break;
                case OrientationColumn:
                    valuesAxis = PivotValuesAxis.Columns;
                    break;
                default:
                    throw new NotSupportedException(
                        "Excel exposed the Values pseudo-field on an unsupported axis.");
            }

            valuesPosition = ReadRequiredInt(
                () => (object?)field.Position,
                "Excel did not expose the Values pseudo-field position required for rollback.");
            if (valuesPosition < 1)
            {
                throw new NotSupportedException(
                    "Excel exposed an invalid Values pseudo-field position for rollback.");
            }
        }

        private static void ApplyValuesAxis(
            dynamic pivot,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            if (valuesAxis == PivotValuesAxis.Automatic)
            {
                // With zero or one Values field Excel owns the hidden/implicit
                // DataPivotField state.  Some real Excel builds expose that
                // pseudo-field but reject an explicit Orientation = xlHidden
                // write with error 1004.  Automatic therefore means no host
                // write; the concrete Rows/Columns cases below remain exact.
                return;
            }

            object dataPivotField = ReadDataPivotField(pivot, required: true)!;
            dynamic field = dataPivotField;
            switch (valuesAxis)
            {
                case PivotValuesAxis.Rows:
                    field.Orientation = OrientationRow;
                    break;
                case PivotValuesAxis.Columns:
                    field.Orientation = OrientationColumn;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(valuesAxis));
            }

            field.Position = valuesPosition;
        }

        private static void VerifyValuesAxis(
            dynamic pivot,
            PivotValuesAxis expectedAxis,
            int expectedPosition,
            int valueCount)
        {
            object? dataPivotField = ReadDataPivotField(
                pivot,
                required: expectedAxis != PivotValuesAxis.Automatic || valueCount > 1);
            if (expectedAxis == PivotValuesAxis.Automatic)
            {
                // Automatic is valid only for zero/one Values fields.  Excel
                // may expose its implicit DataPivotField with a host-specific
                // orientation even though no Values axis is visible, so there
                // is no stable orientation value to assert here.
                return;
            }

            dynamic field = dataPivotField!;
            DemandEqual(
                expectedAxis == PivotValuesAxis.Rows ? OrientationRow : OrientationColumn,
                ReadRequiredInt(
                    () => (object?)field.Orientation,
                    "Excel did not expose the Values pseudo-field orientation during verification."),
                "Excel applied the wrong Values pseudo-field axis.");
            DemandEqual(
                expectedPosition,
                ReadRequiredInt(
                    () => (object?)field.Position,
                    "Excel did not expose the Values pseudo-field position during verification."),
                "Excel applied the wrong Values pseudo-field position.");
        }

        private static int ReadNormalizedAreaPosition(
            object pivotTable,
            CorePivotFieldArea area,
            object field)
        {
            string collectionName = area == CorePivotFieldArea.Row
                ? "RowFields"
                : "ColumnFields";
            IReadOnlyList<object> fields = ReadAreaCollection(pivotTable, collectionName);
            for (var index = 0; index < fields.Count; index++)
            {
                if (ComObjectIdentity.AreSame(fields[index], field))
                {
                    return index + 1;
                }
            }

            // Excel can return a different RCW for PivotFields.Item(...) than
            // it returns while enumerating RowFields/ColumnFields.  IUnknown
            // identity is normally stable, but some Office 2021 builds proxy
            // the two collection paths independently.  Fall back to the
            // native source/CubeField identity, requiring one exact match so
            // verification remains fail-closed rather than caption-based.
            dynamic expectedField = field;
            string? expectedSourceName = ReadOptionalString(expectedField, "SourceName");
            object? expectedCube = ReadObject(expectedField, "CubeField");
            string? expectedCubeName = expectedCube == null
                ? null
                : ReadOptionalString((dynamic)expectedCube, "Name");
            var stableMatches = new List<int>();
            for (var index = 0; index < fields.Count; index++)
            {
                dynamic candidate = fields[index];
                string? candidateSourceName = ReadOptionalString(candidate, "SourceName");
                object? candidateCube = ReadObject(candidate, "CubeField");
                string? candidateCubeName = candidateCube == null
                    ? null
                    : ReadOptionalString((dynamic)candidateCube, "Name");
                bool sourceMatch = !string.IsNullOrWhiteSpace(expectedSourceName) &&
                    string.Equals(
                        expectedSourceName,
                        candidateSourceName,
                        StringComparison.OrdinalIgnoreCase);
                bool cubeMatch = !string.IsNullOrWhiteSpace(expectedCubeName) &&
                    string.Equals(
                        expectedCubeName,
                        candidateCubeName,
                        StringComparison.OrdinalIgnoreCase);
                if (sourceMatch || cubeMatch)
                {
                    stableMatches.Add(index + 1);
                }
            }

            if (stableMatches.Count == 1)
            {
                return stableMatches[0];
            }

            throw new InvalidOperationException(
                "Excel did not expose the placed PivotField in its expected axis collection.");
        }

        private static object ReadPivotCache(dynamic pivot)
        {
            if (TryRead(() => (object?)pivot.PivotCache(), out object? cache) &&
                cache != null)
            {
                return cache;
            }

            if (TryRead(() => (object?)pivot.PivotCache, out cache) && cache != null)
            {
                return cache;
            }

            throw new NotSupportedException(
                "Excel did not expose the selected PivotTable's PivotCache for live source validation.");
        }

        private static string ReadRequiredClassicSourceName(object cacheObject)
        {
            dynamic cache = cacheObject;
            object sourceData = ReadRequiredObject(
                () => (object?)cache.SourceData,
                "Excel did not expose the selected classic PivotCache SourceData.");
            string sourceName;
            if (sourceData is string text)
            {
                sourceName = text;
            }
            else if (sourceData is Array array && array.Length == 1)
            {
                object? item = null;
                foreach (object? candidate in array)
                {
                    item = candidate;
                    break;
                }

                sourceName = Convert.ToString(item, CultureInfo.InvariantCulture) ??
                    string.Empty;
            }
            else
            {
                throw new NotSupportedException(
                    "Excel exposed an unsupported classic PivotCache SourceData shape.");
            }

            DemandSafeSourceName(sourceName);
            return sourceName;
        }

        private static void DemandSafeSourceName(string sourceName)
        {
            if (string.IsNullOrWhiteSpace(sourceName) ||
                sourceName.Length > 255 ||
                sourceName != sourceName.Trim() ||
                sourceName.Any(char.IsControl) ||
                sourceName.IndexOf('\\') >= 0 ||
                sourceName.IndexOf('/') >= 0 ||
                sourceName.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ||
                (sourceName.Length >= 2 &&
                 char.IsLetter(sourceName[0]) &&
                 sourceName[1] == ':'))
            {
                throw new NotSupportedException(
                    "Excel did not expose a path-safe PivotCache source identity.");
            }
        }

        private static bool[] ReadSubtotals(dynamic field)
        {
            var result = new bool[12];
            for (var index = 1; index <= result.Length; index++)
            {
                int captured = index;
                if (!TryRead(
                        () => (object?)field.Subtotals[captured],
                        out object? value) ||
                    value == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose all 12 PivotTable subtotal slots.");
                }

                try
                {
                    result[index - 1] = Convert.ToBoolean(
                        value,
                        CultureInfo.InvariantCulture);
                }
                catch (Exception exception) when (
                    exception is FormatException ||
                    exception is InvalidCastException)
                {
                    throw new NotSupportedException(
                        "Excel exposed an invalid PivotTable subtotal slot.",
                        exception);
                }
            }

            return result;
        }

        private static void ApplySubtotals(dynamic field, PivotSubtotalMode mode)
        {
            var subtotals = new bool[12];
            subtotals[0] = mode == PivotSubtotalMode.Automatic;
            WriteSubtotals(field, subtotals);
        }

        private static void WriteSubtotals(dynamic field, IReadOnlyList<bool> subtotals)
        {
            if (subtotals == null || subtotals.Count != 12)
            {
                throw new ArgumentException(
                    "Excel PivotTable subtotal state must contain exactly 12 slots.",
                    nameof(subtotals));
            }

            for (var index = 1; index <= subtotals.Count; index++)
            {
                field.Subtotals[index] = subtotals[index - 1];
            }
        }

        private static void ApplyCaption(dynamic field, string caption, bool setCaption)
        {
            if (setCaption && !string.IsNullOrWhiteSpace(caption))
            {
                field.Caption = caption;
            }
        }

        private static void ApplyNumberFormat(dynamic field, string? numberFormatCode)
        {
            if (!string.IsNullOrWhiteSpace(numberFormatCode))
            {
                field.NumberFormat = numberFormatCode;
            }
        }

        private static int Orientation(CorePivotFieldArea area)
        {
            switch (area)
            {
                case CorePivotFieldArea.Row: return OrientationRow;
                case CorePivotFieldArea.Column: return OrientationColumn;
                case CorePivotFieldArea.Filter: return OrientationPage;
                case CorePivotFieldArea.Values: return OrientationData;
                default: throw new ArgumentOutOfRangeException(nameof(area));
            }
        }

        private static bool IsClassic(PivotSourceKind sourceKind)
        {
            return sourceKind == PivotSourceKind.WorksheetRange ||
                   sourceKind == PivotSourceKind.WorksheetTable;
        }

        private static int AreaOrder(CorePivotFieldArea area)
        {
            switch (area)
            {
                case CorePivotFieldArea.Row: return 0;
                case CorePivotFieldArea.Column: return 1;
                case CorePivotFieldArea.Filter: return 2;
                case CorePivotFieldArea.Values: return 3;
                default: throw new ArgumentOutOfRangeException(nameof(area));
            }
        }

        private static bool ReadRepeatLabels(dynamic pivot)
        {
            object? row = ReadAreaCollection((object)pivot, "RowFields").FirstOrDefault();
            if (row == null)
            {
                return false;
            }

            dynamic rowField = row;
            return ReadRequiredBoolean(
                () => (object?)rowField.RepeatLabels,
                "Excel did not expose repeated-label state required for rollback.");
        }

        private static object? ReadObject(dynamic owner, string memberName)
        {
            if (owner == null) return null;
            switch (memberName)
            {
                case "CubeField":
                    return TryRead(() => (object?)owner.CubeField, out object? cube) ? cube : null;
                case "PivotField":
                    return TryRead(() => (object?)owner.PivotField, out object? pivotField)
                        ? pivotField
                        : null;
                default:
                    throw new ArgumentOutOfRangeException(nameof(memberName));
            }
        }

        private static string FirstName(dynamic owner, params string[] memberNames)
        {
            foreach (string memberName in memberNames)
            {
                string? value = ReadOptionalString(owner, memberName);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }

            throw new InvalidOperationException("Excel exposed an unnamed PivotTable field.");
        }

        private static bool MatchesAnyName(object candidate, string expected)
        {
            dynamic field = candidate;
            return new[]
                {
                    ReadOptionalString(field, "Name"),
                    ReadOptionalString(field, "Caption"),
                    ReadOptionalString(field, "SourceName")
                }
                .Any(value => string.Equals(value, expected, StringComparison.OrdinalIgnoreCase));
        }

        private static string? ReadOptionalString(dynamic owner, string memberName)
        {
            if (owner == null) return null;
            object? value = null;
            switch (memberName)
            {
                case "Name": TryRead(() => (object?)owner.Name, out value); break;
                case "Caption": TryRead(() => (object?)owner.Caption, out value); break;
                case "SourceName": TryRead(() => (object?)owner.SourceName, out value); break;
                case "NumberFormat": TryRead(() => (object?)owner.NumberFormat, out value); break;
                case "TableStyle2": return ReadPivotTableStyleName(owner);
                default: throw new ArgumentOutOfRangeException(nameof(memberName));
            }

            return value == null
                ? null
                : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string? ReadPivotTableStyleName(dynamic pivot)
        {
            if (!TryRead(() => (object?)pivot.TableStyle2, out object? value) ||
                value == null)
            {
                return null;
            }

            if (value is string text)
            {
                return text;
            }

            // Depending on the Office PIA/build, TableStyle2 can late-bind as
            // the style's name string or as a TableStyle RCW.  Persist and
            // compare the native Name in either case.
            dynamic style = value;
            if (TryRead(() => (object?)style.Name, out object? name) && name != null)
            {
                string? resolvedName = Convert.ToString(name, CultureInfo.InvariantCulture);
                if (!string.Equals(
                        resolvedName,
                        "System.__ComObject",
                        StringComparison.Ordinal))
                {
                    return resolvedName;
                }
            }

            string? fallback = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.Equals(fallback, "System.__ComObject", StringComparison.Ordinal)
                ? null
                : fallback;
        }

        private static int ReadInt(dynamic owner, string memberName, int fallback)
        {
            object? value = null;
            switch (memberName)
            {
                case "Position": TryRead(() => (object?)owner.Position, out value); break;
                case "Orientation": TryRead(() => (object?)owner.Orientation, out value); break;
                case "Function": TryRead(() => (object?)owner.Function, out value); break;
                case "LayoutRowDefault": TryRead(() => (object?)owner.LayoutRowDefault, out value); break;
                default: throw new ArgumentOutOfRangeException(nameof(memberName));
            }

            if (value == null) return fallback;
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return fallback;
            }
        }

        private static int? ReadNullableInt(dynamic owner, string memberName)
        {
            int value = ReadInt(owner, memberName, int.MinValue);
            return value == int.MinValue ? (int?)null : value;
        }

        private static bool ReadBoolean(dynamic owner, string memberName, bool fallback)
        {
            object? value = null;
            switch (memberName)
            {
                case "RowGrand": TryRead(() => (object?)owner.RowGrand, out value); break;
                case "ColumnGrand": TryRead(() => (object?)owner.ColumnGrand, out value); break;
                case "DisplayFieldCaptions": TryRead(() => (object?)owner.DisplayFieldCaptions, out value); break;
                case "PreserveFormatting": TryRead(() => (object?)owner.PreserveFormatting, out value); break;
                case "ShowTableStyleRowStripes": TryRead(() => (object?)owner.ShowTableStyleRowStripes, out value); break;
                case "ShowTableStyleColumnStripes": TryRead(() => (object?)owner.ShowTableStyleColumnStripes, out value); break;
                case "RepeatLabels": TryRead(() => (object?)owner.RepeatLabels, out value); break;
                default: throw new ArgumentOutOfRangeException(nameof(memberName));
            }

            if (value == null) return fallback;
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException)
            {
                return fallback;
            }
        }

        private static bool TryRead(Func<object?> reader, out object? value)
        {
            try
            {
                value = reader();
                return true;
            }
            catch (Exception)
            {
                value = null;
                return false;
            }
        }

        private static object ReadRequiredObject(
            Func<object?> reader,
            string failureMessage)
        {
            if (!TryRead(reader, out object? value) || value == null)
            {
                throw new InvalidOperationException(failureMessage);
            }

            return value;
        }

        private static string ReadRequiredName(
            Func<object?> reader,
            string failureMessage)
        {
            object value = ReadRequiredObject(reader, failureMessage);
            string? name = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException(failureMessage);
            }

            return name;
        }

        private static string? ReadRequiredOptionalString(
            Func<object?> reader,
            string failureMessage)
        {
            if (!TryRead(reader, out object? value))
            {
                throw new NotSupportedException(failureMessage);
            }

            return value == null
                ? null
                : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static int ReadRequiredInt(
            Func<object?> reader,
            string failureMessage)
        {
            object value = ReadRequiredObject(reader, failureMessage);
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new NotSupportedException(failureMessage, exception);
            }
        }

        private static bool ReadRequiredBoolean(
            Func<object?> reader,
            string failureMessage)
        {
            object value = ReadRequiredObject(reader, failureMessage);
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException)
            {
                throw new NotSupportedException(failureMessage, exception);
            }
        }

        private static void DemandEqual<T>(T expected, T actual, string message)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                throw new InvalidOperationException(message);
            }
        }

        private sealed class NativePivotSnapshot
        {
            public PivotSourceKind SourceKind { get; set; }

            public IReadOnlyList<SnapshotField> Fields { get; set; } = Array.Empty<SnapshotField>();

            public IReadOnlyList<SnapshotCubeField> CubeFields { get; set; } =
                Array.Empty<SnapshotCubeField>();

            public NativePivotLayoutCommand Layout { get; set; } = new NativePivotLayoutCommand();
        }

        private sealed class SnapshotCubeField
        {
            public string Name { get; set; } = string.Empty;

            public string Caption { get; set; } = string.Empty;
        }

        private sealed class SnapshotField
        {
            public string FieldName { get; set; } = string.Empty;

            public string Caption { get; set; } = string.Empty;

            public CorePivotFieldArea Area { get; set; }

            public int Position { get; set; }

            public bool IsCubeMeasure { get; set; }

            public int? ConsolidationFunction { get; set; }

            public string? NumberFormatCode { get; set; }

            public bool[] Subtotals { get; set; } = new bool[12];
        }
    }
}
