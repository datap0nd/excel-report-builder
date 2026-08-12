using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Excel.Persistence;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelReportBuilder.Excel.PivotPlus
{
    internal interface IActivePivotTableAccessor
    {
        bool TryGetActivePivotTable(object excelApplication, out object? pivotTable);
    }

    internal sealed class LateBoundActivePivotTableAccessor : IActivePivotTableAccessor
    {
        public bool TryGetActivePivotTable(object excelApplication, out object? pivotTable)
        {
            pivotTable = null;
            dynamic application = excelApplication;
            if (!PivotLateBound.TryRead(() => (object?)application.ActiveCell, out object? activeCell) ||
                activeCell == null)
            {
                return false;
            }

            dynamic cell = activeCell;
            return PivotLateBound.TryRead(() => (object?)cell.PivotTable, out pivotTable) &&
                   pivotTable != null;
        }
    }

    internal interface IWorkbookIdentityResolver
    {
        string Resolve(object workbook);

        void Persist(object workbook, string expectedWorkbookId);
    }

    internal sealed class StoredWorkbookIdentityResolver : IWorkbookIdentityResolver
    {
        private sealed class SessionIdentity
        {
            public SessionIdentity(object workbook, string value)
            {
                Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
                Value = value;
            }

            public object Workbook { get; }

            public string Value { get; }
        }

        private const int MaximumSessionWorkbooks = 64;
        private static readonly object SessionIdentityGate = new object();
        private static readonly List<SessionIdentity> SessionIdentities =
            new List<SessionIdentity>();
        private static readonly ConditionalWeakTable<object, SessionIdentity> ManagedSessionIdentities =
            new ConditionalWeakTable<object, SessionIdentity>();
        private readonly WorkbookIdentityStore store = new WorkbookIdentityStore();

        public string Resolve(object workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            string? stored = store.Load((dynamic)workbook);
            if (!string.IsNullOrWhiteSpace(stored))
            {
                return stored!;
            }

            // Managed test doubles and hosts already have stable object identity. Keeping
            // them in the COM registry would retain every short-lived workbook instance.
            if (!Marshal.IsComObject(workbook))
            {
                return ManagedSessionIdentities.GetValue(
                    workbook,
                    key => new SessionIdentity(
                        key,
                        "workbook_" + Guid.NewGuid().ToString("N"))).Value;
            }

            lock (SessionIdentityGate)
            {
                SessionIdentity? existing = SessionIdentities.FirstOrDefault(item =>
                    ReferenceEquals(item.Workbook, workbook) ||
                    ComObjectIdentity.AreSame(item.Workbook, workbook));
                if (existing != null)
                {
                    return existing.Value;
                }

                if (SessionIdentities.Count >= MaximumSessionWorkbooks)
                {
                    throw new InvalidOperationException(
                        "PivotTable+ has reached its bounded workbook-session limit. Restart Excel before opening more workbooks.");
                }

                var created = new SessionIdentity(
                    workbook,
                    "workbook_" + Guid.NewGuid().ToString("N"));
                SessionIdentities.Add(created);
                return created.Value;
            }
        }

        public void Persist(object workbook, string expectedWorkbookId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (expectedWorkbookId == null)
            {
                throw new ArgumentNullException(nameof(expectedWorkbookId));
            }

            string resolved = Resolve(workbook);
            if (!string.Equals(resolved, expectedWorkbookId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The workbook identity changed after PivotTable discovery.");
            }

            store.Ensure((dynamic)workbook, expectedWorkbookId);
        }
    }

    internal enum PivotConnectionKind
    {
        Classic,
        DataModel,
        ExternalOlap,
        DisconnectedOlap
    }

    internal enum DiscoveredPivotFieldKind
    {
        Regular,
        Hierarchy,
        Measure,
        Set,
        Unknown
    }

    internal sealed class DiscoveredPivotField
    {
        public DiscoveredPivotField(
            string name,
            string caption,
            string sourceName,
            DiscoveredPivotFieldKind kind,
            int position,
            PivotAggregationFunction? aggregation,
            string? numberFormatCode,
            bool isCalculated,
            bool? repeatLabels)
        {
            Name = name;
            Caption = caption;
            SourceName = sourceName;
            Kind = kind;
            Position = position;
            Aggregation = aggregation;
            NumberFormatCode = numberFormatCode;
            IsCalculated = isCalculated;
            RepeatLabels = repeatLabels;
        }

        public string Name { get; }

        public string Caption { get; }

        public string SourceName { get; }

        public string StableName => string.IsNullOrWhiteSpace(SourceName) ? Name : SourceName;

        public DiscoveredPivotFieldKind Kind { get; }

        public int Position { get; }

        public PivotAggregationFunction? Aggregation { get; }

        public string? NumberFormatCode { get; }

        public bool IsCalculated { get; }

        public bool? RepeatLabels { get; }
    }

    internal sealed class DiscoveredPivotLayout
    {
        public DiscoveredPivotLayout(
            IReadOnlyList<DiscoveredPivotField> rows,
            IReadOnlyList<DiscoveredPivotField> columns,
            IReadOnlyList<DiscoveredPivotField> values,
            IReadOnlyList<DiscoveredPivotField> filters,
            PivotValuesAxis valuesAxis,
            int valuesPosition)
        {
            Rows = rows;
            Columns = columns;
            Values = values;
            Filters = filters;
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
        }

        public IReadOnlyList<DiscoveredPivotField> Rows { get; }

        public IReadOnlyList<DiscoveredPivotField> Columns { get; }

        public IReadOnlyList<DiscoveredPivotField> Values { get; }

        public IReadOnlyList<DiscoveredPivotField> Filters { get; }

        public PivotValuesAxis ValuesAxis { get; }

        public int ValuesPosition { get; }

        public IEnumerable<DiscoveredPivotField> AllFields =>
            Rows.Concat(Columns).Concat(Values).Concat(Filters);
    }

    public sealed class PivotTableContextDiscovery
    {
        private const int ConnectionTypeModel = 7;
        private const int PivotSourceTypeDatabase = 1;
        private readonly IActivePivotTableAccessor activePivotTableAccessor;
        private readonly IWorkbookIdentityResolver workbookIdentityResolver;

        public PivotTableContextDiscovery()
            : this(
                new LateBoundActivePivotTableAccessor(),
                new StoredWorkbookIdentityResolver())
        {
        }

        internal PivotTableContextDiscovery(
            IActivePivotTableAccessor activePivotTableAccessor,
            IWorkbookIdentityResolver workbookIdentityResolver)
        {
            this.activePivotTableAccessor = activePivotTableAccessor ??
                throw new ArgumentNullException(nameof(activePivotTableAccessor));
            this.workbookIdentityResolver = workbookIdentityResolver ??
                throw new ArgumentNullException(nameof(workbookIdentityResolver));
        }

        public PivotTableContext Discover(object excelApplication)
        {
            if (excelApplication == null)
            {
                throw new ArgumentNullException(nameof(excelApplication));
            }

            if (!activePivotTableAccessor.TryGetActivePivotTable(excelApplication, out object? pivotObject) ||
                pivotObject == null)
            {
                throw new InvalidOperationException(
                    "Select a cell inside a PivotTable before opening PivotTable Plus.");
            }

            dynamic pivot = pivotObject;
            object cacheObject = ReadPivotCache(pivot);
            PivotConnectionKind connectionKind = Classify(cacheObject);
            PivotTargetIdentity target = ReadTarget(pivot);
            DiscoveredPivotLayout layout = ReadLayout(pivot);
            IReadOnlyList<DiscoveredPivotField> discoveredSourceFields = ReadSourceFields(
                pivot,
                connectionKind,
                layout,
                out bool sourceFieldsComplete);

            IReadOnlyList<DiscoveredPivotField> completeInventory = CompleteInventory(
                discoveredSourceFields,
                layout);
            IReadOnlyList<PivotFieldDescriptor> fields = MapFields(
                completeInventory,
                connectionKind);
            PivotSourceDescriptor source = ReadSourceDescriptor(
                cacheObject,
                pivot,
                connectionKind);
            IReadOnlyList<PivotFieldPlacement> placements = MapPlacements(layout, fields);

            var definition = new PivotLayoutDefinition(
                target,
                source,
                fields,
                placements,
                layout: ReadLayoutMetadata(pivot, layout),
                format: ReadFormatMetadata(pivot),
                clearAll: placements.Count == 0);

            return new PivotTableContext(
                definition,
                isConnected: connectionKind != PivotConnectionKind.DisconnectedOlap,
                sourceFieldsComplete: sourceFieldsComplete);
        }

        private static object ReadPivotCache(dynamic pivot)
        {
            if (PivotLateBound.TryRead(() => (object?)pivot.PivotCache(), out object? cache) &&
                cache != null)
            {
                return cache;
            }

            if (PivotLateBound.TryRead(() => (object?)pivot.PivotCache, out cache) &&
                cache != null)
            {
                return cache;
            }

            throw new InvalidOperationException(
                "Excel did not expose the active PivotTable's PivotCache.");
        }

        private static PivotConnectionKind Classify(object cacheObject)
        {
            dynamic cache = cacheObject;
            if (!PivotLateBound.TryRead(() => (object?)cache.OLAP, out object? olapValue))
            {
                throw new InvalidOperationException(
                    "Excel did not expose whether the active PivotTable uses an OLAP source.");
            }

            bool isOlap;
            try
            {
                isOlap = Convert.ToBoolean(olapValue, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException)
            {
                throw new InvalidOperationException(
                    "Excel returned an invalid OLAP status for the active PivotTable.",
                    exception);
            }

            if (!isOlap)
            {
                if (!PivotLateBound.TryRead(
                        () => (object?)cache.SourceType,
                        out object? sourceTypeValue) ||
                    sourceTypeValue == null)
                {
                    throw new NotSupportedException(
                        "Excel did not expose the non-OLAP PivotCache source type, so PivotTable+ will not assume it is worksheet-backed.");
                }

                int sourceType;
                try
                {
                    sourceType = Convert.ToInt32(sourceTypeValue, CultureInfo.InvariantCulture);
                }
                catch (Exception exception) when (
                    exception is FormatException ||
                    exception is InvalidCastException ||
                    exception is OverflowException)
                {
                    throw new NotSupportedException(
                        "Excel exposed an invalid non-OLAP PivotCache source type.",
                        exception);
                }

                if (sourceType != PivotSourceTypeDatabase)
                {
                    throw new NotSupportedException(
                        "PivotTable+ currently supports worksheet table and range PivotCaches; consolidation, scenario, and PivotTable-derived caches are not treated as upgradeable worksheet sources.");
                }

                return PivotConnectionKind.Classic;
            }

            if (!TryReadWorkbookConnection(cacheObject, out object? connection) ||
                connection == null)
            {
                return PivotConnectionKind.DisconnectedOlap;
            }

            dynamic workbookConnection = connection;
            if (!PivotLateBound.TryRead(
                    () => (object?)workbookConnection.Type,
                    out object? connectionTypeValue))
            {
                return PivotConnectionKind.DisconnectedOlap;
            }

            try
            {
                return Convert.ToInt32(connectionTypeValue, CultureInfo.InvariantCulture) ==
                       ConnectionTypeModel
                    ? PivotConnectionKind.DataModel
                    : PivotConnectionKind.ExternalOlap;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "Excel returned an invalid connection type for the active PivotTable.",
                    exception);
            }
        }

        private static bool TryReadWorkbookConnection(
            object cacheObject,
            out object? connection)
        {
            dynamic cache = cacheObject;
            return PivotLateBound.TryRead(
                () => (object?)cache.WorkbookConnection,
                out connection);
        }

        private PivotTargetIdentity ReadTarget(dynamic pivot)
        {
            string pivotName = ReadRequiredName(
                () => (object?)pivot.Name,
                "Excel did not expose the active PivotTable's name.");
            object worksheetObject = ReadRequiredObject(
                () => (object?)pivot.Parent,
                "Excel did not expose the worksheet containing the active PivotTable.");
            dynamic worksheet = worksheetObject;
            string worksheetName = ReadRequiredName(
                () => (object?)worksheet.Name,
                "Excel did not expose the active PivotTable worksheet's name.");
            object workbookObject = ReadRequiredObject(
                () => (object?)worksheet.Parent,
                "Excel did not expose the workbook containing the active PivotTable.");
            string workbookId = workbookIdentityResolver.Resolve(workbookObject);
            if (string.IsNullOrWhiteSpace(workbookId))
            {
                throw new InvalidOperationException(
                    "Excel did not expose a path-free identity for the active PivotTable workbook.");
            }

            return new PivotTargetIdentity(workbookId, worksheetName, pivotName);
        }

        private static PivotSourceDescriptor ReadSourceDescriptor(
            object cacheObject,
            dynamic pivot,
            PivotConnectionKind connectionKind)
        {
            string fallbackName = ReadRequiredName(
                () => (object?)pivot.Name,
                "Excel did not expose a path-free PivotTable source identity.");
            string sourceName = connectionKind == PivotConnectionKind.Classic
                ? ReadClassicSourceName(cacheObject)
                : ReadConnectionName(cacheObject);
            if (!IsSafeSourceName(sourceName))
            {
                sourceName = fallbackName;
            }

            PivotSourceKind sourceKind;
            switch (connectionKind)
            {
                case PivotConnectionKind.Classic:
                    sourceKind = IsWorkbookTableName(pivot, sourceName)
                        ? PivotSourceKind.WorksheetTable
                        : PivotSourceKind.WorksheetRange;
                    break;
                case PivotConnectionKind.DataModel:
                    sourceKind = PivotSourceKind.DataModel;
                    break;
                case PivotConnectionKind.ExternalOlap:
                case PivotConnectionKind.DisconnectedOlap:
                    sourceKind = PivotSourceKind.ExternalOlap;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(connectionKind));
            }

            return new PivotSourceDescriptor(
                sourceKind,
                sourceName,
                CapabilitiesFor(connectionKind));
        }

        private static bool IsWorkbookTableName(dynamic pivot, string sourceName)
        {
            string candidate = sourceName;
            int qualifier = candidate.LastIndexOf('!');
            if (qualifier >= 0 && qualifier + 1 < candidate.Length)
            {
                candidate = candidate.Substring(qualifier + 1);
            }

            candidate = candidate.Trim().Trim('\'');
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            if (!PivotLateBound.TryRead(() => (object?)pivot.Parent, out object? worksheet) ||
                worksheet == null)
            {
                return false;
            }

            dynamic sheet = worksheet;
            if (!PivotLateBound.TryRead(() => (object?)sheet.Parent, out object? workbook) ||
                workbook == null)
            {
                return false;
            }

            dynamic book = workbook;
            object? worksheets = ReadCollectionMember(
                () => (object?)book.Worksheets,
                () => (object?)book.Worksheets());
            if (worksheets == null || !TryReadCollectionCount(worksheets, out int worksheetCount))
            {
                return false;
            }

            dynamic worksheetCollection = worksheets;
            for (var worksheetIndex = 1; worksheetIndex <= worksheetCount; worksheetIndex++)
            {
                object? currentWorksheet = ReadCollectionItem(worksheetCollection, worksheetIndex);
                if (currentWorksheet == null)
                {
                    continue;
                }

                dynamic currentSheet = currentWorksheet;
                object? tables = ReadCollectionMember(
                    () => (object?)currentSheet.ListObjects,
                    () => (object?)currentSheet.ListObjects());
                if (tables == null || !TryReadCollectionCount(tables, out int tableCount))
                {
                    continue;
                }

                dynamic tableCollection = tables;
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    object? table = ReadCollectionItem(tableCollection, tableIndex);
                    if (table == null)
                    {
                        continue;
                    }

                    dynamic listObject = table;
                    string name = ReadOptionalString(() => (object?)listObject.Name);
                    string displayName = ReadOptionalString(() => (object?)listObject.DisplayName);
                    if (string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(candidate, displayName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string ReadClassicSourceName(object cacheObject)
        {
            dynamic cache = cacheObject;
            if (!PivotLateBound.TryRead(() => (object?)cache.SourceData, out object? value) ||
                value == null)
            {
                return string.Empty;
            }

            if (value is string text)
            {
                return text;
            }

            if (value is Array array && array.Length == 1)
            {
                foreach (object? item in array)
                {
                    return Convert.ToString(item, CultureInfo.InvariantCulture) ?? string.Empty;
                }
            }

            return string.Empty;
        }

        private static string ReadConnectionName(object cacheObject)
        {
            if (!TryReadWorkbookConnection(cacheObject, out object? connection) ||
                connection == null)
            {
                return string.Empty;
            }

            dynamic workbookConnection = connection;
            return ReadOptionalString(() => (object?)workbookConnection.Name);
        }

        private static bool IsSafeSourceName(string value)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value.Length > 255 ||
                value != value.Trim() ||
                value.Any(char.IsControl))
            {
                return false;
            }

            return value.IndexOf('\\') < 0 &&
                   value.IndexOf('/') < 0 &&
                   !value.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
                   !(value.Length >= 2 && char.IsLetter(value[0]) && value[1] == ':');
        }

        private static PivotCapability CapabilitiesFor(PivotConnectionKind connectionKind)
        {
            const PivotCapability connectedBase =
                PivotCapability.NativeFieldPlacement |
                PivotCapability.MemberFiltering |
                PivotCapability.LayoutFormatting |
                PivotCapability.ShowValuesAs |
                PivotCapability.Refresh;

            switch (connectionKind)
            {
                case PivotConnectionKind.Classic:
                    return connectedBase | PivotCapability.UpgradeToDataModel;
                case PivotConnectionKind.DataModel:
                    return connectedBase |
                           PivotCapability.DistinctCount |
                           PivotCapability.DataModel |
                           PivotCapability.ModelMeasures |
                           PivotCapability.CalculatedMembers |
                           PivotCapability.NamedSets |
                           PivotCapability.AsymmetricAxes;
                case PivotConnectionKind.ExternalOlap:
                    return connectedBase |
                           PivotCapability.CalculatedMembers |
                           PivotCapability.NamedSets |
                           PivotCapability.AsymmetricAxes;
                case PivotConnectionKind.DisconnectedOlap:
                    return PivotCapability.None;
                default:
                    throw new ArgumentOutOfRangeException(nameof(connectionKind));
            }
        }

        private static DiscoveredPivotLayout ReadLayout(dynamic pivot)
        {
            IReadOnlyList<DiscoveredPivotField> values =
                ReadLayoutArea(pivot, PivotFieldArea.Values, excludedField: null);
            object? dataPivotField = ReadValuesAxis(
                pivot,
                values.Count,
                out PivotValuesAxis valuesAxis,
                out int valuesPosition);
            return new DiscoveredPivotLayout(
                ReadLayoutArea(pivot, PivotFieldArea.Row, dataPivotField),
                ReadLayoutArea(pivot, PivotFieldArea.Column, dataPivotField),
                values,
                ReadLayoutArea(pivot, PivotFieldArea.Filter, excludedField: null),
                valuesAxis,
                valuesPosition);
        }

        private static IReadOnlyList<DiscoveredPivotField> ReadLayoutArea(
            dynamic pivot,
            PivotFieldArea area,
            object? excludedField)
        {
            object? collection;
            string areaLabel;
            switch (area)
            {
                case PivotFieldArea.Row:
                    areaLabel = "row";
                    collection = ReadCollectionMember(
                        () => (object?)pivot.RowFields,
                        () => (object?)pivot.RowFields());
                    break;
                case PivotFieldArea.Column:
                    areaLabel = "column";
                    collection = ReadCollectionMember(
                        () => (object?)pivot.ColumnFields,
                        () => (object?)pivot.ColumnFields());
                    break;
                case PivotFieldArea.Values:
                    areaLabel = "value";
                    collection = ReadCollectionMember(
                        () => (object?)pivot.DataFields,
                        () => (object?)pivot.DataFields());
                    break;
                case PivotFieldArea.Filter:
                    areaLabel = "filter";
                    collection = ReadCollectionMember(
                        () => (object?)pivot.PageFields,
                        () => (object?)pivot.PageFields());
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(area));
            }

            if (collection == null)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the active PivotTable's " + areaLabel + " fields.");
            }

            return ReadFields(
                collection,
                (field, index) => ReadField(
                    field,
                    index,
                    readNestedCubeField: true,
                    useCollectionPosition: true),
                excludedField);
        }

        private static object? ReadValuesAxis(
            dynamic pivot,
            int valueCount,
            out PivotValuesAxis valuesAxis,
            out int valuesPosition)
        {
            valuesAxis = PivotValuesAxis.Automatic;
            valuesPosition = 1;
            if (!PivotLateBound.TryRead(
                    () => (object?)pivot.DataPivotField,
                    out object? dataPivotField) ||
                dataPivotField == null)
            {
                if (valueCount > 1)
                {
                    throw new InvalidOperationException(
                        "Excel did not expose the Values pseudo-field for a multi-value PivotTable.");
                }

                return null;
            }

            dynamic field = dataPivotField;
            object orientationValue = ReadRequiredObject(
                () => (object?)field.Orientation,
                "Excel did not expose the Values pseudo-field orientation.");
            int orientation;
            try
            {
                orientation = Convert.ToInt32(orientationValue, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "Excel exposed an invalid Values pseudo-field orientation.",
                    exception);
            }

            switch (orientation)
            {
                case 0:
                    if (valueCount > 1)
                    {
                        throw new InvalidOperationException(
                            "A multi-value PivotTable exposed a hidden Values pseudo-field.");
                    }

                    return dataPivotField;
                case 1:
                    valuesAxis = PivotValuesAxis.Rows;
                    break;
                case 2:
                    valuesAxis = PivotValuesAxis.Columns;
                    break;
                default:
                    throw new InvalidOperationException(
                        "Excel exposed the Values pseudo-field on an unsupported axis.");
            }

            object positionValue = ReadRequiredObject(
                () => (object?)field.Position,
                "Excel did not expose the Values pseudo-field position.");
            try
            {
                valuesPosition = Convert.ToInt32(positionValue, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "Excel exposed an invalid Values pseudo-field position.",
                    exception);
            }

            if (valuesPosition < 1)
            {
                throw new InvalidOperationException(
                    "Excel exposed an invalid Values pseudo-field position.");
            }

            return dataPivotField;
        }

        private static IReadOnlyList<DiscoveredPivotField> ReadSourceFields(
            dynamic pivot,
            PivotConnectionKind connectionKind,
            DiscoveredPivotLayout layout,
            out bool complete)
        {
            bool cubeFields = connectionKind != PivotConnectionKind.Classic;
            object? collection = cubeFields
                ? ReadCollectionMember(
                    () => (object?)pivot.CubeFields,
                    () => (object?)pivot.CubeFields())
                : ReadCollectionMember(
                    () => (object?)pivot.PivotFields,
                    () => (object?)pivot.PivotFields());

            if (collection != null)
            {
                complete = true;
                object? dataPivotField = null;
                PivotLateBound.TryRead(
                    () => (object?)pivot.DataPivotField,
                    out dataPivotField);
                return ReadFields(
                    collection,
                    (field, _) => ReadField(field, 0, readNestedCubeField: !cubeFields),
                    dataPivotField);
            }

            complete = false;
            return layout.AllFields
                .GroupBy(field => field.StableName, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
        }

        private static IReadOnlyList<DiscoveredPivotField> CompleteInventory(
            IReadOnlyList<DiscoveredPivotField> sourceFields,
            DiscoveredPivotLayout layout)
        {
            return sourceFields
                .Concat(layout.AllFields)
                .GroupBy(field => field.StableName, StringComparer.OrdinalIgnoreCase)
                .Select(group => group
                    .OrderByDescending(field => FieldKindPriority(field.Kind))
                    .First())
                .ToArray();
        }

        private static int FieldKindPriority(DiscoveredPivotFieldKind kind)
        {
            switch (kind)
            {
                case DiscoveredPivotFieldKind.Measure: return 5;
                case DiscoveredPivotFieldKind.Set: return 4;
                case DiscoveredPivotFieldKind.Hierarchy: return 3;
                case DiscoveredPivotFieldKind.Regular: return 2;
                default: return 1;
            }
        }

        private static IReadOnlyList<PivotFieldDescriptor> MapFields(
            IReadOnlyList<DiscoveredPivotField> sourceFields,
            PivotConnectionKind connectionKind)
        {
            return sourceFields
                .Select(field =>
                {
                    bool isMeasure = IsMeasure(field);
                    return new PivotFieldDescriptor(
                        field.StableName,
                        NullIfEmpty(field.Caption),
                        PivotFieldDataType.Unknown,
                        SupportedAreas(field.Kind, isMeasure),
                        tableName: connectionKind == PivotConnectionKind.DataModel
                            ? ReadModelTableName(field.StableName)
                            : null,
                        isMeasure: isMeasure,
                        isCalculated: field.IsCalculated ||
                                      field.Kind == DiscoveredPivotFieldKind.Set);
                })
                .ToArray();
        }

        private static PivotFieldAreaSupport SupportedAreas(
            DiscoveredPivotFieldKind kind,
            bool isMeasure)
        {
            if (isMeasure)
            {
                return PivotFieldAreaSupport.Values;
            }

            switch (kind)
            {
                case DiscoveredPivotFieldKind.Hierarchy:
                    return PivotFieldAreaSupport.Row |
                           PivotFieldAreaSupport.Column |
                           PivotFieldAreaSupport.Filter;
                case DiscoveredPivotFieldKind.Set:
                    return PivotFieldAreaSupport.Row | PivotFieldAreaSupport.Column;
                default:
                    return PivotFieldAreaSupport.All;
            }
        }

        private static string? ReadModelTableName(string fieldName)
        {
            if (fieldName.Length < 4 || fieldName[0] != '[')
            {
                return null;
            }

            int closingBracket = fieldName.IndexOf(']');
            if (closingBracket <= 1 ||
                closingBracket + 2 >= fieldName.Length ||
                fieldName[closingBracket + 1] != '.' ||
                fieldName[closingBracket + 2] != '[')
            {
                return null;
            }

            string tableName = fieldName.Substring(1, closingBracket - 1)
                .Replace("]]", "]");
            return string.Equals(tableName, "Measures", StringComparison.OrdinalIgnoreCase)
                ? null
                : tableName;
        }

        private static IReadOnlyList<PivotFieldPlacement> MapPlacements(
            DiscoveredPivotLayout layout,
            IReadOnlyList<PivotFieldDescriptor> fields)
        {
            var descriptors = fields.ToDictionary(
                field => field.Name,
                StringComparer.OrdinalIgnoreCase);
            var placements = new List<PivotFieldPlacement>();
            AddPlacements(placements, descriptors, layout.Rows, PivotFieldArea.Row);
            AddPlacements(placements, descriptors, layout.Columns, PivotFieldArea.Column);
            AddPlacements(placements, descriptors, layout.Values, PivotFieldArea.Values);
            AddPlacements(placements, descriptors, layout.Filters, PivotFieldArea.Filter);
            return placements;
        }

        private static void AddPlacements(
            ICollection<PivotFieldPlacement> result,
            IReadOnlyDictionary<string, PivotFieldDescriptor> descriptors,
            IEnumerable<DiscoveredPivotField> fields,
            PivotFieldArea area)
        {
            foreach (DiscoveredPivotField field in fields)
            {
                descriptors.TryGetValue(field.StableName, out PivotFieldDescriptor? descriptor);
                bool isMeasure = descriptor?.IsMeasure == true || IsMeasure(field);
                result.Add(new PivotFieldPlacement(
                    field.StableName,
                    area,
                    field.Position,
                    caption: NullIfEmpty(field.Caption),
                    aggregation: area == PivotFieldArea.Values && !isMeasure
                        ? field.Aggregation
                        : null,
                    numberFormatCode: area == PivotFieldArea.Values
                        ? field.NumberFormatCode
                        : null));
            }
        }

        private static bool IsMeasure(DiscoveredPivotField field)
        {
            return field.Kind == DiscoveredPivotFieldKind.Measure ||
                   field.StableName.StartsWith("[Measures].", StringComparison.OrdinalIgnoreCase);
        }

        private static PivotLayoutMetadata ReadLayoutMetadata(
            dynamic pivot,
            DiscoveredPivotLayout layout)
        {
            PivotLayoutForm form = PivotLayoutForm.Compact;
            if (PivotLateBound.TryRead(() => (object?)pivot.LayoutRowDefault, out object? formValue))
            {
                try
                {
                    switch (Convert.ToInt32(formValue, CultureInfo.InvariantCulture))
                    {
                        case 1:
                            form = PivotLayoutForm.Tabular;
                            break;
                        case 2:
                            form = PivotLayoutForm.Outline;
                            break;
                    }
                }
                catch (Exception exception) when (
                    exception is FormatException ||
                    exception is InvalidCastException ||
                    exception is OverflowException)
                {
                    form = PivotLayoutForm.Compact;
                }
            }

            return new PivotLayoutMetadata(
                form,
                repeatItemLabels: layout.Rows.Count > 0 &&
                                  layout.Rows.All(field => field.RepeatLabels == true),
                showRowGrandTotals: ReadOptionalBoolean(
                    () => (object?)pivot.RowGrand,
                    fallback: true),
                showColumnGrandTotals: ReadOptionalBoolean(
                    () => (object?)pivot.ColumnGrand,
                    fallback: true),
                showFieldHeaders: ReadOptionalBoolean(
                    () => (object?)pivot.DisplayFieldCaptions,
                    fallback: true),
                valuesAxis: layout.ValuesAxis,
                valuesPosition: layout.ValuesPosition);
        }

        private static PivotFormatMetadata ReadFormatMetadata(dynamic pivot)
        {
            return new PivotFormatMetadata(
                pivotTableStyleName: ReadPivotTableStyleName(pivot),
                preserveFormatting: ReadOptionalBoolean(
                    () => (object?)pivot.PreserveFormatting,
                    fallback: true),
                showRowStripes: ReadOptionalBoolean(
                    () => (object?)pivot.ShowTableStyleRowStripes,
                    fallback: false),
                showColumnStripes: ReadOptionalBoolean(
                    () => (object?)pivot.ShowTableStyleColumnStripes,
                    fallback: false));
        }

        private static IReadOnlyList<DiscoveredPivotField> ReadFields(
            object collectionObject,
            Func<object, int, DiscoveredPivotField> reader,
            object? excludedField = null)
        {
            dynamic collection = collectionObject;
            object countValue = ReadRequiredObject(
                () => (object?)collection.Count,
                "Excel did not expose a PivotTable field collection count.");
            int count;
            try
            {
                count = Convert.ToInt32(countValue, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                throw new InvalidOperationException(
                    "Excel returned an invalid PivotTable field collection count.",
                    exception);
            }

            if (count < 0)
            {
                throw new InvalidOperationException(
                    "Excel returned an invalid PivotTable field collection count.");
            }

            var result = new List<DiscoveredPivotField>(count);
            var logicalPosition = 0;
            for (var index = 1; index <= count; index++)
            {
                object field = ReadRequiredObject(
                    () => ReadCollectionItem(collection, index),
                    "Excel did not expose a PivotTable field at index " +
                    index.ToString(CultureInfo.InvariantCulture) + ".");
                if (IsExcludedPseudoField(field, excludedField))
                {
                    continue;
                }

                logicalPosition++;
                result.Add(reader(field, logicalPosition));
            }

            return result;
        }

        private static bool IsExcludedPseudoField(object field, object? excludedField)
        {
            if (ComObjectIdentity.AreSame(field, excludedField)) return true;
            if (excludedField == null) return false;

            dynamic candidate = field;
            dynamic excluded = excludedField;
            string candidateName = ReadOptionalString(() => (object?)candidate.Name);
            string excludedName = ReadOptionalString(() => (object?)excluded.Name);
            if (string.IsNullOrWhiteSpace(candidateName) ||
                string.IsNullOrWhiteSpace(excludedName))
            {
                return false;
            }

            return string.Equals(candidateName, excludedName, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(candidateName, excludedName + "Field", StringComparison.OrdinalIgnoreCase);
        }

        private static object? ReadCollectionItem(dynamic collection, int index)
        {
            if (PivotLateBound.TryRead(
                    () => (object?)collection.Item(index),
                    out object? value))
            {
                return value;
            }

            if (PivotLateBound.TryRead(
                    () => (object?)collection[index],
                    out value))
            {
                return value;
            }

            return null;
        }

        private static bool TryReadCollectionCount(object collectionObject, out int count)
        {
            dynamic collection = collectionObject;
            count = 0;
            if (!PivotLateBound.TryRead(() => (object?)collection.Count, out object? countValue))
            {
                return false;
            }

            try
            {
                count = Convert.ToInt32(countValue, CultureInfo.InvariantCulture);
                return count >= 0;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                count = 0;
                return false;
            }
        }

        private static DiscoveredPivotField ReadField(
            object fieldObject,
            int position,
            bool readNestedCubeField,
            bool useCollectionPosition = false)
        {
            dynamic field = fieldObject;
            string name = ReadOptionalString(() => (object?)field.Name);
            string caption = ReadOptionalString(() => (object?)field.Caption);
            string sourceName = ReadOptionalString(() => (object?)field.SourceName);
            if (string.IsNullOrWhiteSpace(name))
            {
                name = !string.IsNullOrWhiteSpace(sourceName) ? sourceName : caption;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("Excel exposed an unnamed PivotTable field.");
            }

            int actualPosition = position > 0
                ? (useCollectionPosition ? position : ReadPosition(field, position))
                : 0;
            DiscoveredPivotFieldKind kind = readNestedCubeField
                ? ReadNestedCubeFieldKind(field) ?? DiscoveredPivotFieldKind.Regular
                : ReadCubeFieldKind(field);
            if (kind == DiscoveredPivotFieldKind.Regular &&
                sourceName.StartsWith("[Measures].", StringComparison.OrdinalIgnoreCase))
            {
                kind = DiscoveredPivotFieldKind.Measure;
            }

            return new DiscoveredPivotField(
                name,
                string.IsNullOrWhiteSpace(caption) ? name : caption,
                string.IsNullOrWhiteSpace(sourceName) ? name : sourceName,
                kind,
                actualPosition,
                ReadAggregation(field),
                NullIfEmpty(ReadOptionalString(() => (object?)field.NumberFormat)),
                ReadOptionalBoolean(() => (object?)field.IsCalculated, fallback: false),
                ReadOptionalNullableBoolean(() => (object?)field.RepeatLabels));
        }


        private static DiscoveredPivotFieldKind ReadCubeFieldKind(dynamic field)
        {
            if (!PivotLateBound.TryRead(
                    () => (object?)field.CubeFieldType,
                    out object? cubeFieldType))
            {
                return DiscoveredPivotFieldKind.Unknown;
            }

            int value;
            try
            {
                value = Convert.ToInt32(cubeFieldType, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return DiscoveredPivotFieldKind.Unknown;
            }

            switch (value)
            {
                case 1: return DiscoveredPivotFieldKind.Hierarchy;
                case 2: return DiscoveredPivotFieldKind.Measure;
                case 3: return DiscoveredPivotFieldKind.Set;
                default: return DiscoveredPivotFieldKind.Unknown;
            }
        }

        private static DiscoveredPivotFieldKind? ReadNestedCubeFieldKind(dynamic field)
        {
            if (!PivotLateBound.TryRead(() => (object?)field.CubeField, out object? cubeField) ||
                cubeField == null)
            {
                return null;
            }

            return ReadCubeFieldKind((dynamic)cubeField);
        }

        private static PivotAggregationFunction? ReadAggregation(dynamic field)
        {
            if (!PivotLateBound.TryRead(() => (object?)field.Function, out object? functionValue))
            {
                return null;
            }

            int value;
            try
            {
                value = Convert.ToInt32(functionValue, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return null;
            }

            switch (value)
            {
                case -4157: return PivotAggregationFunction.Sum;
                case -4112: return PivotAggregationFunction.Count;
                case -4106: return PivotAggregationFunction.Average;
                case -4139: return PivotAggregationFunction.Minimum;
                case -4136: return PivotAggregationFunction.Maximum;
                case -4149: return PivotAggregationFunction.Product;
                case -4113: return PivotAggregationFunction.CountNumbers;
                case -4155: return PivotAggregationFunction.StandardDeviation;
                case -4156: return PivotAggregationFunction.StandardDeviationPopulation;
                case -4164: return PivotAggregationFunction.Variance;
                case -4165: return PivotAggregationFunction.VariancePopulation;
                case 11: return PivotAggregationFunction.DistinctCount;
                default: return null;
            }
        }

        private static int ReadPosition(dynamic field, int fallback)
        {
            if (!PivotLateBound.TryRead(() => (object?)field.Position, out object? positionValue))
            {
                return fallback;
            }

            try
            {
                int position = Convert.ToInt32(positionValue, CultureInfo.InvariantCulture);
                return position > 0 ? position : fallback;
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException ||
                exception is OverflowException)
            {
                return fallback;
            }
        }

        private static object? ReadCollectionMember(
            Func<object?> propertyReader,
            Func<object?> methodReader)
        {
            if (PivotLateBound.TryRead(propertyReader, out object? collection) &&
                collection != null)
            {
                return collection;
            }

            return PivotLateBound.TryRead(methodReader, out collection)
                ? collection
                : null;
        }

        private static object ReadRequiredObject(
            Func<object?> reader,
            string errorMessage)
        {
            if (!PivotLateBound.TryRead(reader, out object? value) || value == null)
            {
                throw new InvalidOperationException(errorMessage);
            }

            return value;
        }

        private static string ReadRequiredName(Func<object?> reader, string errorMessage)
        {
            string value = ReadOptionalString(reader);
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException(errorMessage);
            }

            return value;
        }

        private static string ReadOptionalString(Func<object?> reader)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string? ReadPivotTableStyleName(dynamic pivot)
        {
            if (!PivotLateBound.TryRead(() => (object?)pivot.TableStyle2, out object? value) ||
                value == null)
            {
                return null;
            }

            if (value is string text)
            {
                return NullIfEmpty(text);
            }

            dynamic style = value;
            if (PivotLateBound.TryRead(() => (object?)style.Name, out object? name) &&
                name != null)
            {
                string resolved = Convert.ToString(name, CultureInfo.InvariantCulture) ?? string.Empty;
                if (!string.Equals(resolved, "System.__ComObject", StringComparison.Ordinal))
                {
                    return NullIfEmpty(resolved);
                }
            }

            string fallback = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return string.Equals(fallback, "System.__ComObject", StringComparison.Ordinal)
                ? null
                : NullIfEmpty(fallback);
        }

        private static bool ReadOptionalBoolean(Func<object?> reader, bool fallback)
        {
            if (!PivotLateBound.TryRead(reader, out object? value))
            {
                return fallback;
            }

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

        private static bool? ReadOptionalNullableBoolean(Func<object?> reader)
        {
            if (!PivotLateBound.TryRead(reader, out object? value) || value == null)
            {
                return null;
            }

            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception exception) when (
                exception is FormatException ||
                exception is InvalidCastException)
            {
                return null;
            }
        }

        private static string? NullIfEmpty(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }
    }

    internal static class PivotLateBound
    {
        public static bool TryRead(Func<object?> reader, out object? value)
        {
            try
            {
                value = reader();
                return true;
            }
            catch (Exception exception) when (IsDispatchFailure(exception))
            {
                value = null;
                return false;
            }
        }

        private static bool IsDispatchFailure(Exception exception)
        {
            return exception is COMException ||
                   exception is RuntimeBinderException ||
                   exception is MissingMemberException ||
                   exception is TargetInvocationException ||
                   exception is ArgumentException ||
                   exception is InvalidOperationException ||
                   exception is NotSupportedException;
        }
    }
}
