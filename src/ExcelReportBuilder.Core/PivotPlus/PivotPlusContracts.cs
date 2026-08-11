using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelReportBuilder.Core.PivotPlus
{
    public enum PivotSourceKind
    {
        Unknown,
        WorksheetRange,
        WorksheetTable,
        DataModel,
        ExternalOlap
    }

    /// <summary>
    /// Capabilities exposed by the selected Excel source. A capability signals
    /// source potential; callers must also require an implementation service
    /// for the specific operation they intend to execute.
    /// </summary>
    [Flags]
    public enum PivotCapability
    {
        None = 0,
        NativeFieldPlacement = 1 << 0,
        MemberFiltering = 1 << 1,
        LayoutFormatting = 1 << 2,
        ShowValuesAs = 1 << 3,
        DistinctCount = 1 << 4,
        DataModel = 1 << 5,
        ModelMeasures = 1 << 6,
        CalculatedMembers = 1 << 7,
        NamedSets = 1 << 8,
        AsymmetricAxes = 1 << 9,
        Refresh = 1 << 10,
        UpgradeToDataModel = 1 << 11
    }

    public enum PivotFieldArea
    {
        Row,
        Column,
        Filter,
        Values
    }

    [Flags]
    public enum PivotFieldAreaSupport
    {
        None = 0,
        Row = 1 << 0,
        Column = 1 << 1,
        Filter = 1 << 2,
        Values = 1 << 3,
        All = Row | Column | Filter | Values
    }

    public enum PivotFieldDataType
    {
        Unknown,
        Text,
        Number,
        Date,
        Boolean
    }

    public enum PivotAggregationFunction
    {
        Sum,
        Count,
        Average,
        Minimum,
        Maximum,
        Product,
        CountNumbers,
        StandardDeviation,
        StandardDeviationPopulation,
        Variance,
        VariancePopulation,
        DistinctCount
    }

    public enum PivotFilterMode
    {
        All,
        Include,
        Exclude
    }

    public enum PivotLayoutForm
    {
        Compact,
        Outline,
        Tabular
    }

    public enum PivotValuesAxis
    {
        Automatic,
        Rows,
        Columns
    }

    public enum PivotSubtotalMode
    {
        None,
        Automatic
    }

    /// <summary>
    /// Identifies a PivotTable without persisting a workbook name or filesystem
    /// path. WorkbookId is a stable host-issued token stored in the workbook.
    /// </summary>
    public sealed class PivotTargetIdentity
    {
        public PivotTargetIdentity(string workbookId, string worksheetName, string pivotTableName)
        {
            WorkbookId = workbookId ?? string.Empty;
            WorksheetName = worksheetName ?? string.Empty;
            PivotTableName = pivotTableName ?? string.Empty;
        }

        public string WorkbookId { get; }

        public string WorksheetName { get; }

        public string PivotTableName { get; }
    }

    /// <summary>
    /// Describes the native source visible to Excel. SourceName is a workbook
    /// object or connection name, never a workbook path.
    /// </summary>
    public sealed class PivotSourceDescriptor
    {
        public PivotSourceDescriptor(
            PivotSourceKind kind,
            string sourceName,
            PivotCapability capabilities,
            string? modelTableName = null)
        {
            Kind = kind;
            SourceName = sourceName ?? string.Empty;
            ModelTableName = modelTableName;
            Capabilities = capabilities;
        }

        public PivotSourceKind Kind { get; }

        public string SourceName { get; }

        public string? ModelTableName { get; }

        public PivotCapability Capabilities { get; }
    }

    public sealed class PivotFieldDescriptor
    {
        public PivotFieldDescriptor(
            string name,
            string? caption,
            PivotFieldDataType dataType,
            PivotFieldAreaSupport supportedAreas,
            string? tableName = null,
            bool isMeasure = false,
            bool isCalculated = false)
        {
            Name = name ?? string.Empty;
            Caption = caption;
            DataType = dataType;
            SupportedAreas = supportedAreas;
            TableName = tableName;
            IsMeasure = isMeasure;
            IsCalculated = isCalculated;
        }

        /// <summary>
        /// The stable field identifier used by Excel. For model sources this
        /// should be the unique cube-field name rather than a display caption.
        /// </summary>
        public string Name { get; }

        public string? Caption { get; }

        public string? TableName { get; }

        public PivotFieldDataType DataType { get; }

        public PivotFieldAreaSupport SupportedAreas { get; }

        public bool IsMeasure { get; }

        public bool IsCalculated { get; }
    }

    public sealed class PivotFieldPlacement
    {
        public PivotFieldPlacement(
            string fieldName,
            PivotFieldArea area,
            int position,
            string? caption = null,
            PivotAggregationFunction? aggregation = null,
            string? numberFormatCode = null,
            PivotSubtotalMode subtotalMode = PivotSubtotalMode.None)
        {
            FieldName = fieldName ?? string.Empty;
            Area = area;
            Position = position;
            Caption = caption;
            Aggregation = aggregation;
            NumberFormatCode = numberFormatCode;
            SubtotalMode = subtotalMode;
        }

        public string FieldName { get; }

        public PivotFieldArea Area { get; }

        /// <summary>
        /// One-based position within the selected field area.
        /// </summary>
        public int Position { get; }

        public string? Caption { get; }

        public PivotAggregationFunction? Aggregation { get; }

        public string? NumberFormatCode { get; }

        public PivotSubtotalMode SubtotalMode { get; }
    }

    public sealed class PivotFieldFilter
    {
        public PivotFieldFilter(
            string fieldName,
            PivotFilterMode mode,
            IEnumerable<string>? members = null,
            bool includeBlank = false)
        {
            FieldName = fieldName ?? string.Empty;
            Mode = mode;
            Members = Copy(members);
            IncludeBlank = includeBlank;
        }

        public string FieldName { get; }

        public PivotFilterMode Mode { get; }

        public IReadOnlyList<string> Members { get; }

        public bool IncludeBlank { get; }

        private static IReadOnlyList<string> Copy(IEnumerable<string>? values)
        {
            return new ReadOnlyCollection<string>((values ?? Enumerable.Empty<string>()).ToList());
        }
    }

    public sealed class PivotLayoutMetadata
    {
        public PivotLayoutMetadata(
            PivotLayoutForm form = PivotLayoutForm.Compact,
            bool repeatItemLabels = false,
            bool showRowGrandTotals = true,
            bool showColumnGrandTotals = true,
            bool showFieldHeaders = true,
            PivotValuesAxis valuesAxis = PivotValuesAxis.Automatic,
            int valuesPosition = 1)
        {
            Form = form;
            RepeatItemLabels = repeatItemLabels;
            ShowRowGrandTotals = showRowGrandTotals;
            ShowColumnGrandTotals = showColumnGrandTotals;
            ShowFieldHeaders = showFieldHeaders;
            ValuesAxis = valuesAxis;
            ValuesPosition = valuesPosition;
        }

        public PivotLayoutForm Form { get; }

        public bool RepeatItemLabels { get; }

        public bool ShowRowGrandTotals { get; }

        public bool ShowColumnGrandTotals { get; }

        public bool ShowFieldHeaders { get; }

        public PivotValuesAxis ValuesAxis { get; }

        /// <summary>
        /// One-based position of Excel's synthetic Values field on the chosen
        /// row or column axis. Regular placement positions exclude that
        /// synthetic field.
        /// </summary>
        public int ValuesPosition { get; }
    }

    public sealed class PivotFormatMetadata
    {
        public PivotFormatMetadata(
            string? pivotTableStyleName = null,
            bool preserveFormatting = true,
            bool showRowStripes = false,
            bool showColumnStripes = false)
        {
            PivotTableStyleName = pivotTableStyleName;
            PreserveFormatting = preserveFormatting;
            ShowRowStripes = showRowStripes;
            ShowColumnStripes = showColumnStripes;
        }

        public string? PivotTableStyleName { get; }

        public bool PreserveFormatting { get; }

        public bool ShowRowStripes { get; }

        public bool ShowColumnStripes { get; }
    }

    public sealed class PivotCapabilityRequirement
    {
        public PivotCapabilityRequirement(PivotCapability capability, string reason)
        {
            Capability = capability;
            Reason = reason ?? string.Empty;
        }

        public PivotCapability Capability { get; }

        public string Reason { get; }
    }

    /// <summary>
    /// A complete discovery/layout snapshot for one real Excel PivotTable.
    /// Collections are defensively copied so later UI edits cannot mutate a
    /// validated definition behind the executor's back.
    /// </summary>
    public sealed class PivotLayoutDefinition
    {
        public PivotLayoutDefinition(
            PivotTargetIdentity target,
            PivotSourceDescriptor source,
            IEnumerable<PivotFieldDescriptor>? fields,
            IEnumerable<PivotFieldPlacement>? placements,
            IEnumerable<PivotFieldFilter>? filters = null,
            PivotLayoutMetadata? layout = null,
            PivotFormatMetadata? format = null,
            IEnumerable<PivotCapabilityRequirement>? capabilityRequirements = null,
            bool clearAll = false)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Fields = Copy(fields);
            Placements = Copy(placements);
            Filters = Copy(filters);
            Layout = layout ?? new PivotLayoutMetadata();
            Format = format ?? new PivotFormatMetadata();
            CapabilityRequirements = Copy(capabilityRequirements);
            ClearAll = clearAll;
        }

        public PivotTargetIdentity Target { get; }

        public PivotSourceDescriptor Source { get; }

        public IReadOnlyList<PivotFieldDescriptor> Fields { get; }

        public IReadOnlyList<PivotFieldPlacement> Placements { get; }

        public IReadOnlyList<PivotFieldFilter> Filters { get; }

        public PivotLayoutMetadata Layout { get; }

        public PivotFormatMetadata Format { get; }

        public IReadOnlyList<PivotCapabilityRequirement> CapabilityRequirements { get; }

        /// <summary>
        /// Explicitly authorizes removal of every native PivotTable placement.
        /// An empty placements collection is otherwise invalid so omitted or
        /// truncated input cannot silently clear a report.
        /// </summary>
        public bool ClearAll { get; }

        private static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>((values ?? Enumerable.Empty<T>()).ToList());
        }
    }
}
