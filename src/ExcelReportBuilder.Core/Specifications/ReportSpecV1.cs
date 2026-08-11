using System;
using System.Collections.Generic;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Specifications
{
    public enum WorkbookSourceKind
    {
        Table,
        NamedRange
    }

    public enum SubtotalMode
    {
        None,
        Automatic
    }

    public enum TotalPlacement
    {
        BeforeMembers,
        AfterMembers
    }

    public enum SortDirection
    {
        SourceOrder,
        Ascending,
        Descending
    }

    public enum TopNDirection
    {
        Top,
        Bottom
    }

    public enum PeriodMappingKind
    {
        LongDateColumn,
        MonthHeaders,
        MetricMonthHeaders
    }

    public enum PeriodGrain
    {
        Day,
        Month,
        Quarter
    }

    public enum PeriodSliceKind
    {
        Current,
        Prior,
        Selected,
        SamePeriodPriorYear
    }

    public enum SpacerAxis
    {
        Row,
        Column
    }

    public enum ReportOutputMode
    {
        StandardMatrix,
        MetricStack,
        DenseGrid
    }

    public enum HorizontalAlignment
    {
        General,
        Left,
        Center,
        Right
    }

    public enum ReportCheckKind
    {
        TotalPreservation,
        NoTruncation,
        RequiredValues,
        NonNegative,
        Balance
    }

    /// <summary>
    /// Version 1 of the single bounded contract edited by both the manual builder
    /// and the AI-assisted builder.
    /// </summary>
    public sealed class ReportSpecV1
    {
        public const string CurrentSchemaVersion = "1.0";

        public string SchemaVersion { get; set; } = CurrentSchemaVersion;

        public string Id { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public WorkbookSourceSpec Source { get; set; } = new WorkbookSourceSpec();

        public string OwnershipId { get; set; } = string.Empty;

        public PeriodMappingSpec? PeriodMapping { get; set; }

        public List<TransformStep> Transforms { get; set; } = new List<TransformStep>();

        public List<MeasureDefinition> Measures { get; set; } = new List<MeasureDefinition>();

        public List<ReportBlockSpec> Blocks { get; set; } = new List<ReportBlockSpec>();

        public List<PresentationStyleSpec> Styles { get; set; } = new List<PresentationStyleSpec>();

        public List<ReportCheckSpec> Checks { get; set; } = new List<ReportCheckSpec>();
    }

    public sealed class PeriodMappingSpec
    {
        public string Id { get; set; } = "periods";

        public PeriodMappingKind Kind { get; set; }

        /// <summary>
        /// The semantic grain represented by each canonical Period value.
        /// Older month-header specifications may omit this member; month is
        /// then the deterministic effective grain for wide mappings.
        /// </summary>
        public PeriodGrain? Grain { get; set; }

        public string? DateColumn { get; set; }

        public List<string> KeyColumns { get; set; } = new List<string>();

        public List<PeriodColumnMapping> Columns { get; set; } = new List<PeriodColumnMapping>();

        public int? ReportingYear { get; set; }

        public string PeriodColumnName { get; set; } = "Period";

        public string ValueColumnName { get; set; } = "Value";

        public string MetricColumnName { get; set; } = "Metric";
    }

    public sealed class ReportBlockSpec
    {
        public string Id { get; set; } = string.Empty;

        public string OwnershipId { get; set; } = string.Empty;

        public string? Title { get; set; }

        public string WorksheetName { get; set; } = "Report";

        public string AnchorCell { get; set; } = "A1";

        public ReportOutputMode OutputMode { get; set; } = ReportOutputMode.DenseGrid;

        /// <summary>
        /// The maximum rectangular area this managed block may write, measured
        /// from AnchorCell. Executors must fail before writing outside it.
        /// </summary>
        public OwnedRangeExtentSpec OwnedExtent { get; set; } = new OwnedRangeExtentSpec();

        public ReportLayoutSpec Layout { get; set; } = new ReportLayoutSpec();

        public List<PeriodSliceSpec> PeriodSlices { get; set; } = new List<PeriodSliceSpec>();

        public List<ReportHeaderSpec> Headers { get; set; } = new List<ReportHeaderSpec>();

        public List<SpacerSpec> Spacers { get; set; } = new List<SpacerSpec>();

        public string? HeaderStyleId { get; set; }

        public string? BodyStyleId { get; set; }

        public string? SubtotalStyleId { get; set; }

        public string? GrandTotalStyleId { get; set; }
    }

    public sealed class OwnedRangeExtentSpec
    {
        public int RowCount { get; set; } = 1000;

        public int ColumnCount { get; set; } = 10;
    }

    public sealed class WorkbookSourceSpec
    {
        public WorkbookSourceKind Kind { get; set; } = WorkbookSourceKind.Table;

        /// <summary>
        /// The name exposed by Excel.CurrentWorkbook. Sheet paths and external
        /// files are intentionally not part of the contract.
        /// </summary>
        public string WorkbookObjectName { get; set; } = string.Empty;

        public int HeaderRowCount { get; set; } = 1;

        /// <summary>
        /// A path-free, value-free identity for the selected source shape.
        /// Saved report setups can use its compatibility key to detect schema
        /// drift without retaining workbook paths or header text.
        /// </summary>
        public SourceFingerprintSpec Fingerprint { get; set; } = new SourceFingerprintSpec();
    }

    public sealed class ReportLayoutSpec
    {
        public List<FieldPlacementSpec> Rows { get; set; } = new List<FieldPlacementSpec>();

        public List<FieldPlacementSpec> Columns { get; set; } = new List<FieldPlacementSpec>();

        public List<ValuePlacementSpec> Values { get; set; } = new List<ValuePlacementSpec>();

        public List<FilterPlacementSpec> Filters { get; set; } = new List<FilterPlacementSpec>();

        public DenseLayoutOptions DenseLayout { get; set; } = new DenseLayoutOptions();

        public GrandTotalsSpec GrandTotals { get; set; } = new GrandTotalsSpec();
    }

    public sealed class FieldPlacementSpec
    {
        public string Field { get; set; } = string.Empty;

        public string? Caption { get; set; }

        public SubtotalSpec Subtotals { get; set; } = new SubtotalSpec();

        public SortDirection Sort { get; set; } = SortDirection.SourceOrder;

        public List<ScalarValue> MemberOrder { get; set; } = new List<ScalarValue>();

        public List<MemberGroupBucketSpec> GroupBuckets { get; set; } = new List<MemberGroupBucketSpec>();

        public TopNSpec? TopN { get; set; }
    }

    public sealed class ValuePlacementSpec
    {
        public string MeasureId { get; set; } = string.Empty;

        public string? Caption { get; set; }

        public string? NumberFormat { get; set; }

        public List<string> PeriodSliceIds { get; set; } = new List<string>();

        public string? StyleId { get; set; }
    }

    public sealed class SubtotalSpec
    {
        public SubtotalMode Mode { get; set; } = SubtotalMode.Automatic;

        public TotalPlacement Placement { get; set; } = TotalPlacement.AfterMembers;

        public string? Label { get; set; }

        public string? StyleId { get; set; }
    }

    public sealed class MemberGroupBucketSpec
    {
        public string Id { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;

        public List<ScalarValue> Members { get; set; } = new List<ScalarValue>();

        public bool IncludeUnmatched { get; set; }
    }

    public sealed class TopNSpec
    {
        public int Count { get; set; } = 10;

        public string MeasureId { get; set; } = string.Empty;

        public TopNDirection Direction { get; set; } = TopNDirection.Top;

        public bool IncludeOthers { get; set; }

        public string OthersLabel { get; set; } = "Others";
    }

    public sealed class GrandTotalsSpec
    {
        public bool ShowRows { get; set; } = true;

        public bool ShowColumns { get; set; } = true;

        public TotalPlacement RowPlacement { get; set; } = TotalPlacement.AfterMembers;

        public TotalPlacement ColumnPlacement { get; set; } = TotalPlacement.AfterMembers;

        public string RowLabel { get; set; } = "Grand Total";

        public string ColumnLabel { get; set; } = "Grand Total";

        public string? StyleId { get; set; }
    }

    public sealed class FilterPlacementSpec
    {
        public string Field { get; set; } = string.Empty;

        public List<ScalarValue> SelectedValues { get; set; } = new List<ScalarValue>();

        public bool IncludeBlank { get; set; }
    }

    public sealed class DenseLayoutOptions
    {
        public bool RepeatRowLabels { get; set; }

        public bool ShowRowGrandTotals { get; set; } = true;

        public bool ShowColumnGrandTotals { get; set; } = true;

        public bool InsertBlankRows { get; set; }

        public int RowIndent { get; set; } = 1;

        public bool FreezeHeaders { get; set; } = true;
    }

    public sealed class ReportCheckSpec
    {
        public string Id { get; set; } = string.Empty;

        public ReportCheckKind Kind { get; set; }

        public string? MeasureId { get; set; }

        public string? ComparedMeasureId { get; set; }

        public decimal Tolerance { get; set; }
    }

    public sealed class PeriodSliceSpec
    {
        public string Id { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;

        public PeriodSliceKind Kind { get; set; }

        public DateTime? SelectedStart { get; set; }

        public DateTime? SelectedEnd { get; set; }

        public string? BasedOnSliceId { get; set; }
    }

    public sealed class ReportHeaderSpec
    {
        public string Text { get; set; } = string.Empty;

        public int RelativeRow { get; set; }

        public int RelativeColumn { get; set; }

        public int ColumnSpan { get; set; } = 1;

        public string? StyleId { get; set; }
    }

    public sealed class SpacerSpec
    {
        public SpacerAxis Axis { get; set; }

        public int BeforeLevel { get; set; }

        public int Count { get; set; } = 1;

        public double? Size { get; set; }
    }

    public sealed class PresentationStyleSpec
    {
        public string Id { get; set; } = string.Empty;

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public string? FontColor { get; set; }

        public string? FillColor { get; set; }

        public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.General;

        public string? NumberFormat { get; set; }

        public int? DecimalPlaces { get; set; }

        public bool TopBorder { get; set; }

        public bool BottomBorder { get; set; }
    }
}
