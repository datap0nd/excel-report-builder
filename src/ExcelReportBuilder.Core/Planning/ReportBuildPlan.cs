using System.Collections.Generic;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Planning
{
    public enum BuildOperationKind
    {
        ProfileSource,
        PrepareSource,
        LoadWorksheet,
        LoadDataModel,
        CreatePivot,
        RenderDenseBlock,
        RenderStandardMatrix,
        RenderMetricStack,
        RenderDenseGrid,
        RunChecks
    }

    public enum DenseBlockRegionKind
    {
        Headers,
        RowHierarchy,
        ColumnHierarchy,
        Values,
        Subtotals,
        GrandTotals,
        Checks
    }

    public enum AggregateComponentRole
    {
        Value,
        Input,
        WeightedNumerator,
        WeightedDenominator
    }

    public enum PivotMemberStageKind
    {
        ApplyMemberOrder,
        GroupMembers,
        SortAscending,
        SortDescending,
        ApplyTopN,
        AggregateOthers
    }

    public enum CheckEvaluationScope
    {
        CanonicalData,
        RenderedOutput
    }

    public enum RowCountExpectation
    {
        ExactProjection,
        ExactPostTransformCount,
        /// <summary>
        /// Retained only so an older in-memory plan fails closed instead of
        /// silently changing meaning. Reconciliation requires the same exact
        /// independent post-transform count as ExactPostTransformCount.
        /// </summary>
        AtMostProjection
    }

    public sealed class ReportBuildPlan
    {
        public string SpecificationId { get; set; } = string.Empty;

        public string SchemaVersion { get; set; } = string.Empty;

        public string SpecificationHash { get; set; } = string.Empty;

        public string PlanHash { get; set; } = string.Empty;

        public string OwnershipId { get; set; } = string.Empty;

        public SourcePreparationPlan Source { get; set; } = new SourcePreparationPlan();

        public List<DenseReportBlockPlan> Blocks { get; set; } = new List<DenseReportBlockPlan>();

        /// <summary>
        /// The complete closed style catalog referenced by block and region
        /// plans. Executors do not need to inspect the source specification.
        /// </summary>
        public List<PresentationStyleSpec> Styles { get; set; } = new List<PresentationStyleSpec>();

        public List<BuildCheckPlan> Checks { get; set; } = new List<BuildCheckPlan>();

        public List<BuildOperation> Operations { get; set; } = new List<BuildOperation>();
    }

    public sealed class SourcePreparationPlan
    {
        public string WorkbookObjectName { get; set; } = string.Empty;

        public SourceFingerprintSpec Fingerprint { get; set; } = new SourceFingerprintSpec();

        public string SavedSetupCompatibilityKey { get; set; } = string.Empty;

        public string ManagedQueryName { get; set; } = string.Empty;

        public string PowerQueryM { get; set; } = string.Empty;

        public SourceLoadRoute Route { get; set; }

        public long SourceRows { get; set; }

        public long ProjectedRows { get; set; }

        public long ExpansionFactor { get; set; }

        public bool TruncationAllowed { get; set; }
    }

    public sealed class DenseReportBlockPlan
    {
        public string BlockId { get; set; } = string.Empty;

        public string OwnershipId { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;

        public string AnchorCell { get; set; } = string.Empty;

        public ReportOutputMode OutputMode { get; set; }

        public OwnedRangePlan OwnedRange { get; set; } = new OwnedRangePlan();

        public string? Title { get; set; }

        public PivotTablePlan Pivot { get; set; } = new PivotTablePlan();

        public DensePresentationPlan Presentation { get; set; } = new DensePresentationPlan();

        public List<DenseBlockRegionPlan> Regions { get; set; } = new List<DenseBlockRegionPlan>();
    }

    public sealed class OwnedRangePlan
    {
        public string AnchorCell { get; set; } = string.Empty;

        public int RowCount { get; set; }

        public int ColumnCount { get; set; }
    }

    public sealed class PivotTablePlan
    {
        public string ManagedPivotName { get; set; } = string.Empty;

        public string ManagedCacheName { get; set; } = string.Empty;

        public bool UseDataModel { get; set; }

        public List<PivotFieldPlan> Rows { get; set; } = new List<PivotFieldPlan>();

        public List<PivotFieldPlan> Columns { get; set; } = new List<PivotFieldPlan>();

        public List<PivotFilterPlan> Filters { get; set; } = new List<PivotFilterPlan>();

        public List<PivotValuePlan> Values { get; set; } = new List<PivotValuePlan>();

        /// <summary>
        /// Hidden aggregate inputs required for ordering, Top N, or Others
        /// even when the measure is not displayed as a Value.
        /// </summary>
        public List<PivotValuePlan> SupportingValues { get; set; } = new List<PivotValuePlan>();

        public GrandTotalsSpec GrandTotals { get; set; } = new GrandTotalsSpec();
    }

    public sealed class PivotFieldPlan
    {
        public string Field { get; set; } = string.Empty;

        public string? Caption { get; set; }

        public SortDirection Sort { get; set; }

        public SubtotalSpec Subtotals { get; set; } = new SubtotalSpec();

        public List<ScalarValue> MemberOrder { get; set; } = new List<ScalarValue>();

        public List<MemberGroupBucketSpec> GroupBuckets { get; set; } = new List<MemberGroupBucketSpec>();

        public TopNSpec? TopN { get; set; }

        /// <summary>
        /// Ordered member operations. AggregateOthers always represents the
        /// members excluded by the immediately preceding Top N stage.
        /// </summary>
        public List<PivotMemberStageKind> MemberStages { get; set; } = new List<PivotMemberStageKind>();
    }

    public sealed class PivotFilterPlan
    {
        public string Field { get; set; } = string.Empty;

        public List<ScalarValue> SelectedValues { get; set; } = new List<ScalarValue>();

        public bool IncludeBlank { get; set; }

        public bool IsSupportingField { get; set; }
    }

    public sealed class PivotValuePlan
    {
        public string MeasureId { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;

        public MeasureValueType ValueType { get; set; }

        public MeasureExpression Expression { get; set; } = new ConstantMeasureExpression();

        public string? NumberFormat { get; set; }

        public List<string> PeriodSliceIds { get; set; } = new List<string>();

        public bool RequiresPostAggregationCalculation { get; set; }

        public List<PivotAggregateComponentPlan> AggregateComponents { get; set; } = new List<PivotAggregateComponentPlan>();
    }

    public sealed class PivotAggregateComponentPlan
    {
        public string Id { get; set; } = string.Empty;

        public AggregateComponentRole Role { get; set; }

        public string Field { get; set; } = string.Empty;

        public AggregateFunction Function { get; set; }

        public MeasureValueType ValueType { get; set; }

        public string? PeriodSliceId { get; set; }

        public List<MeasureFilterSpec> Filters { get; set; } = new List<MeasureFilterSpec>();
    }

    public sealed class DensePresentationPlan
    {
        public DenseLayoutOptions Options { get; set; } = new DenseLayoutOptions();

        public List<PeriodSliceSpec> PeriodSlices { get; set; } = new List<PeriodSliceSpec>();

        public List<ResolvedPeriodSlice> ResolvedPeriodSlices { get; set; } = new List<ResolvedPeriodSlice>();

        public List<ReportHeaderSpec> Headers { get; set; } = new List<ReportHeaderSpec>();

        public List<SpacerSpec> Spacers { get; set; } = new List<SpacerSpec>();

        public string? HeaderStyleId { get; set; }

        public string? BodyStyleId { get; set; }

        public string? SubtotalStyleId { get; set; }

        public string? GrandTotalStyleId { get; set; }
    }

    public sealed class DenseBlockRegionPlan
    {
        public DenseBlockRegionKind Kind { get; set; }

        public int RelativeRow { get; set; }

        public int RelativeColumn { get; set; }

        public int? FixedRowCount { get; set; }

        public int? FixedColumnCount { get; set; }

        public bool DynamicRows { get; set; }

        public bool DynamicColumns { get; set; }

        public string? StyleId { get; set; }
    }

    public sealed class BuildCheckPlan
    {
        public string Id { get; set; } = string.Empty;

        public ReportCheckKind Kind { get; set; }

        public string? MeasureId { get; set; }

        public string? ComparedMeasureId { get; set; }

        public PivotValuePlan? Measure { get; set; }

        public PivotValuePlan? ComparedMeasure { get; set; }

        public decimal Tolerance { get; set; }

        public bool Mandatory { get; set; }

        public CheckEvaluationScope EvaluationScope { get; set; }

        public RowCountExpectation RowCountExpectation { get; set; } = RowCountExpectation.ExactProjection;
    }

    public sealed class BuildOperation
    {
        public int Sequence { get; set; }

        public BuildOperationKind Kind { get; set; }

        public string OwnershipId { get; set; } = string.Empty;

        public string Description { get; set; } = string.Empty;
    }
}
