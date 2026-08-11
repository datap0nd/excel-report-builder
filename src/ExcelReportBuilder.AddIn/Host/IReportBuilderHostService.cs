using System;
using System.Collections.Generic;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.AddIn.Activity;

namespace ExcelReportBuilder.AddIn.Host
{
    /// <summary>
    /// Bounded host boundary for Excel and worker operations. The UI can propose only
    /// these report-builder actions and never receives arbitrary COM or filesystem access.
    /// </summary>
    public interface IReportBuilderHostService
    {
        event EventHandler<HostActivityEventArgs>? ActivityReported;

        bool IsSynthetic { get; }

        SavedEndpointSettingsSnapshot? SavedEndpointSettings { get; }

        Task<SourceSnapshot> SelectCurrentDataAsync(CancellationToken cancellationToken);

        Task ConfirmPeriodMappingAsync(
            PeriodMappingSnapshot periodMapping,
            CancellationToken cancellationToken);

        Task<WideHeaderMappingPreview> PreviewWideHeaderMappingAsync(
            string headerPattern,
            int? reportingYear,
            CancellationToken cancellationToken);

        Task<BuildDraftResult> BuildManagedDraftAsync(
            ReportSpecificationSnapshot specification,
            CancellationToken cancellationToken);

        Task<ChatRunResult> RunChatAsync(
            string request,
            ReportSpecificationSnapshot specification,
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken);

        Task<IReadOnlyList<string>> DiscoverModelsAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken);

        Task<EndpointCheckResult> CheckEndpointAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken);

        Task PersistEndpointSettingsAsync(
            ModelEndpointSettingsSnapshot endpointSettings,
            SecureString? apiKey,
            CancellationToken cancellationToken);

        Task<IReadOnlyList<HostCheckResult>> RunChecksAsync(CancellationToken cancellationToken);

        Task<PublishResult> PublishManagedDraftAsync(CancellationToken cancellationToken);

        void RequestPause();

        void RequestResume();

        void RequestCancel();
    }

    public sealed class HostActivityEventArgs : EventArgs
    {
        public HostActivityEventArgs(
            ActivityStage stage,
            ActivityKind kind,
            string message,
            string detail)
        {
            Stage = stage;
            Kind = kind;
            Message = message ?? throw new ArgumentNullException(nameof(message));
            Detail = detail ?? string.Empty;
        }

        public ActivityStage Stage { get; }

        public ActivityKind Kind { get; }

        public string Message { get; }

        public string Detail { get; }
    }

    public sealed class SourceSnapshot
    {
        public SourceSnapshot(
            string displayName,
            string location,
            int rowCount,
            IReadOnlyList<SourceColumnSnapshot> columns,
            bool isSynthetic,
            ReportSpecificationSnapshot? savedReportSetup = null,
            string savedReportSetupStatus = "")
        {
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
            Location = location ?? throw new ArgumentNullException(nameof(location));
            RowCount = rowCount;
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
            IsSynthetic = isSynthetic;
            SavedReportSetup = savedReportSetup;
            SavedReportSetupStatus = savedReportSetupStatus ?? string.Empty;
        }

        public string DisplayName { get; }

        public string Location { get; }

        public int RowCount { get; }

        public IReadOnlyList<SourceColumnSnapshot> Columns { get; }

        public bool IsSynthetic { get; }

        /// <summary>
        /// A bounded manual-builder projection of a compatible saved setup.
        /// It is present only after the host has matched both the selected
        /// workbook object and its path-free source fingerprint.
        /// </summary>
        public ReportSpecificationSnapshot? SavedReportSetup { get; }

        public string SavedReportSetupStatus { get; }
    }

    public sealed class SourceColumnSnapshot
    {
        public SourceColumnSnapshot(string name, string dataType, string sampleValue)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            DataType = dataType ?? throw new ArgumentNullException(nameof(dataType));
            SampleValue = sampleValue ?? throw new ArgumentNullException(nameof(sampleValue));
        }

        public string Name { get; }

        public string DataType { get; }

        public string SampleValue { get; }
    }

    public sealed class PeriodMappingSnapshot
    {
        public PeriodMappingSnapshot(
            string mode,
            string periodColumn,
            string headerPattern,
            int? reportingYear,
            IReadOnlyList<WideHeaderMappingRowSnapshot> wideHeaderMappings)
        {
            Mode = mode ?? throw new ArgumentNullException(nameof(mode));
            PeriodColumn = periodColumn ?? string.Empty;
            HeaderPattern = headerPattern ?? string.Empty;
            ReportingYear = reportingYear;
            WideHeaderMappings = wideHeaderMappings
                ?? throw new ArgumentNullException(nameof(wideHeaderMappings));
        }

        public string Mode { get; }

        public string PeriodColumn { get; }

        public string HeaderPattern { get; }

        public int? ReportingYear { get; }

        public IReadOnlyList<WideHeaderMappingRowSnapshot> WideHeaderMappings { get; }
    }

    public enum TotalPreservationState
    {
        NotChecked,
        Pass,
        Fail
    }

    public sealed class WideHeaderMappingPreview
    {
        public WideHeaderMappingPreview(
            IReadOnlyList<WideHeaderMappingRowSnapshot> headerMappings,
            long projectedNormalizedRowCount,
            IReadOnlyList<NormalizedSampleRowSnapshot> sampleRows,
            TotalPreservationState totalPreservation,
            string totalPreservationDetail,
            bool requiresReportingYear)
        {
            HeaderMappings = headerMappings ?? throw new ArgumentNullException(nameof(headerMappings));
            ProjectedNormalizedRowCount = projectedNormalizedRowCount;
            SampleRows = sampleRows ?? throw new ArgumentNullException(nameof(sampleRows));
            TotalPreservation = totalPreservation;
            TotalPreservationDetail = totalPreservationDetail
                ?? throw new ArgumentNullException(nameof(totalPreservationDetail));
            RequiresReportingYear = requiresReportingYear;
        }

        public IReadOnlyList<WideHeaderMappingRowSnapshot> HeaderMappings { get; }

        public long ProjectedNormalizedRowCount { get; }

        public IReadOnlyList<NormalizedSampleRowSnapshot> SampleRows { get; }

        public TotalPreservationState TotalPreservation { get; }

        public string TotalPreservationDetail { get; }

        public bool RequiresReportingYear { get; }
    }

    public sealed class WideHeaderMappingRowSnapshot
    {
        public WideHeaderMappingRowSnapshot(
            string sourceHeader,
            string period,
            string metric,
            double confidence)
        {
            SourceHeader = sourceHeader ?? throw new ArgumentNullException(nameof(sourceHeader));
            Period = period ?? throw new ArgumentNullException(nameof(period));
            Metric = metric ?? throw new ArgumentNullException(nameof(metric));
            Confidence = Math.Max(0, Math.Min(1, confidence));
        }

        public string SourceHeader { get; }

        public string Period { get; }

        public string Metric { get; }

        public double Confidence { get; }

        public string ConfidenceLabel => Confidence.ToString("P0");
    }

    public sealed class NormalizedSampleRowSnapshot
    {
        public NormalizedSampleRowSnapshot(
            string sourceRow,
            string period,
            string metric,
            string value)
        {
            SourceRow = sourceRow ?? throw new ArgumentNullException(nameof(sourceRow));
            Period = period ?? throw new ArgumentNullException(nameof(period));
            Metric = metric ?? throw new ArgumentNullException(nameof(metric));
            Value = value ?? throw new ArgumentNullException(nameof(value));
        }

        public string SourceRow { get; }

        public string Period { get; }

        public string Metric { get; }

        public string Value { get; }
    }

    public enum PlacementBucket
    {
        Rows,
        Columns,
        Values,
        Filters
    }

    public sealed class FieldPlacementSnapshot
    {
        public FieldPlacementSnapshot(
            PlacementBucket bucket,
            string columnName,
            string setting,
            bool showSubtotals = true,
            IReadOnlyList<string>? selectedValues = null,
            string subtotalPlacement = "After members",
            IReadOnlyList<string>? memberOrder = null,
            string numberFormat = "#,##0.00")
        {
            Bucket = bucket;
            ColumnName = columnName ?? throw new ArgumentNullException(nameof(columnName));
            Setting = setting ?? string.Empty;
            ShowSubtotals = showSubtotals;
            SelectedValues = selectedValues ?? Array.Empty<string>();
            SubtotalPlacement = subtotalPlacement ?? "After members";
            MemberOrder = memberOrder ?? Array.Empty<string>();
            NumberFormat = numberFormat ?? "General";
        }

        public PlacementBucket Bucket { get; }

        public string ColumnName { get; }

        public string Setting { get; }

        public bool ShowSubtotals { get; }

        public IReadOnlyList<string> SelectedValues { get; }

        public string SubtotalPlacement { get; }

        public IReadOnlyList<string> MemberOrder { get; }

        public string NumberFormat { get; }
    }

    public sealed class ReportSpecificationSnapshot
    {
        public ReportSpecificationSnapshot(
            PeriodMappingSnapshot periodMapping,
            IReadOnlyList<FieldPlacementSnapshot> placements,
            string outputStyle = "Dense management block",
            string canonicalReportSpecJson = "",
            IReadOnlyList<ManualTransformSnapshot>? transforms = null,
            IReadOnlyList<ManualCalculatedMetricSnapshot>? calculatedMetrics = null,
            IReadOnlyList<ManualReportBlockSnapshot>? blocks = null,
            ManualLayoutSnapshot? layout = null,
            IReadOnlyList<ManualCheckSnapshot>? checks = null,
            bool manualProjectionComplete = true)
        {
            PeriodMapping = periodMapping ?? throw new ArgumentNullException(nameof(periodMapping));
            Placements = placements ?? throw new ArgumentNullException(nameof(placements));
            OutputStyle = outputStyle ?? "Dense management block";
            CanonicalReportSpecJson = canonicalReportSpecJson ?? string.Empty;
            Transforms = transforms ?? Array.Empty<ManualTransformSnapshot>();
            CalculatedMetrics = calculatedMetrics ?? Array.Empty<ManualCalculatedMetricSnapshot>();
            Blocks = blocks ?? Array.Empty<ManualReportBlockSnapshot>();
            Layout = layout ?? new ManualLayoutSnapshot();
            Checks = checks ?? Array.Empty<ManualCheckSnapshot>();
            ManualProjectionComplete = manualProjectionComplete;
        }

        public PeriodMappingSnapshot PeriodMapping { get; }

        public IReadOnlyList<FieldPlacementSnapshot> Placements { get; }

        public string OutputStyle { get; }

        /// <summary>
        /// Optional bounded ReportSpecV1 JSON produced by the guarded agent path.
        /// The manual projection remains available for display, but the host uses
        /// this versioned typed contract so advanced blocks and measures are not
        /// silently flattened or discarded.
        /// </summary>
        public string CanonicalReportSpecJson { get; }

        public IReadOnlyList<ManualTransformSnapshot> Transforms { get; }

        public IReadOnlyList<ManualCalculatedMetricSnapshot> CalculatedMetrics { get; }

        public IReadOnlyList<ManualReportBlockSnapshot> Blocks { get; }

        public ManualLayoutSnapshot Layout { get; }

        public IReadOnlyList<ManualCheckSnapshot> Checks { get; }

        /// <summary>
        /// True only when every canonical setting is represented by the manual
        /// editor snapshot. A false value keeps the canonical setup available
        /// for rebuild and Chat, but manual controls must remain read-only so an
        /// edit cannot silently flatten unsupported blocks or calculations.
        /// </summary>
        public bool ManualProjectionComplete { get; }

        public bool HasCanonicalReportSpec => !string.IsNullOrWhiteSpace(CanonicalReportSpecJson);
    }

    public sealed class ManualTransformSnapshot
    {
        public ManualTransformSnapshot(string operation, string column, string outputColumn, string details)
        {
            Operation = operation ?? string.Empty;
            Column = column ?? string.Empty;
            OutputColumn = outputColumn ?? string.Empty;
            Details = details ?? string.Empty;
        }

        public string Operation { get; }

        public string Column { get; }

        public string OutputColumn { get; }

        public string Details { get; }
    }

    public sealed class ManualCalculatedMetricSnapshot
    {
        public ManualCalculatedMetricSnapshot(
            string label,
            string kind,
            string primary,
            string secondary,
            string details,
            string numberFormat)
        {
            Label = label ?? string.Empty;
            Kind = kind ?? string.Empty;
            Primary = primary ?? string.Empty;
            Secondary = secondary ?? string.Empty;
            Details = details ?? string.Empty;
            NumberFormat = numberFormat ?? "General";
        }

        public string Label { get; }

        public string Kind { get; }

        public string Primary { get; }

        public string Secondary { get; }

        public string Details { get; }

        public string NumberFormat { get; }
    }

    public sealed class ManualReportBlockSnapshot
    {
        public ManualReportBlockSnapshot(
            string title,
            string worksheetName,
            string anchorCell,
            string outputStyle,
            string? stableId = null,
            int ownedRows = 500,
            int ownedColumns = 64,
            string? canonicalBlockId = null,
            string? canonicalOwnershipId = null)
        {
            Title = title ?? string.Empty;
            WorksheetName = worksheetName ?? string.Empty;
            AnchorCell = anchorCell ?? string.Empty;
            OutputStyle = outputStyle ?? "Dense management block";
            StableId = string.IsNullOrWhiteSpace(stableId) ? null : stableId;
            OwnedRows = ownedRows;
            OwnedColumns = ownedColumns;
            CanonicalBlockId = string.IsNullOrWhiteSpace(canonicalBlockId)
                ? null
                : canonicalBlockId;
            CanonicalOwnershipId = string.IsNullOrWhiteSpace(canonicalOwnershipId)
                ? null
                : canonicalOwnershipId;
        }

        public string Title { get; }

        public string WorksheetName { get; }

        public string AnchorCell { get; }

        public string OutputStyle { get; }

        /// <summary>
        /// Optional for compatibility with older manual snapshots. New block
        /// rules generate this once so rebuilds retain managed object identity.
        /// </summary>
        public string? StableId { get; }

        public int OwnedRows { get; }

        public int OwnedColumns { get; }

        public string? CanonicalBlockId { get; }

        public string? CanonicalOwnershipId { get; }
    }

    public sealed class ManualLayoutSnapshot
    {
        public bool RepeatRowLabels { get; set; }

        public bool InsertBlankRows { get; set; }

        public bool FreezeHeaders { get; set; } = true;

        public bool ShowRowGrandTotals { get; set; } = true;

        public bool ShowColumnGrandTotals { get; set; } = true;

        public int RowIndent { get; set; } = 1;

        public string RowGrandTotalLabel { get; set; } = "Grand Total";

        public string ColumnGrandTotalLabel { get; set; } = "Grand Total";
    }

    public sealed class ManualCheckSnapshot
    {
        public ManualCheckSnapshot(string kind, string metric, string comparedMetric, decimal tolerance)
        {
            Kind = kind ?? string.Empty;
            Metric = metric ?? string.Empty;
            ComparedMetric = comparedMetric ?? string.Empty;
            Tolerance = tolerance;
        }

        public string Kind { get; }

        public string Metric { get; }

        public string ComparedMetric { get; }

        public decimal Tolerance { get; }
    }

    public sealed class BuildDraftResult
    {
        public BuildDraftResult(string draftName, int outputRows)
        {
            DraftName = draftName ?? throw new ArgumentNullException(nameof(draftName));
            OutputRows = outputRows;
        }

        public string DraftName { get; }

        public int OutputRows { get; }
    }

    public sealed class ReportSetupValidationException : Exception
    {
        public ReportSetupValidationException(string message, Exception innerException)
            : base(Normalize(message), innerException)
        {
        }

        private static string Normalize(string message)
        {
            string normalized = (message ?? "The report setup is not valid.")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Trim();
            return normalized.Length <= 500
                ? normalized
                : normalized.Substring(0, 500);
        }
    }

    public sealed class ChatRunResult
    {
        public ChatRunResult(
            string response,
            ReportSpecificationSnapshot? appliedSpecification = null,
            string managedDraftName = "",
            int outputRows = 0,
            IReadOnlyList<HostCheckResult>? checks = null,
            IReadOnlyList<ChatChangeSnapshot>? changes = null,
            bool allChecksPassed = false,
            bool published = false)
        {
            Response = response ?? throw new ArgumentNullException(nameof(response));
            AppliedSpecification = appliedSpecification;
            ManagedDraftName = managedDraftName ?? string.Empty;
            OutputRows = outputRows;
            Checks = checks ?? Array.Empty<HostCheckResult>();
            Changes = changes ?? Array.Empty<ChatChangeSnapshot>();
            AllChecksPassed = allChecksPassed;
            Published = published;
        }

        public string Response { get; }

        public ReportSpecificationSnapshot? AppliedSpecification { get; }

        public string ManagedDraftName { get; }

        public int OutputRows { get; }

        public IReadOnlyList<HostCheckResult> Checks { get; }

        public IReadOnlyList<ChatChangeSnapshot> Changes { get; }

        public bool HasManagedDraft => AppliedSpecification != null && !string.IsNullOrWhiteSpace(ManagedDraftName);

        public bool AllChecksPassed { get; }

        public bool Published { get; }
    }

    public sealed class ChatChangeSnapshot
    {
        public ChatChangeSnapshot(string category, string description)
        {
            Category = category ?? throw new ArgumentNullException(nameof(category));
            Description = description ?? throw new ArgumentNullException(nameof(description));
        }

        public string Category { get; }

        public string Description { get; }
    }

    public sealed class ModelEndpointSettingsSnapshot
    {
        public ModelEndpointSettingsSnapshot(
            string baseUrl,
            string modelId,
            bool allowRemoteHttp,
            bool allowRemoteWorkbookData = false)
        {
            BaseUrl = baseUrl ?? throw new ArgumentNullException(nameof(baseUrl));
            ModelId = modelId ?? string.Empty;
            AllowRemoteHttp = allowRemoteHttp;
            AllowRemoteWorkbookData = allowRemoteWorkbookData;
        }

        public string BaseUrl { get; }

        public string ModelId { get; }

        public bool AllowRemoteHttp { get; }

        public bool AllowRemoteWorkbookData { get; }
    }

    public sealed class SavedEndpointSettingsSnapshot
    {
        public SavedEndpointSettingsSnapshot(
            string baseUrl,
            string modelId,
            bool allowRemoteHttp,
            bool hasProtectedApiKey,
            bool allowRemoteWorkbookData = false)
        {
            BaseUrl = baseUrl ?? throw new ArgumentNullException(nameof(baseUrl));
            ModelId = modelId ?? string.Empty;
            AllowRemoteHttp = allowRemoteHttp;
            HasProtectedApiKey = hasProtectedApiKey;
            AllowRemoteWorkbookData = allowRemoteWorkbookData;
        }

        public string BaseUrl { get; }

        public string ModelId { get; }

        public bool AllowRemoteHttp { get; }

        public bool HasProtectedApiKey { get; }

        public bool AllowRemoteWorkbookData { get; }
    }

    public sealed class EndpointCheckResult
    {
        public EndpointCheckResult(bool succeeded, string message)
        {
            Succeeded = succeeded;
            Message = message ?? throw new ArgumentNullException(nameof(message));
        }

        public bool Succeeded { get; }

        public string Message { get; }
    }

    public sealed class HostCheckResult
    {
        public HostCheckResult(string name, bool passed, string detail)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Passed = passed;
            Detail = detail ?? throw new ArgumentNullException(nameof(detail));
        }

        public string Name { get; }

        public bool Passed { get; }

        public string Detail { get; }
    }

    public sealed class PublishResult
    {
        public PublishResult(string message)
        {
            Message = message ?? throw new ArgumentNullException(nameof(message));
        }

        public string Message { get; }
    }
}
