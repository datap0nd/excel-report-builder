using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.Validation;
using CorePivotFieldArea = ExcelReportBuilder.Core.PivotPlus.PivotFieldArea;

namespace ExcelReportBuilder.Excel.PivotPlus.Native
{
    public sealed class PivotTableNativeMutationValidationException : Exception
    {
        internal PivotTableNativeMutationValidationException(
            string message,
            IReadOnlyList<ValidationIssue> issues)
            : base(message)
        {
            Issues = issues ?? throw new ArgumentNullException(nameof(issues));
        }

        public IReadOnlyList<ValidationIssue> Issues { get; }
    }

    internal sealed class NativePivotFieldCommand
    {
        public string InstanceId { get; set; } = string.Empty;

        public string FieldName { get; set; } = string.Empty;

        public string Caption { get; set; } = string.Empty;

        public bool SetCaption { get; set; }

        public CorePivotFieldArea Area { get; set; }

        public int Position { get; set; }

        public bool IsMeasure { get; set; }

        public int? ConsolidationFunction { get; set; }

        public string? NumberFormatCode { get; set; }

        public PivotSubtotalMode SubtotalMode { get; set; }
    }

    internal sealed class NativePivotLayoutCommand
    {
        public int RowAxisLayout { get; set; }

        public bool RepeatItemLabels { get; set; }

        public bool ShowRowGrandTotals { get; set; }

        public bool ShowColumnGrandTotals { get; set; }

        public bool ShowFieldHeaders { get; set; }

        public PivotValuesAxis ValuesAxis { get; set; }

        public int ValuesPosition { get; set; } = 1;

        public string? PivotTableStyleName { get; set; }

        public bool SetPivotTableStyle { get; set; }

        public bool PreserveFormatting { get; set; }

        public bool ShowRowStripes { get; set; }

        public bool ShowColumnStripes { get; set; }
    }

    internal sealed class NativePivotMutationPlan
    {
        public PivotSourceKind SourceKind { get; set; }

        public IReadOnlyList<NativePivotFieldCommand> Fields { get; set; } =
            Array.Empty<NativePivotFieldCommand>();

        public NativePivotLayoutCommand Layout { get; set; } = new NativePivotLayoutCommand();
    }

    internal enum NativePivotCacheKind
    {
        ClassicDatabase,
        DataModel,
        ExternalOlap
    }

    internal sealed class NativePivotSourceIdentity
    {
        public NativePivotSourceIdentity(
            NativePivotCacheKind kind,
            string sourceName)
        {
            Kind = kind;
            SourceName = sourceName ?? throw new ArgumentNullException(nameof(sourceName));
        }

        public NativePivotCacheKind Kind { get; }

        public string SourceName { get; }
    }

    internal interface IPivotTableNativeAdapter
    {
        PivotTargetIdentity ReadTarget(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver);

        void PersistWorkbookIdentity(
            object pivotTable,
            IWorkbookIdentityResolver workbookIdentityResolver,
            string expectedWorkbookId);

        NativePivotSourceIdentity ReadSource(object pivotTable);

        object CaptureState(object pivotTable, PivotSourceKind sourceKind);

        void ClearLayout(object pivotTable, PivotSourceKind sourceKind);

        void PlaceField(
            object pivotTable,
            PivotSourceKind sourceKind,
            NativePivotFieldCommand command);

        void ApplyLayout(object pivotTable, NativePivotLayoutCommand command);

        void RestoreState(object pivotTable, object snapshot);

        void Refresh(object pivotTable);

        void Verify(object pivotTable, NativePivotMutationPlan plan);
    }

    /// <summary>
    /// Applies a validated Core PivotTable+ definition directly to one native
    /// PivotTable. It never creates worksheet formulas or a companion report.
    /// All COM access is isolated behind <see cref="IPivotTableNativeAdapter"/>.
    /// </summary>
    public sealed class PivotTableNativeLayoutMutationService
    {
        private const int MaximumFields = 512;
        private const int MaximumPlacements = 512;
        private const int MaximumFilters = 64;

        private readonly IPivotTableNativeAdapter adapter;
        private readonly PivotMutationCoordinator coordinator;
        private readonly IWorkbookIdentityResolver workbookIdentityResolver;

        public PivotTableNativeLayoutMutationService()
            : this(
                new LateBoundPivotTableNativeAdapter(),
                new PivotMutationCoordinator(),
                new StoredWorkbookIdentityResolver())
        {
        }

        internal PivotTableNativeLayoutMutationService(
            IPivotTableNativeAdapter adapter,
            PivotMutationCoordinator coordinator)
            : this(adapter, coordinator, new StoredWorkbookIdentityResolver())
        {
        }

        internal PivotTableNativeLayoutMutationService(
            IPivotTableNativeAdapter adapter,
            PivotMutationCoordinator coordinator,
            IWorkbookIdentityResolver workbookIdentityResolver)
        {
            this.adapter = adapter ?? throw new ArgumentNullException(nameof(adapter));
            this.coordinator = coordinator ?? throw new ArgumentNullException(nameof(coordinator));
            this.workbookIdentityResolver = workbookIdentityResolver ??
                throw new ArgumentNullException(nameof(workbookIdentityResolver));
        }

        public void Apply(
            object pivotTable,
            PivotTableContext context,
            PivotLayoutDefinition definition)
        {
            if (pivotTable == null) throw new ArgumentNullException(nameof(pivotTable));
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (definition == null) throw new ArgumentNullException(nameof(definition));

            NativePivotMutationPlan plan = Compile(context, definition);
            PivotTargetIdentity liveTarget = adapter.ReadTarget(
                pivotTable,
                workbookIdentityResolver);
            DemandTarget(liveTarget, definition.Target);
            NativePivotSourceIdentity liveSource = adapter.ReadSource(pivotTable);
            DemandLiveSource(liveSource, definition.Source);
            object snapshot = adapter.CaptureState(pivotTable, definition.Source.Kind);
            adapter.PersistWorkbookIdentity(
                pivotTable,
                workbookIdentityResolver,
                definition.Target.WorkbookId);
            var steps = new List<PivotMutationStep>
            {
                new PivotMutationStep(
                    "clear-native-layout",
                    () => adapter.ClearLayout(pivotTable, plan.SourceKind),
                    () => adapter.RestoreState(pivotTable, snapshot))
            };
            steps.AddRange(plan.Fields.Select(command => new PivotMutationStep(
                "place-" + command.InstanceId,
                () => adapter.PlaceField(pivotTable, plan.SourceKind, command),
                () => { })));
            steps.Add(new PivotMutationStep(
                "layout-and-format",
                () => adapter.ApplyLayout(pivotTable, plan.Layout),
                () => { }));

            coordinator.Execute(
                pivotTable,
                steps,
                () => adapter.Refresh(pivotTable),
                () => adapter.Verify(pivotTable, plan));
        }

        internal NativePivotMutationPlan Compile(
            PivotTableContext context,
            PivotLayoutDefinition definition)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (definition == null) throw new ArgumentNullException(nameof(definition));

            DemandBoundedDefinition(definition);
            DemandCoreValidation(definition);
            DemandTarget(context.Definition.Target, definition.Target);
            DemandSource(context, definition);

            if ((definition.Placements.Count == 0) != definition.ClearAll)
            {
                throw new InvalidOperationException(
                    "An empty native layout requires explicit clearAll intent.");
            }

            if (definition.Filters.Count > 0)
            {
                throw new NotSupportedException(
                    "Native member-selection filters are not part of this field/layout mutation boundary yet.");
            }

            var descriptors = definition.Fields.ToDictionary(
                field => field.Name,
                StringComparer.OrdinalIgnoreCase);
            var liveFields = context.Definition.Fields
                .GroupBy(field => field.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            var commands = definition.Placements
                .Select(placement => CompilePlacement(
                    definition.Source,
                    descriptors[placement.FieldName],
                    ResolveLiveField(liveFields, context.Definition.Fields, placement.FieldName),
                    placement))
                .OrderBy(command => AreaOrder(command.Area))
                .ThenBy(command => command.Position)
                .ToList();

            DemandUnambiguousValueInstances(definition.Source.Kind, commands);

            return new NativePivotMutationPlan
            {
                SourceKind = definition.Source.Kind,
                Fields = new ReadOnlyCollection<NativePivotFieldCommand>(commands),
                Layout = new NativePivotLayoutCommand
                {
                    RowAxisLayout = RowAxisLayout(definition.Layout.Form),
                    RepeatItemLabels = definition.Layout.RepeatItemLabels,
                    ShowRowGrandTotals = definition.Layout.ShowRowGrandTotals,
                    ShowColumnGrandTotals = definition.Layout.ShowColumnGrandTotals,
                    ShowFieldHeaders = definition.Layout.ShowFieldHeaders,
                    ValuesAxis = definition.Layout.ValuesAxis,
                    ValuesPosition = definition.Layout.ValuesPosition,
                    PivotTableStyleName = definition.Format.PivotTableStyleName,
                    SetPivotTableStyle = definition.Format.PivotTableStyleName != null,
                    PreserveFormatting = definition.Format.PreserveFormatting,
                    ShowRowStripes = definition.Format.ShowRowStripes,
                    ShowColumnStripes = definition.Format.ShowColumnStripes
                }
            };
        }

        internal static int ConsolidationFunction(PivotAggregationFunction function)
        {
            switch (function)
            {
                case PivotAggregationFunction.Sum: return -4157;
                case PivotAggregationFunction.Count: return -4112;
                case PivotAggregationFunction.Average: return -4106;
                case PivotAggregationFunction.Minimum: return -4139;
                case PivotAggregationFunction.Maximum: return -4136;
                case PivotAggregationFunction.Product: return -4149;
                case PivotAggregationFunction.CountNumbers: return -4113;
                case PivotAggregationFunction.StandardDeviation: return -4155;
                case PivotAggregationFunction.StandardDeviationPopulation: return -4156;
                case PivotAggregationFunction.Variance: return -4164;
                case PivotAggregationFunction.VariancePopulation: return -4165;
                case PivotAggregationFunction.DistinctCount: return 11;
                default:
                    throw new ArgumentOutOfRangeException(nameof(function));
            }
        }

        private static NativePivotFieldCommand CompilePlacement(
            PivotSourceDescriptor source,
            PivotFieldDescriptor descriptor,
            PivotFieldDescriptor liveField,
            PivotFieldPlacement placement)
        {
            if (!IsClassic(source.Kind) && descriptor.IsMeasure != liveField.IsMeasure)
            {
                throw new InvalidOperationException(
                    "The live OLAP field kind no longer matches the validated discovery snapshot for '" +
                    placement.FieldName + "'.");
            }

            if (!Supports(liveField.SupportedAreas, placement.Area))
            {
                throw new InvalidOperationException(
                    "The live PivotTable field no longer supports the requested native area for '" +
                    placement.FieldName + "'.");
            }

            if (IsClassic(source.Kind) && descriptor.IsMeasure)
            {
                throw new NotSupportedException(
                    "A classic PivotTable cannot place an OLAP measure field.");
            }

            if (placement.Area == CorePivotFieldArea.Values &&
                source.Kind == PivotSourceKind.ExternalOlap &&
                !descriptor.IsMeasure)
            {
                throw new NotSupportedException(
                    "External OLAP sources can place existing measures in Values but cannot author implicit aggregations.");
            }

            if (placement.Area == CorePivotFieldArea.Values &&
                source.Kind == PivotSourceKind.DataModel &&
                !descriptor.IsMeasure &&
                placement.Aggregation.HasValue &&
                !SupportsDataModelImplicitMeasure(placement.Aggregation.Value))
            {
                throw new NotSupportedException(
                    "Excel CubeFields.GetMeasure supports only Sum, Count, Average, Minimum, and Maximum for implicit Data Model measures.");
            }

            if (placement.Aggregation == PivotAggregationFunction.DistinctCount &&
                source.Kind != PivotSourceKind.DataModel)
            {
                throw new NotSupportedException(
                    "Distinct count requires a Data Model-backed PivotTable.");
            }

            string caption = PivotPlusValueSemantics.ResolveCaption(descriptor, placement);
            return new NativePivotFieldCommand
            {
                InstanceId = placement.Area == CorePivotFieldArea.Values
                    ? "value:" + placement.Position.ToString("D4", CultureInfo.InvariantCulture) + ":" + caption
                    : placement.Area.ToString().ToLowerInvariant() + ":" +
                      placement.Position.ToString("D4", CultureInfo.InvariantCulture) + ":" +
                      descriptor.Name,
                FieldName = descriptor.Name,
                Caption = caption,
                SetCaption = placement.Area == CorePivotFieldArea.Values ||
                             placement.Caption != null,
                Area = placement.Area,
                Position = placement.Position,
                IsMeasure = descriptor.IsMeasure,
                ConsolidationFunction = placement.Aggregation.HasValue
                    ? ConsolidationFunction(placement.Aggregation.Value)
                    : (int?)null,
                NumberFormatCode = placement.NumberFormatCode,
                SubtotalMode = placement.SubtotalMode
            };
        }

        private static bool SupportsDataModelImplicitMeasure(
            PivotAggregationFunction function)
        {
            // Excel documents CubeFields.GetMeasure as supporting only these five functions.
            return function == PivotAggregationFunction.Sum ||
                   function == PivotAggregationFunction.Count ||
                   function == PivotAggregationFunction.Average ||
                   function == PivotAggregationFunction.Minimum ||
                   function == PivotAggregationFunction.Maximum;
        }

        private static void DemandCoreValidation(PivotLayoutDefinition definition)
        {
            ValidationResult validation = PivotPlusValidator.Validate(definition);
            IReadOnlyList<ValidationIssue> blocking = validation.Issues
                .Where(issue => issue.Severity == ValidationSeverity.Error)
                .ToList();
            if (blocking.Count > 0)
            {
                throw new PivotTableNativeMutationValidationException(
                    "The PivotTable+ layout definition is invalid: " + blocking[0].Message,
                    blocking);
            }
        }

        private static void DemandBoundedDefinition(PivotLayoutDefinition definition)
        {
            if (definition.Fields.Count > MaximumFields ||
                definition.Placements.Count > MaximumPlacements ||
                definition.Filters.Count > MaximumFilters)
            {
                throw new InvalidOperationException(
                    "The PivotTable+ native mutation exceeds its bounded field or filter limit.");
            }
        }

        private static void DemandTarget(
            PivotTargetIdentity discovered,
            PivotTargetIdentity target)
        {
            if (!string.Equals(
                    discovered.WorkbookId,
                    target.WorkbookId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    discovered.WorksheetName,
                    target.WorksheetName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    discovered.PivotTableName,
                    target.PivotTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable no longer matches the validated PivotTable+ target.");
            }
        }

        private static void DemandSource(
            PivotTableContext context,
            PivotLayoutDefinition definition)
        {
            PivotSourceDescriptor source = definition.Source;
            if (!context.IsConnected || !context.SourceFieldsComplete)
            {
                throw new NotSupportedException(
                    "The selected PivotTable does not expose a complete editable native layout.");
            }

            PivotSourceDescriptor discovered = context.Definition.Source;
            if (source.Kind == PivotSourceKind.Unknown)
            {
                throw new NotSupportedException("The PivotTable source kind is not mutable.");
            }

            if (discovered.Kind != source.Kind ||
                !string.Equals(
                    discovered.SourceName,
                    source.SourceName,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    discovered.ModelTableName,
                    source.ModelTableName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable source changed after discovery.");
            }

            const PivotCapability required =
                PivotCapability.NativeFieldPlacement |
                PivotCapability.LayoutFormatting |
                PivotCapability.Refresh;
            if ((source.Capabilities & required) != required ||
                (discovered.Capabilities & required) != required)
            {
                throw new NotSupportedException(
                    "The selected PivotTable source lacks native placement, formatting, or refresh capability.");
            }

            if (discovered.Capabilities != source.Capabilities)
            {
                throw new InvalidOperationException(
                    "The selected PivotTable capabilities changed after discovery.");
            }

            if (source.Kind == PivotSourceKind.DataModel &&
                definition.Placements.Any(placement =>
                    placement.Area == CorePivotFieldArea.Values) &&
                (discovered.Capabilities & PivotCapability.ModelMeasures) == 0)
            {
                throw new NotSupportedException(
                    "The selected Data Model PivotTable cannot expose model measures.");
            }
        }

        private static PivotFieldDescriptor ResolveLiveField(
            IReadOnlyDictionary<string, PivotFieldDescriptor> byName,
            IReadOnlyList<PivotFieldDescriptor> allFields,
            string fieldName)
        {
            if (byName.TryGetValue(fieldName, out PivotFieldDescriptor? exact))
            {
                return exact;
            }

            throw new InvalidOperationException(
                "The live PivotTable no longer exposes the field '" + fieldName + "'.");
        }

        private static void DemandUnambiguousValueInstances(
            PivotSourceKind sourceKind,
            IReadOnlyList<NativePivotFieldCommand> commands)
        {
            List<NativePivotFieldCommand> values = commands
                .Where(command => command.Area == CorePivotFieldArea.Values)
                .ToList();
            var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (NativePivotFieldCommand value in values)
            {
                if (!captions.Add(value.Caption))
                {
                    throw new InvalidOperationException(
                        "Each Values instance requires a unique PivotTable caption.");
                }
            }

            if (!IsClassic(sourceKind) &&
                values.Where(value => value.IsMeasure)
                    .GroupBy(value => value.FieldName, StringComparer.OrdinalIgnoreCase)
                    .Any(group => group.Count() > 1))
            {
                throw new NotSupportedException(
                    "Excel cannot place the same existing OLAP measure more than once without a separately authored measure.");
            }

            if (sourceKind == PivotSourceKind.DataModel)
            {
                List<NativePivotFieldCommand> implicitMeasures = values
                    .Where(value => !value.IsMeasure)
                    .ToList();
                for (var index = 0; index < implicitMeasures.Count; index++)
                {
                    for (var other = index + 1; other < implicitMeasures.Count; other++)
                    {
                        if (string.Equals(
                                implicitMeasures[index].FieldName,
                                implicitMeasures[other].FieldName,
                                StringComparison.OrdinalIgnoreCase) &&
                            implicitMeasures[index].ConsolidationFunction ==
                            implicitMeasures[other].ConsolidationFunction)
                        {
                            throw new NotSupportedException(
                                "Excel CubeFields.GetMeasure cannot create two independent implicit measures from the same source field and aggregation; different captions still address the same measure.");
                        }
                    }
                }
            }
        }

        private static void DemandLiveSource(
            NativePivotSourceIdentity live,
            PivotSourceDescriptor expected)
        {
            if (live == null) throw new ArgumentNullException(nameof(live));
            if (expected == null) throw new ArgumentNullException(nameof(expected));

            NativePivotCacheKind expectedKind;
            switch (expected.Kind)
            {
                case PivotSourceKind.WorksheetRange:
                case PivotSourceKind.WorksheetTable:
                    expectedKind = NativePivotCacheKind.ClassicDatabase;
                    break;
                case PivotSourceKind.DataModel:
                    expectedKind = NativePivotCacheKind.DataModel;
                    break;
                case PivotSourceKind.ExternalOlap:
                    expectedKind = NativePivotCacheKind.ExternalOlap;
                    break;
                default:
                    throw new NotSupportedException(
                        "The selected PivotTable source kind is not mutable.");
            }

            if (live.Kind != expectedKind ||
                !string.Equals(
                    live.SourceName,
                    expected.SourceName,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The selected PivotTable's live PivotCache no longer matches the validated PivotTable+ source.");
            }
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

        private static int RowAxisLayout(PivotLayoutForm form)
        {
            switch (form)
            {
                case PivotLayoutForm.Compact: return 0;
                case PivotLayoutForm.Tabular: return 1;
                case PivotLayoutForm.Outline: return 2;
                default: throw new ArgumentOutOfRangeException(nameof(form));
            }
        }

        private static bool IsClassic(PivotSourceKind sourceKind)
        {
            return sourceKind == PivotSourceKind.WorksheetRange ||
                   sourceKind == PivotSourceKind.WorksheetTable;
        }

        private static bool Supports(PivotFieldAreaSupport support, CorePivotFieldArea area)
        {
            switch (area)
            {
                case CorePivotFieldArea.Row: return (support & PivotFieldAreaSupport.Row) != 0;
                case CorePivotFieldArea.Column: return (support & PivotFieldAreaSupport.Column) != 0;
                case CorePivotFieldArea.Filter: return (support & PivotFieldAreaSupport.Filter) != 0;
                case CorePivotFieldArea.Values: return (support & PivotFieldAreaSupport.Values) != 0;
                default: return false;
            }
        }
    }
}
