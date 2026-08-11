using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.PowerQuery;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Planning
{
    public static class ReportBuildPlanner
    {
        public static ReportBuildPlan Create(
            ReportSpecV1 specification,
            SourceProfile sourceProfile,
            PeriodDetectionResult? periodDetection = null)
        {
            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            if (sourceProfile == null)
            {
                throw new ArgumentNullException(nameof(sourceProfile));
            }

            var validation = ReportSpecValidator.Validate(specification, sourceProfile);
            if (periodDetection != null)
            {
                if (periodDetection.IsAmbiguous)
                {
                    validation.AddError(
                        "PERIOD_DETECTION_AMBIGUOUS",
                        "$.periodMapping",
                        "The detected period layout still requires an explicit user resolution.");
                }

                if ((periodDetection.Kind == PeriodLayoutKind.MonthHeaders
                        || periodDetection.Kind == PeriodLayoutKind.MetricMonthHeaders)
                    && specification.PeriodMapping == null)
                {
                    validation.AddError(
                        "PERIOD_MAPPING_REQUIRED",
                        "$.periodMapping",
                        "A detected wide period layout requires an explicit periodMapping in the report specification.");
                }
            }

            if (!validation.IsValid)
            {
                throw new InvalidReportSpecException(validation);
            }

            var projection = specification.PeriodMapping != null
                ? RowProjectionCalculator.Project(sourceProfile.RowCount, specification.PeriodMapping)
                : periodDetection != null
                    ? RowProjectionCalculator.Project(sourceProfile.RowCount, periodDetection)
                    : RowProjectionCalculator.Project(sourceProfile.RowCount, (PeriodMappingSpec?)null);
            if (RequiresDataModel(specification.Measures))
            {
                projection.Route = SourceLoadRoute.DataModel;
                projection.Reason = "A distinct-count measure requires a Data Model-backed aggregate without truncation.";
            }
            var compilation = PowerQueryMCompiler.Compile(specification);
            var plan = new ReportBuildPlan
            {
                SpecificationId = specification.Id,
                SchemaVersion = specification.SchemaVersion,
                SpecificationHash = ReportSpecDigest.Compute(specification),
                OwnershipId = specification.OwnershipId,
                Source = new SourcePreparationPlan
                {
                    WorkbookObjectName = specification.Source.WorkbookObjectName,
                    Fingerprint = specification.Source.Fingerprint,
                    SavedSetupCompatibilityKey = specification.Source.Fingerprint.GetSavedSetupKey(),
                    ManagedQueryName = ManagedName(specification.OwnershipId, "Source"),
                    PowerQueryM = compilation.Query,
                    Route = projection.Route,
                    SourceRows = projection.SourceRows,
                    ProjectedRows = projection.ProjectedRows,
                    ExpansionFactor = projection.ExpansionFactor,
                    TruncationAllowed = false
                },
                Styles = new List<PresentationStyleSpec>(specification.Styles)
            };

            var sequence = 1;
            plan.Operations.Add(Operation(sequence++, BuildOperationKind.ProfileSource, specification.OwnershipId, "Profile the selected Data source."));
            plan.Operations.Add(Operation(sequence++, BuildOperationKind.PrepareSource, specification.OwnershipId, "Apply the validated Data preparation steps."));
            plan.Operations.Add(Operation(
                sequence++,
                projection.Route == SourceLoadRoute.Worksheet ? BuildOperationKind.LoadWorksheet : BuildOperationKind.LoadDataModel,
                specification.OwnershipId,
                projection.Route == SourceLoadRoute.Worksheet
                    ? "Load the prepared rows to a managed worksheet table."
                    : "Load the prepared rows to the Data Model without truncation."));

            var measures = specification.Measures.ToDictionary(measure => measure.Id, StringComparer.OrdinalIgnoreCase);
            foreach (var block in specification.Blocks)
            {
                plan.Blocks.Add(BuildBlock(
                    block,
                    measures,
                    projection.Route,
                    PeriodFieldName(specification.PeriodMapping)));
                plan.Operations.Add(Operation(sequence++, BuildOperationKind.CreatePivot, block.OwnershipId, "Create the managed native PivotTable for block '" + block.Id + "'."));
                plan.Operations.Add(Operation(
                    sequence++,
                    RenderOperation(block.OutputMode),
                    block.OwnershipId,
                    "Render report block '" + block.Id + "' at its owned anchor."));
            }

            AddMandatoryChecks(plan.Checks, specification.Checks, measures, specification.Transforms);
            plan.Operations.Add(Operation(sequence, BuildOperationKind.RunChecks, specification.OwnershipId, "Run all mandatory and configured Checks."));
            plan.PlanHash = ReportBuildPlanDigest.Compute(plan);
            return plan;
        }

        private static DenseReportBlockPlan BuildBlock(
            ReportBlockSpec block,
            Dictionary<string, MeasureDefinition> measures,
            SourceLoadRoute route,
            string? periodFieldName)
        {
            var pivot = new PivotTablePlan
            {
                ManagedPivotName = ManagedName(block.OwnershipId, "Pivot"),
                ManagedCacheName = ManagedName(block.OwnershipId, "Cache"),
                UseDataModel = route == SourceLoadRoute.DataModel,
                GrandTotals = block.Layout.GrandTotals
            };
            pivot.Rows.AddRange(block.Layout.Rows.Select(ToPivotField));
            pivot.Columns.AddRange(block.Layout.Columns.Select(ToPivotField));
            pivot.Filters.AddRange(block.Layout.Filters.Select(filter => new PivotFilterPlan
            {
                Field = filter.Field,
                SelectedValues = new List<ScalarValue>(filter.SelectedValues),
                IncludeBlank = filter.IncludeBlank
            }));
            foreach (var placement in block.Layout.Values)
            {
                var definition = measures[placement.MeasureId];
                pivot.Values.Add(BuildPivotValue(definition, measures, placement));
            }

            var displayedMeasureIds = new HashSet<string>(
                block.Layout.Values.Select(value => value.MeasureId),
                StringComparer.OrdinalIgnoreCase);
            var supportingMeasureIds = block.Layout.Rows
                .Concat(block.Layout.Columns)
                .Where(field => field.TopN != null)
                .Select(field => field.TopN!.MeasureId)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var measureId in supportingMeasureIds)
            {
                if (!displayedMeasureIds.Contains(measureId))
                {
                    pivot.SupportingValues.Add(BuildPivotValue(measures[measureId], measures, null));
                }
            }

            AddSupportingPivotFields(pivot, block, periodFieldName);

            var headerRows = block.Headers.Count == 0
                ? 0
                : block.Headers.Max(header => header.RelativeRow + 1);
            var result = new DenseReportBlockPlan
            {
                BlockId = block.Id,
                OwnershipId = block.OwnershipId,
                WorksheetName = block.WorksheetName,
                AnchorCell = block.AnchorCell.ToUpperInvariant(),
                OutputMode = block.OutputMode,
                OwnedRange = new OwnedRangePlan
                {
                    AnchorCell = block.AnchorCell.ToUpperInvariant(),
                    RowCount = block.OwnedExtent.RowCount,
                    ColumnCount = block.OwnedExtent.ColumnCount
                },
                Title = block.Title,
                Pivot = pivot,
                Presentation = new DensePresentationPlan
                {
                    Options = block.Layout.DenseLayout,
                    PeriodSlices = new List<PeriodSliceSpec>(block.PeriodSlices),
                    ResolvedPeriodSlices = new List<ResolvedPeriodSlice>(PeriodSliceResolver.Resolve(block.PeriodSlices)),
                    Headers = new List<ReportHeaderSpec>(block.Headers),
                    Spacers = new List<SpacerSpec>(block.Spacers),
                    HeaderStyleId = block.HeaderStyleId,
                    BodyStyleId = block.BodyStyleId,
                    SubtotalStyleId = block.SubtotalStyleId,
                    GrandTotalStyleId = block.GrandTotalStyleId
                }
            };

            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.Headers,
                RelativeRow = 0,
                RelativeColumn = 0,
                FixedRowCount = headerRows,
                DynamicColumns = true,
                StyleId = block.HeaderStyleId
            });
            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.RowHierarchy,
                RelativeRow = headerRows,
                RelativeColumn = 0,
                FixedColumnCount = Math.Max(1, block.Layout.Rows.Count),
                DynamicRows = true,
                StyleId = block.BodyStyleId
            });
            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.ColumnHierarchy,
                RelativeRow = headerRows,
                RelativeColumn = Math.Max(1, block.Layout.Rows.Count),
                FixedRowCount = Math.Max(1, block.Layout.Columns.Count),
                DynamicColumns = true,
                StyleId = block.HeaderStyleId
            });
            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.Values,
                RelativeRow = headerRows + Math.Max(1, block.Layout.Columns.Count),
                RelativeColumn = Math.Max(1, block.Layout.Rows.Count),
                DynamicRows = true,
                DynamicColumns = true,
                StyleId = block.BodyStyleId
            });
            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.Subtotals,
                RelativeRow = headerRows,
                RelativeColumn = 0,
                DynamicRows = true,
                DynamicColumns = true,
                StyleId = block.SubtotalStyleId
            });
            result.Regions.Add(new DenseBlockRegionPlan
            {
                Kind = DenseBlockRegionKind.GrandTotals,
                RelativeRow = headerRows,
                RelativeColumn = 0,
                DynamicRows = true,
                DynamicColumns = true,
                StyleId = block.GrandTotalStyleId
            });
            return result;
        }

        private static void AddSupportingPivotFields(
            PivotTablePlan pivot,
            ReportBlockSpec block,
            string? periodFieldName)
        {
            var placed = new HashSet<string>(
                pivot.Rows.Select(field => field.Field)
                    .Concat(pivot.Columns.Select(field => field.Field))
                    .Concat(pivot.Filters.Select(field => field.Field)),
                StringComparer.OrdinalIgnoreCase);
            var required = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in pivot.Values.Concat(pivot.SupportingValues))
            {
                foreach (var component in value.AggregateComponents)
                {
                    foreach (var filter in component.Filters)
                    {
                        required.Add(filter.Field);
                    }
                }
            }

            var usesSlices = block.PeriodSlices.Count != 0
                && pivot.Values.Concat(pivot.SupportingValues).Any(value =>
                    value.PeriodSliceIds.Count != 0
                    || value.AggregateComponents.Any(component => !string.IsNullOrWhiteSpace(component.PeriodSliceId)));
            if (usesSlices)
            {
                if (string.IsNullOrWhiteSpace(periodFieldName))
                {
                    throw new InvalidOperationException("Period slices require an explicit period field.");
                }

                required.Add(periodFieldName!);
            }

            foreach (var field in required
                .Where(field => !placed.Contains(field))
                .OrderBy(field => field, StringComparer.OrdinalIgnoreCase))
            {
                pivot.Filters.Add(new PivotFilterPlan
                {
                    Field = field,
                    IsSupportingField = true
                });
                placed.Add(field);
            }
        }

        private static PivotFieldPlan ToPivotField(FieldPlacementSpec field)
        {
            var plan = new PivotFieldPlan
            {
                Field = field.Field,
                Caption = field.Caption,
                Sort = field.Sort,
                Subtotals = field.Subtotals,
                MemberOrder = new List<ScalarValue>(field.MemberOrder),
                GroupBuckets = new List<MemberGroupBucketSpec>(field.GroupBuckets),
                TopN = field.TopN
            };

            if (field.MemberOrder.Count != 0)
            {
                plan.MemberStages.Add(PivotMemberStageKind.ApplyMemberOrder);
            }

            if (field.GroupBuckets.Count != 0)
            {
                plan.MemberStages.Add(PivotMemberStageKind.GroupMembers);
            }

            if (field.Sort == SortDirection.Ascending)
            {
                plan.MemberStages.Add(PivotMemberStageKind.SortAscending);
            }
            else if (field.Sort == SortDirection.Descending)
            {
                plan.MemberStages.Add(PivotMemberStageKind.SortDescending);
            }

            if (field.TopN != null)
            {
                plan.MemberStages.Add(PivotMemberStageKind.ApplyTopN);
                if (field.TopN.IncludeOthers)
                {
                    plan.MemberStages.Add(PivotMemberStageKind.AggregateOthers);
                }
            }

            return plan;
        }

        private static PivotValuePlan BuildPivotValue(
            MeasureDefinition definition,
            Dictionary<string, MeasureDefinition> measures,
            ValuePlacementSpec? placement)
        {
            return new PivotValuePlan
            {
                MeasureId = definition.Id,
                Label = placement?.Caption ?? definition.Label,
                ValueType = definition.ValueType,
                Expression = definition.Expression,
                NumberFormat = placement?.NumberFormat ?? definition.NumberFormat,
                PeriodSliceIds = placement == null
                    ? new List<string>()
                    : new List<string>(placement.PeriodSliceIds),
                RequiresPostAggregationCalculation = !(definition.Expression is AggregateMeasureExpression),
                AggregateComponents = BuildAggregateComponents(definition, measures)
            };
        }

        private static List<PivotAggregateComponentPlan> BuildAggregateComponents(
            MeasureDefinition definition,
            Dictionary<string, MeasureDefinition> measures)
        {
            var target = new List<PivotAggregateComponentPlan>();
            var visitedMeasures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectAggregateComponents(
                definition.Expression,
                definition.Id,
                AggregateComponentRole.Value,
                measures,
                visitedMeasures,
                target);
            return target;
        }

        private static void CollectAggregateComponents(
            MeasureExpression expression,
            string ownerId,
            AggregateComponentRole role,
            Dictionary<string, MeasureDefinition> measures,
            HashSet<string> visitedMeasures,
            List<PivotAggregateComponentPlan> target)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    AddAggregateComponent(
                        ownerId,
                        role,
                        aggregate.Field,
                        aggregate.Function,
                        aggregate.ResultType,
                        aggregate.PeriodSliceId,
                        new List<MeasureFilterSpec>(),
                        target);
                    break;
                case FilteredAggregateMeasureExpression filtered:
                    AddAggregateComponent(
                        ownerId,
                        role,
                        filtered.Field,
                        filtered.Function,
                        filtered.ResultType,
                        filtered.PeriodSliceId,
                        new List<MeasureFilterSpec>(filtered.Filters),
                        target);
                    break;
                case WeightedAggregateMeasureExpression weighted:
                    CollectAggregateComponents(
                        weighted.Numerator,
                        ownerId,
                        AggregateComponentRole.WeightedNumerator,
                        measures,
                        visitedMeasures,
                        target);
                    CollectAggregateComponents(
                        weighted.Denominator,
                        ownerId,
                        AggregateComponentRole.WeightedDenominator,
                        measures,
                        visitedMeasures,
                        target);
                    break;
                case ReferenceMeasureExpression reference:
                    MeasureDefinition referenced;
                    if (measures.TryGetValue(reference.MeasureId, out referenced)
                        && visitedMeasures.Add(reference.MeasureId))
                    {
                        CollectAggregateComponents(
                            referenced.Expression,
                            ownerId,
                            role == AggregateComponentRole.Value ? AggregateComponentRole.Input : role,
                            measures,
                            visitedMeasures,
                            target);
                        visitedMeasures.Remove(reference.MeasureId);
                    }

                    break;
                case BinaryMeasureExpression binary:
                    CollectAggregateComponents(binary.Left, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    CollectAggregateComponents(binary.Right, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    break;
                case SafeDivideMeasureExpression divide:
                    CollectAggregateComponents(divide.Numerator, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    CollectAggregateComponents(divide.Denominator, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    break;
                case RatioMeasureExpression ratio:
                    CollectAggregateComponents(ratio.Numerator, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    CollectAggregateComponents(ratio.Denominator, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    break;
                case DifferenceMeasureExpression difference:
                    CollectAggregateComponents(difference.Current, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    CollectAggregateComponents(difference.Baseline, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    break;
                case ShareMeasureExpression share:
                    CollectAggregateComponents(share.Part, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    CollectAggregateComponents(share.Whole, ownerId, AggregateComponentRole.Input, measures, visitedMeasures, target);
                    break;
            }
        }

        private static void AddAggregateComponent(
            string ownerId,
            AggregateComponentRole role,
            string field,
            AggregateFunction function,
            MeasureValueType valueType,
            string? periodSliceId,
            List<MeasureFilterSpec> filters,
            List<PivotAggregateComponentPlan> target)
        {
            var existing = target.FirstOrDefault(component => component.Role == role
                && component.Function == function
                && component.ValueType == valueType
                && string.Equals(component.Field, field, StringComparison.OrdinalIgnoreCase)
                && string.Equals(component.PeriodSliceId, periodSliceId, StringComparison.OrdinalIgnoreCase)
                && FiltersEquivalent(component.Filters, filters));
            if (existing != null)
            {
                return;
            }

            target.Add(new PivotAggregateComponentPlan
            {
                Id = ownerId + "_component_" + (target.Count + 1),
                Role = role,
                Field = field,
                Function = function,
                ValueType = valueType,
                PeriodSliceId = periodSliceId,
                Filters = filters
            });
        }

        private static bool FiltersEquivalent(
            IReadOnlyList<MeasureFilterSpec> left,
            IReadOnlyList<MeasureFilterSpec> right)
        {
            if (left.Count != right.Count)
            {
                return false;
            }

            var leftKeys = left.Select(FilterKey).OrderBy(value => value, StringComparer.Ordinal);
            var rightKeys = right.Select(FilterKey).OrderBy(value => value, StringComparer.Ordinal);
            return leftKeys.SequenceEqual(rightKeys, StringComparer.Ordinal);
        }

        private static string FilterKey(MeasureFilterSpec filter)
        {
            if (filter == null)
            {
                return "<null>";
            }

            var values = filter.Values == null
                ? Enumerable.Empty<string>()
                : filter.Values.Select(ScalarKey).OrderBy(value => value, StringComparer.Ordinal);
            return (filter.Field ?? string.Empty).ToUpperInvariant()
                + "|" + filter.Operator
                + "|" + string.Join(",", values);
        }

        private static string ScalarKey(ScalarValue value)
        {
            if (value == null)
            {
                return "<null>";
            }

            return value.Kind + "|" + (value.Text
                ?? (value.Number.HasValue ? value.Number.Value.ToString(CultureInfo.InvariantCulture) : null)
                ?? (value.Boolean.HasValue ? value.Boolean.Value.ToString() : null)
                ?? (value.Temporal.HasValue ? value.Temporal.Value.ToString("O", CultureInfo.InvariantCulture) : string.Empty));
        }

        private static void AddMandatoryChecks(
            List<BuildCheckPlan> target,
            List<ReportCheckSpec> configured,
            Dictionary<string, MeasureDefinition> measures,
            List<TransformStep> transforms)
        {
            var removesRows = transforms.Any(transform => transform is FilterRowsTransform
                || transform is ExcludeTotalRowsTransform);
            target.Add(new BuildCheckPlan
            {
                Id = "mandatory-no-truncation",
                Kind = ReportCheckKind.NoTruncation,
                Mandatory = true,
                EvaluationScope = CheckEvaluationScope.CanonicalData,
                RowCountExpectation = removesRows
                    ? RowCountExpectation.AtMostProjection
                    : RowCountExpectation.ExactProjection
            });
            if (!removesRows)
            {
                target.Add(new BuildCheckPlan
                {
                    Id = "mandatory-total-preservation",
                    Kind = ReportCheckKind.TotalPreservation,
                    Mandatory = true,
                    EvaluationScope = CheckEvaluationScope.CanonicalData
                });
            }

            target.Add(new BuildCheckPlan
            {
                Id = "mandatory-rendered-output-reconciliation",
                Kind = ReportCheckKind.TotalPreservation,
                Mandatory = true,
                EvaluationScope = CheckEvaluationScope.RenderedOutput
            });

            foreach (var check in configured)
            {
                if (target.Any(existing => existing.Kind == check.Kind
                    && string.Equals(existing.MeasureId, check.MeasureId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(existing.ComparedMeasureId, check.ComparedMeasureId, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                target.Add(new BuildCheckPlan
                {
                    Id = check.Id,
                    Kind = check.Kind,
                    MeasureId = check.MeasureId,
                    ComparedMeasureId = check.ComparedMeasureId,
                    Measure = BuildCheckMeasure(check.MeasureId, measures),
                    ComparedMeasure = BuildCheckMeasure(check.ComparedMeasureId, measures),
                    Tolerance = check.Tolerance,
                    Mandatory = false,
                    EvaluationScope = check.Kind == ReportCheckKind.NoTruncation
                        || check.Kind == ReportCheckKind.TotalPreservation
                            ? CheckEvaluationScope.CanonicalData
                            : CheckEvaluationScope.RenderedOutput,
                    RowCountExpectation = removesRows
                        ? RowCountExpectation.AtMostProjection
                        : RowCountExpectation.ExactProjection
                });
            }
        }

        private static PivotValuePlan? BuildCheckMeasure(
            string? measureId,
            Dictionary<string, MeasureDefinition> measures)
        {
            if (string.IsNullOrWhiteSpace(measureId))
            {
                return null;
            }

            return BuildPivotValue(measures[measureId!], measures, null);
        }

        private static BuildOperation Operation(
            int sequence,
            BuildOperationKind kind,
            string ownershipId,
            string description)
        {
            return new BuildOperation
            {
                Sequence = sequence,
                Kind = kind,
                OwnershipId = ownershipId,
                Description = description
            };
        }

        private static BuildOperationKind RenderOperation(ReportOutputMode outputMode)
        {
            switch (outputMode)
            {
                case ReportOutputMode.StandardMatrix:
                    return BuildOperationKind.RenderStandardMatrix;
                case ReportOutputMode.MetricStack:
                    return BuildOperationKind.RenderMetricStack;
                case ReportOutputMode.DenseGrid:
                    return BuildOperationKind.RenderDenseGrid;
                default:
                    throw new ArgumentOutOfRangeException(nameof(outputMode));
            }
        }

        private static string? PeriodFieldName(PeriodMappingSpec? mapping)
        {
            if (mapping == null)
            {
                return null;
            }

            return mapping.Kind == PeriodMappingKind.LongDateColumn
                ? mapping.DateColumn
                : mapping.PeriodColumnName;
        }

        private static bool RequiresDataModel(IEnumerable<MeasureDefinition> measures)
        {
            return measures.Any(measure => measure != null && ContainsDistinctCount(measure.Expression));
        }

        private static bool ContainsDistinctCount(MeasureExpression expression)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    return aggregate.Function == AggregateFunction.DistinctCount;
                case FilteredAggregateMeasureExpression filtered:
                    return filtered.Function == AggregateFunction.DistinctCount;
                case WeightedAggregateMeasureExpression weighted:
                    return ContainsDistinctCount(weighted.Numerator) || ContainsDistinctCount(weighted.Denominator);
                case BinaryMeasureExpression binary:
                    return ContainsDistinctCount(binary.Left) || ContainsDistinctCount(binary.Right);
                case SafeDivideMeasureExpression divide:
                    return ContainsDistinctCount(divide.Numerator) || ContainsDistinctCount(divide.Denominator);
                case RatioMeasureExpression ratio:
                    return ContainsDistinctCount(ratio.Numerator) || ContainsDistinctCount(ratio.Denominator);
                case DifferenceMeasureExpression difference:
                    return ContainsDistinctCount(difference.Current) || ContainsDistinctCount(difference.Baseline);
                case ShareMeasureExpression share:
                    return ContainsDistinctCount(share.Part) || ContainsDistinctCount(share.Whole);
                default:
                    return false;
            }
        }

        private static string ManagedName(string ownershipId, string suffix)
        {
            var encoded = ownershipId
                .Replace("_", "_U")
                .Replace("-", "_D");
            return "ERB_" + encoded + "_" + suffix;
        }
    }
}
