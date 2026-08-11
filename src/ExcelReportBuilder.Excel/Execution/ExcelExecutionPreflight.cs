using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Execution
{
    /// <summary>
    /// Rejects an execution plan before the first workbook mutation when the
    /// installed executor cannot represent it faithfully. A visible blocker is
    /// safer than a partially correct report.
    /// </summary>
    public static class ExcelExecutionPreflight
    {
        public static void DemandSupported(ReportSpecV1 specification, ReportBuildPlan plan)
        {
            if (specification == null) throw new ArgumentNullException(nameof(specification));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (!string.Equals(
                    specification.SchemaVersion,
                    ReportSpecV1.CurrentSchemaVersion,
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException("The report specification version is not supported.");
            }

            if (plan.Source.Route != SourceLoadRoute.Worksheet &&
                plan.Source.Route != SourceLoadRoute.DataModel)
            {
                throw new NotSupportedException(
                    "The report plan contains an unsupported canonical-data route.");
            }

            var useDataModel = plan.Source.Route == SourceLoadRoute.DataModel;
            if (plan.Blocks.Any(block => block.Pivot.UseDataModel != useDataModel))
            {
                throw new NotSupportedException(
                    "Every managed PivotTable must use the same validated backend as the canonical data.");
            }

            DemandCheckMeasuresAreRendered(plan);

            foreach (var block in plan.Blocks)
            {
                DemandPresentationIsConsumed(block);

                var nativeOutput = block.OutputMode == ReportOutputMode.StandardMatrix ||
                                   block.OutputMode == ReportOutputMode.MetricStack;

                foreach (var field in block.Pivot.Rows.Concat(block.Pivot.Columns))
                {
                    if (nativeOutput && field.GroupBuckets.Count > 0)
                    {
                        throw new NotSupportedException(
                            "Layout grouping buckets require a dense output block.");
                    }

                    if (nativeOutput && field.TopN != null && field.TopN.IncludeOthers)
                    {
                        throw new NotSupportedException(
                            "An explicit Others member requires a dense output block.");
                    }

                    if (!nativeOutput && field.TopN != null)
                    {
                        var rankingValue = block.Pivot.Values
                            .Concat(block.Pivot.SupportingValues)
                            .FirstOrDefault(value => string.Equals(
                                value.MeasureId,
                                field.TopN.MeasureId,
                                StringComparison.OrdinalIgnoreCase));
                        if (rankingValue == null ||
                            rankingValue.AggregateComponents.Count != 1 ||
                            rankingValue.AggregateComponents[0].Filters.Count != 0 ||
                            !string.IsNullOrWhiteSpace(rankingValue.AggregateComponents[0].PeriodSliceId))
                        {
                            throw new NotSupportedException(
                                "Dense Top N requires one unsliced direct aggregate Value for deterministic ranking.");
                        }
                    }
                }

                foreach (var value in block.Pivot.Values.Concat(block.Pivot.SupportingValues))
                {
                    if (value.AggregateComponents.Any(component =>
                            component.Function == AggregateFunction.DistinctCount) &&
                        !block.Pivot.UseDataModel)
                    {
                        throw new NotSupportedException(
                            "Distinct count requires the Data Model-backed PivotTable executor.");
                    }

                    if (nativeOutput && value.RequiresPostAggregationCalculation)
                    {
                        throw new NotSupportedException(
                            "Calculated measures require a dense output block in this executor version.");
                    }

                    if (nativeOutput && (value.PeriodSliceIds.Count > 0 || HasPeriodSlice(value.Expression)))
                    {
                        throw new NotSupportedException(
                            "Period-sliced values require a dense output block in this executor version.");
                    }

                    ValidateFilters(value.Expression, nativeOutput);
                    if (!nativeOutput && block.Pivot.Values.Contains(value))
                    {
                        DemandIndependentPivotReadIsBounded(block, value);
                    }
                }
            }
        }

        private static void DemandIndependentPivotReadIsBounded(
            DenseReportBlockPlan block,
            PivotValuePlan value)
        {
            foreach (var component in value.AggregateComponents)
            {
                var filterPairs = block.Pivot.Rows.Count +
                                  block.Pivot.Columns.Count +
                                  component.Filters.Count +
                                  (!string.IsNullOrWhiteSpace(component.PeriodSliceId) ? 1 : 0) +
                                  (value.PeriodSliceIds.Count > 0 ? 1 : 0);
                if (filterPairs > ExcelReportExecutor.MaximumIndependentPivotFilterPairs)
                {
                    throw new NotSupportedException(
                        "A dense Value requires more PivotTable field filters than can be independently validated by this executor version.");
                }
            }
        }

        private static void DemandCheckMeasuresAreRendered(ReportBuildPlan plan)
        {
            var renderedMeasureIds = new HashSet<string>(
                plan.Blocks
                    .SelectMany(block => block.Pivot.Values)
                    .Select(value => value.MeasureId)
                    .Where(value => !string.IsNullOrWhiteSpace(value)),
                StringComparer.OrdinalIgnoreCase);

            foreach (var check in plan.Checks)
            {
                DemandCheckMeasureIsRendered(check.Id, check.MeasureId, renderedMeasureIds);
                DemandCheckMeasureIsRendered(check.Id, check.ComparedMeasureId, renderedMeasureIds);
            }
        }

        private static void DemandCheckMeasureIsRendered(
            string checkId,
            string? measureId,
            ISet<string> renderedMeasureIds)
        {
            if (string.IsNullOrWhiteSpace(measureId))
            {
                return;
            }

            if (!renderedMeasureIds.Contains(measureId!))
            {
                throw new NotSupportedException(
                    "Configured check '" + checkId + "' targets Value '" + measureId +
                    "', but that Value is not rendered by any report block. Add it to a block before building so the check has an independent output to inspect.");
            }
        }

        private static void DemandPresentationIsConsumed(DenseReportBlockPlan block)
        {
            var nativeOutput = block.OutputMode == ReportOutputMode.StandardMatrix ||
                               block.OutputMode == ReportOutputMode.MetricStack;
            if (nativeOutput && block.Presentation.Spacers.Count > 0)
            {
                throw new NotSupportedException(
                    "Explicit row or column spacers require a dense output block.");
            }

            if (block.Presentation.Spacers.Any(spacer =>
                    spacer.Axis == SpacerAxis.Column && spacer.Size > 255d))
            {
                throw new NotSupportedException(
                    "A column spacer exceeds Excel's maximum column width.");
            }

            foreach (var spacer in block.Presentation.Spacers)
            {
                var levels = spacer.Axis == SpacerAxis.Row
                    ? block.Pivot.Rows.Count
                    : block.Pivot.Columns.Count;
                if (levels == 0 || spacer.BeforeLevel >= levels)
                {
                    throw new NotSupportedException(
                        "A configured spacer references a hierarchy level that does not exist in this block.");
                }
            }

            if (block.Presentation.Options.RepeatRowLabels)
            {
                throw new NotSupportedException(
                    "Explicit repeated row labels are not supported by this executor version.");
            }

            if (nativeOutput && !string.IsNullOrWhiteSpace(block.Presentation.SubtotalStyleId))
            {
                throw new NotSupportedException(
                    "A custom subtotal style is not supported by this executor version.");
            }

            if (nativeOutput && block.Pivot.Rows.Concat(block.Pivot.Columns).Any(field =>
                    !string.IsNullOrWhiteSpace(field.Subtotals.StyleId)))
            {
                throw new NotSupportedException(
                    "Per-level subtotal styles are not supported by this executor version.");
            }

            if (nativeOutput &&
                (block.Pivot.GrandTotals.RowPlacement != TotalPlacement.AfterMembers ||
                 block.Pivot.GrandTotals.ColumnPlacement != TotalPlacement.AfterMembers))
            {
                throw new NotSupportedException(
                    "Grand totals before report members are not supported by this executor version.");
            }

            if (nativeOutput && !string.IsNullOrWhiteSpace(block.Pivot.GrandTotals.StyleId))
            {
                throw new NotSupportedException(
                    "The layout-level grand-total style is not supported. Use the report block's grand-total style instead.");
            }

            if (nativeOutput && !string.Equals(
                    block.Pivot.GrandTotals.ColumnLabel,
                    "Grand Total",
                    StringComparison.Ordinal))
            {
                throw new NotSupportedException(
                    "Custom column grand-total labels are not supported by this executor version.");
            }
        }

        private static void ValidateFilters(MeasureExpression expression, bool nativeOutput)
        {
            switch (expression)
            {
                case FilteredAggregateMeasureExpression filtered:
                    if (nativeOutput)
                    {
                        throw new NotSupportedException(
                            "Filtered measures require a dense output block in this executor version.");
                    }

                    foreach (var filter in filtered.Filters)
                    {
                        if (filter.Operator != MeasureFilterOperator.Equal &&
                            filter.Operator != MeasureFilterOperator.In &&
                            filter.Operator != MeasureFilterOperator.IsBlank)
                        {
                            throw new NotSupportedException(
                                "A filtered measure uses an operator that requires a dedicated filtered pivot.");
                        }
                    }

                    break;
                case WeightedAggregateMeasureExpression weighted:
                    ValidateFilters(weighted.Numerator, nativeOutput);
                    ValidateFilters(weighted.Denominator, nativeOutput);
                    break;
                case BinaryMeasureExpression binary:
                    ValidateFilters(binary.Left, nativeOutput);
                    ValidateFilters(binary.Right, nativeOutput);
                    break;
                case SafeDivideMeasureExpression divide:
                    ValidateFilters(divide.Numerator, nativeOutput);
                    ValidateFilters(divide.Denominator, nativeOutput);
                    break;
                case RatioMeasureExpression ratio:
                    ValidateFilters(ratio.Numerator, nativeOutput);
                    ValidateFilters(ratio.Denominator, nativeOutput);
                    break;
                case DifferenceMeasureExpression difference:
                    ValidateFilters(difference.Current, nativeOutput);
                    ValidateFilters(difference.Baseline, nativeOutput);
                    break;
                case ShareMeasureExpression share:
                    ValidateFilters(share.Part, nativeOutput);
                    ValidateFilters(share.Whole, nativeOutput);
                    break;
            }
        }

        private static bool HasPeriodSlice(MeasureExpression expression)
        {
            switch (expression)
            {
                case AggregateMeasureExpression aggregate:
                    return !string.IsNullOrWhiteSpace(aggregate.PeriodSliceId);
                case FilteredAggregateMeasureExpression filtered:
                    return !string.IsNullOrWhiteSpace(filtered.PeriodSliceId);
                case WeightedAggregateMeasureExpression weighted:
                    return HasPeriodSlice(weighted.Numerator) || HasPeriodSlice(weighted.Denominator);
                case BinaryMeasureExpression binary:
                    return HasPeriodSlice(binary.Left) || HasPeriodSlice(binary.Right);
                case SafeDivideMeasureExpression divide:
                    return HasPeriodSlice(divide.Numerator) || HasPeriodSlice(divide.Denominator);
                case RatioMeasureExpression ratio:
                    return HasPeriodSlice(ratio.Numerator) || HasPeriodSlice(ratio.Denominator);
                case DifferenceMeasureExpression difference:
                    return HasPeriodSlice(difference.Current) || HasPeriodSlice(difference.Baseline);
                case ShareMeasureExpression share:
                    return HasPeriodSlice(share.Part) || HasPeriodSlice(share.Whole);
                default:
                    return false;
            }
        }
    }
}
