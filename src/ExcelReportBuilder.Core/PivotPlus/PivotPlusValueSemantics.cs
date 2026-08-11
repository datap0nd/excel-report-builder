using System;

namespace ExcelReportBuilder.Core.PivotPlus
{
    /// <summary>
    /// Deterministic display identity for one native Values instance. The
    /// validator and Excel executor share this logic so caption collisions are
    /// rejected before mutation.
    /// </summary>
    public static class PivotPlusValueSemantics
    {
        public static string ResolveCaption(
            PivotFieldDescriptor field,
            PivotFieldPlacement placement)
        {
            if (field == null) throw new ArgumentNullException(nameof(field));
            if (placement == null) throw new ArgumentNullException(nameof(placement));

            if (!string.IsNullOrWhiteSpace(placement.Caption))
            {
                return placement.Caption!;
            }

            string fieldCaption = string.IsNullOrWhiteSpace(field.Caption)
                ? field.Name
                : field.Caption!;
            if (placement.Area != PivotFieldArea.Values || field.IsMeasure)
            {
                return fieldCaption;
            }

            string aggregate = placement.Aggregation.HasValue
                ? AggregationCaption(placement.Aggregation.Value)
                : "Value";
            return aggregate + " of " + fieldCaption;
        }

        private static string AggregationCaption(PivotAggregationFunction function)
        {
            switch (function)
            {
                case PivotAggregationFunction.Minimum: return "Min";
                case PivotAggregationFunction.Maximum: return "Max";
                case PivotAggregationFunction.CountNumbers: return "Count Numbers";
                case PivotAggregationFunction.StandardDeviation: return "StdDev";
                case PivotAggregationFunction.StandardDeviationPopulation: return "StdDevP";
                case PivotAggregationFunction.Variance: return "Var";
                case PivotAggregationFunction.VariancePopulation: return "VarP";
                case PivotAggregationFunction.DistinctCount: return "Distinct Count";
                default: return function.ToString();
            }
        }
    }
}
