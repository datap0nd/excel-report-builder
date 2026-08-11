using System;
using System.Collections.Generic;

namespace ExcelReportBuilder.Core.Measures
{
    public sealed class WeightedObservation
    {
        public decimal? Value { get; set; }

        public decimal? Weight { get; set; }
    }

    public sealed class WeightedAggregateResult
    {
        public decimal Numerator { get; set; }

        public decimal Denominator { get; set; }

        public decimal? Value { get; set; }

        public long IncludedRows { get; set; }

        public long ExcludedRows { get; set; }
    }

    /// <summary>
    /// Reference arithmetic for synthetic checks and independent reconciliation.
    /// Rows missing either value or weight are excluded from both numerator and
    /// denominator, preserving identical scopes.
    /// </summary>
    public static class WeightedAggregateCalculator
    {
        public static WeightedAggregateResult Calculate(
            IEnumerable<WeightedObservation> observations,
            ZeroDenominatorBehavior onZero = ZeroDenominatorBehavior.Blank)
        {
            if (observations == null)
            {
                throw new ArgumentNullException(nameof(observations));
            }

            var result = new WeightedAggregateResult();
            foreach (var observation in observations)
            {
                if (observation == null || !observation.Value.HasValue || !observation.Weight.HasValue)
                {
                    result.ExcludedRows++;
                    continue;
                }

                result.Numerator = checked(result.Numerator + observation.Value.Value * observation.Weight.Value);
                result.Denominator = checked(result.Denominator + observation.Weight.Value);
                result.IncludedRows++;
            }

            if (result.Denominator != 0m)
            {
                result.Value = result.Numerator / result.Denominator;
                return result;
            }

            switch (onZero)
            {
                case ZeroDenominatorBehavior.Blank:
                    result.Value = null;
                    break;
                case ZeroDenominatorBehavior.Zero:
                    result.Value = 0m;
                    break;
                case ZeroDenominatorBehavior.Error:
                    throw new DivideByZeroException("The weighted aggregate denominator is zero.");
                default:
                    throw new ArgumentOutOfRangeException(nameof(onZero));
            }

            return result;
        }
    }
}
