using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Excel.Validation
{
    /// <summary>
    /// Resolves only mathematically preserved additive lineage back to raw
    /// source columns. Derived or value-changing fields return no lineage and
    /// are validated from canonical data onward instead.
    /// </summary>
    public sealed class SourceTotalLineageResolver
    {
        public IReadOnlyList<string> Resolve(
            ReportSpecV1 specification,
            AggregateMeasureExpression aggregate)
        {
            if (specification == null) throw new ArgumentNullException(nameof(specification));
            if (aggregate == null) throw new ArgumentNullException(nameof(aggregate));
            if (aggregate.Function != AggregateFunction.Sum)
            {
                return Array.Empty<string>();
            }

            IReadOnlyList<string> candidates;
            if (specification.PeriodMapping != null &&
                specification.PeriodMapping.Kind != PeriodMappingKind.LongDateColumn &&
                string.Equals(
                    aggregate.Field,
                    specification.PeriodMapping.ValueColumnName,
                    StringComparison.OrdinalIgnoreCase))
            {
                candidates = specification.PeriodMapping.Columns
                    .Select(column => column.SourceColumn)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else
            {
                candidates = new[] { aggregate.Field };
            }

            var result = new List<string>();
            foreach (string candidate in candidates)
            {
                string? source = Trace(candidate, specification.Transforms);
                if (source == null)
                {
                    return Array.Empty<string>();
                }

                result.Add(source);
            }

            return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string? Trace(string field, IReadOnlyList<TransformStep> transforms)
        {
            string current = field;
            for (var index = transforms.Count - 1; index >= 0; index--)
            {
                TransformStep transform = transforms[index];
                switch (transform)
                {
                    case RenameColumnTransform rename when string.Equals(
                        current,
                        rename.To,
                        StringComparison.OrdinalIgnoreCase):
                        current = rename.From;
                        break;
                    case ChangeColumnTypeTransform change when Affects(change.Column, current):
                    case ReplaceValueTransform replace when Affects(replace.Column, current):
                    case NormalizeBlanksTransform blanks when blanks.Columns.Any(column => Affects(column, current)):
                    case NormalizeErrorsTransform errors when errors.Columns.Any(column => Affects(column, current)):
                    case FillDownTransform fill when fill.Columns.Any(column => Affects(column, current)):
                    case MapValuesTransform map when Affects(map.Column, current):
                        return null;
                    case AddArithmeticColumnTransform arithmetic when Affects(arithmetic.OutputColumn, current):
                        return null;
                    case DerivePeriodPartsTransform period when period.Columns.Any(column => Affects(column.OutputColumn, current)):
                        return null;
                }
            }

            return current;
        }

        private static bool Affects(string left, string right)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }
}
