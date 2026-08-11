using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Periods
{
    public sealed class NormalizedPeriodValue
    {
        public Dictionary<string, object?> Keys { get; set; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        public DateTime Period { get; set; }

        public string? Metric { get; set; }

        public object? Value { get; set; }

        public string SourceColumn { get; set; } = string.Empty;
    }

    /// <summary>
    /// A small in-memory reference implementation used by previews and
    /// independent total-preservation checks. Production-sized sources use the
    /// equivalent M plan and are never materialized here.
    /// </summary>
    public static class WidePeriodNormalizer
    {
        public static IReadOnlyList<NormalizedPeriodValue> Normalize(
            IEnumerable<IReadOnlyDictionary<string, object?>> sourceRows,
            PeriodMappingSpec mapping)
        {
            if (sourceRows == null)
            {
                throw new ArgumentNullException(nameof(sourceRows));
            }

            if (mapping == null)
            {
                throw new ArgumentNullException(nameof(mapping));
            }

            if (mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                throw new ArgumentException("Wide normalization requires month or metric-month header mappings.", nameof(mapping));
            }

            if (mapping.Columns == null || mapping.Columns.Count == 0)
            {
                throw new ArgumentException("At least one mapped period column is required.", nameof(mapping));
            }

            var grain = mapping.Grain ?? PeriodGrain.Month;
            if (grain == PeriodGrain.Day)
            {
                throw new ArgumentException("Wide normalization requires month or quarter grain.", nameof(mapping));
            }

            if (grain == PeriodGrain.Quarter && mapping.Columns.Any(column =>
                column.Month != 1 && column.Month != 4 && column.Month != 7 && column.Month != 10))
            {
                throw new ArgumentException(
                    "Quarter mappings must use the first month of each quarter.",
                    nameof(mapping));
            }

            var output = new List<NormalizedPeriodValue>();
            foreach (var sourceRow in sourceRows)
            {
                if (sourceRow == null)
                {
                    throw new ArgumentException("A source row cannot be null.", nameof(sourceRows));
                }

                var keys = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (var keyColumn in mapping.KeyColumns)
                {
                    keys.Add(keyColumn, GetValue(sourceRow, keyColumn));
                }

                foreach (var periodColumn in mapping.Columns)
                {
                    var year = periodColumn.Year ?? mapping.ReportingYear;
                    if (!year.HasValue)
                    {
                        throw new InvalidOperationException("A reporting year is required; it cannot be inferred.");
                    }

                    output.Add(new NormalizedPeriodValue
                    {
                        Keys = new Dictionary<string, object?>(keys, StringComparer.OrdinalIgnoreCase),
                        Period = new DateTime(year.Value, periodColumn.Month, 1),
                        Metric = periodColumn.Metric,
                        Value = GetValue(sourceRow, periodColumn.SourceColumn),
                        SourceColumn = periodColumn.SourceColumn
                    });
                }
            }

            return output;
        }

        private static object? GetValue(IReadOnlyDictionary<string, object?> row, string column)
        {
            object? value;
            if (row.TryGetValue(column, out value))
            {
                return value;
            }

            var match = row.FirstOrDefault(pair => string.Equals(pair.Key, column, StringComparison.OrdinalIgnoreCase));
            if (match.Key != null)
            {
                return match.Value;
            }

            throw new KeyNotFoundException("The source row does not contain mapped column '" + column + "'.");
        }
    }
}
