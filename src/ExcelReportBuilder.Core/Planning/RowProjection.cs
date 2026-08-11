using System;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Planning
{
    public enum SourceLoadRoute
    {
        Worksheet,
        DataModel
    }

    public sealed class RowProjection
    {
        public const long ExcelWorksheetRowLimit = 1048576L;
        public const long MaximumWorksheetDataRows = ExcelWorksheetRowLimit - 1L;

        public long SourceRows { get; set; }

        public long ExpansionFactor { get; set; }

        public long ProjectedRows { get; set; }

        public SourceLoadRoute Route { get; set; }

        public bool WouldExceedWorksheet => ProjectedRows > MaximumWorksheetDataRows;

        public string Reason { get; set; } = string.Empty;
    }

    public static class RowProjectionCalculator
    {
        public static RowProjection Project(long sourceRows, PeriodDetectionResult detection)
        {
            if (detection == null)
            {
                throw new ArgumentNullException(nameof(detection));
            }

            long factor;
            switch (detection.Kind)
            {
                case PeriodLayoutKind.MonthHeaders:
                    factor = detection.HeaderMatches.Count;
                    break;
                case PeriodLayoutKind.MetricMonthHeaders:
                    factor = detection.HeaderMatches.Count;
                    break;
                default:
                    factor = 1L;
                    break;
            }

            return Build(sourceRows, factor);
        }

        public static RowProjection Project(long sourceRows, PeriodMappingSpec? mapping)
        {
            if (mapping == null || mapping.Kind == PeriodMappingKind.LongDateColumn)
            {
                return Build(sourceRows, 1L);
            }

            var factor = mapping.Columns.Count;
            return Build(sourceRows, factor);
        }

        private static RowProjection Build(long sourceRows, long factor)
        {
            if (sourceRows < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(sourceRows));
            }

            if (factor < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(factor));
            }

            long projected;
            try
            {
                projected = checked(sourceRows * factor);
            }
            catch (OverflowException)
            {
                projected = long.MaxValue;
            }

            var route = projected <= RowProjection.MaximumWorksheetDataRows
                ? SourceLoadRoute.Worksheet
                : SourceLoadRoute.DataModel;
            return new RowProjection
            {
                SourceRows = sourceRows,
                ExpansionFactor = factor,
                ProjectedRows = projected,
                Route = route,
                Reason = route == SourceLoadRoute.Worksheet
                    ? "The projected result fits on one worksheet with a header row."
                    : "The projected result exceeds the worksheet data-row limit and must use the Data Model."
            };
        }
    }
}
