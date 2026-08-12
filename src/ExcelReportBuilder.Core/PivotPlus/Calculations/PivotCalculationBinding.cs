using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelReportBuilder.Core.PivotPlus.Calculations
{
    internal enum PivotCalculationSemanticKind
    {
        Unknown,
        Numeric,
        Ratio,
        PercentagePoints
    }

    internal sealed class PivotBoundField
    {
        public PivotBoundField(
            PivotModelTableSchema table,
            PivotModelFieldSchema field)
        {
            Table = table;
            Field = field;
        }

        public PivotModelTableSchema Table { get; }

        public PivotModelFieldSchema Field { get; }
    }

    internal sealed class PivotCalculationModelIndex
    {
        private readonly Dictionary<string, PivotModelTableSchema> tables =
            new Dictionary<string, PivotModelTableSchema>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, PivotBoundField> fields =
            new Dictionary<string, PivotBoundField>(StringComparer.OrdinalIgnoreCase);

        public PivotCalculationModelIndex(PivotModelSchema schema)
        {
            foreach (PivotModelTableSchema? table in schema.Tables)
            {
                if (table == null || string.IsNullOrWhiteSpace(table.Id) || tables.ContainsKey(table.Id))
                {
                    continue;
                }

                tables.Add(table.Id, table);
                foreach (PivotModelFieldSchema? field in table.Fields)
                {
                    if (field == null || string.IsNullOrWhiteSpace(field.Id) || fields.ContainsKey(field.Id))
                    {
                        continue;
                    }

                    fields.Add(field.Id, new PivotBoundField(table, field));
                }
            }
        }

        public bool TryGetTable(string id, out PivotModelTableSchema table)
        {
            return tables.TryGetValue(id, out table!);
        }

        public bool TryGetField(string id, out PivotBoundField field)
        {
            return fields.TryGetValue(id, out field!);
        }

        public bool TryResolveValue(
            string fieldId,
            PivotFilterValue value,
            out PivotScalarValue scalar)
        {
            scalar = PivotScalarValue.Blank();
            if (!TryGetField(fieldId, out PivotBoundField field))
            {
                return false;
            }

            if (value.Kind == PivotFilterValueKind.Scalar && value.Scalar != null)
            {
                scalar = value.Scalar;
                return true;
            }

            if (value.Kind != PivotFilterValueKind.Member || string.IsNullOrWhiteSpace(value.MemberId))
            {
                return false;
            }

            PivotModelMember? member = field.Field.Members.FirstOrDefault(candidate =>
                candidate != null &&
                string.Equals(candidate.Id, value.MemberId, StringComparison.OrdinalIgnoreCase));
            if (member == null)
            {
                return false;
            }

            scalar = member.Value;
            return true;
        }
    }

    internal static class PivotCalculationCanonical
    {
        public static string ScalarKey(PivotScalarValue value)
        {
            switch (value.Kind)
            {
                case PivotScalarKind.Blank:
                    return "blank";
                case PivotScalarKind.Text:
                    return "text:" + LengthPrefix(value.TextValue ?? string.Empty);
                case PivotScalarKind.WholeNumber:
                    return "whole:" + value.WholeNumberValue.GetValueOrDefault()
                        .ToString(CultureInfo.InvariantCulture);
                case PivotScalarKind.DecimalNumber:
                    return "decimal:" + value.DecimalNumberValue.GetValueOrDefault()
                        .ToString("0.############################", CultureInfo.InvariantCulture);
                case PivotScalarKind.Boolean:
                    return value.BooleanValue.GetValueOrDefault() ? "boolean:true" : "boolean:false";
                case PivotScalarKind.Date:
                    return "date:" + value.TemporalValue.GetValueOrDefault()
                        .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                case PivotScalarKind.DateTime:
                    return "datetime:" + value.TemporalValue.GetValueOrDefault()
                        .ToString("yyyy-MM-ddTHH:mm:ss.fffffff", CultureInfo.InvariantCulture);
                default:
                    return "unknown:" + ((int)value.Kind).ToString(CultureInfo.InvariantCulture);
            }
        }

        public static string PeriodPointKey(PivotPeriodPoint point)
        {
            return ((int)point.Grain).ToString(CultureInfo.InvariantCulture) + ":" +
                   point.Year.ToString(CultureInfo.InvariantCulture) + ":" +
                   (point.Ordinal.HasValue
                       ? point.Ordinal.Value.ToString(CultureInfo.InvariantCulture)
                       : string.Empty) + ":" +
                   (point.Date.HasValue
                       ? point.Date.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                       : string.Empty);
        }

        public static string LengthPrefix(string value)
        {
            return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
        }
    }

    internal static class PivotPeriodRules
    {
        public static int GrainRank(PivotPeriodGrain grain)
        {
            switch (grain)
            {
                case PivotPeriodGrain.Year: return 0;
                case PivotPeriodGrain.Half: return 1;
                case PivotPeriodGrain.Quarter: return 2;
                case PivotPeriodGrain.Month: return 3;
                case PivotPeriodGrain.Date: return 4;
                default: return -1;
            }
        }

        public static bool IsWithin(PivotPeriodPoint candidate, PivotPeriodPoint requested)
        {
            if (candidate.Year != requested.Year)
            {
                return false;
            }

            switch (requested.Grain)
            {
                case PivotPeriodGrain.Year:
                    return true;
                case PivotPeriodGrain.Half:
                    return requested.Ordinal == Half(candidate);
                case PivotPeriodGrain.Quarter:
                    return requested.Ordinal == Quarter(candidate);
                case PivotPeriodGrain.Month:
                    return requested.Ordinal == Month(candidate);
                case PivotPeriodGrain.Date:
                    return candidate.Date.HasValue && requested.Date.HasValue &&
                           candidate.Date.Value.Date == requested.Date.Value.Date;
                default:
                    return false;
            }
        }

        public static int ExpectedBucketCount(
            PivotPeriodPoint requested,
            PivotPeriodGrain sourceGrain)
        {
            int requestRank = GrainRank(requested.Grain);
            int sourceRank = GrainRank(sourceGrain);
            if (requestRank < 0 || sourceRank < requestRank)
            {
                return 0;
            }

            if (requested.Grain == sourceGrain)
            {
                return 1;
            }

            switch (requested.Grain)
            {
                case PivotPeriodGrain.Year:
                    switch (sourceGrain)
                    {
                        case PivotPeriodGrain.Half: return 2;
                        case PivotPeriodGrain.Quarter: return 4;
                        case PivotPeriodGrain.Month: return 12;
                        case PivotPeriodGrain.Date:
                            return DateTime.IsLeapYear(requested.Year) ? 366 : 365;
                    }

                    break;
                case PivotPeriodGrain.Half:
                    switch (sourceGrain)
                    {
                        case PivotPeriodGrain.Quarter: return 2;
                        case PivotPeriodGrain.Month: return 6;
                        case PivotPeriodGrain.Date:
                        {
                            int startMonth = requested.Ordinal == 1 ? 1 : 7;
                            return DaysInMonths(requested.Year, startMonth, 6);
                        }
                    }

                    break;
                case PivotPeriodGrain.Quarter:
                    switch (sourceGrain)
                    {
                        case PivotPeriodGrain.Month: return 3;
                        case PivotPeriodGrain.Date:
                        {
                            int startMonth = ((requested.Ordinal ?? 1) - 1) * 3 + 1;
                            return DaysInMonths(requested.Year, startMonth, 3);
                        }
                    }

                    break;
                case PivotPeriodGrain.Month:
                    if (sourceGrain == PivotPeriodGrain.Date)
                    {
                        return DateTime.DaysInMonth(requested.Year, requested.Ordinal ?? 1);
                    }

                    break;
            }

            return 0;
        }

        public static IEnumerable<PivotPeriodCoverageMember> ResolveCoverage(
            PivotPeriodDefinition periods,
            PivotPeriodSlice slice)
        {
            return periods.Source.Coverage
                .Where(member => member != null && IsWithin(member.Point, slice.Point))
                .OrderBy(member => SortKey(member.Point), StringComparer.Ordinal);
        }

        public static bool TryGetDateRange(
            PivotPeriodPoint point,
            out DateTime start,
            out DateTime end)
        {
            start = default;
            end = default;
            if (point.Year < 1 || point.Year > 9999)
            {
                return false;
            }

            switch (point.Grain)
            {
                case PivotPeriodGrain.Year:
                    start = new DateTime(point.Year, 1, 1);
                    end = new DateTime(point.Year, 12, 31);
                    return true;
                case PivotPeriodGrain.Half:
                {
                    int month = point.Ordinal == 2 ? 7 : 1;
                    start = new DateTime(point.Year, month, 1);
                    end = start.AddMonths(6).AddDays(-1);
                    return true;
                }
                case PivotPeriodGrain.Quarter:
                {
                    int month = ((point.Ordinal ?? 1) - 1) * 3 + 1;
                    start = new DateTime(point.Year, month, 1);
                    end = start.AddMonths(3).AddDays(-1);
                    return true;
                }
                case PivotPeriodGrain.Month:
                {
                    int month = point.Ordinal ?? 1;
                    start = new DateTime(point.Year, month, 1);
                    end = start.AddMonths(1).AddDays(-1);
                    return true;
                }
                case PivotPeriodGrain.Date when point.Date.HasValue:
                    start = point.Date.Value.Date;
                    end = start;
                    return true;
                default:
                    return false;
            }
        }

        private static int? Half(PivotPeriodPoint point)
        {
            switch (point.Grain)
            {
                case PivotPeriodGrain.Half:
                    return point.Ordinal;
                case PivotPeriodGrain.Quarter:
                    return ((point.Ordinal ?? 1) - 1) / 2 + 1;
                case PivotPeriodGrain.Month:
                    return ((point.Ordinal ?? 1) - 1) / 6 + 1;
                case PivotPeriodGrain.Date:
                    return ((point.Date?.Month ?? 1) - 1) / 6 + 1;
                default:
                    return null;
            }
        }

        private static int? Quarter(PivotPeriodPoint point)
        {
            switch (point.Grain)
            {
                case PivotPeriodGrain.Quarter:
                    return point.Ordinal;
                case PivotPeriodGrain.Month:
                    return ((point.Ordinal ?? 1) - 1) / 3 + 1;
                case PivotPeriodGrain.Date:
                    return ((point.Date?.Month ?? 1) - 1) / 3 + 1;
                default:
                    return null;
            }
        }

        private static int? Month(PivotPeriodPoint point)
        {
            switch (point.Grain)
            {
                case PivotPeriodGrain.Month:
                    return point.Ordinal;
                case PivotPeriodGrain.Date:
                    return point.Date?.Month;
                default:
                    return null;
            }
        }

        private static int DaysInMonths(int year, int startMonth, int count)
        {
            var total = 0;
            for (var offset = 0; offset < count; offset++)
            {
                total += DateTime.DaysInMonth(year, startMonth + offset);
            }

            return total;
        }

        private static string SortKey(PivotPeriodPoint point)
        {
            int month;
            int day;
            switch (point.Grain)
            {
                case PivotPeriodGrain.Year:
                    month = 1;
                    day = 1;
                    break;
                case PivotPeriodGrain.Half:
                    month = point.Ordinal == 2 ? 7 : 1;
                    day = 1;
                    break;
                case PivotPeriodGrain.Quarter:
                    month = ((point.Ordinal ?? 1) - 1) * 3 + 1;
                    day = 1;
                    break;
                case PivotPeriodGrain.Month:
                    month = point.Ordinal ?? 1;
                    day = 1;
                    break;
                case PivotPeriodGrain.Date:
                    month = point.Date?.Month ?? 1;
                    day = point.Date?.Day ?? 1;
                    break;
                default:
                    month = 1;
                    day = 1;
                    break;
            }

            return point.Year.ToString("0000", CultureInfo.InvariantCulture) +
                   month.ToString("00", CultureInfo.InvariantCulture) +
                   day.ToString("00", CultureInfo.InvariantCulture);
        }
    }
}
