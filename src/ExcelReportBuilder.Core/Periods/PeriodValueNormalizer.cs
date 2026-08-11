using System;
using System.Globalization;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Periods
{
    /// <summary>
    /// Deterministic reference normalizer for values in a long period column.
    /// It accepts only the same bounded, culture-independent vocabulary used
    /// by source profiling and header detection.
    /// </summary>
    public static class PeriodValueNormalizer
    {
        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "d-MMM-yyyy",
            "dd-MMM-yyyy",
            "MMM d yyyy",
            "MMMM d yyyy"
        };

        private static readonly string[] DateTimeFormats =
        {
            "yyyy-MM-dd HH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss",
            "yyyy-MM-ddTHH:mm:ss.FFFFFFF",
            "yyyy-MM-ddTHH:mm:ssK",
            "yyyy-MM-ddTHH:mm:ss.FFFFFFFK"
        };

        public static DateTime Normalize(
            object value,
            int? reportingYear = null,
            PeriodGrain? expectedGrain = null)
        {
            if (value == null || value == DBNull.Value)
            {
                throw new ArgumentNullException(nameof(value));
            }

            ValidateReportingYear(reportingYear);
            ValidateGrain(expectedGrain);

            DateTime temporal;
            if (TryGetDate(value, out temporal))
            {
                var grain = expectedGrain ?? InferDateGrain(temporal);
                return Canonicalize(temporal, grain);
            }

            ParsedPeriodToken parsed;
            if (!PeriodTextParser.TryParseWholeValue(value, reportingYear, out parsed))
            {
                throw new ArgumentException(
                    "The value is not a supported unambiguous period representation.",
                    nameof(value));
            }

            if (!parsed.CanonicalPeriod.HasValue)
            {
                throw new InvalidOperationException(
                    "A period value without a year requires an explicit reporting year; it cannot be inferred.");
            }

            if (expectedGrain.HasValue && expectedGrain.Value != parsed.Grain)
            {
                throw new InvalidOperationException(
                    "The period value does not match the expected "
                    + expectedGrain.Value.ToString().ToLowerInvariant() + " grain.");
            }

            return Canonicalize(parsed.CanonicalPeriod.Value, parsed.Grain);
        }

        private static bool TryGetDate(object value, out DateTime result)
        {
            if (value is DateTime dateTime)
            {
                result = dateTime;
                return true;
            }

            if (value is DateTimeOffset dateTimeOffset)
            {
                result = dateTimeOffset.DateTime;
                return true;
            }

            if (value is string text)
            {
                if (DateTime.TryParseExact(
                    text.Trim(),
                    DateFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out result))
                {
                    return true;
                }

                if (DateTime.TryParseExact(
                    text.Trim(),
                    DateTimeFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                    out result))
                {
                    return true;
                }
            }

            result = default(DateTime);
            return false;
        }

        private static PeriodGrain InferDateGrain(DateTime value)
        {
            return value.Day == 1 && value.TimeOfDay == TimeSpan.Zero
                ? PeriodGrain.Month
                : PeriodGrain.Day;
        }

        private static DateTime Canonicalize(DateTime value, PeriodGrain grain)
        {
            switch (grain)
            {
                case PeriodGrain.Day:
                    return value.Date;
                case PeriodGrain.Month:
                    return new DateTime(value.Year, value.Month, 1);
                case PeriodGrain.Quarter:
                    var firstMonth = ((value.Month - 1) / 3 * 3) + 1;
                    return new DateTime(value.Year, firstMonth, 1);
                default:
                    throw new ArgumentOutOfRangeException(nameof(grain));
            }
        }

        private static void ValidateReportingYear(int? reportingYear)
        {
            if (reportingYear.HasValue && (reportingYear.Value < 1900 || reportingYear.Value > 9999))
            {
                throw new ArgumentOutOfRangeException(
                    nameof(reportingYear),
                    "The reporting year must be between 1900 and 9999.");
            }
        }

        private static void ValidateGrain(PeriodGrain? grain)
        {
            if (grain.HasValue && !Enum.IsDefined(typeof(PeriodGrain), grain.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(grain));
            }
        }
    }
}
