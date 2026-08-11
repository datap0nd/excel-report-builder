using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Periods
{
    internal sealed class ParsedPeriodToken
    {
        public int Index { get; set; }

        public int Length { get; set; }

        public int Month { get; set; }

        public int? Year { get; set; }

        public PeriodGrain Grain { get; set; }

        public DateTime? CanonicalPeriod { get; set; }

        public bool RequiresReportingYear => !Year.HasValue;
    }

    /// <summary>
    /// Parses the deliberately bounded period vocabulary shared by source
    /// profiling and wide-header detection. It does not use the machine's
    /// current culture, current year, or configurable two-digit-year cutoff.
    /// </summary>
    internal static class PeriodTextParser
    {
        // This fixed window matches the default Windows and Excel convention:
        // 00-29 are 2000-2029 and 30-99 are 1930-1999. Using an explicit
        // cutoff keeps saved setups deterministic across managed desktops.
        internal const int TwoDigitYearMaximum = 2029;

        private const string MonthNamePattern =
            @"jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?";

        private const string FourDigitYearPattern = @"(?:19|20)\d{2}";

        private const string TwoOrFourDigitYearPattern = @"(?:(?:19|20)\d{2}|\d{2})";

        private static readonly Regex QuarterThenYearRegex = Create(
            @"(?<![A-Za-z0-9])Q\s*(?<quarter>[1-4])[\s._/-]+(?<year>" + TwoOrFourDigitYearPattern + @")(?![A-Za-z0-9])");

        private static readonly Regex YearThenQuarterRegex = Create(
            @"(?<![A-Za-z0-9])(?<year>" + TwoOrFourDigitYearPattern + @")[\s._/-]+Q\s*(?<quarter>[1-4])(?![A-Za-z0-9])");

        private static readonly Regex CompactYearMonthRegex = Create(
            @"(?<!\d)(?<year>" + FourDigitYearPattern + @")(?<month>0[1-9]|1[0-2])(?!\d)");

        private static readonly Regex NumericYearMonthRegex = Create(
            @"(?<!\d)(?<year>" + FourDigitYearPattern + @")[-_/ ](?<month>0?[1-9]|1[0-2])(?!\d)");

        private static readonly Regex NumericMonthYearRegex = Create(
            @"(?<!\d)(?<month>0?[1-9]|1[0-2])[-_/ ](?<year>" + FourDigitYearPattern + @")(?!\d)");

        private static readonly Regex MonthThenYearRegex = Create(
            @"(?<![A-Za-z0-9])(?<month>" + MonthNamePattern + @")[\s._/-]+(?<year>" + TwoOrFourDigitYearPattern + @")(?![A-Za-z0-9])");

        private static readonly Regex YearThenMonthRegex = Create(
            @"(?<![A-Za-z0-9])(?<year>" + TwoOrFourDigitYearPattern + @")[\s._/-]+(?<month>" + MonthNamePattern + @")(?![A-Za-z0-9])");

        private static readonly Regex StandaloneQuarterRegex = Create(
            @"(?<![A-Za-z0-9])Q\s*(?<quarter>[1-4])(?![A-Za-z0-9])");

        private static readonly Regex StandaloneMonthRegex = Create(
            @"(?<![A-Za-z])(?<month>" + MonthNamePattern + @")(?![A-Za-z])");

        private static readonly IReadOnlyList<Regex> OrderedPatterns = new[]
        {
            QuarterThenYearRegex,
            YearThenQuarterRegex,
            CompactYearMonthRegex,
            NumericYearMonthRegex,
            NumericMonthYearRegex,
            MonthThenYearRegex,
            YearThenMonthRegex,
            StandaloneQuarterRegex,
            StandaloneMonthRegex
        };

        public static bool TryParseWholeValue(object value, out ParsedPeriodToken token)
        {
            return TryParseWholeValue(value, null, out token);
        }

        public static bool TryParseWholeValue(
            object value,
            int? reportingYear,
            out ParsedPeriodToken token)
        {
            token = new ParsedPeriodToken();
            string text;
            if (value is string stringValue)
            {
                text = stringValue.Trim();
                if (HasMalformedSeparatorSequence(text))
                {
                    return false;
                }
            }
            else if (TryFormatCompactNumericValue(value, out text))
            {
                // Compact numeric periods are common in raw extracts. Format
                // them without culture-dependent separators before parsing.
            }
            else
            {
                return false;
            }

            if (text.Length == 0 || !TryFindToken(text, reportingYear, out token))
            {
                return false;
            }

            return token.Index == 0 && token.Length == text.Length;
        }

        private static bool HasMalformedSeparatorSequence(string text)
        {
            string compact = Regex.Replace(text.Trim(), @"\s+", " ").ToUpperInvariant();
            compact = compact.Replace("Q ", "Q");
            const string separators = " -_./";
            for (var index = 1; index < compact.Length; index++)
            {
                if (separators.IndexOf(compact[index - 1]) >= 0
                    && separators.IndexOf(compact[index]) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool TryFindToken(string text, int? reportingYear, out ParsedPeriodToken token)
        {
            token = new ParsedPeriodToken();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            Match? selected = null;
            Regex? selectedPattern = null;
            foreach (var pattern in OrderedPatterns)
            {
                var candidate = pattern.Match(text);
                if (!candidate.Success)
                {
                    continue;
                }

                if (selected == null || candidate.Index < selected.Index
                    || (candidate.Index == selected.Index && candidate.Length > selected.Length))
                {
                    selected = candidate;
                    selectedPattern = pattern;
                }
            }

            if (selected == null || selectedPattern == null)
            {
                return false;
            }

            var quarterGroup = selected.Groups["quarter"];
            var month = quarterGroup.Success
                ? ((int.Parse(quarterGroup.Value, CultureInfo.InvariantCulture) - 1) * 3) + 1
                : ParseMonth(selected.Groups["month"].Value);
            var grain = quarterGroup.Success ? PeriodGrain.Quarter : PeriodGrain.Month;
            var yearGroup = selected.Groups["year"];
            int? parsedYear = yearGroup.Success ? ParseYear(yearGroup.Value) : (int?)null;
            var effectiveYear = parsedYear ?? reportingYear;
            token = new ParsedPeriodToken
            {
                Index = selected.Index,
                Length = selected.Length,
                Month = month,
                Year = parsedYear,
                Grain = grain,
                CanonicalPeriod = effectiveYear.HasValue
                    ? new DateTime(effectiveYear.Value, month, 1)
                    : (DateTime?)null
            };
            return true;
        }

        public static bool ContainsAnotherToken(string text, ParsedPeriodToken selected)
        {
            var remainder = text.Remove(selected.Index, selected.Length);
            ParsedPeriodToken ignored;
            return TryFindToken(remainder, null, out ignored);
        }

        private static Regex Create(string pattern)
        {
            return new Regex(
                pattern,
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        }

        private static int ParseMonth(string month)
        {
            if (string.IsNullOrWhiteSpace(month))
            {
                throw new InvalidOperationException("A month token is required.");
            }

            int numericMonth;
            if (int.TryParse(month, NumberStyles.None, CultureInfo.InvariantCulture, out numericMonth))
            {
                if (numericMonth < 1 || numericMonth > 12)
                {
                    throw new InvalidOperationException("Unsupported numeric month.");
                }

                return numericMonth;
            }

            switch (month.Substring(0, 3).ToUpperInvariant())
            {
                case "JAN": return 1;
                case "FEB": return 2;
                case "MAR": return 3;
                case "APR": return 4;
                case "MAY": return 5;
                case "JUN": return 6;
                case "JUL": return 7;
                case "AUG": return 8;
                case "SEP": return 9;
                case "OCT": return 10;
                case "NOV": return 11;
                case "DEC": return 12;
                default: throw new InvalidOperationException("Unsupported month name.");
            }
        }

        private static int ParseYear(string year)
        {
            var numericYear = int.Parse(year, CultureInfo.InvariantCulture);
            if (year.Length == 4)
            {
                return numericYear;
            }

            var twoDigitCutoff = TwoDigitYearMaximum % 100;
            var century = TwoDigitYearMaximum - twoDigitCutoff;
            return numericYear <= twoDigitCutoff
                ? century + numericYear
                : century - 100 + numericYear;
        }

        private static bool TryFormatCompactNumericValue(object value, out string text)
        {
            text = string.Empty;
            if (value is bool || value is char || value is DateTime || value is DateTimeOffset)
            {
                return false;
            }

            if (!(value is byte || value is sbyte || value is short || value is ushort
                || value is int || value is uint || value is long || value is ulong
                || value is decimal || value is double || value is float))
            {
                return false;
            }

            try
            {
                var numeric = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                if (decimal.Truncate(numeric) != numeric || numeric < 190001m || numeric > 209912m)
                {
                    return false;
                }

                text = numeric.ToString("000000", CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception exception) when (exception is OverflowException || exception is FormatException)
            {
                return false;
            }
        }
    }
}
