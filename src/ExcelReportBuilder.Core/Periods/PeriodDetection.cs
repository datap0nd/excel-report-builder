using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Periods
{
    public enum PeriodLayoutKind
    {
        NotDetected,
        LongDateColumn,
        MonthHeaders,
        MetricMonthHeaders
    }

    public enum PeriodDetectionSeverity
    {
        Information,
        Warning,
        Error
    }

    public enum PeriodDetectionIssueCode
    {
        MissingReportingYear,
        MultipleDateColumns,
        MixedPeriodHeaderShapes,
        MixedPeriodGrains,
        DuplicatePeriodHeader,
        DuplicateMetricPeriodHeader,
        IncompleteMetricPeriodMatrix
    }

    public sealed class PeriodDetectionIssue
    {
        public PeriodDetectionIssueCode Code { get; set; }

        public PeriodDetectionSeverity Severity { get; set; }

        public string Message { get; set; } = string.Empty;
    }

    public sealed class PeriodHeaderMatch
    {
        public string SourceColumn { get; set; } = string.Empty;

        public int Month { get; set; }

        public int? Year { get; set; }

        public string? Metric { get; set; }

        public PeriodGrain Grain { get; set; } = PeriodGrain.Month;

        public DateTime? CanonicalPeriod { get; set; }
    }

    public sealed class PeriodDetectionResult
    {
        public PeriodLayoutKind Kind { get; set; }

        public string? DateColumn { get; set; }

        public List<string> CandidateDateColumns { get; set; } = new List<string>();

        public List<string> KeyColumns { get; set; } = new List<string>();

        public List<PeriodHeaderMatch> HeaderMatches { get; set; } = new List<PeriodHeaderMatch>();

        public List<PeriodDetectionIssue> Issues { get; set; } = new List<PeriodDetectionIssue>();

        public bool RequiresReportingYear { get; set; }

        public int? ReportingYear { get; set; }

        public PeriodGrain? Grain { get; set; }

        public bool IsAmbiguous => Issues.Any(issue => issue.Severity == PeriodDetectionSeverity.Error);

        public PeriodMappingSpec ToPeriodMapping(string id = "periods")
        {
            if (Kind == PeriodLayoutKind.NotDetected || IsAmbiguous)
            {
                throw new InvalidOperationException("An unambiguous period layout is required.");
            }

            var mapping = new PeriodMappingSpec
            {
                Id = id,
                ReportingYear = ReportingYear,
                DateColumn = DateColumn,
                Grain = Grain,
                KeyColumns = new List<string>(KeyColumns)
            };

            switch (Kind)
            {
                case PeriodLayoutKind.LongDateColumn:
                    mapping.Kind = PeriodMappingKind.LongDateColumn;
                    break;
                case PeriodLayoutKind.MonthHeaders:
                    mapping.Kind = PeriodMappingKind.MonthHeaders;
                    break;
                case PeriodLayoutKind.MetricMonthHeaders:
                    mapping.Kind = PeriodMappingKind.MetricMonthHeaders;
                    break;
                default:
                    throw new InvalidOperationException("Unsupported period layout.");
            }

            foreach (var match in HeaderMatches)
            {
                mapping.Columns.Add(new PeriodColumnMapping
                {
                    SourceColumn = match.SourceColumn,
                    Month = match.Month,
                    Year = match.Year,
                    Metric = match.Metric
                });
            }

            return mapping;
        }
    }

    public static class PeriodDetector
    {
        private const double LongDateThreshold = 0.8d;

        public static PeriodDetectionResult Detect(SourceProfile source, int? reportingYear = null)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            ValidateReportingYear(reportingYear);
            var parsedHeaders = new List<PeriodHeaderMatch>();
            foreach (var column in source.Columns)
            {
                PeriodHeaderMatch match;
                if (TryParseHeader(column.Name, reportingYear, out match))
                {
                    parsedHeaders.Add(match);
                }
            }

            if (IsWideCandidate(parsedHeaders))
            {
                return BuildWideResult(source, parsedHeaders, reportingYear);
            }

            var dateCandidateProfiles = source.Columns
                .Where(column => column.NonBlankCount > 0 && column.PeriodLikeRatio >= LongDateThreshold)
                .ToList();
            var dateCandidates = dateCandidateProfiles.Select(column => column.Name).ToList();

            if (dateCandidates.Count == 1)
            {
                var column = dateCandidateProfiles[0];
                var longResult = new PeriodDetectionResult
                {
                    Kind = PeriodLayoutKind.LongDateColumn,
                    DateColumn = dateCandidates[0],
                    CandidateDateColumns = dateCandidates,
                    KeyColumns = source.Columns
                        .Where(column => !string.Equals(column.Name, dateCandidates[0], StringComparison.OrdinalIgnoreCase))
                        .Select(column => column.Name)
                        .ToList(),
                    ReportingYear = reportingYear,
                    Grain = ResolveGrain(column)
                };

                if (CountObservedGrains(column) > 1)
                {
                    longResult.Issues.Add(new PeriodDetectionIssue
                    {
                        Code = PeriodDetectionIssueCode.MixedPeriodGrains,
                        Severity = PeriodDetectionSeverity.Error,
                        Message = "The period column mixes day, month, or quarter grains. Normalize it to one grain explicitly."
                    });
                }

                if (column.PeriodLikeWithoutYearCount > 0 && !reportingYear.HasValue)
                {
                    longResult.RequiresReportingYear = true;
                    longResult.Issues.Add(new PeriodDetectionIssue
                    {
                        Code = PeriodDetectionIssueCode.MissingReportingYear,
                        Severity = PeriodDetectionSeverity.Error,
                        Message = "One or more period values have no year. A reporting year must be supplied; it will not be inferred."
                    });
                }

                return longResult;
            }

            var result = new PeriodDetectionResult
            {
                Kind = dateCandidates.Count > 0 ? PeriodLayoutKind.LongDateColumn : PeriodLayoutKind.NotDetected,
                CandidateDateColumns = dateCandidates,
                ReportingYear = reportingYear
            };

            if (dateCandidates.Count > 1)
            {
                result.Issues.Add(new PeriodDetectionIssue
                {
                    Code = PeriodDetectionIssueCode.MultipleDateColumns,
                    Severity = PeriodDetectionSeverity.Error,
                    Message = "More than one date-like column was found. Select the reporting date column explicitly."
                });
            }

            return result;
        }

        private static PeriodDetectionResult BuildWideResult(
            SourceProfile source,
            List<PeriodHeaderMatch> matches,
            int? reportingYear)
        {
            var hasPure = matches.Any(match => string.IsNullOrWhiteSpace(match.Metric));
            var hasMetric = matches.Any(match => !string.IsNullOrWhiteSpace(match.Metric));
            var result = new PeriodDetectionResult
            {
                Kind = hasMetric ? PeriodLayoutKind.MetricMonthHeaders : PeriodLayoutKind.MonthHeaders,
                HeaderMatches = matches,
                KeyColumns = source.Columns
                    .Where(column => !matches.Any(match => string.Equals(match.SourceColumn, column.Name, StringComparison.OrdinalIgnoreCase)))
                    .Select(column => column.Name)
                    .ToList(),
                ReportingYear = reportingYear,
                Grain = matches.Select(match => match.Grain).Distinct().Count() == 1
                    ? matches[0].Grain
                    : (PeriodGrain?)null
            };

            if (hasPure && hasMetric)
            {
                result.Issues.Add(new PeriodDetectionIssue
                {
                    Code = PeriodDetectionIssueCode.MixedPeriodHeaderShapes,
                    Severity = PeriodDetectionSeverity.Error,
                    Message = "Month-only and metric-month headers are mixed. Map the period columns explicitly."
                });
                return result;
            }

            if (matches.Select(match => match.Grain).Distinct().Count() > 1)
            {
                result.Issues.Add(new PeriodDetectionIssue
                {
                    Code = PeriodDetectionIssueCode.MixedPeriodGrains,
                    Severity = PeriodDetectionSeverity.Error,
                    Message = "Month and quarter headers are mixed. Map one period grain at a time."
                });
                return result;
            }

            var missingYear = matches.Any(match => !match.Year.HasValue);
            if (missingYear && !reportingYear.HasValue)
            {
                result.RequiresReportingYear = true;
                result.Issues.Add(new PeriodDetectionIssue
                {
                    Code = PeriodDetectionIssueCode.MissingReportingYear,
                    Severity = PeriodDetectionSeverity.Error,
                    Message = "One or more period headers have no year. A reporting year must be supplied; it will not be inferred."
                });
            }

            if (hasMetric)
            {
                ValidateMetricMatrix(result, reportingYear);
            }
            else
            {
                var duplicate = matches
                    .GroupBy(match => PeriodKey(match, reportingYear), StringComparer.Ordinal)
                    .FirstOrDefault(group => group.Count() > 1);
                if (duplicate != null)
                {
                    result.Issues.Add(new PeriodDetectionIssue
                    {
                        Code = PeriodDetectionIssueCode.DuplicatePeriodHeader,
                        Severity = PeriodDetectionSeverity.Error,
                        Message = "More than one source column maps to period " + duplicate.Key + "."
                    });
                }
            }

            return result;
        }

        private static void ValidateMetricMatrix(PeriodDetectionResult result, int? reportingYear)
        {
            var duplicate = result.HeaderMatches
                .GroupBy(
                    match => (match.Metric ?? string.Empty).ToUpperInvariant() + "|" + PeriodKey(match, reportingYear),
                    StringComparer.Ordinal)
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicate != null)
            {
                result.Issues.Add(new PeriodDetectionIssue
                {
                    Code = PeriodDetectionIssueCode.DuplicateMetricPeriodHeader,
                    Severity = PeriodDetectionSeverity.Error,
                    Message = "More than one source column maps to the same metric and period."
                });
            }

            var allPeriods = new HashSet<string>(
                result.HeaderMatches.Select(match => PeriodKey(match, reportingYear)),
                StringComparer.Ordinal);
            foreach (var metricGroup in result.HeaderMatches.GroupBy(
                match => match.Metric ?? string.Empty,
                StringComparer.OrdinalIgnoreCase))
            {
                var metricPeriods = new HashSet<string>(
                    metricGroup.Select(match => PeriodKey(match, reportingYear)),
                    StringComparer.Ordinal);
                if (!metricPeriods.SetEquals(allPeriods))
                {
                    result.Issues.Add(new PeriodDetectionIssue
                    {
                        Code = PeriodDetectionIssueCode.IncompleteMetricPeriodMatrix,
                        Severity = PeriodDetectionSeverity.Error,
                        Message = "Metric '" + metricGroup.Key + "' does not contain the same periods as the other metrics."
                    });
                }
            }
        }

        private static bool IsWideCandidate(IReadOnlyCollection<PeriodHeaderMatch> matches)
        {
            if (matches.Count >= 2)
            {
                return true;
            }

            return matches.Count == 1 && string.IsNullOrWhiteSpace(matches.First().Metric);
        }

        private static bool TryParseHeader(string header, int? reportingYear, out PeriodHeaderMatch result)
        {
            result = new PeriodHeaderMatch();
            if (string.IsNullOrWhiteSpace(header))
            {
                return false;
            }

            var value = header.Trim();
            ParsedPeriodToken parsed;
            if (!PeriodTextParser.TryFindToken(value, reportingYear, out parsed)
                || PeriodTextParser.ContainsAnotherToken(value, parsed))
            {
                return false;
            }

            var metric = CleanMetric(value.Remove(parsed.Index, parsed.Length));
            result = new PeriodHeaderMatch
            {
                SourceColumn = header,
                Month = parsed.Month,
                Year = parsed.Year,
                Metric = string.IsNullOrWhiteSpace(metric) ? null : metric,
                Grain = parsed.Grain,
                CanonicalPeriod = parsed.CanonicalPeriod
            };
            return true;
        }

        private static string CleanMetric(string value)
        {
            return Regex.Replace(value, @"^[\s_\-/.|:()\[\]]+|[\s_\-/.|:()\[\]]+$", string.Empty)
                .Trim();
        }

        private static string PeriodKey(PeriodHeaderMatch match, int? reportingYear)
        {
            var effectiveYear = match.Year ?? reportingYear;
            return match.Grain.ToString().ToUpperInvariant()
                + "|"
                + (effectiveYear.HasValue
                    ? effectiveYear.Value.ToString("0000", CultureInfo.InvariantCulture)
                    : "????")
                + "-"
                + match.Month.ToString("00", CultureInfo.InvariantCulture);
        }

        private static PeriodGrain? ResolveGrain(SourceColumnProfile column)
        {
            if (CountObservedGrains(column) != 1)
            {
                return null;
            }

            if (column.QuarterGrainCount > 0)
            {
                return PeriodGrain.Quarter;
            }

            return column.MonthGrainCount > 0 ? PeriodGrain.Month : PeriodGrain.Day;
        }

        private static int CountObservedGrains(SourceColumnProfile column)
        {
            var count = 0;
            if (column.DayGrainCount > 0) count++;
            if (column.MonthGrainCount > 0) count++;
            if (column.QuarterGrainCount > 0) count++;
            return count;
        }

        private static void ValidateReportingYear(int? reportingYear)
        {
            if (reportingYear.HasValue && (reportingYear.Value < 1900 || reportingYear.Value > 9999))
            {
                throw new ArgumentOutOfRangeException(nameof(reportingYear), "The reporting year must be between 1900 and 9999.");
            }
        }
    }
}
