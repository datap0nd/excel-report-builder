using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Periods
{
    public sealed class ResolvedPeriodSlice
    {
        public string Id { get; set; } = string.Empty;

        public string Label { get; set; } = string.Empty;

        public PeriodSliceKind Kind { get; set; }

        public DateTime StartInclusive { get; set; }

        public DateTime EndInclusive { get; set; }

        public string? BasedOnSliceId { get; set; }
    }

    public sealed class PeriodSliceResolutionException : InvalidOperationException
    {
        public PeriodSliceResolutionException(string code, string message)
            : base(message)
        {
            Code = code;
        }

        public string Code { get; }
    }

    /// <summary>
    /// Resolves every relative period slice into explicit inclusive dates.
    /// Hosts must consume these resolved bounds and must not reinterpret
    /// Current from the available source data.
    /// </summary>
    public static class PeriodSliceResolver
    {
        public static IReadOnlyList<ResolvedPeriodSlice> Resolve(IEnumerable<PeriodSliceSpec> slices)
        {
            if (slices == null)
            {
                throw new ArgumentNullException(nameof(slices));
            }

            var materialized = slices.ToList();
            if (materialized.Any(slice => slice == null))
            {
                throw new PeriodSliceResolutionException("PERIOD_SLICE_REQUIRED", "A period slice cannot be null.");
            }

            var byId = new Dictionary<string, PeriodSliceSpec>(StringComparer.OrdinalIgnoreCase);
            foreach (var slice in materialized)
            {
                if (string.IsNullOrWhiteSpace(slice.Id) || byId.ContainsKey(slice.Id))
                {
                    throw new PeriodSliceResolutionException("PERIOD_SLICE_ID_INVALID", "Period-slice IDs must be nonblank and unique.");
                }

                byId.Add(slice.Id, slice);
            }

            var resolved = new Dictionary<string, ResolvedPeriodSlice>(StringComparer.OrdinalIgnoreCase);
            var resolving = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var slice in materialized)
            {
                try
                {
                    ResolveOne(slice, byId, resolved, resolving);
                }
                catch (Exception exception) when (exception is ArgumentOutOfRangeException || exception is OverflowException)
                {
                    throw new PeriodSliceResolutionException(
                        "PERIOD_SLICE_RANGE_UNSUPPORTED",
                        "The resolved period slice falls outside the supported date range.");
                }
            }

            return materialized.Select(slice => resolved[slice.Id]).ToList();
        }

        private static ResolvedPeriodSlice ResolveOne(
            PeriodSliceSpec slice,
            Dictionary<string, PeriodSliceSpec> byId,
            Dictionary<string, ResolvedPeriodSlice> resolved,
            HashSet<string> resolving)
        {
            ResolvedPeriodSlice existing;
            if (resolved.TryGetValue(slice.Id, out existing))
            {
                return existing;
            }

            if (!resolving.Add(slice.Id))
            {
                throw new PeriodSliceResolutionException(
                    "PERIOD_SLICE_REFERENCE_CYCLE",
                    "The period-slice reference graph contains a cycle.");
            }

            DateTime start;
            DateTime end;
            if (slice.Kind == PeriodSliceKind.Current || slice.Kind == PeriodSliceKind.Selected)
            {
                if (!slice.SelectedStart.HasValue || !slice.SelectedEnd.HasValue)
                {
                    throw new PeriodSliceResolutionException(
                        "ABSOLUTE_SLICE_DATES_REQUIRED",
                        "Current and selected slices require explicit dates.");
                }

                start = slice.SelectedStart.Value;
                end = slice.SelectedEnd.Value;
                if (start.TimeOfDay != TimeSpan.Zero || end.TimeOfDay != TimeSpan.Zero || start > end)
                {
                    throw new PeriodSliceResolutionException(
                        "PERIOD_SLICE_RANGE_INVALID",
                        "Period-slice bounds must be ordered dates without time components.");
                }
            }
            else if (slice.Kind == PeriodSliceKind.Prior || slice.Kind == PeriodSliceKind.SamePeriodPriorYear)
            {
                PeriodSliceSpec baseSlice;
                if (string.IsNullOrWhiteSpace(slice.BasedOnSliceId)
                    || !byId.TryGetValue(slice.BasedOnSliceId!, out baseSlice))
                {
                    throw new PeriodSliceResolutionException(
                        "SLICE_BASE_UNKNOWN",
                        "A relative period slice must reference another slice.");
                }

                var resolvedBase = ResolveOne(baseSlice, byId, resolved, resolving);
                if (slice.Kind == PeriodSliceKind.Prior)
                {
                    if (resolvedBase.StartInclusive.Day == 1
                        && resolvedBase.EndInclusive.Day
                            == DateTime.DaysInMonth(resolvedBase.EndInclusive.Year, resolvedBase.EndInclusive.Month))
                    {
                        var calendarMonths = checked(
                            (resolvedBase.EndInclusive.Year - resolvedBase.StartInclusive.Year) * 12
                            + resolvedBase.EndInclusive.Month
                            - resolvedBase.StartInclusive.Month
                            + 1);
                        start = resolvedBase.StartInclusive.AddMonths(-calendarMonths);
                    }
                    else
                    {
                        var inclusiveDays = checked((resolvedBase.EndInclusive - resolvedBase.StartInclusive).Days + 1);
                        start = resolvedBase.StartInclusive.AddDays(-inclusiveDays);
                    }

                    end = resolvedBase.StartInclusive.AddDays(-1);
                }
                else
                {
                    start = resolvedBase.StartInclusive.AddYears(-1);
                    end = resolvedBase.EndInclusive.AddYears(-1);
                }
            }
            else
            {
                throw new PeriodSliceResolutionException("PERIOD_SLICE_KIND_INVALID", "The period-slice kind is not supported.");
            }

            resolving.Remove(slice.Id);
            var result = new ResolvedPeriodSlice
            {
                Id = slice.Id,
                Label = slice.Label,
                Kind = slice.Kind,
                StartInclusive = start,
                EndInclusive = end,
                BasedOnSliceId = slice.BasedOnSliceId
            };
            resolved.Add(slice.Id, result);
            return result;
        }
    }
}
