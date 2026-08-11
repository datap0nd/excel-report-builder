using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReportBuilder.Core.Periods;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Rendering
{
    public sealed class PeriodMember
    {
        public DateTime Period { get; set; }

        public object PivotValue { get; set; } = string.Empty;
    }

    /// <summary>
    /// Resolves semantic period slices against the periods that actually exist
    /// in the managed PivotTable. Relative slices therefore never invent a
    /// period or silently substitute the current calendar date.
    /// </summary>
    public static class PeriodSliceResolver
    {
        /// <summary>
        /// Binds the compiler's immutable explicit ranges to the period members
        /// that actually exist in the managed PivotTable. The Excel host must
        /// not reinterpret relative period semantics after planning.
        /// </summary>
        public static IReadOnlyDictionary<string, IReadOnlyList<object>> BindResolved(
            IReadOnlyList<ResolvedPeriodSlice> slices,
            IReadOnlyList<PeriodMember> availablePeriods)
        {
            if (slices == null) throw new ArgumentNullException(nameof(slices));
            if (availablePeriods == null) throw new ArgumentNullException(nameof(availablePeriods));

            var periods = availablePeriods
                .GroupBy(member => member.Period.Date)
                .Select(group => group.First())
                .OrderBy(member => member.Period)
                .ToList();
            var result = new Dictionary<string, IReadOnlyList<object>>(StringComparer.OrdinalIgnoreCase);
            foreach (var slice in slices)
            {
                var matches = periods
                    .Where(member => member.Period.Date >= slice.StartInclusive.Date &&
                                     member.Period.Date <= slice.EndInclusive.Date)
                    .Select(member => member.PivotValue)
                    .ToList();
                if (matches.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Period slice '" + slice.Label + "' has no matching periods in the managed data.");
                }

                result.Add(slice.Id, matches);
            }

            return result;
        }

        public static IReadOnlyDictionary<string, IReadOnlyList<object>> Resolve(
            IReadOnlyList<PeriodSliceSpec> slices,
            IReadOnlyList<PeriodMember> availablePeriods)
        {
            if (slices == null) throw new ArgumentNullException(nameof(slices));
            if (availablePeriods == null) throw new ArgumentNullException(nameof(availablePeriods));

            var periods = availablePeriods
                .GroupBy(member => member.Period.Date)
                .Select(group => group.First())
                .OrderBy(member => member.Period)
                .ToList();
            var definitions = slices.ToDictionary(slice => slice.Id, StringComparer.OrdinalIgnoreCase);
            var resolvedDates = new Dictionary<string, IReadOnlyList<DateTime>>(StringComparer.OrdinalIgnoreCase);
            var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var slice in slices)
            {
                ResolveDates(slice.Id, definitions, periods, resolvedDates, active);
            }

            var byDate = periods.ToDictionary(member => member.Period.Date, member => member.PivotValue);
            return resolvedDates.ToDictionary(
                pair => pair.Key,
                pair => (IReadOnlyList<object>)pair.Value.Select(date => byDate[date.Date]).ToList(),
                StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<DateTime> ResolveDates(
            string sliceId,
            IReadOnlyDictionary<string, PeriodSliceSpec> definitions,
            IReadOnlyList<PeriodMember> periods,
            IDictionary<string, IReadOnlyList<DateTime>> resolved,
            ISet<string> active)
        {
            if (resolved.TryGetValue(sliceId, out var existing))
            {
                return existing;
            }

            if (!definitions.TryGetValue(sliceId, out var slice))
            {
                throw new InvalidOperationException("A requested period slice is not defined.");
            }

            if (!active.Add(sliceId))
            {
                throw new InvalidOperationException("The period-slice graph contains a cycle.");
            }

            try
            {
                IReadOnlyList<DateTime> dates;
                switch (slice.Kind)
                {
                    case PeriodSliceKind.Current:
                    case PeriodSliceKind.Selected:
                        if (!slice.SelectedStart.HasValue || !slice.SelectedEnd.HasValue)
                        {
                            throw new InvalidOperationException("A current or selected period slice requires a start and end date.");
                        }

                        dates = periods
                            .Where(member => member.Period.Date >= slice.SelectedStart.Value.Date &&
                                             member.Period.Date <= slice.SelectedEnd.Value.Date)
                            .Select(member => member.Period.Date)
                            .ToList();
                        break;
                    case PeriodSliceKind.Prior:
                    {
                        var basis = ResolveBasis(slice, definitions, periods, resolved, active);
                        var available = periods.Select(member => member.Period.Date).ToList();
                        var prior = new List<DateTime>();
                        foreach (var date in basis)
                        {
                            var index = available.BinarySearch(date.Date);
                            if (index > 0)
                            {
                                prior.Add(available[index - 1]);
                            }
                        }

                        dates = prior.Distinct().OrderBy(value => value).ToList();
                        break;
                    }
                    case PeriodSliceKind.SamePeriodPriorYear:
                    {
                        var basis = ResolveBasis(slice, definitions, periods, resolved, active);
                        var available = new HashSet<DateTime>(periods.Select(member => member.Period.Date));
                        dates = basis
                            .Select(date => date.AddYears(-1).Date)
                            .Where(available.Contains)
                            .Distinct()
                            .OrderBy(value => value)
                            .ToList();
                        break;
                    }
                    default:
                        throw new NotSupportedException("The period-slice kind is not supported.");
                }

                if (dates.Count == 0)
                {
                    throw new InvalidOperationException(
                        "Period slice '" + slice.Label + "' has no matching periods in the managed data.");
                }

                resolved[sliceId] = dates;
                return dates;
            }
            finally
            {
                active.Remove(sliceId);
            }
        }

        private static IReadOnlyList<DateTime> ResolveBasis(
            PeriodSliceSpec slice,
            IReadOnlyDictionary<string, PeriodSliceSpec> definitions,
            IReadOnlyList<PeriodMember> periods,
            IDictionary<string, IReadOnlyList<DateTime>> resolved,
            ISet<string> active)
        {
            if (string.IsNullOrWhiteSpace(slice.BasedOnSliceId))
            {
                throw new InvalidOperationException("A relative period slice requires a base slice.");
            }

            return ResolveDates(slice.BasedOnSliceId!, definitions, periods, resolved, active);
        }
    }
}
