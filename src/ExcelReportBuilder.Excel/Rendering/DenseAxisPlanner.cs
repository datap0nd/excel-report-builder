using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Rendering
{
    internal sealed class DenseAxisPath
    {
        public List<PivotFilterItem> DisplayItems { get; set; } = new List<PivotFilterItem>();

        public List<IReadOnlyList<PivotFilterItem>> MemberFilterSets { get; set; } =
            new List<IReadOnlyList<PivotFilterItem>>();

        public bool IsSubtotal { get; set; }

        public int SubtotalLevel { get; set; } = -1;

        public string? StyleId { get; set; }
    }

    /// <summary>
    /// Applies the typed member pipeline without changing the canonical data or
    /// requiring unsupported native PivotTable groups. Every displayed bucket
    /// retains the exact raw member filter sets used by GETPIVOTDATA.
    /// </summary>
    internal static class DenseAxisPlanner
    {
        public static List<DenseAxisPath> Build(
            IReadOnlyList<List<PivotFilterItem>> rawPaths,
            IReadOnlyList<PivotFieldPlan> fields,
            string? blockSubtotalStyleId,
            Func<string, IReadOnlyList<PivotFilterItem>, decimal> score)
        {
            if (rawPaths == null) throw new ArgumentNullException(nameof(rawPaths));
            if (fields == null) throw new ArgumentNullException(nameof(fields));
            if (score == null) throw new ArgumentNullException(nameof(score));

            if (fields.Count == 0)
            {
                return new List<DenseAxisPath>
                {
                    new DenseAxisPath
                    {
                        MemberFilterSets =
                        {
                            (IReadOnlyList<PivotFilterItem>)Array.Empty<PivotFilterItem>()
                        }
                    }
                };
            }

            var paths = rawPaths
                .Where(path => path.Count >= fields.Count)
                .Select(path => new DenseAxisPath
                {
                    DisplayItems = CloneItems(path.Take(fields.Count)),
                    MemberFilterSets =
                    {
                        (IReadOnlyList<PivotFilterItem>)CloneItems(path)
                    }
                })
                .ToList();

            for (var level = 0; level < fields.Count; level++)
            {
                var field = fields[level];
                foreach (var stage in field.MemberStages)
                {
                    switch (stage)
                    {
                        case PivotMemberStageKind.ApplyMemberOrder:
                            paths = ApplyMemberOrder(paths, level, field.MemberOrder);
                            break;
                        case PivotMemberStageKind.GroupMembers:
                            paths = ApplyGroups(paths, level, field.GroupBuckets);
                            break;
                        case PivotMemberStageKind.SortAscending:
                            paths = SortMembers(paths, level, descending: false);
                            break;
                        case PivotMemberStageKind.SortDescending:
                            paths = SortMembers(paths, level, descending: true);
                            break;
                        case PivotMemberStageKind.ApplyTopN:
                            if (field.TopN != null)
                            {
                                paths = ApplyTopN(paths, level, field.TopN, score);
                            }

                            break;
                        case PivotMemberStageKind.AggregateOthers:
                            // ApplyTopN creates and aggregates the bounded
                            // remainder when IncludeOthers is enabled.
                            break;
                        default:
                            throw new NotSupportedException("The dense member stage is not supported.");
                    }
                }
            }

            return InsertSubtotals(paths, fields, blockSubtotalStyleId, 0);
        }

        private static List<DenseAxisPath> ApplyMemberOrder(
            List<DenseAxisPath> paths,
            int level,
            IReadOnlyList<ScalarValue> order)
        {
            if (order.Count == 0)
            {
                return paths;
            }

            return StableSortWithinParents(
                paths,
                level,
                path =>
                {
                    var rank = IndexOf(order, path.DisplayItems[level].Value);
                    return rank < 0 ? int.MaxValue : rank;
                },
                descending: false);
        }

        private static List<DenseAxisPath> ApplyGroups(
            List<DenseAxisPath> paths,
            int level,
            IReadOnlyList<MemberGroupBucketSpec> buckets)
        {
            if (buckets.Count == 0)
            {
                return paths;
            }

            var unmatched = buckets.FirstOrDefault(bucket => bucket.IncludeUnmatched);
            foreach (var path in paths)
            {
                var value = path.DisplayItems[level].Value;
                var bucket = buckets.FirstOrDefault(candidate =>
                    candidate.Members.Any(member => ScalarMatches(member, value)));
                bucket = bucket ?? unmatched;
                if (bucket != null)
                {
                    path.DisplayItems[level] = new PivotFilterItem
                    {
                        Field = path.DisplayItems[level].Field,
                        Value = bucket.Label
                    };
                }
            }

            return Collapse(paths);
        }

        private static List<DenseAxisPath> ApplyTopN(
            List<DenseAxisPath> paths,
            int level,
            TopNSpec topN,
            Func<string, IReadOnlyList<PivotFilterItem>, decimal> score)
        {
            var keep = new HashSet<string>(StringComparer.Ordinal);
            foreach (var parent in paths.GroupBy(path => PrefixKey(path.DisplayItems, level), StringComparer.Ordinal))
            {
                var members = parent
                    .GroupBy(path => ItemKey(path.DisplayItems[level]), StringComparer.Ordinal)
                    .Select(group => new
                    {
                        Key = group.Key,
                        Score = group.SelectMany(path => path.MemberFilterSets)
                            .Sum(filters => score(topN.MeasureId, filters)),
                        First = paths.IndexOf(group.First())
                    });
                members = topN.Direction == TopNDirection.Top
                    ? members.OrderByDescending(member => member.Score).ThenBy(member => member.First)
                    : members.OrderBy(member => member.Score).ThenBy(member => member.First);
                foreach (var member in members.Take(topN.Count))
                {
                    keep.Add(parent.Key + "\u001d" + member.Key);
                }
            }

            var result = new List<DenseAxisPath>();
            foreach (var path in paths)
            {
                var candidateKey = PrefixKey(path.DisplayItems, level) + "\u001d" +
                                   ItemKey(path.DisplayItems[level]);
                if (keep.Contains(candidateKey))
                {
                    result.Add(path);
                    continue;
                }

                if (!topN.IncludeOthers)
                {
                    continue;
                }

                path.DisplayItems[level] = new PivotFilterItem
                {
                    Field = path.DisplayItems[level].Field,
                    Value = topN.OthersLabel
                };
                for (var childLevel = level + 1; childLevel < path.DisplayItems.Count; childLevel++)
                {
                    path.DisplayItems[childLevel] = new PivotFilterItem
                    {
                        Field = path.DisplayItems[childLevel].Field,
                        Value = null
                    };
                }
                result.Add(path);
            }

            return Collapse(result);
        }

        private static List<DenseAxisPath> SortMembers(
            List<DenseAxisPath> paths,
            int level,
            bool descending)
        {
            return StableSortWithinParents(
                paths,
                level,
                path => Convert.ToString(
                    path.DisplayItems[level].Value,
                    CultureInfo.InvariantCulture) ?? string.Empty,
                descending);
        }

        private static List<DenseAxisPath> StableSortWithinParents<TKey>(
            List<DenseAxisPath> paths,
            int level,
            Func<DenseAxisPath, TKey> selector,
            bool descending)
        {
            var result = new List<DenseAxisPath>();
            foreach (var parent in paths.GroupBy(path => PrefixKey(path.DisplayItems, level), StringComparer.Ordinal))
            {
                var indexed = parent.Select((path, index) => new { Path = path, Index = index });
                result.AddRange(descending
                    ? indexed.OrderByDescending(item => selector(item.Path)).ThenBy(item => item.Index).Select(item => item.Path)
                    : indexed.OrderBy(item => selector(item.Path)).ThenBy(item => item.Index).Select(item => item.Path));
            }

            return result;
        }

        private static List<DenseAxisPath> InsertSubtotals(
            List<DenseAxisPath> details,
            IReadOnlyList<PivotFieldPlan> fields,
            string? blockSubtotalStyleId,
            int level)
        {
            if (level >= fields.Count)
            {
                return details;
            }

            var result = new List<DenseAxisPath>();
            foreach (var group in details.GroupBy(
                         path => PrefixKey(path.DisplayItems, level + 1),
                         StringComparer.Ordinal))
            {
                var groupDetails = group.ToList();
                var children = level + 1 < fields.Count
                    ? InsertSubtotals(groupDetails, fields, blockSubtotalStyleId, level + 1)
                    : groupDetails;
                var subtotal = fields[level].Subtotals.Mode == SubtotalMode.Automatic
                    ? CreateSubtotal(groupDetails, fields[level], level, blockSubtotalStyleId)
                    : null;

                if (subtotal != null && fields[level].Subtotals.Placement == TotalPlacement.BeforeMembers)
                {
                    result.Add(subtotal);
                }

                result.AddRange(children);
                if (subtotal != null && fields[level].Subtotals.Placement == TotalPlacement.AfterMembers)
                {
                    result.Add(subtotal);
                }
            }

            return result;
        }

        private static DenseAxisPath CreateSubtotal(
            IReadOnlyList<DenseAxisPath> details,
            PivotFieldPlan field,
            int level,
            string? blockSubtotalStyleId)
        {
            var display = CloneItems(details[0].DisplayItems.Take(level + 1));
            var sourceLabel = Convert.ToString(display[level].Value, CultureInfo.InvariantCulture) ?? string.Empty;
            display[level] = new PivotFilterItem
            {
                Field = display[level].Field,
                Value = !string.IsNullOrWhiteSpace(field.Subtotals.Label)
                    ? field.Subtotals.Label
                    : sourceLabel + " Total"
            };
            return new DenseAxisPath
            {
                DisplayItems = display,
                MemberFilterSets = details.SelectMany(path => path.MemberFilterSets).ToList(),
                IsSubtotal = true,
                SubtotalLevel = level,
                StyleId = field.Subtotals.StyleId ?? blockSubtotalStyleId
            };
        }

        private static List<DenseAxisPath> Collapse(IEnumerable<DenseAxisPath> paths)
        {
            var result = new List<DenseAxisPath>();
            var indexes = new Dictionary<string, DenseAxisPath>(StringComparer.Ordinal);
            foreach (var path in paths)
            {
                var key = PrefixKey(path.DisplayItems, path.DisplayItems.Count);
                if (indexes.TryGetValue(key, out var existing))
                {
                    existing.MemberFilterSets.AddRange(path.MemberFilterSets);
                }
                else
                {
                    indexes[key] = path;
                    result.Add(path);
                }
            }

            return result;
        }

        private static int IndexOf(IReadOnlyList<ScalarValue> values, object? candidate)
        {
            for (var index = 0; index < values.Count; index++)
            {
                if (ScalarMatches(values[index], candidate))
                {
                    return index;
                }
            }

            return -1;
        }

        private static bool ScalarMatches(ScalarValue value, object? candidate)
        {
            switch (value.Kind)
            {
                case ScalarValueKind.Null:
                    return candidate == null || string.IsNullOrWhiteSpace(
                        Convert.ToString(candidate, CultureInfo.InvariantCulture));
                case ScalarValueKind.Text:
                    return string.Equals(
                        value.Text,
                        Convert.ToString(candidate, CultureInfo.InvariantCulture),
                        StringComparison.OrdinalIgnoreCase);
                case ScalarValueKind.Number:
                    try { return value.Number == Convert.ToDecimal(candidate, CultureInfo.InvariantCulture); }
                    catch (Exception) { return false; }
                case ScalarValueKind.Boolean:
                    try { return value.Boolean == Convert.ToBoolean(candidate, CultureInfo.InvariantCulture); }
                    catch (Exception) { return false; }
                case ScalarValueKind.Date:
                case ScalarValueKind.DateTime:
                    DateTime temporal;
                    if (candidate is DateTime date)
                    {
                        temporal = date;
                    }
                    else if (!DateTime.TryParse(
                                 Convert.ToString(candidate, CultureInfo.InvariantCulture),
                                 CultureInfo.InvariantCulture,
                                 DateTimeStyles.AllowWhiteSpaces,
                                 out temporal))
                    {
                        return false;
                    }

                    return value.Kind == ScalarValueKind.Date
                        ? value.Temporal?.Date == temporal.Date
                        : value.Temporal == temporal;
                default:
                    return false;
            }
        }

        private static string PrefixKey(IReadOnlyList<PivotFilterItem> items, int count)
        {
            return string.Join("\u001f", items.Take(count).Select(ItemKey));
        }

        private static string ItemKey(PivotFilterItem item)
        {
            return item.Field.ToUpperInvariant() + "\u001e" +
                   (Convert.ToString(item.Value, CultureInfo.InvariantCulture) ?? "<blank>").ToUpperInvariant();
        }

        private static List<PivotFilterItem> CloneItems(IEnumerable<PivotFilterItem> items)
        {
            return items.Select(item => new PivotFilterItem
            {
                Field = item.Field,
                Value = item.Value
            }).ToList();
        }
    }
}
