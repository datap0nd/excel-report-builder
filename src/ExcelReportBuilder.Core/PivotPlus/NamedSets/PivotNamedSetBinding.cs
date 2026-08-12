using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportBuilder.Core.PivotPlus.NamedSets
{
    internal sealed class PivotNamedSetBoundLevel
    {
        public PivotNamedSetBoundLevel(
            PivotNamedSetHierarchySchema hierarchy,
            PivotNamedSetLevelSchema level)
        {
            Hierarchy = hierarchy;
            Level = level;
        }

        public PivotNamedSetHierarchySchema Hierarchy { get; }

        public PivotNamedSetLevelSchema Level { get; }
    }

    internal sealed class PivotNamedSetBoundMember
    {
        public PivotNamedSetBoundMember(
            PivotNamedSetHierarchySchema hierarchy,
            PivotNamedSetLevelSchema level,
            PivotNamedSetMemberSchema member)
        {
            Hierarchy = hierarchy;
            Level = level;
            Member = member;
        }

        public PivotNamedSetHierarchySchema Hierarchy { get; }

        public PivotNamedSetLevelSchema Level { get; }

        public PivotNamedSetMemberSchema Member { get; }
    }

    internal sealed class PivotNamedSetSchemaIndex
    {
        private readonly Dictionary<string, PivotNamedSetHierarchySchema> hierarchies =
            new Dictionary<string, PivotNamedSetHierarchySchema>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, PivotNamedSetBoundLevel> levels =
            new Dictionary<string, PivotNamedSetBoundLevel>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, PivotNamedSetBoundMember> members =
            new Dictionary<string, PivotNamedSetBoundMember>(StringComparer.OrdinalIgnoreCase);

        public PivotNamedSetSchemaIndex(PivotNamedSetSchema schema)
        {
            ProviderKind = schema.ProviderKind;
            foreach (PivotNamedSetHierarchySchema? hierarchy in schema.Hierarchies)
            {
                if (hierarchy == null || string.IsNullOrWhiteSpace(hierarchy.Id) ||
                    hierarchies.ContainsKey(hierarchy.Id))
                {
                    continue;
                }

                hierarchies.Add(hierarchy.Id, hierarchy);
                foreach (PivotNamedSetLevelSchema? level in hierarchy.Levels)
                {
                    if (level == null || string.IsNullOrWhiteSpace(level.Id) ||
                        levels.ContainsKey(level.Id))
                    {
                        continue;
                    }

                    var boundLevel = new PivotNamedSetBoundLevel(hierarchy, level);
                    levels.Add(level.Id, boundLevel);
                    foreach (PivotNamedSetMemberSchema? member in level.Members)
                    {
                        if (member == null || string.IsNullOrWhiteSpace(member.Id) ||
                            members.ContainsKey(member.Id))
                        {
                            continue;
                        }

                        members.Add(
                            member.Id,
                            new PivotNamedSetBoundMember(hierarchy, level, member));
                    }
                }
            }
        }

        public PivotNamedSetProviderKind ProviderKind { get; }

        public bool TryGetHierarchy(
            string id,
            out PivotNamedSetHierarchySchema hierarchy)
        {
            return hierarchies.TryGetValue(id, out hierarchy!);
        }

        public bool TryGetLevel(string id, out PivotNamedSetBoundLevel level)
        {
            return levels.TryGetValue(id, out level!);
        }

        public bool TryGetMember(string id, out PivotNamedSetBoundMember member)
        {
            return members.TryGetValue(id, out member!);
        }

        public IReadOnlyList<PivotNamedSetBoundMember> Members => members.Values.ToList();
    }
}
