using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.NamedSets
{
    public static class PivotNamedSetValidator
    {
        internal const int MaximumCompiledFormulaCharacters = 24 * 1024;

        private const int MaximumHierarchies = 64;
        private const int MaximumLevels = 256;
        private const int MaximumMembers = 4096;
        private const int MaximumNamedSets = 2;
        private const int MaximumTupleArity = 8;
        private const int MaximumTuples = 1024;
        private const int MaximumMemberReferences = 4096;
        private const int MaximumTopN = 1000;
        private const int MaximumIdLength = 128;
        private const int MaximumCaptionLength = 255;
        private const int MaximumMemberCaptionLength = 1024;
        private const int MaximumProviderUniqueNameLength = 2048;

        private static readonly Regex IdPattern = new Regex(
            "^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$",
            RegexOptions.CultureInvariant);

        private static readonly Regex FingerprintPattern = new Regex(
            "^[a-z0-9][a-z0-9._-]{0,63}:sha256:[0-9a-f]{64}$",
            RegexOptions.CultureInvariant);

        private static readonly Regex GeneratedSetNamePattern = new Regex(
            "^\\[[A-Za-z0-9][A-Za-z0-9._-]{0,126}\\]$",
            RegexOptions.CultureInvariant);

        public static ValidationResult Validate(PivotNamedSetCompilationRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            var result = new ValidationResult();
            PivotNamedSetCollectionDefinition definition = request.Definition;
            ValidateFingerprint(
                definition.SourceFingerprint,
                "sourceFingerprint",
                "PIVOT_SET_SOURCE_FINGERPRINT_INVALID",
                result);
            ValidateSchema(definition.Schema, result);
            if (!string.Equals(
                    definition.SourceFingerprint,
                    definition.Schema.SourceFingerprint,
                    StringComparison.Ordinal))
            {
                result.AddError(
                    "PIVOT_SET_SOURCE_FINGERPRINT_MISMATCH",
                    "sourceFingerprint",
                    "The action source fingerprint does not match the transient host catalog.");
            }

            var index = new PivotNamedSetSchemaIndex(definition.Schema);
            ValidateNamedSets(definition.NamedSets, request.DaxCompilation, index, result);
            ValidateArtifactBindings(
                definition.NamedSets,
                request.ArtifactBindings,
                result);

            if (result.IsValid)
            {
                ValidateCompiledFormulaBounds(request, index, result);
            }

            return result;
        }

        private static void ValidateSchema(
            PivotNamedSetSchema schema,
            ValidationResult result)
        {
            ValidateFingerprint(
                schema.SourceFingerprint,
                "schema.sourceFingerprint",
                "PIVOT_SET_SCHEMA_SOURCE_FINGERPRINT_INVALID",
                result);
            if (schema.ProviderKind != PivotNamedSetProviderKind.DataModel)
            {
                result.AddError(
                    "PIVOT_SET_SCHEMA_PROVIDER_UNSUPPORTED",
                    "schema.providerKind",
                    "Named sets are supported only for the Excel Data Model provider.");
            }

            if (schema.Hierarchies.Count == 0)
            {
                result.AddError(
                    "PIVOT_SET_SCHEMA_HIERARCHY_REQUIRED",
                    "schema.hierarchies",
                    "At least one exact Data Model hierarchy is required.");
            }
            else if (schema.Hierarchies.Count > MaximumHierarchies)
            {
                result.AddError(
                    "PIVOT_SET_SCHEMA_HIERARCHY_LIMIT",
                    "schema.hierarchies",
                    "The transient hierarchy catalog exceeds its bounded limit.");
            }

            var hierarchyIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var hierarchyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var levelIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var levelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var memberIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var memberNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var totalLevels = 0;
            var totalMembers = 0;

            for (var hierarchyIndex = 0;
                 hierarchyIndex < schema.Hierarchies.Count;
                 hierarchyIndex++)
            {
                PivotNamedSetHierarchySchema? hierarchy =
                    schema.Hierarchies[hierarchyIndex];
                string path = "schema.hierarchies[" +
                    hierarchyIndex.ToString(CultureInfo.InvariantCulture) + "]";
                if (hierarchy == null)
                {
                    result.AddError(
                        "PIVOT_SET_SCHEMA_HIERARCHY_NULL",
                        path,
                        "Hierarchy entries cannot be null.");
                    continue;
                }

                ValidateId(
                    hierarchy.Id,
                    path + ".id",
                    "PIVOT_SET_HIERARCHY_ID_INVALID",
                    result);
                if (!hierarchyIds.Add(hierarchy.Id))
                {
                    result.AddError(
                        "PIVOT_SET_HIERARCHY_ID_DUPLICATE",
                        path + ".id",
                        "Hierarchy IDs must be unique without regard to case.");
                }

                ValidateProviderUniqueName(
                    hierarchy.ProviderUniqueName,
                    path + ".providerUniqueName",
                    "PIVOT_SET_HIERARCHY_UNIQUE_NAME_INVALID",
                    result);
                if (!hierarchyNames.Add(hierarchy.ProviderUniqueName))
                {
                    result.AddError(
                        "PIVOT_SET_HIERARCHY_UNIQUE_NAME_DUPLICATE",
                        path + ".providerUniqueName",
                        "Provider hierarchy unique names must be unambiguous.");
                }

                ValidateOptionalCaption(
                    hierarchy.Caption,
                    MaximumCaptionLength,
                    path + ".caption",
                    "PIVOT_SET_HIERARCHY_CAPTION_INVALID",
                    result);
                if (hierarchy.AllMemberId != null)
                {
                    ValidateId(
                        hierarchy.AllMemberId,
                        path + ".allMemberId",
                        "PIVOT_SET_ALL_MEMBER_ID_INVALID",
                        result);
                }

                if (!hierarchy.IdentityComplete)
                {
                    result.AddError(
                        "PIVOT_SET_HIERARCHY_IDENTITY_INCOMPLETE",
                        path + ".identityComplete",
                        "Named sets require complete host-discovered hierarchy identity.");
                }

                if (hierarchy.Levels.Count == 0)
                {
                    result.AddError(
                        "PIVOT_SET_LEVEL_REQUIRED",
                        path + ".levels",
                        "Every hierarchy requires at least one exact level.");
                }

                totalLevels += hierarchy.Levels.Count;
                var ordinals = new HashSet<int>();
                var hierarchyMembers = new List<PivotNamedSetMemberSchema>();
                for (var levelIndex = 0;
                     levelIndex < hierarchy.Levels.Count;
                     levelIndex++)
                {
                    PivotNamedSetLevelSchema? level = hierarchy.Levels[levelIndex];
                    string levelPath = path + ".levels[" +
                        levelIndex.ToString(CultureInfo.InvariantCulture) + "]";
                    if (level == null)
                    {
                        result.AddError(
                            "PIVOT_SET_LEVEL_NULL",
                            levelPath,
                            "Level entries cannot be null.");
                        continue;
                    }

                    ValidateId(
                        level.Id,
                        levelPath + ".id",
                        "PIVOT_SET_LEVEL_ID_INVALID",
                        result);
                    if (!levelIds.Add(level.Id))
                    {
                        result.AddError(
                            "PIVOT_SET_LEVEL_ID_DUPLICATE",
                            levelPath + ".id",
                            "Level IDs must be globally unique without regard to case.");
                    }

                    ValidateProviderUniqueName(
                        level.ProviderUniqueName,
                        levelPath + ".providerUniqueName",
                        "PIVOT_SET_LEVEL_UNIQUE_NAME_INVALID",
                        result);
                    if (!levelNames.Add(level.ProviderUniqueName))
                    {
                        result.AddError(
                            "PIVOT_SET_LEVEL_UNIQUE_NAME_DUPLICATE",
                            levelPath + ".providerUniqueName",
                            "Provider level unique names must be unambiguous.");
                    }

                    if (level.Ordinal < 0 || !ordinals.Add(level.Ordinal))
                    {
                        result.AddError(
                            "PIVOT_SET_LEVEL_ORDINAL_INVALID",
                            levelPath + ".ordinal",
                            "Level ordinals must be non-negative and unique within a hierarchy.");
                    }

                    totalMembers += level.Members.Count;
                    for (var memberIndex = 0;
                         memberIndex < level.Members.Count;
                         memberIndex++)
                    {
                        PivotNamedSetMemberSchema? member = level.Members[memberIndex];
                        string memberPath = levelPath + ".members[" +
                            memberIndex.ToString(CultureInfo.InvariantCulture) + "]";
                        if (member == null)
                        {
                            result.AddError(
                                "PIVOT_SET_MEMBER_NULL",
                                memberPath,
                                "Member entries cannot be null.");
                            continue;
                        }

                        hierarchyMembers.Add(member);
                        ValidateId(
                            member.Id,
                            memberPath + ".id",
                            "PIVOT_SET_MEMBER_ID_INVALID",
                            result);
                        if (!memberIds.Add(member.Id))
                        {
                            result.AddError(
                                "PIVOT_SET_MEMBER_ID_DUPLICATE",
                                memberPath + ".id",
                                "Member IDs must be globally unique without regard to case.");
                        }

                        ValidateProviderUniqueName(
                            member.ProviderUniqueName,
                            memberPath + ".providerUniqueName",
                            "PIVOT_SET_MEMBER_UNIQUE_NAME_INVALID",
                            result);
                        if (!memberNames.Add(member.ProviderUniqueName))
                        {
                            result.AddError(
                                "PIVOT_SET_MEMBER_UNIQUE_NAME_DUPLICATE",
                                memberPath + ".providerUniqueName",
                                "Provider member unique names must be unambiguous.");
                        }

                        ValidateOptionalCaption(
                            member.Caption,
                            MaximumMemberCaptionLength,
                            memberPath + ".caption",
                            "PIVOT_SET_MEMBER_CAPTION_INVALID",
                            result);
                        if (member.ParentMemberId != null)
                        {
                            ValidateId(
                                member.ParentMemberId,
                                memberPath + ".parentMemberId",
                                "PIVOT_SET_PARENT_MEMBER_ID_INVALID",
                                result);
                            if (string.Equals(
                                    member.Id,
                                    member.ParentMemberId,
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                result.AddError(
                                    "PIVOT_SET_PARENT_MEMBER_SELF_REFERENCE",
                                    memberPath + ".parentMemberId",
                                    "A member cannot be its own parent.");
                            }
                        }
                    }
                }

                if (hierarchy.AllMemberId != null)
                {
                    List<PivotNamedSetMemberSchema> declaredAll = hierarchyMembers
                        .Where(member => member.IsAllMember)
                        .ToList();
                    if (declaredAll.Count != 1 ||
                        !string.Equals(
                            declaredAll[0].Id,
                            hierarchy.AllMemberId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        result.AddError(
                            "PIVOT_SET_ALL_MEMBER_UNRESOLVED",
                            path + ".allMemberId",
                            "A supplied All-member ID must resolve to the one exact catalog member proven as All.");
                    }
                }
            }

            if (totalLevels > MaximumLevels)
            {
                result.AddError(
                    "PIVOT_SET_SCHEMA_LEVEL_LIMIT",
                    "schema.hierarchies",
                    "The transient level catalog exceeds its bounded limit.");
            }

            if (totalMembers > MaximumMembers)
            {
                result.AddError(
                    "PIVOT_SET_SCHEMA_MEMBER_LIMIT",
                    "schema.hierarchies",
                    "The transient member catalog exceeds its bounded limit.");
            }

            ValidateParentBindings(schema, result);
        }

        private static void ValidateParentBindings(
            PivotNamedSetSchema schema,
            ValidationResult result)
        {
            var index = new PivotNamedSetSchemaIndex(schema);
            foreach (PivotNamedSetBoundMember member in index.Members)
            {
                string? parentId = member.Member.ParentMemberId;
                if (parentId == null)
                {
                    continue;
                }

                if (!index.TryGetMember(parentId, out PivotNamedSetBoundMember parent))
                {
                    result.AddError(
                        "PIVOT_SET_PARENT_MEMBER_UNKNOWN",
                        "schema.members[" + member.Member.Id + "].parentMemberId",
                        "A parent member ID does not resolve in the transient catalog.");
                    continue;
                }

                if (!string.Equals(
                        member.Hierarchy.Id,
                        parent.Hierarchy.Id,
                        StringComparison.OrdinalIgnoreCase))
                {
                    result.AddError(
                        "PIVOT_SET_PARENT_MEMBER_HIERARCHY_MISMATCH",
                        "schema.members[" + member.Member.Id + "].parentMemberId",
                        "A member parent must belong to the same hierarchy.");
                }

                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    member.Member.Id
                };
                PivotNamedSetBoundMember cursor = parent;
                while (true)
                {
                    if (!visited.Add(cursor.Member.Id))
                    {
                        result.AddError(
                            "PIVOT_SET_MEMBER_PARENT_CYCLE",
                            "schema.members[" + member.Member.Id + "].parentMemberId",
                            "Member parent identities cannot contain a cycle.");
                        break;
                    }

                    if (cursor.Member.ParentMemberId == null ||
                        !index.TryGetMember(
                            cursor.Member.ParentMemberId,
                            out PivotNamedSetBoundMember next))
                    {
                        break;
                    }

                    cursor = next;
                }
            }
        }

        private static void ValidateNamedSets(
            IReadOnlyList<PivotNamedSetDefinition> namedSets,
            PivotDaxCompilation? daxCompilation,
            PivotNamedSetSchemaIndex index,
            ValidationResult result)
        {
            if (namedSets.Count == 0)
            {
                result.AddError(
                    "PIVOT_SET_DEFINITION_REQUIRED",
                    "namedSets",
                    "At least one named-set definition is required.");
            }
            else if (namedSets.Count > MaximumNamedSets)
            {
                result.AddError(
                    "PIVOT_SET_DEFINITION_LIMIT",
                    "namedSets",
                    "Phase one supports at most one row set and one column set.");
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var captions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var axes = new HashSet<PivotNamedSetAxis>();
            var totalTuples = 0;
            var totalMemberReferences = 0;
            for (var setIndex = 0; setIndex < namedSets.Count; setIndex++)
            {
                PivotNamedSetDefinition? namedSet = namedSets[setIndex];
                string path = "namedSets[" +
                    setIndex.ToString(CultureInfo.InvariantCulture) + "]";
                if (namedSet == null)
                {
                    result.AddError(
                        "PIVOT_SET_DEFINITION_NULL",
                        path,
                        "Named-set entries cannot be null.");
                    continue;
                }

                ValidateId(
                    namedSet.Id,
                    path + ".id",
                    "PIVOT_SET_DEFINITION_ID_INVALID",
                    result);
                if (!ids.Add(namedSet.Id))
                {
                    result.AddError(
                        "PIVOT_SET_DEFINITION_ID_DUPLICATE",
                        path + ".id",
                        "Named-set definition IDs must be unique without regard to case.");
                }

                ValidateCaption(
                    namedSet.Caption,
                    MaximumCaptionLength,
                    path + ".caption",
                    "PIVOT_SET_CAPTION_INVALID",
                    result);
                if (!captions.Add(namedSet.Caption))
                {
                    result.AddError(
                        "PIVOT_SET_CAPTION_DUPLICATE",
                        path + ".caption",
                        "Named-set field captions must be unique without regard to case.");
                }

                if ((namedSet.Axis != PivotNamedSetAxis.Row &&
                     namedSet.Axis != PivotNamedSetAxis.Column) ||
                    !axes.Add(namedSet.Axis))
                {
                    result.AddError(
                        "PIVOT_SET_AXIS_INVALID",
                        path + ".axis",
                        "Each definition requires a unique Row or Column axis.");
                }

                if (namedSet.Expression == null)
                {
                    result.AddError(
                        "PIVOT_SET_EXPRESSION_NULL",
                        path + ".expression",
                        "A typed named-set expression is required.");
                    continue;
                }

                switch (namedSet.Expression)
                {
                    case PivotExplicitOrderedTuplesExpression explicitTuples:
                        ValidateExplicitTuples(
                            explicitTuples,
                            path + ".expression",
                            index,
                            result,
                            ref totalTuples,
                            ref totalMemberReferences);
                        break;
                    case PivotTopNLevelMembersExpression topN:
                        ValidateTopN(
                            topN,
                            path + ".expression",
                            daxCompilation,
                            index,
                            result);
                        break;
                    default:
                        result.AddError(
                            "PIVOT_SET_EXPRESSION_KIND_INVALID",
                            path + ".expression",
                            "Only explicit ordered tuples and typed Top N are supported.");
                        break;
                }
            }

            if (totalTuples > MaximumTuples)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_LIMIT",
                    "namedSets",
                    "The named-set request exceeds the bounded tuple limit.");
            }

            if (totalMemberReferences > MaximumMemberReferences)
            {
                result.AddError(
                    "PIVOT_SET_MEMBER_REFERENCE_LIMIT",
                    "namedSets",
                    "The named-set request exceeds the bounded tuple-member limit.");
            }
        }

        private static void ValidateExplicitTuples(
            PivotExplicitOrderedTuplesExpression expression,
            string path,
            PivotNamedSetSchemaIndex index,
            ValidationResult result,
            ref int totalTuples,
            ref int totalMemberReferences)
        {
            if (expression.Kind != PivotNamedSetExpressionKind.ExplicitOrderedTuples)
            {
                result.AddError(
                    "PIVOT_SET_EXPRESSION_KIND_MISMATCH",
                    path + ".kind",
                    "The expression runtime type and kind do not match.");
            }

            if (expression.HierarchyIds.Count == 0 ||
                expression.HierarchyIds.Count > MaximumTupleArity)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_ARITY_INVALID",
                    path + ".hierarchyIds",
                    "Explicit tuples require between one and eight ordered hierarchies.");
            }

            var hierarchyIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var hierarchyIndex = 0;
                 hierarchyIndex < expression.HierarchyIds.Count;
                 hierarchyIndex++)
            {
                string hierarchyId = expression.HierarchyIds[hierarchyIndex] ?? string.Empty;
                string hierarchyPath = path + ".hierarchyIds[" +
                    hierarchyIndex.ToString(CultureInfo.InvariantCulture) + "]";
                ValidateId(
                    hierarchyId,
                    hierarchyPath,
                    "PIVOT_SET_TUPLE_HIERARCHY_ID_INVALID",
                    result);
                if (!hierarchyIds.Add(hierarchyId))
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_HIERARCHY_DUPLICATE",
                        hierarchyPath,
                        "A tuple cannot contain the same hierarchy more than once.");
                }

                if (!index.TryGetHierarchy(
                        hierarchyId,
                        out PivotNamedSetHierarchySchema hierarchy))
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_HIERARCHY_UNKNOWN",
                        hierarchyPath,
                        "A tuple hierarchy ID does not resolve in the transient catalog.");
                }
                else if (!hierarchy.IdentityComplete)
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_HIERARCHY_IDENTITY_INCOMPLETE",
                        hierarchyPath,
                        "Tuple construction requires complete exact hierarchy identity.");
                }
            }

            if (expression.Tuples.Count == 0)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_REQUIRED",
                    path + ".tuples",
                    "An explicit ordered set requires at least one tuple.");
            }

            totalTuples += expression.Tuples.Count;
            var tupleKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var tupleIndex = 0; tupleIndex < expression.Tuples.Count; tupleIndex++)
            {
                PivotNamedSetTuple? tuple = expression.Tuples[tupleIndex];
                string tuplePath = path + ".tuples[" +
                    tupleIndex.ToString(CultureInfo.InvariantCulture) + "]";
                if (tuple == null)
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_NULL",
                        tuplePath,
                        "Tuple entries cannot be null.");
                    continue;
                }

                totalMemberReferences += tuple.Members.Count;
                if (tuple.Members.Count != expression.HierarchyIds.Count)
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_MEMBER_COUNT_MISMATCH",
                        tuplePath + ".members",
                        "Every tuple must contain one member for each ordered hierarchy.");
                }

                string tupleKey = string.Join(
                    "\u001f",
                    tuple.Members.Select(TupleMemberKey));
                if (!tupleKeys.Add(tupleKey))
                {
                    result.AddError(
                        "PIVOT_SET_TUPLE_DUPLICATE",
                        tuplePath,
                        "Duplicate tuples are not allowed; Excel will not silently de-duplicate them.");
                }

                int comparable = Math.Min(
                    tuple.Members.Count,
                    expression.HierarchyIds.Count);
                for (var memberIndex = 0; memberIndex < tuple.Members.Count; memberIndex++)
                {
                    PivotNamedSetTupleMemberReference? reference =
                        tuple.Members[memberIndex];
                    string memberPath = tuplePath + ".members[" +
                        memberIndex.ToString(CultureInfo.InvariantCulture) + "]";
                    if (reference == null)
                    {
                        result.AddError(
                            "PIVOT_SET_TUPLE_MEMBER_NULL",
                            memberPath,
                            "Tuple-member references cannot be null.");
                        continue;
                    }

                    switch (reference)
                    {
                        case PivotNamedSetCatalogMemberReference catalogMember:
                            ValidateCatalogTupleMember(
                                catalogMember,
                                memberIndex,
                                comparable,
                                expression,
                                memberPath,
                                index,
                                result);
                            break;
                        case PivotNamedSetHierarchyDefaultMemberReference defaultMember:
                            ValidateDefaultTupleMember(
                                defaultMember,
                                memberIndex,
                                comparable,
                                expression,
                                memberPath,
                                index,
                                result);
                            break;
                        default:
                            result.AddError(
                                "PIVOT_SET_TUPLE_MEMBER_KIND_INVALID",
                                memberPath,
                                "Only catalog members and hierarchy DefaultMember references are supported.");
                            break;
                    }
                }
            }
        }

        private static void ValidateCatalogTupleMember(
            PivotNamedSetCatalogMemberReference reference,
            int memberIndex,
            int comparable,
            PivotExplicitOrderedTuplesExpression expression,
            string path,
            PivotNamedSetSchemaIndex index,
            ValidationResult result)
        {
            ValidateId(
                reference.MemberId,
                path + ".memberId",
                "PIVOT_SET_TUPLE_MEMBER_ID_INVALID",
                result);
            if (!index.TryGetMember(
                    reference.MemberId,
                    out PivotNamedSetBoundMember member))
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_MEMBER_UNKNOWN",
                    path + ".memberId",
                    "A tuple member ID does not resolve in the transient catalog.");
                return;
            }

            if (!member.Level.MembersComplete)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_MEMBER_IDENTITY_INCOMPLETE",
                    path + ".memberId",
                    "Catalog tuple members require a complete exact host member catalog.");
            }

            if (memberIndex < comparable &&
                !string.Equals(
                    member.Hierarchy.Id,
                    expression.HierarchyIds[memberIndex],
                    StringComparison.OrdinalIgnoreCase))
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_MEMBER_HIERARCHY_MISMATCH",
                    path + ".memberId",
                    "A tuple member does not belong to the hierarchy at that tuple position.");
            }
        }

        private static void ValidateDefaultTupleMember(
            PivotNamedSetHierarchyDefaultMemberReference reference,
            int memberIndex,
            int comparable,
            PivotExplicitOrderedTuplesExpression expression,
            string path,
            PivotNamedSetSchemaIndex index,
            ValidationResult result)
        {
            ValidateId(
                reference.HierarchyId,
                path + ".hierarchyId",
                "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_ID_INVALID",
                result);
            if (index.ProviderKind != PivotNamedSetProviderKind.DataModel)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_DEFAULT_PROVIDER_UNSUPPORTED",
                    path,
                    "Hierarchy DefaultMember references are supported only for the Excel Data Model provider.");
            }

            if (!index.TryGetHierarchy(
                    reference.HierarchyId,
                    out PivotNamedSetHierarchySchema hierarchy))
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_UNKNOWN",
                    path + ".hierarchyId",
                    "The DefaultMember hierarchy ID does not resolve in the transient catalog.");
                return;
            }

            if (!hierarchy.IdentityComplete)
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_INCOMPLETE",
                    path + ".hierarchyId",
                    "DefaultMember requires complete exact host hierarchy identity.");
            }

            if (memberIndex < comparable &&
                !string.Equals(
                    reference.HierarchyId,
                    expression.HierarchyIds[memberIndex],
                    StringComparison.OrdinalIgnoreCase))
            {
                result.AddError(
                    "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_MISMATCH",
                    path + ".hierarchyId",
                    "DefaultMember must reference the hierarchy at that tuple position.");
            }
        }

        private static string TupleMemberKey(
            PivotNamedSetTupleMemberReference? reference)
        {
            switch (reference)
            {
                case PivotNamedSetCatalogMemberReference catalogMember:
                    return "catalog:" + catalogMember.MemberId;
                case PivotNamedSetHierarchyDefaultMemberReference defaultMember:
                    return "default:" + defaultMember.HierarchyId;
                case null:
                    return "null";
                default:
                    return "unsupported:" + ((int)reference.Kind).ToString(
                        CultureInfo.InvariantCulture);
            }
        }

        private static void ValidateTopN(
            PivotTopNLevelMembersExpression expression,
            string path,
            PivotDaxCompilation? daxCompilation,
            PivotNamedSetSchemaIndex index,
            ValidationResult result)
        {
            if (expression.Kind != PivotNamedSetExpressionKind.TopNLevelMembers)
            {
                result.AddError(
                    "PIVOT_SET_EXPRESSION_KIND_MISMATCH",
                    path + ".kind",
                    "The expression runtime type and kind do not match.");
            }

            ValidateId(
                expression.LevelId,
                path + ".levelId",
                "PIVOT_SET_TOPN_LEVEL_ID_INVALID",
                result);
            if (!index.TryGetLevel(expression.LevelId, out PivotNamedSetBoundLevel level))
            {
                result.AddError(
                    "PIVOT_SET_TOPN_LEVEL_UNKNOWN",
                    path + ".levelId",
                    "The Top N level does not resolve in the transient catalog.");
            }
            else
            {
                if (!level.Hierarchy.IdentityComplete || !level.Level.MembersComplete)
                {
                    result.AddError(
                        "PIVOT_SET_TOPN_LEVEL_IDENTITY_INCOMPLETE",
                        path + ".levelId",
                        "Top N requires complete exact hierarchy and level identity.");
                }

                if (level.Level.Members.Any(member =>
                        member != null && member.IsAllMember))
                {
                    result.AddError(
                        "PIVOT_SET_TOPN_ALL_LEVEL_INVALID",
                        path + ".levelId",
                        "Top N cannot target the hierarchy's All-member level.");
                }
            }

            if (expression.Count < 1 || expression.Count > MaximumTopN)
            {
                result.AddError(
                    "PIVOT_SET_TOPN_COUNT_INVALID",
                    path + ".count",
                    "Top N count must be between one and one thousand.");
            }

            ValidateId(
                expression.MeasureDefinitionId,
                path + ".measureDefinitionId",
                "PIVOT_SET_TOPN_MEASURE_ID_INVALID",
                result);
            if (daxCompilation == null)
            {
                result.AddError(
                    "PIVOT_SET_TOPN_DAX_COMPILATION_REQUIRED",
                    path + ".measureDefinitionId",
                    "Top N requires a measure from the same validated DAX compilation.");
                return;
            }

            List<OwnedPivotMeasureDefinition> matches = daxCompilation.Measures
                .Where(measure => measure != null && string.Equals(
                    measure.DefinitionId,
                    expression.MeasureDefinitionId,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count != 1)
            {
                result.AddError(
                    "PIVOT_SET_TOPN_MEASURE_UNKNOWN",
                    path + ".measureDefinitionId",
                    "Top N must reference exactly one owned measure in the supplied DAX compilation.");
            }
        }

        private static void ValidateArtifactBindings(
            IReadOnlyList<PivotNamedSetDefinition> namedSets,
            IReadOnlyList<PivotNamedSetArtifactBinding> bindings,
            ValidationResult result)
        {
            if (bindings.Count != namedSets.Count)
            {
                result.AddError(
                    "PIVOT_SET_ARTIFACT_BINDING_COUNT_MISMATCH",
                    "artifactBindings",
                    "Every named-set definition requires exactly one generated artifact binding.");
            }

            var definitionIds = new HashSet<string>(
                namedSets.Where(value => value != null).Select(value => value.Id),
                StringComparer.OrdinalIgnoreCase);
            var bindingIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var generatedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < bindings.Count; index++)
            {
                PivotNamedSetArtifactBinding? binding = bindings[index];
                string path = "artifactBindings[" +
                    index.ToString(CultureInfo.InvariantCulture) + "]";
                if (binding == null)
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_BINDING_NULL",
                        path,
                        "Artifact binding entries cannot be null.");
                    continue;
                }

                ValidateId(
                    binding.DefinitionId,
                    path + ".definitionId",
                    "PIVOT_SET_ARTIFACT_DEFINITION_ID_INVALID",
                    result);
                if (!bindingIds.Add(binding.DefinitionId))
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_DEFINITION_ID_DUPLICATE",
                        path + ".definitionId",
                        "Artifact bindings must have unique definition IDs.");
                }

                if (!definitionIds.Contains(binding.DefinitionId))
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_DEFINITION_UNKNOWN",
                        path + ".definitionId",
                        "An artifact binding references an unknown named-set definition.");
                }

                if (!GeneratedSetNamePattern.IsMatch(binding.GeneratedSetName))
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_NAME_INVALID",
                        path + ".generatedSetName",
                        "A generated setup-namespaced single-segment MDX set name is required.");
                }
                else if (!generatedNames.Add(binding.GeneratedSetName))
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_NAME_DUPLICATE",
                        path + ".generatedSetName",
                        "Generated set names must be unique without regard to case.");
                }
            }

            foreach (string definitionId in definitionIds)
            {
                if (!bindingIds.Contains(definitionId))
                {
                    result.AddError(
                        "PIVOT_SET_ARTIFACT_BINDING_MISSING",
                        "artifactBindings",
                        "A named-set definition is missing its generated artifact binding.");
                }
            }
        }

        private static void ValidateCompiledFormulaBounds(
            PivotNamedSetCompilationRequest request,
            PivotNamedSetSchemaIndex index,
            ValidationResult result)
        {
            IReadOnlyDictionary<string, PivotNamedSetArtifactBinding> bindings =
                request.ArtifactBindings.ToDictionary(
                    item => item.DefinitionId,
                    StringComparer.OrdinalIgnoreCase);
            IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures =
                (request.DaxCompilation?.Measures ??
                 Array.Empty<OwnedPivotMeasureDefinition>())
                .ToDictionary(
                    item => item.DefinitionId,
                    StringComparer.OrdinalIgnoreCase);

            for (var indexValue = 0;
                 indexValue < request.Definition.NamedSets.Count;
                 indexValue++)
            {
                PivotNamedSetDefinition namedSet =
                    request.Definition.NamedSets[indexValue];
                string formula = PivotMdxCompiler.CompileFormulaUnchecked(
                    namedSet.Expression,
                    index,
                    measures);
                if (formula.Length > MaximumCompiledFormulaCharacters)
                {
                    result.AddError(
                        "PIVOT_SET_FORMULA_LIMIT",
                        "namedSets[" +
                        indexValue.ToString(CultureInfo.InvariantCulture) +
                        "].expression",
                        "The deterministic MDX formula exceeds the bounded formula limit.");
                }

                // Force exact binding lookup here so a future validator change
                // cannot let compilation proceed with an incomplete map.
                _ = bindings[namedSet.Id];
            }
        }

        private static void ValidateId(
            string value,
            string path,
            string code,
            ValidationResult result)
        {
            if (value == null || value.Length > MaximumIdLength ||
                !IdPattern.IsMatch(value))
            {
                result.AddError(
                    code,
                    path,
                    "A bounded path-free opaque identifier is required.");
            }
        }

        private static void ValidateFingerprint(
            string value,
            string path,
            string code,
            ValidationResult result)
        {
            if (value == null || !FingerprintPattern.IsMatch(value))
            {
                result.AddError(
                    code,
                    path,
                    "A canonical SHA-256 source fingerprint is required.");
            }
        }

        private static void ValidateCaption(
            string value,
            int maximumLength,
            string path,
            string code,
            ValidationResult result)
        {
            if (!IsBoundedPrintableText(value, maximumLength, allowEmpty: false))
            {
                result.AddError(code, path, "A bounded printable caption is required.");
            }
        }

        private static void ValidateOptionalCaption(
            string? value,
            int maximumLength,
            string path,
            string code,
            ValidationResult result)
        {
            if (value != null &&
                !IsBoundedPrintableText(value, maximumLength, allowEmpty: false))
            {
                result.AddError(code, path, "A bounded printable caption is required.");
            }
        }

        private static bool IsBoundedPrintableText(
            string? value,
            int maximumLength,
            bool allowEmpty)
        {
            return value != null &&
                   value.Length <= maximumLength &&
                   (allowEmpty || !string.IsNullOrWhiteSpace(value)) &&
                   string.Equals(value, value.Trim(), StringComparison.Ordinal) &&
                   !value.Any(char.IsControl);
        }

        private static void ValidateProviderUniqueName(
            string value,
            string path,
            string code,
            ValidationResult result)
        {
            if (!PivotMdxUniqueName.IsValid(value, MaximumProviderUniqueNameLength))
            {
                result.AddError(
                    code,
                    path,
                    "A bounded provider-issued MDX unique-name token is required.");
            }
        }
    }

    /// <summary>
    /// Validates provider-issued MDX unique names as identifier tokens, not
    /// expressions. It deliberately accepts provider-specific content inside
    /// brackets while rejecting functions, delimiters, and text outside the
    /// bracketed token grammar.
    /// </summary>
    internal static class PivotMdxUniqueName
    {
        public static bool IsValid(string? value, int maximumLength)
        {
            if (string.IsNullOrWhiteSpace(value) ||
                value!.Length > maximumLength ||
                !string.Equals(value, value.Trim(), StringComparison.Ordinal) ||
                value.Any(char.IsControl))
            {
                return false;
            }

            var index = 0;
            var segment = 0;
            while (index < value.Length)
            {
                if (segment > 0)
                {
                    char delimiter = value[index];
                    if (delimiter != '.' && delimiter != '&')
                    {
                        return false;
                    }

                    index++;
                    if (index >= value.Length)
                    {
                        return false;
                    }

                    if (delimiter == '.' && value[index] == '&')
                    {
                        index++;
                    }
                }

                if (index >= value.Length || value[index] != '[')
                {
                    return false;
                }

                index++;
                var contentCharacters = 0;
                var closed = false;
                while (index < value.Length)
                {
                    char current = value[index];
                    if (current != ']')
                    {
                        contentCharacters++;
                        index++;
                        continue;
                    }

                    if (index + 1 < value.Length && value[index + 1] == ']')
                    {
                        contentCharacters++;
                        index += 2;
                        continue;
                    }

                    index++;
                    closed = true;
                    break;
                }

                if (!closed || contentCharacters == 0)
                {
                    return false;
                }

                segment++;
            }

            return segment > 0;
        }
    }
}
