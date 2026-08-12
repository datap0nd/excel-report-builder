using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.NamedSets
{
    /// <summary>
    /// Deterministically compiles the closed PivotTable+ named-set union. The
    /// compiler accepts only host-bound schema IDs, explicit tuples, and typed
    /// TopCount; it has no arbitrary MDX entry point.
    /// </summary>
    public static class PivotMdxCompiler
    {
        public static PivotMdxCompilation Compile(PivotNamedSetCompilationRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            ValidationResult validation = PivotNamedSetValidator.Validate(request);
            if (!validation.IsValid)
            {
                throw new InvalidPivotNamedSetException(validation);
            }

            var schema = new PivotNamedSetSchemaIndex(request.Definition.Schema);
            IReadOnlyDictionary<string, PivotNamedSetArtifactBinding> bindings =
                request.ArtifactBindings.ToDictionary(
                    binding => binding.DefinitionId,
                    StringComparer.OrdinalIgnoreCase);
            IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures =
                (request.DaxCompilation?.Measures ??
                 Array.Empty<OwnedPivotMeasureDefinition>())
                .ToDictionary(
                    measure => measure.DefinitionId,
                    StringComparer.OrdinalIgnoreCase);

            var compiled = new List<OwnedPivotNamedSetDefinition>(
                request.Definition.NamedSets.Count);
            for (var index = 0;
                 index < request.Definition.NamedSets.Count;
                 index++)
            {
                PivotNamedSetDefinition definition =
                    request.Definition.NamedSets[index];
                PivotNamedSetArtifactBinding binding = bindings[definition.Id];
                string formula = CompileFormulaUnchecked(
                    definition.Expression,
                    schema,
                    measures);
                if (formula.Length > PivotNamedSetValidator.MaximumCompiledFormulaCharacters)
                {
                    throw new InvalidOperationException(
                        "Validated named-set MDX exceeded the bounded compiler limit.");
                }

                IReadOnlyList<PivotNamedSetMeasureDependencyBinding> dependencies =
                    DirectMeasureDependencies(
                        definition.Expression,
                        measures);
                string canonical = CanonicalDefinition(
                    request.Definition.SourceFingerprint,
                    definition,
                    binding,
                    index + 1,
                    measures);
                compiled.Add(new OwnedPivotNamedSetDefinition(
                    definition.Id,
                    index + 1,
                    binding.GeneratedSetName,
                    definition.Caption,
                    definition.Axis,
                    formula,
                    dynamic: definition.Expression is PivotTopNLevelMembersExpression,
                    definition.FlattenHierarchies,
                    dependencies,
                    PivotMdxFingerprint.ComputeDefinition(canonical),
                    PivotMdxFingerprint.ComputeFormula(formula)));
            }

            return new PivotMdxCompilation(
                request.Definition.SourceFingerprint,
                compiled);
        }

        internal static string CompileFormulaUnchecked(
            PivotNamedSetExpression expression,
            PivotNamedSetSchemaIndex schema,
            IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures)
        {
            switch (expression)
            {
                case PivotExplicitOrderedTuplesExpression explicitTuples:
                    return CompileExplicitTuples(explicitTuples, schema);
                case PivotTopNLevelMembersExpression topN:
                    return CompileTopN(topN, schema, measures);
                default:
                    throw new InvalidOperationException(
                        "Validated named-set expression kind was lost.");
            }
        }

        private static string CompileExplicitTuples(
            PivotExplicitOrderedTuplesExpression expression,
            PivotNamedSetSchemaIndex schema)
        {
            var tuples = new List<string>(expression.Tuples.Count);
            foreach (PivotNamedSetTuple tuple in expression.Tuples)
            {
                string[] members = tuple.Members
                    .Select(reference => CompileTupleMember(reference, schema))
                    .ToArray();

                tuples.Add(members.Length == 1
                    ? members[0]
                    : "(" + string.Join(", ", members) + ")");
            }

            return "{" + string.Join(", ", tuples) + "}";
        }

        private static string CompileTupleMember(
            PivotNamedSetTupleMemberReference reference,
            PivotNamedSetSchemaIndex schema)
        {
            switch (reference)
            {
                case PivotNamedSetCatalogMemberReference catalogMember:
                    if (!schema.TryGetMember(
                            catalogMember.MemberId,
                            out PivotNamedSetBoundMember member))
                    {
                        throw new InvalidOperationException(
                            "Validated named-set member binding was lost.");
                    }

                    return member.Member.ProviderUniqueName;
                case PivotNamedSetHierarchyDefaultMemberReference defaultMember:
                    if (!schema.TryGetHierarchy(
                            defaultMember.HierarchyId,
                            out PivotNamedSetHierarchySchema hierarchy))
                    {
                        throw new InvalidOperationException(
                            "Validated named-set hierarchy binding was lost.");
                    }

                    return hierarchy.ProviderUniqueName + ".DefaultMember";
                default:
                    throw new InvalidOperationException(
                        "Validated tuple-member reference kind was lost.");
            }
        }

        private static string CompileTopN(
            PivotTopNLevelMembersExpression expression,
            PivotNamedSetSchemaIndex schema,
            IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures)
        {
            if (!schema.TryGetLevel(
                    expression.LevelId,
                    out PivotNamedSetBoundLevel level) ||
                !measures.TryGetValue(
                    expression.MeasureDefinitionId,
                    out OwnedPivotMeasureDefinition? measure))
            {
                throw new InvalidOperationException(
                    "Validated Top N binding was lost.");
            }

            return "TopCount(" + level.Level.ProviderUniqueName + ".Members, " +
                   expression.Count.ToString(CultureInfo.InvariantCulture) + ", " +
                   MeasureUniqueName(measure.GeneratedMeasureName) + ")";
        }

        private static IReadOnlyList<PivotNamedSetMeasureDependencyBinding>
            DirectMeasureDependencies(
                PivotNamedSetExpression expression,
                IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures)
        {
            if (expression is PivotTopNLevelMembersExpression topN)
            {
                OwnedPivotMeasureDefinition measure = measures[topN.MeasureDefinitionId];
                return new[]
                {
                    new PivotNamedSetMeasureDependencyBinding(
                        measure.DefinitionId,
                        measure.GeneratedMeasureName,
                        measure.DefinitionFingerprint,
                        measure.FormulaFingerprint)
                };
            }

            return Array.Empty<PivotNamedSetMeasureDependencyBinding>();
        }

        private static string CanonicalDefinition(
            string sourceFingerprint,
            PivotNamedSetDefinition definition,
            PivotNamedSetArtifactBinding binding,
            int displayOrder,
            IReadOnlyDictionary<string, OwnedPivotMeasureDefinition> measures)
        {
            var writer = new PivotNamedSetCanonicalWriter();
            writer.Add("sourceFingerprint", sourceFingerprint);
            writer.Add("definitionId", definition.Id);
            writer.Add("displayOrder", displayOrder.ToString(CultureInfo.InvariantCulture));
            writer.Add("generatedSetName", binding.GeneratedSetName);
            writer.Add("caption", definition.Caption);
            writer.Add("axis", ((int)definition.Axis).ToString(CultureInfo.InvariantCulture));
            writer.Add("flattenHierarchies", definition.FlattenHierarchies ? "1" : "0");
            writer.Add("hierarchizeDistinct", "0");
            writer.Add(
                "expressionKind",
                ((int)definition.Expression.Kind).ToString(CultureInfo.InvariantCulture));

            switch (definition.Expression)
            {
                case PivotExplicitOrderedTuplesExpression explicitTuples:
                    writer.Add(
                        "hierarchyCount",
                        explicitTuples.HierarchyIds.Count.ToString(CultureInfo.InvariantCulture));
                    for (var hierarchyIndex = 0;
                         hierarchyIndex < explicitTuples.HierarchyIds.Count;
                         hierarchyIndex++)
                    {
                        writer.Add(
                            "hierarchy" + hierarchyIndex.ToString(CultureInfo.InvariantCulture),
                            explicitTuples.HierarchyIds[hierarchyIndex]);
                    }

                    writer.Add(
                        "tupleCount",
                        explicitTuples.Tuples.Count.ToString(CultureInfo.InvariantCulture));
                    for (var tupleIndex = 0;
                         tupleIndex < explicitTuples.Tuples.Count;
                         tupleIndex++)
                    {
                        PivotNamedSetTuple tuple = explicitTuples.Tuples[tupleIndex];
                        for (var memberIndex = 0;
                             memberIndex < tuple.Members.Count;
                             memberIndex++)
                        {
                            PivotNamedSetTupleMemberReference reference =
                                tuple.Members[memberIndex];
                            writer.Add(
                                "tuple" + tupleIndex.ToString(CultureInfo.InvariantCulture) +
                                "member" + memberIndex.ToString(CultureInfo.InvariantCulture) +
                                "Kind",
                                ((int)reference.Kind).ToString(CultureInfo.InvariantCulture));
                            writer.Add(
                                "tuple" + tupleIndex.ToString(CultureInfo.InvariantCulture) +
                                "member" + memberIndex.ToString(CultureInfo.InvariantCulture),
                                TupleMemberReferenceId(reference));
                        }
                    }

                    break;
                case PivotTopNLevelMembersExpression topN:
                    writer.Add("levelId", topN.LevelId);
                    writer.Add("count", topN.Count.ToString(CultureInfo.InvariantCulture));
                    writer.Add("measureDefinitionId", topN.MeasureDefinitionId);
                    writer.Add(
                        "measureDefinitionFingerprint",
                        measures[topN.MeasureDefinitionId].DefinitionFingerprint);
                    break;
                default:
                    throw new InvalidOperationException(
                        "Validated named-set expression kind was lost.");
            }

            return writer.ToString();
        }

        private static string TupleMemberReferenceId(
            PivotNamedSetTupleMemberReference reference)
        {
            switch (reference)
            {
                case PivotNamedSetCatalogMemberReference catalogMember:
                    return catalogMember.MemberId;
                case PivotNamedSetHierarchyDefaultMemberReference defaultMember:
                    return defaultMember.HierarchyId;
                default:
                    throw new InvalidOperationException(
                        "Validated tuple-member reference kind was lost.");
            }
        }

        private static string MeasureUniqueName(string measureName)
        {
            return "[Measures].[" + measureName.Replace("]", "]]" ) + "]";
        }
    }
}
