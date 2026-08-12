using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.PivotPlus.NamedSets
{
    public enum PivotNamedSetAxis
    {
        Unknown,
        Row,
        Column
    }

    public enum PivotNamedSetExpressionKind
    {
        ExplicitOrderedTuples,
        TopNLevelMembers
    }

    public enum PivotNamedSetProviderKind
    {
        Unknown,
        DataModel
    }

    public enum PivotNamedSetTupleMemberKind
    {
        CatalogMember,
        HierarchyDefaultMember
    }

    internal static class PivotNamedSetCollections
    {
        public static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }

    /// <summary>
    /// One exact, provider-issued member identity captured by the host. Action
    /// contracts reference Id only; ProviderUniqueName is transient compiler
    /// input and is never a model-authored MDX fragment.
    /// </summary>
    public sealed class PivotNamedSetMemberSchema
    {
        public PivotNamedSetMemberSchema(
            string id,
            string providerUniqueName,
            string? caption = null,
            string? parentMemberId = null,
            bool isAllMember = false)
        {
            Id = id ?? string.Empty;
            ProviderUniqueName = providerUniqueName ?? string.Empty;
            Caption = caption;
            ParentMemberId = parentMemberId;
            IsAllMember = isAllMember;
        }

        public string Id { get; }

        public string ProviderUniqueName { get; }

        public string? Caption { get; }

        public string? ParentMemberId { get; }

        public bool IsAllMember { get; }
    }

    /// <summary>
    /// One exact level identity and its bounded host-discovered members.
    /// MembersComplete means the host could establish an unambiguous catalog
    /// for the requested operation without creating PivotFields.
    /// </summary>
    public sealed class PivotNamedSetLevelSchema
    {
        public PivotNamedSetLevelSchema(
            string id,
            string providerUniqueName,
            int ordinal,
            bool membersComplete,
            IEnumerable<PivotNamedSetMemberSchema>? members)
        {
            Id = id ?? string.Empty;
            ProviderUniqueName = providerUniqueName ?? string.Empty;
            Ordinal = ordinal;
            MembersComplete = membersComplete;
            Members = PivotNamedSetCollections.Copy(members);
        }

        public string Id { get; }

        public string ProviderUniqueName { get; }

        public int Ordinal { get; }

        public bool MembersComplete { get; }

        public IReadOnlyList<PivotNamedSetMemberSchema> Members { get; }
    }

    /// <summary>
    /// One hierarchy visible through the selected Data Model PivotTable.
    /// AllMemberId is optional and is populated only when the host can prove
    /// one exact catalog member is the hierarchy's All member. It is distinct
    /// from the hierarchy DefaultMember semantic used by tuple definitions.
    /// </summary>
    public sealed class PivotNamedSetHierarchySchema
    {
        public PivotNamedSetHierarchySchema(
            string id,
            string providerUniqueName,
            bool identityComplete,
            IEnumerable<PivotNamedSetLevelSchema>? levels,
            string? caption = null)
            : this(
                id,
                providerUniqueName,
                null,
                identityComplete,
                levels,
                caption)
        {
        }

        public PivotNamedSetHierarchySchema(
            string id,
            string providerUniqueName,
            string? allMemberId,
            bool identityComplete,
            IEnumerable<PivotNamedSetLevelSchema>? levels,
            string? caption = null)
        {
            Id = id ?? string.Empty;
            ProviderUniqueName = providerUniqueName ?? string.Empty;
            AllMemberId = allMemberId;
            IdentityComplete = identityComplete;
            Levels = PivotNamedSetCollections.Copy(levels);
            Caption = caption;
        }

        public string Id { get; }

        public string ProviderUniqueName { get; }

        public string? AllMemberId { get; }

        public bool IdentityComplete { get; }

        public IReadOnlyList<PivotNamedSetLevelSchema> Levels { get; }

        public string? Caption { get; }
    }

    /// <summary>
    /// Bounded transient catalog produced from the exact selected Data Model
    /// source. It contains no connection string, workbook path, or cell data.
    /// </summary>
    public sealed class PivotNamedSetSchema
    {
        public PivotNamedSetSchema(
            string sourceFingerprint,
            PivotNamedSetProviderKind providerKind,
            IEnumerable<PivotNamedSetHierarchySchema>? hierarchies)
        {
            SourceFingerprint = sourceFingerprint ?? string.Empty;
            ProviderKind = providerKind;
            Hierarchies = PivotNamedSetCollections.Copy(hierarchies);
        }

        public string SourceFingerprint { get; }

        public PivotNamedSetProviderKind ProviderKind { get; }

        public IReadOnlyList<PivotNamedSetHierarchySchema> Hierarchies { get; }
    }

    /// <summary>
    /// Closed tuple-member reference base. References contain only trusted
    /// catalog IDs or hierarchy IDs; there is no raw-MDX alternative.
    /// </summary>
    public abstract class PivotNamedSetTupleMemberReference
    {
        internal PivotNamedSetTupleMemberReference()
        {
        }

        public abstract PivotNamedSetTupleMemberKind Kind { get; }
    }

    public sealed class PivotNamedSetCatalogMemberReference :
        PivotNamedSetTupleMemberReference
    {
        public PivotNamedSetCatalogMemberReference(string memberId)
        {
            MemberId = memberId ?? string.Empty;
        }

        public override PivotNamedSetTupleMemberKind Kind =>
            PivotNamedSetTupleMemberKind.CatalogMember;

        public string MemberId { get; }
    }

    public sealed class PivotNamedSetHierarchyDefaultMemberReference :
        PivotNamedSetTupleMemberReference
    {
        public PivotNamedSetHierarchyDefaultMemberReference(string hierarchyId)
        {
            HierarchyId = hierarchyId ?? string.Empty;
        }

        public override PivotNamedSetTupleMemberKind Kind =>
            PivotNamedSetTupleMemberKind.HierarchyDefaultMember;

        public string HierarchyId { get; }
    }

    public sealed class PivotNamedSetTuple
    {
        public PivotNamedSetTuple(
            IEnumerable<PivotNamedSetTupleMemberReference>? members)
        {
            Members = PivotNamedSetCollections.Copy(members);
        }

        public IReadOnlyList<PivotNamedSetTupleMemberReference> Members { get; }

        public static PivotNamedSetTuple FromCatalogMemberIds(
            IEnumerable<string>? memberIds)
        {
            return new PivotNamedSetTuple(
                (memberIds ?? Enumerable.Empty<string>())
                .Select(memberId =>
                    new PivotNamedSetCatalogMemberReference(memberId)));
        }
    }

    /// <summary>
    /// Closed named-set expression base. There is intentionally no raw-MDX or
    /// arbitrary function expression.
    /// </summary>
    public abstract class PivotNamedSetExpression
    {
        internal PivotNamedSetExpression()
        {
        }

        public abstract PivotNamedSetExpressionKind Kind { get; }
    }

    /// <summary>
    /// Exact tuple order for a row or column axis. Asymmetric/scoped branches
    /// are represented explicitly: for example, a parent-total tuple followed
    /// by only the desired child tuples for that parent.
    /// </summary>
    public sealed class PivotExplicitOrderedTuplesExpression : PivotNamedSetExpression
    {
        public PivotExplicitOrderedTuplesExpression(
            IEnumerable<string>? hierarchyIds,
            IEnumerable<PivotNamedSetTuple>? tuples)
        {
            HierarchyIds = PivotNamedSetCollections.Copy(hierarchyIds);
            Tuples = PivotNamedSetCollections.Copy(tuples);
        }

        public override PivotNamedSetExpressionKind Kind =>
            PivotNamedSetExpressionKind.ExplicitOrderedTuples;

        public IReadOnlyList<string> HierarchyIds { get; }

        public IReadOnlyList<PivotNamedSetTuple> Tuples { get; }
    }

    /// <summary>
    /// Typed Top N over one exact level using one measure authored in the same
    /// validated DAX compilation. The compiler emits only TopCount.
    /// </summary>
    public sealed class PivotTopNLevelMembersExpression : PivotNamedSetExpression
    {
        public PivotTopNLevelMembersExpression(
            string levelId,
            int count,
            string measureDefinitionId)
        {
            LevelId = levelId ?? string.Empty;
            Count = count;
            MeasureDefinitionId = measureDefinitionId ?? string.Empty;
        }

        public override PivotNamedSetExpressionKind Kind =>
            PivotNamedSetExpressionKind.TopNLevelMembers;

        public string LevelId { get; }

        public int Count { get; }

        public string MeasureDefinitionId { get; }
    }

    public sealed class PivotNamedSetDefinition
    {
        public PivotNamedSetDefinition(
            string id,
            string caption,
            PivotNamedSetAxis axis,
            PivotNamedSetExpression expression,
            bool flattenHierarchies = false)
        {
            Id = id ?? string.Empty;
            Caption = caption ?? string.Empty;
            Axis = axis;
            Expression = expression ?? throw new ArgumentNullException(nameof(expression));
            FlattenHierarchies = flattenHierarchies;
        }

        public string Id { get; }

        public string Caption { get; }

        public PivotNamedSetAxis Axis { get; }

        public PivotNamedSetExpression Expression { get; }

        public bool FlattenHierarchies { get; }
    }

    /// <summary>
    /// Complete phase-one named-set request for one exact source snapshot.
    /// At most one named set can own each native axis.
    /// </summary>
    public sealed class PivotNamedSetCollectionDefinition
    {
        public PivotNamedSetCollectionDefinition(
            string sourceFingerprint,
            PivotNamedSetSchema schema,
            IEnumerable<PivotNamedSetDefinition>? namedSets)
        {
            SourceFingerprint = sourceFingerprint ?? string.Empty;
            Schema = schema ?? throw new ArgumentNullException(nameof(schema));
            NamedSets = PivotNamedSetCollections.Copy(namedSets);
        }

        public string SourceFingerprint { get; }

        public PivotNamedSetSchema Schema { get; }

        public IReadOnlyList<PivotNamedSetDefinition> NamedSets { get; }
    }

    /// <summary>
    /// Maps a semantic definition ID to a setup-namespaced native set name.
    /// The native name is generated by the trusted host and is separate from
    /// the user-visible caption.
    /// </summary>
    public sealed class PivotNamedSetArtifactBinding
    {
        public PivotNamedSetArtifactBinding(string definitionId, string generatedSetName)
        {
            DefinitionId = definitionId ?? string.Empty;
            GeneratedSetName = generatedSetName ?? string.Empty;
        }

        public string DefinitionId { get; }

        public string GeneratedSetName { get; }
    }

    public sealed class PivotNamedSetCompilationRequest
    {
        public PivotNamedSetCompilationRequest(
            PivotNamedSetCollectionDefinition definition,
            IEnumerable<PivotNamedSetArtifactBinding>? artifactBindings,
            PivotDaxCompilation? daxCompilation = null)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            ArtifactBindings = PivotNamedSetCollections.Copy(artifactBindings);
            DaxCompilation = daxCompilation;
        }

        public PivotNamedSetCollectionDefinition Definition { get; }

        public IReadOnlyList<PivotNamedSetArtifactBinding> ArtifactBindings { get; }

        public PivotDaxCompilation? DaxCompilation { get; }
    }

    /// <summary>
    /// Transient compiler output. Workbook ownership persists fingerprints,
    /// never MdxFormula.
    /// </summary>
    public sealed class PivotNamedSetMeasureDependencyBinding
    {
        internal PivotNamedSetMeasureDependencyBinding(
            string definitionId,
            string generatedMeasureName,
            string measureDefinitionFingerprint,
            string measureFormulaFingerprint)
        {
            DefinitionId = definitionId;
            GeneratedMeasureName = generatedMeasureName;
            MeasureDefinitionFingerprint = measureDefinitionFingerprint;
            MeasureFormulaFingerprint = measureFormulaFingerprint;
        }

        public string DefinitionId { get; }

        public string GeneratedMeasureName { get; }

        public string MeasureDefinitionFingerprint { get; }

        public string MeasureFormulaFingerprint { get; }
    }

    /// <summary>
    /// Transient compiler output. Workbook ownership persists fingerprints,
    /// never MdxFormula or dependency bindings.
    /// </summary>
    public sealed class OwnedPivotNamedSetDefinition
    {
        internal OwnedPivotNamedSetDefinition(
            string definitionId,
            int displayOrder,
            string generatedSetName,
            string caption,
            PivotNamedSetAxis axis,
            string mdxFormula,
            bool dynamic,
            bool flattenHierarchies,
            IEnumerable<PivotNamedSetMeasureDependencyBinding>? directMeasureDependencies,
            string definitionFingerprint,
            string formulaFingerprint)
        {
            DefinitionId = definitionId;
            DisplayOrder = displayOrder;
            GeneratedSetName = generatedSetName;
            Caption = caption;
            Axis = axis;
            MdxFormula = mdxFormula;
            Dynamic = dynamic;
            FlattenHierarchies = flattenHierarchies;
            DirectMeasureDependencies = PivotNamedSetCollections.Copy(
                directMeasureDependencies);
            DirectMeasureDependencyDefinitionIds = PivotNamedSetCollections.Copy(
                DirectMeasureDependencies.Select(binding => binding.DefinitionId));
            DefinitionFingerprint = definitionFingerprint;
            FormulaFingerprint = formulaFingerprint;
        }

        public string DefinitionId { get; }

        public int DisplayOrder { get; }

        public string GeneratedSetName { get; }

        public string Caption { get; }

        public PivotNamedSetAxis Axis { get; }

        public string MdxFormula { get; }

        public bool Dynamic { get; }

        public bool FlattenHierarchies { get; }

        /// <summary>
        /// Exact custom order requires this to remain false. Duplicate tuples
        /// are rejected by validation rather than silently removed by Excel.
        /// </summary>
        public bool HierarchizeDistinct => false;

        /// <summary>
        /// Exact compiler-issued identity of every directly referenced model
        /// measure. The binding lets the host reject a stale or mixed DAX/MDX
        /// pair without reading or parsing either formula.
        /// </summary>
        public IReadOnlyList<PivotNamedSetMeasureDependencyBinding>
            DirectMeasureDependencies { get; }

        /// <summary>
        /// Compatibility projection for callers that need only semantic IDs.
        /// New compatibility checks must use DirectMeasureDependencies.
        /// </summary>
        public IReadOnlyList<string> DirectMeasureDependencyDefinitionIds { get; }

        public string DefinitionFingerprint { get; }

        public string FormulaFingerprint { get; }
    }

    public sealed class PivotMdxCompilation
    {
        internal PivotMdxCompilation(
            string sourceFingerprint,
            IEnumerable<OwnedPivotNamedSetDefinition>? namedSets)
        {
            SourceFingerprint = sourceFingerprint;
            NamedSets = PivotNamedSetCollections.Copy(namedSets);
            CompilationFingerprint = PivotMdxFingerprint.ComputeCompilation(
                SourceFingerprint,
                NamedSets);
        }

        public string SourceFingerprint { get; }

        public IReadOnlyList<OwnedPivotNamedSetDefinition> NamedSets { get; }

        /// <summary>
        /// Deterministic identity of the exact source-bound compiled set
        /// collection. It contains hashes and identifiers, never raw MDX.
        /// </summary>
        public string CompilationFingerprint { get; }

        public bool HasExactMeasureDependencies(PivotDaxCompilation? daxCompilation)
        {
            IReadOnlyList<PivotNamedSetMeasureDependencyBinding> dependencies =
                NamedSets.SelectMany(set => set.DirectMeasureDependencies).ToList();
            if (dependencies.Count == 0) return true;
            if (daxCompilation == null) return false;

            var measures = daxCompilation.Measures
                .GroupBy(measure => measure.DefinitionId, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() == 1)
                .ToDictionary(
                    group => group.Key,
                    group => group.Single(),
                    StringComparer.OrdinalIgnoreCase);
            return dependencies.All(binding =>
                measures.TryGetValue(
                    binding.DefinitionId,
                    out OwnedPivotMeasureDefinition? measure) &&
                string.Equals(
                    binding.DefinitionId,
                    measure.DefinitionId,
                    StringComparison.Ordinal) &&
                string.Equals(
                    binding.GeneratedMeasureName,
                    measure.GeneratedMeasureName,
                    StringComparison.Ordinal) &&
                string.Equals(
                    binding.MeasureDefinitionFingerprint,
                    measure.DefinitionFingerprint,
                    StringComparison.Ordinal) &&
                string.Equals(
                    binding.MeasureFormulaFingerprint,
                    measure.FormulaFingerprint,
                    StringComparison.Ordinal));
        }
    }

    public sealed class InvalidPivotNamedSetException : Exception
    {
        public InvalidPivotNamedSetException(ValidationResult validation)
            : base("The PivotTable+ named-set definition is invalid.")
        {
            Validation = validation ?? throw new ArgumentNullException(nameof(validation));
        }

        public ValidationResult Validation { get; }
    }
}
