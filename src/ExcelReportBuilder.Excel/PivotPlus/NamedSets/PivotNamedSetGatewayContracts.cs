using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.PivotPlus;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.Measures;
using ExcelReportBuilder.Excel.PivotPlus.Persistence;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    internal enum PivotNamedSetPairState
    {
        Complete,
        CalculatedMemberOnly,
        CubeFieldOnly
    }

    internal sealed class PivotNamedSetDiscoveryDiagnostic
    {
        public PivotNamedSetDiscoveryDiagnostic(string code, string path, string message)
        {
            Code = code ?? string.Empty;
            Path = path ?? string.Empty;
            Message = message ?? string.Empty;
        }

        public string Code { get; }

        public string Path { get; }

        public string Message { get; }
    }

    internal sealed class PivotNamedSetSchemaDiscoveryResult
    {
        public PivotNamedSetSchemaDiscoveryResult(
            PivotNamedSetSchema schema,
            IEnumerable<PivotNamedSetDiscoveryDiagnostic>? diagnostics)
        {
            Schema = schema ?? throw new ArgumentNullException(nameof(schema));
            Diagnostics = Copy(diagnostics);
        }

        public PivotNamedSetSchema Schema { get; }

        public IReadOnlyList<PivotNamedSetDiscoveryDiagnostic> Diagnostics { get; }

        private static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }

    /// <summary>
    /// Internal host DTO created only by PivotNamedSetCompilationAdapter. Its
    /// RawMdx is transient trusted Core compiler output, never model input.
    /// </summary>
    internal sealed class DesiredPivotNamedSet
    {
        internal DesiredPivotNamedSet(
            string sourceFingerprint,
            string compilationFingerprint,
            string definitionId,
            int displayOrder,
            string name,
            string caption,
            PivotNamedSetAxis axis,
            string rawMdx,
            bool dynamic,
            bool flattenHierarchies,
            bool hierarchizeDistinct,
            IEnumerable<DesiredPivotNamedSetMeasureDependency>? directMeasureDependencies,
            string definitionFingerprint,
            string formulaFingerprint,
            string displayFolderMarker)
        {
            SourceFingerprint = sourceFingerprint;
            CompilationFingerprint = compilationFingerprint;
            DefinitionId = definitionId;
            DisplayOrder = displayOrder;
            Name = name;
            Caption = caption;
            Axis = axis;
            RawMdx = rawMdx;
            Dynamic = dynamic;
            FlattenHierarchies = flattenHierarchies;
            HierarchizeDistinct = hierarchizeDistinct;
            DirectMeasureDependencies = new ReadOnlyCollection<DesiredPivotNamedSetMeasureDependency>(
                (directMeasureDependencies ??
                 Enumerable.Empty<DesiredPivotNamedSetMeasureDependency>()).ToList());
            DefinitionFingerprint = definitionFingerprint;
            FormulaFingerprint = formulaFingerprint;
            DisplayFolderMarker = displayFolderMarker;
        }

        public string SourceFingerprint { get; }

        public string CompilationFingerprint { get; }

        public string DefinitionId { get; }

        public int DisplayOrder { get; }

        public string Name { get; }

        public string Caption { get; }

        public PivotNamedSetAxis Axis { get; }

        internal string RawMdx { get; }

        public bool Dynamic { get; }

        public bool FlattenHierarchies { get; }

        public bool HierarchizeDistinct { get; }

        public IReadOnlyList<DesiredPivotNamedSetMeasureDependency> DirectMeasureDependencies
        {
            get;
        }

        public IReadOnlyList<string> DirectMeasureDependencyDefinitionIds =>
            new ReadOnlyCollection<string>(
                DirectMeasureDependencies.Select(binding => binding.DefinitionId).ToList());

        public string DefinitionFingerprint { get; }

        public string FormulaFingerprint { get; }

        public string DisplayFolderMarker { get; }
    }

    internal sealed class DesiredPivotNamedSetMeasureDependency
    {
        public DesiredPivotNamedSetMeasureDependency(
            string definitionId,
            string generatedMeasureName,
            string measureDefinitionFingerprint,
            string measureFormulaFingerprint,
            string expectedDescriptionMarker)
        {
            DefinitionId = definitionId;
            GeneratedMeasureName = generatedMeasureName;
            MeasureDefinitionFingerprint = measureDefinitionFingerprint;
            MeasureFormulaFingerprint = measureFormulaFingerprint;
            ExpectedDescriptionMarker = expectedDescriptionMarker;
        }

        public string DefinitionId { get; }
        public string GeneratedMeasureName { get; }
        public string MeasureDefinitionFingerprint { get; }
        public string MeasureFormulaFingerprint { get; }
        public string ExpectedDescriptionMarker { get; }
    }

    internal sealed class LivePivotNamedSetSnapshot
    {
        public LivePivotNamedSetSnapshot(
            string worksheetName,
            string pivotTableName,
            bool isSelectedTarget,
            string name,
            PivotNamedSetPairState pairState,
            string rawFormula,
            string formulaFingerprint,
            string displayFolder,
            string sourceName,
            string caption,
            int? calculatedMemberType,
            int? cubeFieldType,
            bool? dynamic,
            bool? calculatedMemberFlattenHierarchies,
            bool? cubeFieldFlattenHierarchies,
            bool? calculatedMemberHierarchizeDistinct,
            bool? cubeFieldHierarchizeDistinct,
            bool? showInFieldList,
            int? orientation,
            bool? isValid,
            string sourceFingerprint,
            string modelLineageFingerprint,
            string liveFingerprint)
        {
            WorksheetName = worksheetName;
            PivotTableName = pivotTableName;
            IsSelectedTarget = isSelectedTarget;
            Name = name;
            PairState = pairState;
            RawFormula = rawFormula;
            FormulaFingerprint = formulaFingerprint;
            DisplayFolder = displayFolder;
            SourceName = sourceName;
            Caption = caption;
            CalculatedMemberType = calculatedMemberType;
            CubeFieldType = cubeFieldType;
            Dynamic = dynamic;
            CalculatedMemberFlattenHierarchies = calculatedMemberFlattenHierarchies;
            CubeFieldFlattenHierarchies = cubeFieldFlattenHierarchies;
            CalculatedMemberHierarchizeDistinct = calculatedMemberHierarchizeDistinct;
            CubeFieldHierarchizeDistinct = cubeFieldHierarchizeDistinct;
            ShowInFieldList = showInFieldList;
            Orientation = orientation;
            IsValid = isValid;
            SourceFingerprint = sourceFingerprint;
            ModelLineageFingerprint = modelLineageFingerprint;
            LiveFingerprint = liveFingerprint;
        }

        public string WorksheetName { get; }
        public string PivotTableName { get; }
        public bool IsSelectedTarget { get; }
        public string Name { get; }
        public PivotNamedSetPairState PairState { get; }
        internal string RawFormula { get; }
        public string FormulaFingerprint { get; }
        public string DisplayFolder { get; }
        public string SourceName { get; }
        public string Caption { get; }
        public int? CalculatedMemberType { get; }
        public int? CubeFieldType { get; }
        public bool? Dynamic { get; }
        public bool? CalculatedMemberFlattenHierarchies { get; }
        public bool? CubeFieldFlattenHierarchies { get; }
        public bool? CalculatedMemberHierarchizeDistinct { get; }
        public bool? CubeFieldHierarchizeDistinct { get; }
        public bool? ShowInFieldList { get; }
        public int? Orientation { get; }
        public bool? IsValid { get; }
        public string SourceFingerprint { get; }
        public string ModelLineageFingerprint { get; }
        public string LiveFingerprint { get; }

        public bool IsVisible => Orientation.HasValue && Orientation.Value != 0;
    }

    internal sealed class PivotCalculatedMemberReferenceSnapshot
    {
        public PivotCalculatedMemberReferenceSnapshot(
            string worksheetName,
            string pivotTableName,
            string name,
            int type,
            string rawFormula,
            bool formulaScanComplete)
        {
            WorksheetName = worksheetName;
            PivotTableName = pivotTableName;
            Name = name;
            Type = type;
            RawFormula = rawFormula;
            FormulaScanComplete = formulaScanComplete;
        }

        public string WorksheetName { get; }
        public string PivotTableName { get; }
        public string Name { get; }
        public int Type { get; }
        internal string RawFormula { get; }
        public bool FormulaScanComplete { get; }
    }

    internal sealed class PivotNamedSetPivotSnapshot
    {
        public PivotNamedSetPivotSnapshot(
            string worksheetName,
            string pivotTableName,
            bool isSelectedTarget,
            IEnumerable<LivePivotNamedSetSnapshot>? artifacts,
            IEnumerable<PivotCalculatedMemberReferenceSnapshot>? calculatedMembers,
            bool connectionRefreshed,
            string fingerprint)
        {
            WorksheetName = worksheetName;
            PivotTableName = pivotTableName;
            IsSelectedTarget = isSelectedTarget;
            Artifacts = Copy(artifacts);
            CalculatedMembers = Copy(calculatedMembers);
            ConnectionRefreshed = connectionRefreshed;
            Fingerprint = fingerprint;
        }

        public string WorksheetName { get; }
        public string PivotTableName { get; }
        public bool IsSelectedTarget { get; }
        public IReadOnlyList<LivePivotNamedSetSnapshot> Artifacts { get; }
        public IReadOnlyList<PivotCalculatedMemberReferenceSnapshot> CalculatedMembers { get; }
        public bool ConnectionRefreshed { get; }
        public string Fingerprint { get; }

        private static IReadOnlyList<T> Copy<T>(IEnumerable<T>? values)
        {
            return new ReadOnlyCollection<T>(
                (values ?? Enumerable.Empty<T>()).ToList());
        }
    }

    internal sealed class PivotNamedSetWorkbookSnapshot
    {
        public PivotNamedSetWorkbookSnapshot(
            IEnumerable<PivotNamedSetPivotSnapshot>? pivots,
            string sourceFingerprint,
            string modelLineageFingerprint)
        {
            Pivots = new ReadOnlyCollection<PivotNamedSetPivotSnapshot>(
                (pivots ?? Enumerable.Empty<PivotNamedSetPivotSnapshot>()).ToList());
            SourceFingerprint = sourceFingerprint;
            ModelLineageFingerprint = modelLineageFingerprint;
        }

        public IReadOnlyList<PivotNamedSetPivotSnapshot> Pivots { get; }

        public string SourceFingerprint { get; }

        public string ModelLineageFingerprint { get; }

        public PivotNamedSetPivotSnapshot SelectedPivot => Pivots.Single(
            pivot => pivot.IsSelectedTarget);

        public IReadOnlyList<LivePivotNamedSetSnapshot> Artifacts =>
            new ReadOnlyCollection<LivePivotNamedSetSnapshot>(
                Pivots.SelectMany(pivot => pivot.Artifacts).ToList());
    }

    internal sealed class BoundPivotNamedSetTarget
    {
        public BoundPivotNamedSetTarget(
            object workbook,
            object pivotTable,
            object model,
            object dataModelConnection,
            PivotTargetIdentity identity)
        {
            Workbook = workbook;
            PivotTable = pivotTable;
            Model = model;
            DataModelConnection = dataModelConnection;
            Identity = identity;
        }

        public object Workbook { get; }
        public object PivotTable { get; }
        public object Model { get; }
        public object DataModelConnection { get; }
        public PivotTargetIdentity Identity { get; }
    }

    internal interface IPivotNamedSetGateway
    {
        BoundPivotNamedSetTarget Bind(
            object workbook,
            object pivotTable,
            PivotTableContext context);

        PivotNamedSetSchemaDiscoveryResult DiscoverSchema(
            BoundPivotNamedSetTarget target);

        PivotNamedSetWorkbookSnapshot Capture(BoundPivotNamedSetTarget target);

        LivePivotNamedSetSnapshot CreateSet(
            BoundPivotNamedSetTarget target,
            DesiredPivotNamedSet definition);

        LivePivotNamedSetSnapshot ReplaceSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before,
            DesiredPivotNamedSet definition);

        LivePivotNamedSetSnapshot RestoreSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot before);

        void DeleteSet(
            BoundPivotNamedSetTarget target,
            LivePivotNamedSetSnapshot expected);
    }

    internal sealed class PivotNamedSetRecoveryRequiredException : InvalidOperationException
    {
        public PivotNamedSetRecoveryRequiredException(string message, Exception? inner = null)
            : base(message, inner)
        {
        }
    }

    internal static class PivotNamedSetCompilationAdapter
    {
        private const int MaximumNamedSets = 2;
        private const int MaximumFormulaCharacters = 24 * 1024;
        private static readonly Regex GeneratedNamePattern = new Regex(
            "^\\[[A-Za-z0-9][A-Za-z0-9._-]{0,126}\\]$",
            RegexOptions.CultureInvariant);

        public static IReadOnlyList<DesiredPivotNamedSet> CreateDesired(
            string setupId,
            PivotMdxCompilation compilation)
        {
            if (compilation == null) throw new ArgumentNullException(nameof(compilation));
            PivotPlusMetadataValidator.ValidateId(setupId, "setup identifier");
            PivotPlusMetadataValidator.ValidateFingerprint(
                compilation.SourceFingerprint,
                "named-set source fingerprint");
            PivotPlusMetadataValidator.ValidateFingerprint(
                compilation.CompilationFingerprint,
                "named-set compilation fingerprint");
            if (compilation.NamedSets == null ||
                compilation.NamedSets.Count == 0 ||
                compilation.NamedSets.Count > MaximumNamedSets)
            {
                throw new ArgumentException(
                    "The trusted named-set compilation is empty or exceeds its bounded limit.",
                    nameof(compilation));
            }

            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var displayOrders = new HashSet<int>();
            var result = new List<DesiredPivotNamedSet>();
            foreach (OwnedPivotNamedSetDefinition set in compilation.NamedSets)
            {
                if (set == null)
                {
                    throw new ArgumentException(
                        "A trusted named-set compilation entry cannot be null.",
                        nameof(compilation));
                }

                PivotPlusMetadataValidator.ValidateId(
                    set.DefinitionId,
                    "named-set definition identifier");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    set.DefinitionFingerprint,
                    "named-set definition fingerprint");
                PivotPlusMetadataValidator.ValidateFingerprint(
                    set.FormulaFingerprint,
                    "named-set formula fingerprint");
                if (!GeneratedNamePattern.IsMatch(set.GeneratedSetName) ||
                    !ids.Add(set.DefinitionId) ||
                    !names.Add(set.GeneratedSetName) ||
                    !displayOrders.Add(set.DisplayOrder) ||
                    set.DisplayOrder < 1 ||
                    set.DisplayOrder > compilation.NamedSets.Count ||
                    (set.Axis != PivotNamedSetAxis.Row &&
                     set.Axis != PivotNamedSetAxis.Column) ||
                    string.IsNullOrWhiteSpace(set.Caption) ||
                    set.Caption.Length > 255 ||
                    set.Caption.Any(char.IsControl) ||
                    string.IsNullOrWhiteSpace(set.MdxFormula) ||
                    set.MdxFormula.Length > MaximumFormulaCharacters ||
                    set.MdxFormula.Any(char.IsControl) ||
                    set.HierarchizeDistinct ||
                    !string.Equals(
                        PivotMdxFingerprint.ComputeFormula(set.MdxFormula),
                        set.FormulaFingerprint,
                        StringComparison.Ordinal))
                {
                    throw new ArgumentException(
                        "The trusted named-set compilation is internally inconsistent.",
                        nameof(compilation));
                }

                var dependencies = new List<DesiredPivotNamedSetMeasureDependency>();
                var dependencyIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var dependencyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (PivotNamedSetMeasureDependencyBinding dependency in
                         set.DirectMeasureDependencies)
                {
                    PivotPlusMetadataValidator.ValidateId(
                        dependency.DefinitionId,
                        "named-set measure dependency identifier");
                    PivotPlusMetadataValidator.ValidateArtifactName(
                        dependency.GeneratedMeasureName);
                    PivotPlusMetadataValidator.ValidateFingerprint(
                        dependency.MeasureDefinitionFingerprint,
                        "named-set dependency definition fingerprint");
                    PivotPlusMetadataValidator.ValidateFingerprint(
                        dependency.MeasureFormulaFingerprint,
                        "named-set dependency formula fingerprint");
                    if (!dependencyIds.Add(dependency.DefinitionId) ||
                        !dependencyNames.Add(dependency.GeneratedMeasureName))
                    {
                        throw new ArgumentException(
                            "The trusted named-set compilation has duplicate measure dependencies.",
                            nameof(compilation));
                    }

                    dependencies.Add(new DesiredPivotNamedSetMeasureDependency(
                        dependency.DefinitionId,
                        dependency.GeneratedMeasureName,
                        dependency.MeasureDefinitionFingerprint,
                        dependency.MeasureFormulaFingerprint,
                        PivotModelMeasureCanonical.CreateDescriptionMarker(
                            setupId,
                            dependency.DefinitionId,
                            dependency.MeasureDefinitionFingerprint)));
                }

                result.Add(new DesiredPivotNamedSet(
                    compilation.SourceFingerprint,
                    compilation.CompilationFingerprint,
                    set.DefinitionId,
                    set.DisplayOrder,
                    set.GeneratedSetName,
                    set.Caption,
                    set.Axis,
                    set.MdxFormula,
                    set.Dynamic,
                    set.FlattenHierarchies,
                    set.HierarchizeDistinct,
                    dependencies,
                    set.DefinitionFingerprint,
                    set.FormulaFingerprint,
                    PivotNamedSetCanonical.CreateDisplayFolderMarker(
                        setupId,
                        set.DefinitionId,
                        set.DefinitionFingerprint)));
            }

            return new ReadOnlyCollection<DesiredPivotNamedSet>(
                result.OrderBy(set => set.DisplayOrder).ToList());
        }
    }
}
