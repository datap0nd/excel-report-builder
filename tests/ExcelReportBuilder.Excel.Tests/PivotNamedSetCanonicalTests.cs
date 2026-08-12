using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Excel.PivotPlus.NamedSets;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class PivotNamedSetCanonicalTests
{
    [Fact]
    public void Stable_catalog_ids_and_source_fingerprints_are_order_independent_and_path_free()
    {
        string firstId = PivotNamedSetCanonical.CreateStableCatalogId(
            "hierarchy",
            "[Sales].[Region]");
        string secondId = PivotNamedSetCanonical.CreateStableCatalogId(
            "hierarchy",
            "[Sales].[Department]");
        string first = PivotNamedSetCanonical.CreateSourceFingerprint(
            "namedset.model-lineage.v1:sha256:" + new string('a', 64),
            new[] { "north", "south" });
        string second = PivotNamedSetCanonical.CreateSourceFingerprint(
            "namedset.model-lineage.v1:sha256:" + new string('a', 64),
            new[] { "south", "north" });

        Assert.Matches("^[A-Za-z0-9._-]+$", firstId);
        Assert.NotEqual(firstId, secondId);
        Assert.Equal(first, second);
        Assert.StartsWith("pivot.source.v1:sha256:", first);
        Assert.DoesNotContain("north", first, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Display_folder_marker_is_hash_only()
    {
        string marker = PivotNamedSetCanonical.CreateDisplayFolderMarker(
            "setup_1",
            "row_set",
            "namedset.definition.v2:sha256:" + new string('b', 64));

        Assert.StartsWith("PivotTable+|set|namedset.semantic.v1:sha256:", marker);
        Assert.DoesNotContain("setup_1", marker, StringComparison.Ordinal);
        Assert.DoesNotContain("row_set", marker, StringComparison.Ordinal);
    }

    [Fact]
    public void Scanner_finds_identifiers_only_outside_strings_and_comments()
    {
        const string formula =
            "{[PivotTablePlus_rows], 'ignore [PivotTablePlus_rows]', " +
            "[Measures].[Sales]} -- [PivotTablePlus_rows]\n" +
            "/* [PivotTablePlus_rows] */";

        MdxReferenceScanResult result = MdxNamedSetReferenceScanner.Scan(formula);

        Assert.True(result.IsComplete);
        Assert.Contains("[PivotTablePlus_rows]", result.BracketedIdentifiers);
        Assert.Contains("[Measures].[Sales]", result.BracketedIdentifiers);
        Assert.Contains("ignore [PivotTablePlus_rows]", result.QuotedLiterals);
        Assert.Equal(1, result.BracketedIdentifiers.Count(identifier =>
            identifier == "[PivotTablePlus_rows]"));
        Assert.True(MdxNamedSetReferenceScanner.MightReference(
            formula,
            "[PivotTablePlus_rows]"));
    }

    [Theory]
    [InlineData("{'[PivotTablePlus_rows]'}")]
    [InlineData("StrToSet('[Unrelated]')")]
    [InlineData("NameToSet (\"[Unrelated]\")")]
    [InlineData("StrToMember('[Sales].[Region].&[North]')")]
    [InlineData("StrToTuple('[Sales].[Region].&[North]')")]
    public void Scanner_blocks_quoted_or_dynamic_name_resolution(string formula)
    {
        Assert.True(MdxNamedSetReferenceScanner.MightReference(
            formula,
            "[PivotTablePlus_rows]"));
    }

    [Fact]
    public void Scanner_ignores_dynamic_function_names_in_comments()
    {
        const string formula = "{[Measures].[Sales]} /* StrToSet('[PivotTablePlus_rows]') */";

        MdxReferenceScanResult result = MdxNamedSetReferenceScanner.Scan(formula);

        Assert.True(result.IsComplete);
        Assert.False(result.HasDynamicNameResolution);
        Assert.False(MdxNamedSetReferenceScanner.MightReference(
            formula,
            "[PivotTablePlus_rows]"));
    }

    [Theory]
    [InlineData("{[Unclosed}")]
    [InlineData("{'unterminated}")]
    [InlineData("{/* unterminated")]
    public void Scanner_fails_closed_on_incomplete_syntax(string formula)
    {
        Assert.False(MdxNamedSetReferenceScanner.Scan(formula).IsComplete);
    }

    [Fact]
    public void Trusted_adapter_accepts_only_core_compilation_and_keeps_raw_mdx_internal()
    {
        PivotMdxCompilation compilation = CompileDefaultMemberSet();

        DesiredPivotNamedSet desired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired("setup_1", compilation));

        Assert.Equal("{[Sales].[Region].DefaultMember}", desired.RawMdx);
        Assert.Equal(compilation.SourceFingerprint, desired.SourceFingerprint);
        Assert.Equal(compilation.CompilationFingerprint, desired.CompilationFingerprint);
        Assert.Equal(
            PivotMdxFingerprint.ComputeFormula(desired.RawMdx),
            desired.FormulaFingerprint);
        Assert.DoesNotContain(typeof(PivotNamedSetFormulaTransport).Assembly
            .GetExportedTypes(), type => string.Equals(
                type.Namespace,
                "ExcelReportBuilder.Excel.PivotPlus.NamedSets",
                StringComparison.Ordinal));
        Assert.DoesNotContain(
            typeof(DesiredPivotNamedSet).GetProperties(),
            property => property.Name == "RawMdx" &&
                        property.GetMethod?.IsPublic == true);
    }

    [Fact]
    public void Trusted_adapter_preserves_exact_measure_dependency_bindings()
    {
        (PivotMdxCompilation compilation, PivotDaxCompilation dax) = CompileTopNSet(
            "pivot.source.v1:sha256:" + new string('c', 64));

        DesiredPivotNamedSet desired = Assert.Single(
            PivotNamedSetCompilationAdapter.CreateDesired("setup_1", compilation));
        DesiredPivotNamedSetMeasureDependency dependency = Assert.Single(
            desired.DirectMeasureDependencies);
        PivotNamedSetMeasureDependencyBinding compiledDependency = Assert.Single(
            Assert.Single(compilation.NamedSets).DirectMeasureDependencies);
        OwnedPivotMeasureDefinition measure = Assert.Single(dax.Measures);

        Assert.Equal(compiledDependency.DefinitionId, dependency.DefinitionId);
        Assert.Equal(compiledDependency.GeneratedMeasureName, dependency.GeneratedMeasureName);
        Assert.Equal(
            compiledDependency.MeasureDefinitionFingerprint,
            dependency.MeasureDefinitionFingerprint);
        Assert.Equal(
            compiledDependency.MeasureFormulaFingerprint,
            dependency.MeasureFormulaFingerprint);
        Assert.Equal(measure.FormulaFingerprint, dependency.MeasureFormulaFingerprint);
        Assert.StartsWith("PivotTable+|measure.semantic.v1:sha256:",
            dependency.ExpectedDescriptionMarker);
    }

    internal static PivotMdxCompilation CompileDefaultMemberSet(
        string caption = "Management Rows",
        string generatedName = "[PivotTablePlus_setup_rows]",
        string sourceFingerprint =
            "pivot.source.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
        bool flattenHierarchies = false)
    {
        var hierarchy = new PivotNamedSetHierarchySchema(
            "region_h",
            "[Sales].[Region]",
            true,
            new[]
            {
                new PivotNamedSetLevelSchema(
                    "region_level",
                    "[Sales].[Region].[Region]",
                    1,
                    false,
                    Array.Empty<PivotNamedSetMemberSchema>())
            });
        var schema = new PivotNamedSetSchema(
            sourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { hierarchy });
        var definition = new PivotNamedSetDefinition(
            "row_set",
            caption,
            PivotNamedSetAxis.Row,
            new PivotExplicitOrderedTuplesExpression(
                new[] { "region_h" },
                new[]
                {
                    new PivotNamedSetTuple(new PivotNamedSetTupleMemberReference[]
                    {
                        new PivotNamedSetHierarchyDefaultMemberReference("region_h")
                    })
                }),
            flattenHierarchies);
        return PivotMdxCompiler.Compile(new PivotNamedSetCompilationRequest(
            new PivotNamedSetCollectionDefinition(
                sourceFingerprint,
                schema,
                new[] { definition }),
            new[] { new PivotNamedSetArtifactBinding("row_set", generatedName) }));
    }

    internal static (PivotMdxCompilation Compilation, PivotDaxCompilation Dax)
        CompileTopNSet(string sourceFingerprint)
    {
        var modelSchema = new PivotModelSchema(new[]
        {
            new PivotModelTableSchema(
                "fact",
                "Fact Sales",
                new[]
                {
                    new PivotModelFieldSchema(
                        "amount",
                        "Amount",
                        PivotModelDataType.DecimalNumber)
                })
        });
        PivotDaxCompilation dax = PivotDaxCompiler.Compile(
            new PivotMeasureSetDefinition(
                modelSchema,
                new[]
                {
                    new PivotMeasureDefinition(
                        "sales",
                        "Net Sales",
                        "fact",
                        new PivotMeasureFormat(
                            PivotMeasureFormatKind.DecimalNumber,
                            2,
                            true),
                        new PivotAggregateExpression(
                            "amount",
                            PivotCalculationAggregateFunction.Sum))
                }));
        var hierarchy = new PivotNamedSetHierarchySchema(
            "sku_h",
            "[Product].[SKU]",
            true,
            new[]
            {
                new PivotNamedSetLevelSchema(
                    "sku_level",
                    "[Product].[SKU].[SKU]",
                    0,
                    true,
                    Array.Empty<PivotNamedSetMemberSchema>())
            });
        var schema = new PivotNamedSetSchema(
            sourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { hierarchy });
        var definition = new PivotNamedSetDefinition(
            "row_set",
            "Top products",
            PivotNamedSetAxis.Row,
            new PivotTopNLevelMembersExpression("sku_level", 5, "sales"));
        PivotMdxCompilation compilation = PivotMdxCompiler.Compile(
            new PivotNamedSetCompilationRequest(
                new PivotNamedSetCollectionDefinition(
                    sourceFingerprint,
                    schema,
                    new[] { definition }),
                new[]
                {
                    new PivotNamedSetArtifactBinding(
                        "row_set",
                        "[PivotTablePlus_setup_top_products]")
                },
                dax));
        return (compilation, dax);
    }
}
