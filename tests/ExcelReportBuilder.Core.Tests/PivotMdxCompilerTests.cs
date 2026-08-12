using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotMdxCompilerTests
{
    [Fact]
    public void Compiles_scoped_default_and_detail_branches_in_exact_custom_order()
    {
        PivotMdxCompilation compilation = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request());

        OwnedPivotNamedSetDefinition set = Assert.Single(compilation.NamedSets);
        Assert.Equal(PivotNamedSetTestFactory.SourceFingerprint, compilation.SourceFingerprint);
        Assert.StartsWith(
            "namedset.compilation.v1:sha256:",
            compilation.CompilationFingerprint,
            StringComparison.Ordinal);
        Assert.Equal("row_set", set.DefinitionId);
        Assert.Equal("[PivotTablePlus_setup_row_set]", set.GeneratedSetName);
        Assert.Equal("Management Rows", set.Caption);
        Assert.Equal(PivotNamedSetAxis.Row, set.Axis);
        Assert.False(set.Dynamic);
        Assert.False(set.FlattenHierarchies);
        Assert.False(set.HierarchizeDistinct);
        Assert.Empty(set.DirectMeasureDependencies);
        Assert.Empty(set.DirectMeasureDependencyDefinitionIds);
        Assert.Equal(
            "{([Sales].[Region].&[North], [Sales].[Department].DefaultMember), " +
            "([Sales].[Region].&[North], [Sales].[Department].&[Consumer]), " +
            "([Sales].[Region].&[North], [Sales].[Department].&[Enterprise]), " +
            "([Sales].[Region].&[South], [Sales].[Department].DefaultMember)}",
            set.MdxFormula);
        Assert.StartsWith("namedset.definition.v2:sha256:", set.DefinitionFingerprint);
        Assert.StartsWith("namedset.formula.v1:sha256:", set.FormulaFingerprint);
    }

    [Fact]
    public void Compiles_single_hierarchy_members_without_synthetic_tuple_syntax()
    {
        var rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple("sku_c"),
                    PivotNamedSetTestFactory.Tuple("sku_a")
                }));

        string formula = Assert.Single(PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(new[] { rows })).NamedSets).MdxFormula;

        Assert.Equal("{[Product].[SKU].&[C], [Product].[SKU].&[A]}", formula);
    }

    [Fact]
    public void Compiles_typed_top_count_against_exact_owned_measure_and_escapes_measure_name()
    {
        var topN = PivotNamedSetTestFactory.Rows(
            new PivotTopNLevelMembersExpression("sku_level", 7, "sales"));
        PivotDaxCompilation dax = PivotNamedSetTestFactory.Dax(
            "sales",
            "Net ] Sales");

        OwnedPivotNamedSetDefinition set = Assert.Single(PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(
                new[] { topN },
                dax))
            .NamedSets);

        Assert.Equal(
            "TopCount([Product].[SKU].[SKU].Members, 7, [Measures].[Net ]] Sales])",
            set.MdxFormula);
        Assert.True(set.Dynamic);
        Assert.Equal(new[] { "sales" }, set.DirectMeasureDependencyDefinitionIds);
        PivotNamedSetMeasureDependencyBinding dependency = Assert.Single(
            set.DirectMeasureDependencies);
        OwnedPivotMeasureDefinition measure = Assert.Single(dax.Measures);
        Assert.Equal(measure.DefinitionId, dependency.DefinitionId);
        Assert.Equal(measure.GeneratedMeasureName, dependency.GeneratedMeasureName);
        Assert.Equal(
            measure.DefinitionFingerprint,
            dependency.MeasureDefinitionFingerprint);
        Assert.Equal(measure.FormulaFingerprint, dependency.MeasureFormulaFingerprint);
    }

    [Fact]
    public void Keeps_generated_artifact_identity_separate_from_user_caption()
    {
        PivotNamedSetDefinition definition = PivotNamedSetTestFactory.Rows(
            caption: "North / Selected Detail");
        var binding = new PivotNamedSetArtifactBinding(
            definition.Id,
            "[PivotTablePlus_7f3a_rows]");

        OwnedPivotNamedSetDefinition set = Assert.Single(PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(
                new[] { definition },
                bindings: new[] { binding })).NamedSets);

        Assert.Equal("[PivotTablePlus_7f3a_rows]", set.GeneratedSetName);
        Assert.Equal("North / Selected Detail", set.Caption);
        Assert.DoesNotContain("North / Selected Detail", set.MdxFormula, StringComparison.Ordinal);
    }

    [Fact]
    public void Definition_and_formula_fingerprints_are_deterministic_and_separate()
    {
        PivotNamedSetCompilationRequest firstRequest = PivotNamedSetTestFactory.Request();
        PivotNamedSetCompilationRequest secondRequest = PivotNamedSetTestFactory.Request();
        OwnedPivotNamedSetDefinition first = Assert.Single(
            PivotMdxCompiler.Compile(firstRequest).NamedSets);
        OwnedPivotNamedSetDefinition second = Assert.Single(
            PivotMdxCompiler.Compile(secondRequest).NamedSets);
        var changedCaption = PivotNamedSetTestFactory.Rows(caption: "Changed Caption");
        OwnedPivotNamedSetDefinition captionChange = Assert.Single(
            PivotMdxCompiler.Compile(PivotNamedSetTestFactory.Request(
                new[] { changedCaption })).NamedSets);
        var changedOrder = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "region_h", "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple("south", "department_all"),
                    PivotNamedSetTestFactory.Tuple("north", "department_all"),
                    PivotNamedSetTestFactory.Tuple("north", "consumer"),
                    PivotNamedSetTestFactory.Tuple("north", "enterprise")
                }));
        OwnedPivotNamedSetDefinition orderChange = Assert.Single(
            PivotMdxCompiler.Compile(PivotNamedSetTestFactory.Request(
                new[] { changedOrder })).NamedSets);

        Assert.Equal(first.DefinitionFingerprint, second.DefinitionFingerprint);
        Assert.Equal(first.FormulaFingerprint, second.FormulaFingerprint);
        Assert.NotEqual(first.DefinitionFingerprint, captionChange.DefinitionFingerprint);
        Assert.Equal(first.FormulaFingerprint, captionChange.FormulaFingerprint);
        Assert.NotEqual(first.DefinitionFingerprint, orderChange.DefinitionFingerprint);
        Assert.NotEqual(first.FormulaFingerprint, orderChange.FormulaFingerprint);
    }

    [Fact]
    public void Compilation_fingerprint_is_deterministic_and_binds_the_exact_source()
    {
        PivotMdxCompilation first = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request());
        PivotMdxCompilation second = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request());
        const string otherSource =
            "pivot.source.v1:sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        var otherSchema = new PivotNamedSetSchema(
            otherSource,
            original.ProviderKind,
            original.Hierarchies);
        PivotMdxCompilation other = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(
                schema: otherSchema,
                sourceFingerprint: otherSource));

        Assert.Equal(first.CompilationFingerprint, second.CompilationFingerprint);
        Assert.Equal(otherSource, other.SourceFingerprint);
        Assert.NotEqual(first.CompilationFingerprint, other.CompilationFingerprint);
    }

    [Fact]
    public void Exact_dependency_bindings_detect_stale_or_mixed_dax_compilations()
    {
        PivotDaxCompilation sum = PivotNamedSetTestFactory.Dax("sales", "Net Sales");
        PivotDaxCompilation average = PivotDaxCompiler.Compile(
            PivotCalculationTestFactory.Set(new[]
            {
                PivotCalculationTestFactory.Measure(
                    "sales",
                    "Net Sales",
                    new PivotAggregateExpression(
                        "amount",
                        PivotCalculationAggregateFunction.Average))
            }));
        var topN = PivotNamedSetTestFactory.Rows(
            new PivotTopNLevelMembersExpression("sku_level", 7, "sales"));
        PivotMdxCompilation compiledAgainstSum = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(new[] { topN }, sum));
        PivotMdxCompilation compiledAgainstAverage = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(new[] { topN }, average));

        Assert.True(compiledAgainstSum.HasExactMeasureDependencies(sum));
        Assert.False(compiledAgainstSum.HasExactMeasureDependencies(average));
        Assert.False(compiledAgainstSum.HasExactMeasureDependencies(null));
        Assert.NotEqual(
            compiledAgainstSum.CompilationFingerprint,
            compiledAgainstAverage.CompilationFingerprint);
        Assert.NotEqual(
            Assert.Single(compiledAgainstSum.NamedSets)
                .DirectMeasureDependencies[0].MeasureFormulaFingerprint,
            Assert.Single(compiledAgainstAverage.NamedSets)
                .DirectMeasureDependencies[0].MeasureFormulaFingerprint);
    }

    [Fact]
    public void Explicit_tuple_compilations_need_no_dax_and_expose_immutable_bindings()
    {
        PivotMdxCompilation explicitCompilation = PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request());
        OwnedPivotNamedSetDefinition explicitSet = Assert.Single(
            explicitCompilation.NamedSets);
        Assert.Empty(explicitSet.DirectMeasureDependencies);
        Assert.True(explicitCompilation.HasExactMeasureDependencies(null));

        PivotDaxCompilation dax = PivotNamedSetTestFactory.Dax();
        var topN = PivotNamedSetTestFactory.Rows(
            new PivotTopNLevelMembersExpression("sku_level", 3, "sales"));
        OwnedPivotNamedSetDefinition boundSet = Assert.Single(
            PivotMdxCompiler.Compile(
                PivotNamedSetTestFactory.Request(new[] { topN }, dax)).NamedSets);
        IList<PivotNamedSetMeasureDependencyBinding> mutableView =
            Assert.IsAssignableFrom<IList<PivotNamedSetMeasureDependencyBinding>>(
                boundSet.DirectMeasureDependencies);

        Assert.Throws<NotSupportedException>(() => mutableView.RemoveAt(0));
        Assert.Single(boundSet.DirectMeasureDependencies);
    }

    [Fact]
    public void Preserves_provider_identifier_punctuation_as_one_bound_token()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        PivotNamedSetHierarchySchema sku = original.Hierarchies[2];
        var member = PivotNamedSetTestFactory.Member(
            "punctuation",
            "[Product].[SKU].&[A, TopCount(X)]",
            "sku_all");
        var skuHierarchy = new PivotNamedSetHierarchySchema(
            sku.Id,
            sku.ProviderUniqueName,
            sku.AllMemberId,
            true,
            new[]
            {
                sku.Levels[0],
                new PivotNamedSetLevelSchema(
                    "sku_level",
                    "[Product].[SKU].[SKU]",
                    1,
                    true,
                    new[] { member })
            },
            sku.Caption);
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            original.Hierarchies.Take(2).Concat(new[] { skuHierarchy }));
        var definition = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[] { PivotNamedSetTestFactory.Tuple("punctuation") }));

        string formula = Assert.Single(PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(
                new[] { definition },
                schema: schema)).NamedSets).MdxFormula;

        Assert.Equal("{[Product].[SKU].&[A, TopCount(X)]}", formula);
    }

    [Fact]
    public void Compiles_default_member_without_fabricating_an_all_pivot_item()
    {
        var department = new PivotNamedSetHierarchySchema(
            "department_h",
            "[Sales].[Department]",
            true,
            new[]
            {
                new PivotNamedSetLevelSchema(
                    "department_level",
                    "[Sales].[Department].[Department]",
                    1,
                    false,
                    Array.Empty<PivotNamedSetMemberSchema>())
            });
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { department });
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple(
                        PivotNamedSetTestFactory.DefaultMember("department_h"))
                }));

        OwnedPivotNamedSetDefinition compiled = Assert.Single(
            PivotMdxCompiler.Compile(PivotNamedSetTestFactory.Request(
                new[] { rows },
                schema: schema)).NamedSets);

        Assert.Equal("{[Sales].[Department].DefaultMember}", compiled.MdxFormula);
    }

    [Fact]
    public void Compiles_a_proven_all_catalog_member_as_its_exact_provider_name()
    {
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[] { PivotNamedSetTestFactory.Tuple("department_all") }));

        string formula = Assert.Single(PivotMdxCompiler.Compile(
            PivotNamedSetTestFactory.Request(new[] { rows })).NamedSets).MdxFormula;

        Assert.Equal("{[Sales].[Department].[All]}", formula);
    }

    [Fact]
    public void Fingerprints_distinguish_default_member_from_catalog_all()
    {
        PivotNamedSetDefinition defaultMember = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple(
                        PivotNamedSetTestFactory.DefaultMember("department_h"))
                }));
        PivotNamedSetDefinition catalogAll = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[] { PivotNamedSetTestFactory.Tuple("department_all") }));

        OwnedPivotNamedSetDefinition defaultCompiled = Assert.Single(
            PivotMdxCompiler.Compile(PivotNamedSetTestFactory.Request(
                new[] { defaultMember })).NamedSets);
        OwnedPivotNamedSetDefinition allCompiled = Assert.Single(
            PivotMdxCompiler.Compile(PivotNamedSetTestFactory.Request(
                new[] { catalogAll })).NamedSets);

        Assert.NotEqual(defaultCompiled.DefinitionFingerprint, allCompiled.DefinitionFingerprint);
        Assert.NotEqual(defaultCompiled.FormulaFingerprint, allCompiled.FormulaFingerprint);
    }

    [Fact]
    public void Invalid_contract_throws_typed_exception_without_compiling_mdx()
    {
        PivotNamedSetCompilationRequest invalid = PivotNamedSetTestFactory.Request(
            new[]
            {
                PivotNamedSetTestFactory.Rows(
                    new PivotExplicitOrderedTuplesExpression(
                        new[] { "sku_h" },
                        new[] { PivotNamedSetTestFactory.Tuple("unknown") }))
            });

        InvalidPivotNamedSetException exception = Assert.Throws<InvalidPivotNamedSetException>(
            () => PivotMdxCompiler.Compile(invalid));

        Assert.Contains(exception.Validation.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_UNKNOWN");
    }
}
