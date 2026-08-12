using ExcelReportBuilder.Core.PivotPlus.NamedSets;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotNamedSetValidatorTests
{
    [Fact]
    public void Accepts_exact_ordered_asymmetric_tuples()
    {
        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request());

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(x => x.Code)));
    }

    [Fact]
    public void Requires_exact_matching_canonical_source_fingerprint()
    {
        PivotNamedSetCompilationRequest mismatch = PivotNamedSetTestFactory.Request(
            sourceFingerprint:
            "pivot.source.v1:sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb");
        PivotNamedSetCompilationRequest malformed = PivotNamedSetTestFactory.Request(
            sourceFingerprint: "not-a-fingerprint");

        ValidationResult mismatchResult = PivotNamedSetValidator.Validate(mismatch);
        ValidationResult malformedResult = PivotNamedSetValidator.Validate(malformed);

        Assert.Contains(mismatchResult.Issues,
            issue => issue.Code == "PIVOT_SET_SOURCE_FINGERPRINT_MISMATCH");
        Assert.Contains(malformedResult.Issues,
            issue => issue.Code == "PIVOT_SET_SOURCE_FINGERPRINT_INVALID");
    }

    [Fact]
    public void Rejects_a_second_set_on_the_same_axis_and_duplicate_captions()
    {
        PivotNamedSetDefinition first = PivotNamedSetTestFactory.Rows();
        PivotNamedSetDefinition second = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[] { PivotNamedSetTestFactory.Tuple("sku_a") }),
            id: "other_rows");

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { first, second }));

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_SET_AXIS_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_SET_CAPTION_DUPLICATE");
    }

    [Fact]
    public void Allows_one_row_set_and_one_column_set()
    {
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows();
        var columns = new PivotNamedSetDefinition(
            "column_set",
            "Selected SKUs",
            PivotNamedSetAxis.Column,
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple("sku_c"),
                    PivotNamedSetTestFactory.Tuple("sku_a")
                }));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { rows, columns }));

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(x => x.Code)));
    }

    [Fact]
    public void Rejects_unknown_wrong_hierarchy_duplicate_and_wrong_arity_tuple_members()
    {
        var expression = new PivotExplicitOrderedTuplesExpression(
            new[] { "region_h", "department_h" },
            new[]
            {
                PivotNamedSetTestFactory.Tuple("consumer", "north"),
                PivotNamedSetTestFactory.Tuple("unknown"),
                PivotNamedSetTestFactory.Tuple("consumer", "north")
            });

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[]
            {
                PivotNamedSetTestFactory.Rows(expression)
            }));

        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_HIERARCHY_MISMATCH");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_COUNT_MISMATCH");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_UNKNOWN");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_DUPLICATE");
    }

    [Fact]
    public void Rejects_incomplete_hierarchy_and_referenced_catalog_member_identity()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        PivotNamedSetHierarchySchema region = original.Hierarchies[0];
        var incompleteLevel = new PivotNamedSetLevelSchema(
            region.Levels[1].Id,
            region.Levels[1].ProviderUniqueName,
            region.Levels[1].Ordinal,
            false,
            region.Levels[1].Members);
        var incompleteRegion = new PivotNamedSetHierarchySchema(
            region.Id,
            region.ProviderUniqueName,
            null,
            false,
            new[] { region.Levels[0], incompleteLevel },
            region.Caption);
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { incompleteRegion }.Concat(original.Hierarchies.Skip(1)));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(schema: schema));

        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_HIERARCHY_IDENTITY_INCOMPLETE");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_HIERARCHY_IDENTITY_INCOMPLETE");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_IDENTITY_INCOMPLETE");
        Assert.DoesNotContain(result.Issues,
            issue => issue.Code == "PIVOT_SET_ALL_MEMBER_UNRESOLVED");
    }

    [Fact]
    public void Allows_data_model_default_member_without_all_member_or_member_catalog()
    {
        var hierarchy = new PivotNamedSetHierarchySchema(
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
            },
            "Department");
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { hierarchy });
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple(
                        PivotNamedSetTestFactory.DefaultMember("department_h"))
                }));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { rows }, schema: schema));

        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(x => x.Code)));
    }

    [Fact]
    public void Rejects_default_member_for_a_non_data_model_provider()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.Unknown,
            original.Hierarchies);

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(schema: schema));

        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_SCHEMA_PROVIDER_UNSUPPORTED");
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_DEFAULT_PROVIDER_UNSUPPORTED");
    }

    [Fact]
    public void Rejects_unknown_and_mismatched_default_member_hierarchies()
    {
        PivotNamedSetDefinition unknown = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple(
                        PivotNamedSetTestFactory.DefaultMember("missing_h"))
                }));
        PivotNamedSetDefinition mismatch = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "department_h" },
                new[]
                {
                    PivotNamedSetTestFactory.Tuple(
                        PivotNamedSetTestFactory.DefaultMember("region_h"))
                }));

        ValidationResult unknownResult = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { unknown }));
        ValidationResult mismatchResult = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { mismatch }));

        Assert.Contains(unknownResult.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_UNKNOWN");
        Assert.Contains(mismatchResult.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_DEFAULT_HIERARCHY_MISMATCH");
    }

    [Fact]
    public void Requires_a_supplied_all_member_id_to_resolve_to_the_proven_all_member()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        PivotNamedSetHierarchySchema region = original.Hierarchies[0];
        var invalidRegion = new PivotNamedSetHierarchySchema(
            region.Id,
            region.ProviderUniqueName,
            "north",
            true,
            region.Levels,
            region.Caption);
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { invalidRegion }.Concat(original.Hierarchies.Skip(1)));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(schema: schema));

        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_ALL_MEMBER_UNRESOLVED");
    }

    [Fact]
    public void Catalog_member_still_requires_a_complete_exact_catalog()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        PivotNamedSetHierarchySchema sku = original.Hierarchies[2];
        var incompleteLevel = new PivotNamedSetLevelSchema(
            sku.Levels[1].Id,
            sku.Levels[1].ProviderUniqueName,
            sku.Levels[1].Ordinal,
            false,
            sku.Levels[1].Members);
        var incompleteSku = new PivotNamedSetHierarchySchema(
            sku.Id,
            sku.ProviderUniqueName,
            sku.AllMemberId,
            true,
            new[] { sku.Levels[0], incompleteLevel },
            sku.Caption);
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            original.Hierarchies.Take(2).Concat(new[] { incompleteSku }));
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[] { PivotNamedSetTestFactory.Tuple("sku_a") }));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { rows }, schema: schema));

        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_IDENTITY_INCOMPLETE");
    }

    [Fact]
    public void Rejects_expression_like_provider_tokens_but_accepts_punctuation_inside_brackets()
    {
        PivotNamedSetSchema original = PivotNamedSetTestFactory.Schema();
        PivotNamedSetHierarchySchema sku = original.Hierarchies[2];
        PivotNamedSetLevelSchema skuLevel = sku.Levels[1];
        var unsafeLevel = new PivotNamedSetLevelSchema(
            skuLevel.Id,
            "TopCount([Product].[SKU], 3, [Measures].[X])",
            skuLevel.Ordinal,
            true,
            skuLevel.Members);
        var unsafeHierarchy = new PivotNamedSetHierarchySchema(
            sku.Id,
            sku.ProviderUniqueName,
            sku.AllMemberId,
            true,
            new[] { sku.Levels[0], unsafeLevel },
            sku.Caption);
        var unsafeSchema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            original.Hierarchies.Take(2).Concat(new[] { unsafeHierarchy }));

        ValidationResult unsafeResult = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(schema: unsafeSchema));

        Assert.Contains(unsafeResult.Issues,
            issue => issue.Code == "PIVOT_SET_LEVEL_UNIQUE_NAME_INVALID");

        var punctuationMember = PivotNamedSetTestFactory.Member(
            "sku_punctuation",
            "[Product].[SKU].&[A, TopCount(X)]",
            "sku_all");
        var punctuationLevel = new PivotNamedSetLevelSchema(
            skuLevel.Id,
            skuLevel.ProviderUniqueName,
            skuLevel.Ordinal,
            true,
            new[] { punctuationMember });
        var punctuationHierarchy = new PivotNamedSetHierarchySchema(
            sku.Id,
            sku.ProviderUniqueName,
            sku.AllMemberId,
            true,
            new[] { sku.Levels[0], punctuationLevel },
            sku.Caption);
        var punctuationSchema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            original.Hierarchies.Take(2).Concat(new[] { punctuationHierarchy }));
        var punctuationSet = new PivotNamedSetDefinition(
            "punctuation_set",
            "Punctuation",
            PivotNamedSetAxis.Row,
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[] { PivotNamedSetTestFactory.Tuple("sku_punctuation") }));

        ValidationResult punctuationResult = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(
                new[] { punctuationSet },
                schema: punctuationSchema));

        Assert.True(punctuationResult.IsValid,
            string.Join("; ", punctuationResult.Issues.Select(x => x.Code)));
    }

    [Fact]
    public void Requires_safe_setup_namespaced_artifact_binding_for_every_definition()
    {
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows();
        ValidationResult missing = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(
                new[] { rows },
                bindings: Array.Empty<PivotNamedSetArtifactBinding>()));
        ValidationResult arbitrary = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(
                new[] { rows },
                bindings: new[]
                {
                    new PivotNamedSetArtifactBinding(
                        rows.Id,
                        "[Safe]); DELETE FROM Cube; --]")
                }));

        Assert.Contains(missing.Issues,
            issue => issue.Code == "PIVOT_SET_ARTIFACT_BINDING_COUNT_MISMATCH");
        Assert.Contains(missing.Issues,
            issue => issue.Code == "PIVOT_SET_ARTIFACT_BINDING_MISSING");
        Assert.Contains(arbitrary.Issues,
            issue => issue.Code == "PIVOT_SET_ARTIFACT_NAME_INVALID");
    }

    [Fact]
    public void Top_n_requires_bounded_count_non_all_level_and_owned_dax_measure()
    {
        var topN = new PivotTopNLevelMembersExpression("sku_level", 10, "sales");
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(topN);

        ValidationResult missingDax = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { rows }));
        ValidationResult unknownMeasure = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(
                new[] { rows },
                PivotNamedSetTestFactory.Dax("other")));
        ValidationResult invalidCount = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(
                new[]
                {
                    PivotNamedSetTestFactory.Rows(
                        new PivotTopNLevelMembersExpression(
                            "sku_all_level",
                            1001,
                            "sales"))
                },
                PivotNamedSetTestFactory.Dax()));

        Assert.Contains(missingDax.Issues,
            issue => issue.Code == "PIVOT_SET_TOPN_DAX_COMPILATION_REQUIRED");
        Assert.Contains(unknownMeasure.Issues,
            issue => issue.Code == "PIVOT_SET_TOPN_MEASURE_UNKNOWN");
        Assert.Contains(invalidCount.Issues,
            issue => issue.Code == "PIVOT_SET_TOPN_COUNT_INVALID");
        Assert.Contains(invalidCount.Issues,
            issue => issue.Code == "PIVOT_SET_TOPN_ALL_LEVEL_INVALID");
    }

    [Fact]
    public void Rejects_formula_that_exceeds_the_compiled_mdx_limit()
    {
        var all = PivotNamedSetTestFactory.Member(
            "all",
            "[Synthetic].[Item].[All]",
            isAll: true);
        var members = Enumerable.Range(0, 900)
            .Select(index => PivotNamedSetTestFactory.Member(
                "member_" + index,
                "[Synthetic].[Item].&[Member " + index.ToString("D4") +
                " with a bounded long identity]",
                "all"))
            .ToArray();
        var hierarchy = PivotNamedSetTestFactory.Hierarchy(
            "item_h",
            "[Synthetic].[Item]",
            "all",
            new PivotNamedSetLevelSchema(
                "all_level",
                "[Synthetic].[Item].[All Level]",
                0,
                true,
                new[] { all }),
            new PivotNamedSetLevelSchema(
                "item_level",
                "[Synthetic].[Item].[Item]",
                1,
                true,
                members));
        var schema = new PivotNamedSetSchema(
            PivotNamedSetTestFactory.SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[] { hierarchy });
        var set = new PivotNamedSetDefinition(
            "large_set",
            "Large Set",
            PivotNamedSetAxis.Row,
            new PivotExplicitOrderedTuplesExpression(
                new[] { "item_h" },
                members.Select(member =>
                    PivotNamedSetTestFactory.Tuple(member.Id))));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { set }, schema: schema));

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_SET_FORMULA_LIMIT");
    }

    [Fact]
    public void Contracts_defensively_copy_caller_collections_and_expose_only_closed_kinds()
    {
        var members = new List<PivotNamedSetTupleMemberReference>
        {
            PivotNamedSetTestFactory.Catalog("north"),
            PivotNamedSetTestFactory.DefaultMember("department_h")
        };
        var tuples = new List<PivotNamedSetTuple>
        {
            new PivotNamedSetTuple(members)
        };
        var hierarchies = new List<string> { "region_h", "department_h" };
        var expression = new PivotExplicitOrderedTuplesExpression(hierarchies, tuples);

        members[0] = PivotNamedSetTestFactory.Catalog("south");
        tuples.Clear();
        hierarchies.Clear();

        Assert.Equal(new[] { "region_h", "department_h" }, expression.HierarchyIds);
        PivotNamedSetTuple tuple = Assert.Single(expression.Tuples);
        Assert.Equal(
            "north",
            Assert.IsType<PivotNamedSetCatalogMemberReference>(tuple.Members[0]).MemberId);
        Assert.Equal(
            "department_h",
            Assert.IsType<PivotNamedSetHierarchyDefaultMemberReference>(tuple.Members[1])
                .HierarchyId);
        Assert.Equal(
            new[]
            {
                PivotNamedSetExpressionKind.ExplicitOrderedTuples,
                PivotNamedSetExpressionKind.TopNLevelMembers
            },
            Enum.GetValues<PivotNamedSetExpressionKind>());
        Assert.Equal(
            new[]
            {
                PivotNamedSetTupleMemberKind.CatalogMember,
                PivotNamedSetTupleMemberKind.HierarchyDefaultMember
            },
            Enum.GetValues<PivotNamedSetTupleMemberKind>());
        Assert.Empty(typeof(PivotNamedSetExpression).GetConstructors());
        Assert.Empty(typeof(PivotNamedSetTupleMemberReference).GetConstructors());
    }

    [Fact]
    public void Tuple_member_union_has_no_raw_mdx_variant_or_entry_point()
    {
        Type[] variants = typeof(PivotNamedSetTupleMemberReference).Assembly
            .GetTypes()
            .Where(type => type.BaseType == typeof(PivotNamedSetTupleMemberReference))
            .OrderBy(type => type.Name, StringComparer.Ordinal)
            .ToArray();
        var unsafeReference = new PivotNamedSetCatalogMemberReference(
            "TopCount([Product].[SKU],3,[Measures].[Sales])");
        PivotNamedSetDefinition rows = PivotNamedSetTestFactory.Rows(
            new PivotExplicitOrderedTuplesExpression(
                new[] { "sku_h" },
                new[] { new PivotNamedSetTuple(new[] { unsafeReference }) }));

        ValidationResult result = PivotNamedSetValidator.Validate(
            PivotNamedSetTestFactory.Request(new[] { rows }));

        Assert.Equal(
            new[]
            {
                typeof(PivotNamedSetCatalogMemberReference),
                typeof(PivotNamedSetHierarchyDefaultMemberReference)
            }.OrderBy(type => type.Name, StringComparer.Ordinal),
            variants);
        Assert.Contains(result.Issues,
            issue => issue.Code == "PIVOT_SET_TUPLE_MEMBER_ID_INVALID");
        Assert.DoesNotContain(
            typeof(PivotNamedSetTupleMemberReference).GetProperties(),
            property => property.Name.Contains("Mdx", StringComparison.OrdinalIgnoreCase) ||
                        property.Name.Contains("Formula", StringComparison.OrdinalIgnoreCase));
    }
}
