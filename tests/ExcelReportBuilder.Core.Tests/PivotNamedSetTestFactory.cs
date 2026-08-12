using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.PivotPlus.NamedSets;

namespace ExcelReportBuilder.Core.Tests;

internal static class PivotNamedSetTestFactory
{
    public const string SourceFingerprint =
        "pivot.source.v1:sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    public static PivotNamedSetSchema Schema()
    {
        return new PivotNamedSetSchema(
            SourceFingerprint,
            PivotNamedSetProviderKind.DataModel,
            new[]
            {
                Hierarchy(
                    "region_h",
                    "[Sales].[Region]",
                    "region_all",
                    new PivotNamedSetLevelSchema(
                        "region_all_level",
                        "[Sales].[Region].[All Level]",
                        0,
                        true,
                        new[]
                        {
                            Member(
                                "region_all",
                                "[Sales].[Region].[All]",
                                isAll: true)
                        }),
                    new PivotNamedSetLevelSchema(
                        "region_level",
                        "[Sales].[Region].[Region]",
                        1,
                        true,
                        new[]
                        {
                            Member(
                                "north",
                                "[Sales].[Region].&[North]",
                                "region_all"),
                            Member(
                                "south",
                                "[Sales].[Region].&[South]",
                                "region_all")
                        })),
                Hierarchy(
                    "department_h",
                    "[Sales].[Department]",
                    "department_all",
                    new PivotNamedSetLevelSchema(
                        "department_all_level",
                        "[Sales].[Department].[All Level]",
                        0,
                        true,
                        new[]
                        {
                            Member(
                                "department_all",
                                "[Sales].[Department].[All]",
                                isAll: true)
                        }),
                    new PivotNamedSetLevelSchema(
                        "department_level",
                        "[Sales].[Department].[Department]",
                        1,
                        true,
                        new[]
                        {
                            Member(
                                "consumer",
                                "[Sales].[Department].&[Consumer]",
                                "department_all"),
                            Member(
                                "enterprise",
                                "[Sales].[Department].&[Enterprise]",
                                "department_all")
                        })),
                Hierarchy(
                    "sku_h",
                    "[Product].[SKU]",
                    "sku_all",
                    new PivotNamedSetLevelSchema(
                        "sku_all_level",
                        "[Product].[SKU].[All Level]",
                        0,
                        true,
                        new[]
                        {
                            Member("sku_all", "[Product].[SKU].[All]", isAll: true)
                        }),
                    new PivotNamedSetLevelSchema(
                        "sku_level",
                        "[Product].[SKU].[SKU]",
                        1,
                        true,
                        new[]
                        {
                            Member("sku_a", "[Product].[SKU].&[A]", "sku_all"),
                            Member("sku_b", "[Product].[SKU].&[B]", "sku_all"),
                            Member("sku_c", "[Product].[SKU].&[C]", "sku_all")
                        }))
            });
    }

    public static PivotExplicitOrderedTuplesExpression ScopedDepartmentBranches()
    {
        return new PivotExplicitOrderedTuplesExpression(
            new[] { "region_h", "department_h" },
            new[]
            {
                Tuple(Catalog("north"), DefaultMember("department_h")),
                Tuple("north", "consumer"),
                Tuple("north", "enterprise"),
                Tuple(Catalog("south"), DefaultMember("department_h"))
            });
    }

    public static PivotNamedSetDefinition Rows(
        PivotNamedSetExpression? expression = null,
        string id = "row_set",
        string caption = "Management Rows")
    {
        return new PivotNamedSetDefinition(
            id,
            caption,
            PivotNamedSetAxis.Row,
            expression ?? ScopedDepartmentBranches());
    }

    public static PivotNamedSetCompilationRequest Request(
        IEnumerable<PivotNamedSetDefinition>? namedSets = null,
        PivotDaxCompilation? dax = null,
        PivotNamedSetSchema? schema = null,
        string? sourceFingerprint = null,
        IEnumerable<PivotNamedSetArtifactBinding>? bindings = null)
    {
        PivotNamedSetDefinition[] definitions =
            (namedSets ?? new[] { Rows() }).ToArray();
        return new PivotNamedSetCompilationRequest(
            new PivotNamedSetCollectionDefinition(
                sourceFingerprint ?? SourceFingerprint,
                schema ?? Schema(),
                definitions),
            bindings ?? definitions.Select(definition =>
                new PivotNamedSetArtifactBinding(
                    definition.Id,
                    "[PivotTablePlus_setup_" + definition.Id + "]")),
            dax);
    }

    public static PivotDaxCompilation Dax(
        string definitionId = "sales",
        string caption = "Net Sales")
    {
        return PivotDaxCompiler.Compile(PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                definitionId,
                caption,
                PivotCalculationTestFactory.Sum())
        }));
    }

    public static PivotNamedSetTuple Tuple(params string[] memberIds)
    {
        return PivotNamedSetTuple.FromCatalogMemberIds(memberIds);
    }

    public static PivotNamedSetTuple Tuple(
        params PivotNamedSetTupleMemberReference[] members)
    {
        return new PivotNamedSetTuple(members);
    }

    public static PivotNamedSetCatalogMemberReference Catalog(string memberId)
    {
        return new PivotNamedSetCatalogMemberReference(memberId);
    }

    public static PivotNamedSetHierarchyDefaultMemberReference DefaultMember(
        string hierarchyId)
    {
        return new PivotNamedSetHierarchyDefaultMemberReference(hierarchyId);
    }

    public static PivotNamedSetHierarchySchema Hierarchy(
        string id,
        string uniqueName,
        string? allMemberId,
        params PivotNamedSetLevelSchema[] levels)
    {
        return new PivotNamedSetHierarchySchema(
            id,
            uniqueName,
            allMemberId,
            true,
            levels,
            id + " caption");
    }

    public static PivotNamedSetMemberSchema Member(
        string id,
        string uniqueName,
        string? parentId = null,
        bool isAll = false)
    {
        return new PivotNamedSetMemberSchema(
            id,
            uniqueName,
            id + " caption",
            parentId,
            isAll);
    }
}
