using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class DenseAxisPlannerTests
{
    [Fact]
    public void Groups_orders_ranks_and_aggregates_others_with_typed_subtotals()
    {
        var fields = new[]
        {
            new PivotFieldPlan
            {
                Field = "Region",
                MemberOrder =
                {
                    ScalarValue.FromText("South"),
                    ScalarValue.FromText("North"),
                    ScalarValue.FromText("East")
                },
                GroupBuckets =
                {
                    new MemberGroupBucketSpec
                    {
                        Id = "key",
                        Label = "Key",
                        Members =
                        {
                            ScalarValue.FromText("North"),
                            ScalarValue.FromText("South")
                        }
                    },
                    new MemberGroupBucketSpec
                    {
                        Id = "remaining",
                        Label = "Remaining",
                        IncludeUnmatched = true
                    }
                },
                TopN = new TopNSpec
                {
                    Count = 1,
                    MeasureId = "amount",
                    IncludeOthers = true,
                    OthersLabel = "Others"
                },
                Subtotals = new SubtotalSpec
                {
                    Mode = SubtotalMode.Automatic,
                    Label = "Region subtotal",
                    StyleId = "level-subtotal"
                },
                MemberStages =
                {
                    PivotMemberStageKind.ApplyMemberOrder,
                    PivotMemberStageKind.GroupMembers,
                    PivotMemberStageKind.ApplyTopN,
                    PivotMemberStageKind.AggregateOthers
                }
            },
            new PivotFieldPlan
            {
                Field = "Category",
                Subtotals = new SubtotalSpec { Mode = SubtotalMode.None }
            }
        };
        var raw = new List<List<PivotFilterItem>>
        {
            Path("North", "A"),
            Path("South", "A"),
            Path("East", "B")
        };

        var result = DenseAxisPlanner.Build(
            raw,
            fields,
            "block-subtotal",
            (_, filters) => filters.Any(item => Equals(item.Value, "North"))
                ? 20m
                : filters.Any(item => Equals(item.Value, "South"))
                    ? 10m
                    : 5m);

        Assert.Equal(4, result.Count);
        Assert.Equal("Key", result[0].DisplayItems[0].Value);
        Assert.Equal(2, result[0].MemberFilterSets.Count);
        Assert.True(result[1].IsSubtotal);
        Assert.Equal("Region subtotal", result[1].DisplayItems[0].Value);
        Assert.Equal("level-subtotal", result[1].StyleId);
        Assert.Equal("Others", result[2].DisplayItems[0].Value);
        Assert.Null(result[2].DisplayItems[1].Value);
        Assert.True(result[3].IsSubtotal);
    }

    [Fact]
    public void Applies_explicit_member_order_stably_within_each_parent()
    {
        var fields = new[]
        {
            new PivotFieldPlan
            {
                Field = "Region",
                MemberOrder =
                {
                    ScalarValue.FromText("South"),
                    ScalarValue.FromText("North")
                },
                Subtotals = new SubtotalSpec { Mode = SubtotalMode.None },
                MemberStages = { PivotMemberStageKind.ApplyMemberOrder }
            }
        };

        var result = DenseAxisPlanner.Build(
            new List<List<PivotFilterItem>>
            {
                Single("North"),
                Single("East"),
                Single("South")
            },
            fields,
            null,
            (_, _) => 0m);

        Assert.Equal(
            new object?[] { "South", "North", "East" },
            result.Select(path => path.DisplayItems[0].Value));
    }

    private static List<PivotFilterItem> Path(string region, string category)
    {
        return new List<PivotFilterItem>
        {
            new PivotFilterItem { Field = "Region", Value = region },
            new PivotFilterItem { Field = "Category", Value = category }
        };
    }

    private static List<PivotFilterItem> Single(string region)
    {
        return new List<PivotFilterItem>
        {
            new PivotFilterItem { Field = "Region", Value = region }
        };
    }
}
