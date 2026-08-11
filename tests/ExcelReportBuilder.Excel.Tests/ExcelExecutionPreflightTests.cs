using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ExcelExecutionPreflightTests
{
    [Fact]
    public void Rejects_unsupported_schema_before_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);
        specification.SchemaVersion = "2.0";

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Rejects_calculated_measure_in_native_matrix_before_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);
        plan.Blocks[0].Pivot.Values[0].RequiresPostAggregationCalculation = true;
        plan.Blocks[0].Pivot.Values[0].Expression = new RatioMeasureExpression
        {
            Numerator = new AggregateMeasureExpression { Field = "Amount" },
            Denominator = new AggregateMeasureExpression { Field = "Units" }
        };

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Accepts_direct_aggregate_native_matrix()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Rejects_check_measure_that_is_not_rendered_by_any_block()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Checks.Add(new BuildCheckPlan
        {
            Id = "balance-check",
            Kind = ReportCheckKind.Balance,
            MeasureId = "amount",
            ComparedMeasureId = "unrendered"
        });

        var exception = Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));

        Assert.Contains("unrendered", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Accepts_check_measure_rendered_by_another_block()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks.Add(CreateBlock(ReportOutputMode.DenseGrid, "comparison"));
        plan.Checks.Add(new BuildCheckPlan
        {
            Id = "balance-check",
            Kind = ReportCheckKind.Balance,
            MeasureId = "amount",
            ComparedMeasureId = "comparison"
        });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Accepts_explicit_spacers_for_dense_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan { Field = "Region" });
        plan.Blocks[0].Presentation.Spacers.Add(new SpacerSpec
        {
            Axis = SpacerAxis.Row,
            BeforeLevel = 0,
            Count = 1
        });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Rejects_a_spacer_for_a_missing_hierarchy_level()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Presentation.Spacers.Add(new SpacerSpec
        {
            Axis = SpacerAxis.Row,
            BeforeLevel = 0,
            Count = 1
        });

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Accepts_grouping_buckets_for_dense_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan
        {
            Field = "Region",
            GroupBuckets =
            {
                new MemberGroupBucketSpec
                {
                    Id = "group",
                    Label = "Group",
                    Members = { ScalarValue.FromText("North") }
                }
            }
        });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Rejects_unconsumed_presentation_settings_before_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Presentation.Options.RepeatRowLabels = true;

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Accepts_custom_dense_subtotal_presentation()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan
        {
            Field = "Region",
            Subtotals = new SubtotalSpec
            {
                Mode = SubtotalMode.Automatic,
                Placement = TotalPlacement.AfterMembers,
                Label = "Region total"
            }
        });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Rejects_grouping_buckets_for_native_execution()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan
        {
            Field = "Region",
            GroupBuckets =
            {
                new MemberGroupBucketSpec
                {
                    Id = "group",
                    Label = "Group",
                    Members = { ScalarValue.FromText("North") }
                }
            }
        });

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Accepts_dense_complex_member_and_presentation_pipeline()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Presentation.SubtotalStyleId = "subtotal";
        plan.Blocks[0].Presentation.Spacers.Add(new SpacerSpec
        {
            Axis = SpacerAxis.Row,
            BeforeLevel = 1,
            Count = 1
        });
        plan.Blocks[0].Pivot.GrandTotals.StyleId = "subtotal";
        plan.Blocks[0].Pivot.GrandTotals.ColumnLabel = "All periods";
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan
        {
            Field = "Region",
            MemberOrder = { ScalarValue.FromText("North"), ScalarValue.FromText("South") },
            GroupBuckets =
            {
                new MemberGroupBucketSpec
                {
                    Id = "primary",
                    Label = "Primary",
                    Members = { ScalarValue.FromText("North") }
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
                Count = 5,
                MeasureId = "amount",
                IncludeOthers = true
            },
            Subtotals = new SubtotalSpec
            {
                Mode = SubtotalMode.Automatic,
                Label = "Region total",
                StyleId = "subtotal"
            },
            MemberStages =
            {
                PivotMemberStageKind.ApplyMemberOrder,
                PivotMemberStageKind.GroupMembers,
                PivotMemberStageKind.ApplyTopN,
                PivotMemberStageKind.AggregateOthers
            }
        });
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan
        {
            Field = "Category",
            Subtotals = new SubtotalSpec { Mode = SubtotalMode.None }
        });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Accepts_default_dense_presentation_settings()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.DenseGrid);
        plan.Blocks[0].Pivot.Rows.Add(new PivotFieldPlan { Field = "Region" });

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Rejects_distinct_count_for_worksheet_pivot()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);
        plan.Blocks[0].Pivot.Values[0].AggregateComponents[0].Function = AggregateFunction.DistinctCount;
        plan.Blocks[0].Pivot.UseDataModel = false;

        Assert.Throws<NotSupportedException>(() =>
            ExcelExecutionPreflight.DemandSupported(specification, plan));
    }

    [Fact]
    public void Accepts_distinct_count_for_data_model_pivot()
    {
        var (specification, plan) = CreateDirectPlan(ReportOutputMode.StandardMatrix);
        plan.Blocks[0].Pivot.Values[0].AggregateComponents[0].Function = AggregateFunction.DistinctCount;
        plan.Blocks[0].Pivot.UseDataModel = true;

        ExcelExecutionPreflight.DemandSupported(specification, plan);
    }

    [Fact]
    public void Uses_data_model_distinct_count_consolidation_code()
    {
        Assert.Equal(11, NativePivotTableExecutor.ConsolidationFunction(AggregateFunction.DistinctCount));
    }

    private static (ReportSpecV1 Specification, ReportBuildPlan Plan) CreateDirectPlan(ReportOutputMode mode)
    {
        var expression = new AggregateMeasureExpression
        {
            Field = "Amount",
            Function = AggregateFunction.Sum
        };
        var specification = new ReportSpecV1
        {
            Id = "report",
            OwnershipId = "owned_report",
            Name = "Report"
        };
        var plan = new ReportBuildPlan
        {
            SpecificationId = specification.Id,
            OwnershipId = specification.OwnershipId,
            Blocks = { CreateBlock(mode, "amount", expression) }
        };
        return (specification, plan);
    }

    private static DenseReportBlockPlan CreateBlock(
        ReportOutputMode mode,
        string measureId,
        MeasureExpression? expression = null)
    {
        expression ??= new AggregateMeasureExpression
        {
            Field = "Amount",
            Function = AggregateFunction.Sum
        };
        return new DenseReportBlockPlan
        {
            OutputMode = mode,
            Pivot = new PivotTablePlan
            {
                Values =
                {
                    new PivotValuePlan
                    {
                        MeasureId = measureId,
                        Label = measureId,
                        Expression = expression,
                        RequiresPostAggregationCalculation = false,
                        AggregateComponents =
                        {
                            new PivotAggregateComponentPlan
                            {
                                Id = measureId + "_component_1",
                                Field = "Amount",
                                Function = AggregateFunction.Sum
                            }
                        }
                    }
                }
            }
        };
    }
}
