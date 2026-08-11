using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

public sealed class ReportBuildPlannerTests
{
    [Fact]
    public void Produces_owned_dense_plans_for_each_independent_anchor()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Equal(2, plan.Blocks.Count);
        Assert.Equal(new[] { "B3", "J3" }, plan.Blocks.Select(block => block.AnchorCell));
        Assert.Equal(ReportOutputMode.DenseGrid, plan.Blocks[0].OutputMode);
        Assert.Equal(ReportOutputMode.MetricStack, plan.Blocks[1].OutputMode);
        Assert.All(plan.Blocks, block => Assert.StartsWith("ERB_owned_", block.Pivot.ManagedPivotName));
        Assert.Contains(plan.Blocks[0].Regions, region => region.Kind == DenseBlockRegionKind.Subtotals);
        Assert.Contains(plan.Blocks[0].Regions, region => region.Kind == DenseBlockRegionKind.GrandTotals);
        Assert.Contains(plan.Checks, check => check.Kind == ReportCheckKind.NoTruncation && check.Mandatory);
        Assert.Contains(plan.Checks, check => check.Kind == ReportCheckKind.TotalPreservation && check.Mandatory);
        Assert.False(plan.Source.TruncationAllowed);
        var weighted = plan.Blocks[0].Pivot.Values.Single(value => value.MeasureId == "weighted_rate");
        Assert.True(weighted.RequiresPostAggregationCalculation);
        Assert.Contains(weighted.AggregateComponents, component => component.Role == AggregateComponentRole.WeightedNumerator
            && component.Field == "WeightedUnits");
        Assert.Contains(weighted.AggregateComponents, component => component.Role == AggregateComponentRole.WeightedDenominator
            && component.Field == "Weight");
        Assert.Contains("\"WeightedUnits\"", plan.Source.PowerQueryM);
        Assert.Contains("Number.From([#\"Units\"]) * Number.From([#\"Weight\"])", plan.Source.PowerQueryM);
        Assert.Contains(plan.Blocks[0].Pivot.Filters, filter => filter.IsSupportingField && filter.Field == "Period");
        Assert.Contains(plan.Blocks[0].Pivot.Filters, filter => filter.IsSupportingField && filter.Field == "Units");
        Assert.Contains(plan.Blocks[0].Pivot.Filters, filter => filter.IsSupportingField && filter.Field == "Weight");
        Assert.Equal(spec.Styles.Count, plan.Styles.Count);
        Assert.Matches("^[0-9a-f]{64}$", plan.SpecificationHash);
        Assert.Matches("^[0-9a-f]{64}$", plan.PlanHash);
        Assert.Equal(plan.PlanHash, ReportBuildPlanDigest.Compute(plan));
        Assert.Equal(ReportSpecV1.CurrentSchemaVersion, plan.SchemaVersion);
        Assert.Equal(4, plan.Blocks[0].Presentation.ResolvedPeriodSlices.Count);
        Assert.Equal(new DateTime(2026, 2, 1), plan.Blocks[0].Presentation.ResolvedPeriodSlices.Single(slice => slice.Id == "prior").StartInclusive);
        Assert.Equal(
            new[]
            {
                PivotMemberStageKind.ApplyMemberOrder,
                PivotMemberStageKind.GroupMembers,
                PivotMemberStageKind.ApplyTopN,
                PivotMemberStageKind.AggregateOthers
            },
            plan.Blocks[0].Pivot.Rows[0].MemberStages);
    }

    [Theory]
    [InlineData(ReportOutputMode.StandardMatrix)]
    [InlineData(ReportOutputMode.MetricStack)]
    [InlineData(ReportOutputMode.DenseGrid)]
    public void Carries_each_explicit_output_mode_into_the_build_plan(ReportOutputMode outputMode)
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].OutputMode = outputMode;

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Equal(outputMode, plan.Blocks[0].OutputMode);
        var expectedOperation = outputMode switch
        {
            ReportOutputMode.StandardMatrix => BuildOperationKind.RenderStandardMatrix,
            ReportOutputMode.MetricStack => BuildOperationKind.RenderMetricStack,
            ReportOutputMode.DenseGrid => BuildOperationKind.RenderDenseGrid,
            _ => throw new ArgumentOutOfRangeException(nameof(outputMode))
        };
        Assert.Contains(plan.Operations, operation => operation.Kind == expectedOperation
            && operation.OwnershipId == spec.Blocks[0].OwnershipId);
    }

    [Fact]
    public void Data_model_routing_flows_into_every_pivot_plan()
    {
        var plan = ReportBuildPlanner.Create(
            SyntheticReportFactory.CreateValidLongSpec(),
            SyntheticReportFactory.CreateLongProfile(2_000_000));

        Assert.Equal(SourceLoadRoute.DataModel, plan.Source.Route);
        Assert.All(plan.Blocks, block => Assert.True(block.Pivot.UseDataModel));
        Assert.Contains(plan.Operations, operation => operation.Kind == BuildOperationKind.LoadDataModel);
    }

    [Fact]
    public void Distinct_count_forces_data_model_routing_below_the_row_limit()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Measures.Add(new MeasureDefinition
        {
            Id = "distinct_regions",
            Label = "Distinct regions",
            ValueType = MeasureValueType.WholeNumber,
            Expression = new AggregateMeasureExpression
            {
                Field = "Region",
                Function = AggregateFunction.DistinctCount,
                ResultType = MeasureValueType.WholeNumber
            }
        });
        spec.Blocks[1].Layout.Values.Add(new ValuePlacementSpec { MeasureId = "distinct_regions" });

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile(100));

        Assert.Equal(SourceLoadRoute.DataModel, plan.Source.Route);
        Assert.All(plan.Blocks, block => Assert.True(block.Pivot.UseDataModel));
    }

    [Fact]
    public void Row_removing_transforms_use_an_upper_bound_projection_check()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Transforms.Insert(0, new FilterRowsTransform
        {
            Id = "nonnegative_rows",
            Column = "Amount",
            Operator = RowFilterOperator.GreaterThanOrEqual,
            Value = ScalarValue.FromNumber(0m)
        });

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        var rowCheck = plan.Checks.Single(check => check.Id == "mandatory-no-truncation");
        Assert.Equal(RowCountExpectation.AtMostProjection, rowCheck.RowCountExpectation);
        Assert.DoesNotContain(plan.Checks, check =>
            check.Mandatory &&
            check.Kind == ReportCheckKind.TotalPreservation &&
            check.EvaluationScope == CheckEvaluationScope.CanonicalData);
        Assert.Contains(plan.Checks, check =>
            check.Id == "mandatory-rendered-output-reconciliation" &&
            check.EvaluationScope == CheckEvaluationScope.RenderedOutput);
    }

    [Fact]
    public void Routes_the_exact_worksheet_data_boundary_without_truncation()
    {
        var worksheet = ReportBuildPlanner.Create(
            SyntheticReportFactory.CreateValidLongSpec(),
            SyntheticReportFactory.CreateLongProfile(RowProjection.MaximumWorksheetDataRows));
        var dataModel = ReportBuildPlanner.Create(
            SyntheticReportFactory.CreateValidLongSpec(),
            SyntheticReportFactory.CreateLongProfile(RowProjection.MaximumWorksheetDataRows + 1));

        Assert.Equal(SourceLoadRoute.Worksheet, worksheet.Source.Route);
        Assert.Equal(SourceLoadRoute.DataModel, dataModel.Source.Route);
        Assert.False(worksheet.Source.TruncationAllowed);
        Assert.False(dataModel.Source.TruncationAllowed);
    }

    [Fact]
    public void Adds_a_hidden_supporting_value_for_an_undisplayed_top_n_measure()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].Layout.Values.RemoveAll(value => value.MeasureId == "amount");

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(plan.Blocks[0].Pivot.SupportingValues, value => value.MeasureId == "amount");
    }

    [Fact]
    public void Filtered_aggregates_are_explicit_post_aggregation_calculations()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[1].Layout.Rows[0].Field = "Region";
        spec.Blocks[1].Layout.Values.Add(new ValuePlacementSpec { MeasureId = "filtered_amount" });

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.True(plan.Blocks[1].Pivot.Values.Single(value => value.MeasureId == "filtered_amount")
            .RequiresPostAggregationCalculation);
    }

    [Fact]
    public void Specification_hash_changes_when_content_changes_without_changing_ids()
    {
        var original = SyntheticReportFactory.CreateValidLongSpec();
        var changed = SyntheticReportFactory.CreateValidLongSpec();
        changed.Name = "Changed title";

        var first = ReportBuildPlanner.Create(original, SyntheticReportFactory.CreateLongProfile());
        var second = ReportBuildPlanner.Create(changed, SyntheticReportFactory.CreateLongProfile());

        Assert.Equal(first.SpecificationId, second.SpecificationId);
        Assert.NotEqual(first.SpecificationHash, second.SpecificationHash);
    }

    [Fact]
    public void Canonicalizes_lowercase_a1_anchors_for_host_execution()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].AnchorCell = "$b$3";

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Equal("$B$3", plan.Blocks[0].AnchorCell);
        Assert.Equal("$B$3", plan.Blocks[0].OwnedRange.AnchorCell);
    }

    [Fact]
    public void Plan_hash_detects_mutation_after_validation()
    {
        var plan = ReportBuildPlanner.Create(
            SyntheticReportFactory.CreateValidLongSpec(),
            SyntheticReportFactory.CreateLongProfile());

        plan.Source.PowerQueryM += " ";

        Assert.NotEqual(plan.PlanHash, ReportBuildPlanDigest.Compute(plan));
    }

    [Fact]
    public void Plans_check_only_measures_without_requiring_a_displayed_value()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Checks.Add(new ReportCheckSpec
        {
            Id = "filtered_values_present",
            Kind = ReportCheckKind.RequiredValues,
            MeasureId = "filtered_amount"
        });

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());
        var check = plan.Checks.Single(item => item.Id == "filtered_values_present");

        Assert.NotNull(check.Measure);
        Assert.Equal("filtered_amount", check.Measure!.MeasureId);
        Assert.True(check.Measure.RequiresPostAggregationCalculation);
        Assert.Contains(check.Measure.AggregateComponents, component => component.Filters.Count == 1);
    }
}
