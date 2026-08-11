using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class ReportSpecValidationTests
{
    [Fact]
    public void Accepts_a_dense_multi_block_spec_with_typed_measures_and_layout_controls()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(Format)));
        Assert.Equal(2, spec.Blocks.Count);
        Assert.Contains(spec.Measures, measure => measure.Expression is SafeDivideMeasureExpression);
        Assert.Contains(spec.Measures, measure => measure.Expression is ShareMeasureExpression);
        Assert.Contains(spec.Measures, measure => measure.Expression is WeightedAggregateMeasureExpression);
        Assert.Contains(spec.Measures, measure => measure.Expression is RatioMeasureExpression);
        Assert.Contains(spec.Measures, measure => measure.Expression is FilteredAggregateMeasureExpression);
    }

    [Fact]
    public void Accepts_bounded_normalization_for_a_long_text_period_column()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping!.Grain = PeriodGrain.Month;
        spec.PeriodMapping.ReportingYear = 2026;
        spec.Transforms.Insert(0, new NormalizePeriodsTransform
        {
            Id = "normalize_long_periods",
            PeriodMappingId = spec.PeriodMapping.Id
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(Format)));
    }

    [Fact]
    public void Rejects_a_long_month_column_without_a_reporting_year_before_execution()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping!.Grain = PeriodGrain.Month;
        spec.Transforms.Insert(0, new NormalizePeriodsTransform
        {
            Id = "normalize_long_periods",
            PeriodMappingId = spec.PeriodMapping.Id
        });
        var profile = SyntheticReportFactory.CreateLongProfile();
        SourceColumnProfile period = profile.Columns[0];
        period.InferredType = SourceValueType.Text;
        period.DateLikeCount = 0;
        period.PeriodLikeWithoutYearCount = period.NonBlankCount;
        period.DayGrainCount = 0;
        period.MonthGrainCount = period.NonBlankCount;

        var result = ReportSpecValidator.Validate(spec, profile);

        Assert.Contains(result.Issues, issue => issue.Code == "REPORTING_YEAR_REQUIRED");
    }

    [Fact]
    public void Rejects_long_period_normalization_without_an_explicit_grain()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Transforms.Insert(0, new NormalizePeriodsTransform
        {
            Id = "normalize_long_periods",
            PeriodMappingId = spec.PeriodMapping!.Id
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "LONG_PERIOD_GRAIN_REQUIRED");
    }

    [Fact]
    public void Rejects_missing_year_in_an_explicit_wide_mapping()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MonthHeaders,
            KeyColumns = { "Region" },
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Jan", Month = 1 },
                new PeriodColumnMapping { SourceColumn = "Feb", Month = 2 }
            }
        };
        spec.Transforms.Add(new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" });

        var result = ReportSpecValidator.Validate(spec);

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "REPORTING_YEAR_REQUIRED");
    }

    [Fact]
    public void Rejects_non_start_months_in_quarter_mappings()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MonthHeaders,
            Grain = PeriodGrain.Quarter,
            ReportingYear = 2026,
            KeyColumns = { "Region" },
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Amount", Month = 2 }
            }
        };
        spec.Transforms.Insert(0, new NormalizePeriodsTransform
        {
            Id = "normalize",
            PeriodMappingId = "periods"
        });

        var result = ReportSpecValidator.Validate(spec);

        Assert.Contains(result.Issues, issue => issue.Code == "QUARTER_START_MONTH_INVALID");
    }

    [Fact]
    public void Rejects_measure_reference_cycles()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Measures.Clear();
        spec.Measures.Add(new MeasureDefinition
        {
            Id = "first",
            Label = "First",
            ValueType = MeasureValueType.Number,
            Expression = new ReferenceMeasureExpression
            {
                MeasureId = "second",
                ResultType = MeasureValueType.Number
            }
        });
        spec.Measures.Add(new MeasureDefinition
        {
            Id = "second",
            Label = "Second",
            ValueType = MeasureValueType.Number,
            Expression = new ReferenceMeasureExpression
            {
                MeasureId = "first",
                ResultType = MeasureValueType.Number
            }
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "MEASURE_REFERENCE_CYCLE");
    }

    [Fact]
    public void Validates_percentage_and_percentage_point_difference_types()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Measures.Add(new MeasureDefinition
        {
            Id = "percentage_points",
            Label = "Percentage point difference",
            ValueType = MeasureValueType.Percentage,
            Expression = new DifferenceMeasureExpression
            {
                DifferenceKind = DifferenceKind.PercentagePoints,
                ResultType = MeasureValueType.Percentage,
                Current = new ConstantMeasureExpression
                {
                    ResultType = MeasureValueType.Percentage,
                    Value = 0.25m
                },
                Baseline = new ConstantMeasureExpression
                {
                    ResultType = MeasureValueType.Percentage,
                    Value = 0.20m
                }
            }
        });
        spec.Measures.Add(new MeasureDefinition
        {
            Id = "bad_percentage_points",
            Label = "Bad percentage point difference",
            ValueType = MeasureValueType.Percentage,
            Expression = new DifferenceMeasureExpression
            {
                DifferenceKind = DifferenceKind.PercentagePoints,
                ResultType = MeasureValueType.Percentage,
                Current = new ConstantMeasureExpression { ResultType = MeasureValueType.Number, Value = 25m },
                Baseline = new ConstantMeasureExpression { ResultType = MeasureValueType.Number, Value = 20m }
            }
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "PERCENTAGE_POINT_TYPE_MISMATCH");
    }

    [Fact]
    public void Total_row_exclusion_requires_observed_evidence()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Transforms.Insert(0, new ExcludeTotalRowsTransform
        {
            Id = "exclude_totals",
            Evidence =
            {
                new TotalRowEvidenceSpec
                {
                    Column = "Region",
                    MatchKind = TotalRowMatchKind.EqualsAny,
                    Values = { ScalarValue.FromText("Total") },
                    Source = EvidenceSource.Profile,
                    ObservedMatchCount = 0
                }
            }
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "TOTAL_ROW_MATCH_COUNT_REQUIRED");
    }

    [Fact]
    public void Warns_instead_of_truncating_when_data_model_is_required()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile(2_000_000));

        Assert.True(result.IsValid, string.Join(Environment.NewLine, result.Issues.Select(Format)));
        Assert.Contains(result.Issues, issue => issue.Code == "DATA_MODEL_REQUIRED"
            && issue.Severity == ValidationSeverity.Warning);
    }

    [Fact]
    public void Rejects_overlapping_owned_ranges_on_the_same_worksheet()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[1].AnchorCell = "F3";

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "BLOCK_OWNED_RANGE_OVERLAP");
    }

    [Fact]
    public void Rejects_owned_ranges_that_run_past_worksheet_bounds()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].AnchorCell = "XFD1048576";
        spec.Blocks[0].OwnedExtent = new OwnedRangeExtentSpec { RowCount = 2, ColumnCount = 1 };

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "OWNED_EXTENT_OUT_OF_BOUNDS");
    }

    [Fact]
    public void Rejects_block_ownership_that_reuses_report_ownership()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].OwnershipId = spec.OwnershipId;

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "BLOCK_OWNERSHIP_ID_DUPLICATE");
    }

    [Fact]
    public void Rejects_missing_absolute_slice_dates_and_reference_cycles()
    {
        var missingDates = SyntheticReportFactory.CreateValidLongSpec();
        missingDates.Blocks[0].PeriodSlices[0].SelectedStart = null;

        var missingResult = ReportSpecValidator.Validate(missingDates, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(missingResult.Issues, issue => issue.Code == "ABSOLUTE_SLICE_DATES_REQUIRED");

        var cycle = SyntheticReportFactory.CreateValidLongSpec();
        var current = cycle.Blocks[0].PeriodSlices.Single(slice => slice.Id == "current");
        current.Kind = PeriodSliceKind.Prior;
        current.SelectedStart = null;
        current.SelectedEnd = null;
        current.BasedOnSliceId = "prior";

        var cycleResult = ReportSpecValidator.Validate(cycle, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(cycleResult.Issues, issue => issue.Code == "PERIOD_SLICE_REFERENCE_CYCLE");
    }

    [Fact]
    public void Rejects_ambiguous_measure_and_placement_slice_contexts()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        ((AggregateMeasureExpression)spec.Measures.Single(measure => measure.Id == "amount").Expression)
            .PeriodSliceId = "current";

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "VALUE_SLICE_CONTEXT_CONFLICT");
    }

    [Fact]
    public void Rejects_invalid_filtered_aggregate_arity_and_duplicates()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var expression = (FilteredAggregateMeasureExpression)spec.Measures
            .Single(measure => measure.Id == "filtered_amount").Expression;
        expression.Filters[0].Values.Clear();
        expression.Filters.Add(new MeasureFilterSpec
        {
            Field = "Category",
            Operator = MeasureFilterOperator.Equal,
            Values = { ScalarValue.FromText("Core") }
        });
        expression.Filters.Add(new MeasureFilterSpec
        {
            Field = "Category",
            Operator = MeasureFilterOperator.Equal,
            Values = { ScalarValue.FromText("Core") }
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "MEASURE_FILTER_SINGLE_VALUE_REQUIRED");
        Assert.Contains(result.Issues, issue => issue.Code == "MEASURE_FILTER_DUPLICATE");
    }

    [Fact]
    public void Rejects_duplicate_value_map_inputs()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Transforms.Insert(0, new MapValuesTransform
        {
            Id = "map_categories",
            Column = "Category",
            Entries =
            {
                new ValueMapEntry { From = ScalarValue.FromText("A"), To = ScalarValue.FromText("B") },
                new ValueMapEntry { From = ScalarValue.FromText("A"), To = ScalarValue.FromText("C") }
            }
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "MAP_INPUT_DUPLICATE");
    }

    [Fact]
    public void Rejects_duplicate_group_members_and_conflicting_subtotal_presentation()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        var region = spec.Blocks[0].Layout.Rows[0];
        region.GroupBuckets[1].Members.Add(ScalarValue.FromText("North"));
        var category = spec.Blocks[0].Layout.Rows[1];
        category.Subtotals.Label = "Not allowed";

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "GROUP_MEMBER_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "DISABLED_SUBTOTAL_PRESENTATION_NOT_ALLOWED");
    }

    [Fact]
    public void Allows_empty_filter_selection_to_mean_all_members()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].Layout.Filters.Add(new FilterPlacementSpec
        {
            Field = "Amount"
        });

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.DoesNotContain(result.Issues, issue => issue.Code == "FILTER_SELECTION_EMPTY");
        Assert.True(
            result.IsValid,
            string.Join("; ", result.Issues.Select(issue => issue.Code + ": " + issue.Message)));
    }

    [Fact]
    public void Rejects_fields_on_multiple_axes_and_unrankable_top_n_measures()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Blocks[0].Layout.Filters.Add(new FilterPlacementSpec
        {
            Field = "Region",
            SelectedValues = { ScalarValue.FromText("North") }
        });
        spec.Blocks[0].Layout.Rows[0].TopN!.MeasureId = "average_price";

        var result = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(result.Issues, issue => issue.Code == "FIELD_USED_ON_MULTIPLE_AXES");
        Assert.Contains(result.Issues, issue => issue.Code == "TOP_N_MEASURE_NOT_RANKABLE");
    }

    [Fact]
    public void Rejects_period_output_key_collisions_and_removed_period_fields()
    {
        var collision = SyntheticReportFactory.CreateValidLongSpec();
        collision.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MonthHeaders,
            ReportingYear = 2026,
            KeyColumns = { "Period" },
            PeriodColumnName = "Period",
            Columns = { new PeriodColumnMapping { SourceColumn = "Amount", Month = 1 } }
        };
        collision.Transforms.Add(new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" });

        var collisionResult = ReportSpecValidator.Validate(collision);

        Assert.Contains(collisionResult.Issues, issue => issue.Code == "PERIOD_OUTPUT_KEY_COLLISION");

        var removed = SyntheticReportFactory.CreateValidLongSpec();
        removed.Transforms.Add(new RemoveColumnsTransform { Id = "remove_period", Columns = { "Period" } });

        var removedResult = ReportSpecValidator.Validate(removed, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(removedResult.Issues, issue => issue.Code == "PERIOD_FIELD_REMOVED_BY_TRANSFORM");
    }

    [Fact]
    public void Warns_for_large_unmapped_sources_that_require_the_data_model()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.PeriodMapping = null;
        spec.Blocks[0].PeriodSlices.Clear();
        foreach (var value in spec.Blocks[0].Layout.Values)
        {
            value.PeriodSliceIds.Clear();
        }

        var result = ReportSpecValidator.Validate(
            spec,
            SyntheticReportFactory.CreateLongProfile(RowProjection.MaximumWorksheetDataRows + 1));

        Assert.Contains(result.Issues, issue => issue.Code == "DATA_MODEL_REQUIRED"
            && issue.Severity == ValidationSeverity.Warning);
    }

    private static string Format(ValidationIssue issue)
    {
        return $"{issue.Code} {issue.Path}: {issue.Message}";
    }
}
