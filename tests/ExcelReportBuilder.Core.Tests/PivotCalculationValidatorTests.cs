using ExcelReportBuilder.Core.PivotPlus.Calculations;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotCalculationValidatorTests
{
    [Fact]
    public void Accepts_forward_references_while_preserving_a_separate_display_order()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "variance",
                "Variance",
                new PivotDifferenceExpression(
                    new PivotMeasureReferenceExpression("actual"),
                    new PivotMeasureReferenceExpression("plan"))),
            PivotCalculationTestFactory.Measure("actual", "Actual", PivotCalculationTestFactory.Sum()),
            PivotCalculationTestFactory.Measure(
                "plan", "Plan", new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum))
        });

        ValidationResult validation = PivotCalculationValidator.Validate(definition);
        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.True(validation.IsValid, Format(validation));
        Assert.Equal(new[] { "variance", "actual", "plan" },
            compilation.Measures.Select(measure => measure.DefinitionId));
        Assert.Equal(new[] { "actual", "plan", "variance" },
            compilation.CreationSequence.Select(measure => measure.DefinitionId));
    }

    [Fact]
    public void Rejects_unknown_references_and_dependency_cycles()
    {
        PivotMeasureSetDefinition unknown = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "a", "A", new PivotMeasureReferenceExpression("missing"))
        });
        PivotMeasureSetDefinition cycle = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure("a", "A", new PivotMeasureReferenceExpression("b")),
            PivotCalculationTestFactory.Measure("b", "B", new PivotMeasureReferenceExpression("a"))
        });

        ValidationResult unknownResult = PivotCalculationValidator.Validate(unknown);
        ValidationResult cycleResult = PivotCalculationValidator.Validate(cycle);

        Assert.Contains(unknownResult.Issues, issue => issue.Code == "PIVOT_CALC_REFERENCE_UNKNOWN");
        Assert.Contains(cycleResult.Issues, issue => issue.Code == "PIVOT_CALC_REFERENCE_CYCLE");
        Assert.Throws<InvalidPivotCalculationException>(() => PivotDaxCompiler.Compile(cycle));
    }

    [Fact]
    public void Rejects_yearly_source_for_q1_without_inference()
    {
        PivotPeriodDefinition monthly = PivotCalculationTestFactory.MonthlyPeriods();
        var yearly = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "period",
                PivotPeriodGrain.Year,
                PivotPeriodCoverageStatus.Complete,
                new[]
                {
                    new PivotPeriodCoverageMember(
                        new PivotPeriodPoint(PivotPeriodGrain.Year, 2026),
                        PivotFilterValue.FromMember("jan"),
                        new[] { "actual", "plan" })
                },
                scenarioFieldId: "scenario"),
            monthly.Slices);
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "q1", "Q1", PivotCalculationTestFactory.Sum("actual_q1"))
            },
            yearly);

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_PERIOD_GRAIN_TOO_COARSE");
    }

    [Fact]
    public void Rejects_monthly_q1_when_any_required_month_is_missing()
    {
        PivotPeriodDefinition periods = PivotCalculationTestFactory.MonthlyPeriods(includeMarch: false);
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "q1", "Q1", PivotCalculationTestFactory.Sum("actual_q1"))
            },
            periods);

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_PERIOD_SLICE_COVERAGE_MISSING");
    }

    [Fact]
    public void Rejects_partial_or_unknown_period_coverage()
    {
        PivotPeriodDefinition valid = PivotCalculationTestFactory.MonthlyPeriods();
        var partial = new PivotPeriodDefinition(
            new PivotPeriodSource(
                valid.Source.PeriodFieldId,
                valid.Source.SourceGrain,
                PivotPeriodCoverageStatus.Partial,
                valid.Source.Coverage,
                valid.Source.PeriodContextFieldIds,
                valid.Source.ScenarioFieldId,
                valid.Source.ScenarioContextFieldIds),
            valid.Slices);

        ValidationResult result = PivotCalculationValidator.Validate(
            PivotCalculationTestFactory.Set(
                new[] { PivotCalculationTestFactory.Measure("q1", "Q1", PivotCalculationTestFactory.Sum("actual_q1")) },
                partial));

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_PERIOD_COVERAGE_INCOMPLETE");
    }

    [Fact]
    public void Rejects_date_rollup_without_explicit_calendar_or_continuous_range_evidence()
    {
        var periods = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "transaction_date",
                PivotPeriodGrain.Date,
                PivotPeriodCoverageStatus.Complete,
                new[]
                {
                    new PivotPeriodCoverageMember(
                        new PivotPeriodPoint(PivotPeriodGrain.Date, 2026, date: new DateTime(2026, 1, 1)),
                        PivotFilterValue.FromScalar(PivotScalarValue.Date(new DateTime(2026, 1, 1))))
                }),
            new[]
            {
                new PivotPeriodSlice(
                    "jan",
                    "Jan",
                    new PivotPeriodPoint(PivotPeriodGrain.Month, 2026, 1),
                    null,
                    PivotSliceFilterMode.ReplaceAxisContext)
            });

        ValidationResult result = PivotCalculationValidator.Validate(
            PivotCalculationTestFactory.Set(
                new[] { PivotCalculationTestFactory.Measure("jan", "Jan", PivotCalculationTestFactory.Sum("jan")) },
                periods));

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_DATE_COVERAGE_MODE_REQUIRED");
    }

    [Fact]
    public void Rejects_missing_scenario_coverage_for_one_period_bucket()
    {
        PivotPeriodDefinition valid = PivotCalculationTestFactory.MonthlyPeriods();
        PivotPeriodCoverageMember[] coverage = valid.Source.Coverage
            .Select((member, index) => index == 1
                ? new PivotPeriodCoverageMember(member.Point, member.SourceValue, new[] { "actual" })
                : member)
            .ToArray();
        var periods = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "period",
                PivotPeriodGrain.Month,
                PivotPeriodCoverageStatus.Complete,
                coverage,
                scenarioFieldId: "scenario"),
            valid.Slices);

        ValidationResult result = PivotCalculationValidator.Validate(
            PivotCalculationTestFactory.Set(
                new[] { PivotCalculationTestFactory.Measure("plan", "Plan", PivotCalculationTestFactory.Sum("plan_q1")) },
                periods));

        Assert.Contains(result.Issues, issue =>
            issue.Code == "PIVOT_CALC_PERIOD_SCENARIO_COVERAGE_MISSING");
    }

    [Fact]
    public void Rejects_unknown_members_scalar_type_mismatch_and_filter_arity()
    {
        var filters = new[]
        {
            new PivotCalculationFilter(
                "region",
                PivotCalculationFilterOperator.Equal,
                new[] { PivotFilterValue.FromMember("unknown") }),
            new PivotCalculationFilter(
                "units",
                PivotCalculationFilterOperator.Equal,
                new[] { PivotFilterValue.FromScalar(PivotScalarValue.Text("ten")) }),
            new PivotCalculationFilter(
                "amount",
                PivotCalculationFilterOperator.GreaterThan,
                Array.Empty<PivotFilterValue>())
        };
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "filtered",
                "Filtered",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    filters))
        });

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_FILTER_MEMBER_UNKNOWN");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_FILTER_VALUE_TYPE_MISMATCH");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_FILTER_ARITY_INVALID");
    }

    [Fact]
    public void Rejects_path_like_schema_identifiers_but_allows_escapable_dax_punctuation()
    {
        var schema = new PivotModelSchema(new[]
        {
            new PivotModelTableSchema(
                "bad",
                @"C:\private\Fact",
                new[]
                {
                    new PivotModelFieldSchema("amount", "Amount", PivotModelDataType.DecimalNumber)
                }),
            new PivotModelTableSchema(
                "safe",
                "Fact 'Quoted'",
                new[]
                {
                    new PivotModelFieldSchema("safe_amount", "Amount] Net", PivotModelDataType.DecimalNumber)
                })
        });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "metric",
                    "Metric",
                    new PivotAggregateExpression("safe_amount", PivotCalculationAggregateFunction.Sum),
                    homeTableId: "safe"),
                PivotCalculationTestFactory.Measure(
                    "bad_caption",
                    @"C:\private\Metric",
                    new PivotAggregateExpression("safe_amount", PivotCalculationAggregateFunction.Sum),
                    homeTableId: "safe")
            },
            schema: schema);

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_TABLE_NAME_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_MEASURE_CAPTION_INVALID");
        Assert.DoesNotContain(result.Issues, issue =>
            issue.Code == "PIVOT_CALC_FIELD_NAME_INVALID" && issue.Path.Contains("safe"));
    }

    [Fact]
    public void Rejects_duplicate_measure_ids_and_native_names_without_regard_to_case()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure("metric", "Revenue", PivotCalculationTestFactory.Sum()),
            PivotCalculationTestFactory.Measure("METRIC", "revenue", PivotCalculationTestFactory.Sum())
        });

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_MEASURE_ID_DUPLICATE");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_MEASURE_CAPTION_DUPLICATE");
    }

    [Fact]
    public void Rejects_semantic_format_mismatch_and_unbounded_currency_format()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "ratio",
                "Ratio",
                new PivotSafeRatioExpression(
                    PivotCalculationTestFactory.Sum(),
                    new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Decimal()),
            PivotCalculationTestFactory.Measure(
                "currency",
                "Currency",
                PivotCalculationTestFactory.Sum(),
                new PivotMeasureFormat(PivotMeasureFormatKind.Currency, 2, true, "EUR/USD"))
        });

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_FORMAT_SEMANTIC_MISMATCH");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_FORMAT_CURRENCY_INVALID");
    }

    [Fact]
    public void Rejects_cross_table_weighted_result()
    {
        PivotModelSchema original = PivotCalculationTestFactory.Schema();
        PivotModelTableSchema product = original.Tables[1];
        var schema = new PivotModelSchema(new[]
        {
            original.Tables[0],
            new PivotModelTableSchema(
                product.Id,
                product.Name,
                product.Fields.Concat(new[]
                {
                    new PivotModelFieldSchema("other_weight", "Other Weight", PivotModelDataType.DecimalNumber)
                }))
        });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "weighted",
                    "Weighted",
                    new PivotWeightedResultExpression(
                        "amount",
                        "other_weight",
                        PivotDenominatorBehavior.Blank))
            },
            schema: schema);

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_WEIGHTED_TABLE_MISMATCH");
    }

    [Fact]
    public void Rejects_empty_or_duplicate_parent_and_filtered_total_scope_fields()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "parent",
                "Parent",
                new PivotShareExpression(
                    PivotCalculationTestFactory.Sum(),
                    new PivotParentShareDenominator(Array.Empty<string>()),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage()),
            PivotCalculationTestFactory.Measure(
                "total",
                "Total",
                new PivotShareExpression(
                    PivotCalculationTestFactory.Sum(),
                    new PivotFilteredTotalShareDenominator(new[] { "region", "REGION" }),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage())
        });

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_SHARE_SCOPE_FIELD_REQUIRED");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_SHARE_SCOPE_FIELD_DUPLICATE");
    }

    [Fact]
    public void Rejects_unbounded_measures_expression_depth_and_filter_values()
    {
        PivotMeasureDefinition[] tooMany = Enumerable.Range(0, 129)
            .Select(index => PivotCalculationTestFactory.Measure(
                "m" + index,
                "M " + index,
                PivotCalculationTestFactory.Sum()))
            .ToArray();
        PivotCalculationExpression deep = PivotCalculationTestFactory.Sum();
        for (var index = 0; index < 33; index++)
        {
            deep = new PivotDifferenceExpression(deep, PivotCalculationTestFactory.Sum());
        }

        PivotCalculationExpression wide = PivotCalculationTestFactory.Sum();
        for (var index = 0; index < 8; index++)
        {
            wide = new PivotDifferenceExpression(wide, wide);
        }

        PivotFilterValue[] values = Enumerable.Range(0, 257)
            .Select(index => PivotFilterValue.FromScalar(PivotScalarValue.Text("Member " + index)))
            .ToArray();
        PivotMeasureSetDefinition bounded = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure("deep", "Deep", deep),
            PivotCalculationTestFactory.Measure("wide", "Wide", wide),
            PivotCalculationTestFactory.Measure(
                "filtered",
                "Filtered",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    new[]
                    {
                        new PivotCalculationFilter("region", PivotCalculationFilterOperator.In, values)
                    }))
            ,
            PivotCalculationTestFactory.Measure(
                "many_filters",
                "Many Filters",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    Enumerable.Range(0, 33).Select(index =>
                        new PivotCalculationFilter(
                            "region",
                            PivotCalculationFilterOperator.Equal,
                            new[]
                            {
                                PivotFilterValue.FromScalar(PivotScalarValue.Text("Value " + index))
                            }))))
        });

        ValidationResult measureResult = PivotCalculationValidator.Validate(
            PivotCalculationTestFactory.Set(tooMany));
        ValidationResult boundedResult = PivotCalculationValidator.Validate(bounded);

        Assert.Contains(measureResult.Issues, issue => issue.Code == "PIVOT_CALC_MEASURE_LIMIT");
        Assert.Contains(boundedResult.Issues, issue => issue.Code == "PIVOT_CALC_EXPRESSION_DEPTH_LIMIT");
        Assert.Contains(boundedResult.Issues, issue => issue.Code == "PIVOT_CALC_EXPRESSION_NODE_LIMIT");
        Assert.Contains(boundedResult.Issues, issue => issue.Code == "PIVOT_CALC_FILTER_LIMIT");
        Assert.Contains(boundedResult.Issues, issue => issue.Code == "PIVOT_CALC_FILTER_VALUE_LIMIT");
    }

    [Fact]
    public void Null_collection_elements_fail_as_validation_issues_not_null_reference_exceptions()
    {
        var schema = new PivotModelSchema(new PivotModelTableSchema[] { null! });
        var definition = new PivotMeasureSetDefinition(
            schema,
            new PivotMeasureDefinition[] { null! });

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_SCHEMA_TABLE_NULL");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_MEASURE_NULL");
    }

    [Fact]
    public void Null_context_and_share_field_entries_fail_closed_without_throwing()
    {
        PivotPeriodDefinition valid = PivotCalculationTestFactory.MonthlyPeriods();
        var periods = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "period",
                PivotPeriodGrain.Month,
                PivotPeriodCoverageStatus.Complete,
                valid.Source.Coverage,
                new string[] { null! },
                "scenario",
                new[] { "scenario" }),
            valid.Slices);
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "share",
                    "Share",
                    new PivotShareExpression(
                        PivotCalculationTestFactory.Sum(),
                        new PivotParentShareDenominator(new string[] { null! }),
                        PivotDenominatorBehavior.Blank),
                    PivotCalculationTestFactory.Percentage())
            },
            periods);

        ValidationResult result = PivotCalculationValidator.Validate(definition);

        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_CONTEXT_FIELD_ID_INVALID");
        Assert.Contains(result.Issues, issue => issue.Code == "PIVOT_CALC_SHARE_SCOPE_FIELD_ID_INVALID");
    }

    [Fact]
    public void Contract_collections_are_defensively_copied()
    {
        var members = new List<PivotModelMember>
        {
            new PivotModelMember("north", PivotScalarValue.Text("North"))
        };
        var fields = new List<PivotModelFieldSchema>
        {
            new PivotModelFieldSchema("region", "Region", PivotModelDataType.Text, members)
        };
        var tables = new List<PivotModelTableSchema>
        {
            new PivotModelTableSchema("fact", "Fact", fields)
        };
        var measures = new List<PivotMeasureDefinition>
        {
            PivotCalculationTestFactory.Measure(
                "count",
                "Count",
                new PivotAggregateExpression("region", PivotCalculationAggregateFunction.Count),
                PivotCalculationTestFactory.Whole())
        };
        var definition = new PivotMeasureSetDefinition(new PivotModelSchema(tables), measures);

        members.Clear();
        fields.Clear();
        tables.Clear();
        measures.Clear();

        Assert.Single(definition.Schema.Tables);
        Assert.Single(definition.Schema.Tables[0].Fields);
        Assert.Single(definition.Schema.Tables[0].Fields[0].Members);
        Assert.Single(definition.Measures);
    }

    private static string Format(ValidationResult result)
    {
        return string.Join(Environment.NewLine, result.Issues.Select(issue =>
            issue.Code + " " + issue.Path + " " + issue.Message));
    }
}
