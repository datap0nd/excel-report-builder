using ExcelReportBuilder.Core.PivotPlus.Calculations;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PivotDaxCompilerTests
{
    [Fact]
    public void Compiles_actual_months_q1_plan_and_q1_variance_in_exact_display_order()
    {
        PivotPeriodDefinition periods = PivotCalculationTestFactory.MonthlyPeriods();
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "actual_jan", "Actual Jan", PivotCalculationTestFactory.Sum("actual_jan")),
                PivotCalculationTestFactory.Measure(
                    "actual_feb", "Actual Feb", PivotCalculationTestFactory.Sum("actual_feb")),
                PivotCalculationTestFactory.Measure(
                    "actual_mar", "Actual Mar", PivotCalculationTestFactory.Sum("actual_mar")),
                PivotCalculationTestFactory.Measure(
                    "plan_q1", "Q1 Plan", PivotCalculationTestFactory.Sum("plan_q1")),
                PivotCalculationTestFactory.Measure(
                    "variance_q1",
                    "Q1 Variance",
                    new PivotVarianceExpression(
                        PivotCalculationTestFactory.Sum("actual_q1"),
                        PivotCalculationTestFactory.Sum("plan_q1"),
                        PivotVarianceConvention.ActualMinusPlan))
            },
            periods);

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal(
            new[] { "Actual Jan", "Actual Feb", "Actual Mar", "Q1 Plan", "Q1 Variance" },
            compilation.Measures.Select(measure => measure.GeneratedMeasureName));
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, compilation.Measures.Select(measure => measure.DisplayOrder));
        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "REMOVEFILTERS('Fact Sales'[Month Start]), " +
            "REMOVEFILTERS('Fact Sales'[Scenario]), " +
            "'Fact Sales'[Month Start] IN { DATE(2026, 1, 1) }, " +
            "'Fact Sales'[Scenario] IN { \"Actual\" })",
            compilation.Measures[0].DaxFormula);
        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "REMOVEFILTERS('Fact Sales'[Month Start]), " +
            "REMOVEFILTERS('Fact Sales'[Scenario]), " +
            "'Fact Sales'[Month Start] IN { DATE(2026, 1, 1), DATE(2026, 2, 1), DATE(2026, 3, 1) }, " +
            "'Fact Sales'[Scenario] IN { \"Plan\" })",
            compilation.Measures[3].DaxFormula);
        Assert.Equal(
            "(CALCULATE(SUM('Fact Sales'[Amount]), " +
            "REMOVEFILTERS('Fact Sales'[Month Start]), " +
            "REMOVEFILTERS('Fact Sales'[Scenario]), " +
            "'Fact Sales'[Month Start] IN { DATE(2026, 1, 1), DATE(2026, 2, 1), DATE(2026, 3, 1) }, " +
            "'Fact Sales'[Scenario] IN { \"Actual\" }) - " +
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "REMOVEFILTERS('Fact Sales'[Month Start]), " +
            "REMOVEFILTERS('Fact Sales'[Scenario]), " +
            "'Fact Sales'[Month Start] IN { DATE(2026, 1, 1), DATE(2026, 2, 1), DATE(2026, 3, 1) }, " +
            "'Fact Sales'[Scenario] IN { \"Plan\" }))",
            compilation.Measures[4].DaxFormula);
    }

    [Theory]
    [InlineData(PivotCalculationAggregateFunction.Sum, "SUM")]
    [InlineData(PivotCalculationAggregateFunction.Count, "COUNTA")]
    [InlineData(PivotCalculationAggregateFunction.DistinctCount, "DISTINCTCOUNT")]
    [InlineData(PivotCalculationAggregateFunction.Average, "AVERAGE")]
    [InlineData(PivotCalculationAggregateFunction.Minimum, "MIN")]
    [InlineData(PivotCalculationAggregateFunction.Maximum, "MAX")]
    public void Compiles_the_closed_aggregate_set(
        PivotCalculationAggregateFunction function,
        string daxFunction)
    {
        PivotMeasureFormat format = function == PivotCalculationAggregateFunction.Count ||
            function == PivotCalculationAggregateFunction.DistinctCount
            ? PivotCalculationTestFactory.Whole()
            : PivotCalculationTestFactory.Decimal();
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "metric",
                "Metric",
                new PivotAggregateExpression("amount", function),
                format)
        });

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(daxFunction + "('Fact Sales'[Amount])", measure.DaxFormula);
    }

    [Fact]
    public void Escapes_bound_identifiers_and_text_literals_that_look_like_dax()
    {
        var schema = new PivotModelSchema(new[]
        {
            new PivotModelTableSchema(
                "fact",
                "Fact 'Sales'); EVALUATE",
                new[]
                {
                    new PivotModelFieldSchema(
                        "amount",
                        "Amount] + [Secret",
                        PivotModelDataType.DecimalNumber),
                    new PivotModelFieldSchema("region", "Region", PivotModelDataType.Text)
                })
        });
        var filter = new PivotCalculationFilter(
            "region",
            PivotCalculationFilterOperator.Equal,
            new[]
            {
                PivotFilterValue.FromScalar(
                    PivotScalarValue.Text("North\"), REMOVEFILTERS(Everything)"))
            });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "safe",
                    "Safe ] Measure",
                    new PivotFilteredAggregateExpression(
                        "amount",
                        PivotCalculationAggregateFunction.Sum,
                        new[] { filter }))
            },
            schema: schema);

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "CALCULATE(SUM('Fact ''Sales''); EVALUATE'[Amount]] + [Secret]), " +
            "KEEPFILTERS('Fact ''Sales''); EVALUATE'[Region] IN { " +
            "\"North\"\"), REMOVEFILTERS(Everything)\" }))",
            measure.DaxFormula);
    }

    [Fact]
    public void Compiles_safe_ratio_zero_behaviors_and_ratio_of_ratios_pp_delta()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "delta",
                "Rate Delta",
                new PivotPercentagePointDeltaExpression(
                    new PivotMeasureReferenceExpression("current_ratio"),
                    new PivotMeasureReferenceExpression("plan_ratio")),
                PivotCalculationTestFactory.PercentagePoints()),
            PivotCalculationTestFactory.Measure(
                "current_ratio",
                "Current Ratio",
                new PivotSafeRatioExpression(
                    PivotCalculationTestFactory.Sum(),
                    new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage()),
            PivotCalculationTestFactory.Measure(
                "plan_ratio",
                "Plan Ratio",
                new PivotSafeRatioExpression(
                    PivotCalculationTestFactory.Sum(),
                    new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum),
                    PivotDenominatorBehavior.Zero),
                PivotCalculationTestFactory.Percentage())
        });

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal(
            "IF(OR(ISBLANK([Current Ratio]), ISBLANK([Plan Ratio])), " +
            "BLANK(), 100 * ([Current Ratio] - [Plan Ratio]))",
            compilation.Measures[0].DaxFormula);
        Assert.Equal(
            "DIVIDE(SUM('Fact Sales'[Amount]), SUM('Fact Sales'[Units]), BLANK())",
            compilation.Measures[1].DaxFormula);
        Assert.Equal(
            "DIVIDE(SUM('Fact Sales'[Amount]), SUM('Fact Sales'[Units]), 0)",
            compilation.Measures[2].DaxFormula);
        Assert.Equal(new[] { "current_ratio", "plan_ratio" },
            compilation.Measures[0].DirectDependencyDefinitionIds);
        Assert.Equal(new[] { "current_ratio", "plan_ratio", "delta" },
            compilation.CreationSequence.Select(measure => measure.DefinitionId));
        Assert.Equal(new[] { 3, 1, 2 }, compilation.Measures.Select(measure => measure.CreationOrder));
    }

    [Fact]
    public void Compiles_parent_and_filtered_total_share_with_exact_clear_scopes()
    {
        var part = new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Sum);
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "parent_share",
                "Parent Share",
                new PivotShareExpression(
                    part,
                    new PivotParentShareDenominator(new[] { "department" }),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage()),
            PivotCalculationTestFactory.Measure(
                "filtered_share",
                "Filtered Share",
                new PivotShareExpression(
                    part,
                    new PivotFilteredTotalShareDenominator(new[] { "department", "family" }),
                    PivotDenominatorBehavior.Zero),
                PivotCalculationTestFactory.Percentage())
        });

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal(
            "DIVIDE(SUM('Fact Sales'[Amount]), " +
            "CALCULATE(SUM('Fact Sales'[Amount]), REMOVEFILTERS('Fact Sales'[Department])), BLANK())",
            compilation.Measures[0].DaxFormula);
        Assert.Equal(
            "DIVIDE(SUM('Fact Sales'[Amount]), " +
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "ALLSELECTED('Dim Product'[Product Family]), " +
            "ALLSELECTED('Fact Sales'[Department])), 0)",
            compilation.Measures[1].DaxFormula);
    }

    [Fact]
    public void Compiles_weighted_result_with_one_shared_filtered_scope()
    {
        var filter = new PivotCalculationFilter(
            "region",
            PivotCalculationFilterOperator.In,
            new[]
            {
                PivotFilterValue.FromMember("south"),
                PivotFilterValue.FromMember("north")
            });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "weighted_price",
                "Weighted Price",
                new PivotWeightedResultExpression(
                    "price",
                    "weight",
                    PivotDenominatorBehavior.Blank,
                    new[] { filter }),
                PivotCalculationTestFactory.Currency())
        });

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "VAR __PivotPlusRows = CALCULATETABLE(" +
            "FILTER('Fact Sales', NOT(ISBLANK('Fact Sales'[Unit Price])) && " +
            "NOT(ISBLANK('Fact Sales'[Weight]))), " +
            "KEEPFILTERS('Fact Sales'[Region] IN { \"North\", \"South\" })) " +
            "RETURN DIVIDE(SUMX(__PivotPlusRows, 'Fact Sales'[Unit Price] * 'Fact Sales'[Weight]), " +
            "SUMX(__PivotPlusRows, 'Fact Sales'[Weight]), BLANK())",
            measure.DaxFormula);
    }

    [Fact]
    public void Compiles_boolean_decimal_and_date_literals_through_typed_filters()
    {
        var filters = new[]
        {
            new PivotCalculationFilter(
                "transaction_date",
                PivotCalculationFilterOperator.LessThan,
                new[]
                {
                    PivotFilterValue.FromScalar(PivotScalarValue.Date(new DateTime(2026, 4, 1)))
                }),
            new PivotCalculationFilter(
                "active",
                PivotCalculationFilterOperator.Equal,
                new[] { PivotFilterValue.FromScalar(PivotScalarValue.Boolean(true)) }),
            new PivotCalculationFilter(
                "amount",
                PivotCalculationFilterOperator.GreaterThanOrEqual,
                new[] { PivotFilterValue.FromScalar(PivotScalarValue.DecimalNumber(12.5m)) })
        };
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "typed",
                "Typed",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    filters))
        });

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "KEEPFILTERS('Fact Sales'[Is Active] IN { TRUE() }), " +
            "KEEPFILTERS(NOT(ISBLANK('Fact Sales'[Amount])) && 'Fact Sales'[Amount] >= 12.5), " +
            "KEEPFILTERS(NOT(ISBLANK('Fact Sales'[Transaction Date])) && " +
            "'Fact Sales'[Transaction Date] < DATE(2026, 4, 1)))",
            measure.DaxFormula);
    }

    [Theory]
    [InlineData(PivotCalculationFilterOperator.Equal, "'Fact Sales'[Region] IN { \"North\" }")]
    [InlineData(PivotCalculationFilterOperator.NotEqual, "NOT ('Fact Sales'[Region] IN { \"North\" })")]
    [InlineData(PivotCalculationFilterOperator.In, "'Fact Sales'[Region] IN { \"North\" }")]
    [InlineData(PivotCalculationFilterOperator.NotIn, "NOT ('Fact Sales'[Region] IN { \"North\" })")]
    [InlineData(PivotCalculationFilterOperator.IsBlank, "ISBLANK('Fact Sales'[Region])")]
    [InlineData(PivotCalculationFilterOperator.IsNotBlank, "NOT(ISBLANK('Fact Sales'[Region]))")]
    public void Compiles_each_bounded_equality_set_and_blank_filter_operator(
        PivotCalculationFilterOperator @operator,
        string predicate)
    {
        IReadOnlyList<PivotFilterValue> values =
            @operator == PivotCalculationFilterOperator.IsBlank ||
            @operator == PivotCalculationFilterOperator.IsNotBlank
                ? Array.Empty<PivotFilterValue>()
                : new[] { PivotFilterValue.FromMember("north") };
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "filtered",
                "Filtered",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    new[] { new PivotCalculationFilter("region", @operator, values) }))
        });

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), KEEPFILTERS(" + predicate + "))",
            measure.DaxFormula);
    }

    [Fact]
    public void Intersecting_period_slice_keeps_current_axis_context_explicitly()
    {
        PivotPeriodDefinition source = PivotCalculationTestFactory.MonthlyPeriods();
        var periods = new PivotPeriodDefinition(
            source.Source,
            new[]
            {
                new PivotPeriodSlice(
                    "actual_jan",
                    "Actual Jan",
                    new PivotPeriodPoint(PivotPeriodGrain.Month, 2026, 1),
                    "actual",
                    PivotSliceFilterMode.IntersectCurrentContext)
            });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "actual_jan", "Actual Jan", PivotCalculationTestFactory.Sum("actual_jan"))
            },
            periods);

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "KEEPFILTERS('Fact Sales'[Month Start] IN { DATE(2026, 1, 1) }), " +
            "KEEPFILTERS('Fact Sales'[Scenario] IN { \"Actual\" }))",
            measure.DaxFormula);
    }

    [Fact]
    public void Compiles_difference_explicit_share_and_escaped_measure_references()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "base", "Base ] Amount", PivotCalculationTestFactory.Sum()),
            PivotCalculationTestFactory.Measure(
                "difference",
                "Difference",
                new PivotDifferenceExpression(
                    new PivotMeasureReferenceExpression("base"),
                    new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum))),
            PivotCalculationTestFactory.Measure(
                "explicit_share",
                "Explicit Share",
                new PivotShareExpression(
                    new PivotMeasureReferenceExpression("base"),
                    new PivotExplicitShareDenominator(
                        new PivotAggregateExpression("amount", PivotCalculationAggregateFunction.Sum)),
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage())
        });

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal("([Base ]] Amount] - SUM('Fact Sales'[Units]))", compilation.Measures[1].DaxFormula);
        Assert.Equal(
            "DIVIDE([Base ]] Amount], SUM('Fact Sales'[Amount]), BLANK())",
            compilation.Measures[2].DaxFormula);
    }

    [Fact]
    public void Compiles_growth_achievement_variance_and_variance_percentage_semantics()
    {
        var actual = new PivotMeasureReferenceExpression("actual");
        var plan = new PivotMeasureReferenceExpression("plan");
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure("actual", "Actual", PivotCalculationTestFactory.Sum()),
            PivotCalculationTestFactory.Measure(
                "plan", "Plan", new PivotAggregateExpression("units", PivotCalculationAggregateFunction.Sum)),
            PivotCalculationTestFactory.Measure(
                "growth",
                "Growth",
                new PivotGrowthExpression(actual, plan, PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage()),
            PivotCalculationTestFactory.Measure(
                "achievement",
                "Achievement",
                new PivotAchievementExpression(actual, plan, PivotDenominatorBehavior.Zero),
                PivotCalculationTestFactory.Percentage()),
            PivotCalculationTestFactory.Measure(
                "variance",
                "Variance",
                new PivotVarianceExpression(actual, plan, PivotVarianceConvention.PlanMinusActual)),
            PivotCalculationTestFactory.Measure(
                "variance_pct",
                "Variance %",
                new PivotVariancePercentageExpression(
                    actual,
                    plan,
                    PivotVarianceConvention.ActualMinusPlan,
                    PivotDenominatorBehavior.Blank),
                PivotCalculationTestFactory.Percentage())
        });

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal("DIVIDE(([Actual] - [Plan]), [Plan], BLANK())", compilation.Measures[2].DaxFormula);
        Assert.Equal("DIVIDE([Actual], [Plan], 0)", compilation.Measures[3].DaxFormula);
        Assert.Equal("([Plan] - [Actual])", compilation.Measures[4].DaxFormula);
        Assert.Equal("DIVIDE(([Actual] - [Plan]), [Plan], BLANK())", compilation.Measures[5].DaxFormula);
    }

    [Fact]
    public void Continuous_date_slice_replaces_context_with_an_exact_bounded_range()
    {
        PivotModelSchema schema = PivotCalculationTestFactory.Schema();
        var periods = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "transaction_date",
                PivotPeriodGrain.Date,
                PivotPeriodCoverageStatus.Complete,
                Array.Empty<PivotPeriodCoverageMember>(),
                scenarioFieldId: "scenario",
                dateCoverageMode: PivotDateCoverageMode.ContinuousRange,
                continuousRangeStart: new DateTime(2026, 1, 1),
                continuousRangeEnd: new DateTime(2026, 12, 31),
                continuousRangeScenarioMemberIds: new[] { "actual", "plan" }),
            new[]
            {
                new PivotPeriodSlice(
                    "actual_q1",
                    "Q1 Actual",
                    new PivotPeriodPoint(PivotPeriodGrain.Quarter, 2026, 1),
                    "actual",
                    PivotSliceFilterMode.ReplaceAxisContext)
            });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure(
                    "actual_q1", "Q1 Actual", PivotCalculationTestFactory.Sum("actual_q1"))
            },
            periods,
            schema);

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);

        Assert.Equal(
            "CALCULATE(SUM('Fact Sales'[Amount]), " +
            "REMOVEFILTERS('Fact Sales'[Scenario]), " +
            "REMOVEFILTERS('Fact Sales'[Transaction Date]), " +
            "'Fact Sales'[Transaction Date] >= DATE(2026, 1, 1) && " +
            "'Fact Sales'[Transaction Date] <= DATE(2026, 3, 31), " +
            "'Fact Sales'[Scenario] IN { \"Actual\" })",
            measure.DaxFormula);
    }

    [Fact]
    public void Monthly_coverage_supports_explicit_year_half_quarter_and_month_slices()
    {
        PivotPeriodCoverageMember[] coverage = Enumerable.Range(1, 12)
            .Select(month => new PivotPeriodCoverageMember(
                new PivotPeriodPoint(PivotPeriodGrain.Month, 2026, month),
                PivotFilterValue.FromScalar(PivotScalarValue.Date(new DateTime(2026, month, 1))),
                new[] { "actual", "plan" }))
            .ToArray();
        var periods = new PivotPeriodDefinition(
            new PivotPeriodSource(
                "period",
                PivotPeriodGrain.Month,
                PivotPeriodCoverageStatus.Complete,
                coverage,
                scenarioFieldId: "scenario"),
            new[]
            {
                new PivotPeriodSlice(
                    "year", "Year", new PivotPeriodPoint(PivotPeriodGrain.Year, 2026),
                    "actual", PivotSliceFilterMode.ReplaceAxisContext),
                new PivotPeriodSlice(
                    "h2", "H2", new PivotPeriodPoint(PivotPeriodGrain.Half, 2026, 2),
                    "actual", PivotSliceFilterMode.ReplaceAxisContext),
                new PivotPeriodSlice(
                    "q2", "Q2", new PivotPeriodPoint(PivotPeriodGrain.Quarter, 2026, 2),
                    "actual", PivotSliceFilterMode.ReplaceAxisContext),
                new PivotPeriodSlice(
                    "apr", "Apr", new PivotPeriodPoint(PivotPeriodGrain.Month, 2026, 4),
                    "actual", PivotSliceFilterMode.ReplaceAxisContext)
            });
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(
            new[]
            {
                PivotCalculationTestFactory.Measure("year", "Year", PivotCalculationTestFactory.Sum("year")),
                PivotCalculationTestFactory.Measure("h2", "H2", PivotCalculationTestFactory.Sum("h2")),
                PivotCalculationTestFactory.Measure("q2", "Q2", PivotCalculationTestFactory.Sum("q2")),
                PivotCalculationTestFactory.Measure("apr", "Apr", PivotCalculationTestFactory.Sum("apr"))
            },
            periods);

        PivotDaxCompilation compilation = PivotDaxCompiler.Compile(definition);

        Assert.Equal(4, compilation.Measures.Count);
        Assert.Contains("DATE(2026, 12, 1)", compilation.Measures[0].DaxFormula);
        Assert.DoesNotContain("DATE(2026, 6, 1)", compilation.Measures[1].DaxFormula);
        Assert.Contains(
            "IN { DATE(2026, 4, 1), DATE(2026, 5, 1), DATE(2026, 6, 1) }",
            compilation.Measures[2].DaxFormula);
        Assert.Contains(
            "[Month Start] IN { DATE(2026, 4, 1) }",
            compilation.Measures[3].DaxFormula);
    }

    [Fact]
    public void Semantic_filter_sets_and_fingerprints_are_order_invariant()
    {
        PivotMeasureSetDefinition first = FilterSet("north", "south");
        PivotMeasureSetDefinition second = FilterSet("south", "north");

        OwnedPivotMeasureDefinition firstMeasure = Assert.Single(PivotDaxCompiler.Compile(first).Measures);
        OwnedPivotMeasureDefinition secondMeasure = Assert.Single(PivotDaxCompiler.Compile(second).Measures);

        Assert.Equal(firstMeasure.DaxFormula, secondMeasure.DaxFormula);
        Assert.Equal(firstMeasure.DefinitionFingerprint, secondMeasure.DefinitionFingerprint);
        Assert.Equal(firstMeasure.FormulaFingerprint, secondMeasure.FormulaFingerprint);
    }

    [Fact]
    public void Emits_separate_versioned_definition_and_formula_fingerprints()
    {
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure("revenue", "Revenue", PivotCalculationTestFactory.Sum())
        });

        OwnedPivotMeasureDefinition measure = Assert.Single(PivotDaxCompiler.Compile(definition).Measures);
        Assert.StartsWith("measure.definition.v1:sha256:", measure.DefinitionFingerprint);
        Assert.StartsWith("measure.formula.v1:sha256:", measure.FormulaFingerprint);
        Assert.NotEqual(measure.DefinitionFingerprint, measure.FormulaFingerprint);
        Assert.Equal(93, measure.DefinitionFingerprint.Length);
        Assert.Equal(90, measure.FormulaFingerprint.Length);
        Assert.Equal(
            "measure.formula.v1:sha256:" +
            "b67cb4cfeec64eec6f336a0e4389465e76350450dab7030fdc41d890572edcdd",
            measure.FormulaFingerprint);
    }

    [Fact]
    public void Definition_fingerprint_tracks_typed_format_while_formula_fingerprint_does_not()
    {
        OwnedPivotMeasureDefinition twoDecimals = Assert.Single(PivotDaxCompiler.Compile(
            PivotCalculationTestFactory.Set(new[]
            {
                PivotCalculationTestFactory.Measure(
                    "metric", "Metric", PivotCalculationTestFactory.Sum(),
                    PivotCalculationTestFactory.Decimal(2))
            })).Measures);
        OwnedPivotMeasureDefinition fourDecimals = Assert.Single(PivotDaxCompiler.Compile(
            PivotCalculationTestFactory.Set(new[]
            {
                PivotCalculationTestFactory.Measure(
                    "metric", "Metric", PivotCalculationTestFactory.Sum(),
                    PivotCalculationTestFactory.Decimal(4))
            })).Measures);

        Assert.Equal(twoDecimals.DaxFormula, fourDecimals.DaxFormula);
        Assert.Equal(twoDecimals.FormulaFingerprint, fourDecimals.FormulaFingerprint);
        Assert.NotEqual(twoDecimals.DefinitionFingerprint, fourDecimals.DefinitionFingerprint);
    }

    [Fact]
    public void Rejects_a_compiled_formula_that_exceeds_the_host_safe_length_bound()
    {
        PivotFilterValue[] values = Enumerable.Range(0, 256)
            .Select(index => PivotFilterValue.FromScalar(PivotScalarValue.Text(
                index.ToString("D3") + new string('x', 180))))
            .ToArray();
        PivotMeasureSetDefinition definition = PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "oversized",
                "Oversized",
                new PivotFilteredAggregateExpression(
                    "department",
                    PivotCalculationAggregateFunction.Count,
                    new[]
                    {
                        new PivotCalculationFilter(
                            "region",
                            PivotCalculationFilterOperator.In,
                            values)
                    }))
        });

        InvalidPivotCalculationException exception = Assert.Throws<InvalidPivotCalculationException>(
            () => PivotDaxCompiler.Compile(definition));

        Assert.Contains(
            exception.Validation.Issues,
            issue => issue.Code == "PIVOT_CALC_FORMULA_TOO_LONG");
    }

    private static PivotMeasureSetDefinition FilterSet(string first, string second)
    {
        return PivotCalculationTestFactory.Set(new[]
        {
            PivotCalculationTestFactory.Measure(
                "filtered",
                "Filtered",
                new PivotFilteredAggregateExpression(
                    "amount",
                    PivotCalculationAggregateFunction.Sum,
                    new[]
                    {
                        new PivotCalculationFilter(
                            "region",
                            PivotCalculationFilterOperator.In,
                            new[]
                            {
                                PivotFilterValue.FromMember(first),
                                PivotFilterValue.FromMember(second)
                            })
                    }))
        });
    }
}
