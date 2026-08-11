using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Rendering;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class MeasureFormulaCompilerTests
{
    [Fact]
    public void Typed_ratio_compiles_to_pivot_formulas_with_blank_zero_denominator()
    {
        var measures = CreateMeasures();
        var pivot = new PivotBuildResult
        {
            PivotTableName = "ERB_Pivot",
            WorksheetName = "Managed pivot",
            AnchorCell = "$A$3",
            DataFields = new[]
            {
                new PivotDataFieldDescriptor
                {
                    MeasureId = "rate",
                    ComponentId = "rate_component_1",
                    Role = AggregateComponentRole.Input,
                    SourceField = "Amount",
                    Function = AggregateFunction.Sum,
                    PivotCaption = "ERB rate 1"
                },
                new PivotDataFieldDescriptor
                {
                    MeasureId = "rate",
                    ComponentId = "rate_component_2",
                    Role = AggregateComponentRole.Input,
                    SourceField = "Units",
                    Function = AggregateFunction.Sum,
                    PivotCaption = "ERB rate 2"
                }
            }
        };

        var formula = new MeasureFormulaCompiler().Compile("rate", measures, pivot);

        Assert.StartsWith("=IFERROR(", formula.Value);
        Assert.Contains("GETPIVOTDATA(\"ERB rate 1\"", formula.Value);
        Assert.Contains("GETPIVOTDATA(\"ERB rate 2\"", formula.Value);
        Assert.Contains("=0", formula.Value);
        Assert.EndsWith(",\"\")", formula.Value);
    }

    [Fact]
    public void Typed_compiler_has_no_public_raw_formula_constructor()
    {
        Assert.Empty(typeof(SafeExcelFormula).GetConstructors());
        Assert.DoesNotContain(
            typeof(SafeFormulaFactory).GetMethods(),
            method => method.IsPublic && method.GetParameters().Any(parameter => parameter.ParameterType == typeof(string)));
    }

    [Fact]
    public void Ratio_over_a_multi_period_slice_sums_each_leaf_before_dividing()
    {
        var measures = CreateMeasures();
        var pivot = new PivotBuildResult
        {
            PivotTableName = "ERB_Pivot",
            WorksheetName = "Managed pivot",
            AnchorCell = "$A$3",
            DataFields = new[]
            {
                new PivotDataFieldDescriptor
                {
                    MeasureId = "rate",
                    ComponentId = "rate_component_1",
                    Role = AggregateComponentRole.Input,
                    SourceField = "Amount",
                    Function = AggregateFunction.Sum,
                    PivotCaption = "ERB rate 1"
                },
                new PivotDataFieldDescriptor
                {
                    MeasureId = "rate",
                    ComponentId = "rate_component_2",
                    Role = AggregateComponentRole.Input,
                    SourceField = "Units",
                    Function = AggregateFunction.Sum,
                    PivotCaption = "ERB rate 2"
                }
            }
        };
        IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberSets = new[]
        {
            (IReadOnlyList<PivotFilterItem>)new[]
            {
                new PivotFilterItem { Field = "Period", Value = new DateTime(2026, 1, 1) }
            },
            (IReadOnlyList<PivotFilterItem>)new[]
            {
                new PivotFilterItem { Field = "Period", Value = new DateTime(2026, 2, 1) }
            }
        };

        var formula = new MeasureFormulaCompiler().CompileAcrossMemberSets(
            "rate",
            measures,
            pivot,
            memberSets,
            new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>());

        Assert.Equal(2, Count(formula.Value, "GETPIVOTDATA(\"ERB rate 1\""));
        Assert.Equal(6, Count(formula.Value, "GETPIVOTDATA(\"ERB rate 2\""));
        Assert.Contains("DATE(2026,1,1)", formula.Value);
        Assert.Contains("DATE(2026,2,1)", formula.Value);
    }

    [Fact]
    public void Percentage_point_difference_remains_a_percentage_typed_decimal()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["change"] = new MeasureDefinition
            {
                Id = "change",
                Label = "Change",
                ValueType = MeasureValueType.Percentage,
                Expression = new DifferenceMeasureExpression
                {
                    DifferenceKind = DifferenceKind.PercentagePoints,
                    Current = new AggregateMeasureExpression { Field = "Current", Function = AggregateFunction.Sum },
                    Baseline = new AggregateMeasureExpression { Field = "Baseline", Function = AggregateFunction.Sum }
                }
            }
        };
        var pivot = PivotFor(
            "change",
            ("Current", "Current total"),
            ("Baseline", "Baseline total"));

        var formula = new MeasureFormulaCompiler().Compile("change", measures, pivot);

        Assert.DoesNotContain("100*", formula.Value, StringComparison.Ordinal);
        Assert.Contains("Current total", formula.Value, StringComparison.Ordinal);
        Assert.Contains("Baseline total", formula.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Explicit_error_on_zero_is_not_suppressed_by_an_outer_iferror()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["rate"] = new MeasureDefinition
            {
                Id = "rate",
                Label = "Rate",
                Expression = new RatioMeasureExpression
                {
                    Numerator = new AggregateMeasureExpression { Field = "Amount", Function = AggregateFunction.Sum },
                    Denominator = new AggregateMeasureExpression { Field = "Units", Function = AggregateFunction.Sum },
                    OnZero = ZeroDenominatorBehavior.Error
                }
            }
        };
        var pivot = PivotFor("rate", ("Amount", "Amount total"), ("Units", "Units total"));

        var formula = new MeasureFormulaCompiler().Compile("rate", measures, pivot);

        Assert.DoesNotContain("IFERROR", formula.Value, StringComparison.Ordinal);
        Assert.StartsWith("=", formula.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Matches_equivalent_filtered_components_regardless_of_filter_order()
    {
        var expression = new FilteredAggregateMeasureExpression
        {
            Field = "Amount",
            Function = AggregateFunction.Sum,
            Filters =
            {
                new MeasureFilterSpec
                {
                    Field = "Region",
                    Operator = MeasureFilterOperator.In,
                    Values = { ScalarValue.FromText("B"), ScalarValue.FromText("A") }
                },
                new MeasureFilterSpec
                {
                    Field = "Status",
                    Operator = MeasureFilterOperator.Equal,
                    Values = { ScalarValue.FromText("Active") }
                }
            }
        };
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["filtered"] = new MeasureDefinition { Id = "filtered", Label = "Filtered", Expression = expression }
        };
        var pivot = PivotFor("filtered", ("Amount", "Filtered total"));
        pivot.DataFields[0].Filters = new[]
        {
            new MeasureFilterSpec
            {
                Field = "Status",
                Operator = MeasureFilterOperator.Equal,
                Values = { ScalarValue.FromText("Active") }
            },
            new MeasureFilterSpec
            {
                Field = "Region",
                Operator = MeasureFilterOperator.In,
                Values = { ScalarValue.FromText("A"), ScalarValue.FromText("B") }
            }
        };

        var formula = new MeasureFormulaCompiler().Compile("filtered", measures, pivot);

        Assert.Contains("Filtered total", formula.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Rejects_excessive_getpivotdata_term_expansion()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["amount"] = new MeasureDefinition
            {
                Id = "amount",
                Label = "Amount",
                Expression = new AggregateMeasureExpression { Field = "Amount", Function = AggregateFunction.Sum }
            }
        };
        var pivot = PivotFor("amount", ("Amount", "Amount total"));
        var members = Enumerable.Range(1, MeasureFormulaCompiler.MaximumExpandedAggregateTerms + 1)
            .Select(index => (IReadOnlyList<PivotFilterItem>)new[]
            {
                new PivotFilterItem { Field = "Member", Value = index }
            })
            .ToList();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new MeasureFormulaCompiler().CompileAcrossMemberSets(
                "amount",
                measures,
                pivot,
                members,
                new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>()));

        Assert.Contains("too many", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(ShareDenominatorScope.Parent, 4, 1, 4)]
    [InlineData(ShareDenominatorScope.FilteredReportTotal, 1, 1, 4)]
    public void Share_scope_removes_only_the_intended_row_hierarchy_filters(
        ShareDenominatorScope scope,
        int expectedRegionUses,
        int expectedCategoryUses,
        int expectedPeriodUses)
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["share"] = new MeasureDefinition
            {
                Id = "share",
                Label = "Share",
                ValueType = MeasureValueType.Percentage,
                Expression = new ShareMeasureExpression
                {
                    ResultType = MeasureValueType.Percentage,
                    Part = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    },
                    Whole = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    },
                    Scope = scope,
                    OnZero = ZeroDenominatorBehavior.Blank
                }
            }
        };
        var pivot = PivotFor("share", ("Amount", "Amount total"));
        IReadOnlyList<IReadOnlyList<PivotFilterItem>> memberSets = new[]
        {
            (IReadOnlyList<PivotFilterItem>)new[]
            {
                new PivotFilterItem { Field = "Region", Value = "North" },
                new PivotFilterItem { Field = "Category", Value = "A" },
                new PivotFilterItem { Field = "Period", Value = new DateTime(2026, 1, 1) }
            }
        };

        SafeExcelFormula formula = new MeasureFormulaCompiler().CompileAcrossMemberSets(
            "share",
            measures,
            pivot,
            memberSets,
            new Dictionary<string, IReadOnlyList<IReadOnlyList<PivotFilterItem>>>(),
            new[] { "Region", "Category" });

        // Safe division repeats the denominator for both zero checks and the
        // final division. Counts therefore prove which filters survive into
        // the denominator, not how often Excel evaluates one expression.
        Assert.Equal(expectedRegionUses, Count(formula.Value, "\"Region\""));
        Assert.Equal(expectedCategoryUses, Count(formula.Value, "\"Category\""));
        Assert.Equal(expectedPeriodUses, Count(formula.Value, "\"Period\""));
    }

    [Fact]
    public void Scoped_share_requires_a_rows_hierarchy()
    {
        var measures = new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            ["share"] = new MeasureDefinition
            {
                Id = "share",
                Label = "Share",
                ValueType = MeasureValueType.Percentage,
                Expression = new ShareMeasureExpression
                {
                    Part = new AggregateMeasureExpression { Field = "Amount", Function = AggregateFunction.Sum },
                    Whole = new AggregateMeasureExpression { Field = "Amount", Function = AggregateFunction.Sum },
                    Scope = ShareDenominatorScope.Parent
                }
            }
        };

        var exception = Assert.Throws<InvalidOperationException>(() =>
            new MeasureFormulaCompiler().Compile("share", measures, PivotFor("share", ("Amount", "Amount total"))));

        Assert.Contains("Rows hierarchy", exception.Message, StringComparison.Ordinal);
    }

    private static PivotBuildResult PivotFor(
        string measureId,
        params (string Field, string Caption)[] components)
    {
        return new PivotBuildResult
        {
            PivotTableName = "ERB_Pivot",
            WorksheetName = "Managed pivot",
            AnchorCell = "$A$3",
            DataFields = components.Select((component, index) => new PivotDataFieldDescriptor
            {
                MeasureId = measureId,
                ComponentId = measureId + "_component_" + (index + 1),
                Role = AggregateComponentRole.Input,
                SourceField = component.Field,
                Function = AggregateFunction.Sum,
                PivotCaption = component.Caption
            }).ToArray()
        };
    }

    private static int Count(string value, string text)
    {
        var count = 0;
        var index = 0;
        while ((index = value.IndexOf(text, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += text.Length;
        }

        return count;
    }

    private static IReadOnlyDictionary<string, MeasureDefinition> CreateMeasures()
    {
        var amount = new MeasureDefinition
        {
            Id = "amount",
            Label = "Amount",
            Expression = new AggregateMeasureExpression
            {
                Field = "Amount",
                Function = AggregateFunction.Sum
            }
        };
        var units = new MeasureDefinition
        {
            Id = "units",
            Label = "Units",
            Expression = new AggregateMeasureExpression
            {
                Field = "Units",
                Function = AggregateFunction.Sum
            }
        };
        var rate = new MeasureDefinition
        {
            Id = "rate",
            Label = "Rate",
            Expression = new RatioMeasureExpression
            {
                Numerator = new ReferenceMeasureExpression { MeasureId = "amount" },
                Denominator = new ReferenceMeasureExpression { MeasureId = "units" },
                OnZero = ZeroDenominatorBehavior.Blank
            }
        };
        return new Dictionary<string, MeasureDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            [amount.Id] = amount,
            [units.Id] = units,
            [rate.Id] = rate
        };
    }
}
