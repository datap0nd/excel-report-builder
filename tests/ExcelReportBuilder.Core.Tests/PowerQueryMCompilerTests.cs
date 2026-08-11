using ExcelReportBuilder.Core.PowerQuery;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

public sealed class PowerQueryMCompilerTests
{
    [Fact]
    public void Compiles_the_full_bounded_transform_surface_from_current_workbook()
    {
        var spec = CreateTransformSpec();

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Equal("Excel.CurrentWorkbook", result.SourceConnector);
        Assert.Contains("Source = Excel.CurrentWorkbook(){[Name=\"SourceData\"]}[Content]", result.Query);
        Assert.Contains("Table.SelectColumns", result.Query);
        Assert.Contains("Table.ReorderColumns", result.Query);
        Assert.Contains("Table.RenameColumns", result.Query);
        Assert.Contains("Table.TransformColumnTypes", result.Query);
        Assert.Contains("Table.ReplaceErrorValues", result.Query);
        Assert.Contains("Table.FillDown", result.Query);
        Assert.Contains("Table.SelectRows", result.Query);
        Assert.Contains("Date.QuarterOfYear", result.Query);
        Assert.Contains("Table.AddColumn", result.Query);
        Assert.DoesNotContain("File.Contents(", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Web.Contents(", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Sql.Database(", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Value_maps_are_one_pass_and_do_not_cascade_through_later_entries()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            Transforms =
            {
                new MapValuesTransform
                {
                    Id = "map",
                    Column = "Category",
                    Entries =
                    {
                        new ValueMapEntry { From = ScalarValue.FromText("A"), To = ScalarValue.FromText("B") },
                        new ValueMapEntry { From = ScalarValue.FromText("B"), To = ScalarValue.FromText("C") }
                    }
                }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("each if Value.Equals(_, \"A\") then \"B\" else if Value.Equals(_, \"B\") then \"C\" else _", result.Query);
        Assert.DoesNotContain("Table.ReplaceValue(Table.ReplaceValue", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Rejects_a_connector_injection_as_an_invalid_workbook_object_name()
    {
        var source = new WorkbookSourceSpec
        {
            WorkbookObjectName = "Data\"]}[Content], Evil = Web.Contents(\"https://invalid.example\")"
        };

        var exception = Assert.Throws<MCompilationException>(() =>
            PowerQueryMCompiler.Compile(source, Array.Empty<TransformStep>()));

        Assert.Equal("SOURCE_NAME_INVALID", exception.Code);
    }

    [Fact]
    public void Metric_month_normalization_emits_one_long_row_per_source_metric_period_column()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MetricMonthHeaders,
                ReportingYear = 2026,
                KeyColumns = { "Region" },
                Columns =
                {
                    Map("Revenue Jan", 1, "Revenue"),
                    Map("Cost Jan", 1, "Cost"),
                    Map("Revenue Feb", 2, "Revenue"),
                    Map("Cost Feb", 2, "Cost")
                }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("Table.Unpivot", result.Query);
        Assert.DoesNotContain("Table.Pivot", result.Query);
        Assert.Contains("\"Metric\"", result.Query);
        Assert.Contains("\"Value\"", result.Query);
        Assert.Contains("#date(2026, 1, 1)", result.Query);
    }

    [Fact]
    public void Missing_year_fails_closed_during_compilation()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MonthHeaders,
                KeyColumns = { "Region" },
                Columns = { Map("Jan", 1) }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var exception = Assert.Throws<MCompilationException>(() => PowerQueryMCompiler.Compile(spec));

        Assert.Equal("REPORTING_YEAR_REQUIRED", exception.Code);
    }

    [Fact]
    public void Quarter_normalization_emits_canonical_quarter_start_dates()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MonthHeaders,
                Grain = PeriodGrain.Quarter,
                KeyColumns = { "Region" },
                Columns =
                {
                    new PeriodColumnMapping { SourceColumn = "Q1 2026", Month = 1, Year = 2026 },
                    new PeriodColumnMapping { SourceColumn = "2026-Q2", Month = 4, Year = 2026 }
                }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("#date(2026, 1, 1)", result.Query);
        Assert.Contains("#date(2026, 4, 1)", result.Query);
    }

    [Fact]
    public void Quarter_normalization_rejects_non_start_months_during_compilation()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MonthHeaders,
                Grain = PeriodGrain.Quarter,
                ReportingYear = 2026,
                Columns = { Map("Bad quarter", 2) }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var exception = Assert.Throws<MCompilationException>(() => PowerQueryMCompiler.Compile(spec));

        Assert.Equal("QUARTER_START_MONTH_INVALID", exception.Code);
    }

    [Fact]
    public void Long_text_period_normalization_emits_a_bounded_canonical_date_parser()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.LongDateColumn,
                DateColumn = "Period label",
                Grain = PeriodGrain.Month,
                ReportingYear = 2026
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("Table.TransformColumns", result.Query);
        Assert.Contains("Date.StartOfMonth", result.Query);
        Assert.Contains("expectedGrain = \"month\"", result.Query);
        Assert.Contains("if two <= 29 then 2000 + two else 1900 + two", result.Query);
        Assert.Contains("reportingYear = 2026", result.Query);
        Assert.Contains("Unsupported period", result.Query);
        Assert.Contains("Reporting year required", result.Query);
        Assert.Contains("monthNumber =", result.Query);
        Assert.DoesNotContain("Text.BeforeDelimiter", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("raw is number", result.Query, StringComparison.Ordinal);
        Assert.Contains("parts = if text = null then {} else Text.SplitAny", result.Query);
        Assert.DoesNotContain("List.Select(Text.SplitAny(text", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("DateTime.LocalNow", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Culture.Current", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Long_period_normalization_requires_an_explicit_grain()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.LongDateColumn,
                DateColumn = "Period"
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var exception = Assert.Throws<MCompilationException>(() => PowerQueryMCompiler.Compile(spec));

        Assert.Equal("LONG_PERIOD_GRAIN_REQUIRED", exception.Code);
    }

    private static ReportSpecV1 CreateTransformSpec()
    {
        return new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            Transforms =
            {
                new SelectColumnsTransform
                {
                    Id = "select",
                    Columns = { "Period", "Region", "Category", "Amount", "Units", "Weight" }
                },
                new KeepColumnsTransform
                {
                    Id = "keep",
                    Columns = { "Period", "Region", "Category", "Amount", "Units", "Weight" }
                },
                new ReorderColumnsTransform
                {
                    Id = "reorder",
                    Columns = { "Region", "Category", "Period", "Amount", "Units", "Weight" }
                },
                new RenameColumnTransform { Id = "rename", From = "Category", To = "Group" },
                new ChangeColumnTypeTransform
                {
                    Id = "type_amount",
                    Column = "Amount",
                    DataType = ColumnDataType.DecimalNumber
                },
                new TrimTextTransform { Id = "trim", Columns = { "Region", "Group" } },
                new NormalizeBlanksTransform
                {
                    Id = "blanks",
                    Columns = { "Group" },
                    Replacement = ScalarValue.FromText("Unassigned")
                },
                new NormalizeErrorsTransform
                {
                    Id = "errors",
                    Columns = { "Amount" },
                    Replacement = ScalarValue.FromNumber(0)
                },
                new FillDownTransform { Id = "fill", Columns = { "Region" } },
                new MapValuesTransform
                {
                    Id = "map",
                    Column = "Group",
                    Entries =
                    {
                        new ValueMapEntry
                        {
                            From = ScalarValue.FromText("A"),
                            To = ScalarValue.FromText("Core")
                        }
                    }
                },
                new ReplaceValueTransform
                {
                    Id = "replace",
                    Column = "Region",
                    Find = ScalarValue.FromText("N"),
                    ReplaceWith = ScalarValue.FromText("North")
                },
                new FilterRowsTransform
                {
                    Id = "filter",
                    Column = "Amount",
                    Operator = RowFilterOperator.GreaterThanOrEqual,
                    Value = ScalarValue.FromNumber(0)
                },
                new ExcludeTotalRowsTransform
                {
                    Id = "exclude_totals",
                    Evidence =
                    {
                        new TotalRowEvidenceSpec
                        {
                            Column = "Region",
                            MatchKind = TotalRowMatchKind.EqualsAny,
                            Values = { ScalarValue.FromText("Total") },
                            Source = EvidenceSource.Preview,
                            ObservedMatchCount = 1
                        }
                    }
                },
                new DerivePeriodPartsTransform
                {
                    Id = "derive",
                    DateColumn = "Period",
                    Columns =
                    {
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.Year, OutputColumn = "Year" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.Half, OutputColumn = "Half" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.Quarter, OutputColumn = "Quarter" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.MonthNumber, OutputColumn = "MonthNumber" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.MonthName, OutputColumn = "MonthName" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.YearMonth, OutputColumn = "YearMonth" }
                    }
                },
                new AddArithmeticColumnTransform
                {
                    Id = "derive_rate",
                    OutputColumn = "Rate",
                    Operator = ArithmeticOperator.Divide,
                    Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Amount" },
                    Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Units" },
                    ResultType = ColumnDataType.DecimalNumber,
                    ReturnNullOnZeroDenominator = true
                },
                new RemoveColumnsTransform { Id = "remove", Columns = { "Weight" } }
            }
        };
    }

    private static PeriodColumnMapping Map(string source, int month, string? metric = null)
    {
        return new PeriodColumnMapping { SourceColumn = source, Month = month, Metric = metric };
    }
}
