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
        Assert.DoesNotContain("Table.PromoteHeaders", result.Query, StringComparison.Ordinal);
        Assert.Contains("Table.SelectColumns", result.Query);
        Assert.Contains("Table.ReorderColumns", result.Query);
        Assert.Contains("Table.RenameColumns", result.Query);
        Assert.DoesNotContain("Table.TransformColumnTypes", result.Query, StringComparison.Ordinal);
        Assert.Contains("Type conversions must be finite decimal numbers", result.Query);
        Assert.Contains("Table.ReplaceErrorValues", result.Query);
        Assert.Contains("Table.FillDown", result.Query);
        Assert.Contains("Table.SelectRows", result.Query);
        Assert.Contains("Date.QuarterOfYear", result.Query);
        Assert.Contains("Table.AddColumn", result.Query);
        Assert.Contains("Text.From(_, \"en-US\")", result.Query);
        Assert.Contains("Date.From([#\"Period\"], \"en-US\")", result.Query);
        Assert.Contains("Decimal.From(raw, \"en-US\")", result.Query);
        Assert.Contains("Value.Divide(left, right, Precision.Decimal)", result.Query);
        Assert.DoesNotContain("File.Contents(", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Web.Contents(", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Sql.Database(", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Named_range_promotes_exactly_one_header_row_before_any_transform()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec
            {
                Kind = WorkbookSourceKind.NamedRange,
                WorkbookObjectName = "ManagedSource",
                HeaderRowCount = 1
            },
            Transforms =
            {
                new SelectColumnsTransform
                {
                    Id = "select",
                    Columns = { "Region", "Amount" }
                }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        const string rawSource =
            "RawSource = Excel.CurrentWorkbook(){[Name=\"ManagedSource\"]}[Content]";
        const string promotedSource =
            "Source = Table.PromoteHeaders(RawSource, [PromoteAllScalars = true, Culture = \"en-US\"])";
        Assert.Contains(rawSource, result.Query);
        Assert.Contains(promotedSource, result.Query);
        Assert.True(result.Query.IndexOf(rawSource, StringComparison.Ordinal) <
                    result.Query.IndexOf(promotedSource, StringComparison.Ordinal));
        Assert.True(result.Query.IndexOf(promotedSource, StringComparison.Ordinal) <
                    result.Query.IndexOf("Table.SelectColumns(Source", StringComparison.Ordinal));
        Assert.Equal(
            1,
            result.Query.Split(new[] { "Table.PromoteHeaders" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void Table_source_preserves_Excel_table_headers_without_promotion()
    {
        var source = new WorkbookSourceSpec
        {
            Kind = WorkbookSourceKind.Table,
            WorkbookObjectName = "SourceTable",
            HeaderRowCount = 1
        };

        var result = PowerQueryMCompiler.Compile(source, Array.Empty<TransformStep>());

        Assert.Contains(
            "Source = Excel.CurrentWorkbook(){[Name=\"SourceTable\"]}[Content]",
            result.Query);
        Assert.DoesNotContain("RawSource", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Table.PromoteHeaders", result.Query, StringComparison.Ordinal);
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

        Assert.DoesNotContain("Table.Unpivot", result.Query, StringComparison.Ordinal);
        Assert.Contains("Table.ExpandListColumn", result.Query);
        Assert.Contains("Table.ExpandRecordColumn", result.Query);
        Assert.Contains("Record.FromList({\"Revenue Jan\", [#\"Revenue Jan\"]}", result.Query);
        Assert.Contains("Record.FromList({\"Cost Feb\", [#\"Cost Feb\"]}", result.Query);
        Assert.Equal(
            4,
            result.Query.Split(new[] { "Record.FromList" }, StringSplitOptions.None).Length - 1);
        Assert.DoesNotContain("Table.Pivot", result.Query);
        Assert.Contains("\"Metric\"", result.Query);
        Assert.Contains("\"Value\"", result.Query);
        Assert.Contains("#date(2026, 1, 1)", result.Query);
    }

    [Fact]
    public void Wide_normalization_expands_explicit_cell_records_so_null_values_remain_rows()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MonthHeaders,
                ReportingYear = 2026,
                KeyColumns = { "Region" },
                Columns =
                {
                    Map("Jan", 1),
                    Map("Feb", 2)
                }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.DoesNotContain("Table.Unpivot", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Table.SelectRows", result.Query, StringComparison.Ordinal);
        Assert.Contains("Table.AddColumn(selected, \"__erb_period_cells\"", result.Query);
        Assert.Contains(
            "Record.FromList({\"Jan\", [#\"Jan\"]}, {\"__erb_cell_header\", \"__erb_cell_value\"})",
            result.Query);
        Assert.Contains(
            "Record.FromList({\"Feb\", [#\"Feb\"]}, {\"__erb_cell_header\", \"__erb_cell_value\"})",
            result.Query);
        Assert.Contains(
            "Table.ExpandListColumn(withoutMappedColumns, \"__erb_period_cells\")",
            result.Query);
        Assert.Contains(
            "Table.ExpandRecordColumn(expandedCells, \"__erb_period_cells\"",
            result.Query);
        Assert.Equal(
            2,
            result.Query.Split(new[] { "Record.FromList" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void Wide_normalization_uses_collision_safe_internal_names_and_removes_inputs_before_expansion()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.MonthHeaders,
                ReportingYear = 2026,
                KeyColumns =
                {
                    "__erb_period_header",
                    "__erb_period_cells",
                    "__erb_cell_header",
                    "__erb_cell_value"
                },
                Columns = { Map("Jan", 1) }
            },
            Transforms =
            {
                new NormalizePeriodsTransform { Id = "normalize", PeriodMappingId = "periods" }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("\"__erb_period_header_1\"", result.Query);
        Assert.Contains("\"__erb_period_cells_1\"", result.Query);
        Assert.Contains("\"__erb_cell_header_1\"", result.Query);
        Assert.Contains("\"__erb_cell_value_1\"", result.Query);
        int removal = result.Query.IndexOf(
            "withoutMappedColumns = Table.RemoveColumns",
            StringComparison.Ordinal);
        int expansion = result.Query.IndexOf(
            "expandedCells = Table.ExpandListColumn",
            StringComparison.Ordinal);
        Assert.True(removal >= 0 && expansion > removal);
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
        Assert.Contains("numericCompactCandidate = raw is number", result.Query);
        Assert.Contains("raw is number and not numericCompactCandidate", result.Query);
        Assert.Contains("parts = if text = null then {} else Text.SplitAny", result.Query);
        Assert.DoesNotContain("List.Select(Text.SplitAny(text", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("DateTime.LocalNow", result.Query, StringComparison.Ordinal);
        Assert.DoesNotContain("Culture.Current", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Arithmetic_uses_bounded_decimal_operands_and_explicit_whole_number_coercion()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            Transforms =
            {
                new AddArithmeticColumnTransform
                {
                    Id = "calculate",
                    OutputColumn = "Result",
                    Operator = ArithmeticOperator.Multiply,
                    Left = new ArithmeticOperand
                    {
                        Kind = ArithmeticOperandKind.Column,
                        Column = "Amount"
                    },
                    Right = new ArithmeticOperand
                    {
                        Kind = ArithmeticOperandKind.Number,
                        Number = 2.5m
                    },
                    ResultType = ColumnDataType.WholeNumber
                }
            }
        };

        var result = PowerQueryMCompiler.Compile(spec);

        Assert.Contains("Text.Select(text, {\"0\"..\"9\", \"+\", \"-\", \".\", \",\", \"e\", \"E\"})", result.Query);
        Assert.Contains("Decimal.From(raw, \"en-US\")", result.Query);
        Assert.Contains("Decimal.From(\"2.5\", \"en-US\")", result.Query);
        Assert.Contains("Value.Multiply(left, right, Precision.Decimal)", result.Query);
        Assert.Contains("Int64.From(calculated, \"en-US\")", result.Query);
        Assert.DoesNotContain("Number.From(", result.Query, StringComparison.Ordinal);
    }

    [Fact]
    public void Column_type_changes_compile_the_closed_en_us_conversion_grammar()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            Transforms =
            {
                new ChangeColumnTypeTransform { Id = "text", Column = "Label", DataType = ColumnDataType.Text },
                new ChangeColumnTypeTransform { Id = "whole", Column = "Units", DataType = ColumnDataType.WholeNumber },
                new ChangeColumnTypeTransform { Id = "decimal", Column = "Amount", DataType = ColumnDataType.DecimalNumber },
                new ChangeColumnTypeTransform { Id = "boolean", Column = "Flag", DataType = ColumnDataType.Boolean },
                new ChangeColumnTypeTransform { Id = "date", Column = "Day", DataType = ColumnDataType.Date },
                new ChangeColumnTypeTransform { Id = "datetime", Column = "Stamp", DataType = ColumnDataType.DateTime }
            }
        };

        MCompilationResult result = PowerQueryMCompiler.Compile(spec);

        Assert.DoesNotContain("Table.TransformColumnTypes", result.Query, StringComparison.Ordinal);
        Assert.Contains("Text.From(_, \"en-US\")", result.Query);
        Assert.Contains("Decimal.From(raw, \"en-US\")", result.Query);
        Assert.Contains("Int64.From(converted, \"en-US\")", result.Query);
        Assert.Contains("Text.Lower(Text.Trim(_))", result.Query);
        Assert.Contains("Date.From(raw, \"en-US\")", result.Query);
        Assert.Contains("DateTime.From(raw, \"en-US\")", result.Query);
        Assert.Contains("Date.FromText(text, [Format = \"M/d/yyyy\", Culture = \"en-US\"])", result.Query);
        Assert.Contains("DateTime.FromText(text, [Format = \"M/d/yyyy h:mm:ss tt\", Culture = \"en-US\"])", result.Query);
        Assert.Contains("Text.Select(text, {\"0\"..\"9\", \"+\", \"-\", \".\", \",\", \"e\", \"E\"})", result.Query);
    }

    [Fact]
    public void Division_without_null_on_zero_is_rejected_during_compilation()
    {
        var spec = new ReportSpecV1
        {
            Source = new WorkbookSourceSpec { WorkbookObjectName = "SourceData" },
            Transforms =
            {
                new AddArithmeticColumnTransform
                {
                    Id = "unsafe_divide",
                    OutputColumn = "Result",
                    Operator = ArithmeticOperator.Divide,
                    Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Amount" },
                    Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Units" },
                    ResultType = ColumnDataType.DecimalNumber,
                    ReturnNullOnZeroDenominator = false
                }
            }
        };

        MCompilationException exception = Assert.Throws<MCompilationException>(() =>
            PowerQueryMCompiler.Compile(spec));

        Assert.Equal("ARITHMETIC_DIVIDE_NULL_ON_ZERO_REQUIRED", exception.Code);
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
