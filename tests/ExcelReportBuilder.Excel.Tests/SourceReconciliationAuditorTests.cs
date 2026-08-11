using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;
using ExcelReportBuilder.Excel.Validation;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class SourceReconciliationAuditorTests
{
    private readonly SourceReconciliationAuditor _auditor = new();

    [Fact]
    public void Legitimate_filter_and_confirmed_total_exclusion_produce_exact_rows_and_totals()
    {
        ReportSpecV1 specification = BasicSpecification();
        specification.Transforms.Add(new ExcludeTotalRowsTransform
        {
            Id = "exclude_total",
            Evidence =
            {
                new TotalRowEvidenceSpec
                {
                    Column = "Label",
                    MatchKind = TotalRowMatchKind.EqualsAny,
                    Values = { ScalarValue.FromText("Total") },
                    Source = EvidenceSource.UserConfirmation,
                    ObservedMatchCount = 1
                }
            }
        });
        specification.Transforms.Add(new TrimTextTransform
        {
            Id = "trim_status",
            Columns = { "Status" }
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "keep_open",
            Column = "Status",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("Open")
        });
        var rows = new[]
        {
            Row(("Label", "A"), ("Status", " Open "), ("Amount", 10m)),
            Row(("Label", "B"), ("Status", "Closed"), ("Amount", 100m)),
            Row(("Label", "Total"), ("Status", "Open"), ("Amount", 999m)),
            Row(("Label", "C"), ("Status", "Open"), ("Amount", 20m))
        };

        SourceReconciliationAudit result = _auditor.AuditRows(rows, specification);

        Assert.Equal(4, result.SourceRows);
        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(1, result.RemovedRowsByTransform["exclude_total"]);
        Assert.Equal(1, result.RemovedRowsByTransform["keep_open"]);
        Assert.Equal(30m, result.ExpectedTotals["amount"]);
    }

    [Fact]
    public void Changed_total_row_evidence_fails_closed()
    {
        ReportSpecV1 specification = BasicSpecification();
        specification.Transforms.Add(new ExcludeTotalRowsTransform
        {
            Id = "exclude_total",
            Evidence =
            {
                new TotalRowEvidenceSpec
                {
                    Column = "Label",
                    MatchKind = TotalRowMatchKind.EqualsAny,
                    Values = { ScalarValue.FromText("Total") },
                    Source = EvidenceSource.UserConfirmation,
                    ObservedMatchCount = 2
                }
            }
        });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(
                new[] { Row(("Label", "Total"), ("Status", "Open"), ("Amount", 10m)) },
                specification));

        Assert.Contains("expected 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("contains 1", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Value_changing_transforms_produce_the_expected_additive_total()
    {
        var specification = new ReportSpecV1
        {
            Measures =
            {
                new MeasureDefinition
                {
                    Id = "derived",
                    Label = "Derived",
                    Expression = new AggregateMeasureExpression
                    {
                        Field = "Derived",
                        Function = AggregateFunction.Sum
                    }
                }
            }
        };
        specification.Transforms.Add(new RenameColumnTransform
        {
            Id = "rename",
            From = "Raw input",
            To = "Raw"
        });
        specification.Transforms.Add(new NormalizeErrorsTransform
        {
            Id = "errors",
            Columns = { "Raw" },
            Replacement = ScalarValue.FromNumber(2m)
        });
        specification.Transforms.Add(new NormalizeBlanksTransform
        {
            Id = "blanks",
            Columns = { "Raw" },
            Replacement = ScalarValue.Null(),
            TreatWhitespaceAsBlank = true
        });
        specification.Transforms.Add(new FillDownTransform
        {
            Id = "fill",
            Columns = { "Raw" }
        });
        specification.Transforms.Add(new TrimTextTransform
        {
            Id = "trim",
            Columns = { "Raw" }
        });
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "number",
            Column = "Raw",
            DataType = ColumnDataType.DecimalNumber
        });
        specification.Transforms.Add(new ReplaceValueTransform
        {
            Id = "replace",
            Column = "Raw",
            Find = ScalarValue.FromNumber(10m),
            ReplaceWith = ScalarValue.FromNumber(11m)
        });
        specification.Transforms.Add(new MapValuesTransform
        {
            Id = "map",
            Column = "Raw",
            Entries =
            {
                new ValueMapEntry
                {
                    From = ScalarValue.FromNumber(11m),
                    To = ScalarValue.FromNumber(12m)
                }
            }
        });
        specification.Transforms.Add(new AddArithmeticColumnTransform
        {
            Id = "derive",
            OutputColumn = "Derived",
            Operator = ArithmeticOperator.Multiply,
            Left = new ArithmeticOperand
            {
                Kind = ArithmeticOperandKind.Column,
                Column = "Raw"
            },
            Right = new ArithmeticOperand
            {
                Kind = ArithmeticOperandKind.Number,
                Number = 2m
            }
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "valid",
            Column = "Derived",
            Operator = RowFilterOperator.GreaterThanOrEqual,
            Value = ScalarValue.FromNumber(4m)
        });
        var rows = new[]
        {
            Row(("Raw input", " 10 ")),
            Row(("Raw input", new ErrorWrapper(-2146826281))),
            Row(("Raw input", "   "))
        };

        SourceReconciliationAudit result = _auditor.AuditRows(rows, specification);

        Assert.Equal(3, result.ExpectedNormalizedRows);
        Assert.Equal(32m, result.ExpectedTotals["derived"]);
    }

    [Fact]
    public void Filter_after_wide_normalization_counts_each_expanded_row()
    {
        ReportSpecV1 specification = BasicSpecification();
        specification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MonthHeaders,
            Grain = PeriodGrain.Month,
            KeyColumns = { "Label" },
            PeriodColumnName = "Period",
            ValueColumnName = "Amount",
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Jan", Month = 1, Year = 2026 },
                new PeriodColumnMapping { SourceColumn = "Feb", Month = 2, Year = 2026 }
            }
        };
        specification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize_periods",
            PeriodMappingId = "periods"
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "positive_values",
            Column = "Amount",
            Operator = RowFilterOperator.GreaterThan,
            Value = ScalarValue.FromNumber(0m)
        });
        var rows = new[]
        {
            Row(("Label", "A"), ("Jan", 10m), ("Feb", 0m)),
            Row(("Label", "B"), ("Jan", 5m), ("Feb", -1m))
        };

        SourceReconciliationAudit result = _auditor.AuditRows(rows, specification);

        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(2, result.RemovedRowsByTransform["positive_values"]);
        Assert.Equal(15m, result.ExpectedTotals["amount"]);
    }

    [Fact]
    public void Wide_normalization_counts_a_blank_mapped_cell_as_a_canonical_row()
    {
        ReportSpecV1 specification = BasicSpecification();
        specification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MonthHeaders,
            Grain = PeriodGrain.Month,
            KeyColumns = { "Label" },
            PeriodColumnName = "Period",
            ValueColumnName = "Amount",
            Columns =
            {
                new PeriodColumnMapping { SourceColumn = "Jan", Month = 1, Year = 2026 },
                new PeriodColumnMapping { SourceColumn = "Feb", Month = 2, Year = 2026 }
            }
        };
        specification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize_periods",
            PeriodMappingId = "periods"
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Label", "A"), ("Jan", null), ("Feb", 12m))
            },
            specification);

        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(12m, result.ExpectedTotals["amount"]);
        Assert.Empty(result.RemovedRowsByTransform);
    }

    [Fact]
    public void Metric_month_and_long_period_mappings_are_evaluated_without_materializing_output()
    {
        ReportSpecV1 metricSpecification = BasicSpecification();
        metricSpecification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.MetricMonthHeaders,
            Grain = PeriodGrain.Month,
            KeyColumns = { "Label" },
            PeriodColumnName = "Period",
            MetricColumnName = "Metric",
            ValueColumnName = "Amount",
            Columns =
            {
                new PeriodColumnMapping
                {
                    SourceColumn = "Jan Sales",
                    Month = 1,
                    Year = 2026,
                    Metric = "Sales"
                },
                new PeriodColumnMapping
                {
                    SourceColumn = "Jan Units",
                    Month = 1,
                    Year = 2026,
                    Metric = "Units"
                }
            }
        };
        metricSpecification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize_periods",
            PeriodMappingId = "periods"
        });
        metricSpecification.Transforms.Add(new FilterRowsTransform
        {
            Id = "sales_only",
            Column = "Metric",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("Sales")
        });

        SourceReconciliationAudit metric = _auditor.AuditRows(
            new[]
            {
                Row(("Label", "A"), ("Jan Sales", 25m), ("Jan Units", 3m))
            },
            metricSpecification);

        Assert.Equal(1, metric.ExpectedNormalizedRows);
        Assert.Equal(25m, metric.ExpectedTotals["amount"]);

        ReportSpecV1 longSpecification = BasicSpecification();
        longSpecification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.LongDateColumn,
            DateColumn = "Period",
            Grain = PeriodGrain.Month
        };
        longSpecification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize_periods",
            PeriodMappingId = "periods"
        });
        longSpecification.Transforms.Add(new FilterRowsTransform
        {
            Id = "selected_month",
            Column = "Period",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDate(new DateTime(2026, 1, 1))
        });

        SourceReconciliationAudit longResult = _auditor.AuditRows(
            new[]
            {
                Row(("Period", "Jan-26"), ("Amount", 5m)),
                Row(("Period", "Feb-26"), ("Amount", 7m))
            },
            longSpecification);

        Assert.Equal(1, longResult.ExpectedNormalizedRows);
        Assert.Equal(5m, longResult.ExpectedTotals["amount"]);
    }

    [Fact]
    public void Conversion_errors_remain_repairable_until_normalize_errors()
    {
        ReportSpecV1 specification = SumSpecification("Raw");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "convert",
            Column = "Raw",
            DataType = ColumnDataType.DecimalNumber
        });
        specification.Transforms.Add(new MapValuesTransform
        {
            Id = "map_valid_value",
            Column = "Raw",
            Entries =
            {
                new ValueMapEntry
                {
                    From = ScalarValue.FromNumber(2.5m),
                    To = ScalarValue.FromNumber(3m)
                }
            }
        });
        specification.Transforms.Add(new NormalizeErrorsTransform
        {
            Id = "repair",
            Columns = { "Raw" },
            Replacement = ScalarValue.FromNumber(7m)
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Raw", "2.5")),
                Row(("Raw", "bad"))
            },
            specification);

        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(10m, result.ExpectedTotals["total"]);
    }

    [Fact]
    public void Whole_number_conversion_uses_to_even_rounding()
    {
        ReportSpecV1 specification = SumSpecification("Raw");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "whole",
            Column = "Raw",
            DataType = ColumnDataType.WholeNumber
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Raw", "4.5")),
                Row(("Raw", "5.5"))
            },
            specification);

        Assert.Equal(10m, result.ExpectedTotals["total"]);
    }

    [Fact]
    public void Decimal_conversion_accepts_exponents_but_rejects_percent_text()
    {
        ReportSpecV1 specification = SumSpecification("Raw");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "decimal",
            Column = "Raw",
            DataType = ColumnDataType.DecimalNumber
        });
        specification.Transforms.Add(new NormalizeErrorsTransform
        {
            Id = "repair_percent",
            Columns = { "Raw" },
            Replacement = ScalarValue.FromNumber(7m)
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Raw", "1e3")),
                Row(("Raw", "12.5%"))
            },
            specification);

        Assert.Equal(1007m, result.ExpectedTotals["total"]);
    }

    [Fact]
    public void Boolean_conversion_accepts_only_bounded_logical_inputs()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "logical",
            Column = "Flag",
            DataType = ColumnDataType.Boolean
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "true_only",
            Column = "Flag",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromBoolean(true)
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Flag", " TRUE "), ("Amount", 1m)),
                Row(("Flag", 1m), ("Amount", 2m)),
                Row(("Flag", 0m), ("Amount", 4m))
            },
            specification);

        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(3m, result.ExpectedTotals["total"]);
    }

    [Fact]
    public void Date_conversion_uses_en_us_and_rejects_day_first_text()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "date",
            Column = "Day",
            DataType = ColumnDataType.Date
        });
        specification.Transforms.Add(new NormalizeErrorsTransform
        {
            Id = "repair_invalid_date",
            Columns = { "Day" },
            Replacement = ScalarValue.FromDate(new DateTime(2026, 1, 1))
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "year_end",
            Column = "Day",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDate(new DateTime(2026, 12, 31))
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Day", "12/31/2026"), ("Amount", 1m)),
                Row(("Day", "31/12/2026"), ("Amount", 2m))
            },
            specification);

        Assert.Equal(1, result.ExpectedNormalizedRows);
        Assert.Equal(1m, result.ExpectedTotals["total"]);
    }

    [Theory]
    [InlineData(RowFilterOperator.GreaterThan, 1, 11)]
    [InlineData(RowFilterOperator.GreaterThanOrEqual, 2, 21)]
    [InlineData(RowFilterOperator.LessThan, 1, 9)]
    [InlineData(RowFilterOperator.LessThanOrEqual, 2, 19)]
    public void Relational_filters_never_select_null_cells(
        RowFilterOperator filterOperator,
        int expectedRows,
        int expectedTotal)
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "compare",
            Column = "Amount",
            Operator = filterOperator,
            Value = ScalarValue.FromNumber(10m)
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Amount", null)),
                Row(("Amount", 9m)),
                Row(("Amount", 10m)),
                Row(("Amount", 11m))
            },
            specification);

        Assert.Equal(expectedRows, result.ExpectedNormalizedRows);
        Assert.Equal((decimal)expectedTotal, result.ExpectedTotals["total"]);
    }

    [Fact]
    public void Compact_numeric_periods_are_parsed_before_oadate_serials()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.LongDateColumn,
            DateColumn = "Period",
            Grain = PeriodGrain.Month
        };
        specification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize",
            PeriodMappingId = "periods"
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "january",
            Column = "Period",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDate(new DateTime(2026, 1, 1))
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Period", 202601m), ("Amount", 5m)),
                Row(("Period", new DateTime(2026, 1, 15).ToOADate()), ("Amount", 7m)),
                Row(("Period", 45292m), ("Amount", 3m))
            },
            specification);

        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(12m, result.ExpectedTotals["total"]);
        Assert.Equal(1, result.RemovedRowsByTransform["january"]);
    }

    [Theory]
    [InlineData(202600d)]
    [InlineData(202613d)]
    public void Invalid_six_digit_numeric_periods_do_not_fall_through_to_oadates(double value)
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.PeriodMapping = new PeriodMappingSpec
        {
            Id = "periods",
            Kind = PeriodMappingKind.LongDateColumn,
            DateColumn = "Period",
            Grain = PeriodGrain.Month
        };
        specification.Transforms.Add(new NormalizePeriodsTransform
        {
            Id = "normalize",
            PeriodMappingId = "periods"
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "consume_period",
            Column = "Period",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDate(new DateTime(2026, 1, 1))
        });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(
                new[] { Row(("Period", value), ("Amount", 1m)) },
                specification));

        Assert.Contains("normalize", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("YYYYMM", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_conversion_is_explicitly_en_us_and_preserves_date_kinds()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "day_type",
            Column = "Day",
            DataType = ColumnDataType.Date
        });
        specification.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "stamp_type",
            Column = "Stamp",
            DataType = ColumnDataType.DateTime
        });
        specification.Transforms.Add(new TrimTextTransform
        {
            Id = "to_text",
            Columns = { "Flag", "Number", "Day", "Stamp" }
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "flag_text",
            Column = "Flag",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("true")
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "number_text",
            Column = "Number",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("1234.5")
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "date_text",
            Column = "Day",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("1/2/2026")
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "datetime_text",
            Column = "Stamp",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("1/2/2026 3:04:05 PM")
        });

        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        CultureInfo originalUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("de-DE");
            SourceReconciliationAudit result = _auditor.AuditRows(
                new[]
                {
                    Row(
                        ("Flag", true),
                        ("Number", 1234.5m),
                        ("Day", new DateTime(2026, 1, 2, 15, 4, 5)),
                        ("Stamp", new DateTime(2026, 1, 2, 15, 4, 5)),
                        ("Amount", 1m))
                },
                specification);

            Assert.Equal(1, result.ExpectedNormalizedRows);
            Assert.Equal(1m, result.ExpectedTotals["total"]);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public void Date_and_datetime_equality_remains_type_sensitive()
    {
        ReportSpecV1 mismatch = SumSpecification("Amount");
        mismatch.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "as_date",
            Column = "When",
            DataType = ColumnDataType.Date
        });
        mismatch.Transforms.Add(new FilterRowsTransform
        {
            Id = "datetime_literal",
            Column = "When",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDateTime(new DateTime(2026, 1, 2))
        });

        SourceReconciliationAudit mismatchResult = _auditor.AuditRows(
            new[] { Row(("When", new DateTime(2026, 1, 2)), ("Amount", 1m)) },
            mismatch);

        Assert.Equal(0, mismatchResult.ExpectedNormalizedRows);

        ReportSpecV1 matching = SumSpecification("Amount");
        matching.Transforms.Add(new ChangeColumnTypeTransform
        {
            Id = "as_date",
            Column = "When",
            DataType = ColumnDataType.Date
        });
        matching.Transforms.Add(new FilterRowsTransform
        {
            Id = "date_literal",
            Column = "When",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromDate(new DateTime(2026, 1, 2))
        });

        SourceReconciliationAudit matchingResult = _auditor.AuditRows(
            new[] { Row(("When", new DateTime(2026, 1, 2)), ("Amount", 1m)) },
            matching);

        Assert.Equal(1, matchingResult.ExpectedNormalizedRows);
    }

    [Fact]
    public void Division_returns_blank_for_zero_and_null_denominators()
    {
        ReportSpecV1 specification = SumSpecification("Result");
        specification.Transforms.Add(new AddArithmeticColumnTransform
        {
            Id = "divide",
            OutputColumn = "Result",
            Operator = ArithmeticOperator.Divide,
            Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Numerator" },
            Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Denominator" },
            ResultType = ColumnDataType.DecimalNumber,
            ReturnNullOnZeroDenominator = true
        });

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Numerator", 10m), ("Denominator", 2m)),
                Row(("Numerator", 10m), ("Denominator", 0m)),
                Row(("Numerator", 10m), ("Denominator", null))
            },
            specification);

        Assert.Equal(3, result.ExpectedNormalizedRows);
        Assert.Equal(5m, result.ExpectedTotals["total"]);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Arithmetic_rejects_non_finite_values(double value)
    {
        ReportSpecV1 specification = SumSpecification("Result");
        specification.Transforms.Add(new AddArithmeticColumnTransform
        {
            Id = "calculate",
            OutputColumn = "Result",
            Operator = ArithmeticOperator.Add,
            Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Input" },
            Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Number, Number = 1m },
            ResultType = ColumnDataType.DecimalNumber
        });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(new[] { Row(("Input", value)) }, specification));

        Assert.Contains("calculate", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("numeric", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Audit_rejects_case_only_column_references()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "wrong_case",
            Column = "amount",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromNumber(1m)
        });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(new[] { Row(("Amount", 1m)) }, specification));

        Assert.Contains("cannot resolve column 'amount'", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Additive_totals_reject_non_null_nonnumeric_values()
    {
        ReportSpecV1 specification = SumSpecification("Amount");

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(new[] { Row(("Amount", "10")) }, specification));

        Assert.Contains("nonnumeric", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Amount", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Require_all_total_evidence_short_circuits_after_a_false_condition()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(TotalExclusion(requireAll: true));

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Kind", "Detail"), ("Note", new ErrorWrapper(-1)), ("Amount", 10m)),
                Row(("Kind", "Total"), ("Note", "Summary"), ("Amount", 999m))
            },
            specification);

        Assert.Equal(1, result.ExpectedNormalizedRows);
        Assert.Equal(10m, result.ExpectedTotals["total"]);
        Assert.Equal(1, result.RemovedRowsByTransform["exclude_totals"]);
    }

    [Fact]
    public void Require_any_total_evidence_short_circuits_after_a_true_condition()
    {
        ReportSpecV1 specification = SumSpecification("Amount");
        specification.Transforms.Add(TotalExclusion(requireAll: false));

        SourceReconciliationAudit result = _auditor.AuditRows(
            new[]
            {
                Row(("Kind", "Total"), ("Note", new ErrorWrapper(-1)), ("Amount", 999m)),
                Row(("Kind", "Other"), ("Note", "Summary"), ("Amount", 888m)),
                Row(("Kind", "Other"), ("Note", "Detail"), ("Amount", 10m))
            },
            specification);

        Assert.Equal(1, result.ExpectedNormalizedRows);
        Assert.Equal(10m, result.ExpectedTotals["total"]);
        Assert.Equal(2, result.RemovedRowsByTransform["exclude_totals"]);
    }

    [Fact]
    public void Unknown_transform_kind_fails_closed()
    {
        ReportSpecV1 specification = BasicSpecification();
        specification.Transforms.Add(new UnsupportedTransform());

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            _auditor.AuditRows(
                new[] { Row(("Amount", 1m)) },
                specification));

        Assert.Contains("does not support", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Visible_table_totals_row_is_never_audited_as_source_data()
    {
        ReportSpecV1 specification = BasicSpecification();
        var source = FakeExcelTableRange.Create(
            new[] { "Label", "Status", "Amount" },
            new[]
            {
                new object?[] { "A", "Open", 10m },
                new object?[] { "B", "Open", 20m }
            },
            new object?[] { "Total", null, 999m });

        SourceReconciliationAudit result = _auditor.AuditRange(
            source,
            specification,
            expectedSourceRows: 2);

        Assert.Equal(2, result.SourceRows);
        Assert.Equal(2, result.ExpectedNormalizedRows);
        Assert.Equal(30m, result.ExpectedTotals["amount"]);
    }

    [Fact]
    public void Fill_down_state_survives_a_full_source_read_batch_boundary()
    {
        const int rowCount = 16385;
        ReportSpecV1 specification = BasicSpecification();
        specification.Transforms.Add(new FillDownTransform
        {
            Id = "fill_labels",
            Columns = { "Label" }
        });
        specification.Transforms.Add(new FilterRowsTransform
        {
            Id = "keep_rows",
            Column = "Label",
            Operator = RowFilterOperator.Equal,
            Value = ScalarValue.FromText("Keep")
        });
        var data = new object?[rowCount][];
        for (var index = 0; index < rowCount; index++)
        {
            data[index] = new object?[]
            {
                index == 0 ? "Keep" : null,
                "Open",
                1m
            };
        }
        var source = FakeExcelTableRange.Create(
            new[] { "Label", "Status", "Amount" },
            data,
            totals: null);

        SourceReconciliationAudit result = _auditor.AuditRange(
            source,
            specification,
            expectedSourceRows: rowCount);

        Assert.Equal(rowCount, result.ExpectedNormalizedRows);
        Assert.False(result.RemovedRowsByTransform.ContainsKey("keep_rows"));
        Assert.Equal(rowCount, result.ExpectedTotals["amount"]);
    }

    private static ReportSpecV1 BasicSpecification()
    {
        return new ReportSpecV1
        {
            Measures =
            {
                new MeasureDefinition
                {
                    Id = "amount",
                    Label = "Amount",
                    Expression = new AggregateMeasureExpression
                    {
                        Field = "Amount",
                        Function = AggregateFunction.Sum
                    }
                }
            }
        };
    }

    private static ReportSpecV1 SumSpecification(string field)
    {
        return new ReportSpecV1
        {
            Measures =
            {
                new MeasureDefinition
                {
                    Id = "total",
                    Label = "Total",
                    Expression = new AggregateMeasureExpression
                    {
                        Field = field,
                        Function = AggregateFunction.Sum
                    }
                }
            }
        };
    }

    private static ExcludeTotalRowsTransform TotalExclusion(bool requireAll)
    {
        return new ExcludeTotalRowsTransform
        {
            Id = "exclude_totals",
            RequireAllEvidence = requireAll,
            Evidence =
            {
                new TotalRowEvidenceSpec
                {
                    Column = "Kind",
                    MatchKind = TotalRowMatchKind.EqualsAny,
                    Values = { ScalarValue.FromText("Total") },
                    Source = EvidenceSource.UserConfirmation,
                    ObservedMatchCount = 1
                },
                new TotalRowEvidenceSpec
                {
                    Column = "Note",
                    MatchKind = TotalRowMatchKind.StartsWith,
                    Values = { ScalarValue.FromText("Summary") },
                    Source = EvidenceSource.UserConfirmation,
                    ObservedMatchCount = 1
                }
            }
        };
    }

    private static IReadOnlyDictionary<string, object?> Row(
        params (string Column, object? Value)[] values)
    {
        return values.ToDictionary(
            value => value.Column,
            value => value.Value,
            StringComparer.OrdinalIgnoreCase);
    }

    public sealed class FakeExcelTableRange
    {
        private FakeExcelTableRange(object?[,] values, int dataRows, int columns)
        {
            var worksheet = new FakeWorksheet(values);
            Rows = new FakeCount(dataRows + 2);
            Columns = new FakeCount(columns);
            Cells = new FakeCells(values, 0, 0);
            Worksheet = worksheet;
            ListObject = new FakeListObject(
                new FakeSegment(values, worksheet, 0, 0, 1, columns),
                new FakeSegment(values, worksheet, 1, 0, dataRows, columns),
                dataRows,
                columns);
        }

        public FakeCount Rows { get; }

        public FakeCount Columns { get; }

        public FakeCells Cells { get; }

        public FakeWorksheet Worksheet { get; }

        public FakeListObject ListObject { get; }

        public static FakeExcelTableRange Create(
            IReadOnlyList<string> headers,
            IReadOnlyList<object?[]> rows,
            object?[]? totals)
        {
            int totalRows = 1 + rows.Count + (totals == null ? 0 : 1);
            var values = new object?[totalRows, headers.Count];
            for (var column = 0; column < headers.Count; column++)
            {
                values[0, column] = headers[column];
            }
            for (var row = 0; row < rows.Count; row++)
            {
                for (var column = 0; column < headers.Count; column++)
                {
                    values[row + 1, column] = rows[row][column];
                }
            }
            if (totals != null)
            {
                for (var column = 0; column < headers.Count; column++)
                {
                    values[totalRows - 1, column] = totals[column];
                }
            }

            return new FakeExcelTableRange(values, rows.Count, headers.Count);
        }
    }

    public sealed class FakeListObject
    {
        public FakeListObject(
            FakeSegment header,
            FakeSegment data,
            int rows,
            int columns)
        {
            HeaderRowRange = header;
            DataBodyRange = data;
            ListRows = new FakeCount(rows);
            ListColumns = new FakeCount(columns);
        }

        public FakeSegment HeaderRowRange { get; }

        public FakeSegment DataBodyRange { get; }

        public FakeCount ListRows { get; }

        public FakeCount ListColumns { get; }
    }

    public sealed class FakeSegment
    {
        public FakeSegment(
            object?[,] values,
            FakeWorksheet worksheet,
            int startRow,
            int startColumn,
            int rows,
            int columns)
        {
            Cells = new FakeCells(values, startRow, startColumn);
            Worksheet = worksheet;
            Rows = new FakeCount(rows);
            Columns = new FakeCount(columns);
        }

        public FakeCells Cells { get; }

        public FakeWorksheet Worksheet { get; }

        public FakeCount Rows { get; }

        public FakeCount Columns { get; }
    }

    public sealed class FakeCells
    {
        private readonly object?[,] _values;
        private readonly int _startRow;
        private readonly int _startColumn;

        public FakeCells(object?[,] values, int startRow, int startColumn)
        {
            _values = values;
            _startRow = startRow;
            _startColumn = startColumn;
        }

        public FakeCell this[int row, int column] => new(
            _values,
            _startRow + row - 1,
            _startColumn + column - 1);
    }

    public sealed class FakeCell
    {
        public FakeCell(object?[,] values, int row, int column)
        {
            Values = values;
            Row = row;
            Column = column;
        }

        public object?[,] Values { get; }

        public int Row { get; }

        public int Column { get; }

        public object? Value2 => Values[Row, Column];
    }

    public sealed class FakeWorksheet
    {
        public FakeWorksheet(object?[,] values)
        {
            Range = new FakeRangeIndexer(values);
        }

        public FakeRangeIndexer Range { get; }
    }

    public sealed class FakeRangeIndexer
    {
        private readonly object?[,] _values;

        public FakeRangeIndexer(object?[,] values)
        {
            _values = values;
        }

        public FakeValueRange this[FakeCell first, FakeCell last] =>
            new(_values, first.Row, first.Column, last.Row, last.Column);
    }

    public sealed class FakeValueRange
    {
        private readonly object?[,] _values;
        private readonly int _firstRow;
        private readonly int _firstColumn;
        private readonly int _lastRow;
        private readonly int _lastColumn;

        public FakeValueRange(
            object?[,] values,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn)
        {
            _values = values;
            _firstRow = firstRow;
            _firstColumn = firstColumn;
            _lastRow = lastRow;
            _lastColumn = lastColumn;
        }

        public object? Value2
        {
            get
            {
                int rowCount = _lastRow - _firstRow + 1;
                int columnCount = _lastColumn - _firstColumn + 1;
                if (rowCount == 1 && columnCount == 1)
                {
                    return _values[_firstRow, _firstColumn];
                }

                var result = new object?[rowCount, columnCount];
                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        result[row, column] = _values[_firstRow + row, _firstColumn + column];
                    }
                }

                return result;
            }
        }
    }

    public sealed class FakeCount
    {
        public FakeCount(int count)
        {
            Count = count;
        }

        public int Count { get; }
    }

    private sealed class UnsupportedTransform : TransformStep
    {
        public override TransformKind Kind => (TransformKind)int.MaxValue;
    }
}
