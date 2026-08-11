using ExcelReportBuilder.Core.Measures;
using ExcelReportBuilder.Core.Profiling;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Transforms;

namespace ExcelReportBuilder.Core.Tests;

internal static class SyntheticReportFactory
{
    public static SourceProfile CreateLongProfile(long rowCount = 3)
    {
        return new SourceProfile
        {
            RowCount = rowCount,
            ColumnCount = 6,
            Columns =
            {
                Column(0, "Period", SourceValueType.Date, rowCount, dateLike: rowCount),
                Column(1, "Region", SourceValueType.Text, rowCount),
                Column(2, "Category", SourceValueType.Text, rowCount),
                Column(3, "Amount", SourceValueType.DecimalNumber, rowCount, numeric: rowCount),
                Column(4, "Units", SourceValueType.WholeNumber, rowCount, numeric: rowCount),
                Column(5, "Weight", SourceValueType.DecimalNumber, rowCount, numeric: rowCount)
            }
        };
    }

    public static ReportSpecV1 CreateValidLongSpec()
    {
        var amount = new MeasureDefinition
        {
            Id = "amount",
            Label = "Amount",
            ValueType = MeasureValueType.Currency,
            NumberFormat = "#,##0.00",
            Expression = new AggregateMeasureExpression
            {
                Field = "Amount",
                Function = AggregateFunction.Sum,
                ResultType = MeasureValueType.Currency
            }
        };
        var units = new MeasureDefinition
        {
            Id = "units",
            Label = "Units",
            ValueType = MeasureValueType.Number,
            Expression = new AggregateMeasureExpression
            {
                Field = "Units",
                Function = AggregateFunction.Sum,
                ResultType = MeasureValueType.Number
            }
        };
        var averagePrice = new MeasureDefinition
        {
            Id = "average_price",
            Label = "Average price",
            ValueType = MeasureValueType.Currency,
            Expression = new SafeDivideMeasureExpression
            {
                ResultType = MeasureValueType.Currency,
                Numerator = Reference("amount", MeasureValueType.Currency),
                Denominator = Reference("units", MeasureValueType.Number),
                OnZero = ZeroDenominatorBehavior.Blank,
                AsPercentage = false
            }
        };
        var share = new MeasureDefinition
        {
            Id = "share",
            Label = "Share",
            ValueType = MeasureValueType.Percentage,
            NumberFormat = "0.0%",
            Expression = new ShareMeasureExpression
            {
                ResultType = MeasureValueType.Percentage,
                Part = Reference("amount", MeasureValueType.Currency),
                Whole = Reference("amount", MeasureValueType.Currency),
                OnZero = ZeroDenominatorBehavior.Blank,
                Scope = ShareDenominatorScope.FilteredReportTotal
            }
        };
        var weighted = new MeasureDefinition
        {
            Id = "weighted_rate",
            Label = "Weighted rate",
            ValueType = MeasureValueType.Number,
            Expression = new WeightedAggregateMeasureExpression
            {
                ResultType = MeasureValueType.Number,
                Numerator = new FilteredAggregateMeasureExpression
                {
                    Field = "WeightedUnits",
                    Function = AggregateFunction.Sum,
                    ResultType = MeasureValueType.Number,
                    Filters =
                    {
                        NotBlank("Units"),
                        NotBlank("Weight")
                    }
                },
                Denominator = new FilteredAggregateMeasureExpression
                {
                    Field = "Weight",
                    Function = AggregateFunction.Sum,
                    ResultType = MeasureValueType.Number,
                    Filters =
                    {
                        NotBlank("Units"),
                        NotBlank("Weight")
                    }
                },
                OnZero = ZeroDenominatorBehavior.Blank
            }
        };
        var ratio = new MeasureDefinition
        {
            Id = "unit_ratio",
            Label = "Unit ratio",
            ValueType = MeasureValueType.Number,
            Expression = new RatioMeasureExpression
            {
                ResultType = MeasureValueType.Number,
                Numerator = Reference("units", MeasureValueType.Number),
                Denominator = Reference("units", MeasureValueType.Number),
                OnZero = ZeroDenominatorBehavior.Blank
            }
        };
        var filteredAmount = new MeasureDefinition
        {
            Id = "filtered_amount",
            Label = "Filtered amount",
            ValueType = MeasureValueType.Currency,
            Expression = new FilteredAggregateMeasureExpression
            {
                Field = "Amount",
                Function = AggregateFunction.Sum,
                ResultType = MeasureValueType.Currency,
                Filters =
                {
                    new MeasureFilterSpec
                    {
                        Field = "Category",
                        Operator = MeasureFilterOperator.Equal,
                        Values = { ScalarValue.FromText("Core") }
                    }
                }
            }
        };

        return new ReportSpecV1
        {
            Id = "monthly_report",
            Name = "Monthly report",
            OwnershipId = "owned_report",
            Source = new WorkbookSourceSpec
            {
                Kind = WorkbookSourceKind.Table,
                WorkbookObjectName = "SourceData",
                HeaderRowCount = 1,
                Fingerprint = SourceFingerprint.FromHeaders(new[]
                {
                    "Period", "Region", "Category", "Amount", "Units", "Weight"
                })
            },
            PeriodMapping = new PeriodMappingSpec
            {
                Id = "periods",
                Kind = PeriodMappingKind.LongDateColumn,
                DateColumn = "Period"
            },
            Transforms =
            {
                new TrimTextTransform { Id = "trim_labels", Columns = { "Region", "Category" } },
                new DerivePeriodPartsTransform
                {
                    Id = "derive_periods",
                    DateColumn = "Period",
                    Columns =
                    {
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.Year, OutputColumn = "Year" },
                        new DerivedPeriodColumnSpec { Part = DerivedPeriodPart.Quarter, OutputColumn = "Quarter" }
                    }
                },
                new AddArithmeticColumnTransform
                {
                    Id = "derive_weighted_units",
                    OutputColumn = "WeightedUnits",
                    Operator = ArithmeticOperator.Multiply,
                    Left = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Units" },
                    Right = new ArithmeticOperand { Kind = ArithmeticOperandKind.Column, Column = "Weight" },
                    ResultType = ColumnDataType.DecimalNumber
                }
            },
            Measures = { amount, units, averagePrice, share, weighted, ratio, filteredAmount },
            Styles =
            {
                new PresentationStyleSpec
                {
                    Id = "header",
                    Bold = true,
                    FillColor = "#D9EAF7",
                    BottomBorder = true
                },
                new PresentationStyleSpec
                {
                    Id = "subtotal",
                    Bold = true,
                    TopBorder = true
                }
            },
            Blocks =
            {
                new ReportBlockSpec
                {
                    Id = "summary",
                    OwnershipId = "owned_summary",
                    Title = "Management summary",
                    WorksheetName = "Report",
                    AnchorCell = "B3",
                    OutputMode = ReportOutputMode.DenseGrid,
                    OwnedExtent = new OwnedRangeExtentSpec { RowCount = 500, ColumnCount = 6 },
                    HeaderStyleId = "header",
                    SubtotalStyleId = "subtotal",
                    PeriodSlices =
                    {
                        new PeriodSliceSpec
                        {
                            Id = "current",
                            Label = "Current",
                            Kind = PeriodSliceKind.Current,
                            SelectedStart = new DateTime(2026, 3, 1),
                            SelectedEnd = new DateTime(2026, 3, 31)
                        },
                        new PeriodSliceSpec
                        {
                            Id = "prior",
                            Label = "Prior",
                            Kind = PeriodSliceKind.Prior,
                            BasedOnSliceId = "current"
                        },
                        new PeriodSliceSpec
                        {
                            Id = "selected",
                            Label = "Selected",
                            Kind = PeriodSliceKind.Selected,
                            SelectedStart = new DateTime(2026, 1, 1),
                            SelectedEnd = new DateTime(2026, 3, 31)
                        },
                        new PeriodSliceSpec
                        {
                            Id = "same_prior_year",
                            Label = "Same period prior year",
                            Kind = PeriodSliceKind.SamePeriodPriorYear,
                            BasedOnSliceId = "selected"
                        }
                    },
                    Headers =
                    {
                        new ReportHeaderSpec
                        {
                            Text = "Management summary",
                            RelativeRow = 0,
                            RelativeColumn = 0,
                            ColumnSpan = 4,
                            StyleId = "header"
                        }
                    },
                    Spacers =
                    {
                        new SpacerSpec { Axis = SpacerAxis.Row, BeforeLevel = 1, Count = 1 }
                    },
                    Layout = new ReportLayoutSpec
                    {
                        Rows =
                        {
                            new FieldPlacementSpec
                            {
                                Field = "Region",
                                Subtotals = new SubtotalSpec
                                {
                                    Mode = SubtotalMode.Automatic,
                                    Placement = TotalPlacement.AfterMembers,
                                    StyleId = "subtotal"
                                },
                                MemberOrder =
                                {
                                    ScalarValue.FromText("North"),
                                    ScalarValue.FromText("South")
                                },
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
                                    IncludeOthers = true,
                                    OthersLabel = "Others"
                                }
                            },
                            new FieldPlacementSpec
                            {
                                Field = "Category",
                                Subtotals = new SubtotalSpec { Mode = SubtotalMode.None }
                            }
                        },
                        Columns =
                        {
                            new FieldPlacementSpec
                            {
                                Field = "Quarter",
                                Subtotals = new SubtotalSpec { Mode = SubtotalMode.None }
                            }
                        },
                        Values =
                        {
                            new ValuePlacementSpec
                            {
                                MeasureId = "amount",
                                PeriodSliceIds = { "current", "prior", "selected", "same_prior_year" }
                            },
                            new ValuePlacementSpec { MeasureId = "average_price" },
                            new ValuePlacementSpec { MeasureId = "share" },
                            new ValuePlacementSpec { MeasureId = "weighted_rate" }
                        },
                        GrandTotals = new GrandTotalsSpec
                        {
                            ShowRows = true,
                            ShowColumns = true,
                            RowPlacement = TotalPlacement.AfterMembers,
                            ColumnPlacement = TotalPlacement.AfterMembers,
                            StyleId = "subtotal"
                        }
                    }
                },
                new ReportBlockSpec
                {
                    Id = "detail",
                    OwnershipId = "owned_detail",
                    WorksheetName = "Report",
                    AnchorCell = "J3",
                    OutputMode = ReportOutputMode.MetricStack,
                    OwnedExtent = new OwnedRangeExtentSpec { RowCount = 500, ColumnCount = 6 },
                    Layout = new ReportLayoutSpec
                    {
                        Rows =
                        {
                            new FieldPlacementSpec
                            {
                                Field = "Category",
                                Subtotals = new SubtotalSpec { Mode = SubtotalMode.Automatic }
                            }
                        },
                        Values = { new ValuePlacementSpec { MeasureId = "units" } }
                    }
                }
            },
            Checks =
            {
                new ReportCheckSpec
                {
                    Id = "preserve_amount",
                    Kind = ReportCheckKind.TotalPreservation,
                    MeasureId = "amount",
                    Tolerance = 0.01m
                }
            }
        };
    }

    private static ReferenceMeasureExpression Reference(string id, MeasureValueType valueType)
    {
        return new ReferenceMeasureExpression { MeasureId = id, ResultType = valueType };
    }

    private static MeasureFilterSpec NotBlank(string field)
    {
        return new MeasureFilterSpec
        {
            Field = field,
            Operator = MeasureFilterOperator.IsNotBlank
        };
    }

    private static SourceColumnProfile Column(
        int index,
        string name,
        SourceValueType type,
        long count,
        long dateLike = 0,
        long numeric = 0)
    {
        return new SourceColumnProfile
        {
            Index = index,
            Name = name,
            InferredType = type,
            NonBlankCount = count,
            DistinctCount = count,
            DateLikeCount = dateLike,
            NumericCount = numeric
        };
    }
}
