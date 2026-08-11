using System.Collections.Generic;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Core.Transforms
{
    public enum TransformKind
    {
        SelectColumns,
        KeepColumns,
        RemoveColumns,
        ReorderColumns,
        RenameColumn,
        ChangeColumnType,
        TrimText,
        ReplaceValue,
        NormalizeBlanks,
        NormalizeErrors,
        FillDown,
        MapValues,
        FilterRows,
        ExcludeTotalRows,
        DerivePeriodParts,
        AddArithmeticColumn,
        NormalizePeriods
    }

    public enum ColumnDataType
    {
        Text,
        WholeNumber,
        DecimalNumber,
        Boolean,
        Date,
        DateTime
    }

    public enum RowFilterOperator
    {
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanOrEqual,
        LessThan,
        LessThanOrEqual,
        Contains,
        StartsWith,
        EndsWith,
        IsBlank,
        IsNotBlank
    }

    public enum TotalRowMatchKind
    {
        EqualsAny,
        StartsWith,
        Contains,
        IsBlank
    }

    public enum EvidenceSource
    {
        Profile,
        Preview,
        UserConfirmation
    }

    public enum DerivedPeriodPart
    {
        Year,
        Half,
        Quarter,
        MonthNumber,
        MonthName,
        YearMonth
    }

    public enum ArithmeticOperator
    {
        Add,
        Subtract,
        Multiply,
        Divide
    }

    public enum ArithmeticOperandKind
    {
        Column,
        Number
    }

    public abstract class TransformStep
    {
        public string Id { get; set; } = string.Empty;

        public abstract TransformKind Kind { get; }
    }

    public sealed class KeepColumnsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.KeepColumns;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class SelectColumnsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.SelectColumns;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class RemoveColumnsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.RemoveColumns;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class ReorderColumnsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.ReorderColumns;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class RenameColumnTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.RenameColumn;

        public string From { get; set; } = string.Empty;

        public string To { get; set; } = string.Empty;
    }

    public sealed class ChangeColumnTypeTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.ChangeColumnType;

        public string Column { get; set; } = string.Empty;

        public ColumnDataType DataType { get; set; }
    }

    public sealed class TrimTextTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.TrimText;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class ReplaceValueTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.ReplaceValue;

        public string Column { get; set; } = string.Empty;

        public ScalarValue Find { get; set; } = ScalarValue.Null();

        public ScalarValue ReplaceWith { get; set; } = ScalarValue.Null();
    }

    public sealed class NormalizeBlanksTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.NormalizeBlanks;

        public List<string> Columns { get; set; } = new List<string>();

        public ScalarValue Replacement { get; set; } = ScalarValue.Null();

        public bool TreatWhitespaceAsBlank { get; set; } = true;
    }

    public sealed class NormalizeErrorsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.NormalizeErrors;

        public List<string> Columns { get; set; } = new List<string>();

        public ScalarValue Replacement { get; set; } = ScalarValue.Null();
    }

    public sealed class FillDownTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.FillDown;

        public List<string> Columns { get; set; } = new List<string>();
    }

    public sealed class MapValuesTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.MapValues;

        public string Column { get; set; } = string.Empty;

        public List<ValueMapEntry> Entries { get; set; } = new List<ValueMapEntry>();
    }

    public sealed class ValueMapEntry
    {
        public ScalarValue From { get; set; } = ScalarValue.Null();

        public ScalarValue To { get; set; } = ScalarValue.Null();
    }

    public sealed class FilterRowsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.FilterRows;

        public string Column { get; set; } = string.Empty;

        public RowFilterOperator Operator { get; set; }

        public ScalarValue? Value { get; set; }
    }

    public sealed class ExcludeTotalRowsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.ExcludeTotalRows;

        public List<TotalRowEvidenceSpec> Evidence { get; set; } = new List<TotalRowEvidenceSpec>();

        public bool RequireAllEvidence { get; set; }
    }

    public sealed class TotalRowEvidenceSpec
    {
        public string Column { get; set; } = string.Empty;

        public TotalRowMatchKind MatchKind { get; set; }

        public List<ScalarValue> Values { get; set; } = new List<ScalarValue>();

        public EvidenceSource Source { get; set; }

        public long ObservedMatchCount { get; set; }
    }

    public sealed class DerivePeriodPartsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.DerivePeriodParts;

        public string DateColumn { get; set; } = string.Empty;

        public List<DerivedPeriodColumnSpec> Columns { get; set; } = new List<DerivedPeriodColumnSpec>();
    }

    public sealed class DerivedPeriodColumnSpec
    {
        public DerivedPeriodPart Part { get; set; }

        public string OutputColumn { get; set; } = string.Empty;
    }

    public sealed class AddArithmeticColumnTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.AddArithmeticColumn;

        public string OutputColumn { get; set; } = string.Empty;

        public ArithmeticOperator Operator { get; set; }

        public ArithmeticOperand Left { get; set; } = new ArithmeticOperand();

        public ArithmeticOperand Right { get; set; } = new ArithmeticOperand();

        public ColumnDataType ResultType { get; set; } = ColumnDataType.DecimalNumber;

        public bool ReturnNullOnZeroDenominator { get; set; } = true;
    }

    public sealed class ArithmeticOperand
    {
        public ArithmeticOperandKind Kind { get; set; }

        public string? Column { get; set; }

        public decimal? Number { get; set; }
    }

    public sealed class NormalizePeriodsTransform : TransformStep
    {
        public override TransformKind Kind => TransformKind.NormalizePeriods;

        public string PeriodMappingId { get; set; } = string.Empty;
    }

    public sealed class PeriodColumnMapping
    {
        public string SourceColumn { get; set; } = string.Empty;

        public int Month { get; set; }

        public int? Year { get; set; }

        public string? Metric { get; set; }
    }
}
