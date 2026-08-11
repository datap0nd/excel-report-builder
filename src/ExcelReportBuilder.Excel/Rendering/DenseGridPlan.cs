using System;
using System.Collections.Generic;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Rendering
{
    public enum DenseCellValueKind
    {
        Blank,
        Text,
        Number,
        Date,
        Formula
    }

    public enum DenseFormulaExpectationKind
    {
        Number,
        Blank,
        Error
    }

    /// <summary>
    /// The independently evaluated result expected from a host-generated
    /// formula. This value is calculated from typed measure nodes and native
    /// PivotTable aggregate inputs, never by parsing or evaluating formula text.
    /// </summary>
    public sealed class DenseFormulaExpectation
    {
        private DenseFormulaExpectation(DenseFormulaExpectationKind kind, decimal? numericValue)
        {
            Kind = kind;
            NumericValue = numericValue;
        }

        public DenseFormulaExpectationKind Kind { get; }

        public decimal? NumericValue { get; }

        internal static DenseFormulaExpectation Number(decimal value)
        {
            return new DenseFormulaExpectation(DenseFormulaExpectationKind.Number, value);
        }

        internal static DenseFormulaExpectation Blank()
        {
            return new DenseFormulaExpectation(DenseFormulaExpectationKind.Blank, null);
        }

        internal static DenseFormulaExpectation Error()
        {
            return new DenseFormulaExpectation(DenseFormulaExpectationKind.Error, null);
        }
    }

    public sealed class SafeExcelFormula
    {
        public const int MaximumFormulaCharacters = 8192;

        internal SafeExcelFormula(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value[0] != '=')
            {
                throw new ArgumentException("A host-generated Excel formula is required.", nameof(value));
            }

            if (value.Length > MaximumFormulaCharacters)
            {
                throw new InvalidOperationException(
                    "The generated formula exceeds Excel's supported formula length.");
            }

            Value = value;
        }

        internal string Value { get; }
    }

    public static class SafeFormulaFactory
    {
        internal static SafeExcelFormula FromPivotExpression(string hostGeneratedExpression)
        {
            if (string.IsNullOrWhiteSpace(hostGeneratedExpression) ||
                hostGeneratedExpression.IndexOf("GETPIVOTDATA(", StringComparison.Ordinal) < 0)
            {
                throw new ArgumentException("Only a host-generated PivotTable expression is accepted.", nameof(hostGeneratedExpression));
            }

            return new SafeExcelFormula(hostGeneratedExpression);
        }

        internal static SafeExcelFormula FromTypedMeasure(string formula)
        {
            return new SafeExcelFormula(formula);
        }
    }

    public sealed class DenseCellWrite
    {
        public int RelativeRow { get; set; }

        public int RelativeColumn { get; set; }

        public DenseCellValueKind Kind { get; set; }

        public object? Value { get; set; }

        public SafeExcelFormula? Formula { get; set; }

        public DenseFormulaExpectation? ExpectedFormulaValue { get; set; }

        public string? NumberFormat { get; set; }

        public string? StyleId { get; set; }

        public string? MeasureId { get; set; }

        public bool IsOutputTotal { get; set; }

        public int IndentLevel { get; set; }

        public int ColumnSpan { get; set; } = 1;
    }

    public sealed class DenseDimensionSize
    {
        public int RelativeIndex { get; set; }

        public double Size { get; set; }
    }

    public sealed class DenseGridPlan
    {
        public string BlockId { get; set; } = string.Empty;

        public string AnchorCell { get; set; } = "A1";

        public int OwnedRowCount { get; set; }

        public int OwnedColumnCount { get; set; }

        public bool FreezeHeaders { get; set; }

        public int FreezeRelativeRow { get; set; }

        public List<DenseCellWrite> Cells { get; set; } = new List<DenseCellWrite>();

        public List<DenseDimensionSize> RowSizes { get; set; } = new List<DenseDimensionSize>();

        public List<DenseDimensionSize> ColumnSizes { get; set; } = new List<DenseDimensionSize>();

        public IReadOnlyDictionary<string, PresentationStyleSpec> Styles { get; set; } =
            new Dictionary<string, PresentationStyleSpec>(StringComparer.OrdinalIgnoreCase);
    }
}
