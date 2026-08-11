using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReportBuilder.Excel.Rendering
{
    public sealed class PivotFilterItem
    {
        public string Field { get; set; } = string.Empty;

        public object? Value { get; set; }
    }

    /// <summary>
    /// Generates only GETPIVOTDATA expressions from typed arguments. Model text
    /// never enters this class as an executable formula.
    /// </summary>
    public sealed class GetPivotDataFormulaBuilder
    {
        private static readonly Regex A1Address = new Regex(
            @"^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}$",
            RegexOptions.CultureInvariant);

        public string Build(
            string dataFieldCaption,
            string pivotWorksheet,
            string pivotAnchor,
            IReadOnlyList<PivotFilterItem>? filters = null)
        {
            if (string.IsNullOrWhiteSpace(dataFieldCaption))
            {
                throw new ArgumentException("A PivotTable value caption is required.", nameof(dataFieldCaption));
            }

            if (string.IsNullOrWhiteSpace(pivotWorksheet))
            {
                throw new ArgumentException("A PivotTable worksheet is required.", nameof(pivotWorksheet));
            }

            if (!A1Address.IsMatch(pivotAnchor ?? string.Empty))
            {
                throw new ArgumentException("A valid A1 PivotTable anchor is required.", nameof(pivotAnchor));
            }

            return "=IFERROR(" + BuildExpression(
                dataFieldCaption,
                pivotWorksheet,
                pivotAnchor!,
                filters) + ",\"\")";
        }

        public SafeExcelFormula BuildSafe(
            string dataFieldCaption,
            string pivotWorksheet,
            string pivotAnchor,
            IReadOnlyList<PivotFilterItem>? filters = null)
        {
            return SafeFormulaFactory.FromPivotExpression(Build(
                dataFieldCaption,
                pivotWorksheet,
                pivotAnchor,
                filters));
        }

        public string BuildExpression(
            string dataFieldCaption,
            string pivotWorksheet,
            string pivotAnchor,
            IReadOnlyList<PivotFilterItem>? filters = null)
        {
            if (string.IsNullOrWhiteSpace(dataFieldCaption))
            {
                throw new ArgumentException("A PivotTable value caption is required.", nameof(dataFieldCaption));
            }

            if (string.IsNullOrWhiteSpace(pivotWorksheet))
            {
                throw new ArgumentException("A PivotTable worksheet is required.", nameof(pivotWorksheet));
            }

            if (!A1Address.IsMatch(pivotAnchor ?? string.Empty))
            {
                throw new ArgumentException("A valid A1 PivotTable anchor is required.", nameof(pivotAnchor));
            }

            var builder = new StringBuilder();
            builder.Append("GETPIVOTDATA(");
            builder.Append(ExcelString(dataFieldCaption));
            builder.Append(",'");
            builder.Append(pivotWorksheet.Replace("'", "''"));
            builder.Append("'!");
            builder.Append(pivotAnchor);

            foreach (var filter in filters ?? Array.Empty<PivotFilterItem>())
            {
                if (string.IsNullOrWhiteSpace(filter.Field))
                {
                    throw new ArgumentException("Pivot filter fields cannot be blank.", nameof(filters));
                }

                builder.Append(',');
                builder.Append(ExcelString(filter.Field));
                builder.Append(',');
                builder.Append(ExcelLiteral(filter.Value));
            }

            builder.Append(')');
            return builder.ToString();
        }

        public string SafeDivide(string numeratorCell, string denominatorCell)
        {
            if (!A1Address.IsMatch(numeratorCell ?? string.Empty) ||
                !A1Address.IsMatch(denominatorCell ?? string.Empty))
            {
                throw new ArgumentException("Safe division accepts only A1 cell references.");
            }

            return "=IF(OR(" + denominatorCell + "=0," + denominatorCell + "=\"\"),\"\"," +
                   numeratorCell + "/" + denominatorCell + ")";
        }

        private static string ExcelString(string value)
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static string ExcelLiteral(object? value)
        {
            if (value == null)
            {
                return "\"\"";
            }

            if (value is bool boolean)
            {
                return boolean ? "TRUE" : "FALSE";
            }

            if (value is byte || value is sbyte || value is short || value is ushort ||
                value is int || value is uint || value is long || value is ulong ||
                value is float || value is double || value is decimal)
            {
                return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "0";
            }

            if (value is DateTime date)
            {
                var dateExpression = "DATE(" + date.Year + "," + date.Month + "," + date.Day + ")";
                if (date.TimeOfDay == TimeSpan.Zero)
                {
                    return dateExpression;
                }

                var seconds = date.Second + (date.Millisecond / 1000m);
                return dateExpression + "+TIME(" + date.Hour + "," + date.Minute + "," +
                    seconds.ToString(CultureInfo.InvariantCulture) + ")";
            }

            return ExcelString(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }
}
