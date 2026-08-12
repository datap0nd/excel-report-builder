using System;
using System.Text.RegularExpressions;

namespace ExcelReportBuilder.Core.PivotPlus
{
    public enum PivotPlusWorkbookObjectKind
    {
        Table,
        NamedRange
    }

    /// <summary>
    /// Creates the smallest workbook-only Power Query needed to place a table
    /// or named range in Excel's Data Model. The compiler accepts an identifier,
    /// never an arbitrary query or connector.
    /// </summary>
    public static class PivotPlusSourceQueryCompiler
    {
        private static readonly Regex WorkbookObjectPattern = new Regex(
            @"^[A-Za-z_][A-Za-z0-9_.]*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static string Compile(string workbookObjectName, PivotPlusWorkbookObjectKind kind)
        {
            if (string.IsNullOrWhiteSpace(workbookObjectName) ||
                !PivotPlusPathPolicy.IsPathFree(workbookObjectName) ||
                !WorkbookObjectPattern.IsMatch(workbookObjectName))
            {
                throw new ArgumentException(
                    "A valid Excel table or named-range identifier is required.",
                    nameof(workbookObjectName));
            }

            var source = "Excel.CurrentWorkbook(){[Name=" +
                MString(workbookObjectName) +
                "]}[Content]";

            switch (kind)
            {
                case PivotPlusWorkbookObjectKind.Table:
                    return "let\n    Source = " + source + "\nin\n    Source";
                case PivotPlusWorkbookObjectKind.NamedRange:
                    return "let\n    RawSource = " + source +
                        ",\n    Source = Table.PromoteHeaders(RawSource, " +
                        "[PromoteAllScalars = true, Culture = \"en-US\"])" +
                        "\nin\n    Source";
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static string MString(string value)
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
