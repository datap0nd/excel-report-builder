using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelReportBuilder.Excel.Source
{
    public sealed class SourceSelectionSnapshot
    {
        public string WorkbookObjectName { get; set; } = string.Empty;

        public long RowCount { get; set; }

        public int ColumnCount { get; set; }

        public IReadOnlyList<string> Headers { get; set; } = Array.Empty<string>();

        public IReadOnlyList<IReadOnlyList<object?>> SampleRows { get; set; } = Array.Empty<IReadOnlyList<object?>>();
    }

    public sealed class SourceSelectionInspector
    {
        public const int DefaultSampleRows = 100;

        public SourceSelectionSnapshot Inspect(dynamic excelApplication, int maximumSampleRows = DefaultSampleRows)
        {
            if (excelApplication == null)
            {
                throw new ArgumentNullException(nameof(excelApplication));
            }

            if (maximumSampleRows < 1 || maximumSampleRows > 1000)
            {
                throw new ArgumentOutOfRangeException(nameof(maximumSampleRows));
            }

            object? selectionObject = excelApplication.Selection as object;
            if (selectionObject == null)
            {
                throw new InvalidOperationException("Select one rectangular table or range.");
            }

            dynamic selection = selectionObject;
            if (Convert.ToInt32(selection.Areas.Count, CultureInfo.InvariantCulture) != 1)
            {
                throw new InvalidOperationException("Select one rectangular table or range.");
            }

            var rowCount = Convert.ToInt64(selection.Rows.Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(selection.Columns.Count, CultureInfo.InvariantCulture);
            if (rowCount < 2 || columnCount < 1)
            {
                throw new InvalidOperationException("The selected source must contain one header row and at least one data row.");
            }

            if (columnCount > 16384)
            {
                throw new InvalidOperationException("The selected source exceeds Excel's column capacity.");
            }

            var boundedRows = (int)Math.Min(rowCount, maximumSampleRows + 1L);
            dynamic sampleRange = selection.Resize[boundedRows, columnCount];
            var values = sampleRange.Value2;
            var headers = new List<string>(columnCount);
            var rows = new List<IReadOnlyList<object?>>(Math.Max(0, boundedRows - 1));

            for (var column = 1; column <= columnCount; column++)
            {
                var header = Convert.ToString(Read(values, 1, column), CultureInfo.InvariantCulture)?.Trim();
                if (string.IsNullOrWhiteSpace(header))
                {
                    throw new InvalidOperationException("Every selected source column must have a non-blank header.");
                }

                if (headers.Exists(existing => string.Equals(existing, header, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new InvalidOperationException("Source headers must be unique.");
                }

                headers.Add(header);
            }

            for (var row = 2; row <= boundedRows; row++)
            {
                var resultRow = new object?[columnCount];
                for (var column = 1; column <= columnCount; column++)
                {
                    resultRow[column - 1] = Read(values, row, column);
                }

                rows.Add(resultRow);
            }

            var objectName = TryGetTableName(selection) ?? CreateRangeName(selection);
            return new SourceSelectionSnapshot
            {
                WorkbookObjectName = objectName,
                RowCount = rowCount - 1,
                ColumnCount = columnCount,
                Headers = headers,
                SampleRows = rows
            };
        }

        private static object? Read(object values, int row, int column)
        {
            if (values is Array array && array.Rank == 2)
            {
                return array.GetValue(row, column);
            }

            return row == 1 && column == 1 ? values : null;
        }

        private static string? TryGetTableName(dynamic selection)
        {
            try
            {
                dynamic listObject = selection.ListObject;
                return listObject == null ? null : Convert.ToString(listObject.Name, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string CreateRangeName(dynamic selection)
        {
            var sheetName = Convert.ToString(selection.Worksheet.Name, CultureInfo.InvariantCulture) ?? "Data";
            var address = Convert.ToString(selection.Address[true, true, 1, true], CultureInfo.InvariantCulture) ?? string.Empty;
            return sheetName + "!" + address;
        }
    }
}
