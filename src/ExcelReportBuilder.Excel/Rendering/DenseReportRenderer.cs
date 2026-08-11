using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Rendering
{
    public sealed class DenseRenderResult
    {
        public int CellsWritten { get; set; }

        public int FormulasWritten { get; set; }

        public string BlockId { get; set; } = string.Empty;
    }

    public sealed class DenseReportRenderer
    {
        private const int BorderEdgeTop = 8;
        private const int BorderEdgeBottom = 9;
        private const int LineStyleContinuous = 1;
        private readonly ManagedOwnershipGuard ownershipGuard;

        public DenseReportRenderer(ManagedOwnershipGuard? ownershipGuard = null)
        {
            this.ownershipGuard = ownershipGuard ?? new ManagedOwnershipGuard();
        }

        public DenseRenderResult Render(
            dynamic worksheet,
            ManagedObjectIdentity worksheetIdentity,
            DenseGridPlan plan,
            IExcelProgressSink? progressSink = null)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            ownershipGuard.DemandOwned(worksheet, worksheetIdentity);
            progressSink = progressSink ?? NullExcelProgressSink.Instance;
            var anchor = ValidatePlan(plan);

            dynamic ownedRange = worksheet.Range[
                worksheet.Cells[anchor.Row, anchor.Column],
                worksheet.Cells[
                    anchor.Row + plan.OwnedRowCount - 1,
                    anchor.Column + plan.OwnedColumnCount - 1]];
            ownedRange.UnMerge();
            ownedRange.Clear();

            progressSink.Report(new ExcelProgress
            {
                Stage = ExcelBuildStage.Rendering,
                Operation = "Rendering dense block " + plan.BlockId + ".",
                ManagedObject = plan.BlockId
            });

            var formulas = 0;
            foreach (var write in plan.Cells)
            {
                var row = anchor.Row + write.RelativeRow;
                var column = anchor.Column + write.RelativeColumn;
                dynamic cell = worksheet.Cells[row, column];

                switch (write.Kind)
                {
                    case DenseCellValueKind.Blank:
                        cell.ClearContents();
                        break;
                    case DenseCellValueKind.Text:
                        // Set the cell to Text before assigning any workbook-
                        // derived label. Formula-like labels also receive the
                        // Excel literal prefix as defense in depth.
                        cell.NumberFormat = "@";
                        cell.Value2 = ExcelLiteralText.Prepare(
                            Convert.ToString(write.Value, CultureInfo.InvariantCulture));
                        break;
                    case DenseCellValueKind.Number:
                        cell.Value2 = Convert.ToDouble(write.Value, CultureInfo.InvariantCulture);
                        break;
                    case DenseCellValueKind.Date:
                        cell.Value2 = ((DateTime)write.Value!).ToOADate();
                        break;
                    case DenseCellValueKind.Formula:
                        if (write.Formula == null)
                        {
                            throw new InvalidOperationException("A formula cell is missing its host-generated formula.");
                        }

                        cell.Formula = write.Formula.Value;
                        formulas++;
                        break;
                    default:
                        throw new NotSupportedException("The dense cell kind is not supported.");
                }

                if (!string.IsNullOrWhiteSpace(write.NumberFormat))
                {
                    cell.NumberFormat = write.NumberFormat;
                }

                if (!string.IsNullOrWhiteSpace(write.StyleId))
                {
                    if (!plan.Styles.TryGetValue(write.StyleId!, out var style))
                    {
                        throw new InvalidOperationException("The dense cell references an unknown presentation style.");
                    }

                    ApplyStyle(cell, style);
                }

                if (write.IndentLevel > 0)
                {
                    cell.IndentLevel = write.IndentLevel;
                }

                if (write.ColumnSpan > 1)
                {
                    dynamic span = worksheet.Range[
                        worksheet.Cells[row, column],
                        worksheet.Cells[row, column + write.ColumnSpan - 1]];
                    span.Merge();
                }
            }

            foreach (var rowSize in plan.RowSizes)
            {
                worksheet.Rows[anchor.Row + rowSize.RelativeIndex].RowHeight = rowSize.Size;
            }

            foreach (var columnSize in plan.ColumnSizes)
            {
                worksheet.Columns[anchor.Column + columnSize.RelativeIndex].ColumnWidth = columnSize.Size;
            }

            if (plan.FreezeHeaders && plan.FreezeRelativeRow > 0)
            {
                worksheet.Activate();
                dynamic window = worksheet.Application.ActiveWindow;
                window.FreezePanes = false;
                window.SplitRow = anchor.Row + plan.FreezeRelativeRow - 1;
                window.FreezePanes = true;
            }

            return new DenseRenderResult
            {
                BlockId = plan.BlockId,
                CellsWritten = plan.Cells.Count,
                FormulasWritten = formulas
            };
        }

        internal static CellAddress ValidatePlan(DenseGridPlan plan)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            var anchor = CellAddress.Parse(plan.AnchorCell);
            Validate(plan, anchor);
            return anchor;
        }

        private static void Validate(DenseGridPlan plan, CellAddress anchor)
        {
            if (plan.OwnedRowCount < 1 || plan.OwnedColumnCount < 1)
            {
                throw new InvalidOperationException("A dense block requires a positive owned extent.");
            }

            if (plan.FreezeRelativeRow < 0 || plan.FreezeRelativeRow > plan.OwnedRowCount)
            {
                throw new InvalidOperationException("The freeze-header row exceeds the block's managed owned extent.");
            }

            foreach (var rowSize in plan.RowSizes)
            {
                if (rowSize.RelativeIndex < 0 || rowSize.RelativeIndex >= plan.OwnedRowCount ||
                    rowSize.Size <= 0d || rowSize.Size > 409d)
                {
                    throw new InvalidOperationException("A dense row spacer exceeds its managed extent or Excel size limit.");
                }
            }

            foreach (var columnSize in plan.ColumnSizes)
            {
                if (columnSize.RelativeIndex < 0 || columnSize.RelativeIndex >= plan.OwnedColumnCount ||
                    columnSize.Size <= 0d || columnSize.Size > 255d)
                {
                    throw new InvalidOperationException("A dense column spacer exceeds its managed extent or Excel size limit.");
                }
            }

            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (var cell in plan.Cells)
            {
                if (cell.RelativeRow < 0 || cell.RelativeColumn < 0 || cell.ColumnSpan < 1)
                {
                    throw new InvalidOperationException("Dense cell coordinates and spans must be non-negative.");
                }

                if (cell.IndentLevel < 0 || cell.IndentLevel > 15)
                {
                    throw new InvalidOperationException("Dense cell indentation must be between 0 and 15.");
                }

                if (cell.Kind == DenseCellValueKind.Formula)
                {
                    if (cell.Formula == null)
                    {
                        throw new InvalidOperationException("A formula cell is missing its host-generated formula.");
                    }

                    if (cell.ExpectedFormulaValue == null)
                    {
                        throw new InvalidOperationException(
                            "Every managed formula requires an independent typed result expectation.");
                    }

                    if (string.IsNullOrWhiteSpace(cell.MeasureId))
                    {
                        throw new InvalidOperationException(
                            "Every managed formula requires a typed Value identifier for validation.");
                    }
                }
                else if (cell.ExpectedFormulaValue != null)
                {
                    throw new InvalidOperationException(
                        "Only managed formula cells may carry a typed result expectation.");
                }

                var absoluteRow = anchor.Row + cell.RelativeRow;
                var finalColumn = anchor.Column + cell.RelativeColumn + cell.ColumnSpan - 1;
                if (cell.RelativeRow >= plan.OwnedRowCount ||
                    cell.RelativeColumn + cell.ColumnSpan > plan.OwnedColumnCount)
                {
                    throw new InvalidOperationException(
                        "A dense cell write exceeds the block's managed owned extent.");
                }

                if (absoluteRow > 1_048_576 || finalColumn > 16_384)
                {
                    throw new InvalidOperationException("The dense block would exceed the worksheet boundary.");
                }

                for (var offset = 0; offset < cell.ColumnSpan; offset++)
                {
                    var key = absoluteRow.ToString(CultureInfo.InvariantCulture) + ":" +
                              (anchor.Column + cell.RelativeColumn + offset).ToString(CultureInfo.InvariantCulture);
                    if (!occupied.Add(key))
                    {
                        throw new InvalidOperationException("Dense block cells overlap.");
                    }
                }
            }
        }

        private static void ApplyStyle(dynamic cell, PresentationStyleSpec style)
        {
            cell.Font.Bold = style.Bold;
            cell.Font.Italic = style.Italic;
            if (!string.IsNullOrWhiteSpace(style.FontColor))
            {
                cell.Font.Color = ParseExcelColor(style.FontColor!);
            }

            if (!string.IsNullOrWhiteSpace(style.FillColor))
            {
                cell.Interior.Color = ParseExcelColor(style.FillColor!);
            }

            switch (style.HorizontalAlignment)
            {
                case HorizontalAlignment.Left: cell.HorizontalAlignment = -4131; break;
                case HorizontalAlignment.Center: cell.HorizontalAlignment = -4108; break;
                case HorizontalAlignment.Right: cell.HorizontalAlignment = -4152; break;
                default: cell.HorizontalAlignment = 1; break;
            }

            if (!string.IsNullOrWhiteSpace(style.NumberFormat))
            {
                cell.NumberFormat = style.NumberFormat;
            }

            if (style.TopBorder)
            {
                cell.Borders[BorderEdgeTop].LineStyle = LineStyleContinuous;
            }

            if (style.BottomBorder)
            {
                cell.Borders[BorderEdgeBottom].LineStyle = LineStyleContinuous;
            }
        }

        private static int ParseExcelColor(string value)
        {
            if (!Regex.IsMatch(value, "^#[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant))
            {
                throw new InvalidOperationException("Presentation colors must use #RRGGBB.");
            }

            var red = int.Parse(value.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var green = int.Parse(value.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var blue = int.Parse(value.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return red | (green << 8) | (blue << 16);
        }
    }

    internal readonly struct CellAddress
    {
        private static readonly Regex Pattern = new Regex(
            @"^\$?(?<column>[A-Z]{1,3})\$?(?<row>[1-9][0-9]{0,6})$",
            RegexOptions.CultureInvariant);

        public CellAddress(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }

        public int Column { get; }

        public static CellAddress Parse(string value)
        {
            var match = Pattern.Match(value ?? string.Empty);
            if (!match.Success)
            {
                throw new ArgumentException("A valid A1 cell address is required.", nameof(value));
            }

            var column = 0;
            foreach (var character in match.Groups["column"].Value)
            {
                column = checked(column * 26 + character - 'A' + 1);
            }

            var row = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture);
            if (row > 1_048_576 || column > 16_384)
            {
                throw new ArgumentException("The cell address exceeds the worksheet boundary.", nameof(value));
            }

            return new CellAddress(row, column);
        }
    }
}
