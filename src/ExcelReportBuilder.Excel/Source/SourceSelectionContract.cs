using System;
using System.Globalization;

namespace ExcelReportBuilder.Excel.Source
{
    /// <summary>
    /// Keeps the profiled Excel selection and the object later exposed through
    /// Excel.CurrentWorkbook on the same exact range. A selection contained in
    /// a table is not equivalent to the whole table unless their external
    /// addresses are identical.
    /// </summary>
    internal static class SourceSelectionContract
    {
        public static object? GetExactContainingTableOrThrow(object selectionObject)
        {
            if (selectionObject == null)
            {
                throw new ArgumentNullException(nameof(selectionObject));
            }

            dynamic selection = selectionObject;
            object? tableObject;
            try
            {
                tableObject = selection.ListObject as object;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose table membership for exact source validation.",
                    exception);
            }

            if (tableObject == null)
            {
                return null;
            }

            dynamic table = tableObject;
            object? tableRangeObject;
            try
            {
                tableRangeObject = table.Range as object;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the containing table range for exact source validation.",
                    exception);
            }

            if (tableRangeObject == null || !HaveSameExternalAddress(selectionObject, tableRangeObject))
            {
                throw new InvalidOperationException(
                    "The selected range is inside an Excel table but is not the complete table. " +
                    "Select the entire table, including its header row, or choose a range outside the table.");
            }

            bool showHeaders;
            try
            {
                showHeaders = Convert.ToBoolean(table.ShowHeaders, CultureInfo.InvariantCulture);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the containing table's header state for exact source validation.",
                    exception);
            }

            if (!showHeaders)
            {
                throw new InvalidOperationException(
                    "The selected Excel table must show its header row before it can be used as Data.");
            }

            return tableObject;
        }

        private static bool HaveSameExternalAddress(object leftObject, object rightObject)
        {
            try
            {
                dynamic left = leftObject;
                dynamic right = rightObject;
                string leftAddress = Convert.ToString(
                    left.Address[true, true, 1, true],
                    CultureInfo.InvariantCulture) ?? string.Empty;
                string rightAddress = Convert.ToString(
                    right.Address[true, true, 1, true],
                    CultureInfo.InvariantCulture) ?? string.Empty;
                return leftAddress.Length > 0 && string.Equals(
                    leftAddress,
                    rightAddress,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose a stable address for exact source validation.",
                    exception);
            }
        }
    }
}
