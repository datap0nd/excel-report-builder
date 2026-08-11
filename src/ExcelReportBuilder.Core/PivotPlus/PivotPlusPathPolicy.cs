using System;

namespace ExcelReportBuilder.Core.PivotPlus
{
    /// <summary>
    /// Shared boundary for identifiers that may be persisted or compiled into
    /// workbook-local source definitions. These values identify workbook
    /// objects or connections and must never carry a file-system location.
    /// </summary>
    public static class PivotPlusPathPolicy
    {
        public static bool IsPathFree(string? value)
        {
            if (value == null || value.Length == 0) return true;

            string candidate = value;

            return candidate.IndexOf('\\') < 0 &&
                   candidate.IndexOf('/') < 0 &&
                   !candidate.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
                   !(candidate.Length >= 2 && char.IsLetter(candidate[0]) && candidate[1] == ':');
        }
    }
}
