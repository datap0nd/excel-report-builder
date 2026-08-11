using System;

namespace ExcelReportBuilder.Excel.Rendering
{
    /// <summary>
    /// Converts untrusted workbook labels to literal Excel text. The leading
    /// apostrophe is Excel's text-entry marker and is not part of the displayed
    /// label. It prevents formula interpretation even when an imported member
    /// begins with a formula trigger after whitespace or control characters.
    /// </summary>
    public static class ExcelLiteralText
    {
        public static string Prepare(string? value)
        {
            var text = value ?? string.Empty;
            return CouldBeInterpretedAsFormula(text) ? "'" + text : text;
        }

        public static bool CouldBeInterpretedAsFormula(string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var index = 0;
            while (index < value.Length &&
                   (char.IsWhiteSpace(value[index]) || char.IsControl(value[index])))
            {
                index++;
            }

            if (index == value.Length)
            {
                return false;
            }

            switch (value[index])
            {
                case '=':
                case '+':
                case '-':
                case '@':
                    return true;
                default:
                    return false;
            }
        }
    }
}
