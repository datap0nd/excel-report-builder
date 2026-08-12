using System;
using System.Text;

namespace ExcelReportBuilder.Excel.PivotPlus.NamedSets
{
    /// <summary>
    /// The Core compiler owns raw MDX. Excel's CalculatedMembers.Add transport
    /// uses one apostrophe envelope; readback is normalized only across that
    /// exact, reversible envelope.
    /// </summary>
    internal static class PivotNamedSetFormulaTransport
    {
        private const int MaximumFormulaCharacters = 24 * 1024;

        public static string EncodeForExcel(string rawFormula)
        {
            DemandRaw(rawFormula);
            return "'" + rawFormula.Replace("'", "''") + "'";
        }

        public static bool TryDecodeReadback(string readback, out string rawFormula)
        {
            rawFormula = string.Empty;
            if (!IsBounded(readback)) return false;

            bool startsEnvelope = readback.Length > 0 && readback[0] == '\'';
            bool endsEnvelope = readback.Length > 0 &&
                                readback[readback.Length - 1] == '\'';
            if (!startsEnvelope && !endsEnvelope)
            {
                rawFormula = readback;
                return IsRaw(rawFormula);
            }

            if (!startsEnvelope || !endsEnvelope || readback.Length < 2)
            {
                return false;
            }

            var decoded = new StringBuilder(readback.Length - 2);
            for (var index = 1; index < readback.Length - 1; index++)
            {
                char current = readback[index];
                if (current != '\'')
                {
                    decoded.Append(current);
                    continue;
                }

                if (index + 1 >= readback.Length - 1 || readback[index + 1] != '\'')
                {
                    return false;
                }

                decoded.Append('\'');
                index++;
            }

            rawFormula = decoded.ToString();
            return IsRaw(rawFormula);
        }

        public static string DecodeRequired(string readback)
        {
            if (!TryDecodeReadback(readback, out string rawFormula))
            {
                throw new InvalidOperationException(
                    "Excel exposed a malformed or unsupported named-set formula readback.");
            }

            return rawFormula;
        }

        public static void DemandExactReadback(string readback, string expectedRawFormula)
        {
            DemandRaw(expectedRawFormula);
            string decoded = DecodeRequired(readback);
            if (!string.Equals(decoded, expectedRawFormula, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Excel did not preserve the exact compiled named-set formula.");
            }
        }

        private static void DemandRaw(string rawFormula)
        {
            if (!IsRaw(rawFormula))
            {
                throw new ArgumentException(
                    "A bounded raw Core-compiled named-set formula is required.",
                    nameof(rawFormula));
            }
        }

        private static bool IsRaw(string value)
        {
            return IsBounded(value) &&
                   value.Length > 0 &&
                   value[0] != '\'' &&
                   value[value.Length - 1] != '\'' &&
                   value[0] != '=';
        }

        private static bool IsBounded(string? value)
        {
            if (value == null || value.Length > MaximumFormulaCharacters)
            {
                return false;
            }

            foreach (char character in value)
            {
                if (char.IsControl(character)) return false;
            }

            return true;
        }
    }
}
