using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Core.PivotPlus.Calculations
{
    /// <summary>
    /// Versioned fingerprints for validated compiler definitions and exact
    /// generated formula text. Live Excel artifact fingerprints belong to the
    /// host layer; this utility is not a raw-DAX proposal or execution contract.
    /// </summary>
    public static class PivotDaxFingerprint
    {
        public const string FormulaAlgorithm = "measure.formula.v1:sha256";
        public const string DefinitionAlgorithm = "measure.definition.v1:sha256";

        public static string ComputeFormula(string daxFormula)
        {
            if (daxFormula == null) throw new ArgumentNullException(nameof(daxFormula));

            var writer = new PivotFingerprintCanonicalWriter();
            writer.Add("formula", daxFormula);
            return Hash(FormulaAlgorithm, writer.ToString());
        }

        internal static string ComputeDefinition(string canonicalDefinition)
        {
            if (canonicalDefinition == null)
            {
                throw new ArgumentNullException(nameof(canonicalDefinition));
            }

            return Hash(DefinitionAlgorithm, canonicalDefinition);
        }

        internal static void AddFormat(
            PivotFingerprintCanonicalWriter writer,
            PivotMeasureFormat format)
        {
            writer.Add("formatKind", ((int)format.Kind).ToString(CultureInfo.InvariantCulture));
            writer.Add("decimalPlaces", format.DecimalPlaces.ToString(CultureInfo.InvariantCulture));
            writer.Add("thousands", format.UseThousandsSeparator ? "1" : "0");
            writer.Add("currency", format.CurrencySymbolOrCode ?? string.Empty);
        }

        private static string Hash(string algorithm, string canonical)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(canonical));
                var result = new StringBuilder(bytes.Length * 2);
                foreach (byte value in bytes)
                {
                    result.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }

                return algorithm + ":" + result;
            }
        }
    }

    internal sealed class PivotFingerprintCanonicalWriter
    {
        private readonly StringBuilder builder = new StringBuilder();

        public void Add(string name, string value)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));
            if (value == null) throw new ArgumentNullException(nameof(value));

            builder.Append(name.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(name);
            builder.Append('=');
            builder.Append(value.Length.ToString(CultureInfo.InvariantCulture));
            builder.Append(':');
            builder.Append(value);
            builder.Append(';');
        }

        public override string ToString()
        {
            return builder.ToString();
        }
    }
}
