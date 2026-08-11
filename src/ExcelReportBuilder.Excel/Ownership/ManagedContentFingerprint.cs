using System;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Excel.Ownership
{
    internal static class ManagedContentFingerprint
    {
        public static string Create(string contractKind, string value)
        {
            if (string.IsNullOrWhiteSpace(contractKind))
            {
                throw new ArgumentException("A managed contract kind is required.", nameof(contractKind));
            }

            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var sha256 = SHA256.Create())
            {
                var digest = sha256.ComputeHash(Encoding.UTF8.GetBytes(value));
                return contractKind + ":sha256:" + Convert.ToBase64String(digest);
            }
        }

        public static bool Matches(string? persistedContract, string contractKind, string liveValue)
        {
            return !string.IsNullOrWhiteSpace(persistedContract) &&
                   string.Equals(
                       persistedContract,
                       Create(contractKind, liveValue),
                       StringComparison.Ordinal);
        }
    }
}
