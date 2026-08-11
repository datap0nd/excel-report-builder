using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Excel.Ownership
{
    /// <summary>
    /// Gives every logical output worksheet one stable identity, independent of
    /// the number of report blocks anchored on that worksheet.
    /// </summary>
    public static class ManagedOutputIdentity
    {
        public static ManagedObjectIdentity Draft(string reportId, string worksheetName)
        {
            return Create(reportId, worksheetName, ManagedObjectKind.DraftWorksheet);
        }

        public static ManagedObjectIdentity Published(string reportId, string worksheetName)
        {
            return Create(reportId, worksheetName, ManagedObjectKind.PublishedWorksheet);
        }

        public static ManagedObjectIdentity Rollback(string reportId, string worksheetName)
        {
            return Create(reportId, worksheetName, ManagedObjectKind.RollbackWorksheet);
        }

        public static string LogicalKey(string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                throw new ArgumentException("A logical output worksheet name is required.", nameof(worksheetName));
            }

            return worksheetName.Trim().Normalize(NormalizationForm.FormC).ToUpperInvariant();
        }

        private static ManagedObjectIdentity Create(
            string reportId,
            string worksheetName,
            ManagedObjectKind kind)
        {
            var logicalKey = LogicalKey(worksheetName);
            byte[] digest;
            using (var sha256 = SHA256.Create())
            {
                digest = sha256.ComputeHash(Encoding.UTF8.GetBytes(logicalKey));
            }

            var suffix = new StringBuilder(24);
            for (var index = 0; index < 12; index++)
            {
                suffix.Append(digest[index].ToString("x2", CultureInfo.InvariantCulture));
            }

            return new ManagedObjectIdentity(reportId, "output_" + suffix, kind);
        }
    }
}
