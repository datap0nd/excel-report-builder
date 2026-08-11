using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace ExcelReportBuilder.Core.Specifications
{
    /// <summary>
    /// Binds an immutable build plan to the exact ReportSpec content from
    /// which it was created. Hosts must reject a plan/spec digest mismatch.
    /// </summary>
    public static class ReportSpecDigest
    {
        public static string Compute(ReportSpecV1 specification)
        {
            if (specification == null)
            {
                throw new ArgumentNullException(nameof(specification));
            }

            var canonicalJson = ReportSpecJson.Serialize(specification, Formatting.None);
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(canonicalJson));
                var result = new StringBuilder(bytes.Length * 2);
                foreach (var value in bytes)
                {
                    result.Append(value.ToString("x2", CultureInfo.InvariantCulture));
                }

                return result.ToString();
            }
        }
    }
}
