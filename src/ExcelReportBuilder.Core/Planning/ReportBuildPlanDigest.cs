using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using ExcelReportBuilder.Core.Specifications;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelReportBuilder.Core.Planning
{
    /// <summary>
    /// Detects mutation of a validated build plan between planning and host
    /// execution. Hosts must compare this digest with ReportBuildPlan.PlanHash.
    /// </summary>
    public static class ReportBuildPlanDigest
    {
        public static string Compute(ReportBuildPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var json = JsonConvert.SerializeObject(
                plan,
                Formatting.None,
                ReportSpecJson.CreateSerializerSettings());
            var document = JObject.Parse(json);
            document.Remove("planHash");
            var canonical = document.ToString(Formatting.None);
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(canonical));
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
