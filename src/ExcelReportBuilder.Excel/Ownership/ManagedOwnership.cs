using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelReportBuilder.Excel.Ownership
{
    public enum ManagedObjectKind
    {
        CanonicalTable,
        CanonicalQuery,
        DataModelConnection,
        PivotTable,
        DraftWorksheet,
        PublishedWorksheet,
        RollbackWorksheet,
        ChecksWorksheet,
        Metadata
    }

    public sealed class ManagedObjectIdentity
    {
        public const string MarkerName = "ExcelReportBuilder.Owner";

        public ManagedObjectIdentity(string reportId, string objectId, ManagedObjectKind kind)
        {
            ReportId = RequireIdentifier(reportId, nameof(reportId));
            ObjectId = RequireIdentifier(objectId, nameof(objectId));
            Kind = kind;
        }

        public string ReportId { get; }

        public string ObjectId { get; }

        public ManagedObjectKind Kind { get; }

        public string MarkerValue => string.Format(
            CultureInfo.InvariantCulture,
            "erb:v1:{0}:{1}:{2}",
            ReportId,
            Kind,
            ObjectId);

        public string ExcelName => ManagedName.Create(Kind.ToString(), ReportId, ObjectId);

        public static bool TryParse(string? value, out ManagedObjectIdentity? identity)
        {
            identity = null;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var parts = value!.Split(':');
            if (parts.Length != 5 ||
                !string.Equals(parts[0], "erb", StringComparison.Ordinal) ||
                !string.Equals(parts[1], "v1", StringComparison.Ordinal) ||
                !Enum.TryParse(parts[3], false, out ManagedObjectKind kind))
            {
                return false;
            }

            try
            {
                identity = new ManagedObjectIdentity(parts[2], parts[4], kind);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        private static string RequireIdentifier(string value, string parameter)
        {
            if (string.IsNullOrWhiteSpace(value) || value.IndexOf(':') >= 0)
            {
                throw new ArgumentException("Managed identifiers must be non-empty and cannot contain a colon.", parameter);
            }

            return value;
        }
    }

    public static class ManagedName
    {
        public static string Create(params string[] parts)
        {
            var joined = string.Join("_", parts.Select(Sanitize));
            var result = "ERB_" + joined;
            return result.Length <= 200 ? result : result.Substring(0, 200);
        }

        public static string Worksheet(string label, string stableId)
        {
            var cleanLabel = SanitizeWorksheetText(label);
            var suffix = Sanitize(stableId);
            if (suffix.Length > 8)
            {
                suffix = suffix.Substring(0, 8);
            }

            var reserved = suffix.Length + 3;
            var labelLength = Math.Max(1, 31 - reserved);
            if (cleanLabel.Length > labelLength)
            {
                cleanLabel = cleanLabel.Substring(0, labelLength);
            }

            return cleanLabel + " - " + suffix;
        }

        private static string Sanitize(string value)
        {
            var builder = new StringBuilder(value.Length);
            foreach (var character in value)
            {
                builder.Append(char.IsLetterOrDigit(character) || character == '_' ? character : '_');
            }

            var result = builder.ToString().Trim('_');
            return string.IsNullOrEmpty(result) ? "item" : result;
        }

        private static string SanitizeWorksheetText(string value)
        {
            var forbidden = new HashSet<char>(new[] { ':', '\\', '/', '?', '*', '[', ']' });
            var result = new string(value.Where(character => !forbidden.Contains(character)).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(result) ? "Draft" : result;
        }
    }

    /// <summary>
    /// Guards the only mutable Excel objects. An object without the exact marker
    /// is unmanaged even if its visible name resembles a managed object.
    /// </summary>
    public sealed class ManagedOwnershipGuard
    {
        public bool IsOwned(dynamic excelObject, ManagedObjectIdentity expected)
        {
            if (excelObject == null)
            {
                return false;
            }

            try
            {
                dynamic properties = excelObject.CustomProperties;
                dynamic property = properties.Item(ManagedObjectIdentity.MarkerName);
                var value = Convert.ToString(property.Value, CultureInfo.InvariantCulture);
                return string.Equals(value, expected.MarkerValue, StringComparison.Ordinal);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void MarkOwned(dynamic excelObject, ManagedObjectIdentity identity)
        {
            if (excelObject == null)
            {
                throw new ArgumentNullException(nameof(excelObject));
            }

            try
            {
                dynamic existing = excelObject.CustomProperties.Item(ManagedObjectIdentity.MarkerName);
                existing.Value = identity.MarkerValue;
            }
            catch (Exception)
            {
                excelObject.CustomProperties.Add(ManagedObjectIdentity.MarkerName, identity.MarkerValue);
            }

            if (!IsOwned(excelObject, identity))
            {
                throw new InvalidOperationException("Excel did not retain the managed-object ownership marker.");
            }
        }

        public void DemandOwned(dynamic excelObject, ManagedObjectIdentity expected)
        {
            if (!IsOwned(excelObject, expected))
            {
                throw new InvalidOperationException("The requested Excel object is unmanaged and cannot be changed.");
            }
        }
    }
}
