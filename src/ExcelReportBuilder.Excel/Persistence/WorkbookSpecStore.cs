using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelReportBuilder.Core.Specifications;

namespace ExcelReportBuilder.Excel.Persistence
{
    /// <summary>
    /// Stores report specifications in a versioned Custom XML part. The payload
    /// is base64 so report labels cannot change the XML structure.
    /// </summary>
    public sealed class WorkbookSpecStore
    {
        public const string NamespaceUri = "urn:excel-report-builder:report-spec:v1";

        public void Save(dynamic workbook, ReportSpecV1 reportSpec)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (reportSpec == null)
            {
                throw new ArgumentNullException(nameof(reportSpec));
            }

            if (!string.Equals(reportSpec.SchemaVersion, ReportSpecV1.CurrentSchemaVersion, StringComparison.Ordinal))
            {
                throw new NotSupportedException("Unknown report specification version.");
            }

            if (string.IsNullOrWhiteSpace(reportSpec.Id))
            {
                throw new ArgumentException("A report identifier is required.", nameof(reportSpec));
            }

            RemoveExisting(workbook, reportSpec.Id);
            var json = ReportSpecJson.Serialize(reportSpec);
            var payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
            XNamespace ns = NamespaceUri;
            var document = new XDocument(
                new XElement(
                    ns + "reportSpec",
                    new XAttribute("id", reportSpec.Id),
                    new XAttribute("schemaVersion", reportSpec.SchemaVersion),
                    new XElement(ns + "payload", new XAttribute("encoding", "base64"), payload)));
            workbook.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));
        }

        public IReadOnlyList<ReportSpecV1> LoadAll(dynamic workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = new List<ReportSpecV1>();
            foreach (var part in EnumerateParts(workbook))
            {
                var spec = TryRead(part);
                if (spec != null)
                {
                    result.Add(spec);
                }
            }

            return result;
        }

        public ReportSpecV1? Load(dynamic workbook, string reportId)
        {
            IReadOnlyList<ReportSpecV1> specifications = LoadAll((object)workbook);
            return specifications.SingleOrDefault(spec => string.Equals(spec.Id, reportId, StringComparison.Ordinal));
        }

        private static void RemoveExisting(dynamic workbook, string reportId)
        {
            foreach (var part in EnumerateParts(workbook))
            {
                try
                {
                    var document = XDocument.Parse((string)part.XML, LoadOptions.None);
                    if (document.Root != null &&
                        string.Equals((string?)document.Root.Attribute("id"), reportId, StringComparison.Ordinal))
                    {
                        part.Delete();
                    }
                }
                catch (Exception)
                {
                    // Foreign or malformed Custom XML parts are not ours and are
                    // deliberately left untouched.
                }
            }
        }

        private static ReportSpecV1? TryRead(dynamic part)
        {
            try
            {
                var document = XDocument.Parse((string)part.XML, LoadOptions.None);
                XNamespace ns = NamespaceUri;
                if (document.Root == null || document.Root.Name != ns + "reportSpec")
                {
                    return null;
                }

                var version = (string?)document.Root.Attribute("schemaVersion");
                if (!string.Equals(version, ReportSpecV1.CurrentSchemaVersion, StringComparison.Ordinal))
                {
                    throw new NotSupportedException("Unknown report specification version.");
                }

                var payload = document.Root.Element(ns + "payload");
                if (payload == null || !string.Equals((string?)payload.Attribute("encoding"), "base64", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("The report specification payload is invalid.");
                }

                var json = Encoding.UTF8.GetString(Convert.FromBase64String(payload.Value));
                return ReportSpecJson.Deserialize(json);
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException("A managed report specification could not be read.", exception);
            }
        }

        private static IEnumerable<dynamic> EnumerateParts(dynamic workbook)
        {
            dynamic parts;
            try
            {
                parts = workbook.CustomXMLParts.SelectByNamespace(NamespaceUri);
            }
            catch (Exception)
            {
                yield break;
            }

            var count = Convert.ToInt32(parts.Count);
            for (var index = 1; index <= count; index++)
            {
                yield return parts.Item(index);
            }
        }
    }
}
