using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ExcelReportBuilder.Excel.Ownership
{
    public sealed class ManagedObjectRecord
    {
        public string ReportId { get; set; } = string.Empty;

        public string ObjectId { get; set; } = string.Empty;

        public ManagedObjectKind Kind { get; set; }

        public string ExcelName { get; set; } = string.Empty;
    }

    /// <summary>
    /// Records ownership for Excel objects, such as queries and connections,
    /// that cannot carry worksheet CustomProperties themselves.
    /// </summary>
    public sealed class WorkbookOwnershipRegistry
    {
        public const string NamespaceUri = "urn:excel-report-builder:ownership:v1";

        public IReadOnlyList<ManagedObjectRecord> Load(dynamic workbook)
        {
            foreach (var part in EnumerateParts(workbook))
            {
                return Parse((string)part.XML);
            }

            return Array.Empty<ManagedObjectRecord>();
        }

        public bool IsOwned(dynamic workbook, ManagedObjectIdentity identity, string excelName)
        {
            IReadOnlyList<ManagedObjectRecord> records = Load((object)workbook);
            return records.Any(record =>
                string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == identity.Kind &&
                string.Equals(record.ExcelName, excelName, StringComparison.Ordinal));
        }

        public void Register(dynamic workbook, ManagedObjectIdentity identity, string excelName)
        {
            IReadOnlyList<ManagedObjectRecord> loaded = Load((object)workbook);
            var records = loaded.Where(record =>
                !(string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                  string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                  record.Kind == identity.Kind)).ToList();
            records.Add(new ManagedObjectRecord
            {
                ReportId = identity.ReportId,
                ObjectId = identity.ObjectId,
                Kind = identity.Kind,
                ExcelName = excelName
            });
            Save(workbook, records);
        }

        public void RemoveReport(dynamic workbook, string reportId)
        {
            IReadOnlyList<ManagedObjectRecord> loaded = Load((object)workbook);
            Save(workbook, loaded
                .Where(record => !string.Equals(record.ReportId, reportId, StringComparison.Ordinal))
                .ToList());
        }

        private static IReadOnlyList<ManagedObjectRecord> Parse(string xml)
        {
            XNamespace ns = NamespaceUri;
            var document = XDocument.Parse(xml, LoadOptions.None);
            if (document.Root == null || document.Root.Name != ns + "ownership")
            {
                return Array.Empty<ManagedObjectRecord>();
            }

            var result = new List<ManagedObjectRecord>();
            foreach (var element in document.Root.Elements(ns + "object"))
            {
                if (!Enum.TryParse((string?)element.Attribute("kind"), false, out ManagedObjectKind kind))
                {
                    continue;
                }

                result.Add(new ManagedObjectRecord
                {
                    ReportId = (string?)element.Attribute("reportId") ?? string.Empty,
                    ObjectId = (string?)element.Attribute("objectId") ?? string.Empty,
                    Kind = kind,
                    ExcelName = (string?)element.Attribute("excelName") ?? string.Empty
                });
            }

            return result;
        }

        private static void Save(dynamic workbook, IReadOnlyList<ManagedObjectRecord> records)
        {
            foreach (var part in EnumerateParts(workbook).ToList())
            {
                part.Delete();
            }

            XNamespace ns = NamespaceUri;
            var document = new XDocument(
                new XElement(
                    ns + "ownership",
                    new XAttribute("schemaVersion", "1.0"),
                    records.Select(record => new XElement(
                        ns + "object",
                        new XAttribute("reportId", record.ReportId),
                        new XAttribute("objectId", record.ObjectId),
                        new XAttribute("kind", record.Kind),
                        new XAttribute("excelName", record.ExcelName)))));
            workbook.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));
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
