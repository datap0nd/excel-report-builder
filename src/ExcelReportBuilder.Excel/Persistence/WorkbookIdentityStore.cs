using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ExcelReportBuilder.Excel.Persistence
{
    /// <summary>
    /// Maintains an anonymous identity inside the workbook so task-pane and
    /// worker lifetimes can refer to the same open workbook without persisting
    /// its file name, path, sheet names, or cell contents.
    /// </summary>
    public sealed class WorkbookIdentityStore
    {
        public const string NamespaceUri = "urn:excel-report-builder:workbook-identity";
        public const string CurrentSchemaVersion = "1.0";

        public string GetOrCreate(dynamic workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            IReadOnlyList<string> identities = LoadAll(workbook);
            if (identities.Count > 1)
            {
                throw new InvalidOperationException(
                    "The workbook contains more than one managed workbook identity.");
            }

            if (identities.Count == 1)
            {
                return identities[0];
            }

            string identity = "workbook_" + Guid.NewGuid().ToString("N");
            XNamespace ns = NamespaceUri;
            var document = new XDocument(
                new XElement(
                    ns + "workbookIdentity",
                    new XAttribute("schemaVersion", CurrentSchemaVersion),
                    new XAttribute("id", identity)));
            workbook.CustomXMLParts.Add(document.ToString(SaveOptions.DisableFormatting));
            return identity;
        }

        public string? Load(dynamic workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            IReadOnlyList<string> identities = LoadAll(workbook);
            if (identities.Count > 1)
            {
                throw new InvalidOperationException(
                    "The workbook contains more than one managed workbook identity.");
            }

            return identities.SingleOrDefault();
        }

        private static IReadOnlyList<string> LoadAll(dynamic workbook)
        {
            var identities = new List<string>();
            foreach (dynamic part in EnumerateParts(workbook))
            {
                identities.Add(ReadIdentity(part));
            }

            return identities;
        }

        private static string ReadIdentity(dynamic part)
        {
            try
            {
                var document = XDocument.Parse((string)part.XML, LoadOptions.None);
                XNamespace ns = NamespaceUri;
                if (document.Root == null || document.Root.Name != ns + "workbookIdentity")
                {
                    throw new InvalidOperationException(
                        "The managed workbook identity has an invalid root element.");
                }

                string? schemaVersion = (string?)document.Root.Attribute("schemaVersion");
                if (!string.Equals(schemaVersion, CurrentSchemaVersion, StringComparison.Ordinal))
                {
                    throw new NotSupportedException("Unknown workbook identity version.");
                }

                string? identity = (string?)document.Root.Attribute("id");
                const string prefix = "workbook_";
                if (identity == null ||
                    !identity.StartsWith(prefix, StringComparison.Ordinal) ||
                    !Guid.TryParseExact(identity.Substring(prefix.Length), "N", out _))
                {
                    throw new InvalidOperationException(
                        "The managed workbook identity is invalid.");
                }

                return identity;
            }
            catch (NotSupportedException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "The managed workbook identity could not be read.",
                    exception);
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

            int count = Convert.ToInt32(parts.Count);
            for (int index = 1; index <= count; index++)
            {
                yield return parts.Item(index);
            }
        }
    }
}
