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
            return Ensure(workbook, identity);
        }

        /// <summary>
        /// Persists the exact path-free identity previously resolved for this
        /// workbook. A different existing identity is a collision and is never
        /// replaced.
        /// </summary>
        public string Ensure(dynamic workbook, string identity)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (!IsValidIdentity(identity))
            {
                throw new ArgumentException(
                    "A valid path-free managed workbook identity is required.",
                    nameof(identity));
            }

            IReadOnlyList<string> identities = LoadAll(workbook);
            if (identities.Count > 1)
            {
                throw new InvalidOperationException(
                    "The workbook contains more than one managed workbook identity.");
            }

            if (identities.Count == 1)
            {
                if (!string.Equals(identities[0], identity, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The workbook already contains a different managed workbook identity.");
                }

                return identity!;
            }

            XNamespace ns = NamespaceUri;
            var document = new XDocument(
                new XElement(
                    ns + "workbookIdentity",
                    new XAttribute("schemaVersion", CurrentSchemaVersion),
                    new XAttribute("id", identity)));
            dynamic created;
            try
            {
                created = workbook.CustomXMLParts.Add(
                    document.ToString(SaveOptions.DisableFormatting));
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel could not persist the managed workbook identity.",
                    exception);
            }

            if (created == null ||
                !string.Equals(ReadIdentity(created), identity, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Excel did not return the persisted managed workbook identity.");
            }

            IReadOnlyList<string> persisted = LoadAll(workbook);
            if (persisted.Count != 1 ||
                !string.Equals(persisted[0], identity, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Excel did not persist exactly one managed workbook identity.");
            }

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
                if (!IsValidIdentity(identity))
                {
                    throw new InvalidOperationException(
                        "The managed workbook identity is invalid.");
                }

                return identity!;
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

        private static IReadOnlyList<dynamic> EnumerateParts(dynamic workbook)
        {
            try
            {
                dynamic parts = workbook.CustomXMLParts.SelectByNamespace(NamespaceUri);
                int count = Convert.ToInt32(parts.Count);
                if (count < 0 || count > 16)
                {
                    throw new InvalidOperationException(
                        "The workbook contains an invalid number of managed identity parts.");
                }

                var result = new List<dynamic>(count);
                for (int index = 1; index <= count; index++)
                {
                    result.Add(parts.Item(index));
                }

                return result;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "The managed workbook identity parts could not be enumerated.",
                    exception);
            }
        }

        private static bool IsValidIdentity(string? identity)
        {
            const string prefix = "workbook_";
            return identity != null &&
                   identity.StartsWith(prefix, StringComparison.Ordinal) &&
                   Guid.TryParseExact(identity.Substring(prefix.Length), "N", out _);
        }
    }
}
