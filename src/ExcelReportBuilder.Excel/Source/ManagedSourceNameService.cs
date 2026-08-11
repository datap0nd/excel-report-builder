using System;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Source
{
    /// <summary>
    /// Makes a rectangular selection addressable through Excel.CurrentWorkbook
    /// without changing any source value. Existing table names are reused.
    /// </summary>
    public sealed class ManagedSourceNameService
    {
        private const string RefersToContractKind = "workbook-name-refers-to-v1";
        private readonly WorkbookOwnershipRegistry registry;

        public ManagedSourceNameService(WorkbookOwnershipRegistry? registry = null)
        {
            this.registry = registry ?? new WorkbookOwnershipRegistry();
        }

        public string EnsureWorkbookObject(
            dynamic workbook,
            dynamic selection,
            string reportId,
            string objectId)
        {
            object? tableObject = SourceSelectionContract.GetExactContainingTableOrThrow((object)selection);
            if (tableObject != null)
            {
                string? tableName = Convert.ToString(
                    ((dynamic)tableObject).Name,
                    CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(tableName))
                {
                    throw new InvalidOperationException(
                        "Excel did not expose a stable name for the selected table.");
                }

                return tableName!;
            }

            var identity = new ManagedObjectIdentity(reportId, objectId, ManagedObjectKind.Metadata);
            var name = ManagedName.Create("Source", reportId, objectId);
            var requestedReference = ToAbsoluteExternalReference(selection);
            var records = registry.Load((object)workbook);
            var exact = records.Where(record =>
                string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == identity.Kind).ToList();
            if (exact.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the managed source name identity.");
            }

            if (records.Any(record =>
                    record.Kind == ManagedObjectKind.Metadata &&
                    !exact.Contains(record) &&
                    string.Equals(record.ExcelName, name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "The requested managed source name is already owned by another object.");
            }

            var registration = exact.SingleOrDefault();
            dynamic? existing = TryGetName(workbook, name);
            if (existing != null && registration == null)
            {
                throw new InvalidOperationException("A workbook name with the requested source identity exists but is unmanaged.");
            }

            dynamic managedName;
            if (existing == null)
            {
                if (registration != null)
                {
                    dynamic? differentlyNamed = TryGetName(workbook, registration.ExcelName);
                    if (differentlyNamed != null)
                    {
                        throw new InvalidOperationException(
                            "The managed source identity is registered under a different live workbook name.");
                    }

                    registry.Remove((object)workbook, new[] { identity });
                }

                managedName = workbook.Names.Add(name, requestedReference);
            }
            else
            {
                if (!string.Equals(registration!.ExcelName, name, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The managed source identity is registered under a different workbook name.");
                }

                var liveReference = ReadReference(existing);
                if (!ManagedContentFingerprint.Matches(
                        registration.SourceContract,
                        RefersToContractKind,
                        liveReference))
                {
                    throw new InvalidOperationException(
                        "The managed source workbook name no longer matches its exact owned reference.");
                }

                existing.RefersTo = requestedReference;
                managedName = existing;
            }

            registry.Register(
                workbook,
                identity,
                name,
                null,
                ManagedContentFingerprint.Create(
                    RefersToContractKind,
                    ReadReference(managedName)));
            return name;
        }

        private static string ReadReference(dynamic workbookName)
        {
            try
            {
                return Convert.ToString(
                    workbookName.RefersTo,
                    CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the managed source workbook reference for ownership verification.",
                    exception);
            }
        }

        private static string ToAbsoluteExternalReference(dynamic selection)
        {
            string address;
            try
            {
                address = Convert.ToString(
                    selection.Address[true, true, 1, true],
                    CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose a stable address for the managed source range.",
                    exception);
            }

            if (string.IsNullOrWhiteSpace(address))
            {
                throw new InvalidOperationException(
                    "Excel did not expose a stable address for the managed source range.");
            }

            return "=" + address;
        }

        private static dynamic? TryGetName(dynamic workbook, string name)
        {
            try
            {
                return workbook.Names.Item(name);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
