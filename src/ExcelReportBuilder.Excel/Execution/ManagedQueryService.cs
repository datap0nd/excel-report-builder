using System;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ManagedQueryService
    {
        private const string FormulaContractKind = "workbook-query-formula-v1";
        private readonly WorkbookOwnershipRegistry registry;

        public ManagedQueryService(WorkbookOwnershipRegistry? registry = null)
        {
            this.registry = registry ?? new WorkbookOwnershipRegistry();
        }

        public dynamic ReplaceQuery(
            dynamic workbook,
            ManagedObjectIdentity identity,
            string queryName,
            string restrictedFormula)
        {
            if (identity.Kind != ManagedObjectKind.CanonicalQuery)
            {
                throw new ArgumentException("The ownership identity must describe a canonical query.", nameof(identity));
            }

            if (string.IsNullOrWhiteSpace(restrictedFormula))
            {
                throw new ArgumentException("A compiled query formula is required.", nameof(restrictedFormula));
            }

            var records = registry.Load((object)workbook);
            var exact = records.Where(record =>
                string.Equals(record.ReportId, identity.ReportId, StringComparison.Ordinal) &&
                string.Equals(record.ObjectId, identity.ObjectId, StringComparison.Ordinal) &&
                record.Kind == identity.Kind).ToList();
            if (exact.Count > 1)
            {
                throw new InvalidOperationException(
                    "More than one ownership record claims the managed query identity.");
            }

            if (records.Any(record =>
                    record.Kind == ManagedObjectKind.CanonicalQuery &&
                    !exact.Contains(record) &&
                    string.Equals(record.ExcelName, queryName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException(
                    "The requested managed query name is already owned by another object.");
            }

            var registration = exact.SingleOrDefault();
            dynamic? existing = TryGetQuery(workbook, queryName);
            if (existing != null)
            {
                if (registration == null)
                {
                    throw new InvalidOperationException("A query with the requested name exists but is unmanaged.");
                }

                if (!string.Equals(registration.ExcelName, queryName, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "The managed query identity is registered under a different name.");
                }

                var liveFormula = ReadFormula(existing);
                if (!ManagedContentFingerprint.Matches(
                        registration.SourceContract,
                        FormulaContractKind,
                        liveFormula))
                {
                    throw new InvalidOperationException(
                        "The managed query formula no longer matches its exact owned contract.");
                }

                // WorkbookQuery.Formula is read/write. Updating the exact owned
                // query in place preserves worksheet and Data Model connections
                // that already depend on its stable name.
                existing.Formula = restrictedFormula;
                registry.Register(
                    workbook,
                    identity,
                    queryName,
                    null,
                    ManagedContentFingerprint.Create(
                        FormulaContractKind,
                        ReadFormula(existing)));
                return existing;
            }

            if (registration != null)
            {
                dynamic? differentlyNamed = TryGetQuery(workbook, registration.ExcelName);
                if (differentlyNamed != null)
                {
                    throw new InvalidOperationException(
                        "The managed query identity is registered under a different live query name.");
                }

                registry.Remove((object)workbook, new[] { identity });
            }

            dynamic query = workbook.Queries.Add(queryName, restrictedFormula);
            registry.Register(
                workbook,
                identity,
                queryName,
                null,
                ManagedContentFingerprint.Create(
                    FormulaContractKind,
                    ReadFormula(query)));
            return query;
        }

        private static string ReadFormula(dynamic query)
        {
            try
            {
                return Convert.ToString(query.Formula, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the managed query formula for ownership verification.",
                    exception);
            }
        }

        private static dynamic? TryGetQuery(dynamic workbook, string queryName)
        {
            try
            {
                return workbook.Queries.Item(queryName);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
