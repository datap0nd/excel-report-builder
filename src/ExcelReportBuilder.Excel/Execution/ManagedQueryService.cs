using System;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ManagedQueryService
    {
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

            dynamic? existing = TryGetQuery(workbook, queryName);
            if (existing != null)
            {
                if (!registry.IsOwned(workbook, identity, queryName))
                {
                    throw new InvalidOperationException("A query with the requested name exists but is unmanaged.");
                }

                existing.Delete();
            }

            dynamic query = workbook.Queries.Add(queryName, restrictedFormula);
            registry.Register(workbook, identity, queryName);
            return query;
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
