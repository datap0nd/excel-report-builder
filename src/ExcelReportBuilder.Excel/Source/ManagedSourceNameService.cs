using System;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Source
{
    /// <summary>
    /// Makes a rectangular selection addressable through Excel.CurrentWorkbook
    /// without changing any source value. Existing table names are reused.
    /// </summary>
    public sealed class ManagedSourceNameService
    {
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
            var tableName = TryGetTableName(selection);
            if (!string.IsNullOrWhiteSpace(tableName))
            {
                return tableName!;
            }

            var identity = new ManagedObjectIdentity(reportId, objectId, ManagedObjectKind.Metadata);
            var name = ManagedName.Create("Source", reportId, objectId);
            dynamic? existing = TryGetName(workbook, name);
            if (existing != null && !registry.IsOwned(workbook, identity, name))
            {
                throw new InvalidOperationException("A workbook name with the requested source identity exists but is unmanaged.");
            }

            if (existing == null)
            {
                workbook.Names.Add(name, selection);
            }
            else
            {
                existing.RefersToRange = selection;
            }

            registry.Register(workbook, identity, name);
            return name;
        }

        private static string? TryGetTableName(dynamic selection)
        {
            try
            {
                dynamic table = selection.ListObject;
                return table == null ? null : Convert.ToString(table.Name);
            }
            catch (Exception)
            {
                return null;
            }
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
