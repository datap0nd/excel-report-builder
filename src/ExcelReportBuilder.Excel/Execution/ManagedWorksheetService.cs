using System;
using System.Globalization;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ManagedWorksheetService
    {
        private const int SheetVeryHidden = 2;
        private readonly ManagedOwnershipGuard ownershipGuard;

        public ManagedWorksheetService(ManagedOwnershipGuard? ownershipGuard = null)
        {
            this.ownershipGuard = ownershipGuard ?? new ManagedOwnershipGuard();
        }

        public dynamic GetOrCreateDraft(dynamic workbook, ManagedObjectIdentity identity, string label)
        {
            dynamic? ownedWorksheet = TryFindOwnedWorksheet(workbook, identity);
            if (ownedWorksheet != null)
            {
                return ownedWorksheet;
            }

            var preferredName = ManagedName.Worksheet(label, identity.ObjectId);
            dynamic? existing = TryFindWorksheet(workbook, preferredName);
            if (existing != null)
            {
                preferredName = NextAvailableName(workbook, preferredName, identity.ObjectId);
            }

            dynamic worksheet = workbook.Worksheets.Add();
            worksheet.Name = preferredName;
            ownershipGuard.MarkOwned(worksheet, identity);
            return worksheet;
        }

        public dynamic GetOrCreateHidden(dynamic workbook, ManagedObjectIdentity identity)
        {
            var worksheet = GetOrCreateDraft(workbook, identity, "ERB managed");
            worksheet.Visible = SheetVeryHidden;
            return worksheet;
        }

        public void ClearOwned(dynamic worksheet, ManagedObjectIdentity identity)
        {
            ownershipGuard.DemandOwned(worksheet, identity);
            ClearPivotTables(worksheet);
            DeleteListObjects(worksheet);
            worksheet.Cells.Clear();
        }

        public void DeleteOwned(dynamic worksheet, ManagedObjectIdentity identity)
        {
            ownershipGuard.DemandOwned(worksheet, identity);
            worksheet.Delete();
        }

        private static dynamic? TryFindWorksheet(dynamic workbook, string name)
        {
            try
            {
                return workbook.Worksheets.Item(name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private dynamic? TryFindOwnedWorksheet(dynamic workbook, ManagedObjectIdentity identity)
        {
            dynamic worksheets = workbook.Worksheets;
            var count = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            dynamic? match = null;
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = worksheets.Item(index);
                if (!ownershipGuard.IsOwned(worksheet, identity))
                {
                    continue;
                }

                if (match != null)
                {
                    throw new InvalidOperationException(
                        "More than one worksheet carries the same managed-object ownership marker.");
                }

                match = worksheet;
            }

            return match;
        }

        private static void ClearPivotTables(dynamic worksheet)
        {
            dynamic pivotTables = worksheet.PivotTables();
            var count = Convert.ToInt32(pivotTables.Count, CultureInfo.InvariantCulture);
            for (var index = count; index >= 1; index--)
            {
                dynamic pivotTable = pivotTables.Item(index);
                pivotTable.TableRange2.Clear();
            }

            if (Convert.ToInt32(worksheet.PivotTables().Count, CultureInfo.InvariantCulture) != 0)
            {
                throw new InvalidOperationException("Excel did not remove every PivotTable from the managed worksheet.");
            }
        }

        private static void DeleteListObjects(dynamic worksheet)
        {
            dynamic listObjects = worksheet.ListObjects;
            var count = Convert.ToInt32(listObjects.Count, CultureInfo.InvariantCulture);
            for (var index = count; index >= 1; index--)
            {
                dynamic listObject = listObjects.Item(index);
                TryCancelRefresh(listObject);
                listObject.Delete();
            }

            if (Convert.ToInt32(worksheet.ListObjects.Count, CultureInfo.InvariantCulture) != 0)
            {
                throw new InvalidOperationException("Excel did not remove every table from the managed worksheet.");
            }
        }

        private static void TryCancelRefresh(dynamic listObject)
        {
            try
            {
                dynamic queryTable = listObject.QueryTable;
                if (Convert.ToBoolean(queryTable.Refreshing, CultureInfo.InvariantCulture))
                {
                    queryTable.CancelRefresh();
                }
            }
            catch (Exception)
            {
                // A normal worksheet table has no QueryTable. It can be deleted directly.
            }
        }

        private static string NextAvailableName(dynamic workbook, string preferred, string stableId)
        {
            for (var index = 2; index <= 999; index++)
            {
                var suffix = " " + index.ToString(CultureInfo.InvariantCulture);
                var baseName = preferred;
                if (baseName.Length + suffix.Length > 31)
                {
                    baseName = baseName.Substring(0, 31 - suffix.Length);
                }

                var candidate = baseName + suffix;
                if (TryFindWorksheet(workbook, candidate) == null)
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException("A safe managed worksheet name could not be allocated for " + stableId + ".");
        }
    }
}
