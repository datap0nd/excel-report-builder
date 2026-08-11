using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Execution
{
    public sealed class ManagedWorksheetService
    {
        private const int SheetVeryHidden = 2;
        private const string HiddenPivotSheetSuffix = "_pivot_sheet";
        private readonly ManagedOwnershipGuard ownershipGuard;
        private readonly WorkbookOwnershipRegistry ownershipRegistry;

        public ManagedWorksheetService(
            ManagedOwnershipGuard? ownershipGuard = null,
            WorkbookOwnershipRegistry? ownershipRegistry = null)
        {
            this.ownershipGuard = ownershipGuard ?? new ManagedOwnershipGuard();
            this.ownershipRegistry = ownershipRegistry ?? new WorkbookOwnershipRegistry();
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

        /// <summary>
        /// Reuses the stable managed draft identity, then removes every object
        /// left by the previous rendering pass. A draft worksheet is wholly
        /// owned by the report, so rebuilding it must start from an empty sheet
        /// rather than clearing only the rectangles in the new plan. This also
        /// removes blocks that were moved, shrunk, or deleted from the setup.
        /// </summary>
        public dynamic GetOrCreateClearedDraft(
            dynamic workbook,
            ManagedObjectIdentity identity,
            string label)
        {
            dynamic worksheet = GetOrCreateDraft(workbook, identity, label);
            ClearOwned(worksheet, identity);
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

        public bool DeleteOwnedWorksheetIfPresent(
            dynamic excelApplication,
            dynamic workbook,
            ManagedObjectIdentity identity)
        {
            if (excelApplication == null) throw new ArgumentNullException(nameof(excelApplication));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (identity == null) throw new ArgumentNullException(nameof(identity));

            dynamic? worksheet = TryFindOwnedWorksheet(workbook, identity);
            if (worksheet == null)
            {
                return false;
            }

            bool previousAlerts = Convert.ToBoolean(
                excelApplication.DisplayAlerts,
                CultureInfo.InvariantCulture);
            try
            {
                excelApplication.DisplayAlerts = false;
                ownershipGuard.DemandOwned(worksheet, identity);
                worksheet.Delete();
                return true;
            }
            finally
            {
                excelApplication.DisplayAlerts = previousAlerts;
            }
        }

        public void DeleteStaleHiddenPivotWorksheets(
            dynamic excelApplication,
            dynamic workbook,
            string reportId,
            IReadOnlyCollection<string> activeObjectIds)
        {
            if (excelApplication == null) throw new ArgumentNullException(nameof(excelApplication));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(reportId))
            {
                throw new ArgumentException("A report identifier is required.", nameof(reportId));
            }

            if (activeObjectIds == null) throw new ArgumentNullException(nameof(activeObjectIds));
            var active = new HashSet<string>(activeObjectIds, StringComparer.Ordinal);
            var stale = new List<StaleHiddenPivotWorksheet>();
            dynamic worksheets = workbook.Worksheets;
            int count = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = worksheets.Item(index);
                ManagedObjectIdentity? parsedIdentity = ReadIdentity(worksheet);
                if (parsedIdentity == null)
                {
                    continue;
                }

                ManagedObjectIdentity identity = parsedIdentity;
                if (identity.Kind != ManagedObjectKind.PivotTable ||
                    !string.Equals(identity.ReportId, reportId, StringComparison.Ordinal) ||
                    identity.ObjectId.Length <= HiddenPivotSheetSuffix.Length ||
                    !identity.ObjectId.EndsWith(HiddenPivotSheetSuffix, StringComparison.Ordinal) ||
                    active.Contains(identity.ObjectId))
                {
                    continue;
                }

                stale.Add(new StaleHiddenPivotWorksheet
                {
                    Worksheet = worksheet,
                    WorksheetIdentity = identity,
                    PivotIdentity = new ManagedObjectIdentity(
                        reportId,
                        identity.ObjectId.Substring(
                            0,
                            identity.ObjectId.Length - HiddenPivotSheetSuffix.Length),
                        ManagedObjectKind.PivotTable)
                });
            }

            if (stale.Count == 0)
            {
                return;
            }

            bool previousAlerts = Convert.ToBoolean(
                excelApplication.DisplayAlerts,
                CultureInfo.InvariantCulture);
            var removedRegistryIdentities = new List<ManagedObjectIdentity>();
            try
            {
                excelApplication.DisplayAlerts = false;
                foreach (StaleHiddenPivotWorksheet item in stale)
                {
                    ownershipGuard.DemandOwned(item.Worksheet, item.WorksheetIdentity);
                    item.Worksheet.Delete();
                    removedRegistryIdentities.Add(item.PivotIdentity);
                }
            }
            finally
            {
                excelApplication.DisplayAlerts = previousAlerts;
                ownershipRegistry.Remove((object)workbook, removedRegistryIdentities);
            }
        }

        /// <summary>
        /// Removes draft worksheets that belonged to a logical output in an
        /// earlier version of the same report but are absent from the current
        /// validated plan. Published and rollback worksheets are deliberately
        /// excluded because they may change only during an explicit publish.
        /// </summary>
        public void DeleteStaleDraftWorksheets(
            dynamic excelApplication,
            dynamic workbook,
            string reportId,
            IReadOnlyCollection<string> activeObjectIds)
        {
            if (excelApplication == null) throw new ArgumentNullException(nameof(excelApplication));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(reportId))
            {
                throw new ArgumentException("A report identifier is required.", nameof(reportId));
            }

            if (activeObjectIds == null) throw new ArgumentNullException(nameof(activeObjectIds));
            var active = new HashSet<string>(activeObjectIds, StringComparer.Ordinal);
            var stale = new List<StaleOwnedWorksheet>();
            dynamic worksheets = workbook.Worksheets;
            int count = Convert.ToInt32(worksheets.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = worksheets.Item(index);
                ManagedObjectIdentity? identity = ReadIdentity(worksheet);
                if (identity == null ||
                    identity.Kind != ManagedObjectKind.DraftWorksheet ||
                    !string.Equals(identity.ReportId, reportId, StringComparison.Ordinal) ||
                    active.Contains(identity.ObjectId))
                {
                    continue;
                }

                stale.Add(new StaleOwnedWorksheet
                {
                    Worksheet = worksheet,
                    Identity = identity
                });
            }

            if (stale.Count == 0)
            {
                return;
            }

            bool previousAlerts = Convert.ToBoolean(
                excelApplication.DisplayAlerts,
                CultureInfo.InvariantCulture);
            try
            {
                excelApplication.DisplayAlerts = false;
                foreach (StaleOwnedWorksheet item in stale)
                {
                    ownershipGuard.DemandOwned(item.Worksheet, item.Identity);
                    item.Worksheet.Delete();
                }
            }
            finally
            {
                excelApplication.DisplayAlerts = previousAlerts;
            }
        }

        private static ManagedObjectIdentity? ReadIdentity(dynamic worksheet)
        {
            try
            {
                dynamic property = worksheet.CustomProperties.Item(ManagedObjectIdentity.MarkerName);
                string? marker = Convert.ToString(property.Value, CultureInfo.InvariantCulture);
                return ManagedObjectIdentity.TryParse(marker, out ManagedObjectIdentity? identity)
                    ? identity
                    : null;
            }
            catch (Exception)
            {
                return null;
            }
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

        internal dynamic? TryFindOwnedWorksheet(dynamic workbook, ManagedObjectIdentity identity)
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

        private sealed class StaleHiddenPivotWorksheet
        {
            public dynamic Worksheet { get; set; } = null!;

            public ManagedObjectIdentity WorksheetIdentity { get; set; } = null!;

            public ManagedObjectIdentity PivotIdentity { get; set; } = null!;
        }

        private sealed class StaleOwnedWorksheet
        {
            public dynamic Worksheet { get; set; } = null!;

            public ManagedObjectIdentity Identity { get; set; } = null!;
        }
    }
}
