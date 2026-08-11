using System;
using System.Collections.Generic;
using ExcelReportBuilder.Excel.Execution;
using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Publishing
{
    public sealed class PublishResult
    {
        public string PublishedWorksheetName { get; set; } = string.Empty;

        public string? RollbackWorksheetName { get; set; }
    }

    public sealed class ManagedPublishRequest
    {
        public object DraftWorksheet { get; set; } = null!;

        public ManagedObjectIdentity DraftIdentity { get; set; } = null!;

        public ManagedObjectIdentity PublishedIdentity { get; set; } = null!;

        public ManagedObjectIdentity RollbackIdentity { get; set; } = null!;

        public string FinalWorksheetName { get; set; } = string.Empty;
    }

    public sealed class ManagedPublishService
    {
        private readonly ManagedOwnershipGuard ownershipGuard;

        public ManagedPublishService(ManagedOwnershipGuard? ownershipGuard = null)
        {
            this.ownershipGuard = ownershipGuard ?? new ManagedOwnershipGuard();
        }

        public PublishResult Publish(
            dynamic excelApplication,
            dynamic workbook,
            dynamic draftWorksheet,
            ManagedObjectIdentity draftIdentity,
            ManagedObjectIdentity publishedIdentity,
            ManagedObjectIdentity rollbackIdentity,
            string finalWorksheetName,
            bool userConfirmed)
        {
            if (!userConfirmed)
            {
                throw new InvalidOperationException("Publishing requires an explicit user confirmation.");
            }

            DemandCanPublish(
                workbook,
                draftWorksheet,
                draftIdentity,
                publishedIdentity,
                rollbackIdentity,
                finalWorksheetName);
            ownershipGuard.DemandOwned(draftWorksheet, draftIdentity);
            if (publishedIdentity.Kind != ManagedObjectKind.PublishedWorksheet ||
                rollbackIdentity.Kind != ManagedObjectKind.RollbackWorksheet)
            {
                throw new ArgumentException("Publishing requires published and rollback ownership identities.");
            }

            var safeName = ValidateWorksheetName(finalWorksheetName);
            dynamic? existingPublished = TryGetWorksheet(workbook, safeName);
            string? rollbackName = null;

            var previousAlerts = Convert.ToBoolean(excelApplication.DisplayAlerts);
            try
            {
                excelApplication.DisplayAlerts = false;
                if (existingPublished != null)
                {
                    ownershipGuard.DemandOwned(existingPublished, publishedIdentity);
                    rollbackName = ReplaceRollback(workbook, existingPublished, rollbackIdentity, safeName);
                    existingPublished.Delete();
                }

                draftWorksheet.Copy(After: workbook.Worksheets.Item(workbook.Worksheets.Count));
                dynamic published = workbook.ActiveSheet;
                published.Name = safeName;
                ownershipGuard.MarkOwned(published, publishedIdentity);
            }
            finally
            {
                excelApplication.DisplayAlerts = previousAlerts;
            }

            return new PublishResult
            {
                PublishedWorksheetName = safeName,
                RollbackWorksheetName = rollbackName
            };
        }

        public IReadOnlyList<PublishResult> PublishAll(
            dynamic excelApplication,
            dynamic workbook,
            IReadOnlyList<ManagedPublishRequest> requests,
            bool userConfirmed,
            Action<int, int, string>? beforePublish = null)
        {
            if (!userConfirmed)
            {
                throw new InvalidOperationException("Publishing requires an explicit user confirmation.");
            }

            if (requests == null || requests.Count == 0)
            {
                throw new ArgumentException("At least one managed output is required for publishing.", nameof(requests));
            }

            foreach (var request in requests)
            {
                DemandCanPublish(
                    workbook,
                    request.DraftWorksheet,
                    request.DraftIdentity,
                    request.PublishedIdentity,
                    request.RollbackIdentity,
                    request.FinalWorksheetName);
            }

            var results = new List<PublishResult>(requests.Count);
            for (var index = 0; index < requests.Count; index++)
            {
                var request = requests[index];
                beforePublish?.Invoke(index + 1, requests.Count, request.FinalWorksheetName);
                results.Add(Publish(
                    excelApplication,
                    workbook,
                    request.DraftWorksheet,
                    request.DraftIdentity,
                    request.PublishedIdentity,
                    request.RollbackIdentity,
                    request.FinalWorksheetName,
                    userConfirmed: true));
            }

            return results;
        }

        public void DemandCanPublish(
            dynamic workbook,
            dynamic draftWorksheet,
            ManagedObjectIdentity draftIdentity,
            ManagedObjectIdentity publishedIdentity,
            ManagedObjectIdentity rollbackIdentity,
            string finalWorksheetName)
        {
            ownershipGuard.DemandOwned(draftWorksheet, draftIdentity);
            if (publishedIdentity.Kind != ManagedObjectKind.PublishedWorksheet ||
                rollbackIdentity.Kind != ManagedObjectKind.RollbackWorksheet)
            {
                throw new ArgumentException("Publishing requires published and rollback ownership identities.");
            }

            var safeName = ValidateWorksheetName(finalWorksheetName);
            dynamic? existingPublished = TryGetWorksheet(workbook, safeName);
            if (existingPublished == null)
            {
                return;
            }

            ownershipGuard.DemandOwned(existingPublished, publishedIdentity);
            var rollbackName = ManagedName.Worksheet(
                safeName + " rollback",
                rollbackIdentity.ObjectId);
            dynamic? existingRollback = TryGetWorksheet(workbook, rollbackName);
            if (existingRollback != null)
            {
                ownershipGuard.DemandOwned(existingRollback, rollbackIdentity);
            }
        }

        private string ReplaceRollback(
            dynamic workbook,
            dynamic published,
            ManagedObjectIdentity rollbackIdentity,
            string publishedName)
        {
            var rollbackName = ManagedName.Worksheet(publishedName + " rollback", rollbackIdentity.ObjectId);
            dynamic? existingRollback = TryGetWorksheet(workbook, rollbackName);
            if (existingRollback != null)
            {
                ownershipGuard.DemandOwned(existingRollback, rollbackIdentity);
                existingRollback.Delete();
            }

            published.Copy(After: workbook.Worksheets.Item(workbook.Worksheets.Count));
            dynamic rollback = workbook.ActiveSheet;
            rollback.Name = rollbackName;
            ownershipGuard.MarkOwned(rollback, rollbackIdentity);
            return rollbackName;
        }

        private static string ValidateWorksheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 ||
                name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            {
                throw new ArgumentException("The published worksheet name is invalid.", nameof(name));
            }

            return name.Trim();
        }

        private static dynamic? TryGetWorksheet(dynamic workbook, string name)
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
    }
}
