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
        private const int CellTypeFormulas = -4123;
        private const int PasteFormats = -4122;
        private const int PasteValues = -4163;
        private const int SheetVeryHidden = 2;
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
            var results = PublishAll(
                excelApplication,
                workbook,
                new[]
                {
                    new ManagedPublishRequest
                    {
                        DraftWorksheet = draftWorksheet,
                        DraftIdentity = draftIdentity,
                        PublishedIdentity = publishedIdentity,
                        RollbackIdentity = rollbackIdentity,
                        FinalWorksheetName = finalWorksheetName
                    }
                },
                userConfirmed);
            return results[0];
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

            var finalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var outputIdentities = new HashSet<string>(StringComparer.Ordinal);
            var activeObjectIds = new HashSet<string>(StringComparer.Ordinal);
            var contexts = new List<PublishContext>(requests.Count);
            string? reportId = null;
            foreach (ManagedPublishRequest request in requests)
            {
                if (request == null)
                {
                    throw new ArgumentException("Publish requests cannot contain a null item.", nameof(requests));
                }

                ActiveOutputState activeOutput = DemandCanPublishCore(
                    workbook,
                    request.DraftWorksheet,
                    request.DraftIdentity,
                    request.PublishedIdentity,
                    request.RollbackIdentity,
                    request.FinalWorksheetName);
                if (reportId == null)
                {
                    reportId = request.PublishedIdentity.ReportId;
                }
                else if (!string.Equals(
                             reportId,
                             request.PublishedIdentity.ReportId,
                             StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "One publish transaction cannot combine outputs from different reports.");
                }

                var finalName = ValidateWorksheetName(request.FinalWorksheetName);
                if (!finalNames.Add(finalName))
                {
                    throw new InvalidOperationException(
                        "Two managed outputs cannot publish to the same worksheet name.");
                }

                var identityKey = request.PublishedIdentity.ReportId + "\u001f" +
                    request.PublishedIdentity.ObjectId;
                if (!outputIdentities.Add(identityKey))
                {
                    throw new InvalidOperationException(
                        "Two publish requests cannot use the same managed output identity.");
                }

                activeObjectIds.Add(request.PublishedIdentity.ObjectId);

                var rollbackName = ManagedName.Worksheet(
                    finalName + " rollback",
                    request.RollbackIdentity.ObjectId);
                contexts.Add(new PublishContext
                {
                    Request = request,
                    FinalName = finalName,
                    RollbackName = rollbackName,
                    ExistingPublished = activeOutput.PublishedWorksheet,
                    // A rollback without a current final is recovery state. It
                    // remains active and untouched until a final is replaced.
                    ExistingRollback = activeOutput.PublishedWorksheet == null
                        ? null
                        : activeOutput.RollbackWorksheet
                });
            }

            IReadOnlyList<RetirementContext> retirements = FindStalePublishedOutputs(
                workbook,
                reportId!,
                activeObjectIds);

            var previousAlerts = Convert.ToBoolean(excelApplication.DisplayAlerts);
            try
            {
                excelApplication.DisplayAlerts = false;
                for (var index = 0; index < contexts.Count; index++)
                {
                    PublishContext context = contexts[index];
                    beforePublish?.Invoke(index + 1, contexts.Count, context.FinalName);
                    PrepareStagedCopies(workbook, context);
                }

                foreach (PublishContext context in contexts)
                {
                    CommitStagedCopies(workbook, context);
                }

                foreach (RetirementContext retirement in retirements)
                {
                    CommitRetirement(workbook, retirement);
                }

                // At this point every requested final and rollback is complete.
                // The prior worksheets are now transaction-owned backups and can
                // be discarded without risking any published output.
                foreach (PublishContext context in contexts)
                {
                    DeleteObsoleteBackup(context.ExistingPublished, context.PublishedBackupIdentity);
                    DeleteObsoleteBackup(context.ExistingRollback, context.RollbackBackupIdentity);
                }
                foreach (RetirementContext retirement in retirements)
                {
                    DeleteObsoleteBackup(retirement.Worksheet, retirement.BackupIdentity);
                }

                var results = new List<PublishResult>(contexts.Count);
                foreach (PublishContext context in contexts)
                {
                    results.Add(new PublishResult
                    {
                        PublishedWorksheetName = context.FinalName,
                        RollbackWorksheetName = context.ExistingPublished == null
                            ? null
                            : context.RollbackName
                    });
                }

                return results;
            }
            catch (Exception publishFailure)
            {
                IReadOnlyList<Exception> compensationFailures = Compensate(contexts, retirements);
                if (compensationFailures.Count > 0)
                {
                    var failures = new List<Exception>(compensationFailures.Count + 1)
                    {
                        publishFailure
                    };
                    failures.AddRange(compensationFailures);
                    throw new InvalidOperationException(
                        "Publishing failed and Excel could not fully restore the previous managed outputs.",
                        new AggregateException(failures));
                }

                throw;
            }
            finally
            {
                excelApplication.DisplayAlerts = previousAlerts;
            }
        }

        public void DemandCanPublish(
            dynamic workbook,
            dynamic draftWorksheet,
            ManagedObjectIdentity draftIdentity,
            ManagedObjectIdentity publishedIdentity,
            ManagedObjectIdentity rollbackIdentity,
            string finalWorksheetName)
        {
            DemandCanPublishCore(
                workbook,
                draftWorksheet,
                draftIdentity,
                publishedIdentity,
                rollbackIdentity,
                finalWorksheetName);
        }

        private ActiveOutputState DemandCanPublishCore(
            dynamic workbook,
            dynamic draftWorksheet,
            ManagedObjectIdentity draftIdentity,
            ManagedObjectIdentity publishedIdentity,
            ManagedObjectIdentity rollbackIdentity,
            string finalWorksheetName)
        {
            ownershipGuard.DemandOwned(draftWorksheet, draftIdentity);
            if (draftIdentity.Kind != ManagedObjectKind.DraftWorksheet ||
                publishedIdentity.Kind != ManagedObjectKind.PublishedWorksheet ||
                rollbackIdentity.Kind != ManagedObjectKind.RollbackWorksheet)
            {
                throw new ArgumentException(
                    "Publishing requires matching draft, published, and rollback ownership identities.");
            }

            if (!string.Equals(draftIdentity.ReportId, publishedIdentity.ReportId, StringComparison.Ordinal) ||
                !string.Equals(draftIdentity.ReportId, rollbackIdentity.ReportId, StringComparison.Ordinal) ||
                !string.Equals(draftIdentity.ObjectId, publishedIdentity.ObjectId, StringComparison.Ordinal) ||
                !string.Equals(draftIdentity.ObjectId, rollbackIdentity.ObjectId, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "Publishing identities must describe the same managed report output.");
            }

            var safeName = ValidateWorksheetName(finalWorksheetName);
            var rollbackName = ManagedName.Worksheet(
                safeName + " rollback",
                rollbackIdentity.ObjectId);
            ActiveOutputState active = FindActiveOutputState(
                workbook,
                publishedIdentity,
                rollbackIdentity,
                safeName,
                rollbackName);

            dynamic? worksheetAtFinalName = TryGetWorksheet(workbook, safeName);
            if (worksheetAtFinalName != null)
            {
                ownershipGuard.DemandOwned(worksheetAtFinalName, publishedIdentity);
            }

            dynamic? worksheetAtRollbackName = TryGetWorksheet(workbook, rollbackName);
            if (active.PublishedWorksheet != null && worksheetAtRollbackName != null)
            {
                ownershipGuard.DemandOwned(worksheetAtRollbackName, rollbackIdentity);
            }

            return active;
        }

        private ActiveOutputState FindActiveOutputState(
            dynamic workbook,
            ManagedObjectIdentity publishedIdentity,
            ManagedObjectIdentity rollbackIdentity,
            string finalName,
            string rollbackName)
        {
            object? published = null;
            object? rollback = null;
            dynamic worksheets = workbook.Worksheets;
            var count = Convert.ToInt32(worksheets.Count);
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = worksheets.Item(index);
                if (ownershipGuard.IsOwned(worksheet, publishedIdentity))
                {
                    if (published != null)
                    {
                        throw new InvalidOperationException(
                            "More than one worksheet carries the active published-output identity.");
                    }

                    published = worksheet;
                }

                if (ownershipGuard.IsOwned(worksheet, rollbackIdentity))
                {
                    if (rollback != null)
                    {
                        throw new InvalidOperationException(
                            "More than one worksheet carries the active rollback-output identity.");
                    }

                    rollback = worksheet;
                }
            }

            DemandExpectedName(published, publishedIdentity, finalName, "published output");
            DemandExpectedName(rollback, rollbackIdentity, rollbackName, "rollback output");
            return new ActiveOutputState
            {
                PublishedWorksheet = published,
                RollbackWorksheet = rollback
            };
        }

        private void DemandExpectedName(
            object? worksheetObject,
            ManagedObjectIdentity identity,
            string expectedName,
            string label)
        {
            if (worksheetObject == null)
            {
                return;
            }

            dynamic worksheet = worksheetObject;
            ownershipGuard.DemandOwned(worksheet, identity);
            var actualName = Convert.ToString(worksheet.Name) ?? string.Empty;
            if (!string.Equals(actualName, expectedName, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The active managed " + label + " is not at its expected worksheet name.");
            }
        }

        private void PrepareStagedCopies(dynamic workbook, PublishContext context)
        {
            var token = Guid.NewGuid().ToString("N");
            context.StagedPublishedIdentity = CreateTransactionIdentity(
                context.Request.PublishedIdentity,
                "new",
                token);
            context.StagedPublished = CreateStagedCopy(
                workbook,
                context.Request.DraftWorksheet,
                context.StagedPublishedIdentity,
                "publish new");

            if (context.ExistingPublished == null)
            {
                return;
            }

            context.StagedRollbackIdentity = CreateTransactionIdentity(
                context.Request.RollbackIdentity,
                "rollback",
                token);
            context.StagedRollback = CreateStagedCopy(
                workbook,
                context.ExistingPublished,
                context.StagedRollbackIdentity,
                "publish rollback");
        }

        private dynamic CreateStagedCopy(
            dynamic workbook,
            dynamic sourceWorksheet,
            ManagedObjectIdentity stagingIdentity,
            string label)
        {
            sourceWorksheet.Copy(After: workbook.Worksheets.Item(workbook.Worksheets.Count));
            dynamic staged = workbook.ActiveSheet;
            try
            {
                ownershipGuard.MarkOwned(staged, stagingIdentity);
                staged.Name = TemporaryWorksheetName(workbook, label, stagingIdentity.ObjectId);
                FreezeWorksheetSnapshot(staged);
                return staged;
            }
            catch (Exception)
            {
                TryDeleteCreatedWorksheet(staged);
                throw;
            }
        }

        private void CommitStagedCopies(dynamic workbook, PublishContext context)
        {
            var token = Guid.NewGuid().ToString("N");
            if (context.ExistingPublished != null)
            {
                context.PublishedBackupIdentity = CreateTransactionIdentity(
                    context.Request.PublishedIdentity,
                    "old",
                    token);
                ownershipGuard.DemandOwned(
                    context.ExistingPublished,
                    context.Request.PublishedIdentity);
                ownershipGuard.MarkOwned(
                    context.ExistingPublished,
                    context.PublishedBackupIdentity);
                context.ExistingPublished.Name = TemporaryWorksheetName(
                    workbook,
                    "publish old",
                    context.PublishedBackupIdentity.ObjectId);
            }

            if (context.ExistingRollback != null)
            {
                context.RollbackBackupIdentity = CreateTransactionIdentity(
                    context.Request.RollbackIdentity,
                    "older",
                    token);
                ownershipGuard.DemandOwned(
                    context.ExistingRollback,
                    context.Request.RollbackIdentity);
                ownershipGuard.MarkOwned(
                    context.ExistingRollback,
                    context.RollbackBackupIdentity);
                context.ExistingRollback.Name = TemporaryWorksheetName(
                    workbook,
                    "publish older",
                    context.RollbackBackupIdentity.ObjectId);
            }

            if (context.StagedRollback != null)
            {
                context.StagedRollback.Name = context.RollbackName;
                ownershipGuard.MarkOwned(
                    context.StagedRollback,
                    context.Request.RollbackIdentity);
            }

            if (context.StagedPublished == null)
            {
                throw new InvalidOperationException(
                    "The publish transaction is missing its staged output worksheet.");
            }

            context.StagedPublished.Name = context.FinalName;
            ownershipGuard.MarkOwned(
                context.StagedPublished,
                context.Request.PublishedIdentity);
        }

        private IReadOnlyList<Exception> Compensate(
            IReadOnlyList<PublishContext> contexts,
            IReadOnlyList<RetirementContext> retirements)
        {
            var failures = new List<Exception>();
            for (var index = retirements.Count - 1; index >= 0; index--)
            {
                RetirementContext retirement = retirements[index];
                if (retirement.BackupIdentity != null)
                {
                    TryCompensationStep(
                        () => RestoreOriginal(
                            retirement.Worksheet,
                            retirement.OriginalName,
                            retirement.OriginalIdentity,
                            retirement.BackupIdentity),
                        failures);
                }
            }

            for (var index = contexts.Count - 1; index >= 0; index--)
            {
                PublishContext context = contexts[index];
                TryCompensationStep(
                    () => DeleteTransactionCopy(
                        context.StagedPublished,
                        context.StagedPublishedIdentity,
                        context.Request.PublishedIdentity),
                    failures);
                TryCompensationStep(
                    () => DeleteTransactionCopy(
                        context.StagedRollback,
                        context.StagedRollbackIdentity,
                        context.Request.RollbackIdentity),
                    failures);
                if (context.PublishedBackupIdentity != null)
                {
                    TryCompensationStep(
                        () => RestoreOriginal(
                            context.ExistingPublished,
                            context.FinalName,
                            context.Request.PublishedIdentity,
                            context.PublishedBackupIdentity),
                        failures);
                }

                if (context.RollbackBackupIdentity != null)
                {
                    TryCompensationStep(
                        () => RestoreOriginal(
                            context.ExistingRollback,
                            context.RollbackName,
                            context.Request.RollbackIdentity,
                            context.RollbackBackupIdentity),
                        failures);
                }
            }

            return failures;
        }

        private void CommitRetirement(dynamic workbook, RetirementContext retirement)
        {
            ownershipGuard.DemandOwned(retirement.Worksheet, retirement.OriginalIdentity);
            retirement.BackupIdentity = CreateTransactionIdentity(
                retirement.OriginalIdentity,
                "retired",
                Guid.NewGuid().ToString("N"));
            ownershipGuard.MarkOwned(retirement.Worksheet, retirement.BackupIdentity);
            retirement.Worksheet.Name = TemporaryWorksheetName(
                workbook,
                "publish retired",
                retirement.BackupIdentity.ObjectId);
        }

        private IReadOnlyList<RetirementContext> FindStalePublishedOutputs(
            dynamic workbook,
            string reportId,
            ISet<string> activeObjectIds)
        {
            var result = new List<RetirementContext>();
            dynamic worksheets = workbook.Worksheets;
            var count = Convert.ToInt32(worksheets.Count);
            for (var index = 1; index <= count; index++)
            {
                dynamic worksheet = worksheets.Item(index);
                ManagedObjectIdentity? identity = ReadWorksheetIdentity(worksheet);
                if (identity == null ||
                    !string.Equals(identity.ReportId, reportId, StringComparison.Ordinal) ||
                    (identity.Kind != ManagedObjectKind.PublishedWorksheet &&
                     identity.Kind != ManagedObjectKind.RollbackWorksheet) ||
                    activeObjectIds.Contains(identity.ObjectId))
                {
                    continue;
                }

                ownershipGuard.DemandOwned(worksheet, identity);
                result.Add(new RetirementContext
                {
                    Worksheet = worksheet,
                    OriginalName = Convert.ToString(worksheet.Name) ?? string.Empty,
                    OriginalIdentity = identity
                });
            }

            return result;
        }

        private static ManagedObjectIdentity? ReadWorksheetIdentity(dynamic worksheet)
        {
            try
            {
                dynamic property = worksheet.CustomProperties.Item(ManagedObjectIdentity.MarkerName);
                string? marker = Convert.ToString(property.Value);
                return ManagedObjectIdentity.TryParse(marker, out ManagedObjectIdentity? identity)
                    ? identity
                    : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void DeleteTransactionCopy(
            object? worksheetObject,
            ManagedObjectIdentity? stagingIdentity,
            ManagedObjectIdentity committedIdentity)
        {
            if (worksheetObject == null)
            {
                return;
            }

            dynamic worksheet = worksheetObject;
            if ((stagingIdentity == null || !ownershipGuard.IsOwned(worksheet, stagingIdentity)) &&
                !ownershipGuard.IsOwned(worksheet, committedIdentity))
            {
                throw new InvalidOperationException(
                    "A transaction worksheet lost its managed ownership marker during compensation.");
            }

            worksheet.Delete();
        }

        private void RestoreOriginal(
            object? worksheetObject,
            string originalName,
            ManagedObjectIdentity originalIdentity,
            ManagedObjectIdentity? backupIdentity)
        {
            if (worksheetObject == null)
            {
                return;
            }

            dynamic worksheet = worksheetObject;
            if (!ownershipGuard.IsOwned(worksheet, originalIdentity) &&
                (backupIdentity == null || !ownershipGuard.IsOwned(worksheet, backupIdentity)))
            {
                throw new InvalidOperationException(
                    "An original worksheet lost its managed ownership marker during compensation.");
            }

            worksheet.Name = originalName;
            ownershipGuard.MarkOwned(worksheet, originalIdentity);
        }

        private void DeleteObsoleteBackup(
            object? worksheetObject,
            ManagedObjectIdentity? backupIdentity)
        {
            if (worksheetObject == null || backupIdentity == null)
            {
                return;
            }

            dynamic worksheet = worksheetObject;
            ownershipGuard.DemandOwned(worksheet, backupIdentity);
            try
            {
                worksheet.Delete();
            }
            catch (Exception)
            {
                // The replacement is already complete. Keep an inaccessible,
                // transaction-owned recovery sheet instead of reporting failure
                // after the user's published outputs have changed successfully.
                try
                {
                    worksheet.Visible = SheetVeryHidden;
                }
                catch (Exception)
                {
                    // The worksheet retains a transaction-only ownership marker,
                    // so it cannot be mistaken for a published output or rollback.
                }
            }
        }

        private static void TryCompensationStep(Action action, ICollection<Exception> failures)
        {
            try
            {
                action();
            }
            catch (Exception exception)
            {
                failures.Add(exception);
            }
        }

        private static ManagedObjectIdentity CreateTransactionIdentity(
            ManagedObjectIdentity target,
            string role,
            string token)
        {
            return new ManagedObjectIdentity(
                target.ReportId,
                target.ObjectId + "_publish_" + role + "_" + token,
                ManagedObjectKind.Metadata);
        }

        private static string TemporaryWorksheetName(
            dynamic workbook,
            string label,
            string stableId)
        {
            var stableSuffix = stableId.Length <= 8
                ? stableId
                : stableId.Substring(stableId.Length - 8);
            for (var attempt = 0; attempt < 100; attempt++)
            {
                var candidate = ManagedName.Worksheet(
                    label + " " + attempt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    stableSuffix);
                if (TryGetWorksheet(workbook, candidate) == null)
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException("Excel could not allocate a temporary publish worksheet.");
        }

        private static void FreezeWorksheetSnapshot(dynamic worksheet)
        {
            FreezePivotTables(worksheet);
            FreezeWorksheetFormulas(worksheet);
            DemandNoPivotTables(worksheet);
        }

        private static void FreezePivotTables(dynamic worksheet)
        {
            try
            {
                var remaining = CountPivotTables(worksheet);
                for (var pivotIndex = remaining; pivotIndex >= 1; pivotIndex--)
                {
                    try
                    {
                        dynamic pivotTable = worksheet.PivotTables().Item(pivotIndex);
                        dynamic tableRange = pivotTable.TableRange2;

                        // Copy the complete displayed range before clearing the
                        // PivotTable. Range.Clear removes the native PivotTable;
                        // values-only and formats-only paste operations cannot
                        // restore its live cache or refresh behavior.
                        tableRange.Copy();
                        tableRange.Clear();
                        if (CountPivotTables(worksheet) != pivotIndex - 1)
                        {
                            throw new InvalidOperationException(
                                "Excel did not remove a managed PivotTable while preparing a static publish snapshot.");
                        }

                        tableRange.PasteSpecial(PasteFormats);
                        tableRange.PasteSpecial(PasteValues);
                    }
                    catch (Exception exception)
                    {
                        throw new InvalidOperationException(
                            "Excel could not convert every managed PivotTable to a static publish snapshot.",
                            exception);
                    }

                    if (CountPivotTables(worksheet) != pivotIndex - 1)
                    {
                        throw new InvalidOperationException(
                            "Excel recreated a PivotTable while preparing a static publish snapshot.");
                    }
                }

                DemandNoPivotTables(worksheet);
            }
            finally
            {
                TryClearCopyMode(worksheet);
            }
        }

        private static void TryClearCopyMode(dynamic worksheet)
        {
            try
            {
                worksheet.Application.CutCopyMode = false;
            }
            catch (Exception)
            {
                // Copy-mode cleanup is best effort. It must not turn a fully
                // frozen static worksheet into a failed publish transaction.
            }
        }

        private static int CountPivotTables(dynamic worksheet)
        {
            try
            {
                return Convert.ToInt32(worksheet.PivotTables().Count);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel did not expose the PivotTables collection for static publishing.",
                    exception);
            }
        }

        private static void DemandNoPivotTables(dynamic worksheet)
        {
            if (CountPivotTables(worksheet) != 0)
            {
                throw new InvalidOperationException(
                    "A static published worksheet cannot retain a PivotTable.");
            }
        }

        private static void FreezeWorksheetFormulas(dynamic worksheet)
        {
            dynamic usedRange = worksheet.UsedRange;
            object? hasFormula = usedRange.HasFormula;
            if (hasFormula is bool boolean && !boolean)
            {
                return;
            }

            dynamic formulaCells;
            try
            {
                formulaCells = usedRange.SpecialCells(CellTypeFormulas);
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "Excel reported formulas in the managed output but could not enumerate them for publishing.",
                    exception);
            }

            dynamic areas = formulaCells.Areas;
            var areaCount = Convert.ToInt32(areas.Count);
            for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                dynamic area = areas.Item(areaIndex);
                object? values = area.Value2;
                area.Value2 = values;
            }

            object? remainingFormula = usedRange.HasFormula;
            if (!(remainingFormula is bool remaining && !remaining))
            {
                throw new InvalidOperationException(
                    "Excel did not freeze every managed formula before publishing.");
            }
        }

        private static void TryDeleteCreatedWorksheet(dynamic worksheet)
        {
            try
            {
                worksheet.Delete();
            }
            catch (Exception)
            {
                // Preserve the original staging failure. The created worksheet
                // still carries a transaction-only ownership identity.
            }
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

        private sealed class PublishContext
        {
            public ManagedPublishRequest Request { get; set; } = null!;

            public string FinalName { get; set; } = string.Empty;

            public string RollbackName { get; set; } = string.Empty;

            public dynamic? ExistingPublished { get; set; }

            public dynamic? ExistingRollback { get; set; }

            public dynamic? StagedPublished { get; set; }

            public dynamic? StagedRollback { get; set; }

            public ManagedObjectIdentity? StagedPublishedIdentity { get; set; }

            public ManagedObjectIdentity? StagedRollbackIdentity { get; set; }

            public ManagedObjectIdentity? PublishedBackupIdentity { get; set; }

            public ManagedObjectIdentity? RollbackBackupIdentity { get; set; }
        }

        private sealed class RetirementContext
        {
            public dynamic Worksheet { get; set; } = null!;

            public string OriginalName { get; set; } = string.Empty;

            public ManagedObjectIdentity OriginalIdentity { get; set; } = null!;

            public ManagedObjectIdentity? BackupIdentity { get; set; }
        }

        private sealed class ActiveOutputState
        {
            public dynamic? PublishedWorksheet { get; set; }

            public dynamic? RollbackWorksheet { get; set; }
        }
    }
}
