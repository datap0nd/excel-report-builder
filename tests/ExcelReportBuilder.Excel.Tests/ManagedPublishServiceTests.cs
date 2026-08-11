using ExcelReportBuilder.Excel.Ownership;
using ExcelReportBuilder.Excel.Publishing;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class ManagedPublishServiceTests
{
    [Fact]
    public void Preflight_rejects_an_unmanaged_final_sheet_before_any_publish_mutation()
    {
        var workbook = new ManagedWorksheetServiceTests.FakeWorkbook();
        var draft = workbook.Worksheets.AddExisting("Managed draft");
        var final = workbook.Worksheets.AddExisting("Final report");
        var draftIdentity = ManagedOutputIdentity.Draft("report", "Report");
        var publishedIdentity = ManagedOutputIdentity.Published("report", "Report");
        var rollbackIdentity = ManagedOutputIdentity.Rollback("report", "Report");
        new ManagedOwnershipGuard().MarkOwned(draft, draftIdentity);

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().DemandCanPublish(
                workbook,
                draft,
                draftIdentity,
                publishedIdentity,
                rollbackIdentity,
                final.Name));

        Assert.Empty(draft.Log);
        Assert.Empty(final.Log);
    }

    [Fact]
    public void Preflight_accepts_independent_owned_outputs_and_their_single_rollbacks()
    {
        var workbook = new ManagedWorksheetServiceTests.FakeWorkbook();
        var guard = new ManagedOwnershipGuard();
        var service = new ManagedPublishService();
        foreach (var logicalName in new[] { "Report", "Appendix" })
        {
            var draftIdentity = ManagedOutputIdentity.Draft("report", logicalName);
            var publishedIdentity = ManagedOutputIdentity.Published("report", logicalName);
            var rollbackIdentity = ManagedOutputIdentity.Rollback("report", logicalName);
            var finalName = logicalName + " final";
            var draft = workbook.Worksheets.AddExisting(logicalName + " draft");
            var final = workbook.Worksheets.AddExisting(finalName);
            var rollbackName = ManagedName.Worksheet(
                finalName + " rollback",
                rollbackIdentity.ObjectId);
            var rollback = workbook.Worksheets.AddExisting(rollbackName);
            guard.MarkOwned(draft, draftIdentity);
            guard.MarkOwned(final, publishedIdentity);
            guard.MarkOwned(rollback, rollbackIdentity);

            service.DemandCanPublish(
                workbook,
                draft,
                draftIdentity,
                publishedIdentity,
                rollbackIdentity,
                finalName);
        }
    }

    [Fact]
    public void Batch_publish_replaces_every_output_and_retains_exactly_one_owned_rollback_each()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var requests = new List<ManagedPublishRequest>();
        foreach (var logicalName in new[] { "Report", "Appendix" })
        {
            var draftIdentity = ManagedOutputIdentity.Draft("report", logicalName);
            var publishedIdentity = ManagedOutputIdentity.Published("report", logicalName);
            var rollbackIdentity = ManagedOutputIdentity.Rollback("report", logicalName);
            var finalName = logicalName + " final";
            var rollbackName = ManagedName.Worksheet(
                finalName + " rollback",
                rollbackIdentity.ObjectId);
            var draft = workbook.Worksheets.AddExisting(logicalName + " draft");
            var published = workbook.Worksheets.AddExisting(finalName);
            var rollback = workbook.Worksheets.AddExisting(rollbackName);
            guard.MarkOwned(draft, draftIdentity);
            guard.MarkOwned(published, publishedIdentity);
            guard.MarkOwned(rollback, rollbackIdentity);
            requests.Add(new ManagedPublishRequest
            {
                DraftWorksheet = draft,
                DraftIdentity = draftIdentity,
                PublishedIdentity = publishedIdentity,
                RollbackIdentity = rollbackIdentity,
                FinalWorksheetName = finalName
            });
        }

        var progress = new List<string>();
        var results = new ManagedPublishService().PublishAll(
            application,
            workbook,
            requests,
            userConfirmed: true,
            beforePublish: (current, total, name) =>
                progress.Add(current + "/" + total + ":" + name));

        Assert.Equal(2, results.Count);
        Assert.Equal(new[] { "1/2:Report final", "2/2:Appendix final" }, progress);
        Assert.True(application.DisplayAlerts);
        foreach (var request in requests)
        {
            var published = workbook.Worksheets.Item(request.FinalWorksheetName);
            Assert.True(guard.IsOwned(published, request.PublishedIdentity));
            var rollbackName = ManagedName.Worksheet(
                request.FinalWorksheetName + " rollback",
                request.RollbackIdentity.ObjectId);
            var rollback = workbook.Worksheets.Item(rollbackName);
            Assert.True(guard.IsOwned(rollback, request.RollbackIdentity));
            Assert.Equal(1, workbook.Worksheets.CountOwned(guard, request.RollbackIdentity));
        }
    }

    [Fact]
    public void Publish_freezes_formulas_and_pivots_in_the_final_and_rollback_snapshots()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var draftIdentity = ManagedOutputIdentity.Draft("report", "Report");
        var publishedIdentity = ManagedOutputIdentity.Published("report", "Report");
        var rollbackIdentity = ManagedOutputIdentity.Rollback("report", "Report");
        var finalName = "Report final";
        var rollbackName = ManagedName.Worksheet(
            finalName + " rollback",
            rollbackIdentity.ObjectId);
        var draft = workbook.Worksheets.AddExisting("Report draft");
        draft.HasFormula = true;
        draft.CellValue = 42m;
        draft.AddPivotTable(420m, "draft-pivot-format");
        var published = workbook.Worksheets.AddExisting(finalName);
        published.HasFormula = true;
        published.CellValue = 17m;
        published.AddPivotTable(170m, "published-pivot-format");
        var oldRollback = workbook.Worksheets.AddExisting(rollbackName);
        guard.MarkOwned(draft, draftIdentity);
        guard.MarkOwned(published, publishedIdentity);
        guard.MarkOwned(oldRollback, rollbackIdentity);

        PublishResult result = new ManagedPublishService().Publish(
            application,
            workbook,
            draft,
            draftIdentity,
            publishedIdentity,
            rollbackIdentity,
            finalName,
            userConfirmed: true);

        var final = workbook.Worksheets.Item(finalName);
        var rollback = workbook.Worksheets.Item(rollbackName);
        Assert.False(final.HasFormula);
        Assert.Equal(42m, final.CellValue);
        Assert.Equal(0, final.PivotTables().Count);
        Assert.Equal(new object?[] { 420m }, final.PivotValues);
        Assert.Equal(new[] { "draft-pivot-format" }, final.PivotFormats);
        Assert.False(final.Application.CutCopyMode);
        Assert.Equal(new[] { true, false }, final.Application.Assignments);
        Assert.False(rollback.HasFormula);
        Assert.Equal(17m, rollback.CellValue);
        Assert.Equal(0, rollback.PivotTables().Count);
        Assert.Equal(new object?[] { 170m }, rollback.PivotValues);
        Assert.Equal(new[] { "published-pivot-format" }, rollback.PivotFormats);
        Assert.False(rollback.Application.CutCopyMode);
        Assert.Equal(new[] { true, false }, rollback.Application.Assignments);
        Assert.True(guard.IsOwned(final, publishedIdentity));
        Assert.True(guard.IsOwned(rollback, rollbackIdentity));
        Assert.Equal(rollbackName, result.RollbackWorksheetName);
    }

    [Fact]
    public void Publish_fails_closed_and_removes_the_stage_when_a_pivot_cannot_be_removed()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var draftIdentity = ManagedOutputIdentity.Draft("report", "Report");
        var publishedIdentity = ManagedOutputIdentity.Published("report", "Report");
        var rollbackIdentity = ManagedOutputIdentity.Rollback("report", "Report");
        var draft = workbook.Worksheets.AddExisting("Report draft");
        draft.AddPivotTable(42m, "pivot-format", retainOnClear: true);
        guard.MarkOwned(draft, draftIdentity);

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().Publish(
                application,
                workbook,
                draft,
                draftIdentity,
                publishedIdentity,
                rollbackIdentity,
                "Report final",
                userConfirmed: true));

        Assert.True(application.DisplayAlerts);
        Assert.Single(workbook.Worksheets.All);
        Assert.Same(draft, workbook.Worksheets.Item("Report draft"));
        Assert.Equal(1, draft.PivotTables().Count);
        Assert.Equal(0, workbook.Worksheets.CountOwned(guard, publishedIdentity));
    }

    [Fact]
    public void Batch_publish_removes_staged_copies_when_preparation_fails()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var requests = CreateFirstPublishRequests(workbook, guard);
        workbook.Worksheets.FailCopyOnAttempt = 2;

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().PublishAll(
                application,
                workbook,
                requests,
                userConfirmed: true));

        Assert.True(application.DisplayAlerts);
        Assert.Equal(2, workbook.Worksheets.Count);
        foreach (ManagedPublishRequest request in requests)
        {
            Assert.Same(
                request.DraftWorksheet,
                workbook.Worksheets.Item(((PublishWorksheet)request.DraftWorksheet).Name));
            Assert.Equal(0, workbook.Worksheets.CountOwned(guard, request.PublishedIdentity));
        }
    }

    [Fact]
    public void Batch_publish_restores_every_original_when_a_later_commit_fails()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var requests = new List<ManagedPublishRequest>();
        var originals = new List<(PublishWorksheet Published, PublishWorksheet Rollback)>();
        foreach (var logicalName in new[] { "Report", "Appendix" })
        {
            var draftIdentity = ManagedOutputIdentity.Draft("report", logicalName);
            var publishedIdentity = ManagedOutputIdentity.Published("report", logicalName);
            var rollbackIdentity = ManagedOutputIdentity.Rollback("report", logicalName);
            var finalName = logicalName + " final";
            var rollbackName = ManagedName.Worksheet(
                finalName + " rollback",
                rollbackIdentity.ObjectId);
            var draft = workbook.Worksheets.AddExisting(logicalName + " draft");
            var published = workbook.Worksheets.AddExisting(finalName);
            var rollback = workbook.Worksheets.AddExisting(rollbackName);
            guard.MarkOwned(draft, draftIdentity);
            guard.MarkOwned(published, publishedIdentity);
            guard.MarkOwned(rollback, rollbackIdentity);
            originals.Add((published, rollback));
            requests.Add(new ManagedPublishRequest
            {
                DraftWorksheet = draft,
                DraftIdentity = draftIdentity,
                PublishedIdentity = publishedIdentity,
                RollbackIdentity = rollbackIdentity,
                FinalWorksheetName = finalName
            });
        }

        workbook.Worksheets.FailNextRenameTo = "Appendix final";

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().PublishAll(
                application,
                workbook,
                requests,
                userConfirmed: true));

        Assert.True(application.DisplayAlerts);
        Assert.Equal(6, workbook.Worksheets.Count);
        for (var index = 0; index < requests.Count; index++)
        {
            ManagedPublishRequest request = requests[index];
            var rollbackName = ManagedName.Worksheet(
                request.FinalWorksheetName + " rollback",
                request.RollbackIdentity.ObjectId);
            Assert.Same(originals[index].Published, workbook.Worksheets.Item(request.FinalWorksheetName));
            Assert.Same(originals[index].Rollback, workbook.Worksheets.Item(rollbackName));
            Assert.True(guard.IsOwned(originals[index].Published, request.PublishedIdentity));
            Assert.True(guard.IsOwned(originals[index].Rollback, request.RollbackIdentity));
        }
    }

    [Fact]
    public void Batch_publish_retires_only_stale_same_report_final_outputs()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        var requests = CreateFirstPublishRequests(workbook, guard);
        var stalePublishedIdentity = ManagedOutputIdentity.Published("report", "Removed");
        var staleRollbackIdentity = ManagedOutputIdentity.Rollback("report", "Removed");
        var stalePublished = workbook.Worksheets.AddExisting("Removed final");
        var staleRollback = workbook.Worksheets.AddExisting("Removed rollback");
        guard.MarkOwned(stalePublished, stalePublishedIdentity);
        guard.MarkOwned(staleRollback, staleRollbackIdentity);
        var otherIdentity = ManagedOutputIdentity.Published("other-report", "Other");
        var otherReport = workbook.Worksheets.AddExisting("Other final");
        guard.MarkOwned(otherReport, otherIdentity);
        var unmanaged = workbook.Worksheets.AddExisting("Notes");

        new ManagedPublishService().PublishAll(
            application,
            workbook,
            requests,
            userConfirmed: true);

        Assert.DoesNotContain(stalePublished, workbook.Worksheets.All);
        Assert.DoesNotContain(staleRollback, workbook.Worksheets.All);
        Assert.Contains(otherReport, workbook.Worksheets.All);
        Assert.Contains(unmanaged, workbook.Worksheets.All);
        Assert.True(guard.IsOwned(otherReport, otherIdentity));
        foreach (ManagedPublishRequest request in requests)
        {
            Assert.True(guard.IsOwned(
                workbook.Worksheets.Item(request.FinalWorksheetName),
                request.PublishedIdentity));
        }
    }

    [Fact]
    public void Batch_publish_restores_retired_outputs_when_retirement_commit_fails()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        List<ManagedPublishRequest> requests = CreateFirstPublishRequests(workbook, guard);
        // Use one active output so the rename sequence is deterministic:
        // stage copy, active final, first retirement, second retirement.
        requests.RemoveAt(1);
        var stalePublishedIdentity = ManagedOutputIdentity.Published("report", "Removed");
        var staleRollbackIdentity = ManagedOutputIdentity.Rollback("report", "Removed");
        var stalePublished = workbook.Worksheets.AddExisting("Removed final");
        var staleRollback = workbook.Worksheets.AddExisting("Removed rollback");
        guard.MarkOwned(stalePublished, stalePublishedIdentity);
        guard.MarkOwned(staleRollback, staleRollbackIdentity);
        workbook.Worksheets.FailRenameOnAttempt = 4;

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().PublishAll(
                application,
                workbook,
                requests,
                userConfirmed: true));

        Assert.True(application.DisplayAlerts);
        Assert.Same(stalePublished, workbook.Worksheets.Item("Removed final"));
        Assert.Same(staleRollback, workbook.Worksheets.Item("Removed rollback"));
        Assert.True(guard.IsOwned(stalePublished, stalePublishedIdentity));
        Assert.True(guard.IsOwned(staleRollback, staleRollbackIdentity));
        Assert.Equal(0, workbook.Worksheets.CountOwned(guard, requests[0].PublishedIdentity));
    }

    [Fact]
    public void Preflight_rejects_duplicate_active_published_ownership_markers()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        List<ManagedPublishRequest> requests = CreateFirstPublishRequests(workbook, guard);
        requests.RemoveAt(1);
        ManagedPublishRequest request = requests[0];
        var expected = workbook.Worksheets.AddExisting(request.FinalWorksheetName);
        var duplicate = workbook.Worksheets.AddExisting("Copied final");
        guard.MarkOwned(expected, request.PublishedIdentity);
        guard.MarkOwned(duplicate, request.PublishedIdentity);
        var namesBefore = workbook.Worksheets.All.Select(sheet => sheet.Name).ToArray();

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().PublishAll(
                application,
                workbook,
                requests,
                userConfirmed: true));

        Assert.Equal(namesBefore, workbook.Worksheets.All.Select(sheet => sheet.Name));
        Assert.Equal(0, workbook.Worksheets.CopyAttempts);
        Assert.True(application.DisplayAlerts);
    }

    [Fact]
    public void Preflight_rejects_an_active_rollback_at_the_wrong_name()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        List<ManagedPublishRequest> requests = CreateFirstPublishRequests(workbook, guard);
        requests.RemoveAt(1);
        ManagedPublishRequest request = requests[0];
        var rollback = workbook.Worksheets.AddExisting("Moved rollback");
        guard.MarkOwned(rollback, request.RollbackIdentity);

        Assert.Throws<InvalidOperationException>(() =>
            new ManagedPublishService().PublishAll(
                application,
                workbook,
                requests,
                userConfirmed: true));

        Assert.Same(rollback, workbook.Worksheets.Item("Moved rollback"));
        Assert.True(guard.IsOwned(rollback, request.RollbackIdentity));
        Assert.Equal(0, workbook.Worksheets.CopyAttempts);
    }

    [Fact]
    public void First_publish_preserves_a_valid_orphan_rollback_as_recovery()
    {
        var application = new FakeApplication();
        var workbook = new PublishWorkbook();
        var guard = new ManagedOwnershipGuard();
        List<ManagedPublishRequest> requests = CreateFirstPublishRequests(workbook, guard);
        requests.RemoveAt(1);
        ManagedPublishRequest request = requests[0];
        var rollbackName = ManagedName.Worksheet(
            request.FinalWorksheetName + " rollback",
            request.RollbackIdentity.ObjectId);
        var orphanRollback = workbook.Worksheets.AddExisting(rollbackName);
        guard.MarkOwned(orphanRollback, request.RollbackIdentity);

        IReadOnlyList<PublishResult> results = new ManagedPublishService().PublishAll(
            application,
            workbook,
            requests,
            userConfirmed: true);

        Assert.Same(orphanRollback, workbook.Worksheets.Item(rollbackName));
        Assert.True(guard.IsOwned(orphanRollback, request.RollbackIdentity));
        Assert.True(guard.IsOwned(
            workbook.Worksheets.Item(request.FinalWorksheetName),
            request.PublishedIdentity));
        Assert.Null(results[0].RollbackWorksheetName);
    }

    private static List<ManagedPublishRequest> CreateFirstPublishRequests(
        PublishWorkbook workbook,
        ManagedOwnershipGuard guard)
    {
        var requests = new List<ManagedPublishRequest>();
        foreach (var logicalName in new[] { "Report", "Appendix" })
        {
            var draftIdentity = ManagedOutputIdentity.Draft("report", logicalName);
            var draft = workbook.Worksheets.AddExisting(logicalName + " draft");
            guard.MarkOwned(draft, draftIdentity);
            requests.Add(new ManagedPublishRequest
            {
                DraftWorksheet = draft,
                DraftIdentity = draftIdentity,
                PublishedIdentity = ManagedOutputIdentity.Published("report", logicalName),
                RollbackIdentity = ManagedOutputIdentity.Rollback("report", logicalName),
                FinalWorksheetName = logicalName + " final"
            });
        }

        return requests;
    }

    public sealed class FakeApplication
    {
        public bool DisplayAlerts { get; set; } = true;
    }

    public sealed class PublishWorkbook
    {
        public PublishWorkbook()
        {
            Worksheets = new PublishWorksheetCollection(this);
        }

        public PublishWorksheetCollection Worksheets { get; }

        public PublishWorksheet ActiveSheet { get; set; } = null!;
    }

    public sealed class PublishWorksheetCollection
    {
        private readonly PublishWorkbook workbook;
        private readonly List<PublishWorksheet> values = new();

        public PublishWorksheetCollection(PublishWorkbook workbook)
        {
            this.workbook = workbook;
        }

        public int Count => values.Count;

        public IReadOnlyList<PublishWorksheet> All => values;

        public int? FailCopyOnAttempt { get; set; }

        public string? FailNextRenameTo { get; set; }

        public int? FailRenameOnAttempt { get; set; }

        public int CopyAttempts { get; private set; }

        private int RenameAttempts { get; set; }

        public PublishWorksheet AddExisting(string name)
        {
            var worksheet = new PublishWorksheet(this, name);
            values.Add(worksheet);
            return worksheet;
        }

        public PublishWorksheet Item(int index)
        {
            return values[index - 1];
        }

        public PublishWorksheet Item(string name)
        {
            return values.Single(value => string.Equals(
                value.Name,
                name,
                StringComparison.OrdinalIgnoreCase));
        }

        public int CountOwned(ManagedOwnershipGuard guard, ManagedObjectIdentity identity)
        {
            return values.Count(value => guard.IsOwned(value, identity));
        }

        public void Copy(PublishWorksheet source)
        {
            CopyAttempts++;
            if (FailCopyOnAttempt == CopyAttempts)
            {
                throw new InvalidOperationException("Injected worksheet copy failure.");
            }

            var copy = AddExisting("Sheet" + (values.Count + 1));
            copy.HasFormula = source.HasFormula;
            copy.CellValue = source.CellValue;
            source.CopyPivotStateTo(copy);
            try
            {
                var marker = source.CustomProperties.Item(ManagedObjectIdentity.MarkerName);
                copy.CustomProperties.Add(ManagedObjectIdentity.MarkerName, marker.Value);
            }
            catch (InvalidOperationException)
            {
            }

            workbook.ActiveSheet = copy;
        }

        public void Delete(PublishWorksheet worksheet)
        {
            values.Remove(worksheet);
        }

        public void Rename(PublishWorksheet worksheet, string name)
        {
            RenameAttempts++;
            if (FailRenameOnAttempt == RenameAttempts)
            {
                FailRenameOnAttempt = null;
                throw new InvalidOperationException("Injected worksheet rename failure.");
            }

            if (string.Equals(FailNextRenameTo, name, StringComparison.OrdinalIgnoreCase))
            {
                FailNextRenameTo = null;
                throw new InvalidOperationException("Injected worksheet rename failure.");
            }

            if (values.Any(value => !ReferenceEquals(value, worksheet) && string.Equals(
                    value.Name,
                    name,
                    StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException("A worksheet with that name already exists.");
            }

            worksheet.SetName(name);
        }
    }

    public sealed class PublishWorksheet
    {
        private readonly PublishWorksheetCollection owner;
        private readonly List<PublishPivotSurface> pivotSurfaces = new();
        private readonly PublishPivotTableCollection pivotTables;
        private string name;

        public PublishWorksheet(PublishWorksheetCollection owner, string name)
        {
            this.owner = owner;
            this.name = name;
            UsedRange = new PublishUsedRange(this);
            pivotTables = new PublishPivotTableCollection(Application);
        }

        public string Name
        {
            get => name;
            set => owner.Rename(this, value);
        }

        public bool HasFormula { get; set; }

        public object? CellValue { get; set; }

        public int Visible { get; set; } = -1;

        public PublishUsedRange UsedRange { get; }

        public PublishCopyModeApplication Application { get; } = new();

        public IReadOnlyList<object?> PivotValues =>
            pivotSurfaces.Select(surface => surface.Value).ToList();

        public IReadOnlyList<string> PivotFormats =>
            pivotSurfaces.Select(surface => surface.Format).ToList();

        public ManagedWorksheetServiceTests.FakeCustomProperties CustomProperties { get; } = new();

        public void Copy(object After)
        {
            owner.Copy(this);
        }

        public void Delete()
        {
            owner.Delete(this);
        }

        public void SetName(string value)
        {
            name = value;
        }

        public PublishPivotTableCollection PivotTables()
        {
            return pivotTables;
        }

        public void AddPivotTable(
            object? value,
            string format,
            bool retainOnClear = false)
        {
            var surface = new PublishPivotSurface(value, format, retainOnClear);
            pivotSurfaces.Add(surface);
            pivotTables.Add(surface);
        }

        public void CopyPivotStateTo(PublishWorksheet target)
        {
            foreach (PublishPivotSurface surface in pivotSurfaces)
            {
                var copy = new PublishPivotSurface(
                    surface.Value,
                    surface.Format,
                    surface.RetainOnClear)
                {
                    IsNativePivot = surface.IsNativePivot
                };
                target.pivotSurfaces.Add(copy);
                if (copy.IsNativePivot)
                {
                    target.pivotTables.Add(copy);
                }
            }
        }
    }

    public sealed class PublishPivotTableCollection
    {
        private readonly List<PublishPivotTable> values = new();
        private readonly PublishCopyModeApplication application;

        public PublishPivotTableCollection(PublishCopyModeApplication application)
        {
            this.application = application;
        }

        public int Count => values.Count;

        public PublishPivotTable Item(int index)
        {
            return values[index - 1];
        }

        public void Add(PublishPivotSurface surface)
        {
            PublishPivotTable? pivot = null;
            pivot = new PublishPivotTable(
                surface,
                application,
                () => values.Remove(pivot!));
            values.Add(pivot);
        }
    }

    public sealed class PublishPivotTable
    {
        public PublishPivotTable(
            PublishPivotSurface surface,
            PublishCopyModeApplication application,
            Action remove)
        {
            TableRange2 = new PublishPivotRange(surface, application, remove);
        }

        public PublishPivotRange TableRange2 { get; }
    }

    public sealed class PublishPivotRange
    {
        private readonly PublishPivotSurface surface;
        private readonly PublishCopyModeApplication application;
        private readonly Action remove;
        private object? copiedValue;
        private string? copiedFormat;
        private bool copied;

        public PublishPivotRange(
            PublishPivotSurface surface,
            PublishCopyModeApplication application,
            Action remove)
        {
            this.surface = surface;
            this.application = application;
            this.remove = remove;
        }

        public void Copy()
        {
            copiedValue = surface.Value;
            copiedFormat = surface.Format;
            copied = true;
            application.CutCopyMode = true;
        }

        public void Clear()
        {
            surface.Value = null;
            surface.Format = string.Empty;
            if (!surface.RetainOnClear)
            {
                surface.IsNativePivot = false;
                remove();
            }
        }

        public void PasteSpecial(int Paste)
        {
            if (!copied)
            {
                throw new InvalidOperationException("No PivotTable range was copied.");
            }

            if (Paste == -4122)
            {
                surface.Format = copiedFormat ?? string.Empty;
                return;
            }

            if (Paste == -4163)
            {
                surface.Value = copiedValue;
                return;
            }

            throw new InvalidOperationException("Unsupported synthetic paste type.");
        }
    }

    public sealed class PublishPivotSurface
    {
        public PublishPivotSurface(object? value, string format, bool retainOnClear)
        {
            Value = value;
            Format = format;
            RetainOnClear = retainOnClear;
        }

        public object? Value { get; set; }

        public string Format { get; set; }

        public bool RetainOnClear { get; }

        public bool IsNativePivot { get; set; } = true;
    }

    public sealed class PublishCopyModeApplication
    {
        private bool cutCopyMode;

        public IReadOnlyList<bool> Assignments => assignments;

        private readonly List<bool> assignments = new();

        public bool CutCopyMode
        {
            get => cutCopyMode;
            set
            {
                cutCopyMode = value;
                assignments.Add(value);
            }
        }
    }

    public sealed class PublishUsedRange
    {
        private readonly PublishWorksheet worksheet;

        public PublishUsedRange(PublishWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public bool HasFormula => worksheet.HasFormula;

        public PublishFormulaRange SpecialCells(int cellType)
        {
            if (cellType != -4123 || !worksheet.HasFormula)
            {
                throw new InvalidOperationException("No formula cells were found.");
            }

            return new PublishFormulaRange(worksheet);
        }
    }

    public sealed class PublishFormulaRange
    {
        private readonly PublishWorksheet worksheet;

        public PublishFormulaRange(PublishWorksheet worksheet)
        {
            this.worksheet = worksheet;
            Areas = new PublishFormulaAreas(this);
        }

        public PublishFormulaAreas Areas { get; }

        public object? Value2
        {
            get => worksheet.CellValue;
            set
            {
                worksheet.CellValue = value;
                worksheet.HasFormula = false;
            }
        }
    }

    public sealed class PublishFormulaAreas
    {
        private readonly PublishFormulaRange range;

        public PublishFormulaAreas(PublishFormulaRange range)
        {
            this.range = range;
        }

        public int Count => 1;

        public PublishFormulaRange Item(int index)
        {
            if (index != 1)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return range;
        }
    }
}
