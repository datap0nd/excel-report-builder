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
            var copy = AddExisting("Sheet" + (values.Count + 1));
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
    }

    public sealed class PublishWorksheet
    {
        private readonly PublishWorksheetCollection owner;

        public PublishWorksheet(PublishWorksheetCollection owner, string name)
        {
            this.owner = owner;
            Name = name;
        }

        public string Name { get; set; }

        public ManagedWorksheetServiceTests.FakeCustomProperties CustomProperties { get; } = new();

        public void Copy(object After)
        {
            owner.Copy(this);
        }

        public void Delete()
        {
            owner.Delete(this);
        }
    }
}
