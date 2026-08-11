using ExcelReportBuilder.Excel.Ownership;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class OwnershipTests
{
    [Fact]
    public void Identity_round_trips_through_marker()
    {
        var expected = new ManagedObjectIdentity("report_1", "draft_1", ManagedObjectKind.DraftWorksheet);

        var parsed = ManagedObjectIdentity.TryParse(expected.MarkerValue, out var actual);

        Assert.True(parsed);
        Assert.NotNull(actual);
        Assert.Equal(expected.ReportId, actual.ReportId);
        Assert.Equal(expected.ObjectId, actual.ObjectId);
        Assert.Equal(expected.Kind, actual.Kind);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("erb:v2:report:DraftWorksheet:draft")]
    [InlineData("erb:v1:report:Unknown:draft")]
    public void Identity_rejects_unknown_markers(string? marker)
    {
        Assert.False(ManagedObjectIdentity.TryParse(marker, out _));
    }

    [Fact]
    public void Worksheet_names_are_valid_and_stable()
    {
        var name = ManagedName.Worksheet("Revenue: by / region [draft]", "1234567890abcdef");

        Assert.True(name.Length <= 31);
        Assert.DoesNotContain(':', name);
        Assert.DoesNotContain('/', name);
        Assert.DoesNotContain('[', name);
        Assert.EndsWith("12345678", name);
        Assert.Equal(name, ManagedName.Worksheet("Revenue: by / region [draft]", "1234567890abcdef"));
    }

    [Fact]
    public void Guard_requires_exact_ownership_marker()
    {
        dynamic worksheet = new FakeWorksheet();
        var guard = new ManagedOwnershipGuard();
        var identity = new ManagedObjectIdentity("report", "draft", ManagedObjectKind.DraftWorksheet);

        Assert.False(guard.IsOwned(worksheet, identity));

        guard.MarkOwned(worksheet, identity);

        Assert.True(guard.IsOwned(worksheet, identity));
        Assert.False(guard.IsOwned(
            worksheet,
            new ManagedObjectIdentity("another", "draft", ManagedObjectKind.DraftWorksheet)));
    }

    [Fact]
    public void Registry_remove_deletes_only_exact_report_kind_and_object_id_records()
    {
        var workbook = new WorkbookSpecStoreTests.FakeWorkbook();
        var registry = new WorkbookOwnershipRegistry();
        var removed = new ManagedObjectIdentity("report", "pivot", ManagedObjectKind.PivotTable);
        var otherReport = new ManagedObjectIdentity("other", "pivot", ManagedObjectKind.PivotTable);
        var otherKind = new ManagedObjectIdentity("report", "pivot", ManagedObjectKind.Metadata);
        registry.Register(workbook, removed, "RemovedPivot");
        registry.Register(workbook, otherReport, "OtherReportPivot");
        registry.Register(workbook, otherKind, "OtherKindObject");

        registry.Remove(workbook, new[] { removed });

        Assert.False(registry.IsOwned(workbook, removed, "RemovedPivot"));
        Assert.True(registry.IsOwned(workbook, otherReport, "OtherReportPivot"));
        Assert.True(registry.IsOwned(workbook, otherKind, "OtherKindObject"));
    }

    [Fact]
    public void Registry_round_trips_pivot_cache_locator_and_source_contract()
    {
        var workbook = new WorkbookSpecStoreTests.FakeWorkbook();
        var registry = new WorkbookOwnershipRegistry();
        var identity = new ManagedObjectIdentity(
            "report",
            "block_cache",
            ManagedObjectKind.PivotCache);

        registry.Register(
            workbook,
            identity,
            "ManagedCache",
            "7",
            "M|12:MANAGEDMODEL");

        ManagedObjectRecord record = Assert.Single(registry.Load(workbook));
        Assert.Equal(identity.ReportId, record.ReportId);
        Assert.Equal(identity.ObjectId, record.ObjectId);
        Assert.Equal(identity.Kind, record.Kind);
        Assert.Equal("ManagedCache", record.ExcelName);
        Assert.Equal("7", record.Locator);
        Assert.Equal("M|12:MANAGEDMODEL", record.SourceContract);
        Assert.Contains("locator=\"7\"", Assert.Single(workbook.CustomXMLParts.AllXml));
    }

    [Fact]
    public void Registry_replaces_metadata_only_for_the_same_exact_cache_identity()
    {
        var workbook = new WorkbookSpecStoreTests.FakeWorkbook();
        var registry = new WorkbookOwnershipRegistry();
        var identity = new ManagedObjectIdentity(
            "report",
            "block_cache",
            ManagedObjectKind.PivotCache);
        var other = new ManagedObjectIdentity(
            "other_report",
            "block_cache",
            ManagedObjectKind.PivotCache);
        registry.Register(workbook, identity, "ManagedCache", "1", "W|5:FIRST");
        registry.Register(workbook, other, "OtherCache", "9", "W|5:OTHER");

        registry.Register(workbook, identity, "ManagedCache", "2", "W|6:SECOND");

        IReadOnlyList<ManagedObjectRecord> records = registry.Load(workbook);
        Assert.Equal(2, records.Count);
        ManagedObjectRecord updated = records.Single(record => record.ReportId == "report");
        ManagedObjectRecord untouched = records.Single(record => record.ReportId == "other_report");
        Assert.Equal("2", updated.Locator);
        Assert.Equal("W|6:SECOND", updated.SourceContract);
        Assert.Equal("9", untouched.Locator);
        Assert.Equal("W|5:OTHER", untouched.SourceContract);
    }

    public sealed class FakeWorksheet
    {
        public FakeCustomProperties CustomProperties { get; } = new();
    }

    public sealed class FakeCustomProperties
    {
        private readonly Dictionary<string, FakeCustomProperty> values = new(StringComparer.Ordinal);

        public FakeCustomProperty Item(string name)
        {
            return values.TryGetValue(name, out var value)
                ? value
                : throw new InvalidOperationException("Property not found.");
        }

        public void Add(string name, string value)
        {
            values[name] = new FakeCustomProperty { Value = value };
        }
    }

    public sealed class FakeCustomProperty
    {
        public string Value { get; set; } = string.Empty;
    }
}
