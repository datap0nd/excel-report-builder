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
