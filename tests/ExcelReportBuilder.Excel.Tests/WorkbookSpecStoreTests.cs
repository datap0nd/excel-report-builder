using System.Xml.Linq;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Excel.Persistence;

namespace ExcelReportBuilder.Excel.Tests;

public sealed class WorkbookSpecStoreTests
{
    [Fact]
    public void Workbook_identity_is_anonymous_and_stable_for_the_open_workbook()
    {
        var workbook = new FakeWorkbook();
        var store = new WorkbookIdentityStore();

        string first = store.GetOrCreate(workbook);
        string second = new WorkbookIdentityStore().GetOrCreate(workbook);

        Assert.Equal(first, second);
        Assert.StartsWith("workbook_", first, StringComparison.Ordinal);
        Assert.True(Guid.TryParseExact(first.Substring("workbook_".Length), "N", out _));
        string persistedXml = Assert.Single(workbook.CustomXMLParts.AllXml);
        Assert.DoesNotContain("sheet", persistedXml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("path", persistedXml, StringComparison.OrdinalIgnoreCase);

        var reopenedWorkbook = new FakeWorkbook();
        reopenedWorkbook.CustomXMLParts.Add(persistedXml);
        Assert.Equal(first, new WorkbookIdentityStore().GetOrCreate(reopenedWorkbook));
    }

    [Fact]
    public void Workbook_identity_ensure_persists_the_exact_token_once_and_rejects_collisions()
    {
        var workbook = new FakeWorkbook();
        var store = new WorkbookIdentityStore();
        const string expected = "workbook_11111111111111111111111111111111";
        const string collision = "workbook_22222222222222222222222222222222";

        Assert.Equal(expected, store.Ensure(workbook, expected));
        Assert.Equal(expected, store.Ensure(workbook, expected));
        Assert.Equal(1, workbook.CustomXMLParts.TotalCount);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            store.Ensure(workbook, collision));
        Assert.Contains("different managed workbook identity", exception.Message, StringComparison.Ordinal);
        Assert.Equal(1, workbook.CustomXMLParts.TotalCount);
    }

    [Fact]
    public void Workbook_identity_ensure_rejects_non_managed_tokens_without_writing()
    {
        var workbook = new FakeWorkbook();

        Assert.Throws<ArgumentException>(() =>
            new WorkbookIdentityStore().Ensure(workbook, @"C:\secret\book.xlsx"));

        Assert.Equal(0, workbook.CustomXMLParts.TotalCount);
    }

    [Fact]
    public void Workbook_identity_rejects_unknown_owned_versions()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add(
            "<workbookIdentity xmlns=\"" + WorkbookIdentityStore.NamespaceUri +
            "\" schemaVersion=\"2.0\" id=\"workbook_0123456789abcdef0123456789abcdef\" />");

        var exception = Assert.Throws<NotSupportedException>(
            () => new WorkbookIdentityStore().GetOrCreate(workbook));

        Assert.Equal("Unknown workbook identity version.", exception.Message);
    }

    [Fact]
    public void Workbook_identity_fails_closed_when_owned_parts_cannot_be_enumerated()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.ThrowOnSelect = true;

        var exception = Assert.Throws<InvalidOperationException>(
            () => new WorkbookIdentityStore().GetOrCreate(workbook));

        Assert.Equal(
            "The managed workbook identity parts could not be enumerated.",
            exception.Message);
        Assert.Empty(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Workbook_identity_rejects_a_null_no_op_add_result()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.ReturnNullWithoutAdding = true;

        var exception = Assert.Throws<InvalidOperationException>(
            () => new WorkbookIdentityStore().GetOrCreate(workbook));

        Assert.Equal(
            "Excel did not return the persisted managed workbook identity.",
            exception.Message);
        Assert.Empty(workbook.CustomXMLParts.AllXml);
    }

    [Fact]
    public void Saved_setup_requires_matching_object_kind_identity_and_fingerprint()
    {
        SourceFingerprintSpec fingerprint = SourceFingerprint.FromHeaders(
            new[] { "Period", "Category", "Amount" });
        var saved = new WorkbookSourceSpec
        {
            Kind = WorkbookSourceKind.Table,
            WorkbookObjectName = "SourceTable",
            Fingerprint = fingerprint
        };

        Assert.True(SavedSetupCompatibility.Matches(
            saved,
            WorkbookSourceKind.Table,
            SourceFingerprint.FromHeaders(new[] { "period", "CATEGORY", "amount" }),
            workbookObjectMatches: true));
        Assert.False(SavedSetupCompatibility.Matches(
            saved,
            WorkbookSourceKind.NamedRange,
            fingerprint,
            workbookObjectMatches: true));
        Assert.False(SavedSetupCompatibility.Matches(
            saved,
            WorkbookSourceKind.Table,
            fingerprint,
            workbookObjectMatches: false));
        Assert.False(SavedSetupCompatibility.Matches(
            saved,
            WorkbookSourceKind.Table,
            SourceFingerprint.FromHeaders(new[] { "Period", "Amount", "Category" }),
            workbookObjectMatches: true));
    }

    [Fact]
    public void Store_round_trips_and_replaces_only_the_same_report()
    {
        var workbook = new FakeWorkbook();
        var store = new WorkbookSpecStore();
        var first = new ReportSpecV1 { Id = "report_1", Name = "First" };
        var second = new ReportSpecV1 { Id = "report_2", Name = "Second" };

        store.Save(workbook, first);
        store.Save(workbook, second);
        first.Name = "Updated";
        store.Save(workbook, first);

        var loaded = store.LoadAll(workbook);
        Assert.Equal(2, loaded.Count);
        Assert.Equal("Updated", loaded.Single(item => item.Id == "report_1").Name);
        Assert.Equal("Second", loaded.Single(item => item.Id == "report_2").Name);
    }

    [Fact]
    public void Store_leaves_foreign_custom_xml_untouched()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add("<foreign xmlns=\"urn:another-product\" />");
        var store = new WorkbookSpecStore();

        store.Save(workbook, new ReportSpecV1 { Id = "report", Name = "Report" });

        Assert.Equal(2, workbook.CustomXMLParts.TotalCount);
    }

    [Fact]
    public void Store_rejects_an_unknown_owned_report_specification_version()
    {
        var workbook = new FakeWorkbook();
        workbook.CustomXMLParts.Add(
            "<reportSpec xmlns=\"" + WorkbookSpecStore.NamespaceUri +
            "\" id=\"report\" schemaVersion=\"2.0\"><payload encoding=\"base64\">e30=</payload></reportSpec>");

        var exception = Assert.Throws<NotSupportedException>(
            () => new WorkbookSpecStore().LoadAll(workbook));

        Assert.Equal("Unknown report specification version.", exception.Message);
    }

    public sealed class FakeWorkbook
    {
        public FakeCustomXmlParts CustomXMLParts { get; } = new();
    }

    public sealed class FakeCustomXmlParts
    {
        private readonly List<FakeCustomXmlPart> parts = new();

        public bool ThrowOnSelect { get; set; }

        public bool ReturnNullWithoutAdding { get; set; }

        public int TotalCount => parts.Count;

        public IReadOnlyList<string> AllXml => parts.Select(part => part.XML).ToList();

        public FakeCustomXmlPart? Add(string xml)
        {
            if (ReturnNullWithoutAdding)
            {
                return null;
            }

            FakeCustomXmlPart? holder = null;
            var created = new FakeCustomXmlPart(xml, () => parts.RemoveAll(item => ReferenceEquals(item, holder)));
            holder = created;
            parts.Add(created);
            return created;
        }

        public FakeCustomXmlPartSelection SelectByNamespace(string namespaceUri)
        {
            if (ThrowOnSelect)
            {
                throw new InvalidOperationException("Injected selection failure.");
            }

            var matches = parts.Where(part =>
            {
                var document = XDocument.Parse(part.XML);
                return document.Root?.Name.NamespaceName == namespaceUri;
            }).ToList();
            return new FakeCustomXmlPartSelection(matches);
        }
    }

    public sealed class FakeCustomXmlPartSelection
    {
        private readonly IReadOnlyList<FakeCustomXmlPart> parts;

        public FakeCustomXmlPartSelection(IReadOnlyList<FakeCustomXmlPart> parts)
        {
            this.parts = parts;
        }

        public int Count => parts.Count;

        public FakeCustomXmlPart Item(int oneBasedIndex) => parts[oneBasedIndex - 1];
    }

    public sealed class FakeCustomXmlPart
    {
        private readonly Action delete;

        public FakeCustomXmlPart(string xml, Action delete)
        {
            XML = xml;
            this.delete = delete;
        }

        public string XML { get; }

        public void Delete() => delete();
    }
}
