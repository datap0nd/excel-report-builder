using ExcelReportBuilder.Core.Planning;
using ExcelReportBuilder.Core.Specifications;
using ExcelReportBuilder.Core.Validation;

namespace ExcelReportBuilder.Core.Tests;

public sealed class SourceFingerprintTests
{
    [Fact]
    public void Produces_a_stable_path_free_saved_setup_key()
    {
        var first = SourceFingerprint.FromHeaders(new[] { "Period", "Region", "Amount" });
        var caseOnlyChange = SourceFingerprint.FromHeaders(new[] { "period", "REGION", "amount" });
        var reordered = SourceFingerprint.FromHeaders(new[] { "Region", "Period", "Amount" });

        Assert.Equal(first.HeaderHash, caseOnlyChange.HeaderHash);
        Assert.NotEqual(first.HeaderHash, reordered.HeaderHash);
        Assert.Matches("^[0-9a-f]{64}$", first.HeaderHash);
        Assert.Equal("sha256-v1:3:" + first.HeaderHash, first.GetSavedSetupKey());
        Assert.DoesNotContain("Period", first.GetSavedSetupKey(), StringComparison.Ordinal);
        Assert.DoesNotContain("Region", first.GetSavedSetupKey(), StringComparison.Ordinal);
    }

    [Fact]
    public void Rejects_a_saved_setup_when_the_selected_source_headers_drift()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();
        spec.Source.Fingerprint = SourceFingerprint.FromHeaders(new[]
        {
            "Period", "Region", "Changed", "Amount", "Units", "Weight"
        });

        var validation = ReportSpecValidator.Validate(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Contains(validation.Issues, issue => issue.Code == "SOURCE_FINGERPRINT_HEADER_MISMATCH");
    }

    [Fact]
    public void Build_plan_exposes_the_typed_fingerprint_and_compatibility_key()
    {
        var spec = SyntheticReportFactory.CreateValidLongSpec();

        var plan = ReportBuildPlanner.Create(spec, SyntheticReportFactory.CreateLongProfile());

        Assert.Same(spec.Source.Fingerprint, plan.Source.Fingerprint);
        Assert.Equal(spec.Source.Fingerprint.GetSavedSetupKey(), plan.Source.SavedSetupCompatibilityKey);
    }
}
