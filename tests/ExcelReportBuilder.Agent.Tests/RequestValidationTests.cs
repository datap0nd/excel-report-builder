using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class RequestValidationTests
{
    [Fact]
    public void Validate_AcceptsBoundedSyntheticRequest()
    {
        AgentRequestValidator.Validate(SyntheticJob.Create());
    }

    [Fact]
    public void Validate_RejectsOversizedPromptInsteadOfTruncating()
    {
        var request = SyntheticJob.Create();
        request.UserPrompt = new string('x', AgentRequestLimits.MaximumPromptCharacters + 1);

        var error = Assert.Throws<AgentInputValidationException>(() => AgentRequestValidator.Validate(request));

        Assert.Equal("prompt_too_large", error.Code);
    }

    [Fact]
    public void Validate_RejectsRepairCyclesAboveHardLimit()
    {
        var request = SyntheticJob.Create();
        request.MaxRepairCycles = 99;

        var error = Assert.Throws<AgentInputValidationException>(() => AgentRequestValidator.Validate(request));

        Assert.Equal("repair_cycle_limit_invalid", error.Code);
    }

    [Fact]
    public void Validate_RequiresHostGeneratedWorkbookId()
    {
        var request = SyntheticJob.Create();
        request.WorkbookId = string.Empty;

        var error = Assert.Throws<AgentInputValidationException>(() => AgentRequestValidator.Validate(request));

        Assert.Equal("workbook_id_invalid", error.Code);
    }

    [Fact]
    public void Validate_RejectsUnknownSampleField()
    {
        var request = SyntheticJob.Create();
        request.Data.SampleRows[0].Values[0].Field = "Not a source column";

        var error = Assert.Throws<AgentInputValidationException>(() => AgentRequestValidator.Validate(request));

        Assert.Equal("sample_field_invalid", error.Code);
    }

    [Fact]
    public void Validate_AllowsMissingReportingYearWithoutInferringOne()
    {
        var request = SyntheticJob.Create();
        request.Data.ReportingYear = null;

        AgentRequestValidator.Validate(request);

        Assert.Null(request.Data.ReportingYear);
    }

    [Fact]
    public void Validate_RejectsRemoteEndpointWithoutWorkbookDataConsent()
    {
        var request = SyntheticJob.Create();
        request.Endpoint.BaseUrl = "https://models.example.test/v1";

        var error = Assert.Throws<AgentInputValidationException>(() => AgentRequestValidator.Validate(request));

        Assert.Equal("remote_workbook_data_consent_required", error.Code);
    }

    [Fact]
    public void Validate_AcceptsRemoteEndpointAfterWorkbookDataConsent()
    {
        var request = SyntheticJob.Create();
        request.Endpoint.BaseUrl = "https://models.example.test/v1";
        request.Endpoint.AllowRemoteWorkbookData = true;

        AgentRequestValidator.Validate(request);
    }

    [Fact]
    public void Validate_AcceptsHostValidatedCanonicalCurrentSetup()
    {
        var request = SyntheticJob.Create();
        request.CurrentSpecification.CanonicalReportSpecJson =
            "{\"schemaVersion\":\"1.0\",\"id\":\"managed-report\"}";

        AgentRequestValidator.Validate(request);
    }

    [Fact]
    public void Validate_RejectsUnsupportedCanonicalCurrentSetupVersion()
    {
        var request = SyntheticJob.Create();
        request.CurrentSpecification.CanonicalReportSpecJson =
            "{\"schemaVersion\":\"2.0\"}";

        var error = Assert.Throws<AgentInputValidationException>(() =>
            AgentRequestValidator.Validate(request));

        Assert.Equal("specification_invalid", error.Code);
    }
}
