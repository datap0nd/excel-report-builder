using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Tools;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class ToolCallValidatorTests
{
    [Fact]
    public void Validate_AcceptsApprovedWorkflowInSafeOrder()
    {
        var calls = ApprovedWorkflow();

        var result = AgentToolCallValidator.Validate(
            calls,
            SyntheticJob.Create().Data,
            requireCompleteWorkflow: true);

        Assert.True(result.IsValid);
        Assert.Equal(5, result.ToolCalls.Count);
    }

    [Fact]
    public void Validate_CompleteJobCannotSkipDeterministicBuildAndChecks()
    {
        var result = AgentToolCallValidator.Validate(
            new[] { SyntheticJob.ValidReportSpecCall() },
            SyntheticJob.Create().Data,
            requireCompleteWorkflow: true);

        Assert.False(result.IsValid);
        Assert.Equal("workflow_incomplete", result.ErrorCode);
    }

    [Theory]
    [InlineData("run_shell")]
    [InlineData("read_file")]
    [InlineData("write_formula")]
    [InlineData("call_com")]
    [InlineData("save_workbook")]
    [InlineData("publish_report")]
    [InlineData("delete_sheet")]
    public void Validate_RejectsEveryNonAllowlistedCapability(string toolName)
    {
        var call = new AgentToolCall
        {
            Id = "unsafe-call",
            Name = toolName,
            ArgumentsJson = "{}",
        };

        var result = AgentToolCallValidator.Validate(new[] { call }, SyntheticJob.Create().Data);

        Assert.False(result.IsValid);
        Assert.Equal("tool_not_allowed", result.ErrorCode);
    }

    [Fact]
    public void Validate_RejectsExtraFormulaProperty()
    {
        var call = SyntheticJob.ValidReportSpecCall();
        call.ArgumentsJson = call.ArgumentsJson.TrimEnd('}') + ",\"formula\":\"=SUM(A:A)\"}";

        var result = AgentToolCallValidator.Validate(new[] { call }, SyntheticJob.Create().Data);

        Assert.False(result.IsValid);
        Assert.Equal("tool_arguments_invalid", result.ErrorCode);
    }

    [Fact]
    public void Validate_RejectsInventedSourceColumn()
    {
        var call = SyntheticJob.ValidReportSpecCall();
        call.ArgumentsJson = call.ArgumentsJson.Replace("Department", "Invented Column", StringComparison.Ordinal);

        Assert.False(AgentToolCallValidator.Validate(new[] { call }, SyntheticJob.Create().Data).IsValid);
    }

    [Fact]
    public void Validate_RejectsOutOfOrderStatefulFlow()
    {
        var calls = ApprovedWorkflow();
        (calls[0], calls[1]) = (calls[1], calls[0]);

        var result = AgentToolCallValidator.Validate(calls, SyntheticJob.Create().Data);

        Assert.False(result.IsValid);
        Assert.Equal("tool_order_invalid", result.ErrorCode);
    }

    [Fact]
    public void Validate_DoesNotInventMissingReportingYear()
    {
        var job = SyntheticJob.Create();
        job.Data.ReportingYear = null;
        var mapping = new AgentToolCall
        {
            Id = "mapping-1",
            Name = AgentToolNames.ProposePeriodMapping,
            ArgumentsJson =
                "{\"mode\":\"widePeriods\",\"periodField\":\"\",\"reportingYear\":2026," +
                "\"mappings\":[{\"sourceField\":\"Period\",\"periodLabel\":\"2026-01\"}]}",
        };

        Assert.False(AgentToolCallValidator.Validate(new[] { mapping }, job.Data).IsValid);
    }

    [Fact]
    public void Validate_DateColumnDoesNotRequireAnInventedReportingYear()
    {
        var job = SyntheticJob.Create();
        job.Data.ReportingYear = null;
        var mapping = new AgentToolCall
        {
            Id = "mapping-date",
            Name = AgentToolNames.ProposePeriodMapping,
            ArgumentsJson =
                "{\"mode\":\"dateColumn\",\"periodField\":\"Period\"," +
                "\"reportingYear\":null,\"mappings\":[]}"
        };

        Assert.True(AgentToolCallValidator.Validate(new[] { mapping }, job.Data).IsValid);
    }

    [Fact]
    public void Validate_AllowsReportSpecToReferenceBoundedTransformOutput()
    {
        var transforms = new AgentToolCall
        {
            Id = "transforms-1",
            Name = AgentToolNames.ProposeTransforms,
            ArgumentsJson =
                "{\"transforms\":[{\"kind\":\"convertNumber\",\"sourceField\":\"Net Amount\"," +
                "\"outputField\":\"Normalized Amount\"}]}",
        };
        var spec = SyntheticJob.ValidReportSpecCall("spec-1");
        spec.ArgumentsJson = spec.ArgumentsJson.Replace("Net Amount", "Normalized Amount", StringComparison.Ordinal);

        Assert.True(AgentToolCallValidator.Validate(new[] { transforms, spec }, SyntheticJob.Create().Data).IsValid);
    }

    [Fact]
    public void Validate_AcceptsTypedMeasuresPeriodSlicesAndMultipleManagedBlocks()
    {
        var result = AgentToolCallValidator.Validate(
            new[] { SyntheticJob.ValidAdvancedReportSpecCall() },
            SyntheticJob.Create().Data);

        Assert.True(result.IsValid);
    }

    [Fact]
    public void Validate_AdvancedReportRejectsFormulaBearingNumberFormat()
    {
        var call = SyntheticJob.ValidAdvancedReportSpecCall();
        call.ArgumentsJson = call.ArgumentsJson.Replace(
            "\"numberFormat\":\"0.0%\"",
            "\"numberFormat\":\"=SUM(A:A)\"",
            StringComparison.Ordinal);

        var result = AgentToolCallValidator.Validate(new[] { call }, SyntheticJob.Create().Data);

        Assert.False(result.IsValid);
        Assert.Equal("tool_arguments_invalid", result.ErrorCode);
    }

    [Fact]
    public void Validate_AcceptsExpandedTypedTransformsAndTheirOutputFields()
    {
        var transforms = new AgentToolCall
        {
            Id = "transforms-advanced",
            Name = AgentToolNames.ProposeTransforms,
            ArgumentsJson =
                "{\"transforms\":[" +
                "{\"kind\":\"trimText\",\"sourceField\":\"Department\",\"outputField\":\"Department\"}," +
                "{\"kind\":\"mapValues\",\"sourceField\":\"Department\",\"outputField\":\"Reporting Group\",\"mappings\":[{\"from\":\"Operations\",\"to\":\"Core\"}]}," +
                "{\"kind\":\"derivePeriodPart\",\"sourceField\":\"Period\",\"outputField\":\"Quarter\",\"part\":\"quarter\"}," +
                "{\"kind\":\"addArithmeticColumn\",\"sourceField\":\"Net Amount\",\"outputField\":\"Adjusted Amount\",\"operator\":\"multiply\",\"rightField\":\"\",\"rightNumber\":1.1}]}"
        };

        var result = AgentToolCallValidator.Validate(new[] { transforms }, SyntheticJob.Create().Data);

        Assert.True(result.IsValid);
    }

    [Fact]
    public void Validate_RejectsExecutablePropertyOnTypedTransform()
    {
        var transforms = new AgentToolCall
        {
            Id = "transforms-unsafe",
            Name = AgentToolNames.ProposeTransforms,
            ArgumentsJson =
                "{\"transforms\":[{\"kind\":\"fillDown\",\"sourceField\":\"Department\"," +
                "\"outputField\":\"Department\",\"code\":\"run\"}]}"
        };

        Assert.False(AgentToolCallValidator.Validate(
            new[] { transforms },
            SyntheticJob.Create().Data).IsValid);
    }

    [Fact]
    public void Catalog_ReportSpecSchemaPublishesTheVersionedTypedGrammar()
    {
        var reportTool = AgentToolCatalog.Definitions.Single(definition =>
            definition.Function.Name == AgentToolNames.ProposeReportSpec);

        Assert.Equal("1.0", reportTool.Function.Parameters
            .GetProperty("properties")
            .GetProperty("version")
            .GetProperty("const")
            .GetString());
        Assert.True(reportTool.Function.Parameters.GetProperty("$defs").TryGetProperty("expression", out _));
    }

    [Fact]
    public void Catalog_IsExactlyTheApprovedBoundedWorkflow()
    {
        var names = AgentToolCatalog.Definitions.Select(definition => definition.Function.Name).ToArray();

        Assert.Equal(
            new[]
            {
                AgentToolNames.ProposePeriodMapping,
                AgentToolNames.ProposeTransforms,
                AgentToolNames.ProposeReportSpec,
                AgentToolNames.ValidateSpec,
                AgentToolNames.RequestManagedDraftBuild,
                AgentToolNames.RunChecks,
                AgentToolNames.FinalChangeSummary,
            },
            names);
    }

    internal static List<AgentToolCall> ApprovedWorkflow()
    {
        return new List<AgentToolCall>
        {
            SyntheticJob.ValidReportSpecCall("spec-call"),
            new AgentToolCall
            {
                Id = "validate-call",
                Name = AgentToolNames.ValidateSpec,
                ArgumentsJson = "{\"proposalToolCallId\":\"spec-call\"}",
            },
            new AgentToolCall
            {
                Id = "build-call",
                Name = AgentToolNames.RequestManagedDraftBuild,
                ArgumentsJson = "{\"validatedSpecificationId\":\"validated-spec-1\"}",
            },
            new AgentToolCall
            {
                Id = "checks-call",
                Name = AgentToolNames.RunChecks,
                ArgumentsJson = "{\"managedDraftId\":\"draft-1\",\"checks\":[\"sourceTotals\",\"periodCoverage\"]}",
            },
            new AgentToolCall
            {
                Id = "summary-call",
                Name = AgentToolNames.FinalChangeSummary,
                ArgumentsJson =
                    "{\"managedDraftId\":\"draft-1\",\"allChecksPassed\":true," +
                    "\"changes\":[{\"category\":\"checks\",\"description\":\"All requested checks passed.\"}]}",
            },
        };
    }
}
