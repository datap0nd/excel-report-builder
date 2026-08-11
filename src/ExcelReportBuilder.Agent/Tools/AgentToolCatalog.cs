using System;
using System.Collections.Generic;
using System.Text.Json;

namespace ExcelReportBuilder.Agent.Tools;

public static class AgentToolNames
{
    public const string ProposePeriodMapping = "propose_period_mapping";
    public const string ProposeTransforms = "propose_transforms";
    public const string ProposeReportSpec = "propose_report_spec";
    public const string ValidateSpec = "validate_spec";
    public const string RequestManagedDraftBuild = "request_managed_draft_build";
    public const string RunChecks = "run_checks";
    public const string FinalChangeSummary = "final_change_summary";
}

public sealed class OpenAiToolDefinition
{
    public string Type { get; set; } = "function";

    public OpenAiFunctionDefinition Function { get; set; } = new OpenAiFunctionDefinition();
}

public sealed class OpenAiFunctionDefinition
{
    public string Name { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    public bool Strict { get; set; } = true;

    public JsonElement Parameters { get; set; }
}

public static class AgentToolCatalog
{
    private static readonly HashSet<string> AllowedNames = new HashSet<string>(StringComparer.Ordinal)
    {
        AgentToolNames.ProposePeriodMapping,
        AgentToolNames.ProposeTransforms,
        AgentToolNames.ProposeReportSpec,
        AgentToolNames.ValidateSpec,
        AgentToolNames.RequestManagedDraftBuild,
        AgentToolNames.RunChecks,
        AgentToolNames.FinalChangeSummary,
    };

    private static readonly IReadOnlyList<OpenAiToolDefinition> ToolDefinitions = CreateDefinitions();

    public static IReadOnlyList<OpenAiToolDefinition> Definitions => ToolDefinitions;

    public static bool IsAllowed(string? name)
    {
        return name != null && AllowedNames.Contains(name);
    }

    public static int GetFlowOrder(string name)
    {
        switch (name)
        {
            case AgentToolNames.ProposePeriodMapping: return 0;
            case AgentToolNames.ProposeTransforms: return 1;
            case AgentToolNames.ProposeReportSpec: return 2;
            case AgentToolNames.ValidateSpec: return 3;
            case AgentToolNames.RequestManagedDraftBuild: return 4;
            case AgentToolNames.RunChecks: return 5;
            case AgentToolNames.FinalChangeSummary: return 6;
            default: return int.MaxValue;
        }
    }

    private static IReadOnlyList<OpenAiToolDefinition> CreateDefinitions()
    {
        return new List<OpenAiToolDefinition>
        {
            Define(
                AgentToolNames.ProposePeriodMapping,
                "Propose a bounded date-column, wide-period, or unresolved period mapping. Never invent a reporting year.",
                AgentToolSchemaFactory.PeriodMapping()),
            Define(
                AgentToolNames.ProposeTransforms,
                "Propose only allowlisted declarative data transforms, including type changes, normalization, fill down, value maps, text filters, period parts, and typed arithmetic. This tool cannot execute code or formulas.",
                AgentToolSchemaFactory.AdvancedTransforms()),
            Define(
                AgentToolNames.ProposeReportSpec,
                "Propose a complete versioned ReportSpecV1 using only typed measures, period slices, managed report blocks, layout options, styles, and checks. Never emit formulas or code.",
                AgentToolSchemaFactory.AdvancedReportSpecification()),
            Define(
                AgentToolNames.ValidateSpec,
                "Ask the deterministic host to validate a proposed report specification.",
                ReferenceSchema("proposalToolCallId")),
            Define(
                AgentToolNames.RequestManagedDraftBuild,
                "Ask the deterministic host to build an explicitly managed draft from a validated specification. This does not save or publish.",
                ReferenceSchema("validatedSpecificationId")),
            Define(
                AgentToolNames.RunChecks,
                "Ask the deterministic host to run allowlisted independent checks against a managed draft.",
                @"{""type"":""object"",""properties"":{""managedDraftId"":{""type"":""string"",""maxLength"":128},""checks"":{""type"":""array"",""minItems"":1,""maxItems"":8,""uniqueItems"":true,""items"":{""type"":""string"",""enum"":[""sourceTotals"",""grandTotals"",""rowCounts"",""periodCoverage"",""formulaErrors"",""managedOwnership""]}}},""required"":[""managedDraftId"",""checks""],""additionalProperties"":false}"),
            Define(
                AgentToolNames.FinalChangeSummary,
                "Return a bounded final change and check summary for host display. This tool cannot publish or save.",
                @"{""type"":""object"",""properties"":{""managedDraftId"":{""type"":""string"",""maxLength"":128},""allChecksPassed"":{""type"":""boolean""},""changes"":{""type"":""array"",""maxItems"":20,""items"":{""type"":""object"",""properties"":{""category"":{""type"":""string"",""enum"":[""data"",""rows"",""columns"",""values"",""filters"",""formatting"",""checks""]},""description"":{""type"":""string"",""maxLength"":256}},""required"":[""category"",""description""],""additionalProperties"":false}}},""required"":[""managedDraftId"",""allChecksPassed"",""changes""],""additionalProperties"":false}"),
        };
    }

    private static string ReferenceSchema(string propertyName)
    {
        return "{\"type\":\"object\",\"properties\":{\"" + propertyName +
               "\":{\"type\":\"string\",\"maxLength\":128}},\"required\":[\"" +
               propertyName + "\"],\"additionalProperties\":false}";
    }

    private static OpenAiToolDefinition Define(string name, string description, string schema)
    {
        using (var document = JsonDocument.Parse(schema))
        {
            return new OpenAiToolDefinition
            {
                Function = new OpenAiFunctionDefinition
                {
                    Name = name,
                    Description = description,
                    Parameters = document.RootElement.Clone(),
                },
            };
        }
    }

    private static OpenAiToolDefinition Define(string name, string description, object schema)
    {
        using (var document = JsonDocument.Parse(JsonSerializer.Serialize(schema)))
        {
            return new OpenAiToolDefinition
            {
                Function = new OpenAiFunctionDefinition
                {
                    Name = name,
                    Description = description,
                    Parameters = document.RootElement.Clone(),
                },
            };
        }
    }
}
