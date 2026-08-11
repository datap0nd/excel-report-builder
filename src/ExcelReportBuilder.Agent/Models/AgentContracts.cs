using System;
using System.Collections.Generic;
using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.Agent.Models;

public sealed class AgentJobRequest
{
    public string JobId { get; set; } = string.Empty;

    /// <summary>
    /// Stable host-generated identifier for the workbook. It must not contain
    /// a workbook name or path.
    /// </summary>
    public string WorkbookId { get; set; } = string.Empty;

    public string UserPrompt { get; set; } = string.Empty;

    public AgentDataSnapshot Data { get; set; } = new AgentDataSnapshot();

    public AgentSpecificationSnapshot CurrentSpecification { get; set; } = new AgentSpecificationSnapshot();

    public AgentEndpointSettings Endpoint { get; set; } = new AgentEndpointSettings();

    public int MaxRepairCycles { get; set; } = AgentDefaults.MaxRepairCycles;

    public string? ResumeFromCheckpointId { get; set; }
}

public sealed class AgentDataSnapshot
{
    public string SourceDisplayName { get; set; } = "Selected data";

    public long RowCount { get; set; }

    public List<AgentField> Fields { get; set; } = new List<AgentField>();

    public List<AgentSampleRow> SampleRows { get; set; } = new List<AgentSampleRow>();

    public int? ReportingYear { get; set; }
}

public sealed class AgentField
{
    public string Name { get; set; } = string.Empty;

    public AgentFieldType Type { get; set; }

    public bool AllowsBlank { get; set; }
}

public enum AgentFieldType
{
    Text,
    Number,
    Date,
    Boolean,
}

public sealed class AgentSampleRow
{
    public List<AgentSampleValue> Values { get; set; } = new List<AgentSampleValue>();
}

public sealed class AgentSampleValue
{
    public string Field { get; set; } = string.Empty;

    /// <summary>
    /// A display-only scalar. The worker never interprets this as a formula.
    /// </summary>
    public string? Value { get; set; }
}

public sealed class AgentSpecificationSnapshot
{
    /// <summary>
    /// Optional host-validated ReportSpecV1 JSON for an already applied agent
    /// setup. It is data only and cannot contain executable tools or formulas.
    /// </summary>
    public string CanonicalReportSpecJson { get; set; } = string.Empty;

    public List<string> Rows { get; set; } = new List<string>();

    public List<string> Columns { get; set; } = new List<string>();

    public List<AgentValuePlacement> Values { get; set; } = new List<AgentValuePlacement>();

    public List<AgentFilterPlacement> Filters { get; set; } = new List<AgentFilterPlacement>();
}

public sealed class AgentValuePlacement
{
    public string Field { get; set; } = string.Empty;

    public string Aggregation { get; set; } = "sum";
}

public sealed class AgentFilterPlacement
{
    public string Field { get; set; } = string.Empty;

    public string Operator { get; set; } = "equals";

    public List<string> Values { get; set; } = new List<string>();
}

public sealed class AgentToolCall
{
    public string Id { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string ArgumentsJson { get; set; } = "{}";
}

public sealed class ValidatedAgentToolCall
{
    public string Id { get; set; } = string.Empty;

    public string Name { get; set; } = string.Empty;

    public string ArgumentsJson { get; set; } = "{}";
}

public sealed class AgentModelProposal
{
    public string Model { get; set; } = string.Empty;

    public string? AssistantText { get; set; }

    public List<AgentToolCall> ToolCalls { get; set; } = new List<AgentToolCall>();
}

public sealed class AgentRunResult
{
    public string JobId { get; set; } = string.Empty;

    public string Model { get; set; } = string.Empty;

    public string WorkbookId { get; set; } = string.Empty;

    public int RepairCyclesUsed { get; set; }

    public List<ValidatedAgentToolCall> ToolCalls { get; set; } = new List<ValidatedAgentToolCall>();

    public List<HostToolResultSummary> HostToolResults { get; set; } = new List<HostToolResultSummary>();
}

public sealed class ModelDiscoveryResult
{
    public List<string> ModelIds { get; set; } = new List<string>();

    public string SelectedModel { get; set; } = string.Empty;
}

public sealed class EndpointProbeResult
{
    public bool ModelsEndpointAvailable { get; set; }

    public bool ToolCallingAvailable { get; set; }

    public bool StructuredOutputAvailable { get; set; }

    public string SelectedModel { get; set; } = string.Empty;

    public List<string> DiscoveredModels { get; set; } = new List<string>();

    public string Summary { get; set; } = string.Empty;
}

public sealed class HostToolResultSummary
{
    public string ToolCallId { get; set; } = string.Empty;

    public string ToolName { get; set; } = string.Empty;

    public bool Succeeded { get; set; }

    public string OutcomeCode { get; set; } = string.Empty;

    public int CheckFailureCount { get; set; }
}

public sealed class AgentToolValidationResult
{
    public bool IsValid { get; set; }

    public string ErrorCode { get; set; } = string.Empty;

    public string RepairInstruction { get; set; } = string.Empty;

    public List<ValidatedAgentToolCall> ToolCalls { get; set; } = new List<ValidatedAgentToolCall>();
}
