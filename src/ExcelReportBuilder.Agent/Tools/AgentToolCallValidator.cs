using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Tools;

public static class AgentToolCallValidator
{
    public const int MaximumToolCalls = 16;
    public const int MaximumArgumentsBytes = 128 * 1024;

    public static AgentToolValidationResult Validate(
        IReadOnlyList<AgentToolCall>? toolCalls,
        AgentDataSnapshot data,
        bool requireCompleteWorkflow = false)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (toolCalls == null || toolCalls.Count == 0)
        {
            return Invalid("tool_call_required", "Return at least one allowlisted report workflow tool call.");
        }

        if (toolCalls.Count > MaximumToolCalls)
        {
            return Invalid("too_many_tool_calls", "Return fewer report workflow tool calls.");
        }

        var fields = new Dictionary<string, AgentField>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in data.Fields) fields[field.Name] = field;
        AddProposedOutputFields(toolCalls, fields);

        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        var seenNames = new HashSet<string>(StringComparer.Ordinal);
        var validated = new List<ValidatedAgentToolCall>();
        var lastFlowOrder = -1;

        foreach (var toolCall in toolCalls)
        {
            if (toolCall == null || !IsBoundedIdentifier(toolCall.Id) || !seenIds.Add(toolCall.Id))
            {
                return Invalid("tool_call_id_invalid", "Return a unique bounded ID for every tool call.");
            }

            if (!AgentToolCatalog.IsAllowed(toolCall.Name))
            {
                return Invalid("tool_not_allowed", "Use only the report workflow tools supplied in the request.");
            }

            if (!seenNames.Add(toolCall.Name))
            {
                return Invalid("duplicate_tool", "Call each report workflow tool at most once per proposal.");
            }

            var flowOrder = AgentToolCatalog.GetFlowOrder(toolCall.Name);
            if (flowOrder < lastFlowOrder)
            {
                return Invalid("tool_order_invalid", "Call report workflow tools in proposal, validation, managed draft, checks, and summary order.");
            }

            lastFlowOrder = flowOrder;
            if (string.IsNullOrWhiteSpace(toolCall.ArgumentsJson) ||
                Encoding.UTF8.GetByteCount(toolCall.ArgumentsJson) > MaximumArgumentsBytes)
            {
                return Invalid("tool_arguments_too_large", "Return smaller JSON arguments for each tool call.");
            }

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(toolCall.ArgumentsJson);
            }
            catch (JsonException)
            {
                return Invalid("tool_arguments_invalid_json", "Return valid JSON object arguments for every tool call.");
            }

            using (document)
            {
                if (document.RootElement.ValueKind != JsonValueKind.Object ||
                    !ValidateArguments(toolCall.Name, document.RootElement, fields, data))
                {
                    return Invalid(
                        "tool_arguments_invalid",
                        "Match the supplied schema exactly, use only bounded identifiers and listed columns, and never invent a reporting year.");
                }

                validated.Add(new ValidatedAgentToolCall
                {
                    Id = toolCall.Id,
                    Name = toolCall.Name,
                    ArgumentsJson = document.RootElement.GetRawText(),
                });
            }
        }

        if (requireCompleteWorkflow &&
            (!seenNames.Contains(AgentToolNames.ProposeReportSpec) ||
             !seenNames.Contains(AgentToolNames.ValidateSpec) ||
             !seenNames.Contains(AgentToolNames.RequestManagedDraftBuild) ||
             !seenNames.Contains(AgentToolNames.RunChecks) ||
             !seenNames.Contains(AgentToolNames.FinalChangeSummary)))
        {
            return Invalid(
                "workflow_incomplete",
                "Return the complete guarded workflow: propose_report_spec, validate_spec, request_managed_draft_build, run_checks, and final_change_summary in order.");
        }

        return new AgentToolValidationResult { IsValid = true, ToolCalls = validated };
    }

    private static bool ValidateArguments(
        string name,
        JsonElement arguments,
        Dictionary<string, AgentField> fields,
        AgentDataSnapshot data)
    {
        switch (name)
        {
            case AgentToolNames.ProposePeriodMapping:
                return ValidatePeriodMapping(arguments, fields, data.ReportingYear);
            case AgentToolNames.ProposeTransforms:
                return ValidateTransforms(arguments, fields);
            case AgentToolNames.ProposeReportSpec:
                return ValidateReportSpec(arguments, fields);
            case AgentToolNames.ValidateSpec:
                return ValidateReference(arguments, "proposalToolCallId");
            case AgentToolNames.RequestManagedDraftBuild:
                return ValidateReference(arguments, "validatedSpecificationId");
            case AgentToolNames.RunChecks:
                return ValidateChecks(arguments);
            case AgentToolNames.FinalChangeSummary:
                return ValidateFinalSummary(arguments);
            default:
                return false;
        }
    }

    private static bool ValidatePeriodMapping(
        JsonElement value,
        Dictionary<string, AgentField> fields,
        int? knownReportingYear)
    {
        if (!HasOnlyProperties(value, "mode", "periodField", "reportingYear", "mappings") ||
            !TryGetString(value, "mode", out var mode) ||
            !IsOneOf(mode, "dateColumn", "widePeriods", "unresolved") ||
            !TryGetString(value, "periodField", out var periodField) ||
            periodField!.Length > 128 ||
            !value.TryGetProperty("reportingYear", out var year) ||
            !value.TryGetProperty("mappings", out var mappings) ||
            mappings.ValueKind != JsonValueKind.Array || mappings.GetArrayLength() > 120)
        {
            return false;
        }

        int? proposedYear = null;
        if (year.ValueKind == JsonValueKind.Number)
        {
            if (!year.TryGetInt32(out var parsedYear) || parsedYear < 1900 || parsedYear > 9999) return false;
            proposedYear = parsedYear;
        }
        else if (year.ValueKind != JsonValueKind.Null)
        {
            return false;
        }

        if (!knownReportingYear.HasValue && proposedYear.HasValue)
        {
            return false;
        }

        if (knownReportingYear.HasValue && proposedYear.HasValue && knownReportingYear.Value != proposedYear.Value)
        {
            return false;
        }

        if (string.Equals(mode, "dateColumn", StringComparison.Ordinal) &&
            (!fields.TryGetValue(periodField, out var period) || period.Type != AgentFieldType.Date || mappings.GetArrayLength() != 0))
        {
            return false;
        }

        if (string.Equals(mode, "unresolved", StringComparison.Ordinal) &&
            (periodField.Length != 0 || proposedYear.HasValue || mappings.GetArrayLength() != 0))
        {
            return false;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mapping in mappings.EnumerateArray())
        {
            if (mapping.ValueKind != JsonValueKind.Object ||
                !HasOnlyProperties(mapping, "sourceField", "periodLabel", "metric") ||
                !TryGetString(mapping, "sourceField", out var sourceField) ||
                !TryGetString(mapping, "periodLabel", out var periodLabel) ||
                !TryGetString(mapping, "metric", out var metric) ||
                !fields.ContainsKey(sourceField!) || !seen.Add(sourceField!) ||
                !IsBoundedText(periodLabel, 32) ||
                metric!.Length > 128 || metric.Any(char.IsControl))
            {
                return false;
            }
        }

        return !string.Equals(mode, "widePeriods", StringComparison.Ordinal) ||
               (knownReportingYear.HasValue && proposedYear == knownReportingYear && mappings.GetArrayLength() > 0);
    }

    private static bool ValidateTransforms(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (!HasOnlyProperties(value, "transforms") ||
            !value.TryGetProperty("transforms", out var transforms) ||
            transforms.ValueKind != JsonValueKind.Array || transforms.GetArrayLength() > 32)
        {
            return false;
        }

        var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var transform in transforms.EnumerateArray())
        {
            if (transform.ValueKind != JsonValueKind.Object ||
                !TryGetString(transform, "kind", out var kind) ||
                !TryGetString(transform, "sourceField", out var sourceField) ||
                !fields.TryGetValue(sourceField!, out var source) ||
                !TryGetString(transform, "outputField", out var outputField) ||
                !IsBoundedText(outputField, 128) || !outputs.Add(outputField!) ||
                !ValidateTransformShape(transform, kind!, source, fields, sourceField!, outputField!))
            {
                return false;
            }

            AgentFieldType outputType = source.Type;
            if (string.Equals(kind, "convertNumber", StringComparison.Ordinal) ||
                string.Equals(kind, "addArithmeticColumn", StringComparison.Ordinal))
            {
                outputType = AgentFieldType.Number;
            }
            else if (string.Equals(kind, "convertDate", StringComparison.Ordinal))
            {
                outputType = AgentFieldType.Date;
            }
            else if (string.Equals(kind, "derivePeriodPart", StringComparison.Ordinal))
            {
                outputType = AgentFieldType.Text;
            }

            fields[outputField!] = new AgentField
            {
                Name = outputField!,
                Type = outputType,
                AllowsBlank = source.AllowsBlank
            };
            if (string.Equals(kind, "rename", StringComparison.Ordinal) &&
                !string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase))
            {
                fields.Remove(sourceField!);
            }
        }

        return true;
    }

    private static bool ValidateTransformShape(
        JsonElement transform,
        string kind,
        AgentField source,
        IReadOnlyDictionary<string, AgentField> fields,
        string sourceField,
        string outputField)
    {
        switch (kind)
        {
            case "rename":
            case "trimText":
            case "convertNumber":
            case "convertDate":
            case "replaceBlank":
            case "normalizeBlanks":
            case "normalizeErrors":
            case "fillDown":
                return HasOnlyProperties(transform, "kind", "sourceField", "outputField") &&
                    (!string.Equals(kind, "trimText", StringComparison.Ordinal) || source.Type == AgentFieldType.Text);
            case "filterRows":
                if (!HasOnlyProperties(transform, "kind", "sourceField", "outputField", "operator", "value") ||
                    !string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase) ||
                    source.Type != AgentFieldType.Text ||
                    !TryGetString(transform, "operator", out var filterOperator) ||
                    !IsOneOf(
                        filterOperator,
                        "equal", "notEqual", "contains", "startsWith", "endsWith", "isBlank", "isNotBlank") ||
                    !TryGetString(transform, "value", out var filterValue) ||
                    filterValue!.Length > 1024 || filterValue.Any(char.IsControl))
                {
                    return false;
                }

                return IsOneOf(filterOperator, "isBlank", "isNotBlank")
                    ? filterValue.Length == 0
                    : filterValue.Length != 0;
            case "mapValues":
                if (!HasOnlyProperties(transform, "kind", "sourceField", "outputField", "mappings") ||
                    source.Type != AgentFieldType.Text ||
                    !transform.TryGetProperty("mappings", out var mappings) ||
                    mappings.ValueKind != JsonValueKind.Array || mappings.GetArrayLength() == 0 ||
                    mappings.GetArrayLength() > 200)
                {
                    return false;
                }

                foreach (var mapping in mappings.EnumerateArray())
                {
                    if (mapping.ValueKind != JsonValueKind.Object ||
                        !HasOnlyProperties(mapping, "from", "to") ||
                        !TryGetString(mapping, "from", out var from) || from!.Length > 1024 || from.Any(char.IsControl) ||
                        !TryGetString(mapping, "to", out var to) || to!.Length > 1024 || to.Any(char.IsControl))
                    {
                        return false;
                    }
                }

                return true;
            case "excludeTotalRows":
                if (!HasOnlyProperties(
                        transform,
                        "kind",
                        "sourceField",
                        "outputField",
                        "matchKind",
                        "values",
                        "evidenceSource",
                        "observedMatchCount") ||
                    source.Type != AgentFieldType.Text ||
                    !string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase) ||
                    !TryGetString(transform, "matchKind", out var matchKind) ||
                    !IsOneOf(matchKind, "equalsAny", "startsWith", "contains", "isBlank") ||
                    !transform.TryGetProperty("values", out var totalValues) ||
                    totalValues.ValueKind != JsonValueKind.Array || totalValues.GetArrayLength() > 50 ||
                    !TryGetString(transform, "evidenceSource", out var evidenceSource) ||
                    !IsOneOf(evidenceSource, "profile", "preview", "userConfirmation") ||
                    !transform.TryGetProperty("observedMatchCount", out var observed) ||
                    observed.ValueKind != JsonValueKind.Number ||
                    !observed.TryGetInt64(out long observedCount) || observedCount < 0)
                {
                    return false;
                }

                if ((string.Equals(matchKind, "isBlank", StringComparison.Ordinal) && totalValues.GetArrayLength() != 0) ||
                    (!string.Equals(matchKind, "isBlank", StringComparison.Ordinal) && totalValues.GetArrayLength() == 0))
                {
                    return false;
                }

                foreach (var totalValue in totalValues.EnumerateArray())
                {
                    if (totalValue.ValueKind != JsonValueKind.String ||
                        totalValue.GetString()!.Length > 1024 ||
                        totalValue.GetString()!.Any(char.IsControl))
                    {
                        return false;
                    }
                }

                return true;
            case "derivePeriodPart":
                return HasOnlyProperties(transform, "kind", "sourceField", "outputField", "part") &&
                    source.Type == AgentFieldType.Date &&
                    !string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase) &&
                    TryGetString(transform, "part", out var part) &&
                    IsOneOf(part, "year", "half", "quarter", "monthNumber", "monthName", "yearMonth");
            case "addArithmeticColumn":
                if (!HasOnlyProperties(
                        transform,
                        "kind",
                        "sourceField",
                        "outputField",
                        "operator",
                        "rightField",
                        "rightNumber") ||
                    source.Type != AgentFieldType.Number ||
                    string.Equals(sourceField, outputField, StringComparison.OrdinalIgnoreCase) ||
                    !TryGetString(transform, "operator", out var arithmeticOperator) ||
                    !IsOneOf(arithmeticOperator, "add", "subtract", "multiply", "divide") ||
                    !TryGetString(transform, "rightField", out var rightField) ||
                    !transform.TryGetProperty("rightNumber", out var rightNumber))
                {
                    return false;
                }

                if (rightField!.Length != 0)
                {
                    return rightNumber.ValueKind == JsonValueKind.Null &&
                        fields.TryGetValue(rightField, out var right) &&
                        right.Type == AgentFieldType.Number;
                }

                return rightNumber.ValueKind == JsonValueKind.Number &&
                    rightNumber.TryGetDecimal(out decimal number) &&
                    (!string.Equals(arithmeticOperator, "divide", StringComparison.Ordinal) || number != 0m);
            default:
                return false;
        }
    }

    private static bool ValidateReportSpec(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (value.TryGetProperty("version", out _))
        {
            return AgentAdvancedReportProposalValidator.Validate(value, fields);
        }

        if (!HasOnlyProperties(value, "rows", "columns", "values", "filters", "subtotals", "formatting", "ordering") ||
            !ValidateFieldArray(value, "rows", fields, 32) ||
            !ValidateFieldArray(value, "columns", fields, 16) ||
            !ValidateValues(value, fields) ||
            !ValidateFilters(value, fields) ||
            !ValidateFieldModeArray(value, "subtotals", "mode", fields, "show", "hide") ||
            !ValidateFormatting(value, fields) ||
            !ValidateOrdering(value, fields))
        {
            return false;
        }

        return true;
    }

    private static bool ValidateValues(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (!value.TryGetProperty("values", out var values) || values.ValueKind != JsonValueKind.Array || values.GetArrayLength() > 32) return false;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in values.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || !HasOnlyProperties(item, "field", "aggregation") ||
                !TryGetString(item, "field", out var fieldName) || !fields.TryGetValue(fieldName!, out var field) || !seen.Add(fieldName!) ||
                !TryGetString(item, "aggregation", out var aggregation) ||
                !IsOneOf(aggregation, "sum", "count", "min", "max", "average") ||
                (!string.Equals(aggregation, "count", StringComparison.Ordinal) && field.Type != AgentFieldType.Number))
            {
                return false;
            }
        }

        return true;
    }

    private static bool ValidateFilters(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (!value.TryGetProperty("filters", out var filters) || filters.ValueKind != JsonValueKind.Array || filters.GetArrayLength() > 32) return false;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var filter in filters.EnumerateArray())
        {
            if (filter.ValueKind != JsonValueKind.Object || !HasOnlyProperties(filter, "field", "operator", "values") ||
                !TryGetString(filter, "field", out var field) || !fields.ContainsKey(field!) || !seen.Add(field!) ||
                !TryGetString(filter, "operator", out var filterOperator) ||
                !IsOneOf(filterOperator, "equals", "notEquals", "in", "notIn", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "between", "isBlank", "isNotBlank") ||
                !filter.TryGetProperty("values", out var filterValues) || filterValues.ValueKind != JsonValueKind.Array || filterValues.GetArrayLength() > 50)
            {
                return false;
            }

            var count = filterValues.GetArrayLength();
            if ((IsOneOf(filterOperator, "isBlank", "isNotBlank") && count != 0) ||
                (string.Equals(filterOperator, "between", StringComparison.Ordinal) && count != 2) ||
                (!IsOneOf(filterOperator, "isBlank", "isNotBlank", "between") && count == 0)) return false;
            foreach (var scalar in filterValues.EnumerateArray())
            {
                if (scalar.ValueKind != JsonValueKind.String || !IsBoundedText(scalar.GetString(), 1024)) return false;
            }
        }

        return true;
    }

    private static bool ValidateFieldModeArray(
        JsonElement value,
        string propertyName,
        string modeName,
        Dictionary<string, AgentField> fields,
        params string[] choices)
    {
        if (!value.TryGetProperty(propertyName, out var items) || items.ValueKind != JsonValueKind.Array || items.GetArrayLength() > 32) return false;
        foreach (var item in items.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || !HasOnlyProperties(item, "field", modeName) ||
                !TryGetString(item, "field", out var field) || !fields.ContainsKey(field!) ||
                !TryGetString(item, modeName, out var mode) || !IsOneOf(mode, choices)) return false;
        }

        return true;
    }

    private static bool ValidateFormatting(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (!value.TryGetProperty("formatting", out var items) || items.ValueKind != JsonValueKind.Array || items.GetArrayLength() > 32) return false;
        foreach (var item in items.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || !HasOnlyProperties(item, "field", "numberStyle", "decimalPlaces") ||
                !TryGetString(item, "field", out var fieldName) || !fields.TryGetValue(fieldName!, out var field) || field.Type != AgentFieldType.Number ||
                !TryGetString(item, "numberStyle", out var style) || !IsOneOf(style, "general", "integer", "decimal", "currency", "percentage") ||
                !item.TryGetProperty("decimalPlaces", out var places) || places.ValueKind != JsonValueKind.Number ||
                !places.TryGetInt32(out var parsedPlaces) || parsedPlaces < 0 || parsedPlaces > 6) return false;
        }

        return true;
    }

    private static bool ValidateOrdering(JsonElement value, Dictionary<string, AgentField> fields)
    {
        if (!value.TryGetProperty("ordering", out var items) || items.ValueKind != JsonValueKind.Array || items.GetArrayLength() > 32) return false;
        foreach (var item in items.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || !HasOnlyProperties(item, "field", "direction", "by") ||
                !TryGetString(item, "field", out var field) || !fields.ContainsKey(field!) ||
                !TryGetString(item, "direction", out var direction) || !IsOneOf(direction, "ascending", "descending") ||
                !TryGetString(item, "by", out var by) || !IsOneOf(by, "label", "value")) return false;
        }

        return true;
    }

    private static bool ValidateReference(JsonElement value, string propertyName)
    {
        return HasOnlyProperties(value, propertyName) && TryGetString(value, propertyName, out var id) && IsBoundedIdentifier(id);
    }

    private static bool ValidateChecks(JsonElement value)
    {
        if (!HasOnlyProperties(value, "managedDraftId", "checks") ||
            !TryGetString(value, "managedDraftId", out var draftId) || !IsBoundedIdentifier(draftId) ||
            !value.TryGetProperty("checks", out var checks) || checks.ValueKind != JsonValueKind.Array ||
            checks.GetArrayLength() == 0 || checks.GetArrayLength() > 8) return false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var check in checks.EnumerateArray())
        {
            if (check.ValueKind != JsonValueKind.String ||
                !IsOneOf(check.GetString(), "sourceTotals", "grandTotals", "rowCounts", "periodCoverage", "formulaErrors", "managedOwnership") ||
                !seen.Add(check.GetString()!)) return false;
        }

        return true;
    }

    private static bool ValidateFinalSummary(JsonElement value)
    {
        if (!HasOnlyProperties(value, "managedDraftId", "allChecksPassed", "changes") ||
            !TryGetString(value, "managedDraftId", out var draftId) || !IsBoundedIdentifier(draftId) ||
            !value.TryGetProperty("allChecksPassed", out var passed) ||
            (passed.ValueKind != JsonValueKind.True && passed.ValueKind != JsonValueKind.False) ||
            !value.TryGetProperty("changes", out var changes) || changes.ValueKind != JsonValueKind.Array || changes.GetArrayLength() > 20) return false;
        foreach (var change in changes.EnumerateArray())
        {
            if (change.ValueKind != JsonValueKind.Object || !HasOnlyProperties(change, "category", "description") ||
                !TryGetString(change, "category", out var category) ||
                !IsOneOf(category, "data", "rows", "columns", "values", "filters", "formatting", "checks") ||
                !TryGetString(change, "description", out var description) || !IsBoundedText(description, 256)) return false;
        }

        return true;
    }

    private static bool ValidateFieldArray(JsonElement value, string propertyName, Dictionary<string, AgentField> fields, int maximum)
    {
        if (!value.TryGetProperty(propertyName, out var items) || items.ValueKind != JsonValueKind.Array || items.GetArrayLength() > maximum) return false;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String || !fields.ContainsKey(item.GetString()!) || !seen.Add(item.GetString()!)) return false;
        }

        return true;
    }

    private static void AddProposedOutputFields(IReadOnlyList<AgentToolCall> toolCalls, Dictionary<string, AgentField> fields)
    {
        foreach (var call in toolCalls)
        {
            if (call == null || !string.Equals(call.Name, AgentToolNames.ProposeTransforms, StringComparison.Ordinal)) continue;
            try
            {
                using (var document = JsonDocument.Parse(call.ArgumentsJson))
                {
                    if (!document.RootElement.TryGetProperty("transforms", out var transforms) || transforms.ValueKind != JsonValueKind.Array) continue;
                    foreach (var transform in transforms.EnumerateArray())
                    {
                        if (!TryGetString(transform, "outputField", out var output) || !IsBoundedText(output, 128)) continue;
                        var type = AgentFieldType.Text;
                        if (TryGetString(transform, "kind", out var kind))
                        {
                            if (string.Equals(kind, "convertNumber", StringComparison.Ordinal)) type = AgentFieldType.Number;
                            if (string.Equals(kind, "addArithmeticColumn", StringComparison.Ordinal)) type = AgentFieldType.Number;
                            if (string.Equals(kind, "convertDate", StringComparison.Ordinal)) type = AgentFieldType.Date;
                            if (string.Equals(kind, "derivePeriodPart", StringComparison.Ordinal)) type = AgentFieldType.Text;
                        }

                        if (!fields.ContainsKey(output!)) fields[output!] = new AgentField { Name = output!, Type = type, AllowsBlank = true };
                    }
                }
            }
            catch (JsonException)
            {
            }
        }
    }

    private static bool HasOnlyProperties(JsonElement value, params string[] requiredProperties)
    {
        var names = new HashSet<string>(requiredProperties, StringComparer.Ordinal);
        var count = 0;
        foreach (var property in value.EnumerateObject())
        {
            count++;
            if (!names.Contains(property.Name)) return false;
        }

        return count == names.Count;
    }

    private static bool TryGetString(JsonElement value, string propertyName, out string? result)
    {
        result = null;
        if (!value.TryGetProperty(propertyName, out var property) || property.ValueKind != JsonValueKind.String) return false;
        result = property.GetString();
        return result != null;
    }

    private static bool IsOneOf(string? value, params string[] choices)
    {
        if (value == null) return false;
        foreach (var choice in choices)
        {
            if (string.Equals(value, choice, StringComparison.Ordinal)) return true;
        }

        return false;
    }

    private static bool IsBoundedIdentifier(string? value)
    {
        if (!IsBoundedText(value, 128)) return false;
        foreach (var character in value!)
        {
            var allowed = (character >= 'a' && character <= 'z') ||
                          (character >= 'A' && character <= 'Z') ||
                          (character >= '0' && character <= '9') ||
                          character == '-' || character == '_' || character == '.';
            if (!allowed) return false;
        }

        return true;
    }

    private static bool IsBoundedText(string? value, int maximum)
    {
        if (string.IsNullOrWhiteSpace(value) || value!.Length > maximum) return false;
        foreach (var character in value)
        {
            if (char.IsControl(character)) return false;
        }

        return true;
    }

    private static AgentToolValidationResult Invalid(string code, string repairInstruction)
    {
        return new AgentToolValidationResult
        {
            IsValid = false,
            ErrorCode = code,
            RepairInstruction = repairInstruction,
        };
    }
}
