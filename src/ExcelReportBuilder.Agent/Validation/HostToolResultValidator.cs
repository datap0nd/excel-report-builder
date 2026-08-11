using System;
using System.Text;
using System.Text.Json;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Validation;

public static class HostToolResultValidator
{
    public const int MaximumResultBytes = 64 * 1024;
    public const int MaximumCheckFailures = 50;

    public static void Validate(HostToolRequestEvent request, HostToolResultRequest result)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (result == null) throw new AgentInputValidationException("host_tool_result_missing", "The deterministic host returned no tool result.");
        if (!string.Equals(request.JobId, result.JobId, StringComparison.Ordinal) ||
            !string.Equals(request.ToolCallId, result.ToolCallId, StringComparison.Ordinal))
        {
            throw new AgentInputValidationException("host_tool_result_mismatch", "The deterministic host tool result did not match the pending request.");
        }

        if (!IsIdentifier(result.OutcomeCode))
        {
            throw new AgentInputValidationException("host_tool_outcome_invalid", "The deterministic host returned an invalid outcome code.");
        }

        if (string.IsNullOrWhiteSpace(result.ResultJson) || Encoding.UTF8.GetByteCount(result.ResultJson) > MaximumResultBytes)
        {
            throw new AgentInputValidationException("host_tool_result_too_large", "The deterministic host result exceeded the supported size.");
        }

        try
        {
            using (var document = JsonDocument.Parse(result.ResultJson))
            {
                if (document.RootElement.ValueKind != JsonValueKind.Object)
                {
                    throw new AgentInputValidationException("host_tool_result_invalid", "The deterministic host result must be a JSON object.");
                }
            }
        }
        catch (JsonException exception)
        {
            throw new AgentInputValidationException("host_tool_result_invalid", "The deterministic host result is not valid JSON.", exception);
        }

        if (result.CheckFailures == null || result.CheckFailures.Count > MaximumCheckFailures)
        {
            throw new AgentInputValidationException("host_check_failures_invalid", "The deterministic host returned an invalid check-failure collection.");
        }

        foreach (var failure in result.CheckFailures)
        {
            if (failure == null || !IsIdentifier(failure.Code) || !IsMessage(failure.Message))
            {
                throw new AgentInputValidationException("host_check_failure_invalid", "The deterministic host returned an invalid check failure.");
            }
        }
    }

    private static bool IsIdentifier(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value!.Length > 128) return false;
        foreach (var character in value)
        {
            var allowed = (character >= 'a' && character <= 'z') ||
                          (character >= 'A' && character <= 'Z') ||
                          (character >= '0' && character <= '9') ||
                          character == '_' || character == '-' || character == '.';
            if (!allowed) return false;
        }

        return true;
    }

    private static bool IsMessage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value!.Length > 256) return false;
        foreach (var character in value)
        {
            if (char.IsControl(character)) return false;
        }

        return true;
    }
}
