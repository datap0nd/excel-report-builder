using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;

namespace ExcelReportBuilder.Agent.Validation;

public static class AgentRequestLimits
{
    public const int MaximumPromptCharacters = 16 * 1024;
    public const int MaximumDataPayloadBytes = 512 * 1024;
    public const int MaximumFields = 256;
    public const int MaximumSampleRows = 50;
    public const int MaximumSampleCells = 10_000;
    public const int MaximumFieldNameCharacters = 128;
    public const int MaximumSampleValueCharacters = 1024;
    public const int MaximumSourceNameCharacters = 128;
    public const int MaximumCurrentSpecificationBytes = 256 * 1024;
    public const int MaximumEndpointCharacters = 2048;
    public const int MaximumModelCharacters = 256;
    public const int MaximumApiKeyCharacters = 8192;
    public const int MaximumWorkflowGuidanceCharacters = 8192;
}

public static class AgentRequestValidator
{
    public static void Validate(AgentJobRequest request)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateIdentity(request);

        if (string.IsNullOrWhiteSpace(request.UserPrompt))
        {
            throw new AgentInputValidationException("prompt_required", "A report request is required.");
        }

        if (request.UserPrompt.Length > AgentRequestLimits.MaximumPromptCharacters)
        {
            throw new AgentInputValidationException(
                "prompt_too_large",
                "The report request exceeds the supported prompt size.");
        }

        if (request.MaxRepairCycles < 0 || request.MaxRepairCycles > AgentDefaults.MaximumAllowedRepairCycles)
        {
            throw new AgentInputValidationException(
                "repair_cycle_limit_invalid",
                "The repair-cycle limit is outside the supported range.");
        }

        if (request.Data == null)
        {
            throw new AgentInputValidationException("data_required", "A bounded data description is required.");
        }

        if (request.Endpoint == null)
        {
            throw new AgentInputValidationException("endpoint_required", "AI endpoint settings are required.");
        }

        ValidateEndpointSettings(request.Endpoint);
        Uri endpointUri = AgentEndpointPolicy.Validate(request.Endpoint);
        if (!endpointUri.IsLoopback && !request.Endpoint.AllowRemoteWorkbookData)
        {
            throw new AgentInputValidationException(
                "remote_workbook_data_consent_required",
                "A remote model endpoint can receive workbook column names and bounded sample rows only after explicit consent.");
        }

        ValidateData(request.Data);
        ValidateCurrentSpecification(request.CurrentSpecification ?? new AgentSpecificationSnapshot(), request.Data.Fields);
    }

    public static void ValidateIdentity(AgentJobRequest request)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateIdentifier(request.JobId, "job_id_invalid", "The job ID is required and must be 128 characters or fewer.");
        ValidateIdentifier(request.WorkbookId, "workbook_id_invalid", "The workbook ID is required and must be 128 characters or fewer.");
        if (request.ResumeFromCheckpointId != null)
        {
            ValidateIdentifier(
                request.ResumeFromCheckpointId,
                "checkpoint_id_invalid",
                "The resume checkpoint ID must be 128 characters or fewer.");
        }
    }

    private static void ValidateData(AgentDataSnapshot data)
    {
        if (data.RowCount < 0)
        {
            throw new AgentInputValidationException("row_count_invalid", "The source row count cannot be negative.");
        }

        if (string.IsNullOrWhiteSpace(data.SourceDisplayName) ||
            data.SourceDisplayName.Length > AgentRequestLimits.MaximumSourceNameCharacters)
        {
            throw new AgentInputValidationException("source_name_invalid", "The source display name is invalid.");
        }

        if (data.ReportingYear.HasValue && (data.ReportingYear.Value < 1900 || data.ReportingYear.Value > 9999))
        {
            throw new AgentInputValidationException("reporting_year_invalid", "The reporting year is invalid.");
        }

        if (data.Fields == null || data.Fields.Count == 0)
        {
            throw new AgentInputValidationException("fields_required", "At least one source column is required.");
        }

        if (data.Fields.Count > AgentRequestLimits.MaximumFields)
        {
            throw new AgentInputValidationException("too_many_fields", "The source description contains too many columns.");
        }

        var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in data.Fields)
        {
            if (field == null || string.IsNullOrWhiteSpace(field.Name) ||
                field.Name.Length > AgentRequestLimits.MaximumFieldNameCharacters ||
                ContainsControlCharacter(field.Name))
            {
                throw new AgentInputValidationException("field_name_invalid", "A source column name is invalid.");
            }

            if (!fields.Add(field.Name))
            {
                throw new AgentInputValidationException("field_name_duplicate", "Source column names must be unique.");
            }

            if (!Enum.IsDefined(typeof(AgentFieldType), field.Type))
            {
                throw new AgentInputValidationException("field_type_invalid", "A source column type is invalid.");
            }
        }

        if (data.SampleRows == null)
        {
            throw new AgentInputValidationException("sample_rows_invalid", "The sample row collection is missing.");
        }

        if (data.SampleRows.Count > AgentRequestLimits.MaximumSampleRows)
        {
            throw new AgentInputValidationException("too_many_sample_rows", "The data description contains too many sample rows.");
        }

        var cellCount = 0;
        foreach (var row in data.SampleRows)
        {
            if (row == null || row.Values == null)
            {
                throw new AgentInputValidationException("sample_row_invalid", "A sample row is invalid.");
            }

            var rowFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in row.Values)
            {
                cellCount++;
                if (cellCount > AgentRequestLimits.MaximumSampleCells)
                {
                    throw new AgentInputValidationException("too_many_sample_cells", "The data description contains too many sample values.");
                }

                if (value == null || !fields.Contains(value.Field) || !rowFields.Add(value.Field))
                {
                    throw new AgentInputValidationException("sample_field_invalid", "A sample value references an invalid source column.");
                }

                if (value.Value != null && value.Value.Length > AgentRequestLimits.MaximumSampleValueCharacters)
                {
                    throw new AgentInputValidationException("sample_value_too_large", "A sample value exceeds the supported size.");
                }
            }
        }

        var json = JsonSerializer.Serialize(data, AgentProtocol.JsonOptions);
        if (Encoding.UTF8.GetByteCount(json) > AgentRequestLimits.MaximumDataPayloadBytes)
        {
            throw new AgentInputValidationException("data_payload_too_large", "The bounded data description exceeds the supported size.");
        }
    }

    private static void ValidateCurrentSpecification(
        AgentSpecificationSnapshot specification,
        IEnumerable<AgentField> fields)
    {
        if (specification.CanonicalReportSpecJson == null)
        {
            throw new AgentInputValidationException(
                "specification_invalid",
                "The current canonical report setup is invalid.");
        }

        if (specification.CanonicalReportSpecJson.Length != 0)
        {
            try
            {
                using (var document = JsonDocument.Parse(specification.CanonicalReportSpecJson))
                {
                    if (document.RootElement.ValueKind != JsonValueKind.Object ||
                        !document.RootElement.TryGetProperty("schemaVersion", out var version) ||
                        version.ValueKind != JsonValueKind.String ||
                        !string.Equals(version.GetString(), "1.0", StringComparison.Ordinal))
                    {
                        throw new AgentInputValidationException(
                            "specification_invalid",
                            "The current canonical report setup is not a supported versioned object.");
                    }
                }
            }
            catch (JsonException)
            {
                throw new AgentInputValidationException(
                    "specification_invalid",
                    "The current canonical report setup is not valid JSON.");
            }
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fields)
        {
            names.Add(field.Name);
        }

        ValidateFieldList(specification.Rows, names, 32);
        ValidateFieldList(specification.Columns, names, 16);

        if (specification.Values == null || specification.Filters == null)
        {
            throw new AgentInputValidationException("specification_invalid", "The current report setup is invalid.");
        }

        if (specification.Values.Count > 32 || specification.Filters.Count > 32)
        {
            throw new AgentInputValidationException("specification_too_large", "The current report setup exceeds the supported size.");
        }

        foreach (var value in specification.Values)
        {
            if (value == null || !names.Contains(value.Field) || !IsAggregation(value.Aggregation))
            {
                throw new AgentInputValidationException("specification_invalid", "The current Values setup is invalid.");
            }
        }

        foreach (var filter in specification.Filters)
        {
            if (filter == null || !names.Contains(filter.Field) || filter.Values == null || filter.Values.Count > 50)
            {
                throw new AgentInputValidationException("specification_invalid", "The current Filters setup is invalid.");
            }


            foreach (var filterValue in filter.Values)
            {
                if (filterValue == null || filterValue.Length > AgentRequestLimits.MaximumSampleValueCharacters ||
                    ContainsControlCharacter(filterValue))
                {
                    throw new AgentInputValidationException("specification_invalid", "A current filter value is invalid.");
                }
            }
        }

        var specificationJson = JsonSerializer.Serialize(specification, AgentProtocol.JsonOptions);
        if (Encoding.UTF8.GetByteCount(specificationJson) > AgentRequestLimits.MaximumCurrentSpecificationBytes)
        {
            throw new AgentInputValidationException("specification_too_large", "The current report setup exceeds the supported size.");
        }
    }

    private static void ValidateFieldList(List<string>? values, HashSet<string> fieldNames, int maximumCount)
    {
        if (values == null || values.Count > maximumCount)
        {
            throw new AgentInputValidationException("specification_invalid", "The current report setup is invalid.");
        }

        foreach (var value in values)
        {
            if (!fieldNames.Contains(value))
            {
                throw new AgentInputValidationException("specification_invalid", "The current report setup references an unknown column.");
            }
        }
    }

    private static void ValidateEndpointSettings(AgentEndpointSettings endpoint)
    {
        if (endpoint.BaseUrl == null || endpoint.BaseUrl.Length > AgentRequestLimits.MaximumEndpointCharacters ||
            endpoint.Model == null ||
            endpoint.Model.Length > AgentRequestLimits.MaximumModelCharacters || ContainsControlCharacter(endpoint.Model) ||
            (endpoint.ApiKey != null && endpoint.ApiKey.Length > AgentRequestLimits.MaximumApiKeyCharacters))
        {
            throw new AgentInputValidationException("endpoint_settings_invalid", "The AI endpoint settings exceed supported limits.");
        }
    }

    private static bool IsAggregation(string aggregation)
    {
        return string.Equals(aggregation, "sum", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(aggregation, "count", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(aggregation, "min", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(aggregation, "max", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(aggregation, "average", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsControlCharacter(string value)
    {
        foreach (var character in value)
        {
            if (char.IsControl(character)) return true;
        }

        return false;
    }

    private static void ValidateIdentifier(string value, string code, string message)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 128 || ContainsControlCharacter(value))
        {
            throw new AgentInputValidationException(code, message);
        }

        foreach (var character in value)
        {
            var allowed = (character >= 'a' && character <= 'z') ||
                          (character >= 'A' && character <= 'Z') ||
                          (character >= '0' && character <= '9') ||
                          character == '-' || character == '_' || character == '.';
            if (!allowed)
            {
                throw new AgentInputValidationException(code, message);
            }
        }
    }
}

public sealed class AgentInputValidationException : Exception
{
    public AgentInputValidationException(string code, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
    }

    public string Code { get; }
}
