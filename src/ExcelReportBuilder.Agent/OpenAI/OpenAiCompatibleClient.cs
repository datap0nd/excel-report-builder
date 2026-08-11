using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.Models;
using ExcelReportBuilder.Agent.Protocol;
using ExcelReportBuilder.Agent.Tools;
using ExcelReportBuilder.Agent.Validation;

namespace ExcelReportBuilder.Agent.OpenAI;

public sealed class OpenAiCompatibleClient : IOpenAiCompatibleClient, IDisposable
{
    public const int MaximumResponseBytes = 2 * 1024 * 1024;

    private readonly HttpClient _httpClient;
    private readonly bool _ownsHttpClient;

    public OpenAiCompatibleClient(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        _ownsHttpClient = httpClient == null;

        // Long local inference jobs are cancelled by the caller. A finite
        // HttpClient timeout would be an accidental global job timeout.
        if (_httpClient.Timeout != Timeout.InfiniteTimeSpan)
        {
            _httpClient.Timeout = Timeout.InfiniteTimeSpan;
        }
    }

    public async Task<ModelDiscoveryResult> DiscoverModelsAsync(
        AgentEndpointSettings settings,
        CancellationToken cancellationToken)
    {
        var baseUri = AgentEndpointPolicy.Validate(settings);
        var requestUri = AgentEndpointPolicy.BuildV1Uri(baseUri, "models");
        using (var request = CreateRequest(HttpMethod.Get, requestUri, settings.ApiKey))
        using (var response = await _httpClient.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken).ConfigureAwait(false))
        {
            await EnsureSuccessAsync(response).ConfigureAwait(false);
            var json = await ReadBoundedContentAsync(response, cancellationToken).ConfigureAwait(false);
            try
            {
                using (var document = JsonDocument.Parse(json))
                {
                    JsonElement data;
                    if (!document.RootElement.TryGetProperty("data", out data) || data.ValueKind != JsonValueKind.Array)
                    {
                        throw new AgentEndpointException(
                            "models_response_invalid",
                            "The endpoint returned an invalid models response.",
                            false);
                    }

                    var modelIds = new List<string>();
                    foreach (var item in data.EnumerateArray())
                    {
                        JsonElement id;
                        if (item.ValueKind != JsonValueKind.Object ||
                            !item.TryGetProperty("id", out id) ||
                            id.ValueKind != JsonValueKind.String)
                        {
                            continue;
                        }

                        var modelId = id.GetString();
                        if (!string.IsNullOrWhiteSpace(modelId) && modelId!.Length <= 256 && !modelIds.Contains(modelId, StringComparer.Ordinal))
                        {
                            modelIds.Add(modelId);
                            if (modelIds.Count == 512) break;
                        }
                    }

                    if (modelIds.Count == 0)
                    {
                        throw new AgentEndpointException(
                            "models_empty",
                            "The endpoint did not report any available models.",
                            false);
                    }

                    var selected = modelIds.FirstOrDefault(
                        id => string.Equals(id, settings.Model, StringComparison.Ordinal)) ?? string.Empty;
                    return new ModelDiscoveryResult
                    {
                        ModelIds = modelIds,
                        SelectedModel = selected,
                    };
                }
            }
            catch (JsonException exception)
            {
                throw new AgentEndpointException(
                    "models_response_invalid",
                    "The endpoint returned malformed JSON for model discovery.",
                    false,
                    exception);
            }
        }
    }

    public async Task<AgentModelProposal> RequestToolProposalAsync(
        AgentJobRequest request,
        string? repairInstruction,
        CancellationToken cancellationToken)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        AgentRequestValidator.Validate(request);
        if (repairInstruction != null &&
            repairInstruction.Length > AgentRequestLimits.MaximumWorkflowGuidanceCharacters)
        {
            throw new AgentInputValidationException(
                "workflow_guidance_too_large",
                "The bounded workflow guidance exceeds the supported size.");
        }

        var baseUri = AgentEndpointPolicy.Validate(request.Endpoint);
        var requestUri = AgentEndpointPolicy.BuildV1Uri(baseUri, "chat/completions");
        var body = AgentPromptBuilder.CreateChatCompletionRequest(request, repairInstruction);
        var json = JsonSerializer.Serialize(body, AgentProtocol.JsonOptions);

        using (var message = CreateRequest(HttpMethod.Post, requestUri, request.Endpoint.ApiKey))
        {
            message.Content = new StringContent(json, Encoding.UTF8, "application/json");
            using (var response = await _httpClient.SendAsync(
                message,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken).ConfigureAwait(false))
            {
                await EnsureSuccessAsync(response).ConfigureAwait(false);
                var responseJson = await ReadBoundedContentAsync(response, cancellationToken).ConfigureAwait(false);
                return ParseProposal(responseJson, request.Endpoint.Model);
            }
        }
    }

    public async Task<EndpointProbeResult> CheckToolCallingAsync(
        AgentEndpointSettings settings,
        CancellationToken cancellationToken)
    {
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        ModelDiscoveryResult? discovery = null;
        try
        {
            discovery = await DiscoverModelsAsync(settings, cancellationToken).ConfigureAwait(false);
        }
        catch (AgentEndpointException)
        {
            // Model discovery is optional for OpenAI-compatible endpoints.
            // The two synthetic completion probes below still verify that the
            // configured model is accepted, authenticated, and capable.
        }

        await CheckStructuredOutputAsync(settings, cancellationToken).ConfigureAwait(false);
        var syntheticRequest = CreateSyntheticProbeRequest(settings);
        var proposal = await RequestToolProposalAsync(syntheticRequest, null, cancellationToken).ConfigureAwait(false);
        var validation = AgentToolCallValidator.Validate(proposal.ToolCalls, syntheticRequest.Data);
        if (!validation.IsValid)
        {
            throw new AgentEndpointException(
                "tool_call_check_failed",
                "The endpoint did not return a valid allowlisted tool call for the synthetic check.",
                false);
        }

        return new EndpointProbeResult
        {
            ModelsEndpointAvailable = discovery != null,
            ToolCallingAvailable = true,
            StructuredOutputAvailable = true,
            SelectedModel = !string.IsNullOrEmpty(discovery?.SelectedModel)
                ? discovery!.SelectedModel
                : settings.Model,
            DiscoveredModels = discovery?.ModelIds ?? new List<string>(),
            Summary = discovery == null
                ? "The configured model passed structured-output and synthetic tool-call checks. Model discovery is unavailable but optional."
                : "Model discovery, structured output, and the synthetic tool-call check succeeded.",
        };
    }

    private async Task CheckStructuredOutputAsync(
        AgentEndpointSettings settings,
        CancellationToken cancellationToken)
    {
        var baseUri = AgentEndpointPolicy.Validate(settings);
        var requestUri = AgentEndpointPolicy.BuildV1Uri(baseUri, "chat/completions");
        var body = AgentPromptBuilder.CreateStructuredOutputProbeRequest(settings.Model);
        var json = JsonSerializer.Serialize(body, AgentProtocol.JsonOptions);
        using (var message = CreateRequest(HttpMethod.Post, requestUri, settings.ApiKey))
        {
            message.Content = new StringContent(json, Encoding.UTF8, "application/json");
            using (var response = await _httpClient.SendAsync(
                message,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken).ConfigureAwait(false))
            {
                await EnsureSuccessAsync(response).ConfigureAwait(false);
                var responseJson = await ReadBoundedContentAsync(response, cancellationToken).ConfigureAwait(false);
                ValidateStructuredOutputResponse(responseJson);
            }
        }
    }

    private static void ValidateStructuredOutputResponse(string responseJson)
    {
        try
        {
            using (var document = JsonDocument.Parse(responseJson))
            {
                if (!document.RootElement.TryGetProperty("choices", out var choices) ||
                    choices.ValueKind != JsonValueKind.Array || choices.GetArrayLength() == 0 ||
                    choices[0].ValueKind != JsonValueKind.Object ||
                    !choices[0].TryGetProperty("message", out var message) ||
                    message.ValueKind != JsonValueKind.Object ||
                    !message.TryGetProperty("content", out var content) ||
                    content.ValueKind != JsonValueKind.String)
                {
                    throw StructuredOutputFailure();
                }

                using (var structured = JsonDocument.Parse(content.GetString() ?? string.Empty))
                {
                    var root = structured.RootElement;
                    if (root.ValueKind != JsonValueKind.Object || root.GetRawText().Length > 256 ||
                        root.EnumerateObject().Count() != 2 ||
                        !root.TryGetProperty("status", out var status) || status.ValueKind != JsonValueKind.String || status.GetString() != "ok" ||
                        !root.TryGetProperty("capability", out var capability) || capability.ValueKind != JsonValueKind.String || capability.GetString() != "structured_output")
                    {
                        throw StructuredOutputFailure();
                    }
                }
            }
        }
        catch (JsonException exception)
        {
            throw new AgentEndpointException(
                "structured_output_check_failed",
                "The endpoint did not return valid structured output for the synthetic check.",
                false,
                exception);
        }
    }

    private static AgentEndpointException StructuredOutputFailure()
    {
        return new AgentEndpointException(
            "structured_output_check_failed",
            "The endpoint did not return valid structured output for the synthetic check.",
            false);
    }

    public void Dispose()
    {
        if (_ownsHttpClient)
        {
            _httpClient.Dispose();
        }
    }

    private static HttpRequestMessage CreateRequest(HttpMethod method, Uri uri, string? apiKey)
    {
        var request = new HttpRequestMessage(method, uri);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (!string.IsNullOrWhiteSpace(apiKey))
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        }

        return request;
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response)
    {
        if (response.IsSuccessStatusCode) return;

        // Do not include endpoint response bodies. Local servers can echo
        // prompts, data, credentials, or implementation details in errors.
        await Task.CompletedTask.ConfigureAwait(false);
        var status = (int)response.StatusCode;
        throw new AgentEndpointException(
            "endpoint_http_error",
            "The AI endpoint returned HTTP " + status + " (" + SafeReason(response.StatusCode) + ").",
            IsRetryable(response.StatusCode));
    }

    private static async Task<string> ReadBoundedContentAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        if (response.Content.Headers.ContentLength.HasValue &&
            response.Content.Headers.ContentLength.Value > MaximumResponseBytes)
        {
            throw new AgentEndpointException(
                "endpoint_response_too_large",
                "The AI endpoint response exceeded the supported size.",
                false);
        }

        using (var source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
        using (var destination = new MemoryStream())
        {
            var buffer = new byte[8192];
            var total = 0;
            while (true)
            {
                var read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                if (read == 0) break;
                total += read;
                if (total > MaximumResponseBytes)
                {
                    throw new AgentEndpointException(
                        "endpoint_response_too_large",
                        "The AI endpoint response exceeded the supported size.",
                        false);
                }

                destination.Write(buffer, 0, read);
            }

            return Encoding.UTF8.GetString(destination.ToArray());
        }
    }

    private static AgentModelProposal ParseProposal(string responseJson, string configuredModel)
    {
        try
        {
            using (var document = JsonDocument.Parse(responseJson))
            {
                JsonElement choices;
                if (!document.RootElement.TryGetProperty("choices", out choices) ||
                    choices.ValueKind != JsonValueKind.Array ||
                    choices.GetArrayLength() == 0)
                {
                    throw InvalidCompletion();
                }

                var choice = choices[0];
                JsonElement message;
                if (choice.ValueKind != JsonValueKind.Object ||
                    !choice.TryGetProperty("message", out message) ||
                    message.ValueKind != JsonValueKind.Object)
                {
                    throw InvalidCompletion();
                }

                var proposal = new AgentModelProposal { Model = configuredModel };
                JsonElement model;
                if (document.RootElement.TryGetProperty("model", out model) && model.ValueKind == JsonValueKind.String)
                {
                    proposal.Model = model.GetString() ?? configuredModel;
                }

                JsonElement content;
                if (message.TryGetProperty("content", out content) && content.ValueKind == JsonValueKind.String)
                {
                    var assistantText = content.GetString();
                    proposal.AssistantText = assistantText != null && assistantText.Length > 4096
                        ? assistantText.Substring(0, 4096)
                        : assistantText;
                }

                JsonElement toolCalls;
                if (!message.TryGetProperty("tool_calls", out toolCalls) || toolCalls.ValueKind != JsonValueKind.Array)
                {
                    return proposal;
                }

                foreach (var call in toolCalls.EnumerateArray())
                {
                    JsonElement id;
                    JsonElement function;
                    JsonElement name;
                    JsonElement arguments;
                    if (call.ValueKind != JsonValueKind.Object ||
                        !call.TryGetProperty("id", out id) || id.ValueKind != JsonValueKind.String ||
                        !call.TryGetProperty("function", out function) || function.ValueKind != JsonValueKind.Object ||
                        !function.TryGetProperty("name", out name) || name.ValueKind != JsonValueKind.String ||
                        !function.TryGetProperty("arguments", out arguments) || arguments.ValueKind != JsonValueKind.String)
                    {
                        proposal.ToolCalls.Add(new AgentToolCall());
                        continue;
                    }

                    proposal.ToolCalls.Add(new AgentToolCall
                    {
                        Id = id.GetString() ?? string.Empty,
                        Name = name.GetString() ?? string.Empty,
                        ArgumentsJson = arguments.GetString() ?? string.Empty,
                    });
                }

                return proposal;
            }
        }
        catch (JsonException exception)
        {
            throw new AgentEndpointException(
                "completion_response_invalid",
                "The endpoint returned malformed JSON for a tool proposal.",
                false,
                exception);
        }
    }

    private static AgentEndpointException InvalidCompletion()
    {
        return new AgentEndpointException(
            "completion_response_invalid",
            "The endpoint returned an invalid chat completion response.",
            false);
    }

    private static AgentJobRequest CreateSyntheticProbeRequest(AgentEndpointSettings settings)
    {
        return new AgentJobRequest
        {
            JobId = "synthetic-endpoint-check",
            WorkbookId = "synthetic-workbook-check",
            UserPrompt = "Confirm that the bounded synthetic data needs no transforms by calling propose_transforms.",
            Endpoint = new AgentEndpointSettings
            {
                BaseUrl = settings.BaseUrl,
                Model = string.IsNullOrWhiteSpace(settings.Model) ? AgentDefaults.Model : settings.Model,
                ApiKey = settings.ApiKey,
                AllowRemoteHttp = settings.AllowRemoteHttp,
            },
            MaxRepairCycles = 0,
            Data = new AgentDataSnapshot
            {
                SourceDisplayName = "Synthetic data",
                RowCount = 2,
                ReportingYear = 2026,
                Fields = new List<AgentField>
                {
                    new AgentField { Name = "Category", Type = AgentFieldType.Text },
                    new AgentField { Name = "Amount", Type = AgentFieldType.Number },
                },
                SampleRows = new List<AgentSampleRow>
                {
                    new AgentSampleRow
                    {
                        Values = new List<AgentSampleValue>
                        {
                            new AgentSampleValue { Field = "Category", Value = "A" },
                            new AgentSampleValue { Field = "Amount", Value = "10" },
                        },
                    },
                    new AgentSampleRow
                    {
                        Values = new List<AgentSampleValue>
                        {
                            new AgentSampleValue { Field = "Category", Value = "B" },
                            new AgentSampleValue { Field = "Amount", Value = "20" },
                        },
                    },
                },
            },
        };
    }

    private static bool IsRetryable(HttpStatusCode statusCode)
    {
        var status = (int)statusCode;
        return status == 408 || status == 429 || status >= 500;
    }

    private static string SafeReason(HttpStatusCode statusCode)
    {
        switch (statusCode)
        {
            case HttpStatusCode.BadRequest: return "Bad Request";
            case HttpStatusCode.Unauthorized: return "Unauthorized";
            case HttpStatusCode.Forbidden: return "Forbidden";
            case HttpStatusCode.NotFound: return "Not Found";
            case HttpStatusCode.RequestTimeout: return "Request Timeout";
            case (HttpStatusCode)429: return "Too Many Requests";
            case HttpStatusCode.InternalServerError: return "Internal Server Error";
            case HttpStatusCode.BadGateway: return "Bad Gateway";
            case HttpStatusCode.ServiceUnavailable: return "Service Unavailable";
            case HttpStatusCode.GatewayTimeout: return "Gateway Timeout";
            default: return "Endpoint Error";
        }
    }
}

public sealed class AgentEndpointException : Exception
{
    public AgentEndpointException(
        string code,
        string message,
        bool retryable,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        Retryable = retryable;
    }

    public string Code { get; }

    public bool Retryable { get; }
}
