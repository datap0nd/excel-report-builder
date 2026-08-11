using System.Net;
using System.Net.Http.Headers;
using System.Text;
using ExcelReportBuilder.Agent.Configuration;
using ExcelReportBuilder.Agent.OpenAI;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class OpenAiCompatibleClientTests
{
    [Fact]
    public async Task DiscoverModels_UsesV1ModelsAndSelectsConfiguredModel()
    {
        var handler = new SyntheticEndpointHandler();
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);
        var settings = SyntheticJob.Create().Endpoint;

        var result = await client.DiscoverModelsAsync(settings, CancellationToken.None);

        Assert.Equal("/v1/models", handler.Requests.Single().PathAndQuery);
        Assert.Contains(AgentDefaults.Model, result.ModelIds);
        Assert.Equal(AgentDefaults.Model, result.SelectedModel);
        Assert.Equal(Timeout.InfiniteTimeSpan, httpClient.Timeout);
    }

    [Fact]
    public async Task SyntheticCheck_UsesSyntheticDataAndRequiresValidToolCall()
    {
        var handler = new SyntheticEndpointHandler();
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);
        var settings = SyntheticJob.Create().Endpoint;
        var testApiKey = string.Concat("synthetic", "-api-key");
        settings.ApiKey = testApiKey;

        var result = await client.CheckToolCallingAsync(settings, CancellationToken.None);

        Assert.True(result.ModelsEndpointAvailable);
        Assert.True(result.ToolCallingAvailable);
        Assert.True(result.StructuredOutputAvailable);
        Assert.Equal(3, handler.Requests.Count);
        Assert.Equal("/v1/chat/completions", handler.Requests[2].PathAndQuery);
        Assert.Contains("Synthetic data", handler.Bodies.Last(), StringComparison.Ordinal);
        Assert.Contains("propose_report_spec", handler.Bodies.Last(), StringComparison.Ordinal);
        Assert.Equal("Bearer", handler.AuthorizationSchemes.Last());
        Assert.Equal(testApiKey, handler.AuthorizationParameters.Last());
    }

    [Fact]
    public async Task SyntheticCheck_AllowsAnEndpointWithoutModelDiscovery()
    {
        var handler = new SyntheticEndpointHandler(modelsUnavailable: true);
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);

        var result = await client.CheckToolCallingAsync(
            SyntheticJob.Create().Endpoint,
            CancellationToken.None);

        Assert.False(result.ModelsEndpointAvailable);
        Assert.True(result.StructuredOutputAvailable);
        Assert.True(result.ToolCallingAvailable);
        Assert.Equal(AgentDefaults.Model, result.SelectedModel);
        Assert.Contains("optional", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(3, handler.Requests.Count);
    }

    [Theory]
    [InlineData("qwen3.5-35b-a3b")]
    [InlineData("gemma-3-27b-it")]
    public async Task SyntheticCheck_AcceptsSupportedOpenAiCompatibleModelFamilies(string modelId)
    {
        var handler = new SyntheticEndpointHandler(modelId: modelId);
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);
        var settings = SyntheticJob.Create().Endpoint;
        settings.Model = modelId;

        var result = await client.CheckToolCallingAsync(settings, CancellationToken.None);

        Assert.Equal(modelId, result.SelectedModel);
        Assert.True(result.StructuredOutputAvailable);
        Assert.True(result.ToolCallingAvailable);
    }

    [Fact]
    public async Task HttpFailure_DoesNotEchoEndpointBody()
    {
        var responseMarker = string.Concat("synthetic-response", "-marker");
        var handler = new SyntheticEndpointHandler(HttpStatusCode.BadRequest, responseMarker);
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);

        var error = await Assert.ThrowsAsync<AgentEndpointException>(() =>
            client.DiscoverModelsAsync(SyntheticJob.Create().Endpoint, CancellationToken.None));

        Assert.DoesNotContain(responseMarker, error.Message, StringComparison.Ordinal);
        Assert.Equal("endpoint_http_error", error.Code);
    }

    [Fact]
    public async Task SyntheticCheck_FailsWhenStructuredOutputIsNotExact()
    {
        var handler = new SyntheticEndpointHandler(invalidStructuredOutput: true);
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);

        var error = await Assert.ThrowsAsync<AgentEndpointException>(() =>
            client.CheckToolCallingAsync(SyntheticJob.Create().Endpoint, CancellationToken.None));

        Assert.Equal("structured_output_check_failed", error.Code);
        Assert.Equal(2, handler.Requests.Count);
    }

    [Fact]
    public async Task SyntheticCheck_FailsWhenEndpointReturnsUnlistedTool()
    {
        var handler = new SyntheticEndpointHandler(invalidToolCall: true);
        using var httpClient = new HttpClient(handler);
        using var client = new OpenAiCompatibleClient(httpClient);

        var error = await Assert.ThrowsAsync<AgentEndpointException>(() =>
            client.CheckToolCallingAsync(SyntheticJob.Create().Endpoint, CancellationToken.None));

        Assert.Equal("tool_call_check_failed", error.Code);
        Assert.Equal(3, handler.Requests.Count);
    }

    private sealed class SyntheticEndpointHandler : HttpMessageHandler
    {
        private readonly HttpStatusCode _status;
        private readonly string? _failureBody;
        private readonly bool _invalidStructuredOutput;
        private readonly bool _invalidToolCall;
        private readonly bool _modelsUnavailable;
        private readonly string _modelId;

        public SyntheticEndpointHandler(
            HttpStatusCode status = HttpStatusCode.OK,
            string? failureBody = null,
            bool invalidStructuredOutput = false,
            bool invalidToolCall = false,
            bool modelsUnavailable = false,
            string modelId = "qwen3.5-35b-a3b")
        {
            _status = status;
            _failureBody = failureBody;
            _invalidStructuredOutput = invalidStructuredOutput;
            _invalidToolCall = invalidToolCall;
            _modelsUnavailable = modelsUnavailable;
            _modelId = modelId;
        }

        public List<Uri> Requests { get; } = new();

        public List<string> Bodies { get; } = new();

        public List<string?> AuthorizationSchemes { get; } = new();

        public List<string?> AuthorizationParameters { get; } = new();

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            Requests.Add(request.RequestUri!);
            AuthorizationSchemes.Add(request.Headers.Authorization?.Scheme);
            AuthorizationParameters.Add(request.Headers.Authorization?.Parameter);
            if (request.Content != null)
            {
                Bodies.Add(await request.Content.ReadAsStringAsync(cancellationToken));
            }

            if (_status != HttpStatusCode.OK)
            {
                return new HttpResponseMessage(_status)
                {
                    Content = new StringContent(_failureBody ?? string.Empty, Encoding.UTF8, "application/json"),
                };
            }

            if (request.RequestUri!.AbsolutePath.EndsWith("/models", StringComparison.Ordinal))
            {
                if (_modelsUnavailable)
                {
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }

                return Json("{\"data\":[{\"id\":\"" + _modelId + "\"},{\"id\":\"synthetic-secondary\"}]}");
            }

            var requestBody = Bodies.LastOrDefault() ?? string.Empty;
            if (requestBody.Contains("response_format", StringComparison.Ordinal))
            {
                var structuredContent = _invalidStructuredOutput
                    ? "{\\\"status\\\":\\\"maybe\\\",\\\"capability\\\":\\\"structured_output\\\"}"
                    : "{\\\"status\\\":\\\"ok\\\",\\\"capability\\\":\\\"structured_output\\\"}";
                return Json(
                    "{\"model\":\"" + _modelId + "\",\"choices\":[{\"message\":{" +
                    "\"content\":\"" + structuredContent + "\"}}]}");
            }

            var toolName = _invalidToolCall ? "read_file" : "propose_transforms";
            return Json(
                "{\"model\":\"" + _modelId + "\",\"choices\":[{\"message\":{" +
                "\"content\":null,\"tool_calls\":[{\"id\":\"probe-call\",\"type\":\"function\",\"function\":{" +
                "\"name\":\"" + toolName + "\",\"arguments\":\"{\\\"transforms\\\":[]}\"}}]}}]}");
        }

        private static HttpResponseMessage Json(string json)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json"),
            };
        }
    }
}
