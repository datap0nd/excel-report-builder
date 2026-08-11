using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class EndpointPolicyTests
{
    [Theory]
    [InlineData("http://localhost:1234")]
    [InlineData("http://127.0.0.1:1234")]
    [InlineData("http://[::1]:1234")]
    [InlineData("https://models.example.test")]
    public void Validate_AllowsLoopbackHttpAndAnyHttps(string url)
    {
        var result = AgentEndpointPolicy.Validate(new AgentEndpointSettings { BaseUrl = url });

        Assert.Equal(new Uri(url), result);
    }

    [Fact]
    public void Validate_BlocksRemoteHttpByDefault()
    {
        var settings = new AgentEndpointSettings { BaseUrl = "http://models.example.test" };

        var error = Assert.Throws<AgentEndpointPolicyException>(() => AgentEndpointPolicy.Validate(settings));

        Assert.Contains("explicitly allow", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Validate_AllowsRemoteHttpOnlyWithExplicitOptIn()
    {
        var settings = new AgentEndpointSettings
        {
            BaseUrl = "http://models.example.test",
            AllowRemoteHttp = true,
        };

        Assert.Equal("models.example.test", AgentEndpointPolicy.Validate(settings).Host);
    }

    [Theory]
    [InlineData("https://user:secret@models.example.test")]
    [InlineData("https://models.example.test?api_key=secret")]
    [InlineData("file:///tmp/model")]
    public void Validate_RejectsUnsafeEndpointForms(string url)
    {
        Assert.Throws<AgentEndpointPolicyException>(() =>
            AgentEndpointPolicy.Validate(new AgentEndpointSettings { BaseUrl = url }));
    }

    [Theory]
    [InlineData("http://127.0.0.1:1234", "http://127.0.0.1:1234/v1/models")]
    [InlineData("http://127.0.0.1:1234/v1/", "http://127.0.0.1:1234/v1/models")]
    public void BuildV1Uri_DoesNotDuplicateVersionPath(string root, string expected)
    {
        Assert.Equal(expected, AgentEndpointPolicy.BuildV1Uri(new Uri(root), "models").AbsoluteUri);
    }

    [Fact]
    public void DefaultModel_IsQwen35()
    {
        Assert.Equal("qwen3.5-35b-a3b", new AgentEndpointSettings().Model);
    }

    [Fact]
    public void Validate_RestoresDefaultWhenPersistedModelIsBlank()
    {
        var settings = new AgentEndpointSettings
        {
            BaseUrl = "http://127.0.0.1:1234",
            Model = string.Empty,
        };

        AgentEndpointPolicy.Validate(settings);

        Assert.Equal(AgentDefaults.Model, settings.Model);
    }
}
