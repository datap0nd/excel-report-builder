using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class EndpointCredentialScopeTests
{
    [Fact]
    public void Matches_AllowsCaseInsensitiveOriginAndTrailingSlash()
    {
        Assert.True(AgentEndpointCredentialScope.Matches(
            "https://MODELS.example.test:443/v1/",
            "https://models.example.test/v1"));
    }

    [Fact]
    public void Matches_RejectsCaseChangedPath()
    {
        Assert.False(AgentEndpointCredentialScope.Matches(
            "https://models.example.test/Team-A/v1",
            "https://models.example.test/team-a/v1"));
    }

    [Theory]
    [InlineData("https://models.example.test/v1", "https://other.example.test/v1")]
    [InlineData("https://models.example.test/v1", "https://models.example.test:8443/v1")]
    [InlineData("https://models.example.test/v1", "http://models.example.test/v1")]
    public void Matches_RejectsDifferentOrigin(string saved, string requested)
    {
        Assert.False(AgentEndpointCredentialScope.Matches(saved, requested));
    }
}
