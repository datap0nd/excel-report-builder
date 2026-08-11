using ExcelReportBuilder.Agent.Diagnostics;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class DiagnosticRedactorTests
{
    [Fact]
    public void Redact_RemovesBearerSecretsCredentialsQueriesAndPrompt()
    {
        var prompt = "show the synthetic confidential request";
        var input =
            "Authorization: Bearer synthetic-token " +
            "api_key=another-secret " +
            "https://user:password@models.example.test/v1?token=query-secret " + prompt;

        var result = DiagnosticRedactor.Redact(input, new[] { prompt });

        Assert.DoesNotContain("synthetic-token", result, StringComparison.Ordinal);
        Assert.DoesNotContain("another-secret", result, StringComparison.Ordinal);
        Assert.DoesNotContain("password", result, StringComparison.Ordinal);
        Assert.DoesNotContain("query-secret", result, StringComparison.Ordinal);
        Assert.DoesNotContain(prompt, result, StringComparison.Ordinal);
        Assert.Contains("[redacted]", result, StringComparison.Ordinal);
    }

    [Fact]
    public void Redact_CapsDiagnosticLengthAndRemovesNewlines()
    {
        var result = DiagnosticRedactor.Redact(new string('x', 1000) + "\r\nnext");

        Assert.True(result.Length <= 515);
        Assert.DoesNotContain('\r', result);
        Assert.DoesNotContain('\n', result);
    }
}
