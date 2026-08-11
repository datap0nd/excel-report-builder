using System.Text;
using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.Agent.Tests;

public sealed class SettingsTests
{
    [Fact]
    public void Materializer_KeepsPlaintextOutOfPersistedSettings()
    {
        var protector = new ReversingProtector();
        var testSecret = string.Concat("synthetic", "-secret");
        var runtime = new AgentEndpointSettings
        {
            BaseUrl = "https://models.example.test",
            Model = "synthetic-model",
            ApiKey = testSecret,
            AllowRemoteHttp = true,
            AllowRemoteWorkbookData = true,
        };

        var persisted = AgentSettingsMaterializer.Protect(runtime, protector);

        Assert.NotNull(persisted.ProtectedApiKey);
        Assert.DoesNotContain(testSecret, persisted.ProtectedApiKey, StringComparison.Ordinal);
        var restored = AgentSettingsMaterializer.Unprotect(persisted, protector);
        Assert.Equal(runtime.ApiKey, restored.ApiKey);
        Assert.Equal(runtime.BaseUrl, restored.BaseUrl);
        Assert.True(restored.AllowRemoteHttp);
        Assert.True(restored.AllowRemoteWorkbookData);
    }

    private sealed class ReversingProtector : IAgentSecretProtector
    {
        public byte[] Protect(byte[] clearText) => clearText.Reverse().ToArray();

        public byte[] Unprotect(byte[] protectedData) => protectedData.Reverse().ToArray();
    }
}
