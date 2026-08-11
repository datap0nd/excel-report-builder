using System;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelReportBuilder.Agent.Configuration;

public static class AgentDefaults
{
    public const string Model = "qwen3.5-35b-a3b";
    public const int MaxRepairCycles = 3;
    public const int MaximumAllowedRepairCycles = 5;
}

/// <summary>
/// Runtime settings after a host-owned secret store has decrypted the API key.
/// </summary>
public sealed class AgentEndpointSettings
{
    public string BaseUrl { get; set; } = "http://127.0.0.1:1234";

    public string Model { get; set; } = AgentDefaults.Model;

    public string? ApiKey { get; set; }

    public bool AllowRemoteHttp { get; set; }

    public bool AllowRemoteWorkbookData { get; set; }
}

/// <summary>
/// Persistable settings. The protected secret is opaque to this library so a
/// Windows host can use DPAPI without introducing a platform dependency here.
/// </summary>
public sealed class PersistedAgentSettings
{
    public string BaseUrl { get; set; } = "http://127.0.0.1:1234";

    public string Model { get; set; } = AgentDefaults.Model;

    public bool AllowRemoteHttp { get; set; }

    public bool AllowRemoteWorkbookData { get; set; }

    public string? ProtectedApiKey { get; set; }
}

public interface IAgentSecretProtector
{
    byte[] Protect(byte[] clearText);

    byte[] Unprotect(byte[] protectedData);
}

public interface IAgentSettingsStore
{
    Task<PersistedAgentSettings?> LoadAsync(CancellationToken cancellationToken);

    Task SaveAsync(PersistedAgentSettings settings, CancellationToken cancellationToken);
}

public static class AgentSettingsMaterializer
{
    public static PersistedAgentSettings Protect(
        AgentEndpointSettings settings,
        IAgentSecretProtector protector)
    {
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        if (protector == null) throw new ArgumentNullException(nameof(protector));

        string? protectedApiKey = null;
        if (!string.IsNullOrWhiteSpace(settings.ApiKey))
        {
            var clearText = Encoding.UTF8.GetBytes(settings.ApiKey);
            try
            {
                protectedApiKey = Convert.ToBase64String(protector.Protect(clearText));
            }
            finally
            {
                Array.Clear(clearText, 0, clearText.Length);
            }
        }

        return new PersistedAgentSettings
        {
            BaseUrl = settings.BaseUrl,
            Model = settings.Model,
            AllowRemoteHttp = settings.AllowRemoteHttp,
            AllowRemoteWorkbookData = settings.AllowRemoteWorkbookData,
            ProtectedApiKey = protectedApiKey,
        };
    }

    public static AgentEndpointSettings Unprotect(
        PersistedAgentSettings settings,
        IAgentSecretProtector protector)
    {
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        if (protector == null) throw new ArgumentNullException(nameof(protector));

        string? apiKey = null;
        if (!string.IsNullOrWhiteSpace(settings.ProtectedApiKey))
        {
            byte[] protectedData;
            try
            {
                protectedData = Convert.FromBase64String(settings.ProtectedApiKey);
            }
            catch (FormatException exception)
            {
                throw new AgentSettingsException("The protected API key is not valid base64.", exception);
            }

            var clearText = protector.Unprotect(protectedData);
            try
            {
                apiKey = Encoding.UTF8.GetString(clearText);
            }
            finally
            {
                Array.Clear(clearText, 0, clearText.Length);
                Array.Clear(protectedData, 0, protectedData.Length);
            }
        }

        return new AgentEndpointSettings
        {
            BaseUrl = settings.BaseUrl,
            Model = string.IsNullOrWhiteSpace(settings.Model) ? AgentDefaults.Model : settings.Model,
            AllowRemoteHttp = settings.AllowRemoteHttp,
            AllowRemoteWorkbookData = settings.AllowRemoteWorkbookData,
            ApiKey = apiKey,
        };
    }
}

public sealed class AgentSettingsException : Exception
{
    public AgentSettingsException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}
