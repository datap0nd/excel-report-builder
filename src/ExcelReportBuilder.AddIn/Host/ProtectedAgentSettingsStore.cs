using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Configuration;

namespace ExcelReportBuilder.AddIn.Host
{
    internal sealed class WindowsCurrentUserSecretProtector : IAgentSecretProtector
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes(
            "ExcelReportBuilder.AgentSettings.v1");

        public byte[] Protect(byte[] clearText)
        {
            if (clearText == null)
            {
                throw new ArgumentNullException(nameof(clearText));
            }

            return ProtectedData.Protect(clearText, Entropy, DataProtectionScope.CurrentUser);
        }

        public byte[] Unprotect(byte[] protectedData)
        {
            if (protectedData == null)
            {
                throw new ArgumentNullException(nameof(protectedData));
            }

            return ProtectedData.Unprotect(protectedData, Entropy, DataProtectionScope.CurrentUser);
        }
    }

    /// <summary>
    /// Stores endpoint preferences per Windows user. API keys are always DPAPI
    /// protected before JSON is written beneath LocalApplicationData.
    /// </summary>
    internal sealed class ProtectedAgentSettingsStore : IAgentSettingsStore
    {
        private const int MaximumSettingsBytes = 64 * 1024;
        private readonly string _settingsPath;

        public ProtectedAgentSettingsStore(string? settingsPath = null)
        {
            _settingsPath = settingsPath ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelReportBuilder",
                "agent-settings.json");
        }

        public Task<PersistedAgentSettings?> LoadAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!File.Exists(_settingsPath))
            {
                return Task.FromResult<PersistedAgentSettings?>(null);
            }

            var fileInfo = new FileInfo(_settingsPath);
            if (fileInfo.Length <= 0 || fileInfo.Length > MaximumSettingsBytes)
            {
                throw new AgentSettingsException("The saved endpoint settings file is invalid.");
            }

            string json = File.ReadAllText(_settingsPath, Encoding.UTF8);
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                return Task.FromResult(JsonSerializer.Deserialize<PersistedAgentSettings>(json));
            }
            catch (JsonException exception)
            {
                throw new AgentSettingsException("The saved endpoint settings file is invalid.", exception);
            }
        }

        public Task SaveAsync(PersistedAgentSettings settings, CancellationToken cancellationToken)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            cancellationToken.ThrowIfCancellationRequested();
            string? directory = Path.GetDirectoryName(_settingsPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new AgentSettingsException("The endpoint settings location is invalid.");
            }

            Directory.CreateDirectory(directory!);
            string json = JsonSerializer.Serialize(settings);
            if (Encoding.UTF8.GetByteCount(json) > MaximumSettingsBytes)
            {
                throw new AgentSettingsException("The endpoint settings exceed the supported size.");
            }

            string temporaryPath = _settingsPath + ".tmp";
            File.WriteAllText(temporaryPath, json, new UTF8Encoding(false));
            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(_settingsPath))
            {
                File.Replace(temporaryPath, _settingsPath, null);
            }
            else
            {
                File.Move(temporaryPath, _settingsPath);
            }

            return Task.CompletedTask;
        }

        public PersistedAgentSettings? TryLoad()
        {
            try
            {
                return LoadAsync(CancellationToken.None).GetAwaiter().GetResult();
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
            catch (AgentSettingsException)
            {
                return null;
            }
        }
    }
}
