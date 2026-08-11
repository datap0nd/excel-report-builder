using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Worker;

public interface IJobCheckpointStore
{
    Task<AgentResumeMetadata> SaveCheckpointAsync(
        AgentCheckpointEvent checkpoint,
        CancellationToken cancellationToken);

    Task<AgentResumeMetadata> MarkFailureAsync(
        string jobId,
        string workbookId,
        string failureCode,
        CancellationToken cancellationToken);

    Task<AgentResumeMetadata?> LoadAsync(
        string jobId,
        string workbookId,
        CancellationToken cancellationToken);

    Task<IReadOnlyList<AgentResumeMetadata>> ListAsync(
        string? workbookId,
        CancellationToken cancellationToken);

    Task DeleteAsync(string jobId, string workbookId, CancellationToken cancellationToken);
}

/// <summary>
/// Stores only non-sensitive progress metadata. Prompts, endpoint settings,
/// source descriptions, tool arguments, and tool results are never persisted.
/// </summary>
public sealed class LocalAppDataJobCheckpointStore : IJobCheckpointStore
{
    private readonly string _directory;
    private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);
    private static readonly JsonSerializerOptions JsonOptions = CreateJsonOptions();

    public LocalAppDataJobCheckpointStore(string? directory = null)
    {
        _directory = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelReportBuilder",
            "AgentCheckpoints");
    }

    public async Task<AgentResumeMetadata> SaveCheckpointAsync(
        AgentCheckpointEvent checkpoint,
        CancellationToken cancellationToken)
    {
        if (checkpoint == null) throw new ArgumentNullException(nameof(checkpoint));
        var metadata = new AgentResumeMetadata
        {
            JobId = checkpoint.JobId,
            WorkbookId = checkpoint.WorkbookId,
            CheckpointId = checkpoint.CheckpointId,
            Stage = checkpoint.Stage,
            CompletedRepairCycles = checkpoint.CompletedRepairCycles,
            LastCompletedStep = checkpoint.LastCompletedStep,
            FailureCode = null,
            CanResume = true,
            UpdatedAtUtc = checkpoint.OccurredAtUtc,
        };
        await WriteAsync(metadata, cancellationToken).ConfigureAwait(false);
        return metadata;
    }

    public async Task<AgentResumeMetadata> MarkFailureAsync(
        string jobId,
        string workbookId,
        string failureCode,
        CancellationToken cancellationToken)
    {
        var existing = await LoadAsync(jobId, workbookId, cancellationToken).ConfigureAwait(false);
        var metadata = existing ?? new AgentResumeMetadata
        {
            JobId = jobId,
            WorkbookId = workbookId,
            CheckpointId = string.Empty,
            Stage = AgentProgressStage.Failed,
            LastCompletedStep = "No resumable checkpoint was reached.",
            CanResume = false,
        };
        metadata.FailureCode = SafeCode(failureCode);
        metadata.UpdatedAtUtc = DateTimeOffset.UtcNow;
        await WriteAsync(metadata, cancellationToken).ConfigureAwait(false);
        return metadata;
    }

    public async Task<AgentResumeMetadata?> LoadAsync(
        string jobId,
        string workbookId,
        CancellationToken cancellationToken)
    {
        var path = GetPath(jobId, workbookId);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!File.Exists(path)) return null;
            var json = await File.ReadAllTextAsync(path, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Deserialize<AgentResumeMetadata>(json, JsonOptions);
        }
        catch (JsonException)
        {
            return null;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<IReadOnlyList<AgentResumeMetadata>> ListAsync(
        string? workbookId,
        CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (!Directory.Exists(_directory)) return Array.Empty<AgentResumeMetadata>();
            var results = new List<AgentResumeMetadata>();
            foreach (var path in Directory.EnumerateFiles(_directory, "*.json", SearchOption.TopDirectoryOnly))
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var json = await File.ReadAllTextAsync(path, cancellationToken).ConfigureAwait(false);
                    var item = JsonSerializer.Deserialize<AgentResumeMetadata>(json, JsonOptions);
                    if (item != null &&
                        (string.IsNullOrWhiteSpace(workbookId) || string.Equals(item.WorkbookId, workbookId, StringComparison.Ordinal)))
                    {
                        results.Add(item);
                    }
                }
                catch (JsonException)
                {
                }
                catch (IOException)
                {
                }
            }

            return results.OrderByDescending(item => item.UpdatedAtUtc).Take(100).ToArray();
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task DeleteAsync(string jobId, string workbookId, CancellationToken cancellationToken)
    {
        var path = GetPath(jobId, workbookId);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task WriteAsync(AgentResumeMetadata metadata, CancellationToken cancellationToken)
    {
        var path = GetPath(metadata.JobId, metadata.WorkbookId);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            Directory.CreateDirectory(_directory);
            var temporaryPath = path + ".tmp-" + Guid.NewGuid().ToString("N");
            var json = JsonSerializer.Serialize(metadata, JsonOptions);
            await File.WriteAllTextAsync(temporaryPath, json, new UTF8Encoding(false), cancellationToken).ConfigureAwait(false);
            File.Move(temporaryPath, path, true);
        }
        finally
        {
            _gate.Release();
        }
    }

    private string GetPath(string jobId, string workbookId)
    {
        var identity = Encoding.UTF8.GetBytes(workbookId + "\0" + jobId);
        var hash = Convert.ToHexString(SHA256.HashData(identity)).ToLowerInvariant();
        return Path.Combine(_directory, hash + ".json");
    }

    private static string SafeCode(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "worker_failure";
        var builder = new StringBuilder();
        foreach (var character in value.Take(128))
        {
            var allowed = char.IsLetterOrDigit(character) || character == '_' || character == '-' || character == '.';
            if (allowed) builder.Append(character);
        }

        return builder.Length == 0 ? "worker_failure" : builder.ToString();
    }

    private static JsonSerializerOptions CreateJsonOptions()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, false));
        return options;
    }
}
