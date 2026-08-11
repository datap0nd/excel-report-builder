using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace ExcelReportBuilder.Agent.Execution;

/// <summary>
/// Enforces one agent job per workbook across add-in instances in the current
/// Windows session. The named semaphore contains hashes only.
/// </summary>
public sealed class WorkbookAgentJobLease : IDisposable
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> LocalLeases =
        new ConcurrentDictionary<string, SemaphoreSlim>(StringComparer.Ordinal);

    private readonly Semaphore? _semaphore;
    private readonly SemaphoreSlim? _localSemaphore;
    private bool _ownsLease;

    private WorkbookAgentJobLease(Semaphore semaphore)
    {
        _semaphore = semaphore;
        _ownsLease = true;
    }

    private WorkbookAgentJobLease(SemaphoreSlim semaphore)
    {
        _localSemaphore = semaphore;
        _ownsLease = true;
    }

    public static bool TryAcquire(string workbookId, out WorkbookAgentJobLease? lease)
    {
        if (string.IsNullOrWhiteSpace(workbookId))
        {
            throw new ArgumentException("A stable workbook ID is required.", nameof(workbookId));
        }

        string leaseName = CreateLeaseName(workbookId);
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            SemaphoreSlim localSemaphore = LocalLeases.GetOrAdd(
                leaseName,
                _ => new SemaphoreSlim(1, 1));
            if (!localSemaphore.Wait(0))
            {
                lease = null;
                return false;
            }

            lease = new WorkbookAgentJobLease(localSemaphore);
            return true;
        }

        var semaphore = new Semaphore(1, 1, leaseName);
        bool acquired = semaphore.WaitOne(0);

        if (!acquired)
        {
            semaphore.Dispose();
            lease = null;
            return false;
        }

        lease = new WorkbookAgentJobLease(semaphore);
        return true;
    }

    public void Dispose()
    {
        if (_ownsLease)
        {
            _ownsLease = false;
            if (_semaphore != null)
            {
                _semaphore.Release();
            }
            else
            {
                _localSemaphore!.Release();
            }
        }

        _semaphore?.Dispose();
    }

    internal static string CreateLeaseName(string workbookId)
    {
        string material = (Environment.UserName ?? string.Empty) + "\n" + workbookId;
        byte[] clearText = Encoding.UTF8.GetBytes(material);
        byte[] digest;
        using (var sha256 = SHA256.Create())
        {
            digest = sha256.ComputeHash(clearText);
        }

        try
        {
            var builder = new StringBuilder("Local\\ExcelReportBuilder.AgentJob.", 96);
            for (var index = 0; index < 16; index++)
            {
                builder.Append(digest[index].ToString("x2"));
            }

            return builder.ToString();
        }
        finally
        {
            Array.Clear(clearText, 0, clearText.Length);
            Array.Clear(digest, 0, digest.Length);
        }
    }
}
