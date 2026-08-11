using System;
using System.Collections.Concurrent;

namespace ExcelReportBuilder.Worker;

public sealed class ActiveWorkbookJobRegistry
{
    private readonly ConcurrentDictionary<string, string> _jobsByWorkbook =
        new ConcurrentDictionary<string, string>(StringComparer.Ordinal);

    public bool TryStart(string workbookId, string jobId)
    {
        if (string.IsNullOrWhiteSpace(workbookId)) throw new ArgumentException("A workbook ID is required.", nameof(workbookId));
        if (string.IsNullOrWhiteSpace(jobId)) throw new ArgumentException("A job ID is required.", nameof(jobId));
        return _jobsByWorkbook.TryAdd(workbookId, jobId);
    }

    public bool TryGetActiveJob(string workbookId, out string? jobId)
    {
        return _jobsByWorkbook.TryGetValue(workbookId, out jobId);
    }

    public void Complete(string workbookId, string jobId)
    {
        if (_jobsByWorkbook.TryGetValue(workbookId, out var activeJob) &&
            string.Equals(activeJob, jobId, StringComparison.Ordinal))
        {
            _jobsByWorkbook.TryRemove(workbookId, out _);
        }
    }
}
