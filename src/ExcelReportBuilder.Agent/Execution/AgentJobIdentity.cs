using System;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Execution;

/// <summary>
/// Creates a non-identifying, deterministic job ID from the exact bounded
/// workbook request. The stable workbook ID acts as a local salt, while the
/// prompt and snapshots never appear in checkpoint filenames.
/// </summary>
public static class AgentJobIdentity
{
    public static string Create(AgentJobRequest request)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrWhiteSpace(request.WorkbookId))
        {
            throw new ArgumentException("A stable workbook ID is required.", nameof(request));
        }

        string boundedRequest = JsonSerializer.Serialize(new
        {
            request.WorkbookId,
            request.UserPrompt,
            request.Data,
            request.CurrentSpecification,
            request.MaxRepairCycles
        });
        byte[] clearText = Encoding.UTF8.GetBytes(boundedRequest);
        byte[] digest;
        using (var sha256 = SHA256.Create())
        {
            digest = sha256.ComputeHash(clearText);
        }

        try
        {
            var builder = new StringBuilder("job_", 36);
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
