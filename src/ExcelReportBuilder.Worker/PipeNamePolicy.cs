using System;
using ExcelReportBuilder.Agent.Security;

namespace ExcelReportBuilder.Worker;

public static class PipeNamePolicy
{
    public const string Prefix = WorkerHandshakeAuthenticator.PipePrefix;
    public const int MaximumLength = WorkerHandshakeAuthenticator.MaximumPipeNameLength;

    public static string Validate(string? pipeName)
    {
        if (string.IsNullOrWhiteSpace(pipeName) ||
            pipeName.Length > MaximumLength ||
            !pipeName.StartsWith(Prefix, StringComparison.Ordinal))
        {
            throw new ArgumentException("The worker pipe name is invalid.", nameof(pipeName));
        }

        foreach (var character in pipeName)
        {
            var allowed = (character >= 'a' && character <= 'z') ||
                          (character >= 'A' && character <= 'Z') ||
                          (character >= '0' && character <= '9') ||
                          character == '-' || character == '_' || character == '.';
            if (!allowed)
            {
                throw new ArgumentException("The worker pipe name is invalid.", nameof(pipeName));
            }
        }

        return pipeName;
    }
}
