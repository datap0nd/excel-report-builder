using System;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReportBuilder.Worker;

public static class PipeNamePolicy
{
    public const string Prefix = "excel-report-builder-";
    public const int MaximumLength = 128;

    public static string CreateDefaultForCurrentUser()
    {
        var user = Environment.UserName ?? string.Empty;
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(user));
        var suffix = Convert.ToHexString(bytes, 0, 12).ToLowerInvariant();
        return Prefix + suffix;
    }

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
