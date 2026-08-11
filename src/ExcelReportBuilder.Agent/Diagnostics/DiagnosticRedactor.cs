using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelReportBuilder.Agent.Models;

namespace ExcelReportBuilder.Agent.Diagnostics;

public static class DiagnosticRedactor
{
    private const int MaximumDiagnosticCharacters = 512;

    private static readonly Regex BearerPattern = new Regex(
        @"(?i)(authorization\s*:\s*bearer\s+)[^\s,;]+",
        RegexOptions.CultureInvariant);

    private static readonly Regex SecretAssignmentPattern = new Regex(
        @"(?i)\b(api[-_ ]?key|access[-_ ]?token|token|secret|password)\b\s*[:=]\s*([^\s,;]+)",
        RegexOptions.CultureInvariant);

    private static readonly Regex UriUserInfoPattern = new Regex(
        @"(?i)(https?://)[^/@\s]+@",
        RegexOptions.CultureInvariant);

    private static readonly Regex UrlQueryPattern = new Regex(
        @"(?i)(https?://[^\s?#]+)\?[^\s#]*",
        RegexOptions.CultureInvariant);

    public static string Redact(string? value, IEnumerable<string?>? sensitiveValues = null)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "No diagnostic details were provided.";
        }

        var boundedValue = value!.Length > 8192 ? value.Substring(0, 8192) : value;
        var redacted = BearerPattern.Replace(boundedValue, "$1[redacted]");
        redacted = SecretAssignmentPattern.Replace(redacted, "$1=[redacted]");
        redacted = UriUserInfoPattern.Replace(redacted, "$1[redacted]@");
        redacted = UrlQueryPattern.Replace(redacted, "$1?[redacted]");

        if (sensitiveValues != null)
        {
            foreach (var sensitiveValue in sensitiveValues)
            {
                if (!string.IsNullOrEmpty(sensitiveValue))
                {
                    redacted = ReplaceOrdinalIgnoreCase(redacted, sensitiveValue!, "[redacted]");
                }
            }
        }

        redacted = redacted.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (redacted.Length > MaximumDiagnosticCharacters)
        {
            redacted = redacted.Substring(0, MaximumDiagnosticCharacters) + "...";
        }

        return redacted;
    }

    public static RedactedDiagnostic FromException(
        string code,
        Exception exception,
        bool retryable,
        IEnumerable<string?>? sensitiveValues = null)
    {
        if (exception == null) throw new ArgumentNullException(nameof(exception));
        return new RedactedDiagnostic
        {
            Code = string.IsNullOrWhiteSpace(code) ? "agent_error" : code,
            Message = Redact(exception.Message, sensitiveValues),
            Retryable = retryable,
        };
    }

    private static string ReplaceOrdinalIgnoreCase(string source, string oldValue, string newValue)
    {
        var start = 0;
        while (true)
        {
            var index = source.IndexOf(oldValue, start, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                return source;
            }

            source = source.Substring(0, index) + newValue + source.Substring(index + oldValue.Length);
            start = index + newValue.Length;
        }
    }
}
