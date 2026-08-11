using System;

namespace ExcelReportBuilder.Agent.Configuration;

/// <summary>
/// Compares endpoint origins and paths before a protected API key is reused.
/// URI schemes and host names are case-insensitive; URL paths are not.
/// </summary>
public static class AgentEndpointCredentialScope
{
    public static bool Matches(string savedBaseUrl, string requestedBaseUrl)
    {
        if (!Uri.TryCreate(savedBaseUrl, UriKind.Absolute, out var saved) ||
            !Uri.TryCreate(requestedBaseUrl, UriKind.Absolute, out var requested))
        {
            return false;
        }

        return string.Equals(saved.Scheme, requested.Scheme, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(saved.IdnHost, requested.IdnHost, StringComparison.OrdinalIgnoreCase) &&
            saved.Port == requested.Port &&
            string.Equals(NormalizePath(saved), NormalizePath(requested), StringComparison.Ordinal) &&
            string.Equals(saved.Query, requested.Query, StringComparison.Ordinal) &&
            string.Equals(saved.Fragment, requested.Fragment, StringComparison.Ordinal);
    }

    private static string NormalizePath(Uri endpoint)
    {
        string path = endpoint.GetComponents(UriComponents.Path, UriFormat.UriEscaped)
            .TrimEnd('/');
        return path.Length == 0 ? "/" : "/" + path;
    }
}
