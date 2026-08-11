using System;

namespace ExcelReportBuilder.Agent.Configuration;

public static class AgentEndpointPolicy
{
    public static Uri Validate(AgentEndpointSettings settings)
    {
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        if (string.IsNullOrWhiteSpace(settings.BaseUrl) || settings.BaseUrl.Length > 2048)
        {
            throw new AgentEndpointPolicyException("The AI endpoint is required.");
        }

        if (string.IsNullOrWhiteSpace(settings.Model)) settings.Model = AgentDefaults.Model;
        if (settings.Model.Length > 256 || ContainsControlCharacter(settings.Model) ||
            (settings.ApiKey != null &&
             (settings.ApiKey.Length > 8192 || ContainsControlCharacter(settings.ApiKey))))
        {
            throw new AgentEndpointPolicyException("The AI endpoint settings exceed supported limits.");
        }

        Uri endpoint;
        if (!Uri.TryCreate(settings.BaseUrl, UriKind.Absolute, out endpoint!))
        {
            throw new AgentEndpointPolicyException("The AI endpoint must be an absolute URL.");
        }

        if (!string.Equals(endpoint.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(endpoint.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            throw new AgentEndpointPolicyException("The AI endpoint must use HTTP or HTTPS.");
        }

        if (!string.IsNullOrEmpty(endpoint.UserInfo))
        {
            throw new AgentEndpointPolicyException("Credentials are not allowed in the AI endpoint URL.");
        }

        if (!string.IsNullOrEmpty(endpoint.Query) || !string.IsNullOrEmpty(endpoint.Fragment))
        {
            throw new AgentEndpointPolicyException("The AI endpoint URL cannot contain a query or fragment.");
        }

        if (string.Equals(endpoint.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
            !endpoint.IsLoopback &&
            !settings.AllowRemoteHttp)
        {
            throw new AgentEndpointPolicyException(
                "Remote HTTP endpoints are blocked. Use HTTPS or explicitly allow remote HTTP.");
        }

        return endpoint;
    }

    public static Uri BuildV1Uri(Uri baseUri, string resource)
    {
        if (baseUri == null) throw new ArgumentNullException(nameof(baseUri));
        if (string.IsNullOrWhiteSpace(resource)) throw new ArgumentException("A resource is required.", nameof(resource));

        var root = baseUri.AbsoluteUri.TrimEnd('/');
        if (!root.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
        {
            root += "/v1";
        }

        return new Uri(root + "/" + resource.TrimStart('/'), UriKind.Absolute);
    }

    private static bool ContainsControlCharacter(string value)
    {
        foreach (var character in value)
        {
            if (char.IsControl(character)) return true;
        }

        return false;
    }
}

public sealed class AgentEndpointPolicyException : Exception
{
    public AgentEndpointPolicyException(string message)
        : base(message)
    {
    }
}
