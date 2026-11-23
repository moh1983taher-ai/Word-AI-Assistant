namespace ResearchTools.McpHost.Services;

/// <summary>
/// Selects the local agent runtime without changing the Word-facing REST contract.
/// DuckAI remains the default; LM Studio can be enabled explicitly per machine.
/// </summary>
public sealed class ResearchProviderOptions
{
    public const string ProviderEnvironmentVariable = "RESEARCHTOOLS_AGENT_PROVIDER";
    public const string LmStudioBaseUrlEnvironmentVariable = "RESEARCHTOOLS_LMSTUDIO_BASE_URL";
    public const string LmStudioApiTokenEnvironmentVariable = "RESEARCHTOOLS_LMSTUDIO_API_TOKEN";

    public string Provider { get; }
    public bool UseLmStudio => string.Equals(Provider, "lmstudio", StringComparison.OrdinalIgnoreCase);
    public Uri LmStudioBaseUri { get; }
    public string? LmStudioApiToken { get; }

    internal ResearchProviderOptions(string provider, Uri baseUri, string? apiToken)
    {
        Provider = provider;
        LmStudioBaseUri = baseUri;
        LmStudioApiToken = apiToken;
    }

    public static ResearchProviderOptions FromEnvironment()
    {
        var provider = Environment.GetEnvironmentVariable(ProviderEnvironmentVariable)?.Trim();
        if (string.IsNullOrWhiteSpace(provider))
            provider = "duckai";

        if (!string.Equals(provider, "duckai", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(provider, "lmstudio", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"{ProviderEnvironmentVariable} must be either 'duckai' or 'lmstudio'.");
        }

        var configuredUrl = Environment.GetEnvironmentVariable(LmStudioBaseUrlEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(configuredUrl))
            configuredUrl = "http://127.0.0.1:1234";

        if (!Uri.TryCreate(configuredUrl.Trim(), UriKind.Absolute, out var baseUri) ||
            (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps) ||
            !baseUri.IsLoopback ||
            !string.IsNullOrEmpty(baseUri.UserInfo) ||
            !string.IsNullOrEmpty(baseUri.Query) ||
            !string.IsNullOrEmpty(baseUri.Fragment))
        {
            throw new InvalidOperationException(
                $"{LmStudioBaseUrlEnvironmentVariable} must be an HTTP(S) loopback URL such as http://127.0.0.1:1234.");
        }

        var normalizedBaseUri = new Uri(baseUri.AbsoluteUri.TrimEnd('/') + "/");
        var token = Environment.GetEnvironmentVariable(LmStudioApiTokenEnvironmentVariable)?.Trim();
        if (string.IsNullOrWhiteSpace(token))
            token = null;

        return new ResearchProviderOptions(provider.ToLowerInvariant(), normalizedBaseUri, token);
    }
}
