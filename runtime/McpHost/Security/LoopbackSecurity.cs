namespace ResearchTools.McpHost.Security;

public static class LoopbackSecurity
{
    private static readonly HashSet<string> AllowedOrigins = new(StringComparer.OrdinalIgnoreCase)
    {
        "https://moh1983taher-ai.github.io",
        "https://localhost:3000",
        "https://localhost:3001"
    };

    private static readonly HashSet<string> AllowedHosts = new(StringComparer.OrdinalIgnoreCase)
    {
        "127.0.0.1:18081",
        "localhost:18081"
    };

    public static bool IsAllowedOrigin(string? origin) =>
        !string.IsNullOrWhiteSpace(origin) && AllowedOrigins.Contains(origin.TrimEnd('/'));

    public static bool IsAllowedHost(string? host) =>
        !string.IsNullOrWhiteSpace(host) && AllowedHosts.Contains(host.Trim());

    public static bool IsAuthorizedResearchClient(string? origin, string? client) =>
        IsAllowedOrigin(origin) && string.Equals(client, "word-addin", StringComparison.Ordinal);
}
