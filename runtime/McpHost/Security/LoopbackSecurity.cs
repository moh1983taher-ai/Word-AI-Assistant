namespace ResearchTools.McpHost.Security;

public static class LoopbackSecurity
{
    private static readonly HashSet<string> AllowedOrigins = new(StringComparer.OrdinalIgnoreCase)
    {
        "https://moh1983taher-ai.github.io",
        "https://localhost:3000",
        "https://localhost:3001"
    };

    public static bool IsAllowedOrigin(string? origin) =>
        !string.IsNullOrWhiteSpace(origin) && AllowedOrigins.Contains(origin.TrimEnd('/'));

    public static bool IsAllowedHost(string? host, int port = 18081)
    {
        if (port is < 1025 or > 65535 || string.IsNullOrWhiteSpace(host))
            return false;

        return string.Equals(host, $"127.0.0.1:{port}", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(host, $"localhost:{port}", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsAuthorizedResearchClient(string? origin, string? client) =>
        IsAllowedOrigin(origin) && string.Equals(client, "word-addin", StringComparison.Ordinal);
}
