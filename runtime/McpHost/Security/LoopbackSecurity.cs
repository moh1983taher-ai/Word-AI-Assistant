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

    public static bool IsAllowedMcpOrigin(string? origin, int port = 18081)
    {
        if (string.IsNullOrWhiteSpace(origin))
            return true;

        if (IsAllowedOrigin(origin))
            return true;

        if (port is < 1025 or > 65535 ||
            !Uri.TryCreate(origin, UriKind.Absolute, out var uri))
            return false;

        return uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
               uri.IsLoopback &&
               uri.Port == port &&
               uri.AbsolutePath == "/" &&
               string.IsNullOrEmpty(uri.Query) &&
               string.IsNullOrEmpty(uri.Fragment) &&
               string.IsNullOrEmpty(uri.UserInfo);
    }

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
