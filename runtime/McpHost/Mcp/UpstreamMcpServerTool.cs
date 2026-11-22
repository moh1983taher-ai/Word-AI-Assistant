using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;

namespace ResearchTools.McpHost.Mcp;

/// <summary>
/// Exposes one read-only upstream MCP tool through the local ResearchTools MCP server.
/// The upstream server and original tool name remain private implementation details.
/// </summary>
public sealed class UpstreamMcpServerTool : McpServerTool
{
    private readonly McpSourceRegistry _sources;
    private readonly ToolDescriptor _descriptor;
    private readonly Tool _protocolTool;

    public UpstreamMcpServerTool(McpSourceRegistry sources, ToolDescriptor descriptor)
    {
        _sources = sources ?? throw new ArgumentNullException(nameof(sources));
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

        _protocolTool = new Tool
        {
            Name = BuildToolName(descriptor.Server, descriptor.Name),
            Title = $"{descriptor.ServerLabel}: {descriptor.Name}",
            Description = $"{descriptor.ServerLabel} — {descriptor.Description}",
            InputSchema = descriptor.InputSchema.Clone()
        };
    }

    public override Tool ProtocolTool => _protocolTool;

    public override IReadOnlyList<object> Metadata => Array.Empty<object>();

    public override async ValueTask<CallToolResult> InvokeAsync(
        RequestContext<CallToolRequestParams> request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var arguments = request.Params.Arguments is null
            ? new Dictionary<string, object?>(StringComparer.Ordinal)
            : request.Params.Arguments.ToDictionary(
                pair => pair.Key,
                pair => (object?)pair.Value.Clone(),
                StringComparer.Ordinal);

        return await _sources.CallToolAsync(
            _descriptor.Server,
            _descriptor.Name,
            arguments,
            cancellationToken).ConfigureAwait(false);
    }

    internal static string BuildToolName(string server, string tool)
    {
        static string Safe(string value) =>
            new(value.Select(character =>
                char.IsAsciiLetterOrDigit(character) || character is '_' or '-'
                    ? character
                    : '_').ToArray());

        var name = $"mcp_{Safe(server)}__{Safe(tool)}";
        if (name.Length <= 64)
            return name;

        var suffix = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(
                System.Text.Encoding.UTF8.GetBytes($"{server}/{tool}")))[..8].ToLowerInvariant();
        var prefixLength = 64 - suffix.Length - 1;
        return name[..prefixLength] + "_" + suffix;
    }
}
