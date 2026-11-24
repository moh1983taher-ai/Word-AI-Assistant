using System.Collections.Concurrent;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Security;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Owns one official MCP SDK client per source. Tool discovery is cached for five minutes;
/// tool calls are serialized per source while Kestrel remains free to serve other requests.
/// </summary>
public sealed class McpSourceRegistry : IAsyncDisposable
{
    private static readonly SourceDefinition[] Definitions =
    [
        new("aljam3", "الجامع", "https://aljam3.com/mcp"),
        new("shamela", "المكتبة الشاملة", "https://mcp.shamela.ws")
    ];

    private static readonly TimeSpan CatalogLifetime = TimeSpan.FromMinutes(5);
    private readonly ConcurrentDictionary<string, SourceState> _states =
        new(StringComparer.Ordinal);
    private readonly HostLog _log;

    public McpSourceRegistry(HostLog log)
    {
        _log = log;
        foreach (var definition in Definitions)
            _states.TryAdd(definition.Id, new SourceState(definition));
    }

    public string GetServerEndpoint(string sourceId)
    {
        if (!_states.TryGetValue(sourceId, out var state))
            throw new InvalidOperationException($"MCP source is not configured: {sourceId}");

        return state.Definition.Endpoint;
    }

    public Task<SourceDiscovery> DiscoverAllAsync(CancellationToken cancellationToken) =>
        DiscoverSelectedAsync(ResearchSourceSelection.AllSourceIds, cancellationToken);

    public async Task<SourceDiscovery> DiscoverSelectedAsync(
        IReadOnlyCollection<string> selectedSourceIds,
        CancellationToken cancellationToken)
    {
        var selected = ResearchSourceSelection.Normalize(selectedSourceIds);
        if (selected.Count == 0)
            throw new ArgumentException("يجب تحديد مكتبة واحدة على الأقل للبحث.", nameof(selectedSourceIds));

        var definitions = Definitions
            .Where(definition => selected.Contains(definition.Id, StringComparer.Ordinal))
            .ToArray();

        var tasks = definitions.Select(async definition =>
        {
            try
            {
                return await DiscoverOneAsync(_states[definition.Id], cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception exception)
            {
                _log.Warning($"MCP discovery failed for {definition.Id}: {exception.GetType().Name}");
                return new SourceDiscovery([], [$"{definition.Id}: {exception.GetType().Name}: {exception.Message}"]);
            }
        });

        var sources = await Task.WhenAll(tasks);
        return new SourceDiscovery(
            sources.SelectMany(source => source.Tools).ToArray(),
            sources.SelectMany(source => source.Warnings).ToArray());
    }

    public async Task<CallToolResult> CallToolAsync(
        string sourceId,
        string toolName,
        IReadOnlyDictionary<string, object?> arguments,
        CancellationToken cancellationToken)
    {
        if (!_states.TryGetValue(sourceId, out var state))
            throw new InvalidOperationException($"MCP source is not configured: {sourceId}");

        await state.Gate.WaitAsync(cancellationToken);
        try
        {
            await EnsureCatalogLockedAsync(state, cancellationToken);
            var tool = state.Tools.FirstOrDefault(candidate =>
                string.Equals(candidate.Name, toolName, StringComparison.Ordinal));

            if (tool is null)
                throw new InvalidOperationException($"Tool '{toolName}' is not available in the read-only catalog for {sourceId}.");

            var argumentJson = JsonSerializer.SerializeToElement(arguments, JsonDefaults.Web);
            var validationError = ToolArgumentValidator.Validate(tool.JsonSchema, argumentJson);
            if (validationError is not null)
                throw new InvalidOperationException(validationError);

            return await tool.CallAsync(arguments, cancellationToken: cancellationToken);
        }
        finally
        {
            state.Gate.Release();
        }
    }

    private async Task<SourceDiscovery> DiscoverOneAsync(
        SourceState state,
        CancellationToken cancellationToken)
    {
        await state.Gate.WaitAsync(cancellationToken);
        try
        {
            await EnsureCatalogLockedAsync(state, cancellationToken);

            var descriptors = state.Tools.Select(tool => new ToolDescriptor(
                state.Definition.Id,
                state.Definition.Label,
                tool.Name,
                Limit(tool.Description, 700),
                tool.JsonSchema.Clone())).ToArray();

            return new SourceDiscovery(descriptors, []);
        }
        finally
        {
            state.Gate.Release();
        }
    }

    private async Task EnsureCatalogLockedAsync(SourceState state, CancellationToken cancellationToken)
    {
        if (state.Client is not null && state.CatalogExpiresAt > DateTimeOffset.UtcNow)
            return;

        if (state.Client is not null)
        {
            try
            {
                await RefreshToolsAsync(state, cancellationToken);
                return;
            }
            catch (Exception exception) when (exception is not OperationCanceledException)
            {
                _log.Warning($"MCP catalog refresh failed for {state.Definition.Id}: {exception.GetType().Name}; reconnecting.");
                await DisposeConnectionAsync(state);
            }
        }

        await ConnectAndDiscoverAsync(state, cancellationToken);
    }

    private async Task ConnectAndDiscoverAsync(SourceState state, CancellationToken cancellationToken)
    {
        var transport = new HttpClientTransport(new HttpClientTransportOptions
        {
            Endpoint = new Uri(state.Definition.Endpoint),
            TransportMode = HttpTransportMode.AutoDetect,
            ConnectionTimeout = TimeSpan.FromSeconds(45),
            Name = $"ResearchTools/{state.Definition.Id}"
        });

        McpClient? client = null;
        try
        {
            client = await McpClient.CreateAsync(
                transport,
                new McpClientOptions
                {
                    InitializationTimeout = TimeSpan.FromSeconds(60),
                    DiscoverProbeTimeout = TimeSpan.FromSeconds(5)
                },
                cancellationToken: cancellationToken);

            var tools = await client.ListToolsAsync(cancellationToken: cancellationToken);
            state.Transport = transport;
            state.Client = client;
            state.Tools = tools.Where(candidate => ToolPolicy.IsAllowed(candidate.ProtocolTool)).ToList();
            state.CatalogExpiresAt = DateTimeOffset.UtcNow.Add(CatalogLifetime);

            _log.Info(
                $"MCP connected: {state.Definition.Id}; discovered={tools.Count}; readOnly={state.Tools.Count}; protocol={client.NegotiatedProtocolVersion ?? "per-request"}.");
        }
        catch
        {
            if (client is not null)
                await client.DisposeAsync();
            await transport.DisposeAsync();
            throw;
        }
    }

    private async Task RefreshToolsAsync(SourceState state, CancellationToken cancellationToken)
    {
        var tools = await state.Client!.ListToolsAsync(cancellationToken: cancellationToken);
        state.Tools = tools.Where(candidate => ToolPolicy.IsAllowed(candidate.ProtocolTool)).ToList();
        state.CatalogExpiresAt = DateTimeOffset.UtcNow.Add(CatalogLifetime);
        _log.Info(
            $"MCP catalog refreshed: {state.Definition.Id}; discovered={tools.Count}; readOnly={state.Tools.Count}.");
    }

    private static async Task DisposeConnectionAsync(SourceState state)
    {
        if (state.Client is not null)
        {
            await state.Client.DisposeAsync();
            state.Client = null;
        }

        if (state.Transport is not null)
        {
            await state.Transport.DisposeAsync();
            state.Transport = null;
        }

        state.Tools = [];
        state.CatalogExpiresAt = DateTimeOffset.MinValue;
    }

    private static string Limit(string? value, int maxLength)
    {
        var text = value ?? string.Empty;
        return text.Length <= maxLength ? text : text[..maxLength];
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var state in _states.Values)
        {
            await state.Gate.WaitAsync();
            try
            {
                await DisposeConnectionAsync(state);
            }
            finally
            {
                state.Gate.Release();
                state.Gate.Dispose();
            }
        }
    }

    private sealed record SourceDefinition(string Id, string Label, string Endpoint);

    private sealed class SourceState(SourceDefinition definition)
    {
        public SourceDefinition Definition { get; } = definition;
        public SemaphoreSlim Gate { get; } = new(1, 1);
        public HttpClientTransport? Transport { get; set; }
        public McpClient? Client { get; set; }
        public List<McpClientTool> Tools { get; set; } = [];
        public DateTimeOffset CatalogExpiresAt { get; set; }
    }
}
