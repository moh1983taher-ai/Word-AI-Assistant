using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Mcp;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class NativeMcpToolMappingTests
{
    [Fact]
    public void ExposesEverySelectedSourceToolWithAStableQualifiedNameAndOriginalSchema()
    {
        var tools = new[]
        {
            Tool("aljam3", "الجامع", "search", """
                {
                  "type": "object",
                  "properties": { "query": { "type": "string" } },
                  "required": ["query"],
                  "additionalProperties": false
                }
                """),
            Tool("shamela", "المكتبة الشاملة", "shamela_find", """
                {
                  "type": "object",
                  "properties": { "keywords": { "type": "array", "items": { "type": "string" } } },
                  "required": ["keywords"],
                  "additionalProperties": false
                }
                """)
        };

        var bindings = new Dictionary<string, ToolDescriptor>(StringComparer.Ordinal);
        var definitions = ResearchCoordinator.BuildNativeFunctionTools(tools, bindings);

        Assert.Equal(2, definitions.Count);
        Assert.Equal(2, bindings.Count);

        Assert.Contains("mcp_aljam3__search", bindings.Keys);
        Assert.Contains("mcp_shamela__shamela_find", bindings.Keys);

        var search = definitions.Single(item => item.Function.Name == "mcp_aljam3__search");
        Assert.Equal("object", search.Function.Parameters.GetProperty("type").GetString());
        Assert.True(search.Function.Parameters.GetProperty("properties").TryGetProperty("query", out _));

        var find = definitions.Single(item => item.Function.Name == "mcp_shamela__shamela_find");
        Assert.True(find.Function.Parameters.GetProperty("properties").TryGetProperty("keywords", out _));
    }

    [Fact]
    public async Task LocalMcpProxyPreservesInputSchemaAndAdvertisesReadOnlyHints()
    {
        await using var sources = new McpSourceRegistry(new HostLog());
        var descriptor = Tool(
            "aljam3",
            "الجامع",
            "search",
            """
            {
              "type": "object",
              "properties": { "query": { "type": "string" } },
              "required": ["query"],
              "additionalProperties": false
            }
            """);

        var proxy = new UpstreamMcpServerTool(sources, descriptor);
        var protocolTool = proxy.ProtocolTool;

        Assert.Equal("mcp_aljam3__search", protocolTool.Name);
        Assert.Equal("object", protocolTool.InputSchema.GetProperty("type").GetString());
        Assert.True(protocolTool.InputSchema.GetProperty("properties").TryGetProperty("query", out _));
        Assert.Equal(true, protocolTool.Annotations?.ReadOnlyHint);
        Assert.Equal(false, protocolTool.Annotations?.DestructiveHint);
        Assert.Equal(true, protocolTool.Annotations?.IdempotentHint);
    }

    [Fact]
    public void NamespacesSameToolNameAcrossDifferentLibraries()
    {
        var aljam3 = UpstreamMcpServerTool.BuildToolName("aljam3", "search");
        var shamela = UpstreamMcpServerTool.BuildToolName("shamela", "search");

        Assert.NotEqual(aljam3, shamela);
        Assert.Equal("mcp_aljam3__search", aljam3);
        Assert.Equal("mcp_shamela__search", shamela);
    }

    [Fact]
    public void AddsStableDisambiguatorWhenSanitizationWouldCollide()
    {
        var slashName = UpstreamMcpServerTool.BuildToolName("aljam3", "read/page");
        var spaceName = UpstreamMcpServerTool.BuildToolName("aljam3", "read page");

        Assert.NotEqual(slashName, spaceName);
        Assert.True(slashName.Length <= 64);
        Assert.True(spaceName.Length <= 64);
    }

    [Fact]
    public void QualifiedToolNamesAreSafeAndRemainWithinProviderNameLimit()
    {
        var name = UpstreamMcpServerTool.BuildToolName(
            "aljam3",
            "a tool/name with spaces and punctuation that is longer than the provider limit");

        Assert.True(name.Length <= 64);
        Assert.Matches("^[A-Za-z0-9_-]+$", name);
    }

    private static ToolDescriptor Tool(
        string server,
        string label,
        string name,
        string schemaText)
    {
        using var document = JsonDocument.Parse(schemaText);
        return new ToolDescriptor(
            server,
            label,
            name,
            "Read-only test tool",
            document.RootElement.Clone());
    }
}
