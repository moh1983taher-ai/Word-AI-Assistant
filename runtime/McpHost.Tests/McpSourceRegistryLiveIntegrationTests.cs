using ModelContextProtocol.Protocol;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

/// <summary>
/// Live, read-only integration coverage for the official Aljam3 and Shamela MCP servers.
/// These tests are kept separate from the deterministic unit-test suite because they
/// require outbound network access and the upstream services to be available.
/// </summary>
public sealed class McpSourceRegistryLiveIntegrationTests
{
    [Fact]
    [Trait("Category", "Integration")]
    public async Task ExecutesReadOnlySearchAgainstBothLiveMcpSources()
    {
        await using var registry = new McpSourceRegistry(new HostLog());
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(150));
        var selectedSources = new[] { "aljam3", "shamela" };

        var discovery = await registry.DiscoverSelectedAsync(
            selectedSources,
            timeout.Token);

        Assert.Empty(discovery.Warnings);
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "aljam3" && tool.Name == "search");
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "shamela" && tool.Name == "shamela_find");

        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            discovery.Tools,
            selectedSources,
            "مقاصد الشريعة");

        Assert.Equal(2, plan.Count);
        Assert.Equal(
            new[] { "aljam3", "shamela" },
            plan.Select(invocation => invocation.Tool.Server).ToArray());

        foreach (var invocation in plan)
        {
            var result = await registry.CallToolAsync(
                invocation.Tool.Server,
                invocation.Tool.Name,
                invocation.Arguments,
                timeout.Token);

            Assert.NotEqual(true, result.IsError);

            var textCharacters = result.Content
                .OfType<TextContentBlock>()
                .Sum(block => block.Text?.Length ?? 0);
            var structuredCharacters = result.StructuredContent is { } structured
                ? structured.GetRawText().Length
                : 0;
            var returnedCharacters = textCharacters + structuredCharacters;

            Assert.True(
                returnedCharacters > 0,
                $"{invocation.Tool.Server}/{invocation.Tool.Name} returned no content.");

            Console.WriteLine(
                $"Live MCP search succeeded: source={invocation.Tool.Server}; tool={invocation.Tool.Name}; returnedCharacters={returnedCharacters}.");
        }
    }
}
