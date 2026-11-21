using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchSourceToolRouterTests
{
    [Fact]
    public void KeepsBothSelectedLibrariesAvailableBeforeEitherHasSucceeded()
    {
        var tools = Tools();

        var selected = ResearchSourceToolRouter.SelectForTurn(
            tools,
            ["aljam3", "shamela"],
            []);

        Assert.Equal(new[] { "aljam3", "shamela" }, selected.Select(tool => tool.Server).Distinct());
        Assert.DoesNotContain(selected, tool => tool.Server == "unselected");
    }

    [Fact]
    public void RoutesNextTurnOnlyToTheSelectedLibraryStillPending()
    {
        var tools = Tools();

        var selected = ResearchSourceToolRouter.SelectForTurn(
            tools,
            ["aljam3", "shamela"],
            ["aljam3"]);

        Assert.NotEmpty(selected);
        Assert.All(selected, tool => Assert.Equal("shamela", tool.Server));
        Assert.Contains(selected, tool => tool.Name == "shamela_find");
    }

    [Fact]
    public void RestoresSelectedLibrariesForFollowUpReadingAfterAllHaveSucceeded()
    {
        var tools = Tools();

        var selected = ResearchSourceToolRouter.SelectForTurn(
            tools,
            ["aljam3", "shamela"],
            ["aljam3", "shamela"]);

        Assert.Equal(3, selected.Count);
        Assert.Equal(
            new[] { "aljam3", "shamela" },
            selected.Select(tool => tool.Server).Distinct().OrderBy(server => server));
        Assert.DoesNotContain(selected, tool => tool.Server == "unselected");
    }

    private static IReadOnlyList<ToolDescriptor> Tools()
    {
        using var schema = JsonDocument.Parse("{}");
        var input = schema.RootElement.Clone();

        return
        [
            new ToolDescriptor("aljam3", "الجامع", "search", "search", input),
            new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_find", "find", input),
            new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_open", "open", input),
            new ToolDescriptor("unselected", "مصدر غير مختار", "search", "search", input)
        ];
    }
}
