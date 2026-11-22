using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using System.Text.Json;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchToolResultClassifierTests
{
    [Theory]
    [InlineData("{\"evidence\":{\"sources\":[]}}", false)]
    [InlineData("{\"evidence\":{\"sources\":[{\"text\":\"نص موثق\"}]}}", true)]
    public void DetectsEmptyAndNonEmptyJsonSearchPayloads(string text, bool expected)
    {
        Assert.Equal(expected, ResearchToolResultClassifier.HasMeaningfulText(text));
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("لا توجد نتائج مطابقة", false)]
    [InlineData("No results found", false)]
    [InlineData("عثر البحث على ثلاثة مواضع", true)]
    [InlineData("قواطع الأدلة في الأصول، ص 60", true)]
    public void DistinguishesUsefulTextFromEmptySearchResponses(string? text, bool expected)
    {
        Assert.Equal(expected, ResearchToolResultClassifier.HasMeaningfulText(text));
    }

    [Theory]
    [InlineData("aljam3", "search", false, true)]
    [InlineData("aljam3", "search", true, false)]
    [InlineData("shamela", "shamela_find", false, true)]
    [InlineData("shamela", "shamela_find_scoped", false, true)]
    [InlineData("shamela", "shamela_open", false, false)]
    public void EmptyOrNoMatchSearchResponsesCompleteCoverageButErrorsDoNot(
        string server,
        string name,
        bool isError,
        bool expected)
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            "",
            schema.RootElement.Clone());

        Assert.Equal(
            expected,
            ResearchToolResultClassifier.IsCompletedSearch(descriptor, isError));
    }

    [Theory]
    [InlineData("aljam3", "search", true)]
    [InlineData("aljam3", "get_page", false)]
    [InlineData("shamela", "shamela_find", true)]
    [InlineData("shamela", "shamela_find_scoped", true)]
    [InlineData("shamela", "shamela_open", false)]
    public void CountsOnlySearchOperationsTowardSourceCoverage(
        string server,
        string name,
        bool expected)
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            "",
            schema.RootElement.Clone());

        Assert.Equal(expected, ResearchToolResultClassifier.IsSearchTool(descriptor));
    }
}
