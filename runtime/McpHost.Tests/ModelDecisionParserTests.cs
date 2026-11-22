using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ModelDecisionParserTests
{
    [Fact]
    public void ParsesFinalAnswer()
    {
        const string json = "{\"type\":\"final\",\"answer\":\"النتيجة\"}";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("final", decision.Type);
        Assert.Equal("النتيجة", decision.Answer);
    }

    [Fact]
    public void ParsesToolCallAndArguments()
    {
        const string json = "{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"arguments\":{\"query\":\"مقاصد الشريعة\"}}";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("tool_call", decision.Type);
        Assert.Equal("aljam3", decision.Server);
        Assert.Equal("search", decision.Tool);
        Assert.Equal("مقاصد الشريعة", decision.Arguments.GetProperty("query").GetString());
    }

    [Fact]
    public void ParsesToolCallAndItsExplicitTaskScopeBinding()
    {
        const string json = "{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"scope_item\":\"S2\",\"arguments\":{\"q\":\"الأمر\"}}";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("S2", decision.ScopeItem);
    }

    [Fact]
    public void ParsesLegacyNestedToolCallShape()
    {
        const string json = "{\"tool_call\":{\"server\":\"shamela\",\"name\":\"shamela_find\",\"arguments\":{\"query\":\"الاجتهاد\"}}}";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("tool_call", decision.Type);
        Assert.Equal("shamela", decision.Server);
        Assert.Equal("shamela_find", decision.Tool);
    }

    [Fact]
    public void AcceptsJsonCodeFence()
    {
        const string response = "```json\n{\"type\":\"final\",\"answer\":\"تم\"}\n```";

        Assert.True(ModelDecisionParser.TryParse(response, out var decision));
        Assert.Equal("تم", decision.Answer);
    }

    [Theory]
    [InlineData("{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\"}")]
    [InlineData("{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"arguments\":null}")]
    [InlineData("{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"arguments\":\"{}\"}")]
    public void RejectsToolCallsWithoutObjectArguments(string response)
    {
        Assert.False(ModelDecisionParser.TryParse(response, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("لا يوجد JSON هنا")]
    [InlineData("{\"type\":\"tool_call\",\"server\":\"aljam3\"}")]
    [InlineData("{\"type\":\"final\",\"answer\":\"\"}")]
    public void RejectsIncompleteOrInvalidDecisions(string response)
    {
        Assert.False(ModelDecisionParser.TryParse(response, out _));
    }
}
