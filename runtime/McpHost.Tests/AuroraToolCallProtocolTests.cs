using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class AuroraToolCallProtocolTests
{
    [Fact]
    public void ParsesAuroraTaggedToolCall()
    {
        const string response = """
            <tool_call>
            {"name":"aljam3__search","arguments":{"query":"بناء القوي على الضعيف","limit":5}}
            </tool_call>
            """;

        Assert.True(AuroraToolCallProtocol.TryParse(response, out var decision));
        Assert.Equal("tool_call", decision.Type);
        Assert.Equal("aljam3", decision.Server);
        Assert.Equal("search", decision.Tool);
        Assert.Equal("بناء القوي على الضعيف", decision.Arguments.GetProperty("query").GetString());
        Assert.Equal(5, decision.Arguments.GetProperty("limit").GetInt32());
    }

    [Fact]
    public void ParsesTaggedToolCallSurroundedByWhitespace()
    {
        const string response = """
            
            <tool_call>{"name":"shamela__shamela_find","arguments":{"query":"الاجتهاد"}}</tool_call>
            
            """;

        Assert.True(AuroraToolCallProtocol.TryParse(response, out var decision));
        Assert.Equal("shamela", decision.Server);
        Assert.Equal("shamela_find", decision.Tool);
    }

    [Theory]
    [InlineData("<tool_call>{\"name\":\"search\",\"arguments\":{}}</tool_call>")]
    [InlineData("<tool_call>{\"name\":\"aljam3__search\",\"arguments\":null}</tool_call>")]
    [InlineData("<tool_call>{\"name\":\"aljam3__search\",\"arguments\":[]}</tool_call>")]
    [InlineData("<tool_call>{\"name\":\"aljam3__search\",\"arguments\":{}}")]
    [InlineData("<tool_call>{broken}</tool_call>")]
    public void RejectsMalformedOrUnqualifiedToolCalls(string response)
    {
        Assert.False(AuroraToolCallProtocol.TryParse(response, out _));
    }

    [Fact]
    public void RejectsMultipleToolCallsInOneTurn()
    {
        const string response =
            "<tool_call>{\"name\":\"aljam3__search\",\"arguments\":{}}</tool_call>" +
            "<tool_call>{\"name\":\"shamela__shamela_find\",\"arguments\":{}}</tool_call>";

        Assert.False(AuroraToolCallProtocol.TryParse(response, out _));
    }

    [Fact]
    public void BuildsCompactCatalogWithQualifiedNamesAndRequiredParameters()
    {
        using var schema = JsonDocument.Parse("""
            {
              "type":"object",
              "properties":{
                "query":{"type":"string","description":"عبارة البحث"},
                "limit":{"type":"integer","description":"عدد النتائج","enum":[5,10]}
              },
              "required":["query"]
            }
            """);
        var tools = new[]
        {
            new ToolDescriptor(
                "aljam3",
                "الجامع",
                "search",
                "ابحث في فهرس المصادر",
                schema.RootElement.Clone())
        };

        var catalog = AuroraToolCallProtocol.BuildToolCatalog(tools);

        Assert.Contains("aljam3__search", catalog, StringComparison.Ordinal);
        Assert.Contains("query (string، إلزامي)", catalog, StringComparison.Ordinal);
        Assert.Contains("limit (integer، اختياري)", catalog, StringComparison.Ordinal);
        Assert.Contains("الخيارات: 5، 10", catalog, StringComparison.Ordinal);
        Assert.DoesNotContain("\"inputSchema\"", catalog, StringComparison.Ordinal);
    }

    [Fact]
    public void ToolAliasesKeepTheMcpServerAndToolNameSeparate()
    {
        var descriptor = new ToolDescriptor(
            "shamela",
            "المكتبة الشاملة",
            "shamela_find",
            "بحث",
            JsonDocument.Parse("""{"type":"object","properties":{}}""").RootElement.Clone());

        Assert.Equal("shamela__shamela_find", AuroraToolCallProtocol.BuildToolAlias(descriptor));
    }
}
