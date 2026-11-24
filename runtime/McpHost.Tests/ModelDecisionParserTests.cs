using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ModelDecisionParserTests
{
    [Fact]
    public void ParsesFinalAnswer()
    {
        const string json = "{\"type\":\"final\",\"answer\":\"النتيجة\",\"covered\":[\"محور البحث\"],\"uncovered\":[]}";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("final", decision.Type);
        Assert.Equal("النتيجة", decision.Answer);
    }

    [Fact]
    public void RejectsFinalWithoutExplicitCoverageAttestation()
    {
        const string json = """{"type":"final","answer":"النتيجة"}""";

        Assert.False(ModelDecisionParser.TryParse(json, out _));
    }

    [Fact]
    public void ParsesCoveredAndUncoveredPlanFacets()
    {
        const string json = """{"type":"final","answer":"النتيجة","covered":["التعريف","الأقوال"],"uncovered":["التطبيقات"]}""";

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal(new[] { "التعريف", "الأقوال" }, decision.Covered);
        Assert.Equal(new[] { "التطبيقات" }, decision.Uncovered);
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

    [Fact]
    public void ParsesModelAuthoredResearchPlanWithQueryPurposesAndCoverage()
    {
        const string json = """
            {
              "type": "plan",
              "intent": "استقصاء صيغ قاعدة أصولية",
              "kind": "multiple_forms",
              "queries": [
                { "query": "بناء القوي على الضعيف", "purpose": "الصيغة الحرفية" },
                { "query": "لا يجوز بناء القوي على الضعيف", "purpose": "صيغة الحكم" }
              ],
              "coverage": ["الصيغ", "المصادر", "نسبة الأقوال"]
            }
            """;

        Assert.True(ModelDecisionParser.TryParse(json, out var decision));
        Assert.Equal("plan", decision.Type);
        Assert.NotNull(decision.Plan);
        Assert.Equal("استقصاء صيغ قاعدة أصولية", decision.Plan.Intent);
        Assert.Equal("multiple_forms", decision.Plan.Kind);
        Assert.Equal(2, decision.Plan.Queries.Count);
        Assert.Equal("الصيغة الحرفية", decision.Plan.Queries[0].Purpose);
        Assert.Equal(new[] { "الصيغ", "المصادر", "نسبة الأقوال" }, decision.Plan.Coverage);
    }

    [Fact]
    public void RejectsPlansWithoutExecutableQueries()
    {
        const string json = """{"type":"plan","intent":"بحث عام","kind":"broad","queries":[],"coverage":["الموضوع"]}""";

        Assert.False(ModelDecisionParser.TryParse(json, out _));
    }

    [Fact]
    public void ExtractsOneJsonObjectFromSurroundingModelText()
    {
        const string response = """
            إليك خطة البحث:
            {"type":"plan","intent":"بحث","kind":"broad","queries":[{"query":"الاجتهاد","purpose":"المفهوم"}],"coverage":["التعريف"]}
            """;

        Assert.True(ModelDecisionParser.TryParse(response, out var decision));
        Assert.Equal("plan", decision.Type);
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
