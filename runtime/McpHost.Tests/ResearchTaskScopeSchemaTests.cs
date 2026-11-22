using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchTaskScopeSchemaTests
{
    [Fact]
    public void RequiresAnExplicitScopeItemOnEveryToolCall()
    {
        using var input = JsonDocument.Parse("""
            {
              "type": "object",
              "properties": { "q": { "type": "string" } },
              "required": ["q"],
              "additionalProperties": false
            }
            """);
        var tools = new[]
        {
            new ToolDescriptor("aljam3", "الجامع", "search", "بحث", input.RootElement.Clone())
        };

        using var response = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(
                tools,
                allowFinal: false,
                scopeItemIds: ["S1", "S2"]).GetRawText());

        var schema = response.RootElement.GetProperty("json_schema").GetProperty("schema");
        var scope = schema.GetProperty("properties").GetProperty("scope_item");
        Assert.Equal(
            new[] { "S1", "S2" },
            scope.GetProperty("enum").EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Contains(
            "scope_item",
            schema.GetProperty("required").EnumerateArray().Select(item => item.GetString()));
    }

    [Fact]
    public void FinalDecisionMayUseOnlyTheEmptyScopeBinding()
    {
        using var response = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(
                Array.Empty<ToolDescriptor>(),
                allowFinal: true,
                allowToolCalls: false,
                scopeItemIds: ["S1"]).GetRawText());

        var enumValues = response.RootElement
            .GetProperty("json_schema")
            .GetProperty("schema")
            .GetProperty("properties")
            .GetProperty("scope_item")
            .GetProperty("enum")
            .EnumerateArray()
            .Select(item => item.GetString())
            .ToArray();

        Assert.Equal(new[] { "" }, enumValues);
    }
}
