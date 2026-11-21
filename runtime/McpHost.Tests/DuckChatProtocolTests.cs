using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class DuckChatProtocolTests
{
    [Fact]
    public void ParsesStandardOpenAiFunctionCallResponse()
    {
        using var document = JsonDocument.Parse("""
            {
              "choices": [{
                "message": {
                  "role": "assistant",
                  "content": null,
                  "tool_calls": [{
                    "id": "call_123",
                    "type": "function",
                    "function": {
                      "name": "mcp_aljam3_search",
                      "arguments": "{\"query\":\"مقاصد الشريعة\"}"
                    }
                  }]
                }
              }]
            }
            """);

        var result = DuckAiClient.ParseChatCompletion(document.RootElement, "gpt-6-luna");

        Assert.Empty(result.Text);
        var call = Assert.Single(result.ToolCalls);
        Assert.Equal("call_123", call.Id);
        Assert.Equal("mcp_aljam3_search", call.Name);
        using var arguments = JsonDocument.Parse(call.Arguments);
        Assert.Equal("مقاصد الشريعة", arguments.RootElement.GetProperty("query").GetString());
    }

    [Fact]
    public void ParsesPlainAssistantTextWithoutFunctionCalls()
    {
        using var document = JsonDocument.Parse("""
            {"choices":[{"message":{"role":"assistant","content":"الإجابة النهائية"}}]}
            """);

        var result = DuckAiClient.ParseChatCompletion(document.RootElement, "gpt-6-luna");

        Assert.Equal("الإجابة النهائية", result.Text);
        Assert.Empty(result.ToolCalls);
    }

    [Fact]
    public void SerializesToolCallHistoryUsingOpenAiChatSchema()
    {
        var message = new DuckChatMessage(
            "tool",
            "نتيجة البحث",
            ToolCallId: "call_123");

        using var document = JsonDocument.Parse(
            JsonSerializer.Serialize(message, JsonDefaults.Web));

        Assert.Equal("tool", document.RootElement.GetProperty("role").GetString());
        Assert.Equal("call_123", document.RootElement.GetProperty("tool_call_id").GetString());
        Assert.False(document.RootElement.TryGetProperty("tool_calls", out _));
    }

    [Fact]
    public void SerializesMcpToolDefinitionsAsNativeFunctionTools()
    {
        using var schema = JsonDocument.Parse("""
            {"type":"object","properties":{"query":{"type":"string"}}}
            """);
        var definition = new DuckFunctionTool(
            "function",
            new DuckFunctionDefinition(
                "mcp_aljam3_search",
                "Search Aljam3",
                schema.RootElement.Clone()));

        using var document = JsonDocument.Parse(
            JsonSerializer.Serialize(definition, JsonDefaults.Web));

        var function = document.RootElement.GetProperty("function");
        Assert.Equal("mcp_aljam3_search", function.GetProperty("name").GetString());
        Assert.Equal("object", function.GetProperty("parameters").GetProperty("type").GetString());
    }
}
