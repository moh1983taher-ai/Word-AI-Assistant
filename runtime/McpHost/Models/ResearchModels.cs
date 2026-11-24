using System.Text.Json;
using System.Text.Json.Serialization;

namespace ResearchTools.McpHost.Models;

public sealed class ResearchRequest
{
    [JsonPropertyName("client")]
    public string? Client { get; set; }

    [JsonPropertyName("model")]
    public string? Model { get; set; }

    [JsonPropertyName("servers")]
    public List<string> Servers { get; set; } = [];

    [JsonPropertyName("messages")]
    public List<ResearchMessage> Messages { get; set; } = [];
}

public sealed class ResearchMessage
{
    [JsonPropertyName("role")]
    public string? Role { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("content")]
    public string? Content { get; set; }

    public string GetText() => Text ?? Content ?? string.Empty;
}

public sealed record ToolCallTrace(string Server, string Label, string Tool);

public sealed record ResearchResponse(
    string Answer,
    string Model,
    IReadOnlyList<ToolCallTrace> ToolCalls,
    IReadOnlyList<string> ConnectionWarnings);

public sealed record DuckAiModelSnapshot(
    string BaseAddress,
    string DefaultModel,
    IReadOnlyList<string> Models);

public sealed record ToolDescriptor(
    string Server,
    string ServerLabel,
    string Name,
    string Description,
    JsonElement InputSchema);

public sealed record McpServerInstruction(
    string Server,
    string ServerLabel,
    string Text);

public sealed record SourceDiscovery(
    IReadOnlyList<ToolDescriptor> Tools,
    IReadOnlyList<string> Warnings,
    IReadOnlyList<McpServerInstruction> ServerInstructions);
