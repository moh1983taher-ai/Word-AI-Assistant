using System.Text.Json;

namespace ResearchTools.McpHost.Services;

public sealed record ModelDecision(
    string Type,
    string Answer,
    string Server,
    string Tool,
    JsonElement Arguments,
    string Raw);

public static class ModelDecisionParser
{
    public static bool TryParse(string? response, out ModelDecision decision)
    {
        decision = new ModelDecision("invalid", string.Empty, string.Empty, string.Empty, default, response ?? string.Empty);
        if (string.IsNullOrWhiteSpace(response))
            return false;

        var candidate = response.Trim();
        if (candidate.StartsWith("'", StringComparison.Ordinal) || candidate.StartsWith("\"", StringComparison.Ordinal))
            return false;

        if (candidate.StartsWith("```", StringComparison.Ordinal))
        {
            var firstLine = candidate.IndexOf('\n');
            if (firstLine >= 0)
                candidate = candidate[(firstLine + 1)..];
            var fence = candidate.LastIndexOf("```", StringComparison.Ordinal);
            if (fence >= 0)
                candidate = candidate[..fence];
            candidate = candidate.Trim();
        }

        var start = candidate.IndexOf('{');
        var end = candidate.LastIndexOf('}');
        if (start < 0 || end <= start)
            return false;

        try
        {
            using var document = JsonDocument.Parse(candidate[start..(end + 1)]);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return false;

            if (TryGet(root, "tool_call", out var nested) && nested.ValueKind == JsonValueKind.Object)
            {
                root = nested;
                decision = new ModelDecision(
                    "tool_call",
                    string.Empty,
                    ReadString(root, "server"),
                    ReadString(root, "tool", "name"),
                    ReadObject(root, "arguments"),
                    response);
                return !string.IsNullOrWhiteSpace(decision.Server) &&
                       !string.IsNullOrWhiteSpace(decision.Tool);
            }

            var type = ReadString(root, "type");
            if (type == "final")
            {
                var answer = ReadString(root, "answer");
                if (string.IsNullOrWhiteSpace(answer))
                    return false;

                decision = new ModelDecision("final", answer.Trim(), string.Empty, string.Empty, default, response);
                return true;
            }

            if (type == "tool_call")
            {
                decision = new ModelDecision(
                    "tool_call",
                    string.Empty,
                    ReadString(root, "server"),
                    ReadString(root, "tool", "name"),
                    ReadObject(root, "arguments"),
                    response);
                return !string.IsNullOrWhiteSpace(decision.Server) &&
                       !string.IsNullOrWhiteSpace(decision.Tool);
            }
        }
        catch (JsonException)
        {
            return false;
        }

        return false;
    }

    private static string ReadString(JsonElement element, string primary, string? alternative = null)
    {
        if (TryGet(element, primary, out var value) && value.ValueKind == JsonValueKind.String)
            return value.GetString() ?? string.Empty;

        if (alternative is not null &&
            TryGet(element, alternative, out value) &&
            value.ValueKind == JsonValueKind.String)
            return value.GetString() ?? string.Empty;

        return string.Empty;
    }

    private static JsonElement ReadObject(JsonElement element, string name)
    {
        if (TryGet(element, name, out var value) && value.ValueKind == JsonValueKind.Object)
            return value.Clone();

        using var empty = JsonDocument.Parse("{}");
        return empty.RootElement.Clone();
    }

    private static bool TryGet(JsonElement element, string name, out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }
}
