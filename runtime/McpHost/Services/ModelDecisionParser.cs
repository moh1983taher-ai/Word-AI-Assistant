using System.Text.Json;

namespace ResearchTools.McpHost.Services;

public sealed record ModelDecision(
    string Type,
    string Answer,
    string Server,
    string Tool,
    JsonElement Arguments,
    string Raw)
{
    public ResearchPlan? Plan { get; init; }
}

public static class ModelDecisionParser
{
    private const int MaximumPlanQueries = 8;
    private const int MaximumQueryCharacters = 600;
    private const int MaximumPurposeCharacters = 240;
    private const int MaximumCoverageItems = 12;

    public static bool TryParse(string? response, out ModelDecision decision)
    {
        decision = new ModelDecision("invalid", string.Empty, string.Empty, string.Empty, default, response ?? string.Empty);
        if (!TryExtractJsonObject(response, out var candidate))
            return false;

        try
        {
            using var document = JsonDocument.Parse(candidate);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return false;

            if (TryGet(root, "tool_call", out var nested) && nested.ValueKind == JsonValueKind.Object)
            {
                root = nested;
                if (!TryGet(root, "arguments", out var nestedArguments) ||
                    nestedArguments.ValueKind != JsonValueKind.Object)
                    return false;

                decision = new ModelDecision(
                    "tool_call",
                    string.Empty,
                    ReadString(root, "server"),
                    ReadString(root, "tool", "name"),
                    nestedArguments.Clone(),
                    response ?? string.Empty);
                return !string.IsNullOrWhiteSpace(decision.Server) &&
                       !string.IsNullOrWhiteSpace(decision.Tool);
            }

            var type = ReadString(root, "type");
            if (type == "final")
            {
                var answer = ReadString(root, "answer");
                if (string.IsNullOrWhiteSpace(answer))
                    return false;

                decision = new ModelDecision("final", answer.Trim(), string.Empty, string.Empty, default, response ?? string.Empty);
                return true;
            }

            if (type == "tool_call")
            {
                if (!TryGet(root, "arguments", out var arguments) ||
                    arguments.ValueKind != JsonValueKind.Object)
                    return false;

                decision = new ModelDecision(
                    "tool_call",
                    string.Empty,
                    ReadString(root, "server"),
                    ReadString(root, "tool", "name"),
                    arguments.Clone(),
                    response ?? string.Empty);
                return !string.IsNullOrWhiteSpace(decision.Server) &&
                       !string.IsNullOrWhiteSpace(decision.Tool);
            }

            if (type == "plan" && TryParsePlan(root, out var plan))
            {
                decision = new ModelDecision(
                    "plan",
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    default,
                    response ?? string.Empty)
                {
                    Plan = plan
                };
                return true;
            }
        }
        catch (JsonException)
        {
            return false;
        }

        return false;
    }

    private static bool TryParsePlan(JsonElement root, out ResearchPlan plan)
    {
        plan = new ResearchPlan(string.Empty, string.Empty, Array.Empty<ResearchQuery>(), Array.Empty<string>());
        var intent = Limit(ReadString(root, "intent"), 500).Trim();
        var kind = Limit(ReadString(root, "kind"), 64).Trim();
        if (string.IsNullOrWhiteSpace(intent) ||
            string.IsNullOrWhiteSpace(kind) ||
            !TryGet(root, "queries", out var queriesElement) ||
            queriesElement.ValueKind != JsonValueKind.Array ||
            queriesElement.GetArrayLength() == 0 ||
            queriesElement.GetArrayLength() > MaximumPlanQueries)
            return false;

        var queries = new List<ResearchQuery>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in queriesElement.EnumerateArray())
        {
            var query = item.ValueKind == JsonValueKind.String
                ? item.GetString() ?? string.Empty
                : ReadString(item, "query", "text");
            var purpose = item.ValueKind == JsonValueKind.Object
                ? ReadString(item, "purpose")
                : string.Empty;

            query = Limit(query, MaximumQueryCharacters).Trim();
            purpose = Limit(purpose, MaximumPurposeCharacters).Trim();
            if (string.IsNullOrWhiteSpace(query))
                return false;

            var normalized = string.Join(' ', query.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            if (!seen.Add(normalized))
                continue;

            queries.Add(new ResearchQuery(query, purpose));
        }

        if (queries.Count == 0)
            return false;

        var coverage = new List<string>();
        if (TryGet(root, "coverage", out var coverageElement) &&
            coverageElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in coverageElement.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.String)
                    continue;
                var value = Limit(item.GetString() ?? string.Empty, 240).Trim();
                if (value.Length > 0 &&
                    !coverage.Contains(value, StringComparer.OrdinalIgnoreCase))
                    coverage.Add(value);
                if (coverage.Count >= MaximumCoverageItems)
                    break;
            }
        }

        plan = new ResearchPlan(intent, kind, queries, coverage);
        return true;
    }

    private static bool TryExtractJsonObject(string? response, out string json)
    {
        json = string.Empty;
        if (string.IsNullOrWhiteSpace(response))
            return false;

        var text = response.Trim();
        for (var start = text.IndexOf('{'); start >= 0; start = text.IndexOf('{', start + 1))
        {
            var depth = 0;
            var inString = false;
            var escaped = false;

            for (var index = start; index < text.Length; index++)
            {
                var character = text[index];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                        continue;
                    }
                    if (character == '\\')
                    {
                        escaped = true;
                        continue;
                    }
                    if (character == '"')
                        inString = false;
                    continue;
                }

                if (character == '"')
                {
                    inString = true;
                    continue;
                }

                if (character == '{')
                    depth++;
                else if (character == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        var candidate = text[start..(index + 1)];
                        try
                        {
                            using var parsed = JsonDocument.Parse(candidate);
                            if (parsed.RootElement.ValueKind == JsonValueKind.Object)
                            {
                                json = candidate;
                                return true;
                            }
                        }
                        catch (JsonException)
                        {
                            break;
                        }
                    }
                }
            }
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

    private static string Limit(string value, int maximum) =>
        value.Length <= maximum ? value : value[..maximum];

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
