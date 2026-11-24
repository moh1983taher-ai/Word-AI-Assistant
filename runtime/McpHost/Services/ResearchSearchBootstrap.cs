using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Security;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Creates the first search invocation for each user-selected MCP source.
/// Search coverage is a host responsibility, not a probabilistic model decision.
/// </summary>
internal static class ResearchSearchBootstrap
{
    private const int MaximumQueryCharacters = 1_200;

    private static readonly string[] QueryPropertyNames =
    [
        "query",
        "keywords",
        "q",
        "search_query",
        "search_term",
        "keyword",
        "text"
    ];

    internal static IReadOnlyList<ResearchSearchInvocation> BuildInitialPlan(
        IReadOnlyList<ToolDescriptor> availableTools,
        IReadOnlyList<string> selectedServerIds,
        string query) =>
        BuildInitialPlan(
            availableTools,
            selectedServerIds,
            [new ResearchQuery(query, "استعلام البحث")]);

    internal static IReadOnlyList<ResearchSearchInvocation> BuildInitialPlan(
        IReadOnlyList<ToolDescriptor> availableTools,
        IReadOnlyList<string> selectedServerIds,
        IReadOnlyList<ResearchQuery> researchQueries)
    {
        ArgumentNullException.ThrowIfNull(availableTools);
        ArgumentNullException.ThrowIfNull(selectedServerIds);
        ArgumentNullException.ThrowIfNull(researchQueries);

        if (researchQueries.Count == 0 || researchQueries.Count > 8)
            throw new InvalidOperationException("يجب أن تتضمن خطة البحث استعلاماً واحداً على الأقل، وألا تتجاوز ثمانية استعلامات.");

        var sources = ResearchSourceSelection.Normalize(selectedServerIds);
        if (sources.Count == 0)
            throw new InvalidOperationException("لم تُحدد مكتبة صالحة لتنفيذ خطة البحث.");

        var normalizedQueries = new List<ResearchQuery>();
        var seenQueries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in researchQueries)
        {
            var query = (item?.Query ?? string.Empty).Trim();
            if (query.Length > MaximumQueryCharacters)
                query = query[..MaximumQueryCharacters];
            if (string.IsNullOrWhiteSpace(query))
                throw new InvalidOperationException("تتضمن خطة البحث استعلاماً فارغاً.");

            var normalized = string.Join(' ', query.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            if (!seenQueries.Add(normalized))
                continue;

            normalizedQueries.Add(new ResearchQuery(
                query,
                (item.Purpose ?? string.Empty).Trim()));
        }

        if (normalizedQueries.Count == 0)
            throw new InvalidOperationException("لا تتضمن خطة البحث استعلامات مستقلة قابلة للتنفيذ.");

        var plan = new List<ResearchSearchInvocation>();
        for (var queryIndex = 0; queryIndex < normalizedQueries.Count; queryIndex++)
        {
            var researchQuery = normalizedQueries[queryIndex];
            foreach (var source in sources)
            {
                var candidates = availableTools
                    .Where(tool =>
                        string.Equals(tool.Server, source, StringComparison.Ordinal) &&
                        ResearchToolResultClassifier.IsSearchTool(tool))
                    .OrderBy(tool => SearchPreference(source, tool.Name))
                    .ToArray();

                if (candidates.Length == 0)
                {
                    throw new InvalidOperationException(
                        $"لا تتوفر أداة بحث معروفة للمكتبة المختارة {ResearchSourceSelection.GetLabel(source)} ({source}).");
                }

                string? lastError = null;
                ResearchSearchInvocation? invocation = null;

                foreach (var candidate in candidates)
                {
                    if (TryBuildArguments(candidate, researchQuery.Query, out var arguments, out var error))
                    {
                        invocation = new ResearchSearchInvocation(
                            candidate,
                            arguments,
                            researchQuery.Query,
                            researchQuery.Purpose,
                            queryIndex + 1);
                        break;
                    }

                    lastError = error;
                }

                if (invocation is null)
                {
                    throw new InvalidOperationException(
                        $"تعذر بناء معاملات بحث صحيحة للمكتبة {ResearchSourceSelection.GetLabel(source)}. {lastError}");
                }

                plan.Add(invocation);
            }
        }

        return plan;
    }

    private static int SearchPreference(string source, string toolName)
    {
        if (source == "aljam3")
            return toolName == "search" ? 0 : 10;

        return toolName switch
        {
            "shamela_find" => 0,
            "shamela_find_scoped" => 1,
            _ => 10
        };
    }

    private static bool TryBuildArguments(
        ToolDescriptor tool,
        string query,
        out IReadOnlyDictionary<string, object?> arguments,
        out string error)
    {
        var values = new Dictionary<string, object?>(StringComparer.Ordinal);
        arguments = values;
        error = string.Empty;

        var schema = tool.InputSchema;
        if (schema.ValueKind != JsonValueKind.Object ||
            !TryGetProperty(schema, "properties", out var properties) ||
            properties.ValueKind != JsonValueKind.Object)
        {
            error = $"{tool.Server}/{tool.Name}: لا يحتوي مخطط الأداة على خصائص معاملات قابلة للقراءة.";
            return false;
        }

        var queryProperty = QueryPropertyNames
            .Select(name => FindProperty(properties, name))
            .FirstOrDefault(property => property.HasValue);

        if (queryProperty is null ||
            !TryCreateQueryValue(queryProperty.Value.Value, query, out var queryValue))
        {
            error =
                $"{tool.Server}/{tool.Name}: لم يُعثر على معامل بحث نصي معروف (query أو keywords أو q أو ما يعادلها) في مخطط الأداة.";
            return false;
        }

        values[queryProperty.Value.Key] = queryValue;

        if (TryGetProperty(schema, "required", out var required) &&
            required.ValueKind == JsonValueKind.Array)
        {
            foreach (var requiredProperty in required.EnumerateArray())
            {
                if (requiredProperty.ValueKind != JsonValueKind.String)
                    continue;

                var name = requiredProperty.GetString();
                if (string.IsNullOrWhiteSpace(name) || values.ContainsKey(name))
                    continue;

                var property = FindProperty(properties, name);
                if (property is null)
                {
                    error = $"{tool.Server}/{tool.Name}: المعامل الإلزامي {name} غير معرّف في المخطط.";
                    return false;
                }

                if (TryGetProperty(property.Value.Value, "default", out var defaultValue))
                {
                    values[name] = defaultValue.Clone();
                    continue;
                }

                if (AllowsNull(property.Value.Value))
                {
                    values[name] = null;
                    continue;
                }

                error =
                    $"{tool.Server}/{tool.Name}: المعامل الإلزامي {name} لا يملك قيمة افتراضية آمنة؛ لن يخمّن المضيف قيمته.";
                return false;
            }
        }

        var jsonArguments = JsonSerializer.SerializeToElement(values, JsonDefaults.Web);
        var validationError = ToolArgumentValidator.Validate(schema, jsonArguments);
        if (validationError is not null)
        {
            error = $"{tool.Server}/{tool.Name}: {validationError}";
            return false;
        }

        return true;
    }

    private static KeyValuePair<string, JsonElement>? FindProperty(
        JsonElement properties,
        string name)
    {
        foreach (var property in properties.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                return new KeyValuePair<string, JsonElement>(property.Name, property.Value);
        }

        return null;
    }

    private static bool TryCreateQueryValue(
        JsonElement propertySchema,
        string query,
        out object? value)
    {
        value = null;
        var declaredType = ReadPrimaryType(propertySchema);

        if (declaredType == "string")
        {
            value = query;
            return true;
        }

        if (declaredType == "array" &&
            TryGetProperty(propertySchema, "items", out var items) &&
            ReadPrimaryType(items) == "string")
        {
            value = new[] { query };
            return true;
        }

        return false;
    }

    private static string? ReadPrimaryType(JsonElement schema)
    {
        if (TryGetProperty(schema, "type", out var type))
        {
            if (type.ValueKind == JsonValueKind.String)
                return type.GetString();

            if (type.ValueKind == JsonValueKind.Array)
            {
                return type.EnumerateArray()
                    .Select(item => item.ValueKind == JsonValueKind.String ? item.GetString() : null)
                    .FirstOrDefault(item => item is not null && item != "null");
            }
        }

        if (TryGetProperty(schema, "anyOf", out var anyOf) &&
            anyOf.ValueKind == JsonValueKind.Array)
        {
            foreach (var option in anyOf.EnumerateArray())
            {
                var candidate = ReadPrimaryType(option);
                if (candidate is not null && candidate != "null")
                    return candidate;
            }
        }

        return null;
    }

    private static bool AllowsNull(JsonElement schema)
    {
        if (TryGetProperty(schema, "type", out var type))
        {
            if (type.ValueKind == JsonValueKind.String &&
                string.Equals(type.GetString(), "null", StringComparison.Ordinal))
                return true;

            if (type.ValueKind == JsonValueKind.Array &&
                type.EnumerateArray().Any(item =>
                    item.ValueKind == JsonValueKind.String &&
                    string.Equals(item.GetString(), "null", StringComparison.Ordinal)))
                return true;
        }

        if (TryGetProperty(schema, "anyOf", out var anyOf) &&
            anyOf.ValueKind == JsonValueKind.Array)
        {
            return anyOf.EnumerateArray().Any(AllowsNull);
        }

        return false;
    }

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
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

internal sealed record ResearchSearchInvocation(
    ToolDescriptor Tool,
    IReadOnlyDictionary<string, object?> Arguments,
    string Query,
    string Purpose,
    int QueryIndex);
