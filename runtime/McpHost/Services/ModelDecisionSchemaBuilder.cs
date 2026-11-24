using System.Text.Json;
using System.Text.Json.Nodes;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Builds a per-turn strict JSON Schema for the decision envelope. MCP argument
/// schemas are projected into the strict subset accepted by structured-output APIs.
/// Optional properties become required nullable properties; local JSON references
/// are inlined and null optional values are removed before MCP execution.
/// </summary>
internal static class ModelDecisionSchemaBuilder
{
    internal static JsonElement BuildResponseFormat(
        IReadOnlyList<ToolDescriptor> tools,
        bool allowFinal,
        bool allowToolCalls = true)
    {
        ArgumentNullException.ThrowIfNull(tools);

        if (!allowFinal && !allowToolCalls)
            throw new ArgumentException("At least one decision type must be allowed.");

        if (allowToolCalls && tools.Count == 0)
            throw new ArgumentException("Tool-call decisions require at least one available MCP tool.", nameof(tools));

        var properties = new JsonObject();
        var required = new JsonArray();

        var decisionTypes = new List<string>();
        if (allowToolCalls)
            decisionTypes.Add("tool_call");
        if (allowFinal)
            decisionTypes.Add("final");

        var serverValues = tools
            .Select(tool => tool.Server)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var toolValues = tools
            .Select(tool => tool.Name)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .ToList();

        if (allowFinal)
        {
            serverValues.Insert(0, string.Empty);
            toolValues.Insert(0, string.Empty);
        }

        properties["type"] = StringEnum(decisionTypes);
        properties["server"] = StringEnum(serverValues);
        properties["tool"] = StringEnum(toolValues);

        var argumentVariants = new JsonArray();
        if (allowToolCalls)
        {
            foreach (var tool in tools)
            {
                var argumentSchema = NormalizeSchema(
                    tool.InputSchema,
                    nullable: false,
                    rootSchema: tool.InputSchema,
                    activeReferences: new HashSet<string>(StringComparer.Ordinal));

                if (argumentSchema is JsonObject argumentObject)
                {
                    var toolIdentity = $"{tool.Server}/{tool.Name}";
                    var description = $"Arguments for MCP tool {toolIdentity}";
                    if (!string.IsNullOrWhiteSpace(tool.ServerLabel))
                        description += $" ({tool.ServerLabel})";
                    if (!string.IsNullOrWhiteSpace(tool.Description))
                        description += $": {tool.Description.Trim()}";

                    var schemaDescription = argumentObject["description"]?.GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(schemaDescription))
                        description += $" Schema details: {schemaDescription.Trim()}";

                    argumentObject["description"] = description;
                }

                argumentVariants.Add(argumentSchema);
            }
        }

        if (allowFinal)
            argumentVariants.Add(new JsonObject { ["type"] = "null" });

        properties["arguments"] = argumentVariants.Count == 1
            ? argumentVariants[0]!.DeepClone()
            : new JsonObject { ["anyOf"] = argumentVariants };

        properties["answer"] = allowFinal
            ? new JsonObject { ["type"] = "string" }
            : new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray(JsonValue.Create(string.Empty))
            };

        foreach (var name in new[] { "type", "server", "tool", "arguments", "answer" })
            required.Add(name);

        var schema = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = properties,
            ["required"] = required,
            ["additionalProperties"] = false
        };

        var responseFormat = new JsonObject
        {
            ["type"] = "json_schema",
            ["json_schema"] = new JsonObject
            {
                ["name"] = "mcp_research_decision",
                ["strict"] = true,
                ["schema"] = schema
            }
        };

        using var document = JsonDocument.Parse(responseFormat.ToJsonString());
        return document.RootElement.Clone();
    }

    private static JsonObject StringEnum(IEnumerable<string> values)
    {
        var items = new JsonArray();
        foreach (var value in values.Distinct(StringComparer.Ordinal))
            items.Add(value);

        return new JsonObject
        {
            ["type"] = "string",
            ["enum"] = items
        };
    }

    private static JsonNode NormalizeSchema(
        JsonElement source,
        bool nullable,
        JsonElement rootSchema,
        HashSet<string> activeReferences)
    {
        if (source.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException("MCP tool argument schemas must be JSON objects.");

        var normalized = NormalizeCore(source, rootSchema, activeReferences);
        if (!nullable || ContainsNullAlternative(normalized))
            return normalized;

        return new JsonObject
        {
            ["anyOf"] = new JsonArray(
                normalized,
                new JsonObject { ["type"] = "null" })
        };
    }

    private static JsonNode NormalizeCore(
        JsonElement source,
        JsonElement rootSchema,
        HashSet<string> activeReferences)
    {
        if (TryGet(source, "$ref", out var referenceElement))
        {
            if (referenceElement.ValueKind != JsonValueKind.String ||
                string.IsNullOrWhiteSpace(referenceElement.GetString()))
                throw new InvalidOperationException("An MCP argument schema contains an invalid $ref.");

            var reference = referenceElement.GetString()!;
            if (!activeReferences.Add(reference))
                throw new InvalidOperationException($"Circular MCP JSON Schema reference is not supported: {reference}.");

            try
            {
                var resolved = ResolveLocalReference(rootSchema, reference);
                var normalizedReference = NormalizeSchema(
                    resolved,
                    nullable: false,
                    rootSchema,
                    activeReferences);

                if (normalizedReference is JsonObject normalizedObject)
                {
                    CopyStringProperty(source, normalizedObject, "description");
                    CopyStringProperty(source, normalizedObject, "title");
                }

                return normalizedReference;
            }
            finally
            {
                activeReferences.Remove(reference);
            }
        }

        if (TryGet(source, "anyOf", out var anyOf) && anyOf.ValueKind == JsonValueKind.Array)
            return NormalizeAlternatives(anyOf, rootSchema, activeReferences);

        if (TryGet(source, "oneOf", out var oneOf) && oneOf.ValueKind == JsonValueKind.Array)
            return NormalizeAlternatives(oneOf, rootSchema, activeReferences);

        if (TryGet(source, "type", out var typeElement) &&
            typeElement.ValueKind == JsonValueKind.Array)
        {
            var alternatives = new JsonArray();
            foreach (var item in typeElement.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.String)
                    continue;

                var typeName = item.GetString();
                if (string.Equals(typeName, "null", StringComparison.Ordinal))
                {
                    alternatives.Add(new JsonObject { ["type"] = "null" });
                    continue;
                }

                if (string.IsNullOrWhiteSpace(typeName))
                    continue;

                alternatives.Add(NormalizeByType(
                    source,
                    typeName,
                    rootSchema,
                    activeReferences));
            }

            if (alternatives.Count == 0)
                throw new InvalidOperationException("An MCP argument schema contains an empty type union.");

            return alternatives.Count == 1
                ? alternatives[0]!.DeepClone()
                : new JsonObject { ["anyOf"] = alternatives };
        }

        var type = TryGet(source, "type", out typeElement) &&
                   typeElement.ValueKind == JsonValueKind.String
            ? typeElement.GetString()
            : TryGet(source, "properties", out _)
                ? "object"
                : null;

        if (string.IsNullOrWhiteSpace(type))
        {
            if (TryGet(source, "enum", out var enumValues) &&
                enumValues.ValueKind == JsonValueKind.Array)
            {
                var inferred = InferType(enumValues);
                return NormalizeByType(source, inferred, rootSchema, activeReferences);
            }

            throw new InvalidOperationException(
                "An MCP argument schema has no supported JSON type or resolvable local reference.");
        }

        return NormalizeByType(source, type, rootSchema, activeReferences);
    }

    private static JsonNode NormalizeAlternatives(
        JsonElement alternatives,
        JsonElement rootSchema,
        HashSet<string> activeReferences)
    {
        var result = new JsonArray();
        foreach (var alternative in alternatives.EnumerateArray())
        {
            result.Add(NormalizeSchema(
                alternative,
                nullable: false,
                rootSchema,
                activeReferences));
        }

        if (result.Count == 0)
            throw new InvalidOperationException("An MCP argument schema contains an empty union.");

        return result.Count == 1
            ? result[0]!.DeepClone()
            : new JsonObject { ["anyOf"] = result };
    }

    private static JsonObject NormalizeByType(
        JsonElement source,
        string type,
        JsonElement rootSchema,
        HashSet<string> activeReferences)
    {
        var result = new JsonObject();

        switch (type)
        {
            case "object":
            {
                result["type"] = "object";
                var properties = new JsonObject();
                var requiredNames = new HashSet<string>(StringComparer.Ordinal);

                if (TryGet(source, "required", out var required) &&
                    required.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in required.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.String &&
                            !string.IsNullOrWhiteSpace(item.GetString()))
                            requiredNames.Add(item.GetString()!);
                    }
                }

                if (TryGet(source, "properties", out var sourceProperties) &&
                    sourceProperties.ValueKind == JsonValueKind.Object)
                {
                    foreach (var property in sourceProperties.EnumerateObject())
                    {
                        properties[property.Name] = NormalizeSchema(
                            property.Value,
                            nullable: !requiredNames.Contains(property.Name),
                            rootSchema,
                            activeReferences);
                    }
                }

                var requiredProperties = new JsonArray();
                foreach (var propertyName in properties.Select(property => property.Key))
                    requiredProperties.Add(propertyName);

                result["properties"] = properties;
                result["required"] = requiredProperties;
                result["additionalProperties"] = false;
                break;
            }

            case "array":
                result["type"] = "array";
                if (TryGet(source, "items", out var items))
                {
                    result["items"] = NormalizeSchema(
                        items,
                        nullable: false,
                        rootSchema,
                        activeReferences);
                }
                else
                {
                    result["items"] = new JsonObject { ["type"] = "string" };
                }
                break;

            case "string":
            case "integer":
            case "number":
            case "boolean":
            case "null":
                result["type"] = type;
                break;

            default:
                throw new InvalidOperationException(
                    $"Unsupported MCP argument JSON type in strict response schema: {type}.");
        }

        if (TryGet(source, "enum", out var enumElement) &&
            enumElement.ValueKind == JsonValueKind.Array)
        {
            result["enum"] = JsonNode.Parse(enumElement.GetRawText());
        }
        else if (TryGet(source, "const", out var constElement))
        {
            result["enum"] = new JsonArray(JsonNode.Parse(constElement.GetRawText()));
        }

        CopyStringProperty(source, result, "description");
        CopyStringProperty(source, result, "title");
        return result;
    }

    private static JsonElement ResolveLocalReference(JsonElement rootSchema, string reference)
    {
        if (!reference.StartsWith("#/", StringComparison.Ordinal))
            throw new InvalidOperationException(
                $"External MCP JSON Schema references are not supported: {reference}.");

        var current = rootSchema;
        var segments = reference[2..].Split('/');
        foreach (var rawSegment in segments)
        {
            var segment = rawSegment.Replace("~1", "/", StringComparison.Ordinal)
                .Replace("~0", "~", StringComparison.Ordinal);

            if (current.ValueKind != JsonValueKind.Object ||
                !TryGet(current, segment, out var resolved))
            {
                throw new InvalidOperationException(
                    $"MCP JSON Schema reference could not be resolved: {reference}.");
            }

            current = resolved;
        }

        return current;
    }

    private static bool ContainsNullAlternative(JsonNode node)
    {
        if (node is not JsonObject obj ||
            !obj.TryGetPropertyValue("anyOf", out var alternatives) ||
            alternatives is not JsonArray array)
            return false;

        return array.Any(item =>
            item is JsonObject branch &&
            branch.TryGetPropertyValue("type", out var type) &&
            string.Equals(type?.GetValue<string>(), "null", StringComparison.Ordinal));
    }

    private static string InferType(JsonElement values)
    {
        foreach (var item in values.EnumerateArray())
        {
            return item.ValueKind switch
            {
                JsonValueKind.String => "string",
                JsonValueKind.Number when item.TryGetInt64(out _) => "integer",
                JsonValueKind.Number => "number",
                JsonValueKind.True or JsonValueKind.False => "boolean",
                JsonValueKind.Object => "object",
                JsonValueKind.Array => "array",
                JsonValueKind.Null => "null",
                _ => throw new InvalidOperationException("Cannot infer an MCP argument enum type.")
            };
        }

        throw new InvalidOperationException("An MCP argument enum cannot be empty.");
    }

    private static void CopyStringProperty(JsonElement source, JsonObject destination, string name)
    {
        if (TryGet(source, name, out var value) && value.ValueKind == JsonValueKind.String)
            destination[name] = value.GetString();
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
