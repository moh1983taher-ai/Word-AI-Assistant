using System.Text.Json;

namespace ResearchTools.McpHost.Security;

/// <summary>
/// Performs bounded local checks for common JSON Schema constraints before a tool call.
/// The MCP server remains authoritative and performs its own validation.
/// </summary>
public static class ToolArgumentValidator
{
    public static string? Validate(JsonElement schema, JsonElement arguments)
    {
        if (schema.ValueKind != JsonValueKind.Object)
            return "لم يقدّم الخادم مخططاً صالحاً لمعاملات الأداة.";

        if (arguments.ValueKind != JsonValueKind.Object)
            return "يجب أن تكون معاملات الأداة كائن JSON.";

        return ValidateValue(schema, arguments, "$", 0);
    }

    private static string? ValidateValue(
        JsonElement schema,
        JsonElement value,
        string path,
        int depth)
    {
        if (depth > 12)
            return $"تجاوز المخطط عمق التحقق المسموح عند {path}.";

        if (TryGet(schema, "type", out var type))
        {
            var validType = type.ValueKind switch
            {
                JsonValueKind.String => MatchesType(type.GetString(), value),
                JsonValueKind.Array => type.EnumerateArray().Any(item =>
                    item.ValueKind == JsonValueKind.String && MatchesType(item.GetString(), value)),
                _ => true
            };

            if (!validType)
                return $"نوع غير صالح للمعامل {path}.";
        }

        if (TryGet(schema, "enum", out var allowedValues) &&
            allowedValues.ValueKind == JsonValueKind.Array &&
            !allowedValues.EnumerateArray().Any(item => JsonEquals(item, value)))
            return $"قيمة غير مسموح بها للمعامل {path}.";

        if (value.ValueKind == JsonValueKind.Object)
        {
            if (TryGet(schema, "required", out var required) && required.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in required.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String)
                        continue;

                    var requiredName = item.GetString()!;
                    if (!TryGet(value, requiredName, out _))
                        return $"المعامل المطلوب مفقود: {path}.{requiredName}.";
                }
            }

            if (TryGet(schema, "properties", out var properties) &&
                properties.ValueKind == JsonValueKind.Object)
            {
                var allowAdditional = !TryGet(schema, "additionalProperties", out var additional) ||
                                      additional.ValueKind != JsonValueKind.False;

                foreach (var property in value.EnumerateObject())
                {
                    if (TryGet(properties, property.Name, out var propertySchema))
                    {
                        var error = ValidateValue(
                            propertySchema, property.Value, path + "." + property.Name, depth + 1);
                        if (error is not null)
                            return error;
                    }
                    else if (!allowAdditional)
                    {
                        return $"معامل غير معروف: {path}.{property.Name}.";
                    }
                }
            }
        }

        if (value.ValueKind == JsonValueKind.Array &&
            TryGet(schema, "items", out var itemSchema) &&
            itemSchema.ValueKind == JsonValueKind.Object)
        {
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                var error = ValidateValue(itemSchema, item, $"{path}[{index}]", depth + 1);
                if (error is not null)
                    return error;
                index++;
            }
        }

        return null;
    }

    private static bool MatchesType(string? type, JsonElement value) => type switch
    {
        "object" => value.ValueKind == JsonValueKind.Object,
        "array" => value.ValueKind == JsonValueKind.Array,
        "string" => value.ValueKind == JsonValueKind.String,
        "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
        "number" => value.ValueKind == JsonValueKind.Number,
        "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
        "null" => value.ValueKind == JsonValueKind.Null,
        null or "" => true,
        _ => true
    };

    private static bool JsonEquals(JsonElement left, JsonElement right) =>
        string.Equals(left.GetRawText(), right.GetRawText(), StringComparison.Ordinal);

    private static bool TryGet(JsonElement element, string name, out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.Ordinal))
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
