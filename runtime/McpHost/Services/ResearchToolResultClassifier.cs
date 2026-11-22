using ModelContextProtocol.Protocol;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Distinguishes a completed MCP protocol call from a call that actually returned
/// usable search/read content. Search coverage is counted only for search tools.
/// </summary>
internal static class ResearchToolResultClassifier
{
    internal static bool HasUsableContent(CallToolResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        if (result.IsError == true)
            return false;

        if (result.StructuredContent is { } structured &&
            HasMeaningfulStructuredContent(structured))
            return true;

        foreach (var block in result.Content)
        {
            if (block is TextContentBlock text && HasMeaningfulText(text.Text))
                return true;

            if (block is ResourceLinkBlock)
                return true;
        }

        return false;
    }

    internal static bool IsSearchTool(ToolDescriptor descriptor)
    {
        ArgumentNullException.ThrowIfNull(descriptor);

        return descriptor.Server switch
        {
            "aljam3" => string.Equals(descriptor.Name, "search", StringComparison.Ordinal),
            "shamela" =>
                string.Equals(descriptor.Name, "shamela_find", StringComparison.Ordinal) ||
                string.Equals(descriptor.Name, "shamela_find_scoped", StringComparison.Ordinal),
            _ => false
        };
    }

    /// <summary>
    /// A non-error response from a search operation completes source coverage even
    /// when the search found no matches or returned an empty result set.
    /// </summary>
    internal static bool IsCompletedSearch(ToolDescriptor descriptor, bool isError)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        return !isError && IsSearchTool(descriptor);
    }

    internal static bool HasMeaningfulText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var value = text.Trim();
        if (value.StartsWith("{", StringComparison.Ordinal) ||
            value.StartsWith("[", StringComparison.Ordinal))
        {
            try
            {
                using var document = System.Text.Json.JsonDocument.Parse(value);
                return HasMeaningfulStructuredContent(document.RootElement);
            }
            catch (System.Text.Json.JsonException)
            {
                // Fall through to textual no-result detection for JSON-like snippets.
            }
        }

        var lower = value.ToLowerInvariant();

        string[] noResultPhrases =
        [
            "no results",
            "no matches",
            "0 results",
            "لم يتم العثور على نتائج",
            "لم يتم العثور على نتيجة",
            "لم يعثر على نتائج",
            "لم يُعثر على نتائج",
            "لا توجد نتائج",
            "لا توجد نتيجة",
            "لا يوجد نتائج",
            "لم تسفر عملية البحث عن نتائج"
        ];

        return !noResultPhrases.Any(phrase => lower.Contains(phrase, StringComparison.Ordinal));
    }

    private static bool HasMeaningfulStructuredContent(System.Text.Json.JsonElement value)
    {
        if (value.ValueKind == System.Text.Json.JsonValueKind.Array)
            return value.GetArrayLength() > 0;

        if (value.ValueKind == System.Text.Json.JsonValueKind.Object)
        {
            string[] resultCollectionNames =
            [
                "sources", "results", "matches", "hits", "items", "records", "pages", "books"
            ];

            foreach (var property in value.EnumerateObject())
            {
                if (resultCollectionNames.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                {
                    if (property.Value.ValueKind == System.Text.Json.JsonValueKind.Array)
                        return property.Value.GetArrayLength() > 0;

                    if (property.Value.ValueKind == System.Text.Json.JsonValueKind.Object)
                        return HasMeaningfulStructuredContent(property.Value);
                }
            }

            return value.EnumerateObject().Any(property =>
                property.Value.ValueKind switch
                {
                    System.Text.Json.JsonValueKind.Array =>
                        property.Value.GetArrayLength() > 0,
                    System.Text.Json.JsonValueKind.Object =>
                        HasMeaningfulStructuredContent(property.Value),
                    System.Text.Json.JsonValueKind.String =>
                        !string.IsNullOrWhiteSpace(property.Value.GetString()),
                    System.Text.Json.JsonValueKind.Number or
                    System.Text.Json.JsonValueKind.True or
                    System.Text.Json.JsonValueKind.False => true,
                    _ => false
                });
        }

        return value.ValueKind switch
        {
            System.Text.Json.JsonValueKind.String =>
                HasMeaningfulText(value.GetString()),
            System.Text.Json.JsonValueKind.Number or
            System.Text.Json.JsonValueKind.True or
            System.Text.Json.JsonValueKind.False => true,
            _ => false
        };
    }
}

internal sealed record ToolExecutionResult(string Content, bool Succeeded);
