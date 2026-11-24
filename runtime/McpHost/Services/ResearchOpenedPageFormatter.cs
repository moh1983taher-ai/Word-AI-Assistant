using System.Text;
using System.Text.Json;

namespace ResearchTools.McpHost.Services;

internal static class ResearchOpenedPageFormatter
{
    private static readonly HashSet<string> TextFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "content", "text", "body", "page_text", "pageText", "plain_text", "plainText",
        "markdown", "html", "transcription", "excerpt"
    };

    private static readonly HashSet<string> MetadataFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "id", "book_id", "bookId", "page_id", "pageId", "file_id", "fileId",
        "page_number", "pageNumber", "printed_page", "title", "book_title",
        "book_name", "name", "author", "author_name", "url", "uri", "href", "link"
    };

    internal static string Format(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
            return content;

        if (!TryParseJson(content, out var document, out var jsonEnd))
            return WrapLongText(content);

        using (document)
        {
            var metadata = new List<string>();
            var textParts = new List<string>();
            CollectMetadata(document.RootElement, metadata, depth: 0);
            CollectText(document.RootElement, textParts, depth: 0);
            if (jsonEnd < content.Length)
            {
                var trailingText = content[jsonEnd..].Trim();
                if (!string.IsNullOrWhiteSpace(trailingText))
                    textParts.Add(trailingText);
            }

            if (textParts.Count == 0)
                return WrapLongText(content);

            var builder = new StringBuilder();
            if (metadata.Count > 0)
                builder.Append("بيانات الصفحة المفتوحة: ")
                    .AppendLine(string.Join(" | ", metadata.Distinct(StringComparer.OrdinalIgnoreCase)));
            builder.AppendLine("النص المستخرج من الصفحة المفتوحة:");
            foreach (var part in textParts)
            {
                var normalized = part.Replace("\r\n", "\n", StringComparison.Ordinal)
                    .Replace('\r', '\n').Trim();
                if (normalized.Length == 0)
                    continue;
                builder.AppendLine(WrapLongText(normalized));
            }

            return builder.ToString();
        }
    }

    private static void CollectMetadata(JsonElement element, List<string> values, int depth)
    {
        if (depth > 8)
            return;

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (MetadataFields.Contains(property.Name) &&
                    property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number)
                {
                    var value = property.Value.ValueKind == JsonValueKind.String
                        ? property.Value.GetString()
                        : property.Value.GetRawText();
                    if (!string.IsNullOrWhiteSpace(value))
                        values.Add(property.Name + "=" + NormalizeWhitespace(value));
                }

                if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                    CollectMetadata(property.Value, values, depth + 1);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                CollectMetadata(item, values, depth + 1);
        }
    }

    private static void CollectText(JsonElement element, List<string> values, int depth)
    {
        if (depth > 10)
            return;

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (TextFields.Contains(property.Name) &&
                    property.Value.ValueKind == JsonValueKind.String)
                {
                    var value = property.Value.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                        values.Add(value);
                }
                else if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                {
                    CollectText(property.Value, values, depth + 1);
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                CollectText(item, values, depth + 1);
        }
    }

    private static string WrapLongText(string value)
    {
        const int maximumChunkCharacters = 600;
        var paragraphs = value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n');

        var output = new StringBuilder(value.Length + Math.Max(16, paragraphs.Length));
        foreach (var paragraph in paragraphs)
        {
            var remaining = paragraph.Trim();
            while (remaining.Length > maximumChunkCharacters)
            {
                var splitAt = remaining.LastIndexOf(' ', maximumChunkCharacters);
                if (splitAt < maximumChunkCharacters / 2)
                    splitAt = maximumChunkCharacters;
                output.AppendLine(remaining[..splitAt].TrimEnd());
                remaining = remaining[splitAt..].TrimStart();
            }

            if (remaining.Length > 0)
                output.AppendLine(remaining);
            else
                output.AppendLine();
        }

        return output.ToString().TrimEnd();
    }

    private static bool TryParseJson(string content, out JsonDocument document, out int jsonEnd)
    {
        document = null!;
        jsonEnd = 0;
        for (var start = 0; start < content.Length; start++)
        {
            if (content[start] is not ('{' or '['))
                continue;

            var stack = new Stack<char>();
            var inString = false;
            var escaped = false;
            for (var index = start; index < content.Length; index++)
            {
                var character = content[index];
                if (inString)
                {
                    if (escaped) { escaped = false; continue; }
                    if (character == '\\') { escaped = true; continue; }
                    if (character == '"') inString = false;
                    continue;
                }

                if (character == '"') { inString = true; continue; }
                if (character is '{' or '[') stack.Push(character);
                else if (character is '}' or ']')
                {
                    if (stack.Count == 0) break;
                    var expected = character == '}' ? '{' : '[';
                    if (stack.Pop() != expected) break;
                    if (stack.Count == 0)
                    {
                        try
                        {
                            document = JsonDocument.Parse(content[start..(index + 1)]);
                            jsonEnd = index + 1;
                            return document.RootElement.ValueKind is JsonValueKind.Object or JsonValueKind.Array;
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

    private static string NormalizeWhitespace(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
