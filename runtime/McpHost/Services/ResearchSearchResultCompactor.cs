using System.Text;
using System.Text.Json;

namespace ResearchTools.McpHost.Services;

internal static class ResearchSearchResultCompactor
{
    private static readonly HashSet<string> CollectionNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "results", "matches", "hits", "items", "records", "pages", "sources", "documents"
    };

    private static readonly HashSet<string> PaginationFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "next_cursor", "nextCursor", "next_page", "nextPage", "cursor",
        "has_more", "hasMore", "total", "total_count", "totalCount",
        "count", "page", "per_page", "perPage", "limit"
    };

    private static readonly HashSet<string> UsefulFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "id", "book_id", "bookId", "page_id", "pageId", "file_id", "fileId",
        "page_number", "pageNumber", "page", "title", "book_title", "book_name",
        "name", "author", "author_name", "snippet", "excerpt", "preview",
        "text", "content", "url", "uri", "href", "link", "score", "rank",
        "printed_page"
    };

    internal static string Compact(string content)
    {
        if (string.IsNullOrWhiteSpace(content) || !TryParseJson(content, out var document))
            return content;

        using (document)
        {
            var records = new List<JsonElement>();
            CollectRecords(document.RootElement, records, depth: 0);
            if (records.Count == 0)
                return content;

            var builder = new StringBuilder();
            builder.Append("نتائج البحث المنظمة: ").Append(records.Count)
                .AppendLine(" مرجعاً في هذه الاستجابة. المعرفات والبيانات أدناه من مخرجات المكتبة نفسها؛ المقتطفات للاكتشاف فقط.");
            AppendPaginationMetadata(document.RootElement, builder, depth: 0);
            for (var index = 0; index < records.Count; index++)
            {
                builder.Append('[').Append(index + 1).Append("] ");
                var fields = new List<string>();
                CollectFields(records[index], fields, depth: 0);
                builder.AppendLine(fields.Count == 0 ? "(سجل نتيجة)" : string.Join(" | ", fields));
            }

            return builder.ToString();
        }
    }

    private static void AppendPaginationMetadata(JsonElement element, StringBuilder builder, int depth)
    {
        if (depth > 8 || element.ValueKind is not (JsonValueKind.Object or JsonValueKind.Array))
            return;

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (PaginationFields.Contains(property.Name) &&
                    property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False)
                {
                    var value = property.Value.ValueKind == JsonValueKind.String
                        ? property.Value.GetString()
                        : property.Value.GetRawText();
                    if (!string.IsNullOrWhiteSpace(value))
                        builder.Append("بيانات متابعة البحث: ")
                            .Append(property.Name).Append('=').AppendLine(value);
                }
                else if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                {
                    AppendPaginationMetadata(property.Value, builder, depth + 1);
                }
            }
        }
        else
        {
            foreach (var item in element.EnumerateArray())
                AppendPaginationMetadata(item, builder, depth + 1);
        }
    }

    private static void CollectRecords(JsonElement element, List<JsonElement> records, int depth)
    {
        if (depth > 8 || records.Count >= 200)
            return;

        if (element.ValueKind == JsonValueKind.Object)
        {
            var foundCollection = false;
            foreach (var property in element.EnumerateObject())
            {
                if (CollectionNames.Contains(property.Name) &&
                    property.Value.ValueKind == JsonValueKind.Array)
                {
                    foundCollection = true;
                    foreach (var item in property.Value.EnumerateArray())
                    {
                        if (item.ValueKind == JsonValueKind.Object && LooksLikeRecord(item))
                        {
                            records.Add(item.Clone());
                            if (records.Count >= 200) return;
                        }
                        else
                        {
                            CollectRecords(item, records, depth + 1);
                        }
                    }
                }
            }

            if (foundCollection)
                return;

            foreach (var property in element.EnumerateObject())
                if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                    CollectRecords(property.Value, records, depth + 1);
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Object && LooksLikeRecord(item))
                {
                    records.Add(item.Clone());
                    if (records.Count >= 200) return;
                }
                else
                {
                    CollectRecords(item, records, depth + 1);
                }
            }
        }
    }

    private static bool LooksLikeRecord(JsonElement element)
    {
        if (element.ValueKind != JsonValueKind.Object)
            return false;

        var usefulCount = 0;
        foreach (var property in element.EnumerateObject())
        {
            if (UsefulFields.Contains(property.Name) &&
                property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number)
                usefulCount++;
            if (usefulCount >= 2)
                return true;
        }

        return usefulCount >= 1;
    }

    private static void CollectFields(JsonElement element, List<string> fields, int depth)
    {
        if (depth > 2 || element.ValueKind != JsonValueKind.Object)
            return;

        foreach (var property in element.EnumerateObject())
        {
            if (UsefulFields.Contains(property.Name))
            {
                var value = property.Value.ValueKind switch
                {
                    JsonValueKind.String => property.Value.GetString(),
                    JsonValueKind.Number => property.Value.GetRawText(),
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    _ => null
                };

                if (!string.IsNullOrWhiteSpace(value))
                {
                    var maximum = IsLinkField(property.Name) ? 500 :
                        IsTextField(property.Name) ? 180 : 100;
                    value = NormalizeWhitespace(value);
                    if (value.Length > maximum)
                        value = value[..maximum] + "…";
                    fields.Add(property.Name + "=" + value);
                }
            }
            else if (property.Value.ValueKind == JsonValueKind.Object &&
                     (property.Name.Equals("book", StringComparison.OrdinalIgnoreCase) ||
                      property.Name.Equals("author", StringComparison.OrdinalIgnoreCase) ||
                      property.Name.Equals("page", StringComparison.OrdinalIgnoreCase) ||
                      property.Name.Equals("document", StringComparison.OrdinalIgnoreCase)))
            {
                CollectFields(property.Value, fields, depth + 1);
            }
        }
    }

    private static bool TryParseJson(string content, out JsonDocument document)
    {
        document = null!;
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

    private static bool IsLinkField(string name) =>
        name.Equals("url", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("uri", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("href", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("link", StringComparison.OrdinalIgnoreCase);

    private static bool IsTextField(string name) =>
        name.Equals("snippet", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("excerpt", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("preview", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("text", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("content", StringComparison.OrdinalIgnoreCase);

    private static string NormalizeWhitespace(string value) =>
        string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
