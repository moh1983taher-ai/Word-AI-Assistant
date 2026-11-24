using System.Text.Json;
using System.Text.RegularExpressions;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Tracks page references returned by selected MCP search tools and prevents
/// page-reading tools from opening identifiers absent from the current task's search results.
/// </summary>
internal sealed class McpEvidenceProvenance
{
    private static readonly string[] BookKeys = ["book_id", "bookId", "bookid"];
    private static readonly string[] PageKeys = ["page_id", "pageId", "pageid", "page", "start_page_id", "startPageId"];

    private static readonly Regex BookPagePath = new(
        """(?i)/(?:books?)/(?<book>\d+)/(?:pages?/)?(?<page>\d+)(?:[/?#\s]|$)""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex BookPagePair = new(
        """(?i)\bbook[_\s-]?id["'\s:=]+(?<book>\d+)[\s\S]{0,240}?\bpage[_\s-]?id["'\s:=]+(?<page>\d+)""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ColonPair = new(
        """^\s*(?<book>\d+)\s*:\s*(?<page>\d+)\s*$""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly HashSet<string> _bookIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _pageIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _opaqueIds = new(StringComparer.Ordinal);
    private readonly HashSet<string> _bookPagePairs = new(StringComparer.Ordinal);

    internal int SearchReferenceCount =>
        _bookPagePairs.Count + _pageIds.Count + _opaqueIds.Count;

    internal int VerifiedOpenCount { get; private set; }

    internal static bool IsReferenceReadTool(ToolDescriptor descriptor)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        return descriptor.Server switch
        {
            "aljam3" => string.Equals(descriptor.Name, "get_page", StringComparison.Ordinal),
            "shamela" =>
                string.Equals(descriptor.Name, "shamela_open", StringComparison.Ordinal) ||
                string.Equals(descriptor.Name, "shamela_open_many", StringComparison.Ordinal),
            _ => false
        };
    }

    internal int RecordSearchResult(
        ToolDescriptor descriptor,
        JsonElement? structuredContent,
        IEnumerable<string>? textBlocks,
        bool isError)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        if (isError || !ResearchToolResultClassifier.IsSearchTool(descriptor))
            return 0;

        var before = SearchReferenceCount;

        if (structuredContent is { } structured &&
            structured.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null)
        {
            VisitSearchPayload(structured, inheritedBookId: null);
        }

        if (textBlocks is not null)
        {
            foreach (var text in textBlocks)
                VisitSearchText(text);
        }

        return SearchReferenceCount - before;
    }

    internal string? ValidateReferenceRead(
        ToolDescriptor descriptor,
        IReadOnlyDictionary<string, object?> arguments)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        ArgumentNullException.ThrowIfNull(arguments);

        if (!IsReferenceReadTool(descriptor))
            return null;

        if (SearchReferenceCount == 0)
        {
            return "لا توجد نتيجة بحث سابقة من هذه المكتبة تحمل معرّف صفحة يمكن التحقق منه؛ أعد البحث قبل فتح الصفحة.";
        }

        var input = JsonSerializer.SerializeToElement(arguments);
        var requested = new List<Reference>();
        VisitOpenArguments(input, requested, inheritedBookId: null);

        if (requested.Count == 0)
        {
            return "معاملات فتح الصفحة لا تتضمن معرّف كتاب أو صفحة معروفاً؛ لا تخمّن المعرّفات، واستخدم معرّفاً وارداً في نتيجة البحث.";
        }

        if (requested.Any(reference => !WasReturnedBySearch(reference)))
        {
            return "المعرّف المطلوب فتحه غير موجود في نتائج البحث السابقة لهذه المكتبة؛ اختر نتيجة من المخرجات الفعلية أو أعد البحث. لا تفتح موضعاً مستنتجاً أو منقولاً من مصدر آخر.";
        }

        return null;
    }

    internal void RecordVerifiedOpen(ToolDescriptor descriptor, bool succeeded)
    {
        if (succeeded && IsReferenceReadTool(descriptor))
            VerifiedOpenCount++;
    }

    private bool WasReturnedBySearch(Reference reference)
    {
        if (reference.BookId is not null && reference.PageId is not null)
            return _bookPagePairs.Contains(Pair(reference.BookId, reference.PageId));

        if (reference.BookId is not null)
            return _bookIds.Contains(reference.BookId);

        if (reference.PageId is not null)
            return _pageIds.Contains(reference.PageId) ||
                   _opaqueIds.Contains(reference.PageId);

        return reference.OpaqueId is not null &&
               (_opaqueIds.Contains(reference.OpaqueId) ||
                _pageIds.Contains(reference.OpaqueId));
    }

    private void VisitSearchPayload(JsonElement element, string? inheritedBookId)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var bookId = ReadScalar(element, BookKeys) ?? inheritedBookId;
                var pageId = ReadScalar(element, PageKeys);

                if (bookId is not null)
                    _bookIds.Add(bookId);

                if (bookId is not null && pageId is not null)
                    _bookPagePairs.Add(Pair(bookId, pageId));
                else if (pageId is not null)
                    _pageIds.Add(pageId);

                if (TryGetProperty(element, "id", out var idElement))
                    AddIdentifier(ReadScalar(idElement), bookId);

                foreach (var property in element.EnumerateObject())
                {
                    if (IsLinkKey(property.Name) && property.Value.ValueKind == JsonValueKind.String)
                        VisitSearchText(property.Value.GetString());

                    if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        VisitSearchPayload(property.Value, bookId);
                }
                break;
            }

            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    VisitSearchPayload(item, inheritedBookId);
                break;

            case JsonValueKind.String:
                VisitSearchText(element.GetString());
                break;
        }
    }

    private static void VisitOpenArguments(
        JsonElement element,
        List<Reference> references,
        string? inheritedBookId)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var bookId = ReadScalar(element, BookKeys) ?? inheritedBookId;
                var pageId = ReadScalar(element, PageKeys);

                if (bookId is not null || pageId is not null)
                    references.Add(new Reference(bookId, pageId, null));

                if (TryGetProperty(element, "id", out var idElement))
                    AddTargetIdentifier(ReadScalar(idElement), bookId, references);

                foreach (var property in element.EnumerateObject())
                {
                    if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        VisitOpenArguments(property.Value, references, bookId);
                    else if (property.Value.ValueKind == JsonValueKind.String &&
                             IsIdentifierCollectionKey(property.Name))
                        AddTargetIdentifier(property.Value.GetString(), bookId, references);
                }
                break;
            }

            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    VisitOpenArguments(item, references, inheritedBookId);
                break;

            case JsonValueKind.String:
                AddTargetIdentifier(element.GetString(), inheritedBookId, references);
                break;
        }
    }

    private void VisitSearchText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        foreach (Match match in BookPagePath.Matches(text))
            _bookPagePairs.Add(Pair(match.Groups["book"].Value, match.Groups["page"].Value));

        foreach (Match match in BookPagePair.Matches(text))
            _bookPagePairs.Add(Pair(match.Groups["book"].Value, match.Groups["page"].Value));

        var candidate = text.Trim();
        var colonPair = ColonPair.Match(candidate);
        if (colonPair.Success)
            _bookPagePairs.Add(Pair(colonPair.Groups["book"].Value, colonPair.Groups["page"].Value));

        var firstBrace = candidate.IndexOf('{');
        var lastBrace = candidate.LastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace)
            candidate = candidate[firstBrace..(lastBrace + 1)];

        try
        {
            using var document = JsonDocument.Parse(candidate);
            VisitSearchPayload(document.RootElement, inheritedBookId: null);
        }
        catch (JsonException)
        {
            // Search servers may return Markdown; the URL and key-value patterns above still apply.
        }
    }

    private void AddIdentifier(string? id, string? bookId)
    {
        if (string.IsNullOrWhiteSpace(id))
            return;

        var normalized = id.Trim();
        var pair = ColonPair.Match(normalized);
        if (pair.Success)
        {
            _bookPagePairs.Add(Pair(pair.Groups["book"].Value, pair.Groups["page"].Value));
            return;
        }

        if (bookId is not null)
            _bookPagePairs.Add(Pair(bookId, normalized));
        else
            _opaqueIds.Add(normalized);
    }

    private static void AddTargetIdentifier(
        string? id,
        string? bookId,
        List<Reference> references)
    {
        if (string.IsNullOrWhiteSpace(id))
            return;

        var normalized = id.Trim();
        var pair = ColonPair.Match(normalized);
        if (pair.Success)
            references.Add(new Reference(pair.Groups["book"].Value, pair.Groups["page"].Value, null));
        else if (bookId is not null)
            references.Add(new Reference(bookId, normalized, null));
        else
            references.Add(new Reference(null, null, normalized));
    }

    private static string? ReadScalar(JsonElement element, IReadOnlyCollection<string>? keys = null)
    {
        if (keys is null)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => NormalizeId(element.GetString()),
                JsonValueKind.Number => NormalizeId(element.GetRawText()),
                _ => null
            };
        }

        if (element.ValueKind != JsonValueKind.Object)
            return null;

        foreach (var property in element.EnumerateObject())
        {
            if (keys.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                return ReadScalar(property.Value);
        }

        return null;
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

    private static bool IsLinkKey(string name) =>
        name.Equals("url", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("uri", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("href", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("link", StringComparison.OrdinalIgnoreCase);

    private static bool IsIdentifierCollectionKey(string name) =>
        name.Equals("ids", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("page_ids", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("pageIds", StringComparison.OrdinalIgnoreCase);

    private static string Pair(string bookId, string pageId) => bookId + "/" + pageId;

    private static string? NormalizeId(string? value)
    {
        var normalized = value?.Trim();
        return string.IsNullOrWhiteSpace(normalized) ? null : normalized;
    }

    private sealed record Reference(string? BookId, string? PageId, string? OpaqueId);
}
