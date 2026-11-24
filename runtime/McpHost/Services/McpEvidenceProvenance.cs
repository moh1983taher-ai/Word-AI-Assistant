using System.Text.Json;
using System.Text.RegularExpressions;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Tracks page references returned by each MCP source and rejects page reads that
/// do not use a reference from a prior search in that same source.
/// </summary>
internal sealed class McpEvidenceProvenance
{
    private static readonly string[] BookKeys = ["book_id", "bookId", "bookid"];
    private static readonly string[] FileKeys = ["file_id", "fileId"];
    private static readonly string[] FilePageNumberKeys = ["page_number", "pageNumber"];
    // Search results must expose an internal page identifier, not a printed page number.
    private static readonly string[] SearchPageKeys = ["page_id", "pageId", "pageid", "start_page_id", "startPageId"];
    // The open tool may name its requested internal page identifier simply "page".
    private static readonly string[] OpenPageKeys = ["page_id", "pageId", "pageid", "page", "start_page_id", "startPageId"];

    private static readonly Regex BookPagePath = new(
        """(?i)/(?:books?)/(?<book>\d+)/(?:pages?/)?(?<page>\d+)(?!\d)""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex BookPagePair = new(
        """(?i)\bbook[_\s-]?id["'\s:=]+(?<book>\d+)[\s\S]{0,240}?\bpage[_\s-]?id["'\s:=]+(?<page>\d+)""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ColonPair = new(
        """^\s*(?<book>\d+)\s*:\s*(?<page>\d+)\s*$""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly Dictionary<string, SourceIndex> _sources = new(StringComparer.Ordinal);

    internal int NonEmptySearchResponseCount => _sources.Values.Sum(index => index.NonEmptySearchResponseCount);
    internal int VerifiedPageOpenCount => _sources.Values.Sum(index => index.VerifiedPageOpenCount);
    internal bool HasUnverifiedSearchResults =>
        _sources.Values.Any(index => index.NonEmptySearchResponseCount > 0 && index.VerifiedPageOpenCount == 0);

    internal int SearchReferenceCount =>
        _sources.Values.Sum(index => index.ReferenceCount);

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
        bool isError,
        bool hasUsableContent)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        if (!ResearchToolResultClassifier.IsSearchTool(descriptor))
            return 0;

        if (isError)
            return 0;

        var index = GetSourceIndex(descriptor.Server);
        if (hasUsableContent)
            index.NonEmptySearchResponseCount++;
        if (structuredContent is { } structured &&
            structured.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null)
        {
            VisitSearchPayload(structured, index, inheritedBookId: null, descriptor.Server);
        }

        if (textBlocks is not null)
        {
            foreach (var text in textBlocks)
                VisitSearchText(text, index, descriptor.Server);
        }

        return index.ReferenceCount;
    }

    internal string? ValidateReferenceRead(
        ToolDescriptor descriptor,
        IReadOnlyDictionary<string, object?> arguments)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        ArgumentNullException.ThrowIfNull(arguments);

        if (!IsReferenceReadTool(descriptor))
            return null;

        if (!_sources.TryGetValue(descriptor.Server, out var index) ||
            index.ReferenceCount == 0)
        {
            return "لا توجد نتيجة بحث سابقة من هذه المكتبة تحمل معرّف صفحة يمكن التحقق منه؛ أعد البحث قبل فتح الصفحة.";
        }

        var input = JsonSerializer.SerializeToElement(arguments);
        var requested = new List<Reference>();
        VisitOpenArguments(input, requested, inheritedBookId: null, descriptor.Server);

        if (requested.Count == 0)
        {
            return "معاملات فتح الصفحة لا تتضمن معرّف كتاب أو صفحة معروفاً؛ لا تخمّن المعرّفات، واستخدم معرّفاً وارداً في نتيجة البحث.";
        }

        if (requested.Any(reference => !WasReturnedBySearch(index, reference)))
        {
            return "المعرّف المطلوب فتحه غير موجود في نتائج البحث السابقة لهذه المكتبة؛ اختر نتيجة من المخرجات الفعلية أو أعد البحث. لا تفتح موضعاً مستنتجاً أو منقولاً من مصدر آخر.";
        }

        return null;
    }

    /// <summary>
    /// If an opened page returns its own book/page identifiers, make sure they
    /// identify the same hit that the host authorized for opening. Plain-text
    /// servers that omit response identifiers remain traceable to the authorized input.
    /// </summary>
    internal string? ValidateOpenedResult(
        ToolDescriptor descriptor,
        IReadOnlyDictionary<string, object?> arguments,
        JsonElement? structuredContent,
        IEnumerable<string>? textBlocks)
    {
        if (!IsReferenceReadTool(descriptor))
            return null;

        var requested = new List<Reference>();
        VisitOpenArguments(JsonSerializer.SerializeToElement(arguments), requested, inheritedBookId: null);

        if (requested.Count == 0)
            return "تعذر تحديد الصفحة المطلوبة للتحقق من نتيجة الفتح.";

        var returned = new SourceIndex();
        if (structuredContent is { } structured &&
            structured.ValueKind is not JsonValueKind.Undefined and not JsonValueKind.Null)
        {
            VisitSearchPayload(structured, returned, inheritedBookId: null, descriptor.Server);
        }

        if (textBlocks is not null)
        {
            foreach (var text in textBlocks)
                VisitSearchText(text, returned, descriptor.Server);
        }

        if (returned.ReferenceCount == 0)
            return null;

        if (requested.Any(request => !returned.Contains(request)))
        {
            return "أعادت أداة الفتح بيانات كتاب أو صفحة لا تطابق المعرّف الذي طُلب فتحه؛ لم يعتمد المضيف هذا المحتوى دليلاً.";
        }

        return null;
    }

    internal void RecordVerifiedOpen(ToolDescriptor descriptor, bool succeeded)
    {
        if (succeeded && IsReferenceReadTool(descriptor))
            GetSourceIndex(descriptor.Server).VerifiedPageOpenCount++;
    }

    private SourceIndex GetSourceIndex(string sourceId)
    {
        if (!_sources.TryGetValue(sourceId, out var index))
        {
            index = new SourceIndex();
            _sources.Add(sourceId, index);
        }

        return index;
    }

    private static bool WasReturnedBySearch(SourceIndex index, Reference reference)
    {
        if (reference.FileId is not null || reference.FilePageNumber is not null)
            return reference.FileId is not null && reference.FilePageNumber is not null &&
                   index.FilePagePairs.Contains(Pair(reference.FileId, reference.FilePageNumber));

        if (reference.BookId is not null && reference.PageId is not null)
            return index.BookPagePairs.Contains(Pair(reference.BookId, reference.PageId));

        if (reference.BookId is not null)
            return false;

        if (reference.PageId is not null)
            return index.PageIds.Contains(reference.PageId) ||
                   index.OpaqueIds.Contains(reference.PageId) ||
                   index.HasUnambiguousPageId(reference.PageId);

        return reference.OpaqueId is not null &&
               (index.OpaqueIds.Contains(reference.OpaqueId) ||
                index.PageIds.Contains(reference.OpaqueId));
    }

    private static void VisitSearchPayload(
        JsonElement element,
        SourceIndex index,
        string? inheritedBookId,
        string sourceId)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var bookId = ReadScalar(element, BookKeys) ?? inheritedBookId;
                var pageId = ReadScalar(element, SearchPageKeys);
                if (sourceId == "aljam3")
                {
                    var fileId = ReadScalar(element, FileKeys);
                    var pageNumber = ReadScalar(element, FilePageNumberKeys);
                    if (fileId is not null && pageNumber is not null)
                        index.AddFilePagePair(fileId, pageNumber);
                }

                if (bookId is not null && pageId is not null)
                    index.AddBookPagePair(bookId, pageId);
                else if (pageId is not null)
                    index.PageIds.Add(pageId);

                // Aljam3 search-hit `id` identifies the page record, not the
                // file/page pair required by get_page. Never authorize it as an opaque page ref.
                if (sourceId != "aljam3" && bookId is null &&
                    TryGetProperty(element, "id", out var idElement))
                    AddIdentifier(index, ReadScalar(idElement), bookId);

                foreach (var property in element.EnumerateObject())
                {
                    if (IsLinkKey(property.Name) && property.Value.ValueKind == JsonValueKind.String)
                        VisitSearchText(property.Value.GetString(), index, sourceId);

                    if (IsIdentifierCollectionKey(property.Name) &&
                        property.Value.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in property.Value.EnumerateArray())
                        {
                            if (item.ValueKind is JsonValueKind.String or JsonValueKind.Number)
                                AddIdentifier(index, ReadScalar(item), bookId);
                        }
                    }

                    if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        VisitSearchPayload(property.Value, index, bookId, sourceId);
                }
                break;
            }

            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    VisitSearchPayload(item, index, inheritedBookId);
                break;

            case JsonValueKind.String:
                VisitSearchText(element.GetString(), index);
                break;
        }
    }

    private static void VisitOpenArguments(
        JsonElement element,
        List<Reference> references,
        string? inheritedBookId,
        string sourceId)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var bookId = sourceId == "aljam3"
                    ? inheritedBookId
                    : ReadScalar(element, BookKeys) ?? inheritedBookId;

                if (sourceId == "aljam3")
                {
                    var fileId = ReadScalar(element, FileKeys);
                    var pageNumber = ReadScalar(element, FilePageNumberKeys);
                    if (fileId is not null || pageNumber is not null)
                        references.Add(new Reference(null, null, null, fileId, pageNumber));
                }
                else
                {
                    var pageId = ReadScalar(element, OpenPageKeys);
                    if (bookId is not null || pageId is not null)
                        references.Add(new Reference(bookId, pageId, null));

                    if (pageId is null && TryGetProperty(element, "id", out var idElement))
                        AddTargetIdentifier(ReadScalar(idElement), bookId, references);
                }

                foreach (var property in element.EnumerateObject())
                {
                    if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        VisitOpenArguments(property.Value, references, bookId, sourceId);
                    else if (sourceId != "aljam3" &&
                             property.Value.ValueKind == JsonValueKind.String &&
                             IsIdentifierCollectionKey(property.Name))
                        AddTargetIdentifier(property.Value.GetString(), bookId, references);
                }
                break;
            }

            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    VisitOpenArguments(item, references, inheritedBookId, sourceId);
                break;

            case JsonValueKind.String:
                AddTargetIdentifier(element.GetString(), inheritedBookId, references);
                break;

            case JsonValueKind.Number:
                AddTargetIdentifier(element.GetRawText(), inheritedBookId, references);
                break;
        }
    }

    private static void VisitSearchText(string? text, SourceIndex index, string sourceId)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        foreach (Match match in BookPagePath.Matches(text))
            index.AddBookPagePair(match.Groups["book"].Value, match.Groups["page"].Value);

        foreach (Match match in BookPagePair.Matches(text))
            index.AddBookPagePair(match.Groups["book"].Value, match.Groups["page"].Value);

        var candidate = text.Trim();
        var colonPair = ColonPair.Match(candidate);
        if (colonPair.Success)
            index.AddBookPagePair(colonPair.Groups["book"].Value, colonPair.Groups["page"].Value);

        var firstBrace = candidate.IndexOf('{');
        var lastBrace = candidate.LastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace)
            candidate = candidate[firstBrace..(lastBrace + 1)];

        try
        {
            using var document = JsonDocument.Parse(candidate);
            VisitSearchPayload(document.RootElement, index, inheritedBookId: null, sourceId);
        }
        catch (JsonException)
        {
            // Search servers may return Markdown; URL and key-value patterns above still apply.
        }
    }

    private static void AddIdentifier(SourceIndex index, string? id, string? bookId)
    {
        if (string.IsNullOrWhiteSpace(id))
            return;

        var normalized = id.Trim();
        var pair = ColonPair.Match(normalized);
        if (pair.Success)
        {
            index.AddBookPagePair(pair.Groups["book"].Value, pair.Groups["page"].Value);
            return;
        }

        if (bookId is not null)
            index.AddBookPagePair(bookId, normalized);
        else
            index.OpaqueIds.Add(normalized);
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
            references.Add(new Reference(null, normalized, null));
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
        name.Equals("pages", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("page_ids", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("pageIds", StringComparison.OrdinalIgnoreCase);

    private static string Pair(string bookId, string pageId) => bookId + "/" + pageId;

    private static string? NormalizeId(string? value)
    {
        var normalized = value?.Trim();
        return string.IsNullOrWhiteSpace(normalized) ? null : normalized;
    }

    private sealed class SourceIndex
    {
        internal int NonEmptySearchResponseCount { get; set; }
        internal int VerifiedPageOpenCount { get; set; }
        internal HashSet<string> PageIds { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> OpaqueIds { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> BookPagePairs { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> FilePagePairs { get; } = new(StringComparer.Ordinal);
        private Dictionary<string, HashSet<string>> PageBooks { get; } = new(StringComparer.Ordinal);
        internal int ReferenceCount => PageIds.Count + OpaqueIds.Count + BookPagePairs.Count + FilePagePairs.Count;

        internal void AddBookPagePair(string bookId, string pageId)
        {
            BookPagePairs.Add(Pair(bookId, pageId));
            if (!PageBooks.TryGetValue(pageId, out var books))
            {
                books = new HashSet<string>(StringComparer.Ordinal);
                PageBooks.Add(pageId, books);
            }

            books.Add(bookId);
        }

        internal void AddFilePagePair(string fileId, string pageNumber) =>
            FilePagePairs.Add(Pair(fileId, pageNumber));

        internal bool HasUnambiguousPageId(string pageId) =>
            PageBooks.TryGetValue(pageId, out var books) && books.Count == 1;

        internal bool Contains(Reference reference)
        {
            if (reference.FileId is not null || reference.FilePageNumber is not null)
                return reference.FileId is not null && reference.FilePageNumber is not null &&
                       FilePagePairs.Contains(Pair(reference.FileId, reference.FilePageNumber));

            if (reference.BookId is not null && reference.PageId is not null)
                return BookPagePairs.Contains(Pair(reference.BookId, reference.PageId));

            if (reference.BookId is not null)
                return false;

            if (reference.PageId is not null)
                return PageIds.Contains(reference.PageId) ||
                       OpaqueIds.Contains(reference.PageId) ||
                       HasUnambiguousPageId(reference.PageId);

            return reference.OpaqueId is not null &&
                   (OpaqueIds.Contains(reference.OpaqueId) || PageIds.Contains(reference.OpaqueId));
        }
    }

    private sealed record Reference(
        string? BookId,
        string? PageId,
        string? OpaqueId,
        string? FileId = null,
        string? FilePageNumber = null);
}
