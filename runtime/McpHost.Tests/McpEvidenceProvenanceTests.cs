using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class McpEvidenceProvenanceTests
{
    [Fact]
    public void AllowsOpeningAnExactBookPagePairReturnedBySearch()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157,"title":"المنثور في القواعد"}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":1157}"""));

        Assert.Null(error);
    }

    [Fact]
    public void RejectsAReferenceThatWasNotReturnedBySearch()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":9999}"""));

        Assert.Contains("غير موجود في نتائج البحث السابقة", error);
    }

    [Fact]
    public void DoesNotCombineBookAndPageIdsFromDifferentSearchHits()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":10,"page_id":100},{"book_id":20,"page_id":200}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":10,"page_id":200}"""));

        Assert.Contains("غير موجود في نتائج البحث السابقة", error);
    }

    [Fact]
    public void KeepsIdentifiersIsolatedBetweenSelectedSources()
    {
        var provenance = new McpEvidenceProvenance();
        RegisterSearch(
            provenance,
            Tool("aljam3", "search"),
            """{"results":[{"book_id":10,"page_id":100}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":10,"page_id":100}"""));

        Assert.Contains("لا توجد نتيجة بحث سابقة من هذه المكتبة", error);
    }

    [Fact]
    public void ExtractsBookAndPageIdentifiersFromSourceLinks()
    {
        var provenance = new McpEvidenceProvenance();
        var search = Tool("shamela", "shamela_find");
        using var empty = JsonDocument.Parse("{}");

        provenance.RecordSearchResult(
            search,
            empty.RootElement,
            ["موضع النص: [فتح الصفحة](https://shamela.ws/book/21592/1157)"],
            isError: false,
            hasUsableContent: true);

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":1157}"""));

        Assert.Null(error);
    }

    [Fact]
    public void ExtractsNumericPageIdCollectionsReturnedBySearch()
    {
        var provenance = new McpEvidenceProvenance();
        var search = Tool("shamela", "shamela_find");
        using var result = JsonDocument.Parse(
            """{"book_id":21592,"pages":[1157,1158]}""");

        provenance.RecordSearchResult(
            search,
            result.RootElement,
            Array.Empty<string>(),
            isError: false,
            hasUsableContent: true);

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open_many"),
            Arguments("""{"pages":[1157,1158]}"""));

        Assert.Null(error);
    }

    [Fact]
    public void AcceptsOpenedContentWhenItsReturnedBookAndPageMatch()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");
        using var opened = JsonDocument.Parse(
            """{"book_id":21592,"page_id":1157,"text":"نص الصفحة"}""");

        var error = provenance.ValidateOpenedResult(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":1157}"""),
            opened.RootElement,
            Array.Empty<string>());

        Assert.Null(error);
    }

    [Fact]
    public void DoesNotTreatGenericSearchPageNumberAsAnInternalPageId()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page":369,"printed_page":"369"}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":369}"""));

        Assert.Contains("لا توجد نتيجة بحث سابقة", error);
    }

    [Fact]
    public void AllowsPageIdOnlyWhenItUniquelyIdentifiesASearchedBookPage()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"page_id":1157}"""));

        Assert.Null(error);
    }

    [Fact]
    public void RejectsPageIdOnlyWhenSearchHitsUseItInMoreThanOneBook()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":10,"page_id":1157},{"book_id":20,"page_id":1157}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"page_id":1157}"""));

        Assert.Contains("غير موجود في نتائج البحث السابقة", error);
    }

    [Fact]
    public void DoesNotAllowOpeningABookWithoutAPageReference()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592}"""));

        Assert.Contains("غير موجود في نتائج البحث السابقة", error);
    }

    [Fact]
    public void SupportsOpenToolPageAliasOnlyWhenSearchReturnedAnInternalPageId()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page":1157}"""));

        Assert.Null(error);
    }

    [Fact]
    public void DoesNotInferPageIdFromGenericBookRecordId()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"id":1157,"title":"المنثور في القواعد"}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":1157}"""));

        Assert.Contains("لا توجد نتيجة بحث سابقة", error);
    }

    [Fact]
    public void RequiresEveryPageInMultiOpenRequestToComeFromSearchResults()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":10,"page_id":100},{"book_id":10,"page_id":101}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open_many"),
            Arguments("""{"pages":[{"book_id":10,"page_id":100},{"book_id":10,"page_id":999}]}"""));

        Assert.Contains("غير موجود في نتائج البحث السابقة", error);
    }

    [Fact]
    public void RejectsOpenedContentWhoseReturnedPageDiffersFromRequestedHit()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");
        var open = Tool("shamela", "shamela_open");
        var arguments = Arguments("""{"book_id":21592,"page_id":1157}""");

        using var wrongPage = JsonDocument.Parse(
            """{"book_id":21592,"page_id":1158,"text":"نص من صفحة أخرى"}""");

        var error = provenance.ValidateOpenedResult(
            open,
            arguments,
            wrongPage.RootElement,
            Array.Empty<string>());

        Assert.Contains("لا تطابق", error);
    }

    [Fact]
    public void TracksUnverifiedSearchResultsPerSourceUntilAValidPageIsOpened()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");
        var open = Tool("shamela", "shamela_open");

        Assert.True(provenance.HasUnverifiedSearchResults);
        provenance.RecordVerifiedOpen(open, succeeded: true);
        Assert.False(provenance.HasUnverifiedSearchResults);
    }

    [Fact]
    public void ReportsOnlySourcesWithSearchHitsThatStillNeedPageVerification()
    {
        var provenance = SearchAljam3(
            """{"results":[{"file_id":8291,"page_number":471}]}""");
        RegisterSearch(
            provenance,
            Tool("shamela", "shamela_find"),
            """{"results":[{"book_id":21592,"page_id":1157}]}""");

        Assert.Equal(
            new[] { "aljam3", "shamela" },
            provenance.GetUnverifiedSearchSources());

        provenance.RecordVerifiedOpen(
            Tool("aljam3", "get_page"),
            succeeded: true);

        Assert.Equal(
            new[] { "shamela" },
            provenance.GetUnverifiedSearchSources());

        provenance.RecordVerifiedOpen(
            Tool("shamela", "shamela_open"),
            succeeded: true);

        Assert.Empty(provenance.GetUnverifiedSearchSources());
    }

    [Fact]
    public void DoesNotRequirePageVerificationForSourcesWithNoSearchHits()
    {
        var provenance = new McpEvidenceProvenance();
        var search = Tool("shamela", "shamela_find");
        using var empty = JsonDocument.Parse("""{"results":[]}""");

        provenance.RecordSearchResult(
            search,
            empty.RootElement,
            Array.Empty<string>(),
            isError: false,
            hasUsableContent: false);

        Assert.Empty(provenance.GetUnverifiedSearchSources());
    }

    [Fact]
    public void DoesNotRequireOpeningWhenSearchReturnsNoHits()
    {
        var provenance = new McpEvidenceProvenance();
        var search = Tool("shamela", "shamela_find");
        using var empty = JsonDocument.Parse("""{"results":[]}""");

        provenance.RecordSearchResult(
            search,
            empty.RootElement,
            Array.Empty<string>(),
            isError: false,
            hasUsableContent: false);

        Assert.Equal(0, provenance.NonEmptySearchResponseCount);
        Assert.False(provenance.HasUnverifiedSearchResults);
    }

    [Fact]
    public void BuildsSchemaValidatedFallbackOpenArgumentsFromExactShamelaHit()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157},{"book_id":21592,"page_id":1158}]}""");
        var open = ToolWithSchema(
            "shamela",
            "shamela_open",
            """{"type":"object","properties":{"book_id":{"type":"integer"},"page_id":{"type":"integer"}},"required":["book_id","page_id"],"additionalProperties":false}""");

        Assert.True(provenance.TryBuildVerifiedReadArguments(open, out var arguments));
        Assert.Equal(21592L, Assert.IsType<JsonElement>(arguments["book_id"]).GetInt64());
        Assert.Equal(1157L, Assert.IsType<JsonElement>(arguments["page_id"]).GetInt64());
        Assert.Null(provenance.ValidateReferenceRead(open, arguments));
    }

    [Fact]
    public void BuildsSchemaValidatedFallbackOpenArgumentsFromExactAljam3Hit()
    {
        var provenance = SearchAljam3(
            """{"results":[{"file_id":8291,"page_number":471},{"file_id":8291,"page_number":472}]}""");
        var open = ToolWithSchema(
            "aljam3",
            "get_page",
            """{"type":"object","properties":{"file_id":{"type":"integer"},"page_number":{"type":"integer"}},"required":["file_id","page_number"],"additionalProperties":false}""");

        Assert.True(provenance.TryBuildVerifiedReadArguments(open, out var arguments));
        Assert.Equal(8291L, Assert.IsType<JsonElement>(arguments["file_id"]).GetInt64());
        Assert.Equal(471L, Assert.IsType<JsonElement>(arguments["page_number"]).GetInt64());
        Assert.Null(provenance.ValidateReferenceRead(open, arguments));
    }

    [Fact]
    public void BuildsSchemaValidatedFallbackOpenArgumentsFromSecondShamelaHitInSearchOrder()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157},{"book_id":21592,"page_id":1158}]}""");
        var open = ToolWithSchema(
            "shamela",
            "shamela_open",
            """{"type":"object","properties":{"book_id":{"type":"integer"},"page_id":{"type":"integer"}},"required":["book_id","page_id"],"additionalProperties":false}""");

        Assert.True(provenance.TryBuildVerifiedReadArguments(open, 1, out var arguments));
        Assert.Equal(21592L, Assert.IsType<JsonElement>(arguments["book_id"]).GetInt64());
        Assert.Equal(1158L, Assert.IsType<JsonElement>(arguments["page_id"]).GetInt64());
        Assert.Null(provenance.ValidateReferenceRead(open, arguments));
        Assert.False(provenance.TryBuildVerifiedReadArguments(open, 2, out _));
    }

    [Fact]
    public void BuildsSchemaValidatedFallbackOpenArgumentsFromSecondAljam3HitInSearchOrder()
    {
        var provenance = SearchAljam3(
            """{"results":[{"file_id":8291,"page_number":471},{"file_id":8291,"page_number":472}]}""");
        var open = ToolWithSchema(
            "aljam3",
            "get_page",
            """{"type":"object","properties":{"file_id":{"type":"integer"},"page_number":{"type":"integer"}},"required":["file_id","page_number"],"additionalProperties":false}""");

        Assert.True(provenance.TryBuildVerifiedReadArguments(open, 1, out var arguments));
        Assert.Equal(8291L, Assert.IsType<JsonElement>(arguments["file_id"]).GetInt64());
        Assert.Equal(472L, Assert.IsType<JsonElement>(arguments["page_number"]).GetInt64());
        Assert.Null(provenance.ValidateReferenceRead(open, arguments));
        Assert.False(provenance.TryBuildVerifiedReadArguments(open, 2, out _));
    }

    [Fact]
    public void DoesNotInventFallbackPageWhenNoSearchPagePairExists()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"title":"كتاب بلا معرّف صفحة"}]}""");
        var open = ToolWithSchema(
            "shamela",
            "shamela_open",
            """{"type":"object","properties":{"book_id":{"type":"integer"},"page_id":{"type":"integer"}},"required":["book_id","page_id"],"additionalProperties":false}""");

        Assert.False(provenance.TryBuildVerifiedReadArguments(open, out _));
    }

    private static ToolDescriptor ToolWithSchema(string server, string name, string schemaText)
    {
        using var schema = JsonDocument.Parse(schemaText);
        return new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            string.Empty,
            schema.RootElement.Clone());
    }

    [Fact]
    public void AllowsOpeningAnExactAljam3FilePagePairReturnedBySearch()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":8291,"page_number":471,"book":{"id":3435}}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":8291,"page_number":471}"""));

        Assert.Null(error);
    }

    [Fact]
    public void RejectsAnAljam3FilePagePairNotReturnedBySearch()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":8291,"page_number":471}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":8291,"page_number":472}"""));

        Assert.NotNull(error);
    }

    [Fact]
    public void DoesNotCombineAljam3FileAndPageNumbersFromDifferentHits()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":10,"page_number":100},{"id":502,"file_id":20,"page_number":200}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":10,"page_number":200}"""));

        Assert.NotNull(error);
    }

    [Fact]
    public void AcceptsAljam3OpenedContentWhenItsFileAndPageMatchTheRequestedHit()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":8291,"page_number":471}]}""");
        using var opened = JsonDocument.Parse(
            """{"page":{"id":501,"file_id":8291,"page_number":471},"text":{"content":"نص الصفحة"}}""");

        var error = provenance.ValidateOpenedResult(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":8291,"page_number":471}"""),
            opened.RootElement,
            Array.Empty<string>());

        Assert.Null(error);
    }

    [Fact]
    public void RejectsAljam3OpenedContentWhoseFileOrPageDiffersFromTheRequestedHit()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":8291,"page_number":471}]}""");
        using var opened = JsonDocument.Parse(
            """{"page":{"id":502,"file_id":8291,"page_number":472},"text":{"content":"نص صفحة أخرى"}}""");

        var error = provenance.ValidateOpenedResult(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":8291,"page_number":471}"""),
            opened.RootElement,
            Array.Empty<string>());

        Assert.NotNull(error);
    }

    [Fact]
    public void DoesNotTreatAljam3SearchResultRecordIdAsTheGetPageFileId()
    {
        var provenance = SearchAljam3(
            """{"results":[{"id":501,"file_id":8291,"page_number":471}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("aljam3", "get_page"),
            Arguments("""{"file_id":501,"page_number":471}"""));

        Assert.NotNull(error);
    }

    private static McpEvidenceProvenance SearchAljam3(string json)
    {
        var provenance = new McpEvidenceProvenance();
        RegisterSearch(provenance, Tool("aljam3", "search"), json);
        return provenance;
    }

    private static McpEvidenceProvenance SearchShamela(string json)
    {
        var provenance = new McpEvidenceProvenance();
        RegisterSearch(provenance, Tool("shamela", "shamela_find"), json);
        return provenance;
    }

    private static void RegisterSearch(
        McpEvidenceProvenance provenance,
        ToolDescriptor search,
        string json)
    {
        using var document = JsonDocument.Parse(json);
        provenance.RecordSearchResult(
            search,
            document.RootElement,
            Array.Empty<string>(),
            isError: false,
            hasUsableContent: true);
    }

    private static ToolDescriptor Tool(string server, string name)
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{}}""");
        return new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            string.Empty,
            schema.RootElement.Clone());
    }

    private static IReadOnlyDictionary<string, object?> Arguments(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.EnumerateObject().ToDictionary(
            property => property.Name,
            property => (object?)property.Value.Clone(),
            StringComparer.Ordinal);
    }
}
