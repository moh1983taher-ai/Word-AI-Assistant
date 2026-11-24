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
            ["موضع النص: https://shamela.ws/book/21592/1157"],
            isError: false,
            hasUsableContent: true);

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":1157}"""));

        Assert.Null(error);
    }

    [Fact]
    public void RejectsOpeningWhenSearchOnlyReturnedPrintedPageNumber()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"printed_page":"369"}]}""");

        var error = provenance.ValidateReferenceRead(
            Tool("shamela", "shamela_open"),
            Arguments("""{"book_id":21592,"page_id":369}"""));

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
    public void TracksUnverifiedSearchHitsPerSourceUntilAValidPageIsOpened()
    {
        var provenance = SearchShamela(
            """{"results":[{"book_id":21592,"page_id":1157}]}""");
        var open = Tool("shamela", "shamela_open");

        Assert.True(provenance.HasUnverifiedSearchHits);
        provenance.RecordVerifiedOpen(open, succeeded: true);
        Assert.False(provenance.HasUnverifiedSearchHits);
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

        Assert.Equal(0, provenance.SearchHitCount);
        Assert.False(provenance.HasUnverifiedSearchHits);
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
