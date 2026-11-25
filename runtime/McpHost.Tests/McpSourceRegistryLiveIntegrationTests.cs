using System.Text.Json;
using ModelContextProtocol.Protocol;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

/// <summary>
/// Live, read-only integration coverage for the official Aljam3 and Shamela MCP servers.
/// These tests are kept separate from the deterministic unit-test suite because they
/// require outbound network access and the upstream services to be available.
/// </summary>
public sealed class McpSourceRegistryLiveIntegrationTests
{
    [Fact]
    [Trait("Category", "Integration")]
    public async Task SearchesAndOpensVerifiedPagesOnBothLiveMcpSources()
    {
        await using var registry = new McpSourceRegistry(new HostLog());
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(180));
        var selectedSources = new[] { "aljam3", "shamela" };

        var discovery = await registry.DiscoverSelectedAsync(
            selectedSources,
            timeout.Token);

        Assert.Empty(discovery.Warnings);
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "aljam3" && tool.Name == "search");
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "aljam3" && tool.Name == "get_page");
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "shamela" && tool.Name == "shamela_find");
        Assert.Contains(discovery.Tools, tool =>
            tool.Server == "shamela" && tool.Name == "shamela_open");

        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            discovery.Tools,
            selectedSources,
            "مقاصد الشريعة");

        Assert.Equal(2, plan.Count);
        Assert.Equal(
            new[] { "aljam3", "shamela" },
            plan.Select(invocation => invocation.Tool.Server).ToArray());

        var provenance = new McpEvidenceProvenance();

        foreach (var invocation in plan)
        {
            var searchResult = await registry.CallToolAsync(
                invocation.Tool.Server,
                invocation.Tool.Name,
                invocation.Arguments,
                timeout.Token);

            Assert.NotEqual(true, searchResult.IsError);

            var searchText = GetTextBlocks(searchResult);
            var searchCharacters = GetReturnedCharacterCount(searchResult, searchText);
            Assert.True(
                searchCharacters > 0,
                $"{invocation.Tool.Server}/{invocation.Tool.Name} returned no search content.");

            provenance.RecordSearchResult(
                invocation.Tool,
                searchResult.StructuredContent,
                searchText,
                isError: false,
                hasUsableContent: true);

            Console.WriteLine(
                $"Live MCP search succeeded: source={invocation.Tool.Server}; tool={invocation.Tool.Name}; returnedCharacters={searchCharacters}.");

            var openToolName = invocation.Tool.Server == "aljam3"
                ? "get_page"
                : "shamela_open";
            var openTool = Assert.Single(discovery.Tools.Where(tool =>
                tool.Server == invocation.Tool.Server && tool.Name == openToolName));

            var reference = FindReference(searchResult, invocation.Tool.Server);
            Assert.NotNull(reference);

            var openArguments = BuildOpenArguments(openTool, reference!);
            var referenceError = provenance.ValidateReferenceRead(openTool, openArguments);
            Assert.Null(referenceError);

            var openResult = await registry.CallToolAsync(
                openTool.Server,
                openTool.Name,
                openArguments,
                timeout.Token);

            Assert.NotEqual(true, openResult.IsError);

            var openText = GetTextBlocks(openResult);
            var openCharacters = GetReturnedCharacterCount(openResult, openText);
            Assert.True(
                openCharacters > 0,
                $"{openTool.Server}/{openTool.Name} returned no page content.");

            var identityError = provenance.ValidateOpenedResult(
                openTool,
                openArguments,
                openResult.StructuredContent,
                openText);
            Assert.Null(identityError);

            provenance.RecordVerifiedOpen(openTool, succeeded: true);

            Console.WriteLine(
                $"Live MCP page open succeeded: source={openTool.Server}; tool={openTool.Name}; returnedCharacters={openCharacters}; referenceVerified=true.");
        }

        Assert.False(provenance.HasUnverifiedSearchResults);
        Assert.Equal(2, provenance.VerifiedPageOpenCount);
    }

    private static string[] GetTextBlocks(CallToolResult result) =>
        result.Content
            .OfType<TextContentBlock>()
            .Select(block => block.Text ?? string.Empty)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray();

    private static int GetReturnedCharacterCount(
        CallToolResult result,
        IReadOnlyCollection<string> textBlocks)
    {
        var textCharacters = textBlocks.Sum(text => text.Length);
        var structuredCharacters = result.StructuredContent is { } structured
            ? structured.GetRawText().Length
            : 0;
        return textCharacters + structuredCharacters;
    }

    private static IReadOnlyDictionary<string, JsonElement>? FindReference(
        CallToolResult result,
        string source)
    {
        if (result.StructuredContent is { } structured)
        {
            var match = FindReferenceInElement(structured, source);
            if (match is not null)
                return match;
        }

        foreach (var text in GetTextBlocks(result))
        {
            var firstBrace = text.IndexOf('{');
            var lastBrace = text.LastIndexOf('}');
            if (firstBrace < 0 || lastBrace <= firstBrace)
                continue;

            try
            {
                using var document = JsonDocument.Parse(text[firstBrace..(lastBrace + 1)]);
                var match = FindReferenceInElement(document.RootElement, source);
                if (match is not null)
                    return match;
            }
            catch (JsonException)
            {
                // Some MCP servers return explanatory Markdown beside their JSON result.
            }
        }

        return null;
    }

    private static IReadOnlyDictionary<string, JsonElement>? FindReferenceInElement(
        JsonElement element,
        string source)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            if (source == "aljam3")
            {
                var fileId = FindScalar(element, "file_id", "fileId");
                var pageNumber = FindScalar(element, "page_number", "pageNumber");
                if (fileId is { } file && pageNumber is { } page)
                {
                    return new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["file_id"] = file,
                        ["page_number"] = page
                    };
                }
            }
            else
            {
                var bookId = FindScalar(element, "book_id", "bookId", "bookid");
                var pageId = FindScalar(
                    element,
                    "page_id", "pageId", "pageid", "start_page_id", "startPageId");
                if (bookId is { } book && pageId is { } page)
                {
                    return new Dictionary<string, JsonElement>(StringComparer.Ordinal)
                    {
                        ["book_id"] = book,
                        ["page_id"] = page
                    };
                }
            }

            foreach (var property in element.EnumerateObject())
            {
                var match = FindReferenceInElement(property.Value, source);
                if (match is not null)
                    return match;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var match = FindReferenceInElement(item, source);
                if (match is not null)
                    return match;
            }
        }

        return null;
    }

    private static JsonElement? FindScalar(JsonElement element, params string[] names)
    {
        if (element.ValueKind != JsonValueKind.Object)
            return null;

        foreach (var property in element.EnumerateObject())
        {
            if (names.Contains(property.Name, StringComparer.OrdinalIgnoreCase) &&
                property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number)
                return property.Value.Clone();
        }

        return null;
    }

    private static IReadOnlyDictionary<string, object?> BuildOpenArguments(
        ToolDescriptor openTool,
        IReadOnlyDictionary<string, JsonElement> reference)
    {
        var properties = openTool.InputSchema.GetProperty("properties");
        var arguments = new Dictionary<string, object?>(StringComparer.Ordinal);

        foreach (var (canonicalName, value) in reference)
        {
            var aliases = canonicalName switch
            {
                "file_id" => new[] { "file_id", "fileId" },
                "page_number" => new[] { "page_number", "pageNumber" },
                "book_id" => new[] { "book_id", "bookId", "bookid" },
                "page_id" => new[] { "page_id", "pageId", "pageid", "page", "start_page_id", "startPageId" },
                _ => [canonicalName]
            };

            var actualName = properties.EnumerateObject()
                .Select(property => property.Name)
                .FirstOrDefault(name => aliases.Contains(name, StringComparer.OrdinalIgnoreCase));

            Assert.False(
                string.IsNullOrWhiteSpace(actualName),
                $"{openTool.Server}/{openTool.Name} schema omitted the required reference field {canonicalName}.");

            arguments[actualName!] = value.Clone();
        }

        return arguments;
    }
}
