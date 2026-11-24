using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchSearchBootstrapTests
{
    [Fact]
    public void BuildsOneHostOwnedInitialSearchForEverySelectedLibrary()
    {
        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            Tools(),
            ["aljam3", "shamela"],
            "بناء القوي على الضعيف");

        Assert.Equal(2, plan.Count);
        Assert.Equal(
            new[] { "aljam3", "shamela" },
            plan.Select(item => item.Tool.Server).ToArray());
        Assert.Equal(
            new[] { "search", "shamela_find" },
            plan.Select(item => item.Tool.Name).ToArray());
    }

    [Fact]
    public void MapsAljam3QueryIntoTheActualRequiredStringArgument()
    {
        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            Tools(),
            ["aljam3"],
            "بناء القوي على الضعيف");

        var invocation = Assert.Single(plan);
        Assert.Equal("search", invocation.Tool.Name);
        Assert.Equal(
            "بناء القوي على الضعيف",
            Assert.IsType<string>(invocation.Arguments["query"]));
    }

    [Fact]
    public void MapsShamelaQueryIntoItsRequiredKeywordsArray()
    {
        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            Tools(),
            ["shamela"],
            "بناء القوي على الضعيف");

        var invocation = Assert.Single(plan);
        Assert.Equal("shamela_find", invocation.Tool.Name);
        var keywords = Assert.IsType<string[]>(invocation.Arguments["keywords"]);
        Assert.Equal(new[] { "بناء القوي على الضعيف" }, keywords);
    }

    [Fact]
    public void PrefersGeneralShamelaSearchBeforeScopedSearch()
    {
        var available = Tools().Where(tool => tool.Server == "shamela").ToArray();
        var tools = new[]
        {
            available.Single(tool => tool.Name == "shamela_find_scoped"),
            available.Single(tool => tool.Name == "shamela_find")
        };

        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            tools,
            ["shamela"],
            "الاجتهاد");

        Assert.Equal("shamela_find", Assert.Single(plan).Tool.Name);
    }

    [Fact]
    public void DoesNotScheduleAnyLibraryTheUserDidNotSelect()
    {
        var plan = ResearchSearchBootstrap.BuildInitialPlan(
            Tools(),
            ["shamela"],
            "الاجتهاد");

        Assert.All(plan, item => Assert.Equal("shamela", item.Tool.Server));
    }

    [Fact]
    public void FailsClosedWhenASelectedLibraryHasNoUsableSearchTool()
    {
        var tools = Tools().Where(tool => tool.Server == "aljam3").ToArray();

        var exception = Assert.Throws<InvalidOperationException>(() =>
            ResearchSearchBootstrap.BuildInitialPlan(
                tools,
                ["aljam3", "shamela"],
                "الاجتهاد"));

        Assert.Contains("shamela", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsSearchToolWhoseSchemaHasNoRecognizedQueryInput()
    {
        var tools = new[]
        {
            Tool(
                "aljam3",
                "search",
                """
                {
                  "type": "object",
                  "properties": {
                    "book_id": { "type": "integer" }
                  },
                  "required": ["book_id"],
                  "additionalProperties": false
                }
                """)
        };

        Assert.Throws<InvalidOperationException>(() =>
            ResearchSearchBootstrap.BuildInitialPlan(
                tools,
                ["aljam3"],
                "الاجتهاد"));
    }

    private static IReadOnlyList<ToolDescriptor> Tools() =>
    [
        Tool(
            "aljam3",
            "search",
            """
            {
              "type": "object",
              "properties": {
                "query": { "type": "string" },
                "per_page": { "type": "integer", "default": 10 }
              },
              "required": ["query"],
              "additionalProperties": false
            }
            """),
        Tool(
            "shamela",
            "shamela_find_scoped",
            """
            {
              "type": "object",
              "properties": {
                "keywords": { "type": "array", "items": { "type": "string" } },
                "book_queries": { "type": "array", "items": { "type": "string" }, "default": [] }
              },
              "required": ["keywords"],
              "additionalProperties": false
            }
            """),
        Tool(
            "shamela",
            "shamela_find",
            """
            {
              "type": "object",
              "properties": {
                "keywords": { "type": "array", "items": { "type": "string" } },
                "author_queries": { "type": "array", "items": { "type": "string" }, "default": [] },
                "book_queries": { "type": "array", "items": { "type": "string" }, "default": [] },
                "category_ids": { "type": "array", "items": { "type": "integer" }, "default": [] },
                "cursor": {
                  "anyOf": [{ "type": "string" }, { "type": "null" }],
                  "default": null
                },
                "difficulty": { "type": "string", "enum": ["easy", "medium", "hard"], "default": "easy" },
                "field": { "type": "string", "enum": ["body", "footnotes"], "default": "body" }
              },
              "required": ["keywords"],
              "additionalProperties": false
            }
            """)
    ];

    private static ToolDescriptor Tool(string server, string name, string schemaText)
    {
        using var schema = JsonDocument.Parse(schemaText);
        return new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            string.Empty,
            schema.RootElement.Clone());
    }
}
