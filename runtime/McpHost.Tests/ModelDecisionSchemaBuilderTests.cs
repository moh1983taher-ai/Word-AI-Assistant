using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ModelDecisionSchemaBuilderTests
{
    [Fact]
    public void PendingSourceSchemaAllowsOnlyToolCallsFromAvailableSource()
    {
        var tools = new[]
        {
            Tool("aljam3", "search", """
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
            Tool("aljam3", "get_page", """
                {
                  "type": "object",
                  "properties": {
                    "book_id": { "type": "integer" },
                    "page": { "type": "integer" }
                  },
                  "required": ["book_id", "page"],
                  "additionalProperties": false
                }
                """)
        };

        using var format = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(tools, allowFinal: false).GetRawText());

        var jsonSchema = format.RootElement.GetProperty("json_schema");
        Assert.True(jsonSchema.GetProperty("strict").GetBoolean());

        var schema = jsonSchema.GetProperty("schema");
        Assert.Equal("object", schema.GetProperty("type").GetString());
        Assert.False(schema.GetProperty("additionalProperties").GetBoolean());

        var properties = schema.GetProperty("properties");
        Assert.Equal(
            new[] { "tool_call" },
            properties.GetProperty("type").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Equal(
            new[] { "aljam3" },
            properties.GetProperty("server").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Equal(
            new[] { "search", "get_page" },
            properties.GetProperty("tool").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());

        var arguments = properties.GetProperty("arguments").GetProperty("anyOf");
        var searchArguments = arguments.EnumerateArray().Single(item =>
            item.GetProperty("properties").TryGetProperty("query", out _));
        var searchProperties = searchArguments.GetProperty("properties");

        Assert.Contains(
            "per_page",
            searchArguments.GetProperty("required").EnumerateArray()
                .Select(item => item.GetString()));
        Assert.Contains(
            "null",
            searchProperties.GetProperty("per_page").GetProperty("anyOf")
                .EnumerateArray().Select(item => item.GetProperty("type").GetString()));

        Assert.Equal(
            new[] { "type", "server", "tool", "arguments", "answer" },
            schema.GetProperty("required").EnumerateArray()
                .Select(item => item.GetString()).ToArray());
        Assert.Equal(
            new[] { "" },
            properties.GetProperty("answer").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
    }

    [Fact]
    public void FinalDecisionSchemaAllowsFinalAnswerAndNullToolFields()
    {
        var tools = new[]
        {
            Tool("shamela", "shamela_find", """
                {
                  "type": "object",
                  "properties": {
                    "keywords": {
                      "type": "array",
                      "items": { "type": "string" }
                    }
                  },
                  "required": ["keywords"],
                  "additionalProperties": false
                }
                """)
        };

        using var format = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(tools, allowFinal: true).GetRawText());

        var schema = format.RootElement.GetProperty("json_schema").GetProperty("schema");
        var properties = schema.GetProperty("properties");

        Assert.Equal(
            new[] { "tool_call", "final" },
            properties.GetProperty("type").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Contains(
            "",
            properties.GetProperty("server").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()));
        Assert.Contains(
            "null",
            properties.GetProperty("arguments").GetProperty("anyOf")
                .EnumerateArray().Select(item =>
                    item.TryGetProperty("type", out var type) ? type.GetString() : null));
        Assert.Equal("string", properties.GetProperty("answer").GetProperty("type").GetString());
    }

    [Fact]
    public void FinalOnlySchemaCanBeBuiltWhenToolBudgetIsExhausted()
    {
        using var format = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(
                Array.Empty<ToolDescriptor>(),
                allowFinal: true,
                allowToolCalls: false).GetRawText());

        var properties = format.RootElement
            .GetProperty("json_schema")
            .GetProperty("schema")
            .GetProperty("properties");

        Assert.Equal(
            new[] { "final" },
            properties.GetProperty("type").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Equal(
            new[] { "" },
            properties.GetProperty("server").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.Equal(
            "null",
            properties.GetProperty("arguments").GetProperty("type").GetString());
    }

    [Fact]
    public void NormalizesRealisticShamelaFindSchemaWithOptionalArraysEnumsAndNullableCursor()
    {
        var tools = new[]
        {
            Tool("shamela", "shamela_find", """
                {
                  "type": "object",
                  "properties": {
                    "author_queries": {
                      "type": "array",
                      "items": { "type": "string" },
                      "default": []
                    },
                    "book_queries": {
                      "type": "array",
                      "items": { "type": "string" },
                      "default": []
                    },
                    "category_ids": {
                      "type": "array",
                      "items": { "type": "integer" },
                      "default": []
                    },
                    "cursor": {
                      "anyOf": [
                        { "type": "string" },
                        { "type": "null" }
                      ],
                      "default": null
                    },
                    "difficulty": {
                      "type": "string",
                      "enum": ["easy", "medium", "hard"],
                      "default": "easy"
                    },
                    "field": {
                      "type": "string",
                      "enum": ["body", "footnotes"],
                      "default": "body"
                    },
                    "keywords": {
                      "type": "array",
                      "items": { "type": "string" }
                    }
                  },
                  "required": ["keywords"],
                  "additionalProperties": false
                }
                """)
        };

        using var format = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(tools, allowFinal: false).GetRawText());

        var arguments = format.RootElement
            .GetProperty("json_schema")
            .GetProperty("schema")
            .GetProperty("properties")
            .GetProperty("arguments");

        var properties = arguments.GetProperty("properties");
        var required = arguments.GetProperty("required")
            .EnumerateArray()
            .Select(item => item.GetString())
            .ToArray();

        Assert.Equal(7, required.Length);
        Assert.Contains("author_queries", required);
        Assert.Contains("keywords", required);

        Assert.Contains(
            "null",
            properties.GetProperty("author_queries").GetProperty("anyOf")
                .EnumerateArray().Select(item =>
                    item.TryGetProperty("type", out var type) ? type.GetString() : null));
        Assert.Contains(
            "null",
            properties.GetProperty("cursor").GetProperty("anyOf")
                .EnumerateArray().Select(item =>
                    item.TryGetProperty("type", out var type) ? type.GetString() : null));
        var difficultyType = properties.GetProperty("difficulty")
            .GetProperty("anyOf")
            .EnumerateArray()
            .Single(item =>
                item.TryGetProperty("type", out var type) &&
                type.GetString() == "string");

        Assert.Equal(
            new[] { "easy", "medium", "hard" },
            difficultyType.GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()).ToArray());
        Assert.False(arguments.GetProperty("additionalProperties").GetBoolean());
    }

    [Fact]
    public void InlinesShamelaLocalDefinitionsUsedByScopedSearchAndOpenMany()
    {
        var tools = new[]
        {
            Tool("shamela", "shamela_find_scoped", """
                {
                  "$defs": {
                    "Location": {
                      "type": "object",
                      "properties": {
                        "book_id": { "type": "integer" },
                        "page_id": { "type": "integer" }
                      },
                      "required": ["book_id", "page_id"],
                      "additionalProperties": false
                    },
                    "Open": {
                      "type": "object",
                      "properties": {
                        "book_id": { "type": "integer" },
                        "page_id": { "type": "integer" },
                        "cursor": {
                          "anyOf": [
                            { "type": "string" },
                            { "type": "null" }
                          ]
                        }
                      },
                      "required": ["book_id", "page_id"],
                      "additionalProperties": false
                    }
                  },
                  "type": "object",
                  "properties": {
                    "within_pages": {
                      "type": "array",
                      "items": { "$ref": "#/$defs/Location" }
                    },
                    "pages": {
                      "type": "array",
                      "items": { "$ref": "#/$defs/Open" }
                    }
                  },
                  "required": ["pages"],
                  "additionalProperties": false
                }
                """)
        };

        using var format = JsonDocument.Parse(
            ModelDecisionSchemaBuilder.BuildResponseFormat(tools, allowFinal: false).GetRawText());

        var raw = format.RootElement.GetRawText();
        Assert.False(raw.Contains("$ref", StringComparison.Ordinal));
        Assert.False(raw.Contains("$defs", StringComparison.Ordinal));

        var arguments = format.RootElement
            .GetProperty("json_schema")
            .GetProperty("schema")
            .GetProperty("properties")
            .GetProperty("arguments");

        var properties = arguments.GetProperty("properties");
        var pagesArray = properties.GetProperty("pages");
        var locationArray = properties.GetProperty("within_pages")
            .GetProperty("anyOf")
            .EnumerateArray()
            .Single(item => item.TryGetProperty("type", out var type) && type.GetString() == "array");

        var locationItem = locationArray.GetProperty("items");
        Assert.Equal("object", locationItem.GetProperty("type").GetString());
        Assert.Equal(
            new[] { "book_id", "page_id" },
            locationItem.GetProperty("required").EnumerateArray()
                .Select(item => item.GetString()).ToArray());

        var openItem = pagesArray.GetProperty("items");
        Assert.Equal("object", openItem.GetProperty("type").GetString());
        Assert.Equal(
            new[] { "book_id", "page_id", "cursor" },
            openItem.GetProperty("required").EnumerateArray()
                .Select(item => item.GetString()).ToArray());
        var cursorSchema = openItem.GetProperty("properties").GetProperty("cursor");
        Assert.Contains(
            "null",
            cursorSchema.GetProperty("anyOf").EnumerateArray()
                .Select(item => item.TryGetProperty("type", out var type) ? type.GetString() : null));
    }

    [Fact]
    public void RejectsCircularLocalSchemaReferencesWithAnExplicitError()
    {
        var tool = Tool("shamela", "recursive_test", """
            {
              "$defs": {
                "Node": {
                  "type": "object",
                  "properties": {
                    "next": { "$ref": "#/$defs/Node" }
                  },
                  "required": [],
                  "additionalProperties": false
                }
              },
              "type": "object",
              "properties": {
                "root": { "$ref": "#/$defs/Node" }
              },
              "required": ["root"],
              "additionalProperties": false
            }
            """);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            ModelDecisionSchemaBuilder.BuildResponseFormat([tool], allowFinal: false));

        Assert.Contains("Circular", exception.Message);
    }

    private static ToolDescriptor Tool(string server, string name, string schemaText)
    {
        using var document = JsonDocument.Parse(schemaText);
        return new ToolDescriptor(
            server,
            server == "aljam3" ? "الجامع" : "المكتبة الشاملة",
            name,
            "Diagnostic tool",
            document.RootElement.Clone());
    }
}
