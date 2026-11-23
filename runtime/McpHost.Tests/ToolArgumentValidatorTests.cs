using System.Text.Json;
using ResearchTools.McpHost.Security;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ToolArgumentValidatorTests
{
    [Fact]
    public void AcceptsValidRequiredArguments()
    {
        using var schema = JsonDocument.Parse("""
            {"type":"object","required":["query"],"properties":{"query":{"type":"string"}},"additionalProperties":false}
            """);
        using var args = JsonDocument.Parse("""{"query":"فقه المعاملات"}""");

        Assert.Null(ToolArgumentValidator.Validate(schema.RootElement, args.RootElement));
    }

    [Fact]
    public void RejectsMissingRequiredArgument()
    {
        using var schema = JsonDocument.Parse("""
            {"type":"object","required":["query"],"properties":{"query":{"type":"string"}}}
            """);
        using var args = JsonDocument.Parse("""{}""");

        Assert.Contains("مفقود", ToolArgumentValidator.Validate(schema.RootElement, args.RootElement));
    }

    [Fact]
    public void RejectsUnknownPropertiesWhenSchemaForbidsThem()
    {
        using var schema = JsonDocument.Parse("""
            {"type":"object","properties":{"query":{"type":"string"}},"additionalProperties":false}
            """);
        using var args = JsonDocument.Parse("""{"other":"value"}""");

        Assert.Contains("غير معروف", ToolArgumentValidator.Validate(schema.RootElement, args.RootElement));
    }

    [Fact]
    public void RejectsWrongPropertyType()
    {
        using var schema = JsonDocument.Parse("""
            {"type":"object","properties":{"limit":{"type":"integer"}}}
            """);
        using var args = JsonDocument.Parse("""{"limit":"10"}""");

        Assert.Contains("نوع غير صالح", ToolArgumentValidator.Validate(schema.RootElement, args.RootElement));
    }
}
