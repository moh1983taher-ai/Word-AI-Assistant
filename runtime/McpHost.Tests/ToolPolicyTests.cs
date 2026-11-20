using System.Text.Json;
using ModelContextProtocol.Protocol;
using ResearchTools.McpHost.Security;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ToolPolicyTests
{
    [Theory]
    [InlineData("search", true)]
    [InlineData("list_books", true)]
    [InlineData("get_page", true)]
    [InlineData("shamela_find_scoped", true)]
    [InlineData("shamela_open_many", true)]
    [InlineData("delete_book", false)]
    [InlineData("update_record", false)]
    [InlineData("create_document", false)]
    [InlineData("send_message", false)]
    [InlineData("execute_command", false)]
    [InlineData("admin_search", false)]
    public void AppliesReadOnlyToolAllowList(string name, bool expected)
    {
        Assert.Equal(expected, ToolPolicy.IsAllowed(CreateTool(name)));
    }

    [Fact]
    public void RejectsToolsMarkedAsMutating()
    {
        Assert.False(ToolPolicy.IsAllowed(CreateTool("search", readOnly: false)));
        Assert.False(ToolPolicy.IsAllowed(CreateTool("search", destructive: true)));
    }

    [Fact]
    public void AcceptsExplicitlyReadOnlySearchTool()
    {
        Assert.True(ToolPolicy.IsAllowed(CreateTool("search", readOnly: true, destructive: false)));
    }

    private static Tool CreateTool(string name, bool? readOnly = null, bool? destructive = null)
    {
        using var schema = JsonDocument.Parse("{\"type\":\"object\",\"properties\":{}}");
        return new Tool
        {
            Name = name,
            InputSchema = schema.RootElement.Clone(),
            Annotations = new ToolAnnotations
            {
                ReadOnlyHint = readOnly,
                DestructiveHint = destructive
            }
        };
    }
}
