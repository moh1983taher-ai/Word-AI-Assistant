using ResearchTools.McpHost.Services;
using ResearchTools.McpHost.Models;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchContextWindowTests
{
    [Fact]
    public void CompactPreservesPinnedCurrentRequestAndNewestToolResultWithinBudget()
    {
        var messages = new[]
        {
            new DuckChatMessage("user", new string('o', 5_000)),
            new DuckChatMessage("user", "سؤال المستخدم الحالي"),
            new DuckChatMessage("assistant", "{\"type\":\"tool_call\"}"),
            new DuckChatMessage("user", new string('r', 9_000))
        };

        var compacted = ResearchContextWindow.Compact(
            messages,
            maxConversationCharacters: 8_000,
            maxMessageCharacters: 6_000,
            pinnedMessageIndex: 1);

        Assert.InRange(compacted.Sum(message => message.Content?.Length ?? 0), 1, 8_000);
        Assert.Contains(compacted, message => message.Content == "سؤال المستخدم الحالي");
        Assert.Equal(6_000, compacted[^1].Content!.Length);
        Assert.StartsWith(new string('r', 100), compacted[^1].Content);
        Assert.EndsWith("[اقتُطع لتحديد سياق البحث]", compacted[^1].Content);
        Assert.DoesNotContain(compacted, message => message.Content == new string('o', 5_000));
    }

    [Fact]
    public void CompactRetainsNewestMessagesInOriginalOrderWhenNoRequestIsPinned()
    {
        var messages = new[]
        {
            new DuckChatMessage("user", new string('a', 80)),
            new DuckChatMessage("assistant", new string('b', 80)),
            new DuckChatMessage("user", new string('c', 80))
        };

        var compacted = ResearchContextWindow.Compact(
            messages,
            maxConversationCharacters: 100,
            maxMessageCharacters: 100);

        Assert.Equal(100, compacted.Sum(message => message.Content?.Length ?? 0));
        Assert.Equal(new string('b', 20), compacted[0].Content);
        Assert.Equal(new string('c', 80), compacted[1].Content);
    }

    [Fact]
    public void CompactTruncatesOversizedPinnedRequestWithoutDroppingIt()
    {
        var messages = new[]
        {
            new DuckChatMessage("user", new string('x', 15_000))
        };

        var compacted = ResearchContextWindow.Compact(
            messages,
            maxConversationCharacters: 10_000,
            maxMessageCharacters: 12_000,
            pinnedMessageIndex: 0);

        Assert.Single(compacted);
        Assert.Equal(10_000, compacted[0].Content!.Length);
        Assert.EndsWith("[اقتُطع لتحديد سياق البحث]", compacted[0].Content);
    }

    [Fact]
    public void CompactLeavesMessagesUnchangedWhenTheyFitTheBudget()
    {
        var messages = new[]
        {
            new DuckChatMessage("user", "السؤال"),
            new DuckChatMessage("assistant", "السياق")
        };

        var compacted = ResearchContextWindow.Compact(
            messages,
            maxConversationCharacters: 20_000,
            maxMessageCharacters: 12_000,
            pinnedMessageIndex: 0);

        Assert.Equal(messages, compacted);
    }
}
