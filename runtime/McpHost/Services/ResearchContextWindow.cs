using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Builds a bounded conversation window for each local model turn.
/// The current user request is pinned while the newest surrounding messages
/// are retained in their original chronological order.
/// </summary>
internal static class ResearchContextWindow
{
    internal static IReadOnlyList<DuckChatMessage> Compact(
        IReadOnlyList<DuckChatMessage> messages,
        int maxConversationCharacters,
        int maxMessageCharacters,
        int pinnedMessageIndex = -1)
    {
        ArgumentNullException.ThrowIfNull(messages);
        if (maxConversationCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxConversationCharacters));
        if (maxMessageCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxMessageCharacters));

        if (messages.Count == 0)
            return [];

        var selected = new List<(int Index, DuckChatMessage Message)>();
        var usedCharacters = 0;

        if (pinnedMessageIndex >= 0 && pinnedMessageIndex < messages.Count)
        {
            var pinned = LimitMessage(messages[pinnedMessageIndex], maxMessageCharacters);
            selected.Add((pinnedMessageIndex, pinned));
            usedCharacters = pinned.Content?.Length ?? 0;
        }

        for (var index = messages.Count - 1; index >= 0; index--)
        {
            if (index == pinnedMessageIndex || usedCharacters >= maxConversationCharacters)
                continue;

            var remaining = maxConversationCharacters - usedCharacters;
            var messageLimit = Math.Min(maxMessageCharacters, remaining);
            if (messageLimit <= 0)
                break;

            var message = messages[index];
            var content = message.Content ?? string.Empty;
            if (content.Length > messageLimit)
            {
                var clipped = Clip(content, messageLimit);
                selected.Add((index, message with { Content = clipped }));
                usedCharacters += clipped.Length;
                break;
            }

            selected.Add((index, message));
            usedCharacters += content.Length;
        }

        return selected
            .OrderBy(item => item.Index)
            .Select(item => item.Message)
            .ToArray();
    }

    private static DuckChatMessage LimitMessage(DuckChatMessage message, int maximum)
    {
        var content = message.Content ?? string.Empty;
        return content.Length <= maximum
            ? message
            : message with { Content = Clip(content, maximum) };
    }

    private static string Clip(string content, int maximum)
    {
        const string marker = " [اقتُطع لتحديد سياق البحث]";
        if (content.Length <= maximum)
            return content;
        if (maximum <= marker.Length)
            return content[..maximum];

        return content[..(maximum - marker.Length)] + marker;
    }
}
