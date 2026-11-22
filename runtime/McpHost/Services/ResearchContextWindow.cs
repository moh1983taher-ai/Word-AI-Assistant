using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Builds a bounded conversation window for each local model turn.
/// Native tool-call messages and their matching tool results are kept as an
/// indivisible history group so compaction cannot leave invalid OpenAI tool history.
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

        var groups = BuildGroups(messages);
        var selected = new List<(int Index, DuckChatMessage Message)>();
        var usedCharacters = 0;

        var pinnedGroup = pinnedMessageIndex >= 0 && pinnedMessageIndex < messages.Count
            ? groups.FirstOrDefault(group => group.Any(item => item.Index == pinnedMessageIndex))
            : null;

        if (pinnedGroup is not null)
        {
            var pinned = messages[pinnedMessageIndex];
            var bounded = LimitMessage(
                pinned,
                Math.Min(maxMessageCharacters, maxConversationCharacters));
            selected.Add((pinnedMessageIndex, bounded));
            usedCharacters = GetMessageCharacterCount(bounded);
        }

        for (var groupIndex = groups.Count - 1; groupIndex >= 0; groupIndex--)
        {
            var group = groups[groupIndex];
            if (ReferenceEquals(group, pinnedGroup))
                continue;

            var remaining = maxConversationCharacters - usedCharacters;
            if (remaining <= 0)
                break;

            var groupCharacters = group.Sum(item => GetMessageCharacterCount(item.Message));
            if (groupCharacters <= remaining &&
                group.All(item => (item.Message.Content?.Length ?? 0) <= maxMessageCharacters))
            {
                selected.AddRange(group);
                usedCharacters += groupCharacters;
                continue;
            }

            // Never clip a native tool-call/result group: doing so can orphan
            // tool_call_id entries or change the function arguments. Older groups
            // can be omitted because the independent evidence ledger retains results.
            if (group.Count > 1 ||
                group[0].Message.ToolCalls is { Count: > 0 })
                continue;

            var message = group[0].Message;
            var content = message.Content ?? string.Empty;
            var messageLimit = Math.Min(maxMessageCharacters, remaining);
            if (content.Length > messageLimit)
            {
                var clipped = LimitMessage(message, messageLimit);
                selected.Add((group[0].Index, clipped));
                usedCharacters += GetMessageCharacterCount(clipped);
                break;
            }

            selected.Add(group[0]);
            usedCharacters += GetMessageCharacterCount(message);
        }

        return selected
            .OrderBy(item => item.Index)
            .Select(item => item.Message)
            .ToArray();
    }

    internal static int GetMessageCharacterCount(DuckChatMessage message)
    {
        var count = message.Content?.Length ?? 0;
        count += message.ToolCallId?.Length ?? 0;

        if (message.ToolCalls is not null)
        {
            foreach (var call in message.ToolCalls)
            {
                count += call.Id.Length;
                count += call.Type.Length;
                count += call.Function.Name.Length;
                count += call.Function.Arguments.Length;
            }
        }

        return count;
    }

    private static List<List<(int Index, DuckChatMessage Message)>> BuildGroups(
        IReadOnlyList<DuckChatMessage> messages)
    {
        var groups = new List<List<(int Index, DuckChatMessage Message)>>();
        var index = 0;

        while (index < messages.Count)
        {
            var group = new List<(int Index, DuckChatMessage Message)>
            {
                (index, messages[index])
            };

            if (messages[index].Role == "assistant" &&
                messages[index].ToolCalls is { Count: > 0 })
            {
                var expectedIds = messages[index].ToolCalls
                    .Select(call => call.Id)
                    .ToHashSet(StringComparer.Ordinal);

                var next = index + 1;
                while (next < messages.Count &&
                       messages[next].Role == "tool" &&
                       messages[next].ToolCallId is { } callId &&
                       expectedIds.Contains(callId))
                {
                    group.Add((next, messages[next]));
                    expectedIds.Remove(callId);
                    next++;
                    if (expectedIds.Count == 0)
                        break;
                }

                index = next;
            }
            else
            {
                index++;
            }

            groups.Add(group);
        }

        return groups;
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
