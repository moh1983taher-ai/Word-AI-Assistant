using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Keeps each model turn focused on selected libraries that have not yet returned
/// a successful tool result. Once every selected library is covered, all selected
/// libraries become available again for targeted follow-up reading.
/// </summary>
internal static class ResearchSourceToolRouter
{
    internal static IReadOnlyList<ToolDescriptor> SelectForTurn(
        IReadOnlyList<ToolDescriptor> availableTools,
        IReadOnlyList<string> selectedServerIds,
        IEnumerable<string> successfullyCompletedServerIds)
    {
        ArgumentNullException.ThrowIfNull(availableTools);
        ArgumentNullException.ThrowIfNull(selectedServerIds);
        ArgumentNullException.ThrowIfNull(successfullyCompletedServerIds);

        var selected = new HashSet<string>(selectedServerIds, StringComparer.Ordinal);
        if (selected.Count == 0)
            return [];

        var completed = new HashSet<string>(
            successfullyCompletedServerIds.Where(selected.Contains),
            StringComparer.Ordinal);
        var pending = selected.Where(server => !completed.Contains(server)).ToHashSet(StringComparer.Ordinal);

        // While required source coverage is incomplete, only pending sources are exposed.
        // After all selected sources succeed, targeted follow-up tools from those sources
        // may be used for reading and citation verification.
        var permitted = pending.Count > 0 ? pending : selected;

        return availableTools
            .Where(tool => permitted.Contains(tool.Server))
            .ToArray();
    }
}
