namespace ResearchTools.McpHost.Models;

public static class ResearchSourceSelection
{
    public static readonly string[] AllSourceIds = ["aljam3", "shamela"];

    public static IReadOnlyList<string> Normalize(IEnumerable<string>? requested)
    {
        var requestedIds = (requested ?? [])
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id.Trim())
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        var unknown = requestedIds
            .Where(id => !AllSourceIds.Contains(id, StringComparer.Ordinal))
            .ToArray();

        if (unknown.Length > 0)
            throw new ArgumentException("مصادر MCP غير معروفة: " + string.Join("، ", unknown));

        return AllSourceIds
            .Where(id => requestedIds.Contains(id, StringComparer.Ordinal))
            .ToArray();
    }

    public static IReadOnlyList<string> Missing(
        IEnumerable<string> selected,
        IEnumerable<string> completed)
    {
        var completedIds = new HashSet<string>(completed, StringComparer.Ordinal);
        return selected.Where(id => !completedIds.Contains(id)).ToArray();
    }

    public static string GetLabel(string sourceId) => sourceId switch
    {
        "aljam3" => "الجامع",
        "shamela" => "المكتبة الشاملة",
        _ => sourceId
    };
}
