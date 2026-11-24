namespace ResearchTools.McpHost.Services;

internal static class ResearchCoveragePolicy
{
    internal static IReadOnlyList<string> GetMissingCoverage(
        IReadOnlyList<string> plannedCoverage,
        IReadOnlyList<string> declaredCovered,
        IReadOnlyList<string> declaredUncovered)
    {
        ArgumentNullException.ThrowIfNull(plannedCoverage);
        ArgumentNullException.ThrowIfNull(declaredCovered);
        ArgumentNullException.ThrowIfNull(declaredUncovered);

        var covered = declaredCovered
            .Select(Normalize)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return plannedCoverage
            .Where(item => !covered.Contains(Normalize(item)))
            .Concat(declaredUncovered)
            .Select(item => (item ?? string.Empty).Trim())
            .Where(item => item.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string Normalize(string value) =>
        string.Join(' ', (value ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
