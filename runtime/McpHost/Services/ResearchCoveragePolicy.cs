namespace ResearchTools.McpHost.Services;

internal static class ResearchCoveragePolicy
{
    internal static IReadOnlyList<string> GetUnreportedCoverage(
        IReadOnlyList<string> plannedCoverage,
        IReadOnlyList<string> declaredCovered,
        IReadOnlyList<string> declaredUncovered)
    {
        ArgumentNullException.ThrowIfNull(plannedCoverage);
        ArgumentNullException.ThrowIfNull(declaredCovered);
        ArgumentNullException.ThrowIfNull(declaredUncovered);

        var planned = plannedCoverage
            .Select(Normalize)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var covered = declaredCovered
            .Select(Normalize)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var uncovered = declaredUncovered
            .Select(Normalize)
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var unclassified = plannedCoverage
            .Where(item =>
            {
                var key = Normalize(item);
                return key.Length == 0 ||
                    covered.Contains(key) == uncovered.Contains(key);
            })
            .ToList();

        unclassified.AddRange(
            covered.Concat(uncovered)
                .Where(item => !planned.Contains(item)));

        return unclassified
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string Normalize(string value) =>
        string.Join(' ', (value ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
