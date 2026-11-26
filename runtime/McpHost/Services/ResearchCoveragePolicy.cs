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

        var planned = plannedCoverage.Select(Normalize).ToArray();
        var covered = new HashSet<int>();
        var uncovered = new HashSet<int>();
        var unrecognized = new List<string>();

        AddDeclarations(declaredCovered, planned, covered, unrecognized);
        AddDeclarations(declaredUncovered, planned, uncovered, unrecognized);

        var unclassified = Enumerable.Range(0, planned.Length)
            .Where(index =>
                planned[index].Length == 0 ||
                covered.Contains(index) == uncovered.Contains(index))
            .Select(index => plannedCoverage[index])
            .ToList();

        unclassified.AddRange(unrecognized);
        return unclassified.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static void AddDeclarations(
        IReadOnlyList<string> declarations,
        IReadOnlyList<string> planned,
        HashSet<int> destination,
        List<string> unrecognized)
    {
        foreach (var declaration in declarations)
        {
            var value = Normalize(declaration);
            if (value.Length == 0)
                continue;

            var index = ResolveIndex(value, planned);
            if (index >= 0)
                destination.Add(index);
            else
                unrecognized.Add(declaration);
        }
    }

    private static int ResolveIndex(string value, IReadOnlyList<string> planned)
    {
        // Stable coverage identifiers avoid requiring a model to reproduce
        // long Arabic facet descriptions character-for-character.
        if (value.Length > 1 &&
            (value[0] == 'C' || value[0] == 'c') &&
            int.TryParse(value.AsSpan(1), out var number) &&
            number > 0 && number <= planned.Count)
            return number - 1;

        for (var index = 0; index < planned.Count; index++)
        {
            if (string.Equals(value, planned[index], StringComparison.OrdinalIgnoreCase))
                return index;
        }

        return -1;
    }

    private static string Normalize(string value) =>
        string.Join(' ', (value ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
