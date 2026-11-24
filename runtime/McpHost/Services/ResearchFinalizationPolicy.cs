namespace ResearchTools.McpHost.Services;

/// <summary>
/// Final answers are allowed only after every selected source has completed its
/// initial search and every source with usable hits has at least one verified page open.
/// </summary>
internal static class ResearchFinalizationPolicy
{
    internal static bool CanFinalize(
        int pendingSearchSourceCount,
        int unverifiedPageSourceCount) =>
        pendingSearchSourceCount == 0 && unverifiedPageSourceCount == 0;
}
