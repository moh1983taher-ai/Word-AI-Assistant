using System.Globalization;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Adaptive prompt, completion, and execution budgets shared by all research agents.
/// Defaults are intentionally generous and can be tuned for the selected model.
/// </summary>
internal sealed record ResearchContextBudget(
    int MaxContextCharacters,
    int MaxOutputTokens,
    TimeSpan ResearchTimeout,
    TimeSpan ModelRequestTimeout,
    TimeSpan FinalizationTimeout)
{
    private const int DefaultContextCharacters = 64_000;
    private const int MinimumContextCharacters = 16_000;
    private const int MaximumContextCharacters = 196_000;

    private const int DefaultOutputTokens = 8_192;
    private const int MinimumOutputTokens = 1_024;
    private const int MaximumOutputTokens = 32_768;

    private const int DefaultResearchTimeoutSeconds = 600;
    private const int MinimumResearchTimeoutSeconds = 180;
    private const int MaximumResearchTimeoutSeconds = 1_800;

    private const int DefaultModelRequestTimeoutSeconds = 120;
    private const int MinimumModelRequestTimeoutSeconds = 30;
    private const int MaximumModelRequestTimeoutSeconds = 300;

    private const int DefaultFinalizationTimeoutSeconds = 60;
    private const int MinimumFinalizationTimeoutSeconds = 20;
    private const int MaximumFinalizationTimeoutSeconds = 180;

    private const int ProtocolReserveCharacters = 2_048;

    internal static ResearchContextBudget Load()
    {
        return new ResearchContextBudget(
            ReadBoundedSetting("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", DefaultContextCharacters,
                MinimumContextCharacters, MaximumContextCharacters),
            ReadBoundedSetting("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", DefaultOutputTokens,
                MinimumOutputTokens, MaximumOutputTokens),
            TimeSpan.FromSeconds(ReadBoundedSetting("RESEARCHTOOLS_RESEARCH_TIMEOUT_SECONDS",
                DefaultResearchTimeoutSeconds, MinimumResearchTimeoutSeconds, MaximumResearchTimeoutSeconds)),
            TimeSpan.FromSeconds(ReadBoundedSetting("RESEARCHTOOLS_MODEL_REQUEST_TIMEOUT_SECONDS",
                DefaultModelRequestTimeoutSeconds, MinimumModelRequestTimeoutSeconds, MaximumModelRequestTimeoutSeconds)),
            TimeSpan.FromSeconds(ReadBoundedSetting("RESEARCHTOOLS_FINALIZATION_TIMEOUT_SECONDS",
                DefaultFinalizationTimeoutSeconds, MinimumFinalizationTimeoutSeconds, MaximumFinalizationTimeoutSeconds)));
    }

    internal int GetAvailableInputCharacters(int systemPromptCharacters)
    {
        return Math.Max(
            1,
            MaxContextCharacters - Math.Max(0, systemPromptCharacters) - ProtocolReserveCharacters);
    }

    internal int GetEvidenceBudget(int availableCharacters, int estimatedEvidenceCharacters)
    {
        if (availableCharacters <= 0 || estimatedEvidenceCharacters <= 0)
            return 0;

        // Evidence can use up to two thirds of the available input space when
        // needed; unused evidence space remains available to conversation history.
        return Math.Min(
            Math.Max(1, availableCharacters * 2 / 3),
            estimatedEvidenceCharacters);
    }

    private static int ReadBoundedSetting(
        string name,
        int defaultValue,
        int minimum,
        int maximum)
    {
        var raw = Environment.GetEnvironmentVariable(name);
        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
            return defaultValue;

        return Math.Clamp(value, minimum, maximum);
    }
}
