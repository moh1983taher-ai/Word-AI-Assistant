namespace ResearchTools.McpHost.Services;

/// <summary>
/// Identifies requests whose wording makes source verification mandatory.
/// This is a safety gate, not a substitute for the model's broader tool selection.
/// </summary>
public static class ResearchIntentClassifier
{
    private static readonly string[] SourceIndicators =
    [
        "ابحث",
        "البحث عن",
        "بحثاً عن",
        "بحثًا عن",
        "بحث في",
        "بحث علمي",
        "مصدر",
        "مصادر",
        "مرجع",
        "مراجع",
        "توثيق",
        "وثّق",
        "وثق",
        "استشهد",
        "استشهاد",
        "الدراسات السابقة",
        "دراسات سابقة",
        "دراسات حول",
        "دراسات عن",
        "أبحاث",
        "ابحاث",
        "رسالة علمية",
        "رسائل علمية",
        "المكتبة الشاملة",
        "تخريج",
        "أقوال العلماء",
        "اقوال العلماء",
        "أقوال الأصوليين",
        "اقوال الاصوليين",
        "أقوال الفقهاء",
        "اقوال الفقهاء",
        "نصوص العلماء",
        "نصوص الأصوليين",
        "citations",
        "references",
        "bibliography",
        "sources",
        "search for",
        "search in",
        "find studies",
        "research sources"
    ];

    public static bool RequiresMcpSources(string? userMessage)
    {
        if (string.IsNullOrWhiteSpace(userMessage))
            return false;

        return SourceIndicators.Any(indicator =>
            userMessage.Contains(indicator, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// A source-based request cannot be finalized until at least one selected MCP
    /// tool has completed successfully. General conversational requests remain model-led.
    /// </summary>
    public static bool MustUseMcpBeforeFinal(string? userMessage, int successfulToolCalls) =>
        successfulToolCalls <= 0 && RequiresMcpSources(userMessage);

    /// <summary>
    /// Returns every selected library that still needs a successful search.
    /// This helper is called only by the dedicated MCP research endpoint, not
    /// ordinary chat, so source coverage must not depend on trigger words in the
    /// user's natural-language question.
    /// </summary>
    public static IReadOnlyList<string> GetPendingSearchSources(
        string? userMessage,
        IReadOnlyCollection<string> selectedSourceIds,
        IReadOnlySet<string> successfullySearchedSourceIds)
    {
        ArgumentNullException.ThrowIfNull(selectedSourceIds);
        ArgumentNullException.ThrowIfNull(successfullySearchedSourceIds);

        return selectedSourceIds
            .Where(sourceId => !successfullySearchedSourceIds.Contains(sourceId))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }
}
