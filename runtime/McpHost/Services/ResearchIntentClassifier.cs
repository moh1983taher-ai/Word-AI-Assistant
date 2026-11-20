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
        "دراسة",
        "دراسات",
        "أبحاث",
        "ابحاث",
        "رسالة علمية",
        "رسائل علمية",
        "كتاب",
        "كتب",
        "المكتبة الشاملة",
        "الجامع",
        "تخريج",
        "أقوال العلماء",
        "اقوال العلماء",
        "نصوص العلماء",
        "citations",
        "references",
        "bibliography",
        "sources",
        "search for",
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
}
