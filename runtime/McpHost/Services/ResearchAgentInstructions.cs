namespace ResearchTools.McpHost.Services;

/// <summary>
/// Short standing guidance. The model interprets each request and chooses its own research method.
/// </summary>
internal static class ResearchAgentInstructions
{
    private const string CoreInstructions = """
أنت وكيل بحث يساعد الباحث على تنفيذ ما يطلبه بدقة. افهم السؤال كما قصده صاحبه، ولا تضف أهدافاً أو قوالب لم يطلبها.
اختر بنفسك أدوات MCP والاستعلامات والصفحات المناسبة، وتابع البحث بقدر ما يتطلبه السؤال.
لا تختلق معلومة أو نصاً أو نسبة أو بيانات مرجعية. استند إلى محتوى الصفحات التي فُتحت فعلاً، وصرّح بما تعذر التحقق منه.
قدّم النتيجة مباشرة بالشكل واللغة والحجم الذي طلبه الباحث.
""";

    internal static string Build(
        IReadOnlyList<string> selectedSources,
        IReadOnlyList<string> warnings)
    {
        var builder = new System.Text.StringBuilder(CoreInstructions);
        if (selectedSources.Count > 0)
            builder.AppendLine().Append("المصادر المتاحة: ").Append(string.Join("، ", selectedSources)).Append('.');
        if (warnings.Count > 0)
            builder.AppendLine().Append("تنبيهات التنفيذ: ").Append(string.Join(" | ", warnings));
        return builder.ToString();
    }
}
