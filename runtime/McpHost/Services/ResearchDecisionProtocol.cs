using System.Text;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Human-readable JSON decision protocol for Duck2API endpoints that ignore
/// response_format. The host still validates every parsed decision and tool call.
/// </summary>
internal static class ResearchDecisionProtocol
{
    internal static string BuildTurnInstructions(
        bool allowFinal,
        IReadOnlyList<string> pendingSourceLabels,
        IReadOnlyList<string>? pendingPageSourceLabels = null)
    {
        ArgumentNullException.ThrowIfNull(pendingSourceLabels);
        pendingPageSourceLabels ??= [];

        var builder = new StringBuilder();
        builder.AppendLine("بروتوكول قرار البحث (JSON عادي):");
        builder.AppendLine(
            "أخرج كائناً واحداً فقط، دون Markdown أو أي نص خارجه.");
        builder.AppendLine(
            "لاستدعاء أداة MCP أخرج: {\"type\":\"tool_call\",\"server\":\"aljam3 أو shamela\",\"tool\":\"اسم الأداة\",\"arguments\":{...}}.");
        builder.AppendLine(
            "للإجابة النهائية أخرج: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}.");
        builder.AppendLine(
            "استخدم اسم المكتبة واسم الأداة ومعاملاتها كما تظهر حرفياً في فهرس MCP ومخطط inputSchema. لا تخترع معرفات كتب أو صفحات، ولا تنقل معرفاً بين مكتبتين.");

        if (!allowFinal)
        {
            builder.AppendLine("لا يُسمح بالإجابة النهائية الآن.");
            if (pendingSourceLabels.Count > 0)
            {
                builder.Append("لم يكتمل البحث الأولي في: ")
                    .Append(string.Join("، ", pendingSourceLabels))
                    .AppendLine(". نفّذ بحثاً ناجحاً في كل مصدر متبقٍ.");
            }

            if (pendingPageSourceLabels.Count > 0)
            {
                builder.Append("أعادت المصادر الآتية نتائج فعلية، لكن لم يتحقق المضيف من فتح صفحة مطابقة فيها بعد: ")
                    .Append(string.Join("، ", pendingPageSourceLabels))
                    .AppendLine(". افتح صفحة واحدة على الأقل من كل مصدر مذكور، مستخدماً المعرّفات الواردة حرفياً في نتيجة البحث من المصدر نفسه، ثم افحص النص وسياقه.");
            }

            builder.AppendLine("لا تخرج قرار final ما دام أي من الشرطين السابقين غير مكتمل؛ أخرج قرار tool_call لأداة الفتح المناسبة.");
        }
        else
        {
            builder.AppendLine(
                "اكتمل البحث الأولي، وفتح المضيف صفحة متحققة من كل مصدر أعاد نتائج فعلية. يمكنك الآن صياغة الإجابة مع الاقتصار على الأدلة التي فُتحت وتحققت.");
        }

        return builder.ToString();
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. " +
        "اعرض ما تسمح به الأدلة المتاحة ووضّح أن البحث لم يكتمل. " +
        "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}. " +
        "لا تضف شرحاً أو Markdown خارج JSON.";
}
