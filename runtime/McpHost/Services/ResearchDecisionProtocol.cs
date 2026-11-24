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
        IReadOnlyList<string> pendingSourceLabels)
    {
        ArgumentNullException.ThrowIfNull(pendingSourceLabels);

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
            builder.Append("لا يُسمح بالإجابة النهائية الآن. لم يكتمل البحث في: ")
                .Append(string.Join("، ", pendingSourceLabels))
                .AppendLine(". نفّذ بحثاً ناجحاً في كل مصدر متبقٍ، ثم افتح الصفحات المطابقة للنتائج.");
        }
        else
        {
            builder.AppendLine(
                "اكتمل البحث الأولي في المصادر التي اختارها المستخدم. يمكنك فتح الصفحات التي ستستشهد بها أو تنفيذ بحث تكميلي، ثم أجب عندما تتحقق من الأدلة.");
        }

        return builder.ToString();
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. " +
        "اعرض ما تسمح به الأدلة المتاحة ووضّح أن البحث لم يكتمل. " +
        "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}. " +
        "لا تضف شرحاً أو Markdown خارج JSON.";
}
