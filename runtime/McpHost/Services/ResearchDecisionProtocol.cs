using System.Text;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Minimal transport instructions for the text-based Aurora/MCP bridge.
/// Research choices and answer structure are left to the model and the user request.
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
        builder.AppendLine("استخدم أدوات MCP عند الحاجة. لاستدعاء أداة، أرسل كتلة واحدة بالصيغة التالية، مع نسخ الاسم والمعاملات من فهرس الأدوات:");
        builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        builder.AppendLine("بعد انتهاء البحث، اكتب الإجابة مباشرة وبالطريقة التي طلبها الباحث، دون قالب بحثي مفروض.");

        if (!allowFinal)
        {
            builder.AppendLine("لا تنه المهمة الآن؛ أكمل متطلبات الوصول إلى المصادر:");
            if (pendingSourceLabels.Count > 0)
                builder.Append("نفّذ بحثاً في المصادر التي لم تُبحث بعد: ")
                    .Append(string.Join("، ", pendingSourceLabels))
                    .AppendLine(".");
            if (pendingPageSourceLabels.Count > 0)
                builder.Append("افتح صفحة مناسبة من نتائج البحث في: ")
                    .Append(string.Join("، ", pendingPageSourceLabels))
                    .AppendLine(".");
        }
        else
        {
            builder.AppendLine("اكتملت متطلبات الوصول الأساسية. أجب الآن إن كانت الأدلة كافية، أو واصل البحث وفتح الصفحات إن احتاج السؤال إلى مزيد.");
        }

        return builder.ToString();
    }

    internal static string BuildInvalidDecisionRepair(
        bool allowFinal,
        IReadOnlyList<string> pendingSourceLabels,
        IReadOnlyList<string> pendingPageSourceLabels)
    {
        ArgumentNullException.ThrowIfNull(pendingSourceLabels);
        ArgumentNullException.ThrowIfNull(pendingPageSourceLabels);

        var builder = new StringBuilder();
        builder.AppendLine("لم أتمكن من فهم القرار السابق. أعده بصيغة صحيحة واحدة.");
        if (!allowFinal)
        {
            builder.AppendLine("لا تكتب الإجابة النهائية بعد؛ استخدم كتلة <tool_call> صحيحة لإكمال البحث أو فتح الصفحات المطلوبة.");
            if (pendingSourceLabels.Count > 0)
                builder.Append("المصادر التي لم تُبحث: ").Append(string.Join("، ", pendingSourceLabels)).AppendLine(".");
            if (pendingPageSourceLabels.Count > 0)
                builder.Append("الصفحات التي يلزم فتحها: ").Append(string.Join("، ", pendingPageSourceLabels)).AppendLine(".");
            builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        }
        else
        {
            builder.AppendLine("يمكنك متابعة البحث بكتلة <tool_call> صحيحة أو تقديم الإجابة مباشرة.");
            builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        }

        return builder.ToString();
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث. أجب الآن بما تثبته الصفحات التي فُتحت فعلاً، وميّز ما تحقق مما تعذر التحقق منه. لا تستدعِ أدوات جديدة.";
}
