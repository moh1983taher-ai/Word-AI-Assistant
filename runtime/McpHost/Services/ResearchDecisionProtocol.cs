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
        builder.AppendLine("أخرج كائناً واحداً فقط، دون Markdown أو أي نص خارجه.");
        builder.AppendLine("لاستدعاء أداة MCP أخرج: {\"type\":\"tool_call\",\"server\":\"aljam3 أو shamela\",\"tool\":\"اسم الأداة\",\"arguments\":{...}}.");
        builder.AppendLine("للإجابة النهائية أخرج: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}.");
        builder.AppendLine("التزم حرفياً بأسماء الأدوات ومعاملاتها في مخطط MCP، ولا تخمّن المعرّفات أو تنقلها بين المكتبات.");

        if (!allowFinal)
        {
            builder.AppendLine("لا تُنه البحث الآن.");
            if (pendingSourceLabels.Count > 0)
                builder.Append("لم يكتمل البحث الأولي في: ")
                    .Append(string.Join("، ", pendingSourceLabels))
                    .AppendLine(". نفّذ الاستعلامات المخططة في كل مصدر متبقٍ.");
            if (pendingPageSourceLabels.Count > 0)
                builder.Append("توجد نتائج تحتاج إلى فحص صفحاتها في: ")
                    .Append(string.Join("، ", pendingPageSourceLabels))
                    .AppendLine(". اختر المواضع الأوثق صلة وفق خطة البحث وافتحها من نتائج المكتبة نفسها؛ لا تفتح كل النتائج بلا تمييز.");
            builder.AppendLine("لا تصدر final قبل اكتمال متطلبات الوصول الأساسية، ثم تابع فحص الأدلة وتغطية الخطة قبل تقرير كفاية البحث.");
        }
        else
        {
            builder.AppendLine("اكتملت متطلبات الوصول الأساسية، لكن لا يعني ذلك اكتمال الاستقصاء.");
            builder.AppendLine("راجع نتائج جميع الاستعلامات وتغطية محاور خطة البحث. إذا بقي محور غير مغطى، أو نتيجة قوية لم تُفحص، أو احتجت إلى مقارنة مراجع إضافية، فواصل باستخدام tool_call.");
            builder.AppendLine("لا تكتف بصفحة واحدة ولا بعدد شكلي من المراجع. افتح عدداً مناسباً من الصفحات ذات الصلة بحسب نوع السؤال واتساعه، واستبعد الصفحات التي لا تتناول المسألة.");
            builder.AppendLine("في طلب النص المحدد تحقق من موضع النص وسياقه؛ وفي السؤال العام اجمع معلومات متنوعة تغطي المحاور. أخرج final عندما تصبح الأدلة المفتوحة كافية وتكون جولات البحث الإضافية قليلة الجدوى.");
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
        builder.AppendLine("كان القرار السابق غير صالح بصيغة JSON. أعد قراراً واحداً صحيحاً فقط، دون Markdown أو شرح.");
        builder.AppendLine("استعمل أسماء الخوادم والأدوات ومعاملاتها كما وردت حرفياً في فهرس MCP.");
        if (!allowFinal)
        {
            builder.AppendLine("لا يُسمح الآن بقرار final؛ أخرج tool_call فقط.");
            if (pendingSourceLabels.Count > 0)
                builder.Append("المصادر التي لم يكتمل البحث فيها: ")
                    .Append(string.Join("، ", pendingSourceLabels)).AppendLine(".");
            if (pendingPageSourceLabels.Count > 0)
                builder.Append("مصادر أعادت نتائج وتحتاج إلى فتح مواضع مرشحة: ")
                    .Append(string.Join("، ", pendingPageSourceLabels)).AppendLine(".");
            builder.AppendLine("اختر أداة البحث أو فتح الصفحة المناسبة من المخطط، ولا تكرر فتح صفحة تم التحقق منها إذا بقي مصدر آخر غير مفحوص.");
            builder.AppendLine("{\"type\":\"tool_call\",\"server\":\"aljam3 أو shamela\",\"tool\":\"اسم الأداة\",\"arguments\":{...}}");
        }
        else
        {
            builder.AppendLine("يمكنك إخراج tool_call لمتابعة البحث أو final إذا غطت الأدلة المفتوحة محاور الخطة بالفعل.");
            builder.AppendLine("{\"type\":\"tool_call\",\"server\":\"aljam3 أو shamela\",\"tool\":\"اسم الأداة\",\"arguments\":{...}}");
            builder.AppendLine("{\"type\":\"final\",\"answer\":\"نص الإجابة\"}");
        }

        return builder.ToString();
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. " +
        "اعرض ما تسمح به الأدلة المفتوحة وبيّن المحاور التي لم يكتمل التحقق منها. " +
        "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}. " +
        "لا تضف شرحاً أو Markdown خارج JSON.";
}
