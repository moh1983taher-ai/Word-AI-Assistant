using System.Text;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Aurora-compatible tagged tool-call protocol for DuckAI, with a guarded JSON
/// final-answer shape. The host validates every parsed call against the live MCP schema.
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
        builder.AppendLine("بروتوكول البحث: استدعاء الأدوات بصيغة Aurora النصية، والإجابة النهائية بصيغة JSON.");
        builder.AppendLine("أخرج قراراً واحداً فقط، دون Markdown أو شرح خارج الصيغة المطلوبة.");
        builder.AppendLine("لاستدعاء أداة MCP أخرج كتلة واحدة بالصيغة الآتية، وانسخ الاسم المركب حرفياً من فهرس الأدوات:");
        builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        builder.AppendLine("الاسم المركب يتكون من معرّف المكتبة ثم __ ثم اسم الأداة، مثل aljam3__search. لا تضع server أو tool في حقول منفصلة.");
        builder.AppendLine("للإجابة النهائية أخرج: {\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"محور مغطى من الخطة\"],\"uncovered\":[]}. يجب أن تكون القائمتان موجودتين دائماً.");
        builder.AppendLine("التزم حرفياً بمعاملات الأداة. لا تخمّن المعرّفات ولا تنقلها بين المكتبات.");

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
            builder.AppendLine("راجع نتائج جميع الاستعلامات وتغطية محاور خطة البحث. إذا بقي محور غير مغطى، أو نتيجة قوية لم تُفحص، أو احتجت إلى مقارنة مراجع إضافية، فواصل باستخدام كتلة <tool_call>.");
            builder.AppendLine("لا تكتف بصفحة واحدة ولا بعدد شكلي من المراجع. افتح عدداً مناسباً من الصفحات ذات الصلة بحسب نوع السؤال واتساعه، واستبعد الصفحات التي لا تتناول المسألة.");
            builder.AppendLine("في طلب النص المحدد تحقق من موضع النص وسياقه؛ وفي السؤال العام اجمع معلومات متنوعة تغطي المحاور. صنّف كل محور في covered إذا دعمته الأدلة، أو في uncovered إذا تعذر إثباته بعد استقصاء مناسب، واذكر سبب التعذر وحدود البحث في answer. لا تترك أي محور بلا تصنيف، ولا تدّع تغطية ما لم يثبت.");
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
        builder.AppendLine("كان القرار السابق غير صالح. أعد قراراً واحداً صحيحاً فقط، دون Markdown أو شرح.");
        builder.AppendLine("استعمل الاسم المركب للأداة كما ورد حرفياً في فهرس MCP، وأرسل الاستدعاء داخل كتلة <tool_call>.");
        builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        if (!allowFinal)
        {
            builder.AppendLine("لا يُسمح الآن بقرار final؛ أخرج كتلة <tool_call> فقط.");
            if (pendingSourceLabels.Count > 0)
                builder.Append("المصادر التي لم يكتمل البحث فيها: ")
                    .Append(string.Join("، ", pendingSourceLabels)).AppendLine(".");
            if (pendingPageSourceLabels.Count > 0)
                builder.Append("مصادر أعادت نتائج وتحتاج إلى فتح مواضع مرشحة: ")
                    .Append(string.Join("، ", pendingPageSourceLabels)).AppendLine(".");
            builder.AppendLine("اختر أداة البحث أو فتح الصفحة المناسبة من المخطط، ولا تكرر فتح صفحة تم التحقق منها إذا بقي مصدر آخر غير مفحوص.");
            builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        }
        else
        {
            builder.AppendLine("يمكنك إخراج كتلة <tool_call> لمتابعة البحث أو final إذا غطت الأدلة المفتوحة محاور الخطة بالفعل.");
            builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
            builder.AppendLine("{\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"محور مغطى من الخطة\"],\"uncovered\":[]}");
        }

        return builder.ToString();
    }

    internal static string BuildCoverageCorrection(IReadOnlyList<string> missingCoverage)
    {
        ArgumentNullException.ThrowIfNull(missingCoverage);
        return "لا تعتمد الإجابة النهائية بعد؛ لم تثبت تغطية محاور الخطة الآتية: " +
            string.Join("، ", missingCoverage) +
            ". إن بقيت وسائل بحث مناسبة لم تجربها، فنفذ بحثاً تكميلياً أو افتح مراجع إضافية ذات صلة. وإن استنفدت المسارات المعقولة دون دليل، فضع المحور في uncovered واذكر سبب التعذر وحدود الاستقصاء في answer. لا تترك أي محور بلا تصنيف. " +
            "أخرج كتلة <tool_call> لمتابعة البحث، أو final بعد تصنيف جميع المحاور بين covered وuncovered دون ادعاء. " +
            "صيغة final: {\"type\":\"final\",\"answer\":\"...\",\"covered\":[\"...\"],\"uncovered\":[]}.";
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. " +
        "اعرض ما تسمح به الأدلة المفتوحة وبيّن المحاور التي لم يكتمل التحقق منها. " +
            "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"محور تحقق\"],\"uncovered\":[\"محور لم يكتمل\"]}. " +
        "لا تضف شرحاً أو Markdown خارج JSON.";
}
