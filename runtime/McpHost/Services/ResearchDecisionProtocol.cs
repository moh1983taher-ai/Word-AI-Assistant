using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

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
        IReadOnlyList<string>? pendingPageSourceLabels = null,
        IReadOnlyList<string>? plannedCoverage = null)
    {
        ArgumentNullException.ThrowIfNull(pendingSourceLabels);
        pendingPageSourceLabels ??= [];
        plannedCoverage ??= [];

        var builder = new StringBuilder();
        builder.AppendLine("بروتوكول البحث: استدعاء الأدوات بصيغة Aurora النصية، والإجابة النهائية بصيغة JSON.");
        builder.AppendLine("أخرج قراراً واحداً فقط، دون Markdown أو شرح خارج الصيغة المطلوبة.");
        builder.AppendLine("لاستدعاء أداة MCP أخرج كتلة واحدة بالصيغة الآتية، وانسخ الاسم المركب حرفياً من فهرس الأدوات:");
        builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
        builder.AppendLine("الاسم المركب يتكون من معرّف المكتبة ثم __ ثم اسم الأداة، مثل aljam3__search. لا تضع server أو tool في حقول منفصلة.");
        builder.AppendLine("للإجابة النهائية أخرج: {\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"C1\"],\"uncovered\":[\"C2\"]}. يجب أن تكون القائمتان موجودتين دائماً.");
        builder.AppendLine("التزم حرفياً بمعاملات الأداة. لا تخمّن المعرّفات ولا تنقلها بين المكتبات.");

        AppendCoverageMap(builder, plannedCoverage);

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
            builder.AppendLine("صنّف كل محور من قائمة التغطية باستخدام رمزه فقط: covered إذا دعمته الأدلة، أو uncovered إذا تعذر إثباته بعد استقصاء مناسب. لا تعِد صياغة المحاور ولا تنشئ محاور جديدة. اذكر أسباب التعذر وحدود البحث في answer، ولا تدّع تغطية ما لم يثبت.");
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
            builder.AppendLine("يمكنك إخراج كتلة <tool_call> لمتابعة البحث أو final بعد تصنيف جميع محاور الخطة برموزها الثابتة.");
            builder.AppendLine("<tool_call>{\"name\":\"server__tool_name\",\"arguments\":{...}}</tool_call>");
            builder.AppendLine("{\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"C1\"],\"uncovered\":[\"C2\"]}");
        }

        return builder.ToString();
    }

    internal static string BuildCoverageCorrection(
        IReadOnlyList<string> plannedCoverage,
        IReadOnlyList<string> missingCoverage)
    {
        ArgumentNullException.ThrowIfNull(plannedCoverage);
        ArgumentNullException.ThrowIfNull(missingCoverage);

        var missing = missingCoverage.Select(Normalize)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var missingIds = plannedCoverage
            .Select((item, index) => new { Item = item, Id = $"C{index + 1}" })
            .Where(entry => missing.Contains(Normalize(entry.Item)))
            .Select(entry => entry.Id)
            .ToArray();

        var builder = new StringBuilder();
        builder.AppendLine("لم يعتمد المضيف التصنيف السابق؛ لأن حقلي covered وuncovered يجب أن يحتويَا رموز المحاور لا صياغات من إنشائك.");
        AppendCoverageMap(builder, plannedCoverage);
        if (missingIds.Length > 0)
            builder.Append("المحاور التي يلزم تصنيفها الآن: ")
                .Append(string.Join("، ", missingIds))
                .AppendLine(".");
        builder.AppendLine("أعد قراراً نهائياً بعد وضع كل رمز أصلي مرة واحدة فقط في covered أو uncovered. لا تضف عبارات أخرى إلى القائمتين، ولا تغيّر معرّفات المحاور. إذا لم يثبت محور، ضعه في uncovered وبيّن السبب في answer.");
        builder.AppendLine("صيغة final: {\"type\":\"final\",\"answer\":\"...\",\"covered\":[\"C1\"],\"uncovered\":[\"C2\"]}.");
        return builder.ToString();
    }

    internal static string BuildTimeoutFinalInstructions() =>
        "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. " +
        "اعرض ما تسمح به الأدلة المفتوحة وبيّن المحاور التي لم يكتمل التحقق منها. " +
        "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\",\"covered\":[\"C1\"],\"uncovered\":[\"C2\"]}. " +
        "لا تضف شرحاً أو Markdown خارج JSON.";

    private static void AppendCoverageMap(StringBuilder builder, IReadOnlyList<string> plannedCoverage)
    {
        if (plannedCoverage.Count == 0)
            return;

        builder.AppendLine("قائمة المحاور الأصلية ورموزها الثابتة (انسخ الرمز فقط داخل covered أو uncovered):");
        for (var index = 0; index < plannedCoverage.Count; index++)
        {
            var serialized = JsonSerializer.Serialize(
                plannedCoverage[index],
                new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
            builder.Append("C").Append(index + 1).Append(" = ")
                .AppendLine(serialized);
        }
        builder.AppendLine("يجب تصنيف كل رمز من هذه الرموز مرة واحدة بالضبط؛ لا تستخدم نص المحور بدلاً من الرمز، ولا تنشئ رمزاً أو محوراً جديداً.");
    }

    private static string Normalize(string value) =>
        string.Join(' ', (value ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
