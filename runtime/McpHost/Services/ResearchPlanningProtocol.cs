using System.Text;
using System.Text.Json;
using System.Text.Encodings.Web;

namespace ResearchTools.McpHost.Services;

internal static class ResearchPlanningProtocol
{
    internal static string BuildInitialInstructions(IReadOnlyList<string> selectedSources)
    {
        ArgumentNullException.ThrowIfNull(selectedSources);
        var builder = new StringBuilder();
        builder.AppendLine("أنت مساعد بحث علمي. هذه مرحلة فهم السؤال والتخطيط فقط؛ لا تبحث ولا تجب عن المسألة الآن.");
        builder.AppendLine("افهم مقصد الباحث، وحدد نوع المهمة ونطاقها، ثم ضع بنفسك خطة بحث قابلة للتنفيذ في المكتبات المختارة.");
        builder.AppendLine("ميّز بين طلب نص بعينه، واستقصاء صيغ أو أقوال، وسؤال عام، ومقارنة أو توثيق. لا تستخدم خطة واحدة لكل الأنواع.");
        builder.AppendLine("اختر استعلامات متكاملة لا مكررة: ابدأ بلفظ السؤال عند ملاءمته، وأضف الصيغ والمرادفات والمفاهيم الفرعية اللازمة. في السؤال العام قسّم الموضوع إلى محاور تغطي جوانبه المهمة؛ وفي طلب النص المحدد ركز على العبارة وصيغها القريبة.");
        builder.AppendLine("حدد ما ينبغي أن تغطيه الإجابة، ولا تضع عددًا شكليًا للاستعلامات أو المراجع. اختر عدد الاستعلامات بحسب طبيعة المهمة، وبحد أقصى 8 استعلامات موجزة.");
        builder.AppendLine("أخرج كائن JSON واحدًا فقط، دون Markdown أو نص خارجه، بهذه البنية: " +
            "{\"type\":\"plan\",\"intent\":\"مقصد الباحث\",\"kind\":\"exact_text أو multiple_forms أو broad أو comparison أو attribution أو other\",\"queries\":[{\"query\":\"عبارة بحث\",\"purpose\":\"ما الذي يستهدفه الاستعلام\"}],\"coverage\":[\"محور مطلوب تغطيته\"]}.");
        builder.AppendLine("اجعل كل استعلام مستقلًا ومفيدًا، ولا تضع تعليمات المستخدم مثل (مع التوثيق ورابط الصفحة) ضمن كلمات البحث إلا إذا كانت جزءًا من النص المطلوب.");
        builder.Append("المكتبات المتاحة للاختيار: ").Append(string.Join("، ", selectedSources)).Append('.');
        return builder.ToString();
    }

    internal static string BuildRepairInstructions() =>
        "أعد إخراج خطة البحث فقط بصيغة JSON صحيحة. لا تبدأ البحث ولا تكتب إجابة. " +
        "استخدم هذا الشكل حرفيًا: {\"type\":\"plan\",\"intent\":\"...\",\"kind\":\"broad أو exact_text أو multiple_forms أو comparison أو attribution أو other\",\"queries\":[{\"query\":\"...\",\"purpose\":\"...\"}],\"coverage\":[\"...\"]}. " +
        "يجب أن تحتوي queries على استعلام واحد على الأقل، ولا تتجاوز ثمانية استعلامات، وأن تتضمن coverage محوراً واحداً على الأقل يحدد معيار اكتمال البحث. لا تضع Markdown أو نصاً خارج JSON.";

    internal static string FormatPlan(ResearchPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);
        var payload = new
        {
            intent = plan.Intent,
            kind = plan.Kind,
            queries = plan.Queries.Select(query => new { query = query.Query, purpose = query.Purpose }).ToArray(),
            coverage = plan.Coverage
        };
        return JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
    }
}
