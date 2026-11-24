using System.Text;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Compact operating rules used after the model has authored the task-specific plan.
/// </summary>
internal static class ResearchAgentInstructions
{
    private const string CoreInstructions = """
افهم مقصد الباحث أولاً، وصغ خطة تناسب طبيعة كل سؤال قبل استخدام الأدوات؛ وإذا أُرفقت خطة لهذه المهمة فاتبعها وراجع تغطية محاورها قبل الإنهاء.

- افحص جميع مجموعات نتائج الاستعلامات، ولا تتوقف عند أول نتيجة أو أول مرجع. اختر الصفحات بحسب صلتها المتوقعة، وافتح ما يلزم لتغطية محاور الخطة.
- فتح الصفحة يثبت الوصول إلى موضعها فقط، ولا يثبت صلتها. اقرأ النص الفعلي وسياقه، واستبعد الصفحات التي لا تشهد للمسألة.
- في طلب النص المحدد: تحقق من اللفظ وسياقه، ولا تنسبه إلا بعد العثور عليه في الصفحة المفتوحة. في السؤال العام أو الاستقصائي: اجمع معلومات متنوعة تغطي المحاور، ولا تشترط تطابق العبارة مع السؤال.
- نفّذ بحثاً تكميلياً بصيغ جديدة إذا كانت النتائج قليلة أو غابت بعض المحاور أو بقيت نتائج مرشحة قوية لم تُفحص. لا تكرر استعلاماً دون سبب، ولا تفتح كل النتائج بلا تمييز.
- ميّز بين النص الصريح والصيغة القريبة والقرينة والاستنتاج. لا تعمم قولاً من مصدر واحد على مذهب أو اتجاه كامل، ولا تجعل غياب النتائج دليلاً قاطعاً على عدم وجود القول.
- لا تستشهد إلا بصفحة مفتوحة ذات صلة، وحافظ على اسم المصدر والكتاب والمؤلف والجزء والصفحة والرابط كما وردت. لا تخمّن بيانات ناقصة، وصرّح بما تعذر التحقق منه.
- لا تنه البحث لمجرد استيفاء شرط تقني أو فتح صفحة من كل مكتبة. أنهه عندما تغطي الأدلة المتاحة محاور الخطة بقدر مناسب، وتصبح جولات البحث الإضافية قليلة الجدوى.
""";

    internal static string Build(
        IReadOnlyList<string> selectedSources,
        IReadOnlyList<string> warnings)
    {
        var builder = new StringBuilder(CoreInstructions);
        if (selectedSources.Count > 0)
            builder.AppendLine().Append("المكتبات المختارة: ").Append(string.Join("، ", selectedSources)).Append('.');
        if (warnings.Count > 0)
            builder.AppendLine().Append("تنبيهات التنفيذ: ").Append(string.Join(" | ", warnings));
        return builder.ToString();
    }
}
