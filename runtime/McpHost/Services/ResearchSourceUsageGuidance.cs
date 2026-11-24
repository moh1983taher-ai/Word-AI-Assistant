using System.Text;
using System.Text.Json;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Host-authored workflow hints for known source adapters. Hints are emitted only
/// when the corresponding tools and argument properties are present in the live catalog.
/// </summary>
internal static class ResearchSourceUsageGuidance
{
    public static string Build(IReadOnlyList<ToolDescriptor> tools)
    {
        if (tools.Count == 0)
            return string.Empty;

        var builder = new StringBuilder();
        builder.AppendLine("إرشادات تشغيلية خاصة بالمصادر المتاحة:");
        builder.AppendLine("استعمل الأوصاف ومخططات المعاملات الحية بوصفها المرجع الحاسم؛ لا تفترض توفر أداة أو معامل لم يظهر في الفهرس.");

        var shamelaTools = tools
            .Where(tool => string.Equals(tool.Server, "shamela", StringComparison.Ordinal))
            .ToArray();

        if (shamelaTools.Length > 0)
            AppendShamelaGuidance(builder, shamelaTools);

        return builder.ToString().TrimEnd();
    }

    private static void AppendShamelaGuidance(
        StringBuilder builder,
        IReadOnlyList<ToolDescriptor> tools)
    {
        var find = FindTool(tools, "shamela_find");
        var scoped = FindTool(tools, "shamela_find_scoped");
        var openMany = FindTool(tools, "shamela_open_many");
        var open = FindTool(tools, "shamela_open");

        builder.AppendLine()
            .AppendLine("المكتبة الشاملة:");

        if (find is not null)
        {
            builder.AppendLine(
                "ابدأ البحث العام عبر shamela_find قبل الانتقال إلى البحث المحصور. صغ عبارات عربية قصيرة أقرب إلى ألفاظ المسألة في الكتب، وجرّب أكثر من صيغة عند الحاجة بدلاً من الاكتفاء بسؤال المستخدم الكامل أو لفظ مفرد.");

            if (HasRootProperty(find, "keywords"))
                builder.AppendLine("مرّر عبارات البحث في المعامل keywords كما يحدده مخطط الأداة؛ لا تستبدله باسم query.");

            if (HasRootProperty(find, "book_queries") || HasRootProperty(find, "author_queries"))
                builder.AppendLine("لا تستخدم book_queries أو author_queries إلا إذا عيّن المستخدم كتاباً أو مؤلفاً بعينه؛ لا تقيّد البحث من تلقاء نفسك.");

            if (scoped is not null)
                builder.AppendLine("استخدم shamela_find_scoped كخطوة متابعة بعد البحث العام عندما تقتضي النتائج تضييق النطاق؛ لا تبدأ به.");

            if (openMany is not null || open is not null)
            {
                var readTool = openMany is not null ? "shamela_open_many" : "shamela_open";
                builder.Append("بعد تحديد الموضع، افتح النص الكامل عبر ")
                    .Append(readTool)
                    .AppendLine(" متى كان ذلك متاحاً، ولا تعتمد في نقل الأقوال على مقتطف البحث وحده.");
            }

            if (openMany is not null && HasRootProperty(openMany, "context"))
                builder.AppendLine("اطلب سياقاً إضافياً عبر معامل context عند انقطاع المقطع أو نقص سياقه، وفق قيم المخطط، ولا ترسل معاملات غير معلنة.");

            builder.AppendLine(
                "انقل النص من الصفحات التي أعادتها أدوات القراءة، وانسب كل قول إلى قائله، واستخدم بيانات الإحالة والرابط التي يعيدها الخادم دون إنشاء بيانات صفحات أو روابط من عندك.");
        }
    }

    private static ToolDescriptor? FindTool(
        IEnumerable<ToolDescriptor> tools,
        string name) =>
        tools.FirstOrDefault(tool =>
            string.Equals(tool.Name, name, StringComparison.Ordinal));

    private static bool HasRootProperty(ToolDescriptor tool, string propertyName)
    {
        if (tool.InputSchema.ValueKind != JsonValueKind.Object ||
            !TryGetProperty(tool.InputSchema, "properties", out var properties) ||
            properties.ValueKind != JsonValueKind.Object)
            return false;

        return TryGetProperty(properties, propertyName, out _);
    }

    private static bool TryGetProperty(
        JsonElement element,
        string name,
        out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.Ordinal))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }
}
