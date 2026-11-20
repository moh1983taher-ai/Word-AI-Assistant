using System.Text.RegularExpressions;
using ModelContextProtocol.Protocol;

namespace ResearchTools.McpHost.Security;

/// <summary>
/// Conservative allow-list for read-only research tools. Server annotations are treated as hints,
/// never as the sole security boundary.
/// </summary>
public static partial class ToolPolicy
{
    private static readonly Regex DeniedName = new(
        "(delete|remove|update|edit|write|create|upload|submit|send|publish|purchase|execute|admin|حذف|تعديل|كتابة|إنشاء|رفع|إرسال|نشر|شراء)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex ReadName = new(
        "(search|find|query|lookup|get|fetch|read|open|list|retrieve|browse|catalog|book|page|text|بحث|استعلام|قراءة|جلب|استرجاع|عرض|فهرس)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static bool IsAllowed(Tool tool)
    {
        ArgumentNullException.ThrowIfNull(tool);

        var name = tool.Name?.Trim();
        if (string.IsNullOrEmpty(name) || DeniedName.IsMatch(name))
            return false;

        var annotations = tool.Annotations;
        if (annotations?.ReadOnlyHint == false || annotations?.DestructiveHint == true)
            return false;

        return ReadName.IsMatch(name);
    }
}
