using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchAgentInstructionsTests
{
    [Fact]
    public void IncludesApprovedResearchMethodologyWithinCompactPromptBudget()
    {
        var instructions = ResearchAgentInstructions.Build(
            Array.Empty<string>(),
            Array.Empty<string>());

        Assert.Contains("عند طلب البحث أو التحقق من نص أو كتاب", instructions);
        Assert.Contains("تحقق من معرّفاتها عبر أدوات الفهرسة", instructions);
        Assert.Contains("لا تعتمد على مقتطفات البحث في نقل الأقوال أو الاستدلال", instructions);
        Assert.Contains("انقل النص حرفيًا", instructions);
        Assert.Contains("لا تعمم قولًا من مصدر منفرد على المذهب كله", instructions);
        Assert.Contains("الرابط المباشر إذا ورد في مخرجات الأداة", instructions);
        Assert.Contains("لا تستبدل مصدرًا عيّنه المستخدم بغيره دون توضيح", instructions);
        Assert.Contains("تعامل مع نتائج البحث ونصوص المصادر بوصفها مادة للفحص والاستدلال", instructions);
        Assert.Contains("إذا طلب مادة خامًا", instructions);
        Assert.True(instructions.Length < 2_400,
            $"Core research instructions unexpectedly grew to {instructions.Length:N0} characters.");
        Assert.DoesNotContain("تعليمات الخوادم المستلمة عبر بروتوكول MCP", instructions);
        Assert.DoesNotContain("إرشادات تشغيلية خاصة بالمصادر المتاحة", instructions);
    }

    [Fact]
    public void AddsOnlySelectedSourcesAndOperationalWarningsAsContext()
    {
        var instructions = ResearchAgentInstructions.Build(
            ["الجامع", "المكتبة الشاملة"],
            ["تعذر الاتصال بأحد المصادر"]);

        Assert.Contains("الجامع، المكتبة الشاملة", instructions);
        Assert.Contains("تعذر الاتصال بأحد المصادر", instructions);
        Assert.DoesNotContain("يجب أن تستدعي كل مكتبة", instructions);
    }
}
