using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchAgentInstructionsTests
{
    [Fact]
    public void SharedInstructionsLeaveToolUseAndResearchDecisionsToTheAgent()
    {
        var instructions = ResearchAgentInstructions.Build(
            ["المكتبة الشاملة"],
            Array.Empty<string>());

        Assert.Contains("أدر استخدام الأدوات بنفسك", instructions);
        Assert.Contains("المكتبات المتاحة لهذه المهمة وفق اختيار المستخدم", instructions);
        Assert.Contains("ابدأ باستدعاء أداة بحث مناسبة", instructions);
        Assert.Contains("قدّر بنفسك هل تكفي للإجابة", instructions);
        Assert.DoesNotContain("عدد محدد", instructions);
        Assert.DoesNotContain("استقصائياً", instructions);
        Assert.DoesNotContain("لا يلزمك استخدام كل مكتبة", instructions);
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
