using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchAgentInstructionsTests
{
    [Fact]
    public void IncludesConciseOperationalRulesAfterPlanIsAuthored()
    {
        var instructions = ResearchAgentInstructions.Build(
            Array.Empty<string>(),
            Array.Empty<string>());

        Assert.Contains("افهم مقصد الباحث أولاً", instructions);
        Assert.Contains("لا تتوقف عند أول نتيجة", instructions);
        Assert.Contains("فتح الصفحة يثبت الوصول", instructions);
        Assert.Contains("ولا يثبت صلتها", instructions);
        Assert.Contains("لا تكتف بصفحة واحدة", instructions);
        Assert.True(instructions.Length < 1_800,
            $"Operational instructions unexpectedly grew to {instructions.Length:N0} characters.");
    }

    [Fact]
    public void UsesShortOperationalInstructionsFocusedOnUnderstandingAndCoverage()
    {
        var instructions = ResearchAgentInstructions.Build(
            ["الجامع", "المكتبة الشاملة"],
            Array.Empty<string>());

        Assert.Contains("اتبع خطة البحث", instructions);
        Assert.Contains("لا تتوقف عند أول نتيجة", instructions);
        Assert.Contains("ولا يثبت صلتها", instructions);
        Assert.DoesNotContain("التزم بمنهجية المستخدم وهيكله", instructions);
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
