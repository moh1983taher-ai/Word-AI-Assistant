using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchAgentInstructionsTests
{
    [Fact]
    public void UsesBriefGuidanceThatLeavesResearchChoicesToTheModel()
    {
        var instructions = ResearchAgentInstructions.Build(
            Array.Empty<string>(),
            Array.Empty<string>());

        Assert.Contains("افهم السؤال كما قصده صاحبه", instructions);
        Assert.Contains("اختر بنفسك أدوات MCP", instructions);
        Assert.Contains("لا تضف أهدافاً أو قوالب", instructions);
        Assert.Contains("لا تختلق معلومة", instructions);
        Assert.Contains("بالشكل واللغة والحجم الذي طلبه الباحث", instructions);
        Assert.True(instructions.Length < 1_000,
            $"Standing instructions unexpectedly grew to {instructions.Length:N0} characters.");
    }

    [Fact]
    public void DoesNotHardcodePlansOrCoverageRequirements()
    {
        var instructions = ResearchAgentInstructions.Build(
            ["الجامع", "المكتبة الشاملة"],
            Array.Empty<string>());

        Assert.Contains("الجامع، المكتبة الشاملة", instructions);
        Assert.DoesNotContain("خطة البحث", instructions);
        Assert.DoesNotContain("محاور التغطية", instructions);
        Assert.DoesNotContain("لا تكتف بصفحة واحدة", instructions);
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
