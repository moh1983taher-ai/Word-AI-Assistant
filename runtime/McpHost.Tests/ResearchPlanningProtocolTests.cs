using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchPlanningProtocolTests
{
    [Fact]
    public void AsksModelToUnderstandAndPlanBeforeSearching()
    {
        var prompt = ResearchPlanningProtocol.BuildInitialInstructions(
            ["الجامع", "المكتبة الشاملة"]);

        Assert.Contains("مرحلة فهم السؤال والتخطيط فقط", prompt);
        Assert.Contains("السؤال العام", prompt);
        Assert.Contains("طلب نص بعينه", prompt);
        Assert.Contains("الجامع، المكتبة الشاملة", prompt);
        Assert.Contains("\"type\":\"plan\"", prompt);
        Assert.DoesNotContain("\"type\":\"tool_call\"", prompt);
    }

    [Fact]
    public void CarriesTheModelsPlanIntoEveryResearchTurn()
    {
        var plan = new ResearchPlan(
            "استقصاء صيغ قاعدة",
            "multiple_forms",
            [
                new ResearchQuery("بناء القوي على الضعيف", "الصيغة الحرفية"),
                new ResearchQuery("لا يجوز بناء القوي على الضعيف", "صيغة الحكم")
            ],
            ["الصيغ", "نسبة الأقوال"]);

        var context = ResearchPlanningProtocol.FormatPlan(plan);

        Assert.Contains("بناء القوي على الضعيف", context);
        Assert.Contains("صيغة الحكم", context);
        Assert.Contains("نسبة الأقوال", context);
    }
}
