using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchDecisionProtocolTests
{
    [Fact]
    public void KeepsOnlyTheMinimalToolTransportInstructions()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: true,
            pendingSourceLabels: []);

        Assert.Contains("<tool_call>", prompt, StringComparison.Ordinal);
        Assert.Contains("server__tool_name", prompt, StringComparison.Ordinal);
        Assert.Contains("اكتب الإجابة مباشرة", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("covered", prompt, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("uncovered", prompt, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("C1", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("خطة بحث", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void PreventsFinalWhileSelectedSourcesRemainPending()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: false,
            pendingSourceLabels: ["المكتبة الشاملة"]);

        Assert.Contains("لا تنه المهمة الآن", prompt, StringComparison.Ordinal);
        Assert.Contains("المكتبة الشاملة", prompt, StringComparison.Ordinal);
        Assert.Contains("<tool_call>", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void GivesPageVerificationInstructionsWhenSearchSourcesAreComplete()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: false,
            pendingSourceLabels: [],
            pendingPageSourceLabels: ["المكتبة الشاملة"]);

        Assert.Contains("لا تنه المهمة الآن", prompt, StringComparison.Ordinal);
        Assert.Contains("المكتبة الشاملة", prompt, StringComparison.Ordinal);
        Assert.Contains("افتح صفحة", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("نفّذ بحثاً في المصادر", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void AllowsTheModelToContinueResearchAfterMinimumAccessChecks()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: true,
            pendingSourceLabels: []);

        Assert.Contains("أجب الآن إن كانت الأدلة كافية", prompt, StringComparison.Ordinal);
        Assert.Contains("واصل البحث", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidDecisionRepairDoesNotOfferFinalWhilePageReadsRemainPending()
    {
        var prompt = ResearchDecisionProtocol.BuildInvalidDecisionRepair(
            allowFinal: false,
            pendingSourceLabels: [],
            pendingPageSourceLabels: ["المكتبة الشاملة"]);

        Assert.Contains("<tool_call>", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("\"type\":\"final\"", prompt, StringComparison.Ordinal);
        Assert.Contains("المكتبة الشاملة", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void TimeoutInstructionsAllowOnlyEvidenceBoundNaturalLanguage()
    {
        var prompt = ResearchDecisionProtocol.BuildTimeoutFinalInstructions();

        Assert.Contains("أجب الآن", prompt, StringComparison.Ordinal);
        Assert.Contains("الصفحات التي فُتحت فعلاً", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("<tool_call>", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("JSON", prompt, StringComparison.Ordinal);
    }
}
