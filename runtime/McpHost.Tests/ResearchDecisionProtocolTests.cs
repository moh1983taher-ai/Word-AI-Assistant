using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchDecisionProtocolTests
{
    [Fact]
    public void DescribesExactJsonShapesWithoutDependingOnResponseFormat()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: true,
            pendingSourceLabels: []);

        Assert.Contains("\"type\":\"tool_call\"", prompt, StringComparison.Ordinal);
        Assert.Contains("\"type\":\"final\"", prompt, StringComparison.Ordinal);
        Assert.Contains("inputSchema", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("response_format", prompt, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("json_schema", prompt, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PreventsFinalWhileSelectedSourcesRemainPending()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: false,
            pendingSourceLabels: ["المكتبة الشاملة"]);

        Assert.Contains("لا يُسمح بالإجابة النهائية الآن", prompt, StringComparison.Ordinal);
        Assert.Contains("المكتبة الشاملة", prompt, StringComparison.Ordinal);
        Assert.Contains("\"type\":\"tool_call\"", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void PreventsFinalWhenSearchIsCompleteButAResultPageIsNotVerified()
    {
        Assert.False(ResearchFinalizationPolicy.CanFinalize(
            pendingSearchSourceCount: 0,
            unverifiedPageSourceCount: 1));
    }

    [Fact]
    public void GivesPageVerificationInstructionsWhenNoSearchSourceRemainsPending()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: false,
            pendingSourceLabels: [],
            pendingPageSourceLabels: ["المكتبة الشاملة"]);

        Assert.Contains("لا يُسمح بالإجابة النهائية الآن", prompt, StringComparison.Ordinal);
        Assert.Contains("المكتبة الشاملة", prompt, StringComparison.Ordinal);
        Assert.Contains("افتح صفحة", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("لم يكتمل البحث في:", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void TimeoutFinalizationOnlyPermitsFinalJson()
    {
        var prompt = ResearchDecisionProtocol.BuildTimeoutFinalInstructions();

        Assert.Contains("\"type\":\"final\"", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("\"type\":\"tool_call\"", prompt, StringComparison.Ordinal);
    }
}
