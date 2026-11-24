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
        Assert.Contains("\\"arguments\\":{...}", prompt, StringComparison.Ordinal);
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

    [Theory]
    [InlineData(0, 0, true)]
    [InlineData(1, 0, false)]
    [InlineData(0, 1, false)]
    [InlineData(2, 3, false)]
    public void AllowsFinalOnlyWhenSearchAndPageVerificationAreComplete(
        int pendingSearchSourceCount,
        int unverifiedPageSourceCount,
        bool expected)
    {
        Assert.Equal(
            expected,
            ResearchFinalizationPolicy.CanFinalize(
                pendingSearchSourceCount,
                unverifiedPageSourceCount));
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
    public void DoesNotTreatOneOpenedPageAsEvidenceThatResearchIsComplete()
    {
        var prompt = ResearchDecisionProtocol.BuildTurnInstructions(
            allowFinal: true,
            pendingSourceLabels: []);

        Assert.Contains("لا يعني ذلك اكتمال الاستقصاء", prompt, StringComparison.Ordinal);
        Assert.Contains("لا تكتف بصفحة واحدة", prompt, StringComparison.Ordinal);
        Assert.Contains("تغطية محاور خطة البحث", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void TimeoutFinalizationOnlyPermitsFinalJson()
    {
        var prompt = ResearchDecisionProtocol.BuildTimeoutFinalInstructions();

        Assert.Contains("\"type\":\"final\"", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("\"type\":\"tool_call\"", prompt, StringComparison.Ordinal);
    }
}
