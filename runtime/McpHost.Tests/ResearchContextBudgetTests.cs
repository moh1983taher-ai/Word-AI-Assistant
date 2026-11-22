using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchContextBudgetTests
{
    [Fact]
    public void DefaultsProvideMoreRoomThanTheFormerFixedEvidenceAndHistoryBudgets()
    {
        ClearSettings();
        try
        {
            var budget = ResearchContextBudget.Load();

            Assert.Equal(64_000, budget.MaxContextCharacters);
            Assert.Equal(8_192, budget.MaxOutputTokens);
            Assert.True(budget.GetAvailableInputCharacters(5_000) > 10_000);
        }
        finally
        {
            ClearSettings();
        }
    }

    [Fact]
    public void EnvironmentCanRaiseBudgetsWithinTheSupportedRange()
    {
        try
        {
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", "120000");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", "12000");

            var budget = ResearchContextBudget.Load();

            Assert.Equal(120_000, budget.MaxContextCharacters);
            Assert.Equal(12_000, budget.MaxOutputTokens);
        }
        finally
        {
            ClearSettings();
        }
    }

    [Fact]
    public void EvidenceUsesOnlyTheSpaceItNeedsAndNeverConsumesTheWholePromptBudget()
    {
        var budget = new ResearchContextBudget(64_000, 8_192);

        Assert.Equal(4_000, budget.GetEvidenceBudget(30_000, 4_000));
        Assert.Equal(20_000, budget.GetEvidenceBudget(30_000, 40_000));
        Assert.Equal(0, budget.GetEvidenceBudget(30_000, 0));
    }

    private static void ClearSettings()
    {
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", null);
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", null);
    }
}
