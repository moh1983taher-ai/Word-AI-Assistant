using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchContextBudgetTests
{
    [Fact]
    public void DefaultsProvideExpandedContextOutputAndExecutionTime()
    {
        ClearSettings();
        try
        {
            var budget = ResearchContextBudget.Load();

            Assert.Equal(64_000, budget.MaxContextCharacters);
            Assert.Equal(8_192, budget.MaxOutputTokens);
            Assert.Equal(TimeSpan.FromMinutes(10), budget.ResearchTimeout);
            Assert.Equal(TimeSpan.FromSeconds(120), budget.ModelRequestTimeout);
            Assert.Equal(TimeSpan.FromSeconds(60), budget.FinalizationTimeout);
            Assert.True(budget.GetAvailableInputCharacters(5_000) > 10_000);
        }
        finally
        {
            ClearSettings();
        }
    }

    [Fact]
    public void EnvironmentCanRaiseContextOutputAndTimeBudgets()
    {
        try
        {
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", "120000");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", "12000");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_RESEARCH_TIMEOUT_SECONDS", "900");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_MODEL_REQUEST_TIMEOUT_SECONDS", "180");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_FINALIZATION_TIMEOUT_SECONDS", "90");

            var budget = ResearchContextBudget.Load();

            Assert.Equal(120_000, budget.MaxContextCharacters);
            Assert.Equal(12_000, budget.MaxOutputTokens);
            Assert.Equal(TimeSpan.FromMinutes(15), budget.ResearchTimeout);
            Assert.Equal(TimeSpan.FromMinutes(3), budget.ModelRequestTimeout);
            Assert.Equal(TimeSpan.FromSeconds(90), budget.FinalizationTimeout);
        }
        finally
        {
            ClearSettings();
        }
    }

    [Fact]
    public void EnvironmentValuesAreClampedToSupportedOperationalBounds()
    {
        try
        {
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", "99999999");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", "99999999");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_RESEARCH_TIMEOUT_SECONDS", "99999999");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_MODEL_REQUEST_TIMEOUT_SECONDS", "1");
            Environment.SetEnvironmentVariable("RESEARCHTOOLS_FINALIZATION_TIMEOUT_SECONDS", "1");

            var budget = ResearchContextBudget.Load();

            Assert.Equal(196_000, budget.MaxContextCharacters);
            Assert.Equal(32_768, budget.MaxOutputTokens);
            Assert.Equal(TimeSpan.FromMinutes(30), budget.ResearchTimeout);
            Assert.Equal(TimeSpan.FromSeconds(30), budget.ModelRequestTimeout);
            Assert.Equal(TimeSpan.FromSeconds(20), budget.FinalizationTimeout);
        }
        finally
        {
            ClearSettings();
        }
    }

    [Fact]
    public void EvidenceUsesOnlyTheSpaceItNeedsAndNeverConsumesTheWholePromptBudget()
    {
        var budget = new ResearchContextBudget(
            64_000,
            8_192,
            TimeSpan.FromMinutes(10),
            TimeSpan.FromSeconds(120),
            TimeSpan.FromSeconds(60));

        Assert.Equal(4_000, budget.GetEvidenceBudget(30_000, 4_000));
        Assert.Equal(20_000, budget.GetEvidenceBudget(30_000, 40_000));
        Assert.Equal(0, budget.GetEvidenceBudget(30_000, 0));
    }

    private static void ClearSettings()
    {
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS", null);
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS", null);
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_RESEARCH_TIMEOUT_SECONDS", null);
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_MODEL_REQUEST_TIMEOUT_SECONDS", null);
        Environment.SetEnvironmentVariable("RESEARCHTOOLS_FINALIZATION_TIMEOUT_SECONDS", null);
    }
}
