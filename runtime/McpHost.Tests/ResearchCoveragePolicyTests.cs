using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchCoveragePolicyTests
{
    [Fact]
    public void AcceptsFinalOnlyWhenEveryPlannedFacetIsDeclaredCovered()
    {
        var missing = ResearchCoveragePolicy.GetMissingCoverage(
            ["التعريف", "الصيغ", "النسبة"],
            [" التعريف ", "الصيغ", "النسبة"],
            []);

        Assert.Empty(missing);
    }

    [Fact]
    public void ReturnsEveryUncoveredOrUnattestedFacet()
    {
        var missing = ResearchCoveragePolicy.GetMissingCoverage(
            ["التعريف", "الصيغ", "النسبة"],
            ["التعريف"],
            ["النسبة"]);

        Assert.Equal(new[] { "الصيغ", "النسبة" }, missing);
    }

    [Fact]
    public void NormalizesWhitespaceWhenComparingPlanCoverage()
    {
        var missing = ResearchCoveragePolicy.GetMissingCoverage(
            ["أثر المصطلح في الاستنباط"],
            ["أثر   المصطلح في الاستنباط"],
            []);

        Assert.Empty(missing);
    }
}
