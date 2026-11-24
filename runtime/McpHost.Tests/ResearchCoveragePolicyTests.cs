using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchCoveragePolicyTests
{
    [Fact]
    public void AcceptsFinalWhenEveryPlannedFacetIsCovered()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["التعريف", "الصيغ", "النسبة"],
            [" التعريف ", "الصيغ", "النسبة"],
            []);

        Assert.Empty(unreported);
    }

    [Fact]
    public void AcceptsExplicitlyUnresolvedFacetsForTransparentPartialAnswers()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["التعريف", "الصيغ", "النسبة"],
            ["التعريف"],
            ["الصيغ", "النسبة"]);

        Assert.Empty(unreported);
    }

    [Fact]
    public void RejectsFacetsOmittedFromBothCoverageLists()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["التعريف", "الصيغ", "النسبة"],
            ["التعريف"],
            ["النسبة"]);

        Assert.Equal(new[] { "الصيغ" }, unreported);
    }

    [Fact]
    public void RejectsAPlanFacetDeclaredBothCoveredAndUncovered()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["التعريف", "الصيغ"],
            ["التعريف", "الصيغ"],
            ["الصيغ"]);

        Assert.Equal(new[] { "الصيغ" }, unreported);
    }

    [Fact]
    public void NormalizesWhitespaceWhenComparingPlanCoverage()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["أثر المصطلح في الاستنباط"],
            ["أثر   المصطلح في الاستنباط"],
            []);

        Assert.Empty(unreported);
    }
}
