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

    [Fact]
    public void AcceptsStableCoverageIdentifiersWithoutRepeatingArabicLabels()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["الموضع الأصلي", "نص القاعدة", "بيانات المصدر"],
            ["C1", "C3"],
            ["C2"]);
        Assert.Empty(unreported);
    }

    [Fact]
    public void RejectsUnknownCoverageIdentifiersAndPreservesMissingOriginalLabels()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["الموضع الأصلي", "نص القاعدة"],
            ["C1", "C9"],
            []);
        Assert.Equal(new[] { "نص القاعدة", "C9" }, unreported);
    }

    [Fact]
    public void RejectsStableIdentifierDeclaredBothCoveredAndUncovered()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["الموضع الأصلي"],
            ["C1"],
            ["C1"]);
        Assert.Equal(new[] { "الموضع الأصلي" }, unreported);
    }

    [Fact]
    public void RejectsParaphrasedCoverageLabelsInsteadOfSilentlyAssumingEquivalence()
    {
        var unreported = ResearchCoveragePolicy.GetUnreportedCoverage(
            ["موضع أصلي من الجامع"],
            ["صفحة من كتاب في الجامع"],
            []);
        Assert.Equal(
            new[] { "موضع أصلي من الجامع", "صفحة من كتاب في الجامع" },
            unreported);
    }
}
