using ResearchTools.McpHost.Models;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchSourceSelectionTests
{
    [Fact]
    public void NormalizeKeepsOnlyTheSelectedSourcesInCanonicalOrder()
    {
        var selected = ResearchSourceSelection.Normalize(["shamela", "aljam3", "shamela"]);

        Assert.Equal(["aljam3", "shamela"], selected);
    }

    [Fact]
    public void NormalizeAllowsSelectingEitherSingleSource()
    {
        Assert.Equal(["aljam3"], ResearchSourceSelection.Normalize(["aljam3"]));
        Assert.Equal(["shamela"], ResearchSourceSelection.Normalize(["shamela"]));
    }

    [Fact]
    public void NormalizeRejectsUnknownSourceIds()
    {
        Assert.Throws<ArgumentException>(() =>
            ResearchSourceSelection.Normalize(["aljam3", "other-library"]));
    }

    [Fact]
    public void MissingReturnsEverySelectedSourceWithoutASuccessfulToolCall()
    {
        var missing = ResearchSourceSelection.Missing(
            ["aljam3", "shamela"],
            ["aljam3"]);

        Assert.Equal(["shamela"], missing);
    }
}
