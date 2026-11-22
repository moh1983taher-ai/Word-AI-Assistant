using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchTaskProgressTests
{
    [Fact]
    public void PreventsTheSameSearchFromBeingExecutedAgainEvenIfPaginationChanges()
    {
        using var schema = JsonDocument.Parse("""{"type":"object"}""");
        var tool = new ToolDescriptor(
            "aljam3", "الجامع", "search", "بحث", schema.RootElement.Clone());
        using var first = JsonDocument.Parse("""{"q":"الأمر الوارد بعد الحظر","per_page":10}""");
        using var second = JsonDocument.Parse("""{"per_page":50,"q":"الأمر الوارد بعد الحظر"}""");
        var progress = new ResearchTaskProgress(["S1"]);

        Assert.True(progress.TryRegisterToolCall("S1", tool, first.RootElement, out _));
        Assert.False(progress.TryRegisterToolCall("S1", tool, second.RootElement, out var reason));
        Assert.Contains("يكرر", reason);
        Assert.Equal(1, progress.DuplicateToolCalls);
    }

    [Fact]
    public void AllowsTheSameSearchForDifferentExplicitScopeItems()
    {
        using var schema = JsonDocument.Parse("""{"type":"object"}""");
        var tool = new ToolDescriptor(
            "shamela", "المكتبة الشاملة", "shamela_find", "بحث", schema.RootElement.Clone());
        using var args = JsonDocument.Parse("""{"keywords":["الاجتهاد"]}""");
        var progress = new ResearchTaskProgress(["S1", "S2"]);

        Assert.True(progress.TryRegisterToolCall("S1", tool, args.RootElement, out _));
        Assert.True(progress.TryRegisterToolCall("S2", tool, args.RootElement, out _));
    }

    [Fact]
    public void RecordsCoveragePerScopeItemAndResetsNoProgressOnNewEvidence()
    {
        using var schema = JsonDocument.Parse("""{"type":"object"}""");
        var tool = new ToolDescriptor(
            "shamela", "المكتبة الشاملة", "shamela_find", "بحث", schema.RootElement.Clone());
        var progress = new ResearchTaskProgress(["S1", "S2"]);

        Assert.False(progress.RecordToolResult("S1", tool, "لا توجد نتائج", false, false));
        Assert.Equal(1, progress.ConsecutiveNoProgress);

        Assert.True(progress.RecordToolResult("S1", tool, "نص الدليل الأول", true, true));
        Assert.Contains("S1", progress.CoveredScopeItemIds);
        Assert.Equal(0, progress.ConsecutiveNoProgress);
        Assert.Contains("S2", progress.MissingScopeItemIds(["S1", "S2"]));
    }

    [Fact]
    public void CountsThreeRepeatedNonProgressResultsAsAStopCondition()
    {
        using var schema = JsonDocument.Parse("""{"type":"object"}""");
        var tool = new ToolDescriptor(
            "aljam3", "الجامع", "search", "بحث", schema.RootElement.Clone());
        var progress = new ResearchTaskProgress(["S1"]);

        progress.RecordToolResult("S1", tool, "لا توجد نتيجة", false, false);
        progress.RecordToolResult("S1", tool, "لا توجد نتيجة", false, false);
        progress.RecordToolResult("S1", tool, "لا توجد نتيجة", false, false);

        Assert.Equal(3, progress.ConsecutiveNoProgress);
    }
}
