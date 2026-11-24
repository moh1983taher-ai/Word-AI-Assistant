using System.Text.Json;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class ResearchEvidenceLedgerTests
{
    [Fact]
    public void MarksSearchExcerptsAsDiscoveryOnlyUntilThePageIsOpened()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var search = new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_find", "", schema.RootElement.Clone());
        var open = new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_open", "", schema.RootElement.Clone());

        var ledger = new ResearchEvidenceLedger();
        ledger.Add(search, "مقتطف نتيجة البحث", succeeded: true);
        ledger.Add(open, "النص الكامل للصفحة", succeeded: true);

        var context = ledger.BuildContext(
            ["shamela"],
            Array.Empty<string>(),
            maxCharacters: 2_000);

        Assert.Contains("نتائج البحث الأولية مخصصة لاكتشاف المراجع", context);
        Assert.Contains("مقتطف نتيجة البحث", context);
        Assert.Contains("النص الكامل للصفحة", context);
    }

    [Fact]
    public void RebuildsBoundedContextWithEvidenceFromEverySelectedSource()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var aljam3 = new ToolDescriptor("aljam3", "الجامع", "search", "", schema.RootElement.Clone());
        var shamela = new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_find", "", schema.RootElement.Clone());

        var ledger = new ResearchEvidenceLedger();
        var aljam3Result = "ALJAM3_EVIDENCE " + new string('A', 3_000);
        var shamelaResult = "SHAMELA_EVIDENCE " + new string('S', 3_000);

        ledger.Add(aljam3, aljam3Result, succeeded: true);
        ledger.Add(shamela, shamelaResult, succeeded: true);

        var context = ledger.BuildContext(
            ["aljam3", "shamela"],
            Array.Empty<string>(),
            maxCharacters: 1_000);

        Assert.InRange(context.Length, 1, 1_000);
        Assert.Contains("الجامع", context);
        Assert.Contains("المكتبة الشاملة", context);
        Assert.Contains("ALJAM3_EVIDENCE", context);
        Assert.Contains("SHAMELA_EVIDENCE", context);
    }

    [Fact]
    public void PreservesExplicitNoMatchResponsesForTheFinalAnswer()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor(
            "aljam3", "الجامع", "search", "", schema.RootElement.Clone());

        var ledger = new ResearchEvidenceLedger();
        ledger.Add(descriptor, "لا توجد نتائج مطابقة لعبارة البحث في هذه المكتبة.", succeeded: true);

        var context = ledger.BuildContext(
            ["aljam3"],
            Array.Empty<string>(),
            maxCharacters: 2_000);

        Assert.Contains("لا توجد نتائج مطابقة", context);
        Assert.True(Assert.Single(ledger.Entries).Succeeded);
    }

    [Fact]
    public void KeepsFullToolResultsInMemoryWhenModelContextUsesExcerpts()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor(
            "shamela", "المكتبة الشاملة", "shamela_find", "", schema.RootElement.Clone());
        var fullResult = new string('ن', 20_000);

        var ledger = new ResearchEvidenceLedger();
        ledger.Add(descriptor, fullResult, succeeded: true);

        var context = ledger.BuildContext(
            ["shamela"],
            Array.Empty<string>(),
            maxCharacters: 2_000);

        Assert.True(context.Length <= 2_000);
        Assert.Equal(fullResult, Assert.Single(ledger.Entries).Content);
        Assert.Contains("مقتطف", context);
    }

    [Fact]
    public void KeepsLargeToolOutputInTheTaskLedgerAndSendsOnlyABoundedExcerpt()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor(
            "shamela", "المكتبة الشاملة", "shamela_open_many", "", schema.RootElement.Clone());
        var fullResult = "بداية النتيجة " + new string('ن', 24_000) + " نهاية النتيجة";

        var ledger = new ResearchEvidenceLedger();
        ledger.Add(descriptor, fullResult, succeeded: true);

        var context = ledger.BuildContext(
            ["shamela"],
            Array.Empty<string>(),
            maxCharacters: 10_000);

        Assert.InRange(context.Length, 1, 10_000);
        Assert.DoesNotContain(fullResult, context);
        Assert.Contains("مقتطف", context);
        Assert.Equal(fullResult, Assert.Single(ledger.Entries).Content);
    }

    [Fact]
    public void IdentifiesPendingSourcesInTheRebuiltEvidenceContext()
    {
        using var schema = JsonDocument.Parse("""{"type":"object","properties":{},"required":[]}""");
        var descriptor = new ToolDescriptor("shamela", "المكتبة الشاملة", "shamela_find", "", schema.RootElement.Clone());

        var ledger = new ResearchEvidenceLedger();
        ledger.Add(descriptor, "نتيجة موثقة", succeeded: true);

        var context = ledger.BuildContext(
            ["shamela", "aljam3"],
            ["aljam3"],
            maxCharacters: 2_000);

        Assert.Contains("الجامع", context);
        Assert.Contains("المكتبة الشاملة", context);
    }
}
