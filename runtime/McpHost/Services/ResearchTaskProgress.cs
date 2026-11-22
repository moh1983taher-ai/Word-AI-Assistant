using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Enforces task binding and detects repeated calls/results without imposing
/// a small fixed cycle count on useful research.
/// </summary>
internal sealed class ResearchTaskProgress
{
    private readonly HashSet<string> _scopeItemIds;
    private readonly HashSet<string> _coveredScopeItems = new(StringComparer.Ordinal);
    private readonly HashSet<string> _toolCallSignatures = new(StringComparer.Ordinal);
    private readonly HashSet<string> _resultSignatures = new(StringComparer.Ordinal);

    internal ResearchTaskProgress(IEnumerable<string> scopeItemIds)
    {
        ArgumentNullException.ThrowIfNull(scopeItemIds);
        _scopeItemIds = scopeItemIds.ToHashSet(StringComparer.Ordinal);
        if (_scopeItemIds.Count == 0)
            throw new ArgumentException("At least one research scope item is required.", nameof(scopeItemIds));
    }

    internal IReadOnlySet<string> CoveredScopeItemIds => _coveredScopeItems;
    internal int ConsecutiveNoProgress { get; private set; }
    internal int DuplicateToolCalls { get; private set; }
    internal int NewEvidenceCount => _resultSignatures.Count;

    internal IReadOnlyList<string> MissingScopeItemIds(IEnumerable<string> orderedScopeIds) =>
        orderedScopeIds.Where(id => !_coveredScopeItems.Contains(id)).ToArray();

    internal bool TryRegisterToolCall(
        string scopeItemId,
        ToolDescriptor descriptor,
        JsonElement arguments,
        out string reason)
    {
        ArgumentNullException.ThrowIfNull(descriptor);
        reason = string.Empty;

        if (!_scopeItemIds.Contains(scopeItemId))
        {
            reason = "قرار الأداة غير مرتبط بعنصر صالح من عناصر نطاق المهمة.";
            ConsecutiveNoProgress++;
            return false;
        }

        var canonicalArguments = Canonicalize(arguments);
        var signature = Hash(
            scopeItemId + "|" + descriptor.Server + "|" + descriptor.Name + "|" + canonicalArguments);

        if (_toolCallSignatures.Add(signature))
            return true;

        DuplicateToolCalls++;
        ConsecutiveNoProgress++;
        reason = "رُفض الاستدعاء لأنه يكرر أداة وعبارة بحث أو معاملات سبق تنفيذها في عنصر النطاق نفسه.";
        return false;
    }

    internal bool RecordToolResult(
        string scopeItemId,
        ToolDescriptor descriptor,
        string content,
        bool succeeded,
        bool newlyCoveredSource)
    {
        if (!succeeded)
        {
            ConsecutiveNoProgress++;
            return false;
        }

        var newScopeEvidence = _coveredScopeItems.Add(scopeItemId);
        var signature = Hash(
            descriptor.Server + "|" + descriptor.Name + "|" + NormalizeText(content ?? string.Empty));
        var newContent = _resultSignatures.Add(signature);
        var madeProgress = newScopeEvidence || newContent || newlyCoveredSource;

        if (madeProgress)
            ConsecutiveNoProgress = 0;
        else
            ConsecutiveNoProgress++;

        return madeProgress;
    }

    internal void RecordRejectedDecision() => ConsecutiveNoProgress++;

    private static string Canonicalize(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.Object => "{" + string.Join(",", element.EnumerateObject()
                .OrderBy(property => property.Name, StringComparer.Ordinal)
                .Select(property => JsonSerializer.Serialize(property.Name) + ":" +
                                    Canonicalize(property.Value))) + "}",
            JsonValueKind.Array => "[" + string.Join(",", element.EnumerateArray()
                .Select(item => Canonicalize(item))) + "]",
            JsonValueKind.String => JsonSerializer.Serialize(NormalizeText(element.GetString() ?? string.Empty)),
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "null",
            _ => element.GetRawText()
        };
    }

    private static string NormalizeText(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value.Normalize(NormalizationForm.FormKC))
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(character);
            if (category is UnicodeCategory.NonSpacingMark or
                UnicodeCategory.SpacingCombiningMark or
                UnicodeCategory.EnclosingMark)
                continue;

            builder.Append(char.ToLowerInvariant(character));
        }

        return System.Text.RegularExpressions.Regex
            .Replace(builder.ToString(), @"\s+", " ")
            .Trim();
    }

    private static string Hash(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
}
