using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

internal enum ResearchDepth
{
    Focused,
    Standard,
    Deep
}

internal sealed record ResearchScopeItem(string Id, string Description);

internal sealed class ResearchTaskPlan
{
    private ResearchTaskPlan(
        string goal,
        IReadOnlyList<ResearchScopeItem> scopeItems,
        IReadOnlyList<string> constraints,
        IReadOnlyList<string> exclusions,
        ResearchDepth depth,
        int maxToolCalls,
        int timeoutSeconds)
    {
        Goal = goal;
        ScopeItems = scopeItems;
        Constraints = constraints;
        Exclusions = exclusions;
        Depth = depth;
        MaxToolCalls = maxToolCalls;
        TimeoutSeconds = timeoutSeconds;
    }

    internal string Goal { get; }
    internal IReadOnlyList<ResearchScopeItem> ScopeItems { get; }
    internal IReadOnlyList<string> Constraints { get; }
    internal IReadOnlyList<string> Exclusions { get; }
    internal ResearchDepth Depth { get; }
    internal bool RequiresVerification =>
        Depth == ResearchDepth.Deep ||
        new[]
        {
            "تحقق", "دقق", "تدقيق", "توثيق", "وثّق", "وثق",
            "تخريج", "راجع الإحالات", "تأكد من صحة", "اقتباسات موثقة"
        }.Any(indicator => Goal.Contains(indicator, StringComparison.OrdinalIgnoreCase));
    internal int MaxToolCalls { get; }
    internal int TimeoutSeconds { get; }

    internal static ResearchTaskPlan Create(
        string goal,
        IEnumerable<string> scopeDescriptions,
        IEnumerable<string>? constraints,
        IEnumerable<string>? exclusions,
        int selectedSourceCount)
    {
        var safeGoal = string.IsNullOrWhiteSpace(goal)
            ? "إنجاز طلب المستخدم الحالي"
            : goal.Trim();

        var descriptions = NormalizeList(scopeDescriptions, maximumItems: 6, maximumLength: 280);
        if (descriptions.Count == 0)
            descriptions.Add(safeGoal);

        var items = descriptions
            .Select((description, index) => new ResearchScopeItem($"S{index + 1}", description))
            .ToArray();

        var safeConstraints = NormalizeList(constraints ?? [], 8, 240);
        var safeExclusions = NormalizeList(exclusions ?? [], 8, 240);
        var depth = InferDepth(safeGoal);

        // This is a task-specific allowance, not a fixed stopping point.
        // The independent upper ceiling is only a runaway-loop/resource safeguard.
        var depthAllowance = depth switch
        {
            ResearchDepth.Focused => 0,
            ResearchDepth.Standard => 2,
            ResearchDepth.Deep => 10,
            _ => 0
        };

        var sourceCount = Math.Clamp(selectedSourceCount, 1, 2);
        var callAllowance = 6 + (items.Length * 2) + (sourceCount * 2) + depthAllowance;
        var maxToolCalls = Math.Clamp(callAllowance, 10, 32);

        var timeAllowance = 150 + (items.Length * 15) + (sourceCount * 20) +
                            (depth == ResearchDepth.Deep ? 60 :
                             depth == ResearchDepth.Standard ? 30 : 0);
        var timeoutSeconds = Math.Clamp(timeAllowance, 180, 300);

        return new ResearchTaskPlan(
            safeGoal,
            items,
            safeConstraints,
            safeExclusions,
            depth,
            maxToolCalls,
            timeoutSeconds);
    }

    internal IReadOnlyList<string> MissingScopeItems(IReadOnlySet<string> coveredScopeItemIds) =>
        ScopeItems
            .Where(item => !coveredScopeItemIds.Contains(item.Id))
            .Select(item => item.Id)
            .ToArray();

    internal string BuildProgressInstructions(
        IReadOnlySet<string> coveredScopeItemIds,
        IReadOnlyList<string> pendingSourceIds,
        int toolCallsUsed,
        int noProgressStreak,
        bool verificationRequired,
        IReadOnlySet<string> verifiedScopeItemIds,
        bool verificationToolsAvailable)
    {
        var phase = pendingSourceIds.Count > 0
            ? "الاستكشاف وتغطية المكتبات المختارة"
            : MissingScopeItems(coveredScopeItemIds).Count > 0
                ? "استكمال عناصر النطاق التي لم تحصل على أدلة"
                : verificationRequired && verificationToolsAvailable &&
                  ScopeItems.Any(item => !verifiedScopeItemIds.Contains(item.Id))
                    ? "التحقق من المواضع الأصلية والنقول"
                    : "تركيب الأدلة ومراجعة كفايتها";

        var builder = new System.Text.StringBuilder();
        builder.AppendLine("=== بطاقة المهمة وحدودها ===");
        builder.AppendLine("الهدف الملزم كما طلبه المستخدم: " + Goal);
        builder.AppendLine("درجة العمق المقررة من صياغة الطلب: " + DepthLabel(Depth));
        builder.AppendLine("المرحلة الحالية: " + phase);
        builder.AppendLine($"ميزانية المهمة: {toolCallsUsed}/{MaxToolCalls} استدعاءً منفذاً؛ هذه ميزانية مرتبطة بحجم المهمة وليست شرطاً لإنهائها قبل استيفاء نطاقها.");
        builder.AppendLine("عناصر النطاق المطلوبة حصراً:");
        foreach (var item in ScopeItems)
        {
            var status = coveredScopeItemIds.Contains(item.Id) ? "وصل دليل أولي" : "لم يصل دليل بعد";
            builder.AppendLine($"- {item.Id}: {item.Description} [{status}]");
        }

        if (Constraints.Count > 0)
        {
            builder.AppendLine("قيود صريحة في الطلب:");
            foreach (var constraint in Constraints)
                builder.AppendLine("- " + constraint);
        }

        if (Exclusions.Count > 0)
        {
            builder.AppendLine("موضوعات استبعدها المستخدم صراحة:");
            foreach (var exclusion in Exclusions)
                builder.AppendLine("- " + exclusion);
        }

        builder.AppendLine("قواعد النطاق:");
        builder.AppendLine("- لا تبحث ولا تشرح ولا تضف محوراً لم يرد في الهدف أو عناصر النطاق.");
        builder.AppendLine("- استخدم سياق المحادثة السابقة لفهم الإحالات والضمائر فقط، ولا تحوّله إلى تكليف جديد.");
        builder.AppendLine("- يجب أن يحمل كل قرار tool_call قيمة scope_item واحدة من المعرفات المعروضة، وأن تكون معاملات البحث موجهة إلى وصف ذلك العنصر تحديداً.");
        builder.AppendLine("- لا توسّع البحث إلى خلفيات تاريخية أو تراجم أو موضوعات مجاورة إلا إذا كان ذلك لازماً مباشرة للإجابة أو طلبه المستخدم.");
        builder.AppendLine("- لا تعتبر كثرة النتائج دليلاً على الكفاية؛ افحص صلتها بالسؤال، ولا تكرر استعلاماً أو نتيجة سبق الحصول عليها.");
        builder.AppendLine("- لا تنتقل إلى الصياغة النهائية قبل جمع دليل لكل عنصر مطلوب، وتغطية كل مكتبة اختارها المستخدم، وتنفيذ التحقق المطلوب إن كانت المهمة معمّقة.");
        builder.AppendLine("- عند اكتمال عناصر النطاق والتحقق اللازم، توقف عن البحث وصغ الإجابة؛ لا تستمر لتحصيل معلومات إضافية لا تغيّر الإجابة.");
        
        if (pendingSourceIds.Count > 0)
        {
            builder.AppendLine("أولوية المرحلة: استكمل البحث في المكتبات المختارة التي لم تصل منها نتيجة بحث صالحة، مع ربط كل بحث بعنصر نطاق مناسب.");
            builder.AppendLine("المكتبات المتبقية: " + string.Join("، ", pendingSourceIds.Select(ResearchSourceSelection.GetLabel)));
        }
        else if (MissingScopeItems(coveredScopeItemIds).Count > 0)
        {
            builder.AppendLine("أولوية المرحلة: اختر عنصراً غير مغطى من القائمة، وابحث عنه دون إعادة فتح عناصر مكتملة إلا إذا كان التحقق منها ضرورياً.");
            builder.AppendLine("العناصر التي لم تصلها أدلة: " + string.Join("، ", MissingScopeItems(coveredScopeItemIds)));
        }
        else if (verificationRequired && verificationToolsAvailable &&
                 ScopeItems.Any(item => !verifiedScopeItemIds.Contains(item.Id)))
        {
            var pendingVerification = ScopeItems
                .Where(item => !verifiedScopeItemIds.Contains(item.Id))
                .Select(item => item.Id + ": " + item.Description);
            builder.AppendLine("أولوية المرحلة: اقرأ المواضع الأصلية لكل عنصر لم يُتحقق منه، ولا تنشئ بحثاً موضوعياً جديداً.");
            builder.AppendLine("عناصر النطاق التي ما زالت تحتاج إلى تحقق: " + string.Join("؛ ", pendingVerification));
        }
        else if (verificationRequired && !verificationToolsAvailable)
        {
            builder.AppendLine("اكتملت تغطية عناصر السؤال والمكتبات، لكن تعذر التحقق من المواضع الأصلية لعدم إتاحة أدوات قراءة مناسبة؛ أخرج الإجابة مع بيان هذا القيد صراحة.");
        }
        else
        {
            builder.AppendLine("اكتملت التغطية الأولية لجميع عناصر السؤال والمكتبات المختارة.");
            builder.AppendLine("راجع الآن كفاية الأدلة لكل عنصر: إن كانت كافية فأخرج الإجابة النهائية؛ وإن بقي نقص محدد يؤثر مباشرة في جواب عنصر موجود، فاستدع أداة أخرى مرتبطة بذلك العنصر نفسه.");
            builder.AppendLine("لا تضف عناصر جديدة ولا تفتح موضوعات مجاورة. لا تواصل البحث لمجرد العثور على مزيد من النتائج؛ أوقفه عندما لا تضيف الأدلة المتوقعة قيمة لازمة للإجابة.");
        }

        if (noProgressStreak >= 2)
            builder.AppendLine("لم يحرز آخر استدعاءين تقدماً جديداً. لا تكرر نمط البحث نفسه؛ إن لم توجد خطوة مختلفة ومباشرة تسد نقصاً محدداً، أنهِ البحث بنتيجة جزئية واذكر حدودها.");

        return builder.ToString();
    }

    internal static string DepthLabel(ResearchDepth depth) => depth switch
    {
        ResearchDepth.Focused => "محدد",
        ResearchDepth.Standard => "عادي",
        ResearchDepth.Deep => "معمّق",
        _ => "محدد"
    };

    private static ResearchDepth InferDepth(string goal)
    {
        var text = goal.ToLowerInvariant();
        string[] deepIndicators =
        [
            "بحث معمق", "بحث معمّق", "استقصائي", "استقصاء", "استقراء",
            "دراسة مقارنة", "مقارنة بين", "قارن بين", "مقارنة",
            "من جميع جوانبه", "بالتفصيل", "بحث شامل", "دراسة شاملة",
            "تحليل نقدي", "تتبع الأقوال", "in-depth", "comprehensive", "comparative"
        ];

        if (deepIndicators.Any(indicator => text.Contains(indicator, StringComparison.Ordinal)))
            return ResearchDepth.Deep;

        if (ResearchIntentClassifier.RequiresMcpSources(goal) ||
            text.Contains("حلل", StringComparison.Ordinal) ||
            text.Contains("ناقش", StringComparison.Ordinal) ||
            text.Contains("أثر", StringComparison.Ordinal))
            return ResearchDepth.Standard;

        return ResearchDepth.Focused;
    }

    private static List<string> NormalizeList(
        IEnumerable<string> values,
        int maximumItems,
        int maximumLength)
    {
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var value in values)
        {
            if (string.IsNullOrWhiteSpace(value))
                continue;

            var normalized = System.Text.RegularExpressions.Regex
                .Replace(value.Trim(), @"\s+", " ");
            if (normalized.Length > maximumLength)
                normalized = normalized[..maximumLength];

            if (!seen.Add(normalized))
                continue;

            result.Add(normalized);
            if (result.Count >= maximumItems)
                break;
        }

        return result;
    }
}
