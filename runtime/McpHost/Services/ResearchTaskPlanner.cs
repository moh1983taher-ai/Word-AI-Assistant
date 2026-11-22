using System.Text.Json;
using System.Text.Json.Nodes;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

internal sealed record ResearchTaskPlanningResult(
    ResearchTaskPlan Plan,
    string Model,
    string? Warning);

/// <summary>
/// Creates a bounded task plan before MCP research begins. Prior conversation is
/// supplied only to resolve references; the latest user message remains the scope boundary.
/// </summary>
internal static class ResearchTaskPlanner
{
    private const int MaxContextMessages = 8;
    private const int MaxContextCharacters = 8_000;

    private const string SystemPrompt = """
أنت مخطط مهمة بحثية في «أدوات الباحث». لا تنفذ البحث ولا تجب عن السؤال.
اقرأ الطلب الحالي، واستعن بالرسائل السابقة فقط لفهم الإحالات والضمائر والأسماء التي يشير إليها الطلب الحالي.
استخرج عناصر العمل التي طلبها المستخدم صراحة أو التي لا يمكن إنجاز المطلوب دونها مباشرة.
إذا كان الطلب سؤالاً واحداً محدداً، فأنشئ عنصراً واحداً فقط. لا تقسّم السؤال إلى خلفية وسياق وتاريخ ونتائج وتطبيقات إلا إذا طلب المستخدم هذه الجوانب.
في الطلبات متعددة الأجزاء، افصل الأجزاء المطلوبة فعلاً، بحد أقصى ستة عناصر.
لا تضف موضوعات مجاورة أو اقتراحات بحثية أو توسعات أكاديمية لم يطلبها المستخدم.
القيود لا تتضمن إلا ما صرّح به المستخدم، مثل الفترة الزمنية أو نوع المصدر أو النطاق أو مستوى التفصيل.
الموضوعات المستبعدة لا تتضمن إلا ما استبعده المستخدم صراحة.
أعد JSON فقط وفق المخطط الصارم. اجعل القوائم فارغة إذا لم يذكر المستخدم قيوداً أو استبعادات صريحة.
""";

    internal static async Task<ResearchTaskPlanningResult> CreateAsync(
        DuckAiClient duckAi,
        IReadOnlyList<DuckChatMessage> conversation,
        int currentUserMessageIndex,
        string requestedModel,
        int selectedSourceCount,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(duckAi);
        ArgumentNullException.ThrowIfNull(conversation);

        var activeUserIndex = currentUserMessageIndex >= 0 &&
                              currentUserMessageIndex < conversation.Count
            ? currentUserMessageIndex
            : conversation.ToList().FindLastIndex(message => message.Role == "user");
        if (activeUserIndex < 0)
            throw new InvalidOperationException("لا توجد رسالة مستخدم حالية يمكن بناء نطاق البحث عليها.");

        var goal = conversation[activeUserIndex].Content?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(goal))
            throw new InvalidOperationException("تعذر تحديد طلب المستخدم الحالي بوصفه نطاقاً للمهمة.");

        try
        {
            var context = BuildContext(conversation, activeUserIndex);
            var response = await duckAi.CompleteAsync(
                requestedModel,
                [
                    new DuckChatMessage("system", SystemPrompt),
                    new DuckChatMessage("user", context)
                ],
                functionTools: null,
                cancellationToken: cancellationToken,
                responseFormat: BuildResponseFormat());

            if (!TryParse(response.Text, out var scopeItems, out var constraints, out var exclusions))
            {
                var fallback = ResearchTaskPlan.Create(goal, [goal], [], [], selectedSourceCount);
                return new ResearchTaskPlanningResult(
                    fallback,
                    response.Model,
                    "تعذر التحقق من خطة البحث؛ استُخدم نص الطلب الحالي نطاقاً مباشراً دون إضافة محاور.");
            }

            return new ResearchTaskPlanningResult(
                ResearchTaskPlan.Create(goal, scopeItems, constraints, exclusions, selectedSourceCount),
                response.Model,
                null);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            var fallback = ResearchTaskPlan.Create(goal, [goal], [], [], selectedSourceCount);
            return new ResearchTaskPlanningResult(
                fallback,
                requestedModel,
                "تعذر بناء خطة مستقلة (" + exception.GetType().Name +
                ")؛ استُخدم نص الطلب الحالي نطاقاً مباشراً دون إضافة محاور.");
        }
    }

    internal static JsonElement BuildResponseFormat()
    {
        var schema = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["scope_items"] = new JsonObject
                {
                    ["type"] = "array",
                    ["items"] = new JsonObject { ["type"] = "string" }
                },
                ["constraints"] = new JsonObject
                {
                    ["type"] = "array",
                    ["items"] = new JsonObject { ["type"] = "string" }
                },
                ["exclusions"] = new JsonObject
                {
                    ["type"] = "array",
                    ["items"] = new JsonObject { ["type"] = "string" }
                }
            },
            ["required"] = new JsonArray("scope_items", "constraints", "exclusions"),
            ["additionalProperties"] = false
        };

        var responseFormat = new JsonObject
        {
            ["type"] = "json_schema",
            ["json_schema"] = new JsonObject
            {
                ["name"] = "research_task_plan",
                ["strict"] = true,
                ["schema"] = schema
            }
        };

        using var document = JsonDocument.Parse(responseFormat.ToJsonString());
        return document.RootElement.Clone();
    }

    internal static bool TryParse(
        string? response,
        out IReadOnlyList<string> scopeItems,
        out IReadOnlyList<string> constraints,
        out IReadOnlyList<string> exclusions)
    {
        scopeItems = [];
        constraints = [];
        exclusions = [];

        if (string.IsNullOrWhiteSpace(response))
            return false;

        var candidate = response.Trim();
        var start = candidate.IndexOf('{');
        var end = candidate.LastIndexOf('}');
        if (start < 0 || end <= start)
            return false;

        try
        {
            using var document = JsonDocument.Parse(candidate[start..(end + 1)]);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object ||
                !TryReadArray(root, "scope_items", out scopeItems) ||
                !TryReadArray(root, "constraints", out constraints) ||
                !TryReadArray(root, "exclusions", out exclusions))
                return false;

            if (scopeItems.Count == 0)
                return false;

            return true;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static bool TryReadArray(
        JsonElement root,
        string propertyName,
        out IReadOnlyList<string> values)
    {
        values = [];
        if (!root.TryGetProperty(propertyName, out var array) ||
            array.ValueKind != JsonValueKind.Array)
            return false;

        var result = new List<string>();
        foreach (var item in array.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
                continue;

            var value = item.GetString()?.Trim();
            if (!string.IsNullOrWhiteSpace(value))
                result.Add(value);
            if (result.Count >= 8)
                break;
        }

        values = result;
        return true;
    }

    private static string BuildContext(
        IReadOnlyList<DuckChatMessage> conversation,
        int currentUserMessageIndex)
    {
        var priorMessages = conversation
            .Select((message, index) => (message, index))
            .Where(item => item.index < currentUserMessageIndex)
            .TakeLast(MaxContextMessages)
            .ToArray();

        var builder = new System.Text.StringBuilder();
        if (priorMessages.Length > 0)
        {
            builder.AppendLine("سياق سابق لفهم الإحالات فقط، وليس تكليفاً مستقلاً:");
            foreach (var item in priorMessages)
            {
                var role = item.message.Role == "assistant" ? "المساعد" : "المستخدم";
                var content = item.message.Content ?? string.Empty;
                if (content.Length > 1_200)
                    content = content[..1_200] + " [اقتطاع]";
                builder.Append(role).Append(": ").AppendLine(content);
            }
        }

        builder.AppendLine("الطلب الحالي الملزم وحده هو:");
        builder.AppendLine(conversation[currentUserMessageIndex].Content);
        var value = builder.ToString();
        return value.Length <= MaxContextCharacters
            ? value
            : value[^MaxContextCharacters..];
    }
}
