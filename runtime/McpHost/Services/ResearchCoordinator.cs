using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;
using ModelContextProtocol.Protocol;

namespace ResearchTools.McpHost.Services;

public sealed class ResearchBusyException : Exception
{
    public ResearchBusyException() : base("محرك البحث المحلي مشغول بطلبات أخرى؛ أعد المحاولة بعد قليل.") { }
}

/// <summary>
/// Orchestrates JSON-decision turns against official MCP clients. The installed DuckAI
/// endpoint returns plain text instead of tool_calls even when a function is forced, so
/// the host uses the explicit JSON protocol from the first model turn.
/// </summary>
public sealed class ResearchCoordinator
{
    private const int MaxConversationMessages = 40;

    private readonly DuckAiClient _duckAi;
    private readonly McpSourceRegistry _sources;
    private readonly HostLog _log;
    private readonly ResearchProviderOptions _providerOptions;
    private readonly LmStudioResearchAgent _lmStudioResearchAgent;
    private readonly SemaphoreSlim _concurrency = new(2, 2);

    public ResearchCoordinator(
        DuckAiClient duckAi,
        McpSourceRegistry sources,
        HostLog log,
        ResearchProviderOptions providerOptions,
        LmStudioResearchAgent lmStudioResearchAgent)
    {
        _duckAi = duckAi;
        _sources = sources;
        _log = log;
        _providerOptions = providerOptions;
        _lmStudioResearchAgent = lmStudioResearchAgent;
    }

    public async Task<ResearchResponse> RunAsync(
        ResearchRequest request,
        CancellationToken requestCancellation)
    {
        if (!await _concurrency.WaitAsync(0, requestCancellation))
            throw new ResearchBusyException();

        var taskId = Guid.NewGuid().ToString("N")[..10];
        _log.Info($"Research started: task={taskId}.");

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
        timeout.CancelAfter(TimeSpan.FromSeconds(300));

        var selectedServerIds = ResearchSourceSelection.Normalize(request.Servers);
        var warnings = new List<string>();
        var messages = new List<DuckChatMessage>();
        var trace = new List<ToolCallTrace>();
        var evidenceLedger = new ResearchEvidenceLedger();
        var model = request.Model ?? string.Empty;

        try
        {
            if (_providerOptions.UseLmStudio)
                return await _lmStudioResearchAgent.RunAsync(request, requestCancellation, taskId);

            var incoming = (request.Messages ?? [])
                .Where(message =>
                    message is not null &&
                    (message.Role == "user" || message.Role == "assistant") &&
                    !string.IsNullOrWhiteSpace(message.GetText()))
                .TakeLast(MaxConversationMessages)
                .Select(message => new DuckChatMessage(
                    message.Role!,
                    message.GetText().Trim()))
                .ToList();

            if (incoming.Count == 0)
                throw new InvalidOperationException("لا توجد رسائل بحث صالحة.");

            // Keep the latest user request available in every model turn, even when
            // older history and tool results need to be trimmed to the context budget.
            var currentUserMessageIndex = incoming.FindLastIndex(message => message.Role == "user");
            messages.AddRange(incoming);
            if (selectedServerIds.Count == 0)
                throw new InvalidOperationException("يجب تحديد الجامع أو المكتبة الشاملة أو كليهما قبل البحث.");

            var discovery = await _sources.DiscoverSelectedAsync(selectedServerIds, timeout.Token);
            warnings.AddRange(discovery.Warnings);
            var activeTools = discovery.Tools;
            var unavailableSources = selectedServerIds
                .Where(server => !activeTools.Any(tool =>
                    string.Equals(tool.Server, server, StringComparison.Ordinal)))
                .ToArray();

            if (unavailableSources.Length > 0)
            {
                var unavailableLabels = unavailableSources.Select(ResearchSourceSelection.GetLabel);
                warnings.Add(
                    "لم تتوفر أدوات قراءة مسموح بها في المكتبات المختارة الآتية: " +
                    string.Join("، ", unavailableLabels) +
                    ". سيواصل النموذج باستخدام الأدوات المتاحة، ولن تُعد المكتبات غير المتاحة دليلاً على عدم وجود مصادر.");
            }

            if (activeTools.Count == 0)
            {
                var unavailableLabels = selectedServerIds.Select(ResearchSourceSelection.GetLabel);
                var answer =
                    "تعذر بدء البحث؛ لم تتوفر أدوات قراءة مسموح بها في المكتبات التي اخترتها: " +
                    string.Join("، ", unavailableLabels) +
                    ". لم يُنفذ بحث يمكن الاستناد إلى نتائجه، ولا يصح استنتاج عدم وجود مصادر للمسألة.";
                return new ResearchResponse(answer, model, trace.ToArray(), warnings.ToArray());
            }

            var selectedLabels = selectedServerIds
                .Select(ResearchSourceSelection.GetLabel)
                .ToArray();
            var contextBudget = ResearchContextBudget.Load();

            // All selected-source tools remain visible throughout the task.
            var repairAttempted = false;
            var toolCallsAttempted = 0;

            while (true)
            {
                IReadOnlyList<ToolDescriptor> turnTools = activeTools;
                var turnSystemPrompt = BuildSystemPrompt(turnTools, selectedLabels, warnings) +
                    Environment.NewLine + Environment.NewLine +
                    "تواصل مع مضيف MCP عبر بروتوكول قرارات JSON. أخرج كائن JSON واحداً مطابقاً للمخطط المرفق؛ في tool_call املأ type وserver وtool وarguments واترك answer فارغاً، وفي final املأ type وanswer واجعل server وtool فارغين وarguments null.";

                var availableInputCharacters = contextBudget.GetAvailableInputCharacters(turnSystemPrompt.Length);
                if (availableInputCharacters <= 1)
                    throw new InvalidOperationException(
                        "تجاوزت أوصاف ومخططات أدوات MCP ميزانية السياق المتاحة. ارفع RESEARCHTOOLS_CONTEXT_BUDGET_CHARS بما يناسب سعة النموذج.");

                var evidenceBudget = contextBudget.GetEvidenceBudget(
                    availableInputCharacters,
                    evidenceLedger.EstimatedContextCharacters);
                var evidenceContext = evidenceBudget > 0
                    ? evidenceLedger.BuildContext(selectedServerIds, Array.Empty<string>(), evidenceBudget)
                    : string.Empty;
                var historyBudget = Math.Max(1, availableInputCharacters - evidenceContext.Length);
                var conversation = ResearchContextWindow.Compact(
                    messages.ToArray(),
                    historyBudget,
                    contextBudget.MaxContextCharacters,
                    currentUserMessageIndex);

                var requestMessages = new List<DuckChatMessage>(conversation.Count + 2)
                {
                    new("system", turnSystemPrompt)
                };
                requestMessages.AddRange(conversation);
                if (!string.IsNullOrWhiteSpace(evidenceContext))
                    requestMessages.Add(new DuckChatMessage("user", evidenceContext));

                var originalContextCharacters = messages
                    .Sum(message => message.Content?.Length ?? 0);
                var boundedContextCharacters = conversation
                    .Sum(message => message.Content?.Length ?? 0);

                if (boundedContextCharacters < originalContextCharacters ||
                    conversation.Count < messages.Count - 1 ||
                    !string.IsNullOrWhiteSpace(evidenceContext))
                {
                    _log.Info(
                        $"Research context bounded: task={taskId}; messages={conversation.Count}/{messages.Count - 1}; historyCharacters={boundedContextCharacters}/{originalContextCharacters}; evidenceCharacters={evidenceContext.Length}; budget={contextBudget.MaxContextCharacters}.");
                }

                var responseFormat = ModelDecisionSchemaBuilder.BuildResponseFormat(
                    turnTools,
                    allowFinal: true,
                    allowToolCalls: true);

                var completion = await _duckAi.CompleteAsync(
                    model,
                    requestMessages,
                    functionTools: null,
                    cancellationToken: timeout.Token,
                    responseFormat: responseFormat);

                model = completion.Model;

                if (completion.ToolCalls.Count > 0)
                    throw new InvalidOperationException(
                        "أعاد DuckAI استدعاءات وظائف غير مطلوبة؛ تم رفضها لأن المضيف يستخدم بروتوكول قرارات JSON.");

                if (!ModelDecisionParser.TryParse(completion.Text, out var decision))
                {
                    _log.Warning(
                        $"Research model decision invalid: task={taskId}; responseCharacters={completion.Text.Length}; responseFormat=json_schema.");

                    if (repairAttempted)
                        throw new InvalidOperationException("تعذر التحقق من استجابة النموذج وفق مخطط JSON الصارم بعد محاولة الإصلاح.");

                    repairAttempted = true;
                    messages.Add(new DuckChatMessage("assistant", completion.Text));
                    messages.Add(new DuckChatMessage(
                        "user",
                        "لم تطابق استجابتك مخطط JSON الصارم. أعد القرار وفق المخطط المرفق حرفياً، دون أي شرح أو Markdown. لا تغيّر نوع القرار أو المكتبة أو الأداة المسموح بها."));
                    continue;
                }

                repairAttempted = false;

                if (decision.Type == "final")
                {
                    _log.Info($"Research completed: task={taskId}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}; mode=json-schema; model={model}.");
                    return new ResearchResponse(
                        decision.Answer,
                        model,
                        trace.ToArray(),
                        warnings.ToArray());
                }

                var descriptorForJson = turnTools.FirstOrDefault(tool =>
                    tool.Server == decision.Server && tool.Name == decision.Tool);

                messages.Add(new DuckChatMessage("assistant", decision.Raw));

                if (descriptorForJson is null)
                {
                    _log.Warning(
                        $"Research tool rejected: task={taskId}; source={decision.Server}; tool={decision.Tool}; reason=not-in-selected-source-catalog.");
                    messages.Add(new DuckChatMessage(
                        "user",
                        "الأداة التي اخترتها غير موجودة في فهرس المكتبات المختارة. استخدم أداة ظاهرة في فهرس MCP فقط، ولا تخترع أسماء أدوات أو مصادر."));
                    continue;
                }

                toolCallsAttempted++;
                _log.Info(
                    $"Research tool decision: task={taskId}; call={toolCallsAttempted}; source={decision.Server}; tool={decision.Tool}.");

                var execution = await ExecuteToolAsync(
                    descriptorForJson,
                    ToArguments(decision.Arguments),
                    warnings,
                    taskId,
                    timeout.Token);

                evidenceLedger.Add(descriptorForJson, execution.Content, execution.Succeeded);

                if (execution.Succeeded)
                    trace.Add(new ToolCallTrace(
                        descriptorForJson.Server,
                        descriptorForJson.ServerLabel,
                        descriptorForJson.Name));

                _log.Info(
                    $"Research tool result: task={taskId}; source={decision.Server}; tool={decision.Tool}; success={execution.Succeeded}; evidenceEntries={evidenceLedger.Entries.Count}.");
            }

        }
        catch (OperationCanceledException) when (!requestCancellation.IsCancellationRequested)
        {
            _log.Warning($"Research search window elapsed: task={taskId}; toolCalls={trace.Count}.");
            warnings.Add("انتهت مهلة البحث البالغة خمس دقائق؛ أُوقف البحث الجديد وانتقل النظام إلى عرض ما جُمع حتى لحظة التوقف.");

            var pendingSources = Array.Empty<string>();
            var partial = await FinalizeTimedOutResearchAsync(
                model,
                messages,
                evidenceLedger,
                selectedServerIds,
                pendingSources,
                warnings,
                requestCancellation,
                taskId);

            _log.Info(
                $"Research returned partial result after timeout: task={taskId}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}; model={partial.Model}.");

            return new ResearchResponse(
                partial.Answer,
                partial.Model,
                trace.ToArray(),
                warnings.ToArray());
        }
        catch (Exception exception)
        {
            _log.Error($"Research failed: task={taskId}", exception);
            throw;
        }
        finally
        {
            _concurrency.Release();
        }
    }


    private async Task<(string Answer, string Model)> FinalizeTimedOutResearchAsync(
        string model,
        IReadOnlyList<DuckChatMessage> messages,
        ResearchEvidenceLedger evidenceLedger,
        IReadOnlyList<string> selectedServerIds,
        IReadOnlyList<string> pendingSourceIds,
        List<string> warnings,
        CancellationToken requestCancellation,
        string taskId)
    {
        var systemPrompt = ResearchAgentInstructions.Build(
            selectedServerIds.Select(ResearchSourceSelection.GetLabel).ToArray(),
            warnings) +
            Environment.NewLine + Environment.NewLine +
            "انتهت مهلة البحث؛ هذه مرحلة صياغة نهائية ولا تتاح فيها استدعاءات أدوات جديدة. اعرض ما تسمح به الأدلة المتاحة ووضّح أن البحث لم يكتمل." +
            Environment.NewLine +
            "أخرج JSON واحداً وفق المخطط: type=final، وanswer نص الإجابة، وserver وtool فارغان، وarguments=null."

        try
        {
            using var finalizationTimeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
            finalizationTimeout.CancelAfter(TimeSpan.FromSeconds(30));

            var contextBudget = ResearchContextBudget.Load();
            var availableInputCharacters = contextBudget.GetAvailableInputCharacters(systemPrompt.Length);
            var evidenceBudget = contextBudget.GetEvidenceBudget(
                availableInputCharacters,
                evidenceLedger.EstimatedContextCharacters);
            var evidenceContext = evidenceBudget > 0
                ? evidenceLedger.BuildContext(selectedServerIds, pendingSourceIds, evidenceBudget)
                : string.Empty;

            var historyBudget = Math.Max(1, availableInputCharacters - evidenceContext.Length);
            var currentUserMessageIndex = messages.ToList()
                .FindLastIndex(message => message.Role == "user");
            var conversation = messages.Count == 0
                ? Array.Empty<DuckChatMessage>()
                : ResearchContextWindow.Compact(
                    messages.ToArray(),
                    historyBudget,
                    contextBudget.MaxContextCharacters,
                    currentUserMessageIndex);

            var requestMessages = new List<DuckChatMessage>(conversation.Count + 2)
            {
                new("system", systemPrompt +
                    Environment.NewLine + Environment.NewLine +
                    "المكتبات التي اختارها المستخدم: " +
                    string.Join("، ", selectedServerIds.Select(ResearchSourceSelection.GetLabel)) + "." +
                    Environment.NewLine +
                    "المكتبات التي لم يكتمل فيها استدعاء بحث صالح: " +
                    (pendingSourceIds.Count == 0
                        ? "لا يوجد."
                        : string.Join("، ", pendingSourceIds.Select(ResearchSourceSelection.GetLabel))) +
                    Environment.NewLine +
                    "ملاحظات الاتصال والتنفيذ: " +
                    (warnings.Count == 0 ? "لا توجد." : string.Join(" | ", warnings)))
            };
            requestMessages.AddRange(conversation);
            if (!string.IsNullOrWhiteSpace(evidenceContext))
                requestMessages.Add(new DuckChatMessage("user", evidenceContext));

            var responseFormat = ModelDecisionSchemaBuilder.BuildResponseFormat(
                Array.Empty<ToolDescriptor>(),
                allowFinal: true,
                allowToolCalls: false);

            var completion = await _duckAi.CompleteAsync(
                model,
                requestMessages,
                functionTools: null,
                cancellationToken: finalizationTimeout.Token,
                responseFormat: responseFormat);

            if (completion.ToolCalls.Count > 0)
                throw new InvalidOperationException("أعاد النموذج استدعاء أداة في مرحلة الإغلاق التي لا تسمح باستدعاء أدوات.");

            if (ModelDecisionParser.TryParse(completion.Text, out var decision) &&
                decision.Type == "final" &&
                !string.IsNullOrWhiteSpace(decision.Answer))
            {
                return (decision.Answer, completion.Model);
            }

            _log.Warning(
                $"Research timeout finalization returned an invalid final decision: task={taskId}; characters={completion.Text.Length}.");
            warnings.Add("تعذر التحقق من الصياغة النهائية بعد انتهاء المهلة؛ عُرضت مقتطفات النتائج التي جُمعت بدلاً من إسقاطها.");
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception exception)
        {
            _log.Warning(
                $"Research timeout finalization failed: task={taskId}; {exception.GetType().Name}: {exception.Message}");
            warnings.Add("تعذر توليد التحليل النهائي بعد انتهاء المهلة؛ عُرضت مقتطفات النتائج التي جُمعت بدلاً من إسقاطها.");
        }

        return (BuildTimeoutFallback(
            evidenceLedger,
            selectedServerIds,
            pendingSourceIds), model);
    }

    private static string BuildTimeoutFallback(
        ResearchEvidenceLedger evidenceLedger,
        IReadOnlyList<string> selectedServerIds,
        IReadOnlyList<string> pendingSourceIds)
    {
        var entries = evidenceLedger.Entries;
        var builder = new System.Text.StringBuilder();

        if (entries.Count == 0)
        {
            builder.Append("انتهت مهلة البحث قبل الحصول على محتوى يمكن عرضه من المكتبات المختارة: ")
                .Append(string.Join("، ", selectedServerIds.Select(ResearchSourceSelection.GetLabel)))
                .AppendLine(".");
            builder.Append("لم ترد نتائج موثقة يمكن الاستناد إليها خلال مدة البحث. وهذا لا يثبت عدم وجود مصادر للمسألة، وإنما يعني أن البحث لم يتمكن من الوصول إلى نتائج متاحة قبل انتهاء المهلة.");
            return builder.ToString();
        }

        builder.AppendLine("انتهت مهلة البحث، ولذلك لم يكتمل الاستقصاء. فيما يأتي ما أمكن جمعه قبل التوقف، دون إسقاط النتائج المتاحة:");

        var maxFallbackCharacters = ResearchContextBudget.Load().MaxContextCharacters;
        foreach (var server in selectedServerIds)
        {
            if (builder.Length >= maxFallbackCharacters)
                break;

            var sourceEntries = entries
                .Where(entry => string.Equals(entry.Server, server, StringComparison.Ordinal))
                .ToArray();

            builder.AppendLine();
            builder.Append("### ").AppendLine(ResearchSourceSelection.GetLabel(server));

            if (sourceEntries.Length == 0)
            {
                builder.AppendLine("لم تصل من هذه المكتبة نتيجة بحث قبل انتهاء المهلة.");
                continue;
            }

            foreach (var entry in sourceEntries)
            {
                if (builder.Length >= maxFallbackCharacters)
                    break;

                if (!entry.Succeeded)
                    builder.AppendLine("ملاحظة: لم تُصنّف استجابة الأداة على أنها نتيجة صالحة.");

                var remaining = maxFallbackCharacters - builder.Length;
                var contentLimit = Math.Max(0, remaining - 80);
                if (contentLimit == 0)
                    break;

                var content = entry.Content ?? string.Empty;
                if (content.Length > contentLimit)
                    content = content[..contentLimit] + " [اقتُطع حجم العرض، مع بقاء الأصل في سجل المهمة]";

                builder.AppendLine(content);
            }
        }

        if (builder.Length >= maxFallbackCharacters)
            builder.AppendLine("اقتُطع العرض الاحتياطي عند الحد الآمن لطول الإجابة؛ ظلت بقية المواد محفوظة في سجل المهمة.");

        if (pendingSourceIds.Count > 0)
        {
            builder.AppendLine();
            builder.Append("لم يكتمل البحث في: ")
                .Append(string.Join("، ", pendingSourceIds.Select(ResearchSourceSelection.GetLabel)))
                .Append('.');
        }

        return builder.ToString();
    }

    private async Task<ToolExecutionResult> ExecuteToolAsync(
        ToolDescriptor descriptor,
        IReadOnlyDictionary<string, object?> arguments,
        List<string> warnings,
        string taskId,
        CancellationToken cancellationToken)
    {
        try
        {
            var result = await _sources.CallToolAsync(
                descriptor.Server,
                descriptor.Name,
                arguments,
                cancellationToken);

            var formattedResult = FormatToolResult(result);
            if (result.IsError == true)
            {
                warnings.Add($"{descriptor.Server}/{descriptor.Name}: MCP reported a tool execution error");
                return new ToolExecutionResult(formattedResult, false);
            }

            if (!ResearchToolResultClassifier.HasUsableContent(result))
            {
                if (ResearchToolResultClassifier.IsCompletedSearch(descriptor, result.IsError == true))
                {
                    // A successful search with zero matches is still a completed search.
                    // Keep its explicit no-result/empty response in the evidence ledger
                    // and do not force the model to repeat that source indefinitely.
                    warnings.Add($"{descriptor.Server}/{descriptor.Name}: search completed without usable matches");
                    return new ToolExecutionResult(formattedResult, true);
                }

                warnings.Add($"{descriptor.Server}/{descriptor.Name}: MCP returned no usable content");
                return new ToolExecutionResult(formattedResult, false);
            }

            return new ToolExecutionResult(formattedResult, true);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            _log.Warning(
                $"Tool call failed: task={taskId}; source={descriptor.Server}; tool={descriptor.Name}; {exception.GetType().Name}");
            warnings.Add($"{descriptor.Server}/{descriptor.Name}: {exception.GetType().Name}");
            return new ToolExecutionResult(
                "تعذر تنفيذ أداة القراءة المطلوبة. السبب: " + exception.Message +
                ". لا تدّع الحصول على نتائج منها؛ جرّب أداة قراءة أخرى إن كانت مناسبة، أو وضّح حدود الوصول للمصادر.",
                false);
        }
    }

    private static string BuildSystemPrompt(
        IReadOnlyList<ToolDescriptor> tools,
        IReadOnlyList<string> selectedLabels,
        IReadOnlyList<string> warnings)
    {
        var builder = new System.Text.StringBuilder(
            ResearchAgentInstructions.Build(selectedLabels, warnings));
        builder.AppendLine().AppendLine("أدوات MCP المتاحة ومخططات معاملاتها:");
        builder.AppendLine(JsonSerializer.Serialize(tools, JsonDefaults.Web));
        return builder.ToString();
    }

    private static IReadOnlyDictionary<string, object?> ToArguments(JsonElement arguments)
    {
        if (arguments.ValueKind != JsonValueKind.Object)
            throw new InvalidOperationException("يجب أن تكون معاملات الأداة كائن JSON.");

        var values = new Dictionary<string, object?>(StringComparer.Ordinal);
        foreach (var property in arguments.EnumerateObject())
        {
            // Strict output schemas require nullable optional arguments to be present.
            // Omit null-valued optional properties before validating the MCP call.
            if (property.Value.ValueKind == JsonValueKind.Null)
                continue;

            values[property.Name] = property.Value.Clone();
        }

        return values;
    }

    private static string FormatToolResult(CallToolResult result)
    {
        var parts = new List<string>();

        if (result.StructuredContent is { } structured)
            parts.Add(structured.GetRawText());

        foreach (var block in result.Content)
        {
            if (block is TextContentBlock text && !string.IsNullOrWhiteSpace(text.Text))
                parts.Add(text.Text);
            else if (block is ResourceLinkBlock link)
                parts.Add($"رابط مورد: {link.Name} — {link.Uri}");
            else if (block is ImageContentBlock)
                parts.Add("[أعادت الأداة محتوى بصرياً لم يُحوّل إلى نص]");
        }

        if (parts.Count == 0)
            parts.Add("لم تُرجع الأداة محتوى نصياً.");

        if (result.IsError == true)
            parts.Insert(0, "أبلغ خادم MCP عن خطأ في تنفيذ الأداة:");

        return string.Join(Environment.NewLine, parts);
    }

    private sealed record ToolExecutionResult(string Content, bool Succeeded);
}
