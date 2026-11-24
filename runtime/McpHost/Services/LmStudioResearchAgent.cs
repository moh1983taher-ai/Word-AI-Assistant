using System.Text;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Delegates the complete MCP tool-use loop to LM Studio's native MCP runtime.
/// The host only discovers safe tools, selects the user's source boundary, records
/// completed tool outputs, and validates result provenance.
/// </summary>
public sealed class LmStudioResearchAgent
{
    private readonly LmStudioClient _lmStudio;
    private readonly McpSourceRegistry _sources;
    private readonly HostLog _log;

    public LmStudioResearchAgent(
        LmStudioClient lmStudio,
        McpSourceRegistry sources,
        HostLog log)
    {
        _lmStudio = lmStudio;
        _sources = sources;
        _log = log;
    }

    public async Task<ResearchResponse> RunAsync(
        ResearchRequest request,
        CancellationToken requestCancellation,
        string taskId)
    {
        var selectedServerIds = ResearchSourceSelection.Normalize(request.Servers);
        if (selectedServerIds.Count == 0)
            throw new InvalidOperationException("يجب تحديد الجامع أو المكتبة الشاملة أو كليهما قبل البحث.");

        var messages = NormalizeMessages(request.Messages);
        if (messages.Count == 0)
            throw new InvalidOperationException("لا توجد رسائل بحث صالحة.");

        var contextBudget = ResearchContextBudget.Load();
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
        timeout.CancelAfter(contextBudget.ResearchTimeout);

        var warnings = new List<string>();
        var trace = new List<ToolCallTrace>();
        var evidenceLedger = new ResearchEvidenceLedger();
        var seenToolCalls = new HashSet<string>(StringComparer.Ordinal);
        IReadOnlyList<ToolDescriptor> safeTools = [];
        IReadOnlyList<McpServerInstruction> serverInstructions = [];

        try
        {
            // Discovery is still performed by our official MCP clients. ToolPolicy
            // filters the catalog before any names are handed to LM Studio.
            var discovery = await _sources.DiscoverSelectedAsync(selectedServerIds, timeout.Token);
            warnings.AddRange(discovery.Warnings);

            safeTools = discovery.Tools;
            serverInstructions = discovery.ServerInstructions;
            foreach (var sourceId in selectedServerIds)
            {
                if (!safeTools.Any(tool => string.Equals(tool.Server, sourceId, StringComparison.Ordinal)))
                {
                    warnings.Add(
                        $"لم تعرض مكتبة {ResearchSourceSelection.GetLabel(sourceId)} أدوات قراءة مسموحاً بها؛ لن تُتاح أدواتها لهذا الطلب.");
                }
            }

            if (safeTools.Count == 0)
            {
                return new ResearchResponse(
                    "تعذر بدء البحث؛ لم تتوفر أدوات قراءة آمنة في المكتبات التي اخترتها. لم يُنفذ بحث يمكن الاستناد إلى نتائجه، ولا يدل ذلك على خلو المكتبات من المصادر.",
                    request.Model ?? string.Empty,
                    trace,
                    warnings);
            }

            var modelSnapshot = await _lmStudio.GetModelsAsync(timeout.Token);
            var requestedModel = request.Model?.Trim() ?? string.Empty;
            var selectedModel = requestedModel;
            if (string.IsNullOrWhiteSpace(selectedModel) ||
                !modelSnapshot.Models.Any(model => string.Equals(model.Id, selectedModel, StringComparison.Ordinal)))
            {
                selectedModel = modelSnapshot.DefaultModel;
                if (!string.IsNullOrWhiteSpace(requestedModel))
                {
                    warnings.Add(
                        $"النموذج المحدد سابقاً ({requestedModel}) غير متاح في LM Studio؛ استُخدم النموذج الافتراضي المتاح ({selectedModel}).");
                }
            }

            var modelInfo = modelSnapshot.Models.First(model =>
                string.Equals(model.Id, selectedModel, StringComparison.Ordinal));
            if (!modelInfo.TrainedForToolUse)
            {
                warnings.Add(
                    "لا يعلن النموذج المحدد في LM Studio أنه مدرّب على استخدام الأدوات؛ ستُتاح الأدوات له، لكن جودة استدعائها تعتمد على قدرته الفعلية.");
            }

            var integrations = selectedServerIds
                .Select(sourceId => new
                {
                    Source = sourceId,
                    Tools = safeTools
                        .Where(tool => string.Equals(tool.Server, sourceId, StringComparison.Ordinal))
                        .Select(tool => tool.Name)
                        .Distinct(StringComparer.Ordinal)
                        .ToArray()
                })
                .Where(item => item.Tools.Length > 0)
                .Select(item => new LmStudioIntegration(
                    item.Source,
                    _sources.GetServerEndpoint(item.Source),
                    item.Tools))
                .ToArray();

            var systemPrompt = BuildSystemPrompt(selectedServerIds, warnings, safeTools, serverInstructions);
            var availableInputCharacters = contextBudget.GetAvailableInputCharacters(systemPrompt.Length);
            if (availableInputCharacters <= 1)
            {
                var serverInstructionCharacters = serverInstructions.Sum(instruction => instruction.Text.Length);
                _log.Error(
                    $"LM Studio MCP instructions exceed the configured context budget: task={taskId}; serverInstructionCharacters={serverInstructionCharacters}; systemPromptCharacters={systemPrompt.Length}; effectiveContextBudgetCharacters={contextBudget.MaxContextCharacters}.");
                throw new InvalidOperationException(
                    $"تعذر تجهيز تعليمات خوادم MCP ضمن ميزانية السياق الحالية في LM Studio. بلغ حجم تعليمات النظام {systemPrompt.Length:N0} حرفاً، والحد الفعلي {contextBudget.MaxContextCharacters:N0} حرف.");
            }

            var input = BuildTranscript(
                messages,
                availableInputCharacters,
                contextBudget.MaxContextCharacters);

            async Task RecordToolResultAsync(LmStudioToolResult result)
            {
                var descriptor = safeTools.FirstOrDefault(tool =>
                    string.Equals(tool.Server, result.ServerLabel, StringComparison.Ordinal) &&
                    string.Equals(tool.Name, result.Tool, StringComparison.Ordinal));

                if (descriptor is null ||
                    !selectedServerIds.Contains(result.ServerLabel, StringComparer.Ordinal))
                {
                    warnings.Add(
                        $"أعاد LM Studio نتيجة من أداة غير مطابقة لفهرس الأدوات المسموح به ({result.ServerLabel}/{result.Tool})؛ لم تُستخدم دليلاً.");
                    _log.Warning(
                        $"LM Studio MCP result rejected: task={taskId}; source={result.ServerLabel}; tool={result.Tool}; reason=not-in-selected-safe-catalog.");
                    return;
                }

                var signature = result.ServerLabel + "\n" + result.Tool + "\n" + result.Output;
                if (!seenToolCalls.Add(signature))
                    return;

                evidenceLedger.Add(descriptor, result.Output, result.Succeeded);
                if (result.Succeeded)
                {
                    trace.Add(new ToolCallTrace(
                        descriptor.Server,
                        descriptor.ServerLabel,
                        descriptor.Name));
                }

                _log.Info(
                    $"LM Studio MCP tool completed: task={taskId}; source={descriptor.Server}; tool={descriptor.Name}; succeeded={result.Succeeded}; evidenceEntries={evidenceLedger.Entries.Count}.");
            }

            var completion = await _lmStudio.ChatAsync(
                selectedModel,
                systemPrompt,
                input,
                integrations,
                RecordToolResultAsync,
                timeout.Token);

            warnings.AddRange(completion.Warnings);
            if (!completion.IsComplete)
            {
                warnings.Add("انقطع اتصال بث LM Studio قبل الحدث الختامي؛ لا تُعد الرسالة الجزئية إجابة مكتملة.");
                if (evidenceLedger.Entries.Count > 0)
                {
                    var partial = await FinalizeAfterTimeoutAsync(
                        completion.Model,
                        messages,
                        evidenceLedger,
                        selectedServerIds,
                        safeTools,
                        serverInstructions,
                        warnings,
                        requestCancellation,
                        taskId);
                    return new ResearchResponse(partial.Answer, partial.Model, trace, warnings);
                }

                return new ResearchResponse(
                    string.IsNullOrWhiteSpace(completion.Answer)
                        ? "انقطع اتصال LM Studio قبل إكمال الإجابة، ولم تصل نتائج أدوات يمكن عرضها."
                        : "إجابة غير مكتملة بسبب انقطاع البث قبل الإنهاء:" + Environment.NewLine + completion.Answer,
                    completion.Model,
                    trace,
                    warnings);
            }

            if (string.IsNullOrWhiteSpace(completion.Answer))
            {
                if (evidenceLedger.Entries.Count > 0)
                {
                    warnings.Add("لم يُرجع LM Studio إجابة نهائية قابلة للعرض؛ عُرضت الأدلة التي جُمعت بدلاً من إسقاطها.");
                    return new ResearchResponse(
                        BuildEvidenceFallback(evidenceLedger, selectedServerIds),
                        completion.Model,
                        trace,
                        warnings);
                }

                throw new InvalidOperationException("انتهى LM Studio دون إجابة أو نتائج أدوات قابلة للتحقق.");
            }

            _log.Info(
                $"LM Studio research completed: task={taskId}; model={completion.Model}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}.");

            return new ResearchResponse(
                completion.Answer,
                completion.Model,
                trace,
                warnings);
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            throw;
        }
        catch (OperationCanceledException) when (timeout.IsCancellationRequested)
        {
            warnings.Add("انتهت مهلة البحث البالغة خمس دقائق؛ أُوقف البحث الجديد وحاول النظام صياغة إجابة جزئية من الأدلة المجموعة.");
            _log.Warning(
                $"LM Studio research timed out: task={taskId}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}.");

            var partial = await FinalizeAfterTimeoutAsync(
                request.Model ?? string.Empty,
                messages,
                evidenceLedger,
                selectedServerIds,
                safeTools,
                serverInstructions,
                warnings,
                requestCancellation,
                taskId);

            return new ResearchResponse(partial.Answer, partial.Model, trace, warnings);
        }
    }

    private async Task<(string Answer, string Model)> FinalizeAfterTimeoutAsync(
        string model,
        IReadOnlyList<DuckChatMessage> messages,
        ResearchEvidenceLedger evidenceLedger,
        IReadOnlyList<string> selectedServerIds,
        IReadOnlyList<ToolDescriptor> tools,
        IReadOnlyList<McpServerInstruction> serverInstructions,
        List<string> warnings,
        CancellationToken requestCancellation,
        string taskId)
    {
        var finalizationPrompt = ResearchAgentInstructions.Build(
            selectedServerIds.Select(ResearchSourceSelection.GetLabel).ToArray(),
            warnings,
            tools,
            serverInstructions) +
            Environment.NewLine + Environment.NewLine +
            "انتهت مهلة البحث؛ هذه صياغة نهائية جزئية ولا تتاح فيها استدعاءات أدوات جديدة. استند إلى الأدلة المتاحة وبيّن أن البحث لم يكتمل.";

        var contextBudget = ResearchContextBudget.Load();
        var availableInputCharacters = contextBudget.GetAvailableInputCharacters(finalizationPrompt.Length);
        if (availableInputCharacters <= 1)
        {
            warnings.Add("لم تتسع ميزانية السياق لتعليمات الخادم ومرحلة الصياغة الجزئية؛ عُرضت الأدلة التي جُمعت دون طلب توليد جديد.");
            return (BuildEvidenceFallback(evidenceLedger, selectedServerIds), model);
        }

        var evidenceBudget = contextBudget.GetEvidenceBudget(
            availableInputCharacters,
            evidenceLedger.EstimatedContextCharacters);
        var evidence = evidenceBudget > 0
            ? evidenceLedger.BuildContext(selectedServerIds, Array.Empty<string>(), evidenceBudget)
            : string.Empty;

        if (string.IsNullOrWhiteSpace(evidence))
            return (BuildEvidenceFallback(evidenceLedger, selectedServerIds), model);

        var historyBudget = Math.Max(1, availableInputCharacters - evidence.Length);
        var transcript = BuildTranscript(
            messages,
            historyBudget,
            contextBudget.MaxContextCharacters) +
            Environment.NewLine + Environment.NewLine + evidence;

        try
        {
            using var finalizationTimeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
            finalizationTimeout.CancelAfter(contextBudget.FinalizationTimeout);

            var completion = await _lmStudio.ChatAsync(
                model,
                finalizationPrompt,
                transcript,
                Array.Empty<LmStudioIntegration>(),
                onToolResult: null,
                finalizationTimeout.Token);

            if (!string.IsNullOrWhiteSpace(completion.Answer))
                return (completion.Answer, completion.Model);
        }
        catch (OperationCanceledException) when (requestCancellation.IsCancellationRequested)
        {
            throw;
        }
        catch (Exception exception)
        {
            _log.Warning(
                $"LM Studio timeout finalization failed: task={taskId}; {exception.GetType().Name}: {exception.Message}");
            warnings.Add("تعذرت صياغة التحليل النهائي بعد انتهاء المهلة؛ عُرضت مقتطفات الأدلة التي جُمعت.");
        }

        return (BuildEvidenceFallback(evidenceLedger, selectedServerIds), model);
    }

    private static List<DuckChatMessage> NormalizeMessages(
        IReadOnlyList<ResearchMessage>? incoming)
    {
        return (incoming ?? [])
            .Where(message =>
                message is not null &&
                (message.Role == "user" || message.Role == "assistant") &&
                !string.IsNullOrWhiteSpace(message.GetText()))
            .Select(message => new DuckChatMessage(
                message.Role!,
                message.GetText().Trim()))
            .ToList();
    }

    private static string BuildTranscript(
        IReadOnlyList<DuckChatMessage> messages,
        int maxConversationCharacters,
        int maxMessageCharacters)
    {
        var compacted = ResearchContextWindow.Compact(
            messages.ToArray(),
            maxConversationCharacters,
            maxMessageCharacters,
            messages.ToList().FindLastIndex(message => message.Role == "user"));

        var builder = new StringBuilder();
        foreach (var message in compacted)
        {
            if (builder.Length > 0)
                builder.AppendLine().AppendLine();
            builder.Append(message.Role == "assistant" ? "المساعد:" : "المستخدم:")
                .AppendLine()
                .Append(message.Content);
        }

        return builder.ToString();
    }

    private static string BuildSystemPrompt(
        IReadOnlyList<string> selectedServerIds,
        IReadOnlyList<string> warnings,
        IReadOnlyList<ToolDescriptor> tools,
        IReadOnlyList<McpServerInstruction> serverInstructions)
    {
        return ResearchAgentInstructions.Build(
            selectedServerIds.Select(ResearchSourceSelection.GetLabel).ToArray(),
            warnings,
            tools,
            serverInstructions);
    }

    private static string BuildEvidenceFallback(
        ResearchEvidenceLedger evidenceLedger,
        IReadOnlyList<string> selectedServerIds)
    {
        if (evidenceLedger.Entries.Count == 0)
        {
            return "انتهت مهلة البحث قبل الوصول إلى محتوى يمكن الاستناد إليه من المكتبات المختارة: " +
                string.Join("، ", selectedServerIds.Select(ResearchSourceSelection.GetLabel)) +
                ". لا يدل ذلك على خلوها من المصادر.";
        }

        var builder = new StringBuilder();
        var maxFallbackCharacters = ResearchContextBudget.Load().MaxContextCharacters;
        builder.AppendLine("انتهت مهلة البحث، ولم يكتمل البحث. فيما يأتي ما أمكن جمعه قبل التوقف:");

        foreach (var sourceId in selectedServerIds)
        {
            var entries = evidenceLedger.Entries
                .Where(entry => string.Equals(entry.Server, sourceId, StringComparison.Ordinal))
                .ToArray();
            if (entries.Length == 0)
                continue;

            builder.AppendLine().Append("### ").AppendLine(ResearchSourceSelection.GetLabel(sourceId));
            foreach (var entry in entries)
            {
                if (builder.Length >= maxFallbackCharacters)
                    break;

                var content = entry.Content ?? string.Empty;
                var remaining = Math.Max(0, maxFallbackCharacters - builder.Length);
                if (content.Length > remaining)
                    content = content[..remaining];

                builder.AppendLine(content);
            }
        }

        return builder.ToString();
    }
}
