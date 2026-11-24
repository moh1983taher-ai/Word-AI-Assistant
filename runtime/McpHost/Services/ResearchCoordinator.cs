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
    private const int MaxDuckAiMessageCharacters = 12_000;
    private const int MaxDuckAiEvidenceContextCharacters = 10_000;

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
        var contextBudget = _providerOptions.UseLmStudio
            ? ResearchContextBudget.Load()
            : ResearchContextBudget.LoadDuckAi();
        _log.Info(
            $"Research context profile: task={taskId}; provider={_providerOptions.Provider}; contextBudgetCharacters={contextBudget.MaxContextCharacters}; maxOutputTokens={contextBudget.MaxOutputTokens}; contextBudgetOverride={!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS"))}; outputTokensOverride={!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("RESEARCHTOOLS_MAX_OUTPUT_TOKENS"))}.");

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
        timeout.CancelAfter(contextBudget.ResearchTimeout);

        var selectedServerIds = ResearchSourceSelection.Normalize(request.Servers);
        var warnings = new List<string>();
        var messages = new List<DuckChatMessage>();
        var trace = new List<ToolCallTrace>();
        var evidenceLedger = new ResearchEvidenceLedger();
        var evidenceProvenance = new McpEvidenceProvenance();
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
                .Select(message => new DuckChatMessage(
                    message.Role!,
                    message.GetText().Trim()))
                .ToList();

            if (incoming.Count == 0)
                throw new InvalidOperationException("لا توجد رسائل بحث صالحة.");

            // Keep the latest user request available in every model turn, even when
            // older history and tool results need to be trimmed to the context budget.
            var currentUserMessageIndex = incoming.FindLastIndex(message => message.Role == "user");
            var latestUserMessage = currentUserMessageIndex >= 0
                ? incoming[currentUserMessageIndex].Content
                : null;
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

            // All selected-source tools remain visible throughout the task.
            var repairAttempted = false;
            var searchDecisionCorrectionAttempted = false;
            var pageVerificationCorrectionAttempted = false;
            var toolCallsAttempted = 0;
            var searchedSourceIds = new HashSet<string>(StringComparer.Ordinal);

            // The host, not the model, guarantees the first search call to every
            // selected library. Arguments are derived from each live MCP input schema.
            var initialSearchPlan = ResearchSearchBootstrap.BuildInitialPlan(
                activeTools,
                selectedServerIds,
                latestUserMessage ?? string.Empty);
            _log.Info(
                $"Research host search plan: task={taskId}; selectedSources={string.Join(",", selectedServerIds)}; calls={initialSearchPlan.Count}; protocol=host-owned.");

            foreach (var invocation in initialSearchPlan)
            {
                toolCallsAttempted++;
                _log.Info(
                    $"Research host search dispatch: task={taskId}; call={toolCallsAttempted}; source={invocation.Tool.Server}; tool={invocation.Tool.Name}.");

                var execution = await ExecuteToolAsync(
                    invocation.Tool,
                    invocation.Arguments,
                    warnings,
                    evidenceProvenance,
                    taskId,
                    timeout.Token);

                evidenceLedger.Add(invocation.Tool, execution.Content, execution.Succeeded);

                var completedSearch = execution.Succeeded &&
                    ResearchToolResultClassifier.IsCompletedSearch(invocation.Tool, isError: false);
                if (completedSearch)
                    searchedSourceIds.Add(invocation.Tool.Server);

                if (execution.Succeeded)
                    trace.Add(new ToolCallTrace(
                        invocation.Tool.Server,
                        invocation.Tool.ServerLabel,
                        invocation.Tool.Name));

                _log.Info(
                    $"Research host search result: task={taskId}; source={invocation.Tool.Server}; tool={invocation.Tool.Name}; success={execution.Succeeded}; completedSearch={completedSearch}; searchCoverage={searchedSourceIds.Count}/{selectedServerIds.Count}; evidenceEntries={evidenceLedger.Entries.Count}.");
            }

            while (true)
            {
                IReadOnlyList<ToolDescriptor> turnTools = activeTools;
                var pendingSearchSources = ResearchIntentClassifier.GetPendingSearchSources(
                    latestUserMessage,
                    selectedServerIds,
                    searchedSourceIds);
                var allowFinal = pendingSearchSources.Count == 0;
                var turnSystemPrompt = BuildSystemPrompt(
                    turnTools,
                    selectedLabels,
                    warnings,
                    out var toolCatalogCharacters) +
                    Environment.NewLine + Environment.NewLine +
                    (allowFinal
                        ? "أعد كائن JSON واحداً فقط وفق بروتوكول القرارات الموضح أدناه. استخدم final مع answer عند اكتمال البحث، أو tool_call مع server وtool وarguments عند الحاجة إلى فتح صفحة أو بحث تكميلي. لا تضع Markdown أو شرحاً خارج JSON."
                        : "هذه مهمة بحث مصدري لم يكتمل البحث فيها بعد. المكتبات التي يجب أن تبحث فيها الآن: " +
                          string.Join("، ", pendingSearchSources.Select(source =>
                              $"{ResearchSourceSelection.GetLabel(source)} ({(source == "aljam3" ? "search" : "shamela_find أو shamela_find_scoped")})")) +
                          ". لا تصدر final ولا تجب من الذاكرة. أخرج قرار tool_call صحيحاً وفق مخططات الأدوات المعروضة في تعليمات النظام.");

                // Tool input schemas are included in the system prompt as ordinary data.
                // Duck2API does not enforce response_format=json_schema, so no hidden
                // schema payload is sent or counted here.
                var fixedRequestCharacters = turnSystemPrompt.Length;
                var availableInputCharacters = contextBudget.GetAvailableInputCharacters(fixedRequestCharacters);
                if (availableInputCharacters <= 1)
                {
                    var minimumRequiredCharacters = fixedRequestCharacters + 2_048 + 2;
                    var configuredBudget = Environment.GetEnvironmentVariable("RESEARCHTOOLS_CONTEXT_BUDGET_CHARS");
                    _log.Error(
                        $"MCP tool catalog exceeds the configured context budget: task={taskId}; model={model}; toolCount={turnTools.Count}; toolCatalogCharacters={toolCatalogCharacters}; systemPromptCharacters={turnSystemPrompt.Length}; minimumRequiredContextCharacters={minimumRequiredCharacters}; effectiveContextBudgetCharacters={contextBudget.MaxContextCharacters}; configuredContextBudget={configuredBudget ?? "(unset)"}.");
                    throw new InvalidOperationException(
                        $"تعذر تجهيز فهرس أدوات MCP وتعليماته ضمن ميزانية السياق الحالية. حجم تعليمات الأدوات ومخططاتها {turnSystemPrompt.Length:N0} حرف، والحد الفعلي {contextBudget.MaxContextCharacters:N0} حرف، والحد الأدنى اللازم مع احتياطي البروتوكول {minimumRequiredCharacters:N0} حرف. راجع سجل المضيف واضبط RESEARCHTOOLS_CONTEXT_BUDGET_CHARS فقط إذا كان النموذج والبوابة يدعمان سعة أكبر.");
                }

                var evidenceBudget = contextBudget.GetEvidenceBudget(
                    availableInputCharacters,
                    evidenceLedger.EstimatedContextCharacters,
                    MaxDuckAiEvidenceContextCharacters);
                var evidenceContext = evidenceBudget > 0
                    ? evidenceLedger.BuildContext(selectedServerIds, Array.Empty<string>(), evidenceBudget)
                    : string.Empty;
                var historyBudget = Math.Max(1, availableInputCharacters - evidenceContext.Length);
                var conversation = ResearchContextWindow.Compact(
                    messages.ToArray(),
                    historyBudget,
                    MaxDuckAiMessageCharacters,
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

                var requestCharacters = requestMessages.Sum(message => message.Content?.Length ?? 0);
                _log.Info(
                    $"Research prompt prepared: task={taskId}; selectedSources={string.Join(",", selectedServerIds)}; messages={requestMessages.Count}; systemCharacters={turnSystemPrompt.Length}; toolCatalogCharacters={toolCatalogCharacters}; responseFormat=omitted; pendingSearchSources={pendingSearchSources.Count}; finalAllowed={allowFinal}; historyCharacters={boundedContextCharacters}/{originalContextCharacters}; evidenceCharacters={evidenceContext.Length}; requestCharacters={requestCharacters}; fixedRequestCharacters={fixedRequestCharacters}; contextBudget={contextBudget.MaxContextCharacters}; maxOutputTokens={contextBudget.MaxOutputTokens}.");

                var completion = await _duckAi.CompleteAsync(
                    model,
                    requestMessages,
                    functionTools: null,
                    cancellationToken: timeout.Token);

                model = completion.Model;

                if (completion.ToolCalls.Count > 0)
                    throw new InvalidOperationException(
                        "أعاد DuckAI استدعاءات وظائف غير مطلوبة؛ تم رفضها لأن المضيف يستخدم بروتوكول قرارات JSON.");

                if (!ModelDecisionParser.TryParse(completion.Text, out var decision))
                {
                    _log.Warning(
                        $"Research model decision invalid: task={taskId}; responseCharacters={completion.Text.Length}; responseFormat=omitted; protocol=prompted-json.");

                    if (repairAttempted)
                        throw new InvalidOperationException("تعذر التحقق من استجابة النموذج وفق مخطط JSON الصارم بعد محاولة الإصلاح.");

                    repairAttempted = true;
                    messages.Add(new DuckChatMessage("assistant", completion.Text));
                    messages.Add(new DuckChatMessage(
                        "user",
                        "لم تكن استجابتك قرار JSON صالحاً. أخرج كائناً واحداً فقط: إما {\"type\":\"tool_call\",\"server\":\"معرّف المكتبة\",\"tool\":\"اسم الأداة\",\"arguments\":{...}}، أو {\"type\":\"final\",\"answer\":\"نص الإجابة\"}. استخدم اسم أداة ومعاملاتها كما وردت في فهرس MCP ومخططها المعروضين في تعليمات النظام. لا تضع أي نص خارج الكائن."));
                    continue;
                }

                repairAttempted = false;

                if (decision.Type == "final")
                {
                    var searchWasSkipped = pendingSearchSources.Count > 0;
                    // Every request reaching this coordinator is an explicit source-research
                    // request. Verify opened pages regardless of the wording used by the user.
                    var matchingPageWasNotOpened =
                        evidenceProvenance.HasUnverifiedSearchResults;

                    if (searchWasSkipped)
                    {
                        if (!searchDecisionCorrectionAttempted)
                        {
                            searchDecisionCorrectionAttempted = true;
                            _log.Warning(
                                $"Research premature final deferred: task={taskId}; reason=explicit_source_request_without_successful_tool_call; recovery=one_search_correction.");

                            messages.Add(new DuckChatMessage("assistant", decision.Raw));
                            messages.Add(new DuckChatMessage(
                                "user",
                                "تنبيه تحقق: لا يزال البحث غير منفذ في المكتبات الآتية: " +
                                string.Join("، ", pendingSearchSources.Select(ResearchSourceSelection.GetLabel)) +
                                ". لا تصدر إجابة نهائية ولا تعتمد على الذاكرة. نفّذ أداة البحث الصحيحة في كل مكتبة متبقية، ثم افتح الصفحات المطابقة للنتائج التي ستستشهد بها."));
                            continue;
                        }

                        _log.Warning(
                            $"Research final rejected: task={taskId}; reason=explicit_source_request_without_successful_tool_call; correctionAttempted=true.");
                        throw new InvalidOperationException(
                            "تعذر إتمام البحث المطلوب: أنهى النموذج المهمة دون تنفيذ استدعاء ناجح لأداة بحث من المكتبات المختارة. لم تُعرض إجابته بوصفها بحثاً موثقاً.");
                    }

                    if (matchingPageWasNotOpened)
                    {
                        if (!pageVerificationCorrectionAttempted)
                        {
                            pageVerificationCorrectionAttempted = true;
                            _log.Warning(
                                $"Research premature final deferred: task={taskId}; reason=search_results_without_verified_page_open; searchResponses={evidenceProvenance.NonEmptySearchResponseCount}; searchReferences={evidenceProvenance.SearchReferenceCount}; verifiedOpens={evidenceProvenance.VerifiedPageOpenCount}; recovery=one_page_verification_correction.");

                            messages.Add(new DuckChatMessage("assistant", decision.Raw));
                            messages.Add(new DuckChatMessage(
                                "user",
                                "تنبيه تحقق: أعادت مكتبة واحدة أو أكثر نتائج بحث فعلية، لكن لم يُفتح موضع مطابق من كل مكتبة ظهرت منها نتائج. افتح صفحة مطابقة واحدة على الأقل من كل مكتبة ذات نتائج، مستخدماً معرّفات الكتاب والصفحة الواردة حرفياً في نتائج تلك المكتبة نفسها. افحص النص وسياقه قبل الاستشهاد. إذا لم يتوفر معرّف قابل للتحقق، فصرّح بتعذر فتح المصدر ولا تستشهد بمقتطف البحث."));
                            continue;
                        }

                        _log.Warning(
                            $"Research final rejected: task={taskId}; reason=search_results_without_verified_page_open; correctionAttempted=true.");
                        throw new InvalidOperationException(
                            "تعذر اعتماد إجابة البحث: بقيت مكتبة أعادت نتائج فعلية من دون فتح صفحة مطابقة والتحقق من معرّفها. أُوقف تمرير الإجابة غير المكتملة بدلاً من عرضها بوصفها بحثاً موثقاً.");
                    }

                    _log.Info($"Research completed: task={taskId}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}; mode=prompted-json; model={model}.");
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

                var toolArguments = ToArguments(decision.Arguments);
                var referenceError = evidenceProvenance.ValidateReferenceRead(
                    descriptorForJson,
                    toolArguments);
                if (referenceError is not null)
                {
                    warnings.Add($"{descriptorForJson.Server}/{descriptorForJson.Name}: rejected unverified page reference");
                    _log.Warning(
                        $"MCP page reference rejected: task={taskId}; source={descriptorForJson.Server}; tool={descriptorForJson.Name}; knownReferences={evidenceProvenance.SearchReferenceCount}.");

                    messages.Add(new DuckChatMessage(
                        "user",
                        "رفض مضيف البحث تنفيذ فتح الصفحة لأن المعرّف لم يثبت وروده في نتيجة بحث سابقة من المكتبة نفسها. " +
                        referenceError +
                        " أعد البحث عند الحاجة، ثم افتح مرجعاً مطابقاً تماماً لما أعادته أداة البحث."));

                    continue;
                }

                var execution = await ExecuteToolAsync(
                    descriptorForJson,
                    toolArguments,
                    warnings,
                    evidenceProvenance,
                    taskId,
                    timeout.Token);

                evidenceLedger.Add(descriptorForJson, execution.Content, execution.Succeeded);

                var completedSearch = execution.Succeeded &&
                    ResearchToolResultClassifier.IsCompletedSearch(descriptorForJson, isError: false);
                if (completedSearch)
                    searchedSourceIds.Add(descriptorForJson.Server);

                if (execution.Succeeded)
                    trace.Add(new ToolCallTrace(
                        descriptorForJson.Server,
                        descriptorForJson.ServerLabel,
                        descriptorForJson.Name));

                _log.Info(
                    $"Research tool result: task={taskId}; source={decision.Server}; tool={decision.Tool}; success={execution.Succeeded}; completedSearch={completedSearch}; searchCoverage={searchedSourceIds.Count}/{selectedServerIds.Count}; evidenceEntries={evidenceLedger.Entries.Count}.");
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
            "أخرج كائناً واحداً فقط: {\"type\":\"final\",\"answer\":\"نص الإجابة\"}. لا تضف شرحاً أو Markdown خارج JSON.";

        try
        {
            using var finalizationTimeout = CancellationTokenSource.CreateLinkedTokenSource(requestCancellation);
            finalizationTimeout.CancelAfter(TimeSpan.FromSeconds(30));

            var contextBudget = ResearchContextBudget.LoadDuckAi();
            var availableInputCharacters = contextBudget.GetAvailableInputCharacters(systemPrompt.Length);
            var evidenceBudget = contextBudget.GetEvidenceBudget(
                availableInputCharacters,
                evidenceLedger.EstimatedContextCharacters,
                MaxDuckAiEvidenceContextCharacters);
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
                    MaxDuckAiMessageCharacters,
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

            var completion = await _duckAi.CompleteAsync(
                model,
                requestMessages,
                functionTools: null,
                cancellationToken: finalizationTimeout.Token);

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

        var maxFallbackCharacters = ResearchContextBudget.LoadDuckAi().MaxContextCharacters;
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
        McpEvidenceProvenance evidenceProvenance,
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
            var textBlocks = result.Content
                .OfType<TextContentBlock>()
                .Select(block => block.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();
            var hasUsableContent = ResearchToolResultClassifier.HasUsableContent(result);
            var sourceReferenceCount = evidenceProvenance.RecordSearchResult(
                descriptor,
                result.StructuredContent,
                textBlocks,
                result.IsError == true,
                hasUsableContent);

            if (ResearchToolResultClassifier.IsSearchTool(descriptor) &&
                result.IsError != true &&
                hasUsableContent &&
                sourceReferenceCount == 0)
            {
                warnings.Add($"{descriptor.Server}/{descriptor.Name}: search response contained no recognizable page identifiers");
                _log.Warning(
                    $"MCP search result had no page identifiers: task={taskId}; source={descriptor.Server}; tool={descriptor.Name}.");
            }

            if (result.IsError == true)
            {
                warnings.Add($"{descriptor.Server}/{descriptor.Name}: MCP reported a tool execution error");
                evidenceProvenance.RecordVerifiedOpen(descriptor, succeeded: false);
                return new ToolExecutionResult(formattedResult, false);
            }

            if (!hasUsableContent)
            {
                if (ResearchToolResultClassifier.IsCompletedSearch(descriptor, result.IsError == true))
                {
                    // A successful search with zero matches completes the search attempt,
                    // but it does not create a page reference that can be opened or cited.
                    warnings.Add($"{descriptor.Server}/{descriptor.Name}: search completed without usable matches");
                    evidenceProvenance.RecordVerifiedOpen(descriptor, succeeded: false);
                    return new ToolExecutionResult(formattedResult, true);
                }

                warnings.Add($"{descriptor.Server}/{descriptor.Name}: MCP returned no usable content");
                evidenceProvenance.RecordVerifiedOpen(descriptor, succeeded: false);
                return new ToolExecutionResult(formattedResult, false);
            }

            if (McpEvidenceProvenance.IsReferenceReadTool(descriptor))
            {
                var openedPageError = evidenceProvenance.ValidateOpenedResult(
                    descriptor,
                    arguments,
                    result.StructuredContent,
                    textBlocks);
                if (openedPageError is not null)
                {
                    warnings.Add($"{descriptor.Server}/{descriptor.Name}: opened page identity did not match the requested result");
                    _log.Warning(
                        $"MCP opened page rejected: task={taskId}; source={descriptor.Server}; tool={descriptor.Name}; reason=response-reference-mismatch.");
                    evidenceProvenance.RecordVerifiedOpen(descriptor, succeeded: false);
                    return new ToolExecutionResult(openedPageError, false);
                }
            }

            evidenceProvenance.RecordVerifiedOpen(descriptor, succeeded: true);
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
        IReadOnlyList<string> warnings,
        out int toolCatalogCharacters)
    {
        // Duck2API drops response_format, so the actual read-only MCP input schemas
        // must be visible in the ordinary system prompt for model-directed page reads.
        var toolDirectory = tools.Select(tool => new
        {
            server = tool.Server,
            source = tool.ServerLabel,
            name = tool.Name,
            description = tool.Description,
            inputSchema = tool.InputSchema
        }).ToArray();
        var serializedTools = JsonSerializer.Serialize(toolDirectory, JsonDefaults.Web);
        toolCatalogCharacters = serializedTools.Length;

        var builder = new System.Text.StringBuilder(
            ResearchAgentInstructions.Build(selectedLabels, warnings));
        builder.AppendLine().AppendLine(
            "أدوات MCP المتاحة وأوصافها ومخطط inputSchema الفعلي لكل أداة. استخدم المخطط لتحديد أسماء المعاملات وأنواعها والقيم الإلزامية والافتراضية، ولا تخمّن أي معرّف أو معامل:");
        builder.AppendLine(serializedTools);
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
