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
/// Orchestrates Aurora-compatible tagged tool-call turns against official MCP clients.
/// Duck2API returns text rather than native function calls, so the host translates
/// the model's tagged response, validates it against MCP, and executes it locally.
/// </summary>
public sealed class ResearchCoordinator
{
    private const int MaxDuckAiMessageCharacters = 12_000;
    private const int MaxDuckAiEvidenceContextCharacters = 22_000;

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
                var unavailableLabels = unavailableSources
                    .Select(ResearchSourceSelection.GetLabel)
                    .ToArray();
                warnings.Add(
                    "تعذر الاتصال بأدوات البحث في المكتبات المختارة الآتية: " +
                    string.Join("، ", unavailableLabels) +
                    ". أُوقف البحث قبل التنفيذ حتى لا تُعرض تغطية جزئية بوصفها بحثاً في جميع المصادر.");

                var answer =
                    "تعذر بدء البحث في جميع المكتبات التي اخترتها. لم تتوفر أدوات بحث في: " +
                    string.Join("، ", unavailableLabels) +
                    ". لم يُنفذ بحث جزئي، ولا يصح استنتاج عدم وجود مصادر في هذه المكتبات.";
                return new ResearchResponse(answer, model, trace.ToArray(), warnings.ToArray());
            }

            if (activeTools.Count == 0)
            {
                var answer =
                    "تعذر بدء البحث؛ لم تتوفر أدوات قراءة مسموح بها في المكتبات التي اخترتها. " +
                    "لم يُنفذ بحث يمكن الاستناد إلى نتائجه، ولا يصح استنتاج عدم وجود مصادر للمسألة.";
                return new ResearchResponse(answer, model, trace.ToArray(), warnings.ToArray());
            }

            var selectedLabels = selectedServerIds
                .Select(ResearchSourceSelection.GetLabel)
                .ToArray();

            // Give the model the user's request directly. It chooses when and how to use MCP.
            var invalidDecisionAttempts = 0;
            var searchDecisionCorrectionAttempted = false;
            var pageVerificationCorrectionAttempted = false;
            var pageReadFallbackAttempted = false;
            var toolCallsAttempted = 0;
            var searchedSourceIds = new HashSet<string>(StringComparer.Ordinal);

            while (true)
            {
                IReadOnlyList<ToolDescriptor> turnTools = activeTools;
                var pendingSearchSources = ResearchIntentClassifier.GetPendingSearchSources(
                    latestUserMessage,
                    selectedServerIds,
                    searchedSourceIds);
                var pendingPageSources = evidenceProvenance.GetUnverifiedSearchSources();
                var allowFinal = ResearchFinalizationPolicy.CanFinalize(
                    pendingSearchSources.Count,
                    pendingPageSources.Count);
                var pendingSourceLabels = pendingSearchSources
                    .Select(ResearchSourceSelection.GetLabel)
                    .ToArray();
                var pendingPageSourceLabels = pendingPageSources
                    .Select(ResearchSourceSelection.GetLabel)
                    .ToArray();
                var turnSystemPrompt = BuildSystemPrompt(
                    turnTools,
                    selectedLabels,
                    warnings,
                    out var toolCatalogCharacters) +
                    Environment.NewLine + Environment.NewLine +
                    ResearchDecisionProtocol.BuildTurnInstructions(
                        allowFinal,
                        pendingSourceLabels,
                        pendingPageSourceLabels);

                // The compact tool directory is included as ordinary prompt text.
                // The original MCP schema remains authoritative during host-side validation.
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
                    $"Research prompt prepared: task={taskId}; selectedSources={string.Join(",", selectedServerIds)}; messages={requestMessages.Count}; systemCharacters={turnSystemPrompt.Length}; toolCatalogCharacters={toolCatalogCharacters}; responseFormat=omitted; pendingSearchSources={pendingSearchSources.Count}; pendingPageSources={pendingPageSources.Count}; finalAllowed={allowFinal}; historyCharacters={boundedContextCharacters}/{originalContextCharacters}; evidenceCharacters={evidenceContext.Length}; requestCharacters={requestCharacters}; fixedRequestCharacters={fixedRequestCharacters}; contextBudget={contextBudget.MaxContextCharacters}; maxOutputTokens={contextBudget.MaxOutputTokens}.");

                var completion = await _duckAi.CompleteAsync(
                    model,
                    requestMessages,
                    functionTools: null,
                    cancellationToken: timeout.Token);

                model = completion.Model;

                ModelDecision? parsedDecision = null;
                if (completion.ToolCalls.Count == 0 &&
                    (AuroraToolCallProtocol.TryParse(completion.Text, out var candidateDecision) ||
                     ModelDecisionParser.TryParse(completion.Text, out candidateDecision)))
                {
                    parsedDecision = candidateDecision;
                }

                if (parsedDecision is null)
                {
                    _log.Warning(
                        $"Research model decision invalid: task={taskId}; responseCharacters={completion.Text.Length}; toolCalls={completion.ToolCalls.Count}; responseFormat=omitted; protocol=aurora-tool-call+json-final; repairAttempt={invalidDecisionAttempts + 1}.");

                    if (invalidDecisionAttempts >= 2)
                    {
                        warnings.Add("تعذر متابعة البحث لأن النموذج لم ينتج قراراً صالحاً بعد محاولتي إصلاح؛ لم يعتمد المضيف إجابة نهائية غير متحققة.");
                        _log.Warning(
                            $"Research stopped with preserved evidence: task={taskId}; reason=invalid_model_decision; evidenceEntries={evidenceLedger.Entries.Count}; verifiedOpens={evidenceProvenance.VerifiedPageOpenCount}.");
                        return new ResearchResponse(
                            $"توقف البحث قبل اكتمال التحليل؛ جمع النظام {evidenceLedger.Entries.Count} مخرجاً من أدوات المكتبات، لكن النموذج لم ينتج قراراً صالحاً لمتابعة فتح المراجع وتحليلها. لم تُعتمد إجابة نهائية أو إحالات غير متحققة. أعد المحاولة أو استخدم نموذجاً آخر.",
                            model,
                            trace.ToArray(),
                            warnings.ToArray());
                    }

                    messages.Add(new DuckChatMessage("assistant", completion.Text));
                    messages.Add(new DuckChatMessage(
                        "user",
                        ResearchDecisionProtocol.BuildInvalidDecisionRepair(
                            allowFinal,
                            pendingSourceLabels,
                            pendingPageSourceLabels)));
                    invalidDecisionAttempts++;
                    continue;
                }

                invalidDecisionAttempts = 0;
                var decision = parsedDecision!;

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
                                $"Research premature final deferred: task={taskId}; reason=model_final_before_tool_call; recovery=host_exact_request_search; pendingSources={string.Join(",", pendingSearchSources)}.");

                            // Recovery only: if the model returns text instead of its first MCP
                            // tool call, search the exact user request in the still-pending
                            // selected sources, then return control to the model for page choice,
                            // evidence assessment, any further searches, and final writing.
                            messages.Add(new DuckChatMessage("assistant", decision.Raw));
                            if (string.IsNullOrWhiteSpace(latestUserMessage))
                                throw new InvalidOperationException(
                                    "تعذر بدء البحث الاحتياطي؛ لم يُعثر على نص سؤال الباحث.");

                            var recoveryPlan = ResearchSearchBootstrap.BuildInitialPlan(
                                activeTools,
                                pendingSearchSources,
                                latestUserMessage);

                            foreach (var invocation in recoveryPlan)
                            {
                                toolCallsAttempted++;
                                _log.Info(
                                    $"Research fallback search dispatch: task={taskId}; source={invocation.Tool.Server}; tool={invocation.Tool.Name}; querySource=user_request; queryCharacters={invocation.Query.Length}.");

                                var recoveryExecution = await ExecuteToolAsync(
                                    invocation.Tool,
                                    invocation.Arguments,
                                    warnings,
                                    evidenceProvenance,
                                    taskId,
                                    timeout.Token);

                                evidenceLedger.Add(invocation.Tool, recoveryExecution.Content, recoveryExecution.Succeeded);

                                var recoveryCompletedSearch = recoveryExecution.Succeeded &&
                                    ResearchToolResultClassifier.IsCompletedSearch(invocation.Tool, isError: false);
                                if (recoveryCompletedSearch)
                                    searchedSourceIds.Add(invocation.Tool.Server);

                                if (recoveryExecution.Succeeded)
                                    trace.Add(new ToolCallTrace(
                                        invocation.Tool.Server,
                                        invocation.Tool.ServerLabel,
                                        invocation.Tool.Name));

                                _log.Info(
                                    $"Research fallback search result: task={taskId}; source={invocation.Tool.Server}; tool={invocation.Tool.Name}; success={recoveryExecution.Succeeded}; completedSearch={recoveryCompletedSearch}; references={evidenceProvenance.SearchReferenceCount}; evidenceEntries={evidenceLedger.Entries.Count}.");
                            }

                            var stillPending = pendingSearchSources
                                .Where(source => !searchedSourceIds.Contains(source))
                                .ToArray();
                            if (stillPending.Length > 0)
                            {
                                _log.Warning(
                                    $"Research fallback search incomplete: task={taskId}; pendingSources={string.Join(",", stillPending)}.");
                                throw new InvalidOperationException(
                                    "تعذر تنفيذ البحث الأولي بنجاح في جميع المكتبات المختارة: " +
                                    string.Join("، ", stillPending.Select(ResearchSourceSelection.GetLabel)) +
                                    ". لم تُعتمد إجابة النموذج بوصفها بحثاً موثقاً.");
                            }

                            _log.Info(
                                $"Research fallback search completed: task={taskId}; sources={pendingSearchSources.Count}; querySource=user_request; handoff=model.");
                            continue;
                        }

                        _log.Warning(
                            $"Research final rejected: task={taskId}; reason=explicit_source_request_without_successful_tool_call; fallbackAttempted=true.");
                        throw new InvalidOperationException(
                            "تعذر إتمام البحث المطلوب: لم ينتج النموذج قراراً صالحاً لمتابعة البحث بعد تنفيذ البحث الاحتياطي. لم تُعرض إجابته بوصفها بحثاً موثقاً.");
                    }

                    if (matchingPageWasNotOpened)
                    {
                        if (!pageVerificationCorrectionAttempted)
                        {
                            pageVerificationCorrectionAttempted = true;
                            _log.Warning(
                                $"Research premature final deferred: task={taskId}; reason=search_results_without_verified_page_open; finalAllowed={allowFinal}; pendingPageSources={string.Join(",", pendingPageSources)}; searchResponses={evidenceProvenance.NonEmptySearchResponseCount}; searchReferences={evidenceProvenance.SearchReferenceCount}; verifiedOpens={evidenceProvenance.VerifiedPageOpenCount}; recovery=one_page_verification_correction.");

                            messages.Add(new DuckChatMessage("assistant", decision.Raw));
                            messages.Add(new DuckChatMessage(
                                "user",
                                "تنبيه تحقق: أعادت مكتبة واحدة أو أكثر نتائج بحث فعلية، لكن لم يُفتح موضع مطابق من كل مكتبة ظهرت منها نتائج. افتح صفحة مطابقة واحدة على الأقل من كل مكتبة ذات نتائج، مستخدماً معرّفات الكتاب والصفحة الواردة حرفياً في نتائج تلك المكتبة نفسها. افحص النص وسياقه قبل الاستشهاد. إذا لم يتوفر معرّف قابل للتحقق، فصرّح بتعذر فتح المصدر ولا تستشهد بمقتطف البحث."));
                            continue;
                        }

                        if (!pageReadFallbackAttempted)
                        {
                            pageReadFallbackAttempted = true;
                            _log.Warning(
                                $"Research page-open recovery started: task={taskId}; pendingPageSources={string.Join(",", pendingPageSources)}; recovery=first_verified_search_hit.");

                            foreach (var sourceId in pendingPageSources)
                            {
                                var openTool = activeTools.FirstOrDefault(tool =>
                                    tool.Server == sourceId &&
                                    (sourceId == "aljam3"
                                        ? tool.Name == "get_page"
                                        : tool.Name == "shamela_open"));
                                if (openTool is null ||
                                    !evidenceProvenance.TryBuildVerifiedReadArguments(openTool, out var openArguments))
                                {
                                    _log.Warning(
                                        $"Research page-open recovery unavailable: task={taskId}; source={sourceId}; reason=no_schema_valid_verified_reference.");
                                    continue;
                                }

                                var referenceError = evidenceProvenance.ValidateReferenceRead(openTool, openArguments);
                                if (referenceError is not null)
                                {
                                    _log.Warning(
                                        $"Research page-open recovery rejected: task={taskId}; source={sourceId}; reason=reference_validation_failed.");
                                    continue;
                                }

                                toolCallsAttempted++;
                                _log.Info(
                                    $"Research page-open recovery dispatch: task={taskId}; source={sourceId}; tool={openTool.Name}; reference=first_verified_search_hit.");

                                var openExecution = await ExecuteToolAsync(
                                    openTool,
                                    openArguments,
                                    warnings,
                                    evidenceProvenance,
                                    taskId,
                                    timeout.Token);
                                evidenceLedger.Add(openTool, openExecution.Content, openExecution.Succeeded);

                                if (openExecution.Succeeded)
                                    trace.Add(new ToolCallTrace(
                                        openTool.Server,
                                        openTool.ServerLabel,
                                        openTool.Name));

                                _log.Info(
                                    $"Research page-open recovery result: task={taskId}; source={sourceId}; tool={openTool.Name}; success={openExecution.Succeeded}; verifiedOpens={evidenceProvenance.VerifiedPageOpenCount}.");
                            }

                            var stillUnverifiedSources = evidenceProvenance.GetUnverifiedSearchSources();
                            if (stillUnverifiedSources.Count == 0)
                            {
                                messages.Add(new DuckChatMessage(
                                    "user",
                                    "فتح المضيف موضعاً أصلياً مطابقاً من نتائج كل مكتبة كانت تحتاج إلى تحقق. اقرأ النصوص المفتوحة وسياقها، واختر بنفسك ما إذا كان يلزم بحث أو فتح مراجع إضافية، ثم أجب عن سؤال الباحث كما طلبه."));
                                _log.Info(
                                    $"Research page-open recovery completed: task={taskId}; verifiedOpens={evidenceProvenance.VerifiedPageOpenCount}; handoff=model.");
                                continue;
                            }

                            _log.Warning(
                                $"Research page-open recovery incomplete: task={taskId}; pendingPageSources={string.Join(",", stillUnverifiedSources)}.");
                            throw new InvalidOperationException(
                                "تعذر فتح صفحة مطابقة من المكتبات الآتية رغم الاسترداد المحدود: " +
                                string.Join("، ", stillUnverifiedSources.Select(ResearchSourceSelection.GetLabel)) +
                                ". لم تُعتمد إجابة غير موثقة.");
                        }

                        _log.Warning(
                            $"Research final rejected: task={taskId}; reason=search_results_without_verified_page_open; finalAllowed={allowFinal}; fallbackAttempted=true.");
                        throw new InvalidOperationException(
                            "تعذر اعتماد إجابة البحث بعد محاولة فتح الصفحات الاحتياطية؛ لم تُمرر إجابة غير موثقة.");
                    }

                    _log.Info($"Research completed: task={taskId}; toolCalls={trace.Count}; evidenceEntries={evidenceLedger.Entries.Count}; mode=aurora-tool-call+json-final; model={model}.");
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
            ResearchDecisionProtocol.BuildTimeoutFinalInstructions();

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
        // Send a compact Aurora-style catalog instead of full JSON schemas.
        // The host still validates every call against the original MCP schema.
        var compactToolCatalog = AuroraToolCallProtocol.BuildToolCatalog(tools);
        toolCatalogCharacters = compactToolCatalog.Length;

        var builder = new System.Text.StringBuilder(
            ResearchAgentInstructions.Build(selectedLabels, warnings));
        builder.AppendLine().AppendLine(
            "أدوات MCP المتاحة بصيغة أسماء مركبة: اسم المكتبة ثم اسم الأداة. أرسل الاستدعاء في كتلة <tool_call>، ويظل المضيف مسؤولاً عن التحقق من المخطط الأصلي قبل التنفيذ:");
        builder.AppendLine(compactToolCatalog);
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
