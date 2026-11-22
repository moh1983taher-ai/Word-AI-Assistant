using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;

namespace ResearchTools.McpHost.Services;

public sealed record LmStudioModelInfo(
    string Id,
    string DisplayName,
    bool TrainedForToolUse,
    bool IsLoaded);

public sealed record LmStudioModelSnapshot(
    IReadOnlyList<LmStudioModelInfo> Models,
    string DefaultModel);

public sealed record LmStudioIntegration(
    string ServerLabel,
    string ServerUrl,
    IReadOnlyList<string> AllowedTools);

public sealed record LmStudioToolResult(
    string ServerLabel,
    string Tool,
    string Output,
    bool Succeeded);

public sealed record LmStudioChatResult(
    string Model,
    string Answer,
    string? ResponseId,
    IReadOnlyList<string> Warnings);

/// <summary>
/// Client for LM Studio's native v1 API. LM Studio owns the MCP agent loop: it
/// discovers the explicitly allow-listed tools, executes model-selected calls,
/// returns their outputs to the model, and streams the final aggregated response.
/// </summary>
public sealed class LmStudioClient
{
    private readonly HttpClient _http;
    private readonly ResearchProviderOptions _options;
    private readonly HostLog _log;
    private readonly SemaphoreSlim _modelsGate = new(1, 1);
    private LmStudioModelSnapshot? _cachedModels;
    private DateTimeOffset _modelsExpireAt;

    public LmStudioClient(
        IHttpClientFactory factory,
        ResearchProviderOptions options,
        HostLog log)
    {
        _http = factory.CreateClient("LMStudio");
        _options = options;
        _log = log;
    }

    public Uri BaseUri => _options.LmStudioBaseUri;

    public async Task<LmStudioModelSnapshot> GetModelsAsync(CancellationToken cancellationToken)
    {
        if (_cachedModels is not null && _modelsExpireAt > DateTimeOffset.UtcNow)
            return _cachedModels;

        await _modelsGate.WaitAsync(cancellationToken);
        try
        {
            if (_cachedModels is not null && _modelsExpireAt > DateTimeOffset.UtcNow)
                return _cachedModels;

            using var request = CreateRequest(HttpMethod.Get, "api/v1/models");
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(15));

            using var response = await _http.SendAsync(
                request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
            var body = await response.Content.ReadAsStringAsync(timeout.Token);
            EnsureSuccess(response, body, "model discovery");

            using var document = JsonDocument.Parse(body);
            if (!TryGet(document.RootElement, "models", out var modelRows) ||
                modelRows.ValueKind != JsonValueKind.Array)
            {
                throw new InvalidOperationException("LM Studio returned an invalid models catalog.");
            }

            var models = new List<LmStudioModelInfo>();
            foreach (var row in modelRows.EnumerateArray())
            {
                if (!string.Equals(ReadString(row, "type"), "llm", StringComparison.OrdinalIgnoreCase))
                    continue;

                var id = ReadString(row, "key");
                if (string.IsNullOrWhiteSpace(id))
                    continue;

                var displayName = ReadString(row, "display_name");
                var trainedForToolUse =
                    TryGet(row, "capabilities", out var capabilities) &&
                    TryGet(capabilities, "trained_for_tool_use", out var toolUse) &&
                    toolUse.ValueKind == JsonValueKind.True;
                var isLoaded =
                    TryGet(row, "loaded_instances", out var loaded) &&
                    loaded.ValueKind == JsonValueKind.Array &&
                    loaded.GetArrayLength() > 0;

                models.Add(new LmStudioModelInfo(
                    id,
                    string.IsNullOrWhiteSpace(displayName) ? id : displayName,
                    trainedForToolUse,
                    isLoaded));
            }

            if (models.Count == 0)
                throw new InvalidOperationException("LM Studio is reachable, but no language models are available.");

            var defaultModel = models
                .OrderByDescending(model => model.IsLoaded && model.TrainedForToolUse)
                .ThenByDescending(model => model.TrainedForToolUse)
                .ThenByDescending(model => model.IsLoaded)
                .First()
                .Id;

            _cachedModels = new LmStudioModelSnapshot(models, defaultModel);
            _modelsExpireAt = DateTimeOffset.UtcNow.AddSeconds(30);
            return _cachedModels;
        }
        finally
        {
            _modelsGate.Release();
        }
    }

    public async Task<LmStudioChatResult> ChatAsync(
        string requestedModel,
        string systemPrompt,
        string input,
        IReadOnlyList<LmStudioIntegration> integrations,
        Func<LmStudioToolResult, Task>? onToolResult,
        CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(systemPrompt);
        ArgumentException.ThrowIfNullOrWhiteSpace(input);

        var snapshot = await GetModelsAsync(cancellationToken);
        var model = string.IsNullOrWhiteSpace(requestedModel)
            ? snapshot.DefaultModel
            : requestedModel.Trim();

        if (!snapshot.Models.Any(candidate => string.Equals(candidate.Id, model, StringComparison.Ordinal)))
            throw new InvalidOperationException(
                $"The selected LM Studio model '{model}' is not available. Refresh the model list and select an installed language model.");

        var payload = new Dictionary<string, object?>
        {
            ["model"] = model,
            ["input"] = input,
            ["system_prompt"] = systemPrompt,
            ["temperature"] = 0.0,
            ["max_output_tokens"] = 5000,
            ["stream"] = true,
            ["store"] = false
        };

        if (integrations.Count > 0)
        {
            payload["integrations"] = integrations.Select(integration => new
            {
                type = "ephemeral_mcp",
                server_label = integration.ServerLabel,
                server_url = integration.ServerUrl,
                allowed_tools = integration.AllowedTools
            }).ToArray();
        }

        using var request = CreateRequest(HttpMethod.Post, "api/v1/chat");
        request.Content = JsonContent.Create(payload, options: JsonDefaults.Web);

        using var response = await _http.SendAsync(
            request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync(cancellationToken);
            throw new HttpRequestException(
                $"LM Studio chat failed during MCP-enabled inference: HTTP {(int)response.StatusCode}: {Limit(errorBody, 1800)}",
                null,
                response.StatusCode);
        }

        var toolResults = new List<LmStudioToolResult>();
        var warnings = new List<string>();
        var streamedMessage = new StringBuilder();
        JsonElement? finalResult = null;
        string? eventName = null;
        var dataLines = new List<string>();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var reader = new StreamReader(stream, Encoding.UTF8);

        async Task DispatchEventAsync()
        {
            if (dataLines.Count == 0)
            {
                eventName = null;
                return;
            }

            var data = string.Join("\n", dataLines);
            dataLines.Clear();
            var currentEvent = eventName;
            eventName = null;

            using var eventDocument = JsonDocument.Parse(data);
            var root = eventDocument.RootElement;
            var type = ReadString(root, "type");
            if (string.IsNullOrWhiteSpace(type))
                type = currentEvent ?? string.Empty;

            switch (type)
            {
                case "message.delta":
                    var delta = ReadString(root, "content");
                    if (!string.IsNullOrEmpty(delta))
                        streamedMessage.Append(delta);
                    break;

                case "tool_call.success":
                    await CaptureToolResultAsync(root, succeeded: true);
                    break;

                case "tool_call.failure":
                    var reason = ReadString(root, "reason");
                    if (!string.IsNullOrWhiteSpace(reason))
                        warnings.Add("أخفق استدعاء إحدى أدوات MCP في LM Studio: " + Limit(reason, 500));
                    break;

                case "error":
                    if (TryGet(root, "error", out var error))
                    {
                        var errorMessage = ReadString(error, "message");
                        if (!string.IsNullOrWhiteSpace(errorMessage))
                            warnings.Add("خطأ من LM Studio: " + Limit(errorMessage, 500));
                    }
                    break;

                case "chat.end":
                    if (TryGet(root, "result", out var result) && result.ValueKind == JsonValueKind.Object)
                        finalResult = result.Clone();
                    break;
            }
        }

        async Task CaptureToolResultAsync(JsonElement row, bool succeeded)
        {
            var tool = ReadString(row, "tool");
            var output = ReadString(row, "output");
            var serverLabel = string.Empty;

            if (TryGet(row, "provider_info", out var providerInfo))
                serverLabel = ReadString(providerInfo, "server_label");

            if (string.IsNullOrWhiteSpace(tool) || string.IsNullOrWhiteSpace(serverLabel))
            {
                warnings.Add("أعاد LM Studio نتيجة أداة من دون هوية مكتبة أو اسم أداة صالح؛ لم تُضف إلى سجل الأدلة.");
                return;
            }

            var key = serverLabel + "\n" + tool + "\n" + output;
            if (toolResults.Any(existing =>
                string.Equals(existing.ServerLabel + "\n" + existing.Tool + "\n" + existing.Output, key, StringComparison.Ordinal)))
                return;

            var result = new LmStudioToolResult(serverLabel, tool, output, succeeded);
            toolResults.Add(result);
            if (onToolResult is not null)
                await onToolResult(result);
        }

        while (true)
        {
            var line = await reader.ReadLineAsync(cancellationToken);
            if (line is null)
                break;

            if (line.Length == 0)
            {
                await DispatchEventAsync();
                continue;
            }

            if (line.StartsWith("event:", StringComparison.Ordinal))
                eventName = line[6..].Trim();
            else if (line.StartsWith("data:", StringComparison.Ordinal))
                dataLines.Add(line[5..].TrimStart());
        }

        await DispatchEventAsync();

        if (finalResult is null)
            throw new InvalidOperationException(
                "انتهى بث LM Studio دون حدث chat.end؛ لا يمكن التحقق من اكتمال الإجابة.");

        var final = finalResult.Value;
        var finalModel = ReadString(final, "model_instance_id");
        if (string.IsNullOrWhiteSpace(finalModel))
            finalModel = model;

        var answer = ExtractFinalMessage(final);
        if (string.IsNullOrWhiteSpace(answer))
            answer = streamedMessage.ToString().Trim();

        return new LmStudioChatResult(
            finalModel,
            answer,
            ReadString(final, "response_id"),
            warnings);
    }

    private HttpRequestMessage CreateRequest(HttpMethod method, string relativePath)
    {
        var request = new HttpRequestMessage(method, new Uri(_options.LmStudioBaseUri, relativePath));
        if (!string.IsNullOrWhiteSpace(_options.LmStudioApiToken))
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.LmStudioApiToken);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (method == HttpMethod.Post)
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
        return request;
    }

    private static void EnsureSuccess(HttpResponseMessage response, string body, string operation)
    {
        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException(
                $"LM Studio {operation} failed: HTTP {(int)response.StatusCode}: {Limit(body, 1200)}",
                null,
                response.StatusCode);
    }

    private static string ExtractFinalMessage(JsonElement result)
    {
        if (!TryGet(result, "output", out var output) || output.ValueKind != JsonValueKind.Array)
            return string.Empty;

        var messages = new List<string>();
        foreach (var item in output.EnumerateArray())
        {
            if (!string.Equals(ReadString(item, "type"), "message", StringComparison.Ordinal))
                continue;

            var content = ReadString(item, "content");
            if (!string.IsNullOrWhiteSpace(content))
                messages.Add(content);
        }

        return messages.Count == 0 ? string.Empty : messages[^1].Trim();
    }

    private static string ReadString(JsonElement element, string name) =>
        TryGet(element, name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static bool TryGet(JsonElement element, string name, out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string Limit(string? value, int maximum)
    {
        var text = value ?? string.Empty;
        return text.Length <= maximum ? text : text[..maximum] + " [truncated]";
    }
}
