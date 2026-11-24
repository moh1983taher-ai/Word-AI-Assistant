using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Adapter for DuckAI's local OpenAI-compatible HTTP API. The research coordinator
/// uses prompt-guided JSON and host-owned initial MCP searches because the bundled
/// Duck2API endpoint does not enforce response_format JSON Schema. Optional response
/// formats remain available to callers using a compatible endpoint.
/// </summary>
public sealed class DuckAiClient
{
    internal static readonly string[] CandidateBaseAddresses =
    [
        // Use the direct DuckAI inference backend first. The local agent proxy on
        // 8080 may answer model-list requests while stalling chat completions.
        "http://127.0.0.1:18080",
        "http://127.0.0.1:8080"
    ];

    // Keep the selected endpoint stable throughout a multi-turn MCP research task.
    private static readonly TimeSpan ModelCacheLifetime = TimeSpan.FromMinutes(5);
    private readonly HttpClient _http;
    private readonly HostLog _log;
    private readonly SemaphoreSlim _discoveryGate = new(1, 1);
    private DuckAiModelSnapshot? _cachedSnapshot;
    private DateTimeOffset _snapshotExpiresAt;

    public DuckAiClient(IHttpClientFactory factory, HostLog log)
    {
        _http = factory.CreateClient("DuckAI");
        _log = log;
    }

    public async Task<DuckAiModelSnapshot> GetModelsAsync(CancellationToken cancellationToken)
    {
        if (_cachedSnapshot is not null && _snapshotExpiresAt > DateTimeOffset.UtcNow)
            return _cachedSnapshot;

        await _discoveryGate.WaitAsync(cancellationToken);
        try
        {
            if (_cachedSnapshot is not null && _snapshotExpiresAt > DateTimeOffset.UtcNow)
                return _cachedSnapshot;

            var failures = new List<string>();
            foreach (var baseAddress in CandidateBaseAddresses)
            {
                try
                {
                    var snapshot = await DiscoverModelsAtEndpointAsync(baseAddress, cancellationToken);
                    CacheSnapshot(snapshot);
                    _log.Info($"DuckAI model endpoint discovered: {baseAddress}; modelCount={snapshot.Models.Count}; default={snapshot.DefaultModel}.");
                    return snapshot;
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    failures.Add($"{baseAddress}: {exception.GetType().Name}: {exception.Message}");
                }
            }

            throw new InvalidOperationException(
                "لم يمكن العثور على واجهة DuckAI المحلية: " + string.Join(" | ", failures));
        }
        finally
        {
            _discoveryGate.Release();
        }
    }

    private async Task<DuckAiModelSnapshot> DiscoverModelsAtEndpointAsync(
        string baseAddress,
        CancellationToken cancellationToken)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(8));

        using var response = await _http.GetAsync(
            $"{baseAddress}/v1/models",
            HttpCompletionOption.ResponseHeadersRead,
            timeout.Token);
        var body = await response.Content.ReadAsStringAsync(timeout.Token);

        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException(
                $"DuckAI model-list HTTP {(int)response.StatusCode} at {baseAddress}: {Limit(body, 500)}",
                null,
                response.StatusCode);

        using var document = JsonDocument.Parse(body);
        var models = ReadModels(document.RootElement);
        if (models.Count == 0)
            throw new InvalidOperationException($"{baseAddress} returned an empty model list.");

        var defaultModel = models.FirstOrDefault(model =>
            model.Contains("luna", StringComparison.OrdinalIgnoreCase)) ?? models[0];

        return new DuckAiModelSnapshot(baseAddress, defaultModel, models);
    }

    private void CacheSnapshot(DuckAiModelSnapshot snapshot)
    {
        _cachedSnapshot = snapshot;
        _snapshotExpiresAt = DateTimeOffset.UtcNow.Add(ModelCacheLifetime);
    }

    public async Task<DuckAiCompletion> CompleteAsync(
        string requestedModel,
        IReadOnlyList<DuckChatMessage> messages,
        IReadOnlyList<DuckFunctionTool>? functionTools,
        CancellationToken cancellationToken,
        JsonElement? responseFormat = null)
    {
        var preferredSnapshot = await GetModelsAsync(cancellationToken);
        var endpoints = new[] { preferredSnapshot.BaseAddress }
            .Concat(CandidateBaseAddresses.Where(address =>
                !string.Equals(address, preferredSnapshot.BaseAddress, StringComparison.OrdinalIgnoreCase)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var failures = new List<string>();
        foreach (var baseAddress in endpoints)
        {
            cancellationToken.ThrowIfCancellationRequested();

            DuckAiModelSnapshot snapshot;
            try
            {
                snapshot = string.Equals(
                    baseAddress, preferredSnapshot.BaseAddress, StringComparison.OrdinalIgnoreCase)
                    ? preferredSnapshot
                    : await DiscoverModelsAtEndpointAsync(baseAddress, cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception exception)
            {
                failures.Add($"{baseAddress} model discovery: {exception.GetType().Name}: {exception.Message}");
                _log.Warning($"DuckAI model discovery failed at {baseAddress}: {exception.GetType().Name}.");
                continue;
            }

            var model = string.IsNullOrWhiteSpace(requestedModel)
                ? snapshot.DefaultModel
                : requestedModel.Trim();

            if (!snapshot.Models.Contains(model, StringComparer.Ordinal))
            {
                failures.Add($"{baseAddress}: requested model is not listed ({model}).");
                continue;
            }

            try
            {
                var completion = await CompleteAtEndpointAsync(
                    snapshot, model, messages, functionTools, responseFormat, cancellationToken);

                CacheSnapshot(snapshot);
                if (!string.Equals(
                    snapshot.BaseAddress, preferredSnapshot.BaseAddress, StringComparison.OrdinalIgnoreCase))
                {
                    _log.Warning(
                        $"DuckAI inference endpoint switched from {preferredSnapshot.BaseAddress} to {snapshot.BaseAddress} after the preferred endpoint failed.");
                }

                return completion;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                failures.Add($"{baseAddress}: inference timed out after {ResearchContextBudget.LoadDuckAi().ModelRequestTimeout.TotalSeconds:0} seconds.");
                _log.Warning($"DuckAI inference timed out at {baseAddress}; trying the next local endpoint.");
            }
            catch (HttpRequestException exception) when (IsRetryableEndpointFailure(exception.StatusCode))
            {
                var status = exception.StatusCode is null ? "connection failure" : $"HTTP {(int)exception.StatusCode.Value}";
                failures.Add($"{baseAddress} {status}: {exception.Message}");
                _log.Warning($"DuckAI inference failed at {baseAddress} ({status}); trying the next local endpoint.");
            }
        }

        throw new HttpRequestException(
            "تعذر توليد الإجابة عبر واجهات DuckAI المحلية. جُرّبت جميع الواجهات المتاحة: " +
            string.Join(" | ", failures));
    }

    private async Task<DuckAiCompletion> CompleteAtEndpointAsync(
        DuckAiModelSnapshot snapshot,
        string model,
        IReadOnlyList<DuckChatMessage> messages,
        IReadOnlyList<DuckFunctionTool>? functionTools,
        JsonElement? responseFormat,
        CancellationToken cancellationToken,
        bool allowJsonObjectFallback = true)
    {
        var payload = new Dictionary<string, object?>
        {
            ["model"] = model,
            ["messages"] = messages,
            ["temperature"] = 0.0,
            ["max_tokens"] = ResearchContextBudget.LoadDuckAi().MaxOutputTokens,
            ["stream"] = false
        };

        var hasFunctionTools = functionTools is { Count: > 0 };
        if (hasFunctionTools)
        {
            payload["tools"] = functionTools;
            payload["tool_choice"] = "auto";
        }

        if (responseFormat is { } format)
            payload["response_format"] = format;

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(ResearchContextBudget.LoadDuckAi().ModelRequestTimeout);

        using var request = new HttpRequestMessage(
            HttpMethod.Post, $"{snapshot.BaseAddress}/v1/chat/completions")
        {
            Content = JsonContent.Create(payload, options: JsonDefaults.Web)
        };

        using var response = await _http.SendAsync(
            request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
        var responseBody = await response.Content.ReadAsStringAsync(timeout.Token);

        if (!response.IsSuccessStatusCode)
        {
            if (responseFormat is not null &&
                allowJsonObjectFallback &&
                (response.StatusCode is HttpStatusCode.BadRequest or HttpStatusCode.UnprocessableEntity) &&
                IsResponseFormatRejection(responseBody))
            {
                _log.Warning(
                    $"DuckAI endpoint {snapshot.BaseAddress} rejected strict JSON Schema output; retrying this turn with JSON-object mode.");

                using var jsonObjectFormat = JsonDocument.Parse("""{"type":"json_object"}""");
                return await CompleteAtEndpointAsync(
                    snapshot,
                    model,
                    messages,
                    functionTools,
                    jsonObjectFormat.RootElement.Clone(),
                    cancellationToken,
                    allowJsonObjectFallback: false);
            }

            if (hasFunctionTools && IsFunctionCallingUnsupportedStatus(response.StatusCode))
            {
                throw new NativeToolCallingNotSupportedException(
                    $"DuckAI rejected native function tools (HTTP {(int)response.StatusCode}).");
            }

            if (!hasFunctionTools &&
                (response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.MethodNotAllowed ||
                 (int)response.StatusCode == 501))
            {
                return await CompleteWithResponsesApiAsync(
                    snapshot, model, messages, responseFormat, cancellationToken);
            }

            throw new HttpRequestException(
                $"DuckAI HTTP {(int)response.StatusCode} at {snapshot.BaseAddress}: {Limit(responseBody, 1800)}",
                null, response.StatusCode);
        }

        using var document = JsonDocument.Parse(responseBody);
        return ParseChatCompletion(document.RootElement, model);
    }

    private static bool IsRetryableEndpointFailure(HttpStatusCode? statusCode)
    {
        if (statusCode is null) return true;

        return statusCode is HttpStatusCode.BadRequest
            or HttpStatusCode.UnprocessableEntity
            or HttpStatusCode.RequestTimeout
            or HttpStatusCode.InternalServerError
            or HttpStatusCode.BadGateway
            or HttpStatusCode.ServiceUnavailable
            or HttpStatusCode.GatewayTimeout
            or HttpStatusCode.NotFound
            or HttpStatusCode.MethodNotAllowed
            || (int)statusCode.Value == 501;
    }

    private async Task<DuckAiCompletion> CompleteWithResponsesApiAsync(
        DuckAiModelSnapshot snapshot,
        string model,
        IReadOnlyList<DuckChatMessage> messages,
        JsonElement? responseFormat,
        CancellationToken cancellationToken)
    {
        var payload = new Dictionary<string, object?>
        {
            ["model"] = model,
            ["input"] = messages.Select(message => new
            {
                role = message.Role,
                content = message.Content
            }),
            ["stream"] = false,
            ["max_output_tokens"] = 5000
        };

        if (responseFormat is { } chatFormat &&
            TryProperty(chatFormat, "type", out var formatType) &&
            formatType.ValueKind == JsonValueKind.String)
        {
            if (string.Equals(formatType.GetString(), "json_schema", StringComparison.Ordinal) &&
                TryProperty(chatFormat, "json_schema", out var jsonSchema))
            {
                payload["text"] = new
                {
                    format = new
                    {
                        type = "json_schema",
                        name = ReadString(jsonSchema, "name"),
                        strict = TryProperty(jsonSchema, "strict", out var strict) &&
                                 strict.ValueKind == JsonValueKind.True,
                        schema = TryProperty(jsonSchema, "schema", out var schema)
                            ? schema
                            : default
                    }
                };
            }
            else if (string.Equals(formatType.GetString(), "json_object", StringComparison.Ordinal))
            {
                payload["text"] = new { format = new { type = "json_object" } };
            }
        }

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(55));
        using var response = await _http.PostAsJsonAsync(
            $"{snapshot.BaseAddress}/v1/responses", payload, JsonDefaults.Web, timeout.Token);
        var body = await response.Content.ReadAsStringAsync(timeout.Token);

        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException(
                $"واجهتا DuckAI للمحادثة غير متاحتين. HTTP {(int)response.StatusCode}: {Limit(body, 1800)}",
                null, response.StatusCode);

        using var document = JsonDocument.Parse(body);
        var text = ExtractResponsesText(document.RootElement);
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("DuckAI Responses returned no text.");

        return new DuckAiCompletion(text.Trim(), [], model);
    }

    private static bool IsFunctionCallingUnsupportedStatus(HttpStatusCode statusCode) =>
        statusCode is HttpStatusCode.BadRequest or HttpStatusCode.UnprocessableEntity or
            HttpStatusCode.NotImplemented or HttpStatusCode.NotFound or
            HttpStatusCode.MethodNotAllowed ||
        (int)statusCode == 501;

    private static bool IsResponseFormatRejection(string responseBody)
    {
        var body = responseBody.ToLowerInvariant();
        return body.Contains("response_format", StringComparison.Ordinal) ||
               body.Contains("json_schema", StringComparison.Ordinal) ||
               body.Contains("structured output", StringComparison.Ordinal) ||
               body.Contains("schema", StringComparison.Ordinal);
    }

    internal static DuckAiCompletion ParseChatCompletion(JsonElement root, string model)
    {
        var choices = TryProperty(root, "choices", out var choiceRows) &&
                      choiceRows.ValueKind == JsonValueKind.Array
            ? choiceRows
            : default;

        if (choices.ValueKind != JsonValueKind.Array || choices.GetArrayLength() == 0)
        {
            var text = ExtractResponsesText(root);
            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException("DuckAI returned no assistant message.");
            return new DuckAiCompletion(text.Trim(), [], model);
        }

        var first = choices[0];
        if (!TryProperty(first, "message", out var message))
            throw new InvalidOperationException("DuckAI chat response omitted choices[0].message.");

        var content = TryProperty(message, "content", out var contentElement)
            ? ReadContent(contentElement)
            : string.Empty;

        var calls = new List<DuckFunctionCall>();
        if (TryProperty(message, "tool_calls", out var callRows) &&
            callRows.ValueKind == JsonValueKind.Array)
        {
            foreach (var call in callRows.EnumerateArray())
            {
                var id = ReadString(call, "id");
                if (!TryProperty(call, "function", out var function))
                    continue;

                var name = ReadString(function, "name");
                var arguments = ReadRawArgumentString(function);
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name))
                    continue;

                calls.Add(new DuckFunctionCall(id, name, arguments));
            }
        }

        return new DuckAiCompletion(content.Trim(), calls, model);
    }

    private static string ReadRawArgumentString(JsonElement function)
    {
        if (!TryProperty(function, "arguments", out var arguments))
            return "{}";

        return arguments.ValueKind == JsonValueKind.String
            ? arguments.GetString() ?? "{}"
            : arguments.GetRawText();
    }

    private static string ExtractResponsesText(JsonElement root)
    {
        if (TryProperty(root, "output_text", out var outputText) &&
            outputText.ValueKind == JsonValueKind.String)
            return outputText.GetString() ?? string.Empty;

        if (!TryProperty(root, "output", out var output) || output.ValueKind != JsonValueKind.Array)
            return string.Empty;

        var parts = new List<string>();
        foreach (var item in output.EnumerateArray())
        {
            if (!TryProperty(item, "content", out var content) ||
                content.ValueKind != JsonValueKind.Array)
                continue;

            foreach (var block in content.EnumerateArray())
            {
                if (TryProperty(block, "text", out var text) &&
                    text.ValueKind == JsonValueKind.String)
                    parts.Add(text.GetString() ?? string.Empty);
            }
        }

        return string.Join(Environment.NewLine, parts).Trim();
    }

    private static IReadOnlyList<string> ReadModels(JsonElement root)
    {
        var rows = TryProperty(root, "data", out var data) && data.ValueKind == JsonValueKind.Array
            ? data
            : TryProperty(root, "models", out var fallback) && fallback.ValueKind == JsonValueKind.Array
                ? fallback
                : default;

        if (rows.ValueKind != JsonValueKind.Array)
            return [];

        var models = new List<string>();
        foreach (var row in rows.EnumerateArray())
        {
            if (row.ValueKind == JsonValueKind.String)
            {
                var value = row.GetString();
                if (!string.IsNullOrWhiteSpace(value)) models.Add(value);
            }
            else if (TryProperty(row, "id", out var id) && id.ValueKind == JsonValueKind.String)
            {
                var value = id.GetString();
                if (!string.IsNullOrWhiteSpace(value)) models.Add(value);
            }
        }

        return models.Distinct(StringComparer.Ordinal).ToArray();
    }

    private static string ReadContent(JsonElement content)
    {
        if (content.ValueKind == JsonValueKind.String)
            return content.GetString() ?? string.Empty;

        if (content.ValueKind == JsonValueKind.Array)
        {
            var parts = new List<string>();
            foreach (var block in content.EnumerateArray())
            {
                if (TryProperty(block, "text", out var text) && text.ValueKind == JsonValueKind.String)
                    parts.Add(text.GetString() ?? string.Empty);
            }
            return string.Join(Environment.NewLine, parts).Trim();
        }

        return string.Empty;
    }

    private static string ReadString(JsonElement element, string name) =>
        TryProperty(element, name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static bool TryProperty(JsonElement element, string name, out JsonElement value)
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

    private static string Limit(string value, int maximum) =>
        value.Length <= maximum ? value : value[..maximum] + " [truncated]";
}

public sealed class NativeToolCallingNotSupportedException(string message) : Exception(message);

public sealed record DuckChatMessage(
    [property: JsonPropertyName("role")] string Role,
    [property: JsonPropertyName("content")] string? Content,
    [property: JsonPropertyName("tool_call_id"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    string? ToolCallId = null,
    [property: JsonPropertyName("tool_calls"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    IReadOnlyList<DuckChatToolCall>? ToolCalls = null);

public sealed record DuckChatToolCall(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("function")] DuckChatFunction Function);

public sealed record DuckChatFunction(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("arguments")] string Arguments);

public sealed record DuckFunctionTool(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("function")] DuckFunctionDefinition Function);

public sealed record DuckFunctionDefinition(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("description")] string Description,
    [property: JsonPropertyName("parameters")] JsonElement Parameters);

public sealed record DuckFunctionCall(string Id, string Name, string Arguments);

public sealed record DuckAiCompletion(
    string Text,
    IReadOnlyList<DuckFunctionCall> ToolCalls,
    string Model);
