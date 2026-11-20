using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;

namespace ResearchTools.McpHost.Services;

/// <summary>
/// Adapter for DuckAI's local OpenAI-compatible HTTP API. No cloud credentials are used.
/// </summary>
public sealed class DuckAiClient
{
    private static readonly string[] CandidateBaseAddresses =
    [
        "http://127.0.0.1:8080",
        "http://127.0.0.1:18080"
    ];

    private static readonly TimeSpan ModelCacheLifetime = TimeSpan.FromSeconds(15);
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
                    using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    timeout.CancelAfter(TimeSpan.FromSeconds(8));
                    using var response = await _http.GetAsync(
                        $"{baseAddress}/v1/models",
                        HttpCompletionOption.ResponseHeadersRead,
                        timeout.Token);
                    var body = await response.Content.ReadAsStringAsync(timeout.Token);

                    if (!response.IsSuccessStatusCode)
                    {
                        failures.Add($"{baseAddress} HTTP {(int)response.StatusCode}");
                        continue;
                    }

                    using var document = JsonDocument.Parse(body);
                    var models = ReadModels(document.RootElement);
                    if (models.Count == 0)
                    {
                        failures.Add($"{baseAddress} returned an empty model list");
                        continue;
                    }

                    var defaultModel = models.FirstOrDefault(model =>
                        model.Contains("luna", StringComparison.OrdinalIgnoreCase)) ?? models[0];

                    _cachedSnapshot = new DuckAiModelSnapshot(baseAddress, defaultModel, models);
                    _snapshotExpiresAt = DateTimeOffset.UtcNow.Add(ModelCacheLifetime);
                    _log.Info($"DuckAI endpoint selected: {baseAddress}; modelCount={models.Count}; default={defaultModel}.");
                    return _cachedSnapshot;
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    failures.Add($"{baseAddress}: {exception.Message}");
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

    public async Task<(string Text, string Model)> CompleteAsync(
        string requestedModel,
        IReadOnlyList<DuckChatMessage> messages,
        CancellationToken cancellationToken)
    {
        var snapshot = await GetModelsAsync(cancellationToken);
        var model = string.IsNullOrWhiteSpace(requestedModel)
            ? snapshot.DefaultModel
            : requestedModel.Trim();

        if (!snapshot.Models.Contains(model, StringComparer.Ordinal))
            throw new InvalidOperationException($"النموذج غير متاح محلياً: {model}");

        var payload = new
        {
            model,
            messages,
            temperature = 0.1,
            max_tokens = 5000,
            stream = false
        };

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(175));

        using var request = new HttpRequestMessage(
            HttpMethod.Post, $"{snapshot.BaseAddress}/v1/chat/completions")
        {
            Content = JsonContent.Create(payload, options: JsonDefaults.Web)
        };

        using var response = await _http.SendAsync(
            request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
        var responseBody = await response.Content.ReadAsStringAsync(timeout.Token);

        if (response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.MethodNotAllowed ||
            (int)response.StatusCode == 501)
        {
            return await CompleteWithResponsesApiAsync(
                snapshot, model, messages, cancellationToken);
        }

        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException(
                $"DuckAI HTTP {(int)response.StatusCode}: {Limit(responseBody, 1800)}",
                null, response.StatusCode);

        using var document = JsonDocument.Parse(responseBody);
        var text = ExtractAssistantText(document.RootElement);
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("DuckAI returned an empty chat response.");

        return (text.Trim(), model);
    }

    private async Task<(string Text, string Model)> CompleteWithResponsesApiAsync(
        DuckAiModelSnapshot snapshot,
        string model,
        IReadOnlyList<DuckChatMessage> messages,
        CancellationToken cancellationToken)
    {
        var payload = new
        {
            model,
            input = messages.Select(message => new
            {
                role = message.Role,
                content = message.Content
            }),
            stream = false,
            max_output_tokens = 5000
        };

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(175));
        using var response = await _http.PostAsJsonAsync(
            $"{snapshot.BaseAddress}/v1/responses", payload, JsonDefaults.Web, timeout.Token);
        var body = await response.Content.ReadAsStringAsync(timeout.Token);

        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException(
                $"واجهتا DuckAI للمحادثة غير متاحتين. HTTP {(int)response.StatusCode}: {Limit(body, 1800)}",
                null, response.StatusCode);

        using var document = JsonDocument.Parse(body);
        var text = ExtractAssistantText(document.RootElement);
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("DuckAI Responses returned no text.");

        return (text.Trim(), model);
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

    private static string ExtractAssistantText(JsonElement root)
    {
        if (TryProperty(root, "output_text", out var outputText) &&
            outputText.ValueKind == JsonValueKind.String)
            return outputText.GetString() ?? string.Empty;

        if (TryProperty(root, "choices", out var choices) && choices.ValueKind == JsonValueKind.Array &&
            choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (TryProperty(first, "message", out var message) &&
                TryProperty(message, "content", out var content))
                return ReadContent(content);
        }

        if (TryProperty(root, "output", out var output) && output.ValueKind == JsonValueKind.Array)
        {
            var parts = new List<string>();
            foreach (var item in output.EnumerateArray())
            {
                if (!TryProperty(item, "content", out var content) || content.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var block in content.EnumerateArray())
                {
                    if (TryProperty(block, "text", out var text) && text.ValueKind == JsonValueKind.String)
                        parts.Add(text.GetString() ?? string.Empty);
                }
            }

            return string.Join(Environment.NewLine, parts).Trim();
        }

        return string.Empty;
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

public sealed record DuckChatMessage(string Role, string Content);
