using System.Collections.Concurrent;
using System.Net;
using System.Text;
using Microsoft.Extensions.Http;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class DuckAiEndpointSelectionTests
{
    [Fact]
    public void PrefersDirectBackendBeforeLocalAgentProxyForInitialDiscovery()
    {
        Assert.Equal("http://127.0.0.1:18080", DuckAiClient.CandidateBaseAddresses[0]);
        Assert.Equal("http://127.0.0.1:8080", DuckAiClient.CandidateBaseAddresses[1]);
    }

    [Fact]
    public async Task RetriesGenerationOnAlternateEndpointAfterBadGateway()
    {
        var visited = new ConcurrentQueue<string>();
        var handler = new StubHandler(request =>
        {
            var address = request.RequestUri!.GetLeftPart(UriPartial.Authority);
            visited.Enqueue($"{request.Method} {address}{request.RequestUri.AbsolutePath}");

            if (request.Method == HttpMethod.Get && request.RequestUri.AbsolutePath == "/v1/models")
                return Json(HttpStatusCode.OK, """{"data":[{"id":"luna-test"}]}""");

            if (request.Method == HttpMethod.Post &&
                address == "http://127.0.0.1:18080" &&
                request.RequestUri.AbsolutePath == "/v1/chat/completions")
                return Json(HttpStatusCode.BadGateway, """{"error":"Duck.ai request failed after retries"}""");

            if (request.Method == HttpMethod.Post &&
                address == "http://127.0.0.1:8080" &&
                request.RequestUri.AbsolutePath == "/v1/chat/completions")
            {
                return Json(
                    HttpStatusCode.OK,
                    """{"choices":[{"message":{"content":"{\"type\":\"final\",\"answer\":\"نجح التحويل إلى المنفذ البديل\"}"}}]}""");
            }

            return Json(HttpStatusCode.NotFound, """{"error":"not found"}""");
        });

        var client = new DuckAiClient(new StubHttpClientFactory(handler), new HostLog());
        var completion = await client.CompleteAsync(
            requestedModel: "",
            messages: [new DuckChatMessage("user", "اختبار توليد")],
            functionTools: null,
            cancellationToken: CancellationToken.None);

        Assert.Contains("نجح التحويل إلى المنفذ البديل", completion.Text);
        Assert.Equal(
            new[]
            {
                "GET http://127.0.0.1:18080/v1/models",
                "POST http://127.0.0.1:18080/v1/chat/completions",
                "GET http://127.0.0.1:8080/v1/models",
                "POST http://127.0.0.1:8080/v1/chat/completions"
            },
            visited.ToArray());

        var selected = await client.GetModelsAsync(CancellationToken.None);
        Assert.Equal("http://127.0.0.1:8080", selected.BaseAddress);
    }

    [Fact]
    public async Task SendsStrictJsonSchemaAndDeterministicTemperature()
    {
        string? capturedPayload = null;
        var handler = new StubHandler(request =>
        {
            if (request.Method == HttpMethod.Get &&
                request.RequestUri!.AbsolutePath == "/v1/models")
                return Json(HttpStatusCode.OK, """{"data":[{"id":"luna-test"}]}""");

            if (request.Method == HttpMethod.Post &&
                request.RequestUri!.AbsolutePath == "/v1/chat/completions")
            {
                capturedPayload = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();
                return Json(
                    HttpStatusCode.OK,
                    """{"choices":[{"message":{"content":"{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"arguments\":{\"q\":\"أصول الفقه\"}}"}}]}""");
            }

            return Json(HttpStatusCode.NotFound, """{"error":"not found"}""");
        });

        using var formatDocument = System.Text.Json.JsonDocument.Parse("""
            {
              "type": "json_schema",
              "json_schema": {
                "name": "research_decision",
                "strict": true,
                "schema": {
                  "type": "object",
                  "properties": {
                    "type": { "type": "string", "enum": ["tool_call"] }
                  },
                  "required": ["type"],
                  "additionalProperties": false
                }
              }
            }
            """);

        var client = new DuckAiClient(new StubHttpClientFactory(handler), new HostLog());
        var completion = await client.CompleteAsync(
            requestedModel: "",
            messages: [new DuckChatMessage("user", "اختبار المخطط الصارم")],
            functionTools: null,
            cancellationToken: CancellationToken.None,
            responseFormat: formatDocument.RootElement.Clone());

        Assert.Contains("tool_call", completion.Text);
        Assert.NotNull(capturedPayload);

        using var payload = System.Text.Json.JsonDocument.Parse(capturedPayload!);
        Assert.Equal(0, payload.RootElement.GetProperty("temperature").GetDouble());
        Assert.Equal(5_000, payload.RootElement.GetProperty("max_tokens").GetInt32());
        var responseFormat = payload.RootElement.GetProperty("response_format");
        Assert.Equal("json_schema", responseFormat.GetProperty("type").GetString());
        Assert.True(responseFormat.GetProperty("json_schema").GetProperty("strict").GetBoolean());
    }

    [Fact]
    public async Task FallsBackToJsonObjectOnlyWhenEndpointRejectsStrictSchema()
    {
        var receivedFormats = new ConcurrentQueue<string>();
        var handler = new StubHandler(request =>
        {
            if (request.Method == HttpMethod.Get &&
                request.RequestUri!.AbsolutePath == "/v1/models")
                return Json(HttpStatusCode.OK, """{"data":[{"id":"luna-test"}]}""");

            if (request.Method == HttpMethod.Post &&
                request.RequestUri!.AbsolutePath == "/v1/chat/completions")
            {
                var body = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();
                using var document = System.Text.Json.JsonDocument.Parse(body);
                var formatType = document.RootElement
                    .GetProperty("response_format")
                    .GetProperty("type")
                    .GetString()!;
                receivedFormats.Enqueue(formatType);

                if (formatType == "json_schema")
                    return Json(
                        HttpStatusCode.BadRequest,
                        """{"error":"response_format json_schema is not supported"}""");

                return Json(
                    HttpStatusCode.OK,
                    """{"choices":[{"message":{"content":"{\"type\":\"tool_call\",\"server\":\"aljam3\",\"tool\":\"search\",\"arguments\":{\"q\":\"الأمر\"}}"}}]}""");
            }

            return Json(HttpStatusCode.NotFound, """{"error":"not found"}""");
        });

        using var formatDocument = System.Text.Json.JsonDocument.Parse("""
            {
              "type": "json_schema",
              "json_schema": {
                "name": "research_decision",
                "strict": true,
                "schema": {
                  "type": "object",
                  "properties": { "type": { "type": "string" } },
                  "required": ["type"],
                  "additionalProperties": false
                }
              }
            }
            """);

        var client = new DuckAiClient(new StubHttpClientFactory(handler), new HostLog());
        var completion = await client.CompleteAsync(
            requestedModel: "",
            messages: [new DuckChatMessage("user", "اختبار الاحتياط")],
            functionTools: null,
            cancellationToken: CancellationToken.None,
            responseFormat: formatDocument.RootElement.Clone());

        Assert.Contains("tool_call", completion.Text);
        Assert.Equal(new[] { "json_schema", "json_object" }, receivedFormats.ToArray());
    }

    private static HttpResponseMessage Json(HttpStatusCode status, string body) =>
        new(status)
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json")
        };

    private sealed class StubHttpClientFactory(HttpMessageHandler handler) : IHttpClientFactory
    {
        public HttpClient CreateClient(string name) => new(handler, disposeHandler: false);
    }

    private sealed class StubHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
        : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) =>
            Task.FromResult(responder(request));
    }
}
