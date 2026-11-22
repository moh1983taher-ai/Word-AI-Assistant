using System.Net;
using System.Text;
using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Services;
using Xunit;

namespace ResearchTools.McpHost.Tests;

public sealed class LmStudioClientTests
{
    [Fact]
    public async Task DiscoversOnlyLanguageModelsAndPrefersLoadedToolCallingModel()
    {
        var handler = new StubHandler(request =>
            {
                Assert.Equal(HttpMethod.Get, request.Method);
                Assert.Equal("/api/v1/models", request.RequestUri!.AbsolutePath);
                return Json(HttpStatusCode.OK, """
                    {
                      "models": [
                        {
                          "type": "embedding",
                          "key": "nomic/embed",
                          "display_name": "Embedding",
                          "loaded_instances": []
                        },
                        {
                          "type": "llm",
                          "key": "small-chat",
                          "display_name": "Small Chat",
                          "capabilities": { "trained_for_tool_use": true },
                          "loaded_instances": []
                        },
                        {
                          "type": "llm",
                          "key": "research-tool-model",
                          "display_name": "Research Tool Model",
                          "capabilities": { "trained_for_tool_use": true },
                          "loaded_instances": [{ "id": "instance-1" }]
                        },
                        {
                          "type": "llm",
                          "key": "not-tool-trained",
                          "display_name": "No Tool Use",
                          "capabilities": { "trained_for_tool_use": false },
                          "loaded_instances": []
                        }
                      ]
                    }
                    """);
        });

        var client = new LmStudioClient(
                new StubHttpClientFactory(handler),
                TestOptions(),
                new HostLog());

        var snapshot = await client.GetModelsAsync(CancellationToken.None);

        Assert.Equal(3, snapshot.Models.Count);
        Assert.DoesNotContain(snapshot.Models, model => model.Id == "nomic/embed");
        Assert.Equal("research-tool-model", snapshot.DefaultModel);
        Assert.True(snapshot.Models.Single(model => model.Id == "research-tool-model").TrainedForToolUse);
    }

    [Fact]
    public async Task SendsSelectedMcpServersAndAllowListsAndCapturesStreamedToolResults()
    {
        string? capturedPayload = null;
        var handler = new StubHandler(request =>
            {
                if (request.Method == HttpMethod.Get && request.RequestUri.AbsolutePath == "/api/v1/models")
                {
                    return Json(HttpStatusCode.OK, """
                        {
                          "models": [{
                            "type": "llm",
                            "key": "research-model",
                            "display_name": "Research Model",
                            "capabilities": { "trained_for_tool_use": true },
                            "loaded_instances": [{ "id": "research-model" }]
                          }]
                        }
                        """);
                }

                Assert.Equal(HttpMethod.Post, request.Method);
                Assert.Equal("/api/v1/chat", request.RequestUri!.AbsolutePath);
                capturedPayload = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();

                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""
                        event: chat.start
                        data: {"type":"chat.start","model_instance_id":"research-model"}

                        event: tool_call.success
                        data: {"type":"tool_call.success","tool":"search","arguments":{"q":"أصول الفقه"},"output":"[نتائج الجامع]","provider_info":{"type":"ephemeral_mcp","server_label":"aljam3"}}

                        event: message.delta
                        data: {"type":"message.delta","content":"الجواب"}

                        event: chat.end
                        data: {"type":"chat.end","result":{"model_instance_id":"research-model","output":[{"type":"tool_call","tool":"search","output":"[نتائج الجامع]","provider_info":{"type":"ephemeral_mcp","server_label":"aljam3"}},{"type":"message","content":"الجواب النهائي"}],"response_id":"resp_test"}}

                        """, Encoding.UTF8, "text/event-stream")
                };
            });

            var client = new LmStudioClient(
                new StubHttpClientFactory(handler),
                TestOptions(),
                new HostLog());

        var captured = new List<LmStudioToolResult>();
        var result = await client.ChatAsync(
                "research-model",
                "تعليمات النظام",
                "سؤال البحث",
                [
                    new LmStudioIntegration("aljam3", "https://aljam3.com/mcp", ["search", "get_page"]),
                    new LmStudioIntegration("shamela", "https://mcp.shamela.ws", ["shamela_find"])
                ],
                tool => { captured.Add(tool); return Task.CompletedTask; },
                CancellationToken.None);

        Assert.Equal("الجواب النهائي", result.Answer);
        Assert.Equal("research-model", result.Model);
        Assert.Null(result.ResponseId);
        var tool = Assert.Single(captured);
        Assert.Equal("aljam3", tool.ServerLabel);
        Assert.Equal("search", tool.Tool);
        Assert.Equal("[نتائج الجامع]", tool.Output);

        Assert.NotNull(capturedPayload);
        using var payload = JsonDocument.Parse(capturedPayload!);
        Assert.Equal("research-model", payload.RootElement.GetProperty("model").GetString());
        Assert.True(payload.RootElement.GetProperty("stream").GetBoolean());
        Assert.False(payload.RootElement.GetProperty("store").GetBoolean());
        var integrations = payload.RootElement.GetProperty("integrations");
        Assert.Equal(2, integrations.GetArrayLength());

        var aljam3 = integrations.EnumerateArray().Single(item =>
                item.GetProperty("server_label").GetString() == "aljam3");
        Assert.Equal("https://aljam3.com/mcp", aljam3.GetProperty("server_url").GetString());
        Assert.Equal(
            new[] { "search", "get_page" },
                aljam3.GetProperty("allowed_tools").EnumerateArray()
                .Select(item => item.GetString()).ToArray());
    }

    [Fact]
    public async Task DoesNotAddMcpIntegrationsDuringTimeoutFinalization()
    {
        string? capturedPayload = null;
            var handler = new StubHandler(request =>
            {
                if (request.Method == HttpMethod.Get)
                {
                    return Json(HttpStatusCode.OK, """
                        {"models":[{"type":"llm","key":"model-1","display_name":"Model","capabilities":{"trained_for_tool_use":true},"loaded_instances":[]}]}
                        """);
                }

                capturedPayload = request.Content!.ReadAsStringAsync().GetAwaiter().GetResult();
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""
                        event: chat.end
                        data: {"type":"chat.end","result":{"model_instance_id":"model-1","output":[{"type":"message","content":"صياغة جزئية"}]}}

                        """, Encoding.UTF8, "text/event-stream")
                };
            });

            var client = new LmStudioClient(
                new StubHttpClientFactory(handler),
                TestOptions(),
                new HostLog());

        var result = await client.ChatAsync(
                "model-1",
                "لا تستخدم الأدوات",
                "سياق الأدلة",
                Array.Empty<LmStudioIntegration>(),
                onToolResult: null,
                CancellationToken.None);

        Assert.Equal("صياغة جزئية", result.Answer);
        Assert.NotNull(capturedPayload);
        using var payload = JsonDocument.Parse(capturedPayload!);
        Assert.False(payload.RootElement.TryGetProperty("integrations", out _));
    }

    private static ResearchProviderOptions TestOptions() =>
        new("lmstudio", new Uri("http://127.0.0.1:1234/"), null);

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
