using System.Net;
using System.Text.Json;
using ResearchTools.McpHost.Infrastructure;
using ResearchTools.McpHost.Models;
using ResearchTools.McpHost.Security;
using ResearchTools.McpHost.Services;

var builder = WebApplication.CreateBuilder(args);

var configuredPort = Environment.GetEnvironmentVariable("RESEARCHTOOLS_MCP_PORT");
var port = int.TryParse(configuredPort, out var parsedPort) && parsedPort is > 1024 and < 65536
    ? parsedPort
    : 18081;

builder.WebHost.ConfigureKestrel(options =>
{
    options.Listen(IPAddress.Loopback, port);
    options.Limits.MaxRequestBodySize = 1024 * 1024;
    options.Limits.KeepAliveTimeout = TimeSpan.FromMinutes(2);
});

builder.Logging.ClearProviders();
builder.Services.AddHttpClient("DuckAI", client =>
{
    client.Timeout = Timeout.InfiniteTimeSpan;
    client.DefaultRequestHeaders.UserAgent.ParseAdd("ResearchTools-McpHost/1.1");
});
builder.Services.AddSingleton<HostLog>();
builder.Services.AddSingleton<DuckAiClient>();
builder.Services.AddSingleton<McpSourceRegistry>();
builder.Services.AddSingleton<ResearchCoordinator>();

var app = builder.Build();
var log = app.Services.GetRequiredService<HostLog>();

app.Use(async (context, next) =>
{
    context.Response.Headers["Cache-Control"] = "no-store";
    context.Response.Headers["X-Content-Type-Options"] = "nosniff";

    if (!LoopbackSecurity.IsAllowedHost(context.Request.Host.Value, port))
    {
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        await context.Response.WriteAsJsonAsync(new { error = "Invalid Host header." });
        return;
    }

    var origin = context.Request.Headers["Origin"].ToString();
    if (!string.IsNullOrEmpty(origin) && !LoopbackSecurity.IsAllowedOrigin(origin))
    {
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        await context.Response.WriteAsJsonAsync(new { error = "Origin is not allowed." });
        return;
    }

    if (LoopbackSecurity.IsAllowedOrigin(origin))
    {
        context.Response.Headers["Access-Control-Allow-Origin"] = origin;
        context.Response.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
        context.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type";
        context.Response.Headers["Access-Control-Allow-Private-Network"] = "true";
        context.Response.Headers["Vary"] = "Origin";
    }

    if (HttpMethods.IsOptions(context.Request.Method))
    {
        context.Response.StatusCode = LoopbackSecurity.IsAllowedOrigin(origin)
            ? StatusCodes.Status204NoContent
            : StatusCodes.Status403Forbidden;
        return;
    }

    await next();
});

app.MapGet("/ping", () => Results.Json(new
{
    service = "ResearchTools.McpHost",
    listening = true,
    version = "1.1.0"
}, JsonDefaults.Web));

app.MapGet("/health", async (DuckAiClient duckAi, CancellationToken cancellationToken) =>
{
    try
    {
        var snapshot = await duckAi.GetModelsAsync(cancellationToken);
        return Results.Json(new
        {
            ready = true,
            version = "1.1.0",
            provider = "DuckAI",
            model = snapshot.DefaultModel,
            models = snapshot.Models
        }, JsonDefaults.Web);
    }
    catch (Exception exception) when (exception is not OperationCanceledException)
    {
        return Results.Json(new
        {
            ready = false,
            version = "1.1.0",
            error = exception.Message
        }, JsonDefaults.Web, statusCode: StatusCodes.Status503ServiceUnavailable);
    }
});

app.MapGet("/v1/models", async (DuckAiClient duckAi, CancellationToken cancellationToken) =>
{
    try
    {
        var snapshot = await duckAi.GetModelsAsync(cancellationToken);
        return Results.Json(new
        {
            data = snapshot.Models.Select(model => new { id = model, @object = "model" }),
            models = snapshot.Models
        }, JsonDefaults.Web);
    }
    catch (Exception exception) when (exception is not OperationCanceledException)
    {
        return Results.Json(new { error = exception.Message }, JsonDefaults.Web,
            statusCode: StatusCodes.Status503ServiceUnavailable);
    }
});

// Explicit diagnostic endpoint: it discovers both source catalogs and reports their read-only tools.
// It does not invoke DuckAI or execute a research query.
app.MapGet("/v1/mcp/sources", async (
    McpSourceRegistry sources,
    CancellationToken cancellationToken) =>
{
    var discovery = await sources.DiscoverAllAsync(cancellationToken);
    var ids = new[] { "aljam3", "shamela" };
    var labels = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["aljam3"] = "الجامع",
        ["shamela"] = "المكتبة الشاملة"
    };
    var status = ids.Select(id =>
    {
        var tools = discovery.Tools.Where(tool => tool.Server == id).ToArray();
        var warning = discovery.Warnings.FirstOrDefault(item =>
            item.StartsWith(id + ":", StringComparison.Ordinal));
        return new
        {
            id,
            label = labels[id],
            connected = warning is null,
            toolCount = tools.Length,
            tools = tools.Select(tool => tool.Name).ToArray(),
            warning
        };
    }).ToArray();

    return Results.Json(new
    {
        connected = status.Count(source => source.connected),
        totalTools = status.Sum(source => source.toolCount),
        sources = status,
        connectionWarnings = discovery.Warnings
    }, JsonDefaults.Web);
});

app.MapPost("/v1/research", async (
    HttpContext context,
    ResearchCoordinator coordinator,
    HostLog hostLog) =>
{
    var origin = context.Request.Headers["Origin"].ToString();
    if (!LoopbackSecurity.IsAllowedOrigin(origin))
        return Results.Json(new { error = "Research origin is not allowed." },
            statusCode: StatusCodes.Status403Forbidden);

    byte[] payloadBytes;
    await using (var buffer = new MemoryStream())
    {
        var chunk = new byte[16 * 1024];
        while (true)
        {
            var read = await context.Request.Body.ReadAsync(chunk.AsMemory(), context.RequestAborted);
            if (read == 0) break;
            if (buffer.Length + read > 1024 * 1024)
                return Results.Json(new { error = "Research request is too large." },
                    statusCode: StatusCodes.Status413PayloadTooLarge);
            await buffer.WriteAsync(chunk.AsMemory(0, read), context.RequestAborted);
        }
        payloadBytes = buffer.ToArray();
    }

    ResearchRequest? request;
    try
    {
        request = JsonSerializer.Deserialize<ResearchRequest>(payloadBytes, JsonDefaults.Web);
    }
    catch (JsonException)
    {
        return Results.Json(new { error = "Invalid research payload." },
            statusCode: StatusCodes.Status400BadRequest);
    }

    if (request is null || !LoopbackSecurity.IsAuthorizedResearchClient(origin, request.Client))
        return Results.Json(new { error = "Research origin or client is not allowed." },
            statusCode: StatusCodes.Status403Forbidden);

    try
    {
        var result = await coordinator.RunAsync(request, context.RequestAborted);
        return Results.Json(result, JsonDefaults.Web);
    }
    catch (ResearchBusyException exception)
    {
        context.Response.Headers["Retry-After"] = "3";
        return Results.Json(new { error = exception.Message }, JsonDefaults.Web,
            statusCode: StatusCodes.Status429TooManyRequests);
    }
    catch (TimeoutException exception)
    {
        hostLog.Warning($"Research request timed out: {exception.Message}");
        return Results.Json(new { error = exception.Message }, JsonDefaults.Web,
            statusCode: StatusCodes.Status504GatewayTimeout);
    }
    catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested)
    {
        return Results.Empty;
    }
    catch (Exception exception)
    {
        hostLog.Error("Research request failed", exception);
        return Results.Json(new { error = exception.Message }, JsonDefaults.Web,
            statusCode: StatusCodes.Status502BadGateway);
    }
});

app.Lifetime.ApplicationStarted.Register(() =>
    log.Info($"MCP host started on 127.0.0.1:{port}; runtime=.NET 10; MCP transport=official SDK AutoDetect."));
app.Lifetime.ApplicationStopping.Register(() =>
    log.Info("MCP host is stopping."));

await app.RunAsync();
