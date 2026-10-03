using System.Net;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using LedgerLens.Core;
using LedgerLens.Service;

var builder = WebApplication.CreateBuilder(args);
var root = Environment.GetEnvironmentVariable("LEDGERLENS_ROOT") ?? Directory.GetCurrentDirectory();
var port = int.TryParse(Environment.GetEnvironmentVariable("LEDGERLENS_PORT"), out var configuredPort) ? configuredPort : 17843;
var baseUrl = $"http://127.0.0.1:{port}";
var officeUrl = Environment.GetEnvironmentVariable("LEDGERLENS_HTTPS") == "1" ? $"https://127.0.0.1:{port + 1}" : null;
builder.WebHost.UseUrls(officeUrl == null ? [baseUrl] : [baseUrl, officeUrl]);
builder.WebHost.ConfigureKestrel(options => options.Limits.MaxRequestBodySize = 128 * 1024);
builder.Logging.ClearProviders();
builder.Logging.AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; });
builder.Logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
builder.Services.AddSingleton(new FinancialStore(Path.Combine(root, "data", "financials.json"), Path.Combine(root, ".runtime", "latest-financials.json")));
builder.Services.AddSingleton<ConnectionState>();
builder.Services.AddSingleton<CircuitBreaker>();
builder.Services.AddSingleton<SecHttpClient>();
builder.Services.AddSingleton<FactService>();
builder.Services.AddSingleton<EventHub>();
builder.Services.AddSingleton(new HttpClient(new HttpClientHandler { AutomaticDecompression = DecompressionMethods.All }) { Timeout = TimeSpan.FromSeconds(90) });
builder.Services.AddSingleton<ResearchService>();
builder.Services.AddSingleton<SecClient>();
builder.Services.ConfigureHttpJsonOptions(o => o.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase);
var app = builder.Build();
app.UseWebSockets(new WebSocketOptions { KeepAliveInterval = TimeSpan.FromSeconds(15) });
app.Use(async (context, next) =>
{
    if (context.Request.Host.Host != "127.0.0.1")
    {
        context.Response.StatusCode = 400;
        return;
    }
    context.Response.Headers["X-Content-Type-Options"] = "nosniff";
    context.Response.Headers["Referrer-Policy"] = "no-referrer";
    context.Response.Headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self'; img-src 'self' data:; connect-src 'self' ws://127.0.0.1:* wss://127.0.0.1:*; frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com; base-uri 'none'; object-src 'none'";
    if (context.Request.Path.StartsWithSegments("/api"))
    {
        context.Response.Headers.CacheControl = "no-store";
        var origin = context.Request.Headers.Origin.ToString();
        if (origin.Length != 0 && origin != baseUrl && origin != officeUrl)
        {
            context.Response.StatusCode = 403;
            return;
        }
        var supplied = context.Request.Headers.Authorization.ToString();
        if (supplied.StartsWith("Bearer ", StringComparison.Ordinal))
            supplied = supplied.Substring(7);
        else
            supplied = "";
        // Browsers cannot set WebSocket Authorization; token is sent as a subprotocol, never in URLs/logs.
        if (context.WebSockets.IsWebSocketRequest)
            supplied = context.WebSockets.WebSocketRequestedProtocols.FirstOrDefault(s => s.StartsWith("ll-auth.", StringComparison.Ordinal))?.Substring(8) ?? "";
        if (supplied.Length != token.Length || !CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(supplied), Encoding.UTF8.GetBytes(token)))
        {
            await Results.Json(new
            {
                code = "unauthorized",
                message = "Start LedgerLens using its launcher to connect this workspace."
            }, statusCode: 401).ExecuteAsync(context);
            return;
        }
    }
    try
    {
        await next(context);
    }
    catch (OperationCanceledException) when (context.RequestAborted.IsCancellationRequested) { }
    catch (BadHttpRequestException e) { context.Response.StatusCode = e.StatusCode; }
    catch (SecUnavailableException error)
    {
        await Results.Json(new
        {
            code = error.Code,
            message = error.Message
        }, statusCode: 502).ExecuteAsync(context);
    }
    catch (Exception e) when (e is ArgumentException or KeyNotFoundException or ResearchUnavailableException or InvalidDataException)
    {
        var code = e is ResearchUnavailableException ai ? ai.Code : e is KeyNotFoundException ? "not_found" : "invalid_request";
        var status = e is ResearchUnavailableException ? 502 : e is KeyNotFoundException ? 404 : 400;
        await Results.Json(new
        {
            code,
            message = e.Message
        }, statusCode: status).ExecuteAsync(context);
    }
    catch (Exception e)
    {
        app.Logger.LogError("Request failed: {Type}", e.GetType().Name);
        if (!context.Response.HasStarted)
            await Results.Json(new
            {
                code = "internal_error",
                message = "The operation failed. Your workbook was not changed. See diagnostics and retry."
            }, statusCode: 500).ExecuteAsync(context);
    }
});

app.MapGet("/health", () => new { status = "ready", version = ProductInfo.Version });
app.MapGet("/api/catalog", (FinancialStore store) =>
{
    var snapshot = store.Dataset;
    return new
    {
        snapshot.Companies,
        metrics = MetricCatalog.Labels.Select(p => new { id = p.Key, label = p.Value }),
        periods = new[] { "FY2023", "FY2024", "FY2025" },
        snapshot.SnapshotDate,
        snapshot.Description,
        facts = snapshot.Facts
    };
});
app.MapGet("/api/facts/{ticker}/{metric}/{period}", (string ticker, string metric, string period, FactService facts, CancellationToken ct) => facts.Get(new FactKey(ticker, metric, period), ct));
app.MapPost("/api/facts/batch", (FactKey[] keys, FactService facts, CancellationToken ct) =>
{
    if (keys == null || keys.Length == 0 || keys.Length > 128 || keys.Any(key => key == null))
        throw new ArgumentException("Batch requests require 1 to 128 non-null facts.");
    return facts.GetBatch(keys, ct);
});
app.MapPost("/api/research", async (ResearchRequest request, ResearchService research, EventHub events, CancellationToken ct) =>
{
    var answer = await research.AskAsync(request, ct);
    events.Publish("research", $"{request.Ticker.ToUpperInvariant()} research ready · {answer.Provider}");
    return answer;
});
app.MapGet("/api/diagnostics", (FinancialStore store, FactService facts, ConnectionState state, SecHttpClient sec, ResearchService research, EventHub events) => new
{
    version = ProductInfo.Version,
    mode = state.Mode.ToString().ToLowerInvariant(),
    secCircuit = sec.Circuit,
    secRequests = sec.Requests,
    secActiveRequests = sec.Active,
    secPeakConcurrency = sec.PeakActive,
    secRetries = sec.Retries,
    snapshotReads = facts.Reads,
    sourceFacts = store.Count,
    snapshotDate = store.Dataset.SnapshotDate,
    aiConfigured = research.Configured,
    aiModel = research.Model,
    aiCalls = research.Calls,
    streamSubscribers = events.Subscribers,
    sequence = events.Sequence,
    recentEvents = events.Recent,
    workingSetMb = Environment.WorkingSet / 1048576
});
app.MapPost("/api/connection", (ConnectionRequest request, ConnectionState state, EventHub events) =>
{
    if (!Enum.TryParse<ConnectionMode>(request.Mode, true, out var mode) || !Enum.IsDefined(mode))
        throw new ArgumentException("Choose online or offline.");
    state.SetMode(mode);
    return events.Publish("connection", "Connection mode: " + mode.ToString().ToLowerInvariant());
});
app.MapPost("/api/cache/clear", (ResearchService research, EventHub events) => { research.Invalidate(); return events.Publish("cache", "Research cache cleared. Saved SEC evidence retained."); });
app.MapPost("/api/shutdown", (IHostApplicationLifetime lifetime) => { lifetime.StopApplication(); return Results.Json(new { status = "stopping" }); });
app.MapPost("/api/sync/{ticker}", async (string ticker, FinancialStore store, SecClient sec, ResearchService research, ConnectionState state, EventHub events, CancellationToken ct) =>
{
    if (state.Mode == ConnectionMode.Offline)
        throw new ArgumentException("Reconnect before syncing SEC filings.");
    var company = store.Dataset.Companies.SingleOrDefault(c => c.Ticker.Equals(ticker, StringComparison.OrdinalIgnoreCase)) ?? throw new KeyNotFoundException("Company is not in this research universe.");
    var replacements = await sec.RefreshAsync(company, store.ForCompany(company.Ticker), ct);
    store.Replace(replacements);
    research.Invalidate();
    events.Publish("filings", company.Ticker + " SEC facts revalidated against the live companyfacts API.");
    return new
    {
        count = replacements.Length,
        ticker = company.Ticker,
        syncedAt = DateTimeOffset.UtcNow
    };
});
app.MapPost("/api/refresh/preview", (PreviewRequest request, FinancialStore store) => RefreshPlanner.Create(request.WorkbookId, request.Cells, store.CaptureReader(), request.Dependencies));
app.MapPost("/api/events/test", (EventHub events) => events.Publish("diagnostic", "Notification channel test received."));
app.Map("/api/events", async (HttpContext context, EventHub events, IHostApplicationLifetime lifetime) =>
{
    if (!context.WebSockets.IsWebSocketRequest || !context.WebSockets.WebSocketRequestedProtocols.Contains("ledgerlens.v1"))
    {
        context.Response.StatusCode = 400;
        return;
    }
    using var socket = await context.WebSockets.AcceptWebSocketAsync("ledgerlens.v1");
    var (id, reader) = events.Subscribe();
    using var stop = CancellationTokenSource.CreateLinkedTokenSource(context.RequestAborted, lifetime.ApplicationStopping);
    try
    {
        async Task Send(object value) => await socket.SendAsync(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase })), WebSocketMessageType.Text, true, stop.Token);
        var receive = Task.Run(async () =>
        {
            try
            {
                var buffer = new byte[1024];
                while (socket.State == WebSocketState.Open)
                {
                    var result = await socket.ReceiveAsync(buffer, stop.Token);
                    if (result.MessageType == WebSocketMessageType.Close)
                        break;
                }
            }
            finally { stop.Cancel(); }
        }, stop.Token);
        try
        {
            await foreach (var item in reader.ReadAllAsync(stop.Token))
                await Send(item);
        }
        finally { stop.Cancel(); try { await receive; } catch (OperationCanceledException) { } }
    }
    catch (Exception e) when (e is OperationCanceledException or WebSocketException) { }
    finally
    {
        events.Unsubscribe(id);
        if (socket.State == WebSocketState.CloseReceived)
        {
            using var closing = new CancellationTokenSource(TimeSpan.FromSeconds(1));
            try
            {
                await socket.CloseOutputAsync(WebSocketCloseStatus.NormalClosure, "Disconnected", closing.Token);
            }
            catch (Exception e) when (e is OperationCanceledException or WebSocketException) { }
        }
    }
});

var webRoot = Path.Combine(root, "web");
app.UseDefaultFiles(new DefaultFilesOptions { FileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(webRoot) });
app.UseStaticFiles(new StaticFileOptions { FileProvider = new Microsoft.Extensions.FileProviders.PhysicalFileProvider(webRoot) });
var runtimeDir = Path.Combine(root, ".runtime");
Directory.CreateDirectory(runtimeDir);
var endpointPath = Path.Combine(runtimeDir, "endpoint.json");
await app.StartAsync();
await File.WriteAllTextAsync(endpointPath + ".tmp", JsonSerializer.Serialize(new RuntimeEndpoint { BaseUrl = baseUrl, OfficeUrl = officeUrl, Token = token, ProcessId = Environment.ProcessId }));
File.Move(endpointPath + ".tmp", endpointPath, true);
app.Services.GetRequiredService<EventHub>().Publish("service", "LedgerLens research service ready.");
await app.WaitForShutdownAsync();

public sealed class ConnectionRequest
{
    public string Mode { get; set; } = "online";
}
public sealed class PreviewRequest
{
    public string WorkbookId { get; set; } = ""; public ModelCell[] Cells { get; set; } = []; public Dictionary<string, string> Dependencies { get; set; } = new();
}
public partial class Program
{
}
