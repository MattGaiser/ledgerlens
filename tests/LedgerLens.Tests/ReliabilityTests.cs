using System.Net;
using System.Text;
using System.Text.Json;
using LedgerLens.Core;
using LedgerLens.Service;
using Microsoft.Extensions.Hosting;
using Newtonsoft.Json.Linq;
using Xunit;

namespace LedgerLens.Tests;

public sealed class ReliabilityTests
{
    private static FinancialStore Store() => new(Path.Combine(CoreTests.Root, "data", "financials.json"));
    private static FinancialFact[] Facts => CoreTests.Data.Facts.Where(f => f.Ticker == "MSFT").ToArray();

    [Fact]
    public void CircuitPermitsOnlyOneRecoveryProbeAndIgnoresPreviousConnectionEpochs()
    {
        var now = DateTimeOffset.UtcNow; var state = new ResilienceState(() => now);
        for (var i = 0; i < 3; i++) { Assert.True(state.TryEnter(out var lease, out _)); state.Failed(lease); }
        Assert.Equal("open", state.Circuit); Assert.False(state.TryEnter(out _, out _));
        now += TimeSpan.FromSeconds(9); Assert.Equal("half-open", state.Circuit);
        Assert.True(state.TryEnter(out var probe, out _)); Assert.False(state.TryEnter(out _, out _));
        state.Succeeded(probe); Assert.Equal("closed", state.Circuit);
        state.SetMode(ConnectionMode.Online); state.Failed(probe); state.Failed(probe); state.Failed(probe);
        Assert.Equal("closed", state.Circuit);
    }

    [Fact]
    public async Task InvalidationSeparatesNewCallersAndCannotBeUndoneByAnOldLoad()
    {
        var cache = new AsyncCache<string, int>(TimeSpan.FromMinutes(1));
        var release = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        var old = cache.GetAsync("a", () => release.Task);
        cache.Clear();
        Assert.Equal(2, await cache.GetAsync("a", () => Task.FromResult(2)));
        release.SetResult(1); Assert.Equal(1, await old);
        Assert.Equal(2, await cache.GetAsync("a", () => Task.FromResult(999)));
    }

    [Fact]
    public void PreviewDependenciesDetectCompanyAndMetricMapEdits()
    {
        var store = Store();
        var plan = RefreshPlanner.Create("book", [new ModelCell { Address = "Model!E10", Content = "e:", Ticker = "MSFT", Metric = "Revenue", Period = "FY2025" }], store.Get,
            new Dictionary<string, string> { ["Model!B4"] = "s:MSFT", ["Model!B10"] = "s:Revenue" });
        Assert.Empty(RefreshPlanner.Conflicts(plan, "book", a => a == "Model!E10" ? "e:" : plan.Dependencies[a]));
        Assert.Equal(["Model!B4"], RefreshPlanner.Conflicts(plan, "book", a => a == "Model!E10" ? "e:" : a == "Model!B4" ? "s:AAPL" : "s:Revenue"));
    }

    [Fact]
    public void InvalidReplacementNeverPartiallyUpdatesTheStore()
    {
        var store = Store(); var original = store.ForCompany("MSFT"); var replacements = original.Select(f => f.Copy()).ToArray();
        replacements[0].RawValue *= 2; replacements[0].Value *= 2; replacements[^1].Unit = "wrong";
        Assert.Throws<InvalidOperationException>(() => store.Replace(replacements));
        Assert.Equal(original.Select(f => f.Value), store.ForCompany("MSFT").Select(f => f.Value));
    }

    [Fact]
    public async Task ReadersSeeOneCompleteSnapshotDuringAConcurrentReplacement()
    {
        var store = Store(); var before = store.ForCompany("MSFT");
        var after = before.Select(f => { var copy = f.Copy(); copy.RawValue *= 2; copy.Value *= 2; return copy; }).ToArray();
        var first = before.Select(f => f.Value).ToArray(); var second = after.Select(f => f.Value).ToArray();
        var writer = Task.Run(() => { for (var i = 0; i < 100; i++) store.Replace(i % 2 == 0 ? after : before); });
        for (var i = 0; i < 500; i++)
        {
            var observed = store.ForCompany("MSFT").Select(f => f.Value).ToArray();
            Assert.True(observed.SequenceEqual(first) || observed.SequenceEqual(second));
        }
        await writer;
    }

    [Fact]
    public void SyncedEvidenceSurvivesRestartAndCatalogReflectsIt()
    {
        var folder = Path.Combine(Path.GetTempPath(), "LedgerLens-test-" + Guid.NewGuid()); Directory.CreateDirectory(folder);
        try
        {
            var file = Path.Combine(folder, "facts.json"); var source = Path.Combine(CoreTests.Root, "data", "financials.json");
            var store = new FinancialStore(source, file); var updated = store.ForCompany("MSFT");
            foreach (var fact in updated) fact.AcquiredAt = fact.AcquiredAt.AddMinutes(1);
            store.Replace(updated); var restarted = new FinancialStore(source, file);
            Assert.Equal(updated[0].AcquiredAt, restarted.Get(updated[0].Key).AcquiredAt);
            Assert.Equal(updated[0].AcquiredAt, restarted.Dataset.Facts.Single(f => f.Key.Equals(updated[0].Key)).AcquiredAt);
        }
        finally
        {
            var resolved = Path.GetFullPath(folder);
            if (!resolved.StartsWith(Path.GetFullPath(Path.GetTempPath()), StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(resolved).StartsWith("LedgerLens-test-", StringComparison.Ordinal)) throw new InvalidOperationException("Unexpected test cleanup path.");
            Directory.Delete(resolved, true);
        }
    }

    [Fact]
    public void AnnualSelectionIgnoresQuarterlyAndFutureFiledFacts()
    {
        var prior = Facts.Single(f => f.Metric == "Revenue" && f.Period == "FY2025");
        var concept = prior.Concept.Split(':')[1]; var good = new { form = "10-K", start = "2024-07-01", end = "2025-06-30", filed = "2025-07-30", accn = "0000789019-25-000001", val = 281724000000m };
        var json = JObject.Parse(JsonSerializer.Serialize(new { facts = new Dictionary<string, object> { ["us-gaap"] = new Dictionary<string, object> { [concept] = new { units = new Dictionary<string, object> { ["USD"] = new object[] { good, new { form = "10-K", start = "2025-04-01", end = "2025-06-30", filed = "2025-08-01", accn = good.accn, val = 1 }, new { form = "10-K", start = good.start, end = good.end, filed = "2099-01-01", accn = good.accn, val = 2 } } } } } } }));
        var actual = SecClient.SelectAnnualFact(json, prior, "0000789019", DateTimeOffset.Parse("2026-10-02T00:00:00Z"));
        Assert.Equal(281724m, actual.Value); Assert.Equal("2025-07-30", actual.Filed);
    }

    [Theory]
    [InlineData("not json")]
    [InlineData("[]")]
    [InlineData("{\"status\":\"completed\"}")]
    [InlineData("{\"status\":\"completed\",\"output\":null}")]
    [InlineData("{\"status\":\"completed\",\"output\":[{\"content\":[{\"type\":\"output_text\",\"text\":\"{bad}\"}]}]}")]
    public void MalformedProviderResponsesProduceSafeValidationFailures(string raw)
    { Assert.Equal("validation", Assert.Throws<ResearchUnavailableException>(() => ResearchService.ParseResponse(raw, Facts)).Code); }

    [Theory]
    [InlineData(401, "invalid_key")]
    [InlineData(429, "rate_or_quota")]
    [InlineData(500, "provider_error")]
    public async Task ProviderErrorsNeverExposeTheKeyOrRawResponse(int status, string code)
    {
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent("sensitive-provider-payload fake-key-secret") })));
        var service = Research(client);
        var error = await Assert.ThrowsAsync<ResearchUnavailableException>(() => service.AskAsync(new ResearchRequest(), default));
        Assert.Equal(code, error.Code); Assert.DoesNotContain("secret", error.Message); Assert.DoesNotContain("sensitive", error.Message);
    }

    [Fact]
    public async Task MissingKeyUsesLabeledCalculatedAnalysisWithoutAnHttpCall()
    {
        using var client = new HttpClient(new Handler((_, _) => throw new Exception("Must never send")));
        var service = new ResearchService(client, Store(), new ResilienceState(), new Lifetime(), new ResearchOptions { ReadKey = () => null });
        var answer = await service.AskAsync(new ResearchRequest(), default);
        Assert.False(answer.IsAiGenerated); Assert.Contains(answer.Caveats, c => c.Contains("No OpenAI key"));
    }

    [Fact]
    public async Task AiTimeoutReleasesConcurrencySlotsAndReturnsAnActionableError()
    {
        using var client = new HttpClient(new Handler(async (_, token) => { await Task.Delay(Timeout.Infinite, token); throw new InvalidOperationException(); }));
        var service = Research(client, TimeSpan.FromMilliseconds(60));
        for (var i = 0; i < 3; i++)
        {
            var error = await Assert.ThrowsAsync<ResearchUnavailableException>(() => service.AskAsync(new ResearchRequest(), default));
            Assert.Equal("timeout", error.Code);
        }
    }

    [Fact]
    public async Task IdenticalAiRequestsCoalesceAndReturnedObjectsAreIsolated()
    {
        var response = JsonSerializer.Serialize(new { status = "completed", model = "fixture-model", output = new[] { new { content = new[] { new { type = "output_text", text = JsonSerializer.Serialize(ResearchService.Describe(Facts, "Fixture")) } } } } });
        var calls = 0;
        using var client = new HttpClient(new Handler(async (_, token) => { Interlocked.Increment(ref calls); await Task.Delay(50, token); return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(response, Encoding.UTF8, "application/json") }; }));
        var service = Research(client);
        var answers = await Task.WhenAll(Enumerable.Range(0, 20).Select(_ => service.AskAsync(new ResearchRequest(), default)));
        Assert.Equal(1, calls); Assert.All(answers, a => Assert.True(a.IsAiGenerated));
        answers[0].Claims[0].Text = "Changed only here";
        Assert.DoesNotContain(answers.Skip(1), a => a.Claims[0].Text == "Changed only here");
        Assert.True((await service.AskAsync(new ResearchRequest(), default)).Cached);
    }

    private static ResearchService Research(HttpClient http, TimeSpan? timeout = null) => new(http, Store(), new ResilienceState(), new Lifetime(), new ResearchOptions { ReadKey = () => "fake-key-secret", Timeout = timeout ?? TimeSpan.FromSeconds(3) });
    private sealed class Handler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
    { protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request, cancellationToken); }
    private sealed class Lifetime : IHostApplicationLifetime
    {
        public CancellationToken ApplicationStarted => default;
        public CancellationToken ApplicationStopping => default;
        public CancellationToken ApplicationStopped => default;
        public void StopApplication() { }
    }
}
