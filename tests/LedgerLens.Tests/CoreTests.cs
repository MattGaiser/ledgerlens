using LedgerLens.Core;
using LedgerLens.Service;
using Newtonsoft.Json;
using Xunit;

namespace LedgerLens.Tests;

public sealed class CoreTests
{
    public static string Root => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../../"));
    public static FinancialDataset Data => JsonConvert.DeserializeObject<FinancialDataset>(File.ReadAllText(Path.Combine(Root, "data", "financials.json")))!;

    [Theory]
    [InlineData(" msft ", "operating_income", "2025", "MSFT/OperatingIncome/FY2025")]
    [InlineData("aapl", "diluted-eps", "fy2024", "AAPL/DilutedEPS/FY2024")]
    public void KeysNormalizeWithoutChangingMeaning(string ticker, string metric, string period, string expected)
    { Assert.Equal(expected, new FactKey(ticker, metric, period).ToString()); }

    [Theory]
    [InlineData("=HYPERLINK(\"bad\")", "Revenue", "FY2025")]
    [InlineData("MSFT", "EBITDA", "FY2025")]
    [InlineData("MSFT", "Revenue", "Q12025")]
    [InlineData("", "Revenue", "FY2025")]
    public void InvalidRequestsFailBeforeNetwork(string ticker, string metric, string period)
    { Assert.Throws<ArgumentException>(() => new FactKey(ticker, metric, period)); }

    [Fact]
    public void All81SourceFactsReconcile()
    {
        var dataset = Data;
        Assert.Equal(81, dataset.Facts.Length);
        Assert.Equal(81, dataset.Facts.Select(f => f.Key).Distinct().Count());
        foreach (var fact in dataset.Facts) fact.Validate();
        Assert.Equal(281724m, dataset.Facts.Single(f => f.Key.Equals(new FactKey("MSFT", "Revenue", "FY2025"))).Value);
        Assert.Equal(416161m, dataset.Facts.Single(f => f.Key.Equals(new FactKey("AAPL", "Revenue", "FY2025"))).Value);
        Assert.Equal(130497m, dataset.Facts.Single(f => f.Key.Equals(new FactKey("NVDA", "Revenue", "FY2025"))).Value);
    }
    [Fact]
    public void UnitsAndAnnualDurationAreValidated()
    {
        var fact = Data.Facts[0].Copy(); fact.Value += 1;
        Assert.Throws<InvalidOperationException>(fact.Validate);
        fact = Data.Facts[0].Copy(); fact.Start = fact.End;
        Assert.Throws<InvalidOperationException>(fact.Validate);
        fact = Data.Facts[0].Copy(); fact.SourceUrl = "https://www.sec.gov.evil.test/file";
        Assert.Throws<InvalidOperationException>(fact.Validate);
    }
    [Fact]
    public void RefreshPreservesFormulasAndInputsAndDetectsConflicts()
    {
        var source = Data.Facts.Single(f => f.Key.Equals(new FactKey("MSFT", "Revenue", "FY2025")));
        var plan = RefreshPlanner.Create("book-one", new[]
        {
            new ModelCell { Address = "Model!E10", Content = "e:", Ticker = "MSFT", Metric = "Revenue", Period = "FY2025" },
            new ModelCell { Address = "Model!F10", Content = "f:=E10*1.1", HasFormula = true },
            new ModelCell { Address = "Model!F5", Content = "n:0.1", IsAnalystInput = true }
        }, _ => source);
        Assert.Single(plan.Changes); Assert.Equal(2, plan.PreservedAddresses.Length);
        Assert.Empty(RefreshPlanner.Conflicts(plan, "book-one", _ => "e:"));
        Assert.Single(RefreshPlanner.Conflicts(plan, "book-one", _ => "n:123"));
        Assert.Single(RefreshPlanner.Conflicts(plan, "different-book", _ => "e:"));
        Assert.Empty(RefreshPlanner.RollbackConflicts(plan, _ => "n:281724"));
        Assert.Single(RefreshPlanner.RollbackConflicts(plan, _ => "n:281725"));
    }
    [Fact]
    public void DuplicateRefreshAddressesAreRejected()
    {
        var cell = new ModelCell { Address = "Model!F5", IsAnalystInput = true };
        Assert.Throws<ArgumentException>(() => RefreshPlanner.Create("book", new[] { cell, cell }, _ => throw new Exception()));
    }
    [Fact]
    public async Task ConcurrentConsumersShareOneLoad()
    {
        var cache = new AsyncCache<string, int>(TimeSpan.FromMinutes(1));
        var release = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously); var calls = 0;
        var tasks = Enumerable.Range(0, 200).Select(_ => cache.GetAsync("same-key", () => { Interlocked.Increment(ref calls); return release.Task; })).ToArray();
        Assert.Equal(1, calls); release.SetResult(42);
        Assert.All(await Task.WhenAll(tasks), value => Assert.Equal(42, value));
        Assert.Equal(199, cache.Coalesced);
    }
    [Fact]
    public async Task CancellationDoesNotPoisonOtherConsumers()
    {
        var cache = new AsyncCache<string, int>(TimeSpan.FromMinutes(1)); var release = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var canceled = new CancellationTokenSource();
        var first = cache.GetAsync("a", () => release.Task, canceled.Token); var second = cache.GetAsync("a", () => throw new Exception("should coalesce"));
        canceled.Cancel(); await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first);
        release.SetResult(7); Assert.Equal(7, await second);
    }
    [Fact]
    public async Task FailedLoadsCanBeRetriedAndExpiredValuesReload()
    {
        var now = DateTimeOffset.UtcNow; var cache = new AsyncCache<string, int>(TimeSpan.FromSeconds(2), clock: () => now);
        await Assert.ThrowsAsync<IOException>(() => cache.GetAsync("a", () => Task.FromException<int>(new IOException())));
        Assert.Equal(1, await cache.GetAsync("a", () => Task.FromResult(1)));
        Assert.Equal(1, await cache.GetAsync("a", () => Task.FromResult(2)));
        now += TimeSpan.FromSeconds(3);
        Assert.Equal(2, await cache.GetAsync("a", () => Task.FromResult(2)));
    }
    [Fact]
    public async Task CacheCapacityIsBoundedUnderConcurrentCompletion()
    {
        var cache = new AsyncCache<int, int>(TimeSpan.FromMinutes(1), 8);
        await Task.WhenAll(Enumerable.Range(0, 500).Select(i => Task.Run(() => cache.GetAsync(i, () => Task.FromResult(i)))));
        Assert.InRange(cache.Count, 1, 8);
    }
    [Theory]
    [InlineData("MSFT")][InlineData("AAPL")][InlineData("NVDA")]
    public void DeterministicResearchHasResolvableCitations(string ticker)
    {
        var facts = Data.Facts.Where(f => f.Ticker == ticker).ToArray();
        var answer = ResearchService.Describe(facts, "Test"); ResearchValidation.Validate(answer, facts);
        Assert.False(answer.IsAiGenerated); Assert.Equal(3, answer.Claims.Length);
        answer.Claims[0].SourceIds = new[] { "fabricated-source" };
        Assert.Throws<InvalidOperationException>(() => ResearchValidation.Validate(answer, facts));
    }
    [Fact]
    public void IncompleteAndRefusedAiResponsesAreNotAccepted()
    {
        Assert.Equal("incomplete", Assert.Throws<ResearchUnavailableException>(() => ResearchService.ParseResponse("{\"status\":\"incomplete\"}", Data.Facts)).Code);
        Assert.Equal("refusal", Assert.Throws<ResearchUnavailableException>(() => ResearchService.ParseResponse("{\"status\":\"completed\",\"output\":[{\"content\":[{\"type\":\"refusal\"}]}]}", Data.Facts)).Code);
    }
}
