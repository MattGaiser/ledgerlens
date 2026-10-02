using LedgerLens.Core;
using LedgerLens.Service;
using Xunit;

namespace LedgerLens.Tests;

public sealed class QualityRegressionTests
{
    [Fact]
    public async Task CancelingAQueuedWorkbookActionPreventsAnyLaterMutation()
    {
        using var stop = new CancellationTokenSource();
        Action? dispatch = null;
        var writes = 0;
        var task = QueuedAction.Run(callback => dispatch = callback, () => ++writes, stop.Token);
        stop.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
        dispatch!();
        Assert.Equal(0, writes);
    }

    [Fact]
    public async Task CancellationAfterAWriteStartsCannotMisreportItsCommittedOutcome()
    {
        using var stop = new CancellationTokenSource();
        var task = QueuedAction.Run(callback => callback(), () => { stop.Cancel(); return "committed"; }, stop.Token);
        Assert.Equal("committed", await task);
    }

    [Fact]
    public async Task DispatcherFailuresReachTheCallerWithoutExecutingTheAction()
    {
        var writes = 0;
        var task = QueuedAction.Run<int>(_ => throw new InvalidOperationException("Host unloading"), () => ++writes, default);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => task);
        Assert.Equal("Host unloading", error.Message);
        Assert.Equal(0, writes);
    }

    [Fact]
    public void CopiedWorkbooksHaveIndependentSessionsAndClosingOnePreservesTheOther()
    {
        var sessions = new WorkbookSessions();
        var original = new object();
        var copy = new object();
        var documentId = Guid.NewGuid().ToString("N");
        var originalId = sessions.GetId(original, documentId);
        var copyId = sessions.GetId(copy, documentId);
        Assert.NotEqual(originalId, copyId);
        Assert.True(sessions.Remove(copy, out var removed));
        Assert.Equal(copyId, removed);
        Assert.Equal(originalId, sessions.GetId(original, documentId));
        Assert.Throws<InvalidOperationException>(() => sessions.GetId(original, Guid.NewGuid().ToString("N")));
        sessions.Clear();
        Assert.NotEqual(originalId, sessions.GetId(original, documentId));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("foreign_url")]
    [InlineData("wrong_units")]
    public void ResearchExportRejectsIncompleteOrInvalidEvidence(string defect)
    {
        var answer = ResearchService.Describe(CoreTests.Data.Facts.Where(f => f.Ticker == "MSFT").ToArray(), "Fixture");
        ResearchValidation.ValidateExport(answer);
        switch (defect)
        {
            case "missing":
                answer.Sources = answer.Sources.Skip(1).ToArray();
                break;
            case "duplicate":
                answer.Sources = answer.Sources.Concat(new[] { answer.Sources[0] }).ToArray();
                break;
            case "foreign_url":
                answer.Sources[0].SourceUrl = "https://example.com/forged";
                break;
            case "wrong_units":
                answer.Sources[0].Unit = "EUR";
                break;
        }
        Assert.Throws<InvalidOperationException>(() => ResearchValidation.ValidateExport(answer));
    }

    [Fact]
    public async Task EventSubscriptionBeginsWithItsHandshakeBeforeConcurrentPublications()
    {
        var hub = new EventHub();
        hub.Publish("test", "Before subscription");
        var (id, reader) = hub.Subscribe();
        hub.Publish("test", "After subscription");
        var first = await reader.ReadAsync();
        var second = await reader.ReadAsync();
        Assert.Equal("connected", first.Type);
        Assert.True(second.Sequence > first.Sequence);
        hub.Unsubscribe(id);
    }

    [Fact]
    public async Task ImmediatelyCompletingLoadsStillCoalesceUnderContention()
    {
        for (var round = 0; round < 100; round++)
        {
            var calls = 0;
            var cache = new AsyncCache<string, int>(TimeSpan.FromMinutes(1));
            var readers = Enumerable.Range(0, 64).Select(_ => Task.Run(() => cache.GetAsync("same", () =>
            {
                Interlocked.Increment(ref calls);
                return Task.FromResult(42);
            })));
            Assert.All(await Task.WhenAll(readers), value => Assert.Equal(42, value));
            Assert.Equal(1, calls);
            Assert.Equal(0, cache.InFlight);
        }
    }

    [Fact]
    public void EditingAnAlreadyCurrentHistoricalValueInvalidatesThePreview()
    {
        var store = new FinancialStore(Path.Combine(CoreTests.Root, "data", "financials.json"));
        var plan = RefreshPlanner.Create("book", new[]
        {
            new ModelCell { Address = "Model!E10", Content = "n:281724", Ticker = "MSFT", Metric = "Revenue", Period = "FY2025" },
            new ModelCell { Address = "Model!E11", Content = "e:", Ticker = "MSFT", Metric = "GrossProfit", Period = "FY2025" }
        }, store.Get);
        Assert.Single(plan.Changes);
        Assert.Equal(new[] { "Model!E10" }, RefreshPlanner.Conflicts(plan, "book", address => address == "Model!E10" ? "n:999" : "e:"));
    }

    [Fact]
    public void AResolverCannotSubstituteAnotherCompanysFact()
    {
        var wrong = CoreTests.Data.Facts.Single(f => f.Ticker == "AAPL" && f.Metric == "Revenue" && f.Period == "FY2025");
        Assert.Throws<InvalidOperationException>(() => RefreshPlanner.Create("book", new[]
        {
            new ModelCell { Address = "Model!E10", Content = "e:", Ticker = "MSFT", Metric = "Revenue", Period = "FY2025" }
        }, _ => wrong));
    }

    [Fact]
    public async Task SlowEventConsumersRetainOnlyTheNewestBoundedNotifications()
    {
        var hub = new EventHub();
        var (id, reader) = hub.Subscribe();
        for (var i = 0; i < 1000; i++)
            hub.Publish("test", i.ToString());
        hub.Unsubscribe(id);
        var received = new List<FeedEvent>();
        await foreach (var item in reader.ReadAllAsync())
            received.Add(item);
        Assert.Equal(64, received.Count);
        Assert.Equal(1000, received[^1].Sequence);
        Assert.Equal(0, hub.Subscribers);
    }

    [Fact]
    public async Task ConcurrentPublishersPreserveSequenceOrderForEverySubscriber()
    {
        var hub = new EventHub();
        var (id, reader) = hub.Subscribe();
        await Task.WhenAll(Enumerable.Range(0, 16).Select(worker => Task.Run(() =>
        {
            for (var i = 0; i < 100; i++)
                hub.Publish("test", worker + ":" + i);
        })));
        hub.Unsubscribe(id);
        var received = new List<long>();
        await foreach (var item in reader.ReadAllAsync())
            received.Add(item.Sequence);
        Assert.True(received.SequenceEqual(received.OrderBy(value => value)));
        Assert.Equal(1600, received[^1]);
        Assert.Equal(30, hub.Recent.Length);
        Assert.True(hub.Recent.Select(item => item.Sequence).SequenceEqual(hub.Recent.Select(item => item.Sequence).OrderBy(value => value)));
    }
}
