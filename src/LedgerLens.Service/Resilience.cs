using System.Collections.Concurrent;
using LedgerLens.Core;

namespace LedgerLens.Service;

public enum ConnectionMode { Online, Offline, Slow, Unavailable }

public sealed class ResilienceState(Func<DateTimeOffset>? clock = null)
{
    private DateTimeOffset Now => clock?.Invoke() ?? DateTimeOffset.UtcNow;
    private readonly object gate = new();
    private ConnectionMode mode;
    private int failures;
    private DateTimeOffset openUntil;
    private bool probing;
    private long epoch;
    private int slowDelayMs = 2000;
    public int SlowDelayMs { get { lock (gate) return slowDelayMs; } }
    private long providerCalls, active, maximumActive, retryCount, fallbackCount;
    public ConnectionMode Mode { get { lock (gate) return mode; } }
    public long ProviderCalls => Interlocked.Read(ref providerCalls);
    public long Active => Interlocked.Read(ref active);
    public long MaximumActive => Interlocked.Read(ref maximumActive);
    public long RetryCount => Interlocked.Read(ref retryCount);
    public long FallbackCount => Interlocked.Read(ref fallbackCount);
    public string Circuit { get { lock (gate) return openUntil > Now ? "open" : openUntil != default ? "half-open" : "closed"; } }
    public void SetMode(ConnectionMode value, int delayMs = 2000)
    {
        if (delayMs < 40 || delayMs > 6000) throw new ArgumentException("Test latency must be between 40 and 6,000 milliseconds.");
        lock (gate) { mode = value; slowDelayMs = delayMs; epoch++; failures = 0; openUntil = default; probing = false; }
    }
    public bool TryEnter(out long lease, out ConnectionMode attemptMode)
    {
        lock (gate)
        {
            lease = epoch; attemptMode = mode;
            if (openUntil > Now) return false;
            if (openUntil != default) { if (probing) return false; probing = true; }
            return true;
        }
    }
    public void Succeeded(long lease) { lock (gate) { if (lease != epoch) return; failures = 0; openUntil = default; probing = false; } }
    public void Failed(long lease) { lock (gate) { if (lease != epoch) return; probing = false; if (++failures >= 3) openUntil = Now.AddSeconds(8); } }
    public void BeginCall()
    {
        Interlocked.Increment(ref providerCalls);
        var count = Interlocked.Increment(ref active);
        long previous;
        do { previous = Interlocked.Read(ref maximumActive); if (count <= previous) break; } while (Interlocked.CompareExchange(ref maximumActive, count, previous) != previous);
    }
    public void EndCall() => Interlocked.Decrement(ref active);
    public void Retried() => Interlocked.Increment(ref retryCount);
    public void Fallback() => Interlocked.Increment(ref fallbackCount);
}

public sealed class FactService(FinancialStore store, ResilienceState state, IHostApplicationLifetime lifetime)
{
    private readonly AsyncCache<string, FinancialFact> cache = new(TimeSpan.FromMinutes(5), 256);
    private readonly SemaphoreSlim concurrent = new(4);
    private long generation;
    public long CacheHits => cache.Hits;
    public long Coalesced => cache.Coalesced;
    public int CacheCount => cache.Count;
    public int InFlight => cache.InFlight;
    public void Invalidate() { Interlocked.Increment(ref generation); cache.Clear(); }
    public async Task<FactResult> GetAsync(FactKey key, CancellationToken cancellation)
    {
        // Validate existence before retrying: a missing metric is a permanent input error.
        var snapshot = store.Get(key);
        if (state.Mode == ConnectionMode.Offline) return Fallback(snapshot, "Offline — using the saved SEC snapshot.");
        var cacheKey = Interlocked.Read(ref generation) + ":" + key;
        try
        {
            var fact = await cache.GetAsync(cacheKey, async () =>
            {
                using var timeout = CancellationTokenSource.CreateLinkedTokenSource(lifetime.ApplicationStopping);
                timeout.CancelAfter(TimeSpan.FromSeconds(8));
                await concurrent.WaitAsync(timeout.Token);
                try
                {
                  if (!state.TryEnter(out var lease, out var attemptMode)) throw new ProviderUnavailableException("Circuit open; waiting before another provider attempt.");
                  try
                  {
                    for (var attempt = 0; ; attempt++)
                    {
                        state.BeginCall();
                        try
                        {
                            await Task.Delay(attemptMode == ConnectionMode.Slow ? state.SlowDelayMs : 40, timeout.Token);
                            if (attemptMode == ConnectionMode.Unavailable) throw new ProviderUnavailableException("Simulated provider outage.");
                            state.Succeeded(lease);
                            return store.Get(key);
                        }
                        catch (ProviderUnavailableException) when (attempt < 2)
                        { state.Retried(); await Task.Delay(80 * (attempt + 1), timeout.Token); }
                        finally { state.EndCall(); }
                    }
                  }
                  catch { state.Failed(lease); throw; }
                }
                finally { concurrent.Release(); }
            }, cancellation);
            return new FactResult { Fact = fact.Copy(), Freshness = "snapshot", Provider = "SEC snapshot", ServedAt = DateTimeOffset.UtcNow, Message = "Reported annual data. Use Sync SEC to check for revised filings." };
        }
        catch (ProviderUnavailableException e) { return Fallback(snapshot, e.Message); }
        catch (OperationCanceledException) when (!cancellation.IsCancellationRequested && !lifetime.ApplicationStopping.IsCancellationRequested) { return Fallback(snapshot, "Provider timed out; using the saved SEC snapshot."); }
    }
    private FactResult Fallback(FinancialFact fact, string message)
    { state.Fallback(); return new FactResult { Fact = fact, Freshness = "cached", Provider = "Saved SEC snapshot", ServedAt = DateTimeOffset.UtcNow, Message = message }; }
}
public sealed class ProviderUnavailableException(string message) : Exception(message);
