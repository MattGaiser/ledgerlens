using LedgerLens.Core;

namespace LedgerLens.Service;

public enum ConnectionMode
{
    Online, Offline
}

public sealed class ConnectionState
{
    private int offline;
    public ConnectionMode Mode => Volatile.Read(ref offline) == 0 ? ConnectionMode.Online : ConnectionMode.Offline;
    public void SetMode(ConnectionMode mode) => Volatile.Write(ref offline, mode == ConnectionMode.Offline ? 1 : 0);
}

// Reading the validated in-memory snapshot requires neither retries nor network I/O.
// Excel's SessionClient coalesces/cache-shares HTTP reads across formula callers.
public sealed class FactService(FinancialStore store, ConnectionState connection)
{
    private long reads;
    public long Reads => Interlocked.Read(ref reads);

    public FactResult Get(FactKey key, CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        return Result(store.Get(key), connection.Mode);
    }

    public FactResult[] GetBatch(FactKey[] keys, CancellationToken cancellation)
    {
        var read = store.CaptureReader();
        var mode = connection.Mode;
        return keys.Select(key =>
        {
            cancellation.ThrowIfCancellationRequested();
            return Result(read(key), mode);
        }).ToArray();
    }

    private FactResult Result(FinancialFact fact, ConnectionMode mode)
    {
        Interlocked.Increment(ref reads);
        var offline = mode == ConnectionMode.Offline;
        return new FactResult
        {
            Fact = fact,
            Freshness = offline ? "cached" : "snapshot",
            Provider = "Saved SEC snapshot",
            ServedAt = DateTimeOffset.UtcNow,
            Message = offline ? "Offline: using saved SEC evidence." : "Reported annual data. Use Sync SEC to check for revised filings."
        };
    }
}
