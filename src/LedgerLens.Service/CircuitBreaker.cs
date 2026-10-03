namespace LedgerLens.Service;

public sealed class CircuitBreaker(Func<DateTimeOffset>? clock = null)
{
    private readonly object gate = new();
    private DateTimeOffset Now => clock?.Invoke() ?? DateTimeOffset.UtcNow;
    private DateTimeOffset openUntil;
    private int failures;
    private long generation;
    private bool probing;

    public string State
    {
        get
        {
            lock (gate)
                return openUntil > Now ? "open" : openUntil != default ? "half-open" : "closed";
        }
    }

    public bool TryEnter(out long lease)
    {
        lock (gate)
        {
            lease = generation;
            if (openUntil > Now || probing)
                return false;
            if (openUntil != default)
                probing = true;
            return true;
        }
    }

    public void Succeeded(long lease)
    {
        lock (gate)
        {
            if (lease != generation)
                return;
            failures = 0;
            openUntil = default;
            probing = false;
        }
    }

    public void Failed(long lease)
    {
        lock (gate)
        {
            if (lease != generation)
                return;
            probing = false;
            if (++failures >= 3)
            {
                openUntil = Now.AddSeconds(8);
                generation++; // Completions admitted before this opening no longer own its state.
            }
        }
    }

    public void Canceled(long lease)
    {
        lock (gate)
        {
            if (lease == generation)
                probing = false;
        }
    }
}
