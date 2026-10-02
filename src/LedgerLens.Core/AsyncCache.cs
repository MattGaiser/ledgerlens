using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace LedgerLens.Core
{
    /// <summary>Coalesces concurrent loads. Caller cancellation never cancels another caller's shared work.</summary>
    public sealed class AsyncCache<TKey, TValue> where TKey : notnull
    {
        private sealed class Entry
        {
            public TValue Value
            {
                get;
            }
            public DateTimeOffset Expires
            {
                get;
            }
            public Entry(TValue value, DateTimeOffset expires)
            {
                Value = value;
                Expires = expires;
            }
        }
        private readonly ConcurrentDictionary<TKey, Entry> entries = new ConcurrentDictionary<TKey, Entry>();
        private readonly ConcurrentDictionary<Tuple<long, TKey>, Lazy<Task<TValue>>> pending = new ConcurrentDictionary<Tuple<long, TKey>, Lazy<Task<TValue>>>();
        private readonly Func<DateTimeOffset> clock;
        private readonly TimeSpan lifetime;
        private readonly int capacity;
        private readonly object insertionGate = new object();
        private long hits, misses, coalesced;
        private long generation;
        public long Hits => Interlocked.Read(ref hits);
        public long Misses => Interlocked.Read(ref misses);
        public long Coalesced => Interlocked.Read(ref coalesced);
        public int Count => entries.Count;
        public int InFlight => pending.Count;

        public AsyncCache(TimeSpan lifetime, int capacity = 512, Func<DateTimeOffset>? clock = null)
        {
            if (lifetime <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(lifetime));
            if (capacity < 1)
                throw new ArgumentOutOfRangeException(nameof(capacity));
            this.lifetime = lifetime;
            this.capacity = capacity;
            this.clock = clock ?? (() => DateTimeOffset.UtcNow);
        }
        public bool TryGet(TKey key, out TValue value)
        {
            if (entries.TryGetValue(key, out var entry) && entry.Expires > clock())
            {
                value = entry.Value;
                Interlocked.Increment(ref hits);
                return true;
            }
            value = default!;
            return false;
        }
        public async Task<TValue> GetAsync(TKey key, Func<Task<TValue>> load, CancellationToken cancellation = default)
        {
            cancellation.ThrowIfCancellationRequested();
            if (load == null)
                throw new ArgumentNullException(nameof(load));
            Tuple<long, TKey> pendingKey;
            Lazy<Task<TValue>> active;
            bool ownsLoad;
            lock (insertionGate)
            {
                if (TryGet(key, out var cached))
                    return cached;
                var epoch = generation;
                pendingKey = Tuple.Create(epoch, key);
                var proposed = new Lazy<Task<TValue>>(async () =>
                {
                    Interlocked.Increment(ref misses);
                    var result = await load().ConfigureAwait(false);
                    lock (insertionGate)
                    {
                        if (generation == epoch)
                        {
                            if (entries.Count >= capacity && !entries.ContainsKey(key))
                                Evict();
                            entries[key] = new Entry(result, clock() + lifetime);
                        }
                    }
                    return result;
                }, LazyThreadSafetyMode.ExecutionAndPublication);
                // Reserve the load in the same critical section as the cache miss.
                // Otherwise a fast load can finish and be removed between those steps.
                active = pending.GetOrAdd(pendingKey, proposed);
                ownsLoad = ReferenceEquals(active, proposed);
            }
            if (!ownsLoad)
                Interlocked.Increment(ref coalesced);
            var task = active.Value;
            if (ownsLoad)
                _ = task.ContinueWith(completed =>
            {
                // Observe failures even if every waiter cancels before the shared operation completes.
                if (completed.IsFaulted)
                    _ = completed.Exception;
                ((System.Collections.Generic.ICollection<System.Collections.Generic.KeyValuePair<Tuple<long, TKey>, Lazy<Task<TValue>>>>)pending)
                    .Remove(new System.Collections.Generic.KeyValuePair<Tuple<long, TKey>, Lazy<Task<TValue>>>(pendingKey, active));
            }, CancellationToken.None, TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
            return await TaskCancellation.WaitAsync(task, cancellation).ConfigureAwait(false);
        }
        public void Clear()
        {
            lock (insertionGate)
            {
                generation++;
                entries.Clear();
            }
        }
        private void Evict()
        {
            TKey oldestKey = default!;
            var oldest = DateTimeOffset.MaxValue;
            var found = false;
            foreach (var entry in entries)
                if (entry.Value.Expires < oldest)
                {
                    oldestKey = entry.Key;
                    oldest = entry.Value.Expires;
                    found = true;
                }
            if (found)
                entries.TryRemove(oldestKey, out _);
        }
    }

    public static class TaskCancellation
    {
        public static async Task<T> WaitAsync<T>(Task<T> task, CancellationToken cancellation)
        {
            if (!cancellation.CanBeCanceled)
                return await task.ConfigureAwait(false);
            var canceled = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using (cancellation.Register(() => canceled.TrySetResult(true)))
            {
                if (task != await Task.WhenAny(task, canceled.Task).ConfigureAwait(false))
                    throw new OperationCanceledException(cancellation);
                return await task.ConfigureAwait(false);
            }
        }
    }
}
