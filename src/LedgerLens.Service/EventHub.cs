using System.Collections.Concurrent;
using System.Threading.Channels;

namespace LedgerLens.Service;

public sealed record FeedEvent(long Sequence, string Type, string Message, DateTimeOffset At);

public sealed class EventHub
{
    private readonly ConcurrentDictionary<Guid, Channel<FeedEvent>> subscribers = new();
    private readonly ConcurrentQueue<FeedEvent> recent = new();
    private readonly object publicationGate = new();
    private long sequence;
    public long Sequence => Interlocked.Read(ref sequence);
    public int Subscribers => subscribers.Count;
    public FeedEvent[] Recent => recent.ToArray();
    public FeedEvent Publish(string type, string message)
    {
        lock (publicationGate)
        {
            var item = new FeedEvent(Interlocked.Increment(ref sequence), type, message, DateTimeOffset.UtcNow);
            recent.Enqueue(item);
            while (recent.Count > 30)
                recent.TryDequeue(out _);
            foreach (var channel in subscribers.Values)
                channel.Writer.TryWrite(item);
            return item;
        }
    }
    public (Guid Id, ChannelReader<FeedEvent> Reader) Subscribe()
    {
        lock (publicationGate)
        {
            var id = Guid.NewGuid();
            var channel = Channel.CreateBounded<FeedEvent>(new BoundedChannelOptions(64) { FullMode = BoundedChannelFullMode.DropOldest, SingleReader = true, SingleWriter = false });
            subscribers[id] = channel;
            channel.Writer.TryWrite(new FeedEvent(sequence, "connected", "Live notification channel connected.", DateTimeOffset.UtcNow));
            return (id, channel.Reader);
        }
    }
    public void Unsubscribe(Guid id)
    {
        if (subscribers.TryRemove(id, out var channel))
            channel.Writer.TryComplete();
    }
}
