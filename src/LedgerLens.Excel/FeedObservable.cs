using System;
using System.Collections.Generic;
using System.IO;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LedgerLens.Core;
using Newtonsoft.Json.Linq;

namespace LedgerLens.Excel
{
    internal sealed class FeedObservable : IObservable<object>, IDisposable
    {
        private readonly object gate = new object();
        private readonly Dictionary<long, IObserver<object>> observers = new Dictionary<long, IObserver<object>>();
        private readonly RuntimeEndpoint endpoint;
        private CancellationTokenSource? stop;
        private long nextId;
        private bool disposed;
        private string latest = "Connecting to research notifications…";
        public FeedObservable(RuntimeEndpoint endpoint) { this.endpoint = endpoint; }
        public IDisposable Subscribe(IObserver<object> observer)
        {
            lock (gate)
            {
                if (disposed) { observer.OnCompleted(); return new Subscription(() => { }); }
                var id = ++nextId; observers[id] = observer;
                observer.OnNext(latest);
                if (stop == null) { stop = new CancellationTokenSource(); var token = stop.Token; _ = Task.Run(() => RunAsync(token)); }
                return new Subscription(() => { lock (gate) { observers.Remove(id); if (observers.Count == 0) { stop?.Cancel(); stop?.Dispose(); stop = null; } } });
            }
        }
        private void Publish(string value, CancellationToken cancellation)
        {
            IObserver<object>[] snapshot;
            lock (gate) { if (disposed || cancellation.IsCancellationRequested) return; latest = value; snapshot = new List<IObserver<object>>(observers.Values).ToArray(); }
            foreach (var observer in snapshot) { try { observer.OnNext(value); } catch (Exception e) { HostRuntime.RecordError("Stream observer: " + e.GetType().Name); } }
        }
        private async Task RunAsync(CancellationToken cancellation)
        {
            var retry = 0;
            while (!cancellation.IsCancellationRequested)
            {
                using (var socket = new ClientWebSocket())
                {
                    socket.Options.Proxy = null;
                    socket.Options.AddSubProtocol("ledgerlens.v1"); socket.Options.AddSubProtocol("ll-auth." + endpoint.Token);
                    try
                    {
                        var uri = new Uri(endpoint.BaseUrl.Replace("http:", "ws:") + "/api/events");
                        await socket.ConnectAsync(uri, cancellation).ConfigureAwait(false); retry = 0;
                        var buffer = new byte[8192];
                        while (socket.State == WebSocketState.Open && !cancellation.IsCancellationRequested)
                        {
                            using var message = new MemoryStream(); WebSocketReceiveResult part;
                            do
                            {
                                part = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellation).ConfigureAwait(false);
                                if (part.MessageType == WebSocketMessageType.Close) break;
                                if (part.MessageType != WebSocketMessageType.Text) throw new InvalidDataException("Expected a text notification.");
                                message.Write(buffer, 0, part.Count);
                                if (message.Length > 32768) throw new InvalidOperationException("Stream message is too large.");
                            } while (!part.EndOfMessage);
                            if (part.MessageType == WebSocketMessageType.Close) break;
                            var item = JObject.Parse(Encoding.UTF8.GetString(message.ToArray()));
                            Publish("#" + item.Value<long>("sequence") + " · " + item.Value<string>("message"), cancellation);
                        }
                    }
                    catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { return; }
                    catch (Exception e) { HostRuntime.RecordError("Notification connection: " + e.GetType().Name); Publish("Reconnecting to research notifications…", cancellation); }
                }
                try { await Task.Delay(Math.Min(1000 * (1 << Math.Min(retry++, 4)), 15000), cancellation).ConfigureAwait(false); }
                catch (OperationCanceledException) { return; }
            }
        }
        public void Dispose()
        {
            lock (gate) { disposed = true; stop?.Cancel(); stop?.Dispose(); stop = null; observers.Clear(); }
        }
        private sealed class Subscription : IDisposable
        { private Action? release; public Subscription(Action release) { this.release = release; } public void Dispose() => Interlocked.Exchange(ref release, null)?.Invoke(); }
    }
}
