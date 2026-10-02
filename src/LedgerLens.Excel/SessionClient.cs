using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LedgerLens.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LedgerLens.Excel
{
    internal sealed class SessionClient : IDisposable
    {
        private readonly HttpClient http;
        private readonly AsyncCache<string, FactResult> cache = new AsyncCache<string, FactResult>(TimeSpan.FromMinutes(5), 256);
        private readonly CancellationTokenSource stop = new CancellationTokenSource();
        private readonly FinancialDataset snapshot;
        private int disposed;
        public RuntimeEndpoint Endpoint
        {
            get;
        }
        public string Root
        {
            get;
        }
        public SessionClient(string root)
        {
            Root = root;
            Endpoint = JsonConvert.DeserializeObject<RuntimeEndpoint>(File.ReadAllText(Path.Combine(root, ".runtime", "endpoint.json"))) ?? throw new InvalidDataException("Start the LedgerLens research service first.");
            if (!Uri.TryCreate(Endpoint.BaseUrl, UriKind.Absolute, out var uri) || uri.Host != "127.0.0.1" || uri.Scheme != "http" || Endpoint.Token.Length != 64)
                throw new InvalidDataException("Invalid local service configuration.");
            http = new HttpClient(new HttpClientHandler { UseProxy = false }) { BaseAddress = uri, Timeout = TimeSpan.FromSeconds(90) };
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", Endpoint.Token);
            snapshot = JsonConvert.DeserializeObject<FinancialDataset>(File.ReadAllText(Path.Combine(root, "data", "financials.json"))) ?? throw new InvalidDataException("The evidence snapshot is missing.");
        }
        public FinancialDataset Snapshot => snapshot;
        public async Task<T> GetAsync<T>(string path, CancellationToken ct = default) => await SendAsync<T>(HttpMethod.Get, path, null, ct).ConfigureAwait(false);
        public async Task<T> PostAsync<T>(string path, object? payload = null, CancellationToken ct = default) => await SendAsync<T>(HttpMethod.Post, path, payload, ct).ConfigureAwait(false);
        private async Task<T> SendAsync<T>(HttpMethod method, string path, object? payload, CancellationToken ct)
        {
            using (var combined = CancellationTokenSource.CreateLinkedTokenSource(ct, stop.Token))
            using (var request = new HttpRequestMessage(method, path))
            {
                if (payload != null)
                    request.Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                using (var response = await http.SendAsync(request, combined.Token).ConfigureAwait(false))
                {
                    var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        string message;
                        try
                        {
                            message = JObject.Parse(text).Value<string>("message") ?? "The research service rejected the request.";
                        }
                        catch (JsonException) { message = "The research service rejected the request."; }
                        throw new ServiceException((int)response.StatusCode, message);
                    }
                    return JsonConvert.DeserializeObject<T>(text) ?? throw new InvalidDataException("The service returned an empty response.");
                }
            }
        }
        public async Task<FactResult> FactAsync(FactKey key, double revision, CancellationToken ct)
        {
            try
            {
                return await cache.GetAsync(key + ":" + revision.ToString("R", System.Globalization.CultureInfo.InvariantCulture), () => GetAsync<FactResult>("/api/facts/" + key.Ticker + "/" + key.Metric + "/" + key.Period, stop.Token), ct).ConfigureAwait(false);
            }
            catch (Exception e) when (e is HttpRequestException || e is TaskCanceledException && !ct.IsCancellationRequested && !stop.IsCancellationRequested)
            {
                var saved = Array.Find(snapshot.Facts, f => f.Key.Equals(key));
                if (saved == null)
                    throw;
                return new FactResult { Fact = saved.Copy(), Freshness = "cached", Provider = "Bundled SEC snapshot", ServedAt = DateTimeOffset.UtcNow, Message = "Service unavailable. Showing saved reported evidence." };
            }
        }
        public void ClearCache() => cache.Clear();
        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) == 0)
            {
                stop.Cancel();
                http.Dispose();
                stop.Dispose();
            }
        }
    }
    internal sealed class ServiceException : Exception
    {
        public int Status
        {
            get;
        }
        public ServiceException(int status, string message) : base(message) { Status = status; }
    }
}
