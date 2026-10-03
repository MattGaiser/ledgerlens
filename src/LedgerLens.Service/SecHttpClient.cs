using System.Net;
using LedgerLens.Core;

namespace LedgerLens.Service;

// This policy is used only for idempotent SEC GETs. OpenAI POSTs are not replayed.
public sealed class SecHttpClient(HttpClient http, CircuitBreaker breaker, SecHttpOptions? options = null)
{
    private readonly SemaphoreSlim slots = new(4);
    private long requests, retries, active, peakActive;
    public long Requests => Interlocked.Read(ref requests);
    public long Retries => Interlocked.Read(ref retries);
    public long Active => Interlocked.Read(ref active);
    public long PeakActive => Interlocked.Read(ref peakActive);
    public string Circuit => breaker.State;

    public async Task<string> GetCompanyFactsAsync(string cik, CancellationToken cancellation)
    {
        if (!System.Text.RegularExpressions.Regex.IsMatch(cik, @"^\d{10}$"))
            throw new ArgumentException("A ten-digit SEC CIK is required.", nameof(cik));
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(options?.Timeout ?? TimeSpan.FromSeconds(20));
        var acquired = false;
        long? lease = null;
        try
        {
            await slots.WaitAsync(deadline.Token);
            acquired = true;
            if (!breaker.TryEnter(out var admitted))
                throw new SecUnavailableException("sec_circuit_open", "SEC requests are paused after repeated failures. Saved evidence is unchanged; retry shortly.");
            lease = admitted;
            for (var attempt = 0; ; attempt++)
            {
                var retryDelay = TimeSpan.FromMilliseconds((options?.RetryDelayMilliseconds ?? 250) * (attempt + 1));
                try
                {
                    using var request = new HttpRequestMessage(HttpMethod.Get, $"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json");
                    request.Headers.UserAgent.ParseAdd($"LedgerLens/{ProductInfo.Version} (individual financial research prototype)");
                    BeginRequest();
                    try
                    {
                        using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token);
                        if (response.Headers.RetryAfter is { } retryAfter)
                        {
                            var requestedDelay = retryAfter.Delta ?? (retryAfter.Date - DateTimeOffset.UtcNow);
                            if (requestedDelay > retryDelay)
                                retryDelay = requestedDelay.Value;
                        }
                        response.EnsureSuccessStatusCode();
                        // The same deadline covers queueing, headers, body and backoff.
                        var body = await response.Content.ReadAsStringAsync(deadline.Token);
                        breaker.Succeeded(admitted);
                        return body;
                    }
                    finally { Interlocked.Decrement(ref active); }
                }
                catch (HttpRequestException error) when (IsTransient(error.StatusCode) && attempt < 2)
                {
                    Interlocked.Increment(ref retries);
                }
                if (retryDelay > (options?.Timeout ?? TimeSpan.FromSeconds(20)))
                    await Task.Delay(Timeout.Infinite, deadline.Token);
                else
                    await Task.Delay(retryDelay, deadline.Token);
            }
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            if (lease.HasValue)
                breaker.Failed(lease.Value);
            throw new SecUnavailableException("sec_timeout", "SEC did not finish within the request deadline. Saved evidence is unchanged.");
        }
        catch (HttpRequestException error)
        {
            if (lease.HasValue)
            {
                if (IsTransient(error.StatusCode))
                    breaker.Failed(lease.Value);
                else
                    breaker.Succeeded(lease.Value); // A permanent HTTP rejection is not a connectivity outage.
            }
            var message = error.StatusCode.HasValue ? $"SEC returned HTTP {(int)error.StatusCode.Value}." : "SEC could not be reached.";
            throw new SecUnavailableException("sec_unavailable", message + " Saved evidence is unchanged; retry later.");
        }
        finally
        {
            if (lease.HasValue)
                breaker.Canceled(lease.Value);
            if (acquired)
                slots.Release();
        }
    }

    private static bool IsTransient(HttpStatusCode? status) => !status.HasValue || status == HttpStatusCode.RequestTimeout || status == HttpStatusCode.TooManyRequests || (int)status.Value >= 500;

    private void BeginRequest()
    {
        Interlocked.Increment(ref requests);
        var current = Interlocked.Increment(ref active);
        long previous;
        do
        {
            previous = Interlocked.Read(ref peakActive);
            if (current <= previous)
                return;
        } while (Interlocked.CompareExchange(ref peakActive, current, previous) != previous);
    }
}

public sealed class SecHttpOptions
{
    public TimeSpan Timeout { get; init; } = TimeSpan.FromSeconds(20);
    public int RetryDelayMilliseconds { get; init; } = 250;
}

public sealed class SecUnavailableException(string code, string message) : Exception(message)
{
    public string Code { get; } = code;
}
