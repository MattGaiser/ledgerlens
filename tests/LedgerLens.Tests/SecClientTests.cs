using System.Net;
using System.Net.Http.Headers;
using LedgerLens.Core;
using LedgerLens.Service;
using Newtonsoft.Json.Linq;
using Xunit;

namespace LedgerLens.Tests;

public sealed class SecClientTests
{
    private static FinancialFact Revenue => CoreTests.Data.Facts.Single(f => f.Ticker == "MSFT" && f.Metric == "Revenue" && f.Period == "FY2025");
    private static SecHttpClient Transport(HttpClient client, CircuitBreaker? breaker = null, int timeoutMs = 2000) =>
        new(client, breaker ?? new CircuitBreaker(), new SecHttpOptions { Timeout = TimeSpan.FromMilliseconds(timeoutMs), RetryDelayMilliseconds = 0 });

    [Theory]
    [InlineData(408)]
    [InlineData(429)]
    [InlineData(503)]
    public async Task TransientSecGetFailuresRetryAndReturnTheResponseBody(int status)
    {
        var calls = 0;
        using var client = new HttpClient(new Handler((request, _) =>
        {
            Assert.Equal(HttpMethod.Get, request.Method);
            Assert.Equal("https://data.sec.gov/api/xbrl/companyfacts/CIK0000789019.json", request.RequestUri!.AbsoluteUri);
            return Task.FromResult(Response(++calls < 3 ? status : 200));
        }));
        var transport = Transport(client);
        Assert.Equal("{}", await transport.GetCompanyFactsAsync("0000789019", default));
        Assert.Equal(3, transport.Requests);
        Assert.Equal(2, transport.Retries);
        Assert.Equal(0, transport.Active);
    }

    [Fact]
    public async Task ATransientNetworkFailureIsRetried()
    {
        var calls = 0;
        using var client = new HttpClient(new Handler((_, _) => ++calls == 1
            ? Task.FromException<HttpResponseMessage>(new HttpRequestException("Sensitive transport details"))
            : Task.FromResult(Response(200))));
        var transport = Transport(client);
        Assert.Equal("{}", await transport.GetCompanyFactsAsync("0000789019", default));
        Assert.Equal(2, transport.Requests);
    }

    [Fact]
    public async Task PermanentHttpRejectionIsNotRetriedAndDoesNotTripTheCircuit()
    {
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(Response(403))));
        var transport = Transport(client);
        for (var i = 0; i < 4; i++)
        {
            var error = await Assert.ThrowsAsync<SecUnavailableException>(() => transport.GetCompanyFactsAsync("0000789019", default));
            Assert.Contains("HTTP 403", error.Message);
            Assert.DoesNotContain("sensitive", error.Message);
        }
        Assert.Equal(4, transport.Requests);
        Assert.Equal(0, transport.Retries);
        Assert.Equal("closed", transport.Circuit);
    }

    [Fact]
    public async Task RepeatedTransportFailuresOpenTheCircuitThenPermitRecovery()
    {
        var now = DateTimeOffset.UtcNow;
        var healthy = false;
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(Response(healthy ? 200 : 503))));
        var transport = Transport(client, new CircuitBreaker(() => now));
        for (var i = 0; i < 3; i++)
            await Assert.ThrowsAsync<SecUnavailableException>(() => transport.GetCompanyFactsAsync("0000789019", default));
        Assert.Equal(9, transport.Requests);
        Assert.Equal("open", transport.Circuit);
        Assert.Equal("sec_circuit_open", (await Assert.ThrowsAsync<SecUnavailableException>(() => transport.GetCompanyFactsAsync("0000789019", default))).Code);
        Assert.Equal(9, transport.Requests);
        now += TimeSpan.FromSeconds(9);
        healthy = true;
        Assert.Equal("{}", await transport.GetCompanyFactsAsync("0000789019", default));
        Assert.Equal("closed", transport.Circuit);
    }

    [Fact]
    public void CancelingARecoveryProbeAllowsExactlyOneReplacement()
    {
        var now = DateTimeOffset.UtcNow;
        var breaker = new CircuitBreaker(() => now);
        for (var i = 0; i < 3; i++)
        {
            Assert.True(breaker.TryEnter(out var lease));
            breaker.Failed(lease);
        }
        now += TimeSpan.FromSeconds(9);
        Assert.True(breaker.TryEnter(out var probe));
        breaker.Canceled(probe);
        Assert.True(breaker.TryEnter(out _));
        Assert.False(breaker.TryEnter(out _));
    }

    [Fact]
    public async Task ALongRetryAfterDoesNotOverflowOrRetryBeforeTheDeadline()
    {
        using var client = new HttpClient(new Handler((_, _) =>
        {
            var response = Response(429);
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromDays(100));
            return Task.FromResult(response);
        }));
        var transport = Transport(client, timeoutMs: 40);
        Assert.Equal("sec_timeout", (await Assert.ThrowsAsync<SecUnavailableException>(() => transport.GetCompanyFactsAsync("0000789019", default))).Code);
        Assert.Equal(1, transport.Requests);
        Assert.Equal(0, transport.Active);
    }

    [Fact]
    public async Task DeadlineCoversResponseBodyAndReleasesTheSlot()
    {
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new WaitingContent() })));
        var transport = Transport(client, timeoutMs: 40);
        for (var i = 0; i < 2; i++)
            Assert.Equal("sec_timeout", (await Assert.ThrowsAsync<SecUnavailableException>(() => transport.GetCompanyFactsAsync("0000789019", default))).Code);
        Assert.Equal(0, transport.Active);
    }

    [Fact]
    public async Task CallerCancellationIsNotAnOutageAndQueuedRequestsDoNotReachHttp()
    {
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var calls = 0;
        using var canceled = new CancellationTokenSource();
        using var client = new HttpClient(new Handler(async (_, token) =>
        {
            if (Interlocked.Increment(ref calls) == 4)
                entered.SetResult();
            await Task.Delay(Timeout.Infinite, token);
            return Response(200);
        }));
        var transport = Transport(client);
        var requests = Enumerable.Range(0, 10).Select(_ => transport.GetCompanyFactsAsync("0000789019", canceled.Token)).ToArray();
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(1));
        Assert.Equal(4, transport.Active);
        canceled.Cancel();
        foreach (var task in requests)
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
        Assert.Equal(4, transport.Requests);
        Assert.Equal(0, transport.Active);
        Assert.Equal("closed", transport.Circuit);
    }

    [Theory]
    [InlineData(ConnectionMode.Online)]
    [InlineData(ConnectionMode.Offline)]
    public void SnapshotReadsHonorCancellationInBothModes(ConnectionMode mode)
    {
        var state = new ConnectionState();
        state.SetMode(mode);
        var service = new FactService(new FinancialStore(Path.Combine(CoreTests.Root, "data/financials.json")), state);
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => service.Get(Revenue.Key, canceled.Token));
        Assert.Equal(0, service.Reads);
    }

    [Theory]
    [InlineData("missing_value")]
    [InlineData("null_value")]
    [InlineData("string_value")]
    [InlineData("missing_accession")]
    [InlineData("null_accession")]
    [InlineData("object_accession")]
    public async Task MalformedSecRecordsAreRejectedWithoutChangingSavedValues(string defect)
    {
        var previous = Revenue;
        var row = AnnualRow(previous);
        switch (defect)
        {
            case "missing_value":
                row.Remove("val");
                break;
            case "null_value":
                row["val"] = null;
                break;
            case "string_value":
                row["val"] = "123";
                break;
            case "missing_accession":
                row.Remove("accn");
                break;
            case "null_accession":
                row["accn"] = null;
                break;
            case "object_accession":
                row["accn"] = new JObject();
                break;
        }
        var json = CompanyFacts(previous, row);
        Assert.Throws<InvalidDataException>(() => SecClient.SelectAnnualFact(json, previous, "0000789019", DateTimeOffset.UtcNow));
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(json.ToString()) })));
        var sec = new SecClient(Transport(client));
        var error = await Assert.ThrowsAsync<SecUnavailableException>(() => sec.RefreshAsync(new Company { Cik = "0000789019" }, [previous], default));
        Assert.Equal("sec_invalid_data", error.Code);
        Assert.Equal(281724m, previous.Value);
    }

    [Fact]
    public async Task AReportedZeroRemainsValidAfterATransientHttpFailure()
    {
        var previous = Revenue;
        var row = AnnualRow(previous);
        row["val"] = 0;
        var calls = 0;
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(++calls == 1
            ? Response(503)
            : new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(CompanyFacts(previous, row).ToString()) })));
        var refreshed = await new SecClient(Transport(client)).RefreshAsync(new Company { Cik = "0000789019" }, [previous], default);
        Assert.Equal(2, calls);
        Assert.Equal(0m, Assert.Single(refreshed).Value);
        Assert.Equal(previous.SourceUrl, refreshed[0].SourceUrl);
        Assert.Equal(281724m, previous.Value);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("320193")]
    [InlineData("\"789019\"")]
    [InlineData("{}")]
    public async Task MissingOrMismatchedCompanyIdentityIsRejected(string identity)
    {
        var previous = Revenue;
        var json = CompanyFacts(previous, AnnualRow(previous));
        json["cik"] = JToken.Parse(identity);
        using var client = new HttpClient(new Handler((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(json.ToString()) })));
        var error = await Assert.ThrowsAsync<SecUnavailableException>(() => new SecClient(Transport(client)).RefreshAsync(new Company { Cik = "0000789019" }, [previous], default));
        Assert.Equal("sec_invalid_data", error.Code);
    }

    private static JObject AnnualRow(FinancialFact fact) => JObject.FromObject(new
    {
        form = "10-K",
        start = fact.Start,
        end = fact.End,
        filed = fact.Filed,
        accn = fact.Accession,
        val = fact.RawValue
    });

    private static JObject CompanyFacts(FinancialFact fact, JObject row) => new()
    {
        ["cik"] = 789019,
        ["facts"] = new JObject
        {
            ["us-gaap"] = new JObject
            {
                [fact.Concept.Split(':')[1]] = new JObject { ["units"] = new JObject { ["USD"] = new JArray(row) } }
            }
        }
    };

    private static HttpResponseMessage Response(int status) => new((HttpStatusCode)status) { Content = new StringContent(status == 200 ? "{}" : "sensitive provider payload") };
    private sealed class Handler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request, cancellationToken);
    }
    private sealed class WaitingContent : HttpContent
    {
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) => throw new InvalidOperationException("The body must be read with cancellation.");
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context, CancellationToken cancellationToken) => Task.Delay(Timeout.Infinite, cancellationToken);
        protected override bool TryComputeLength(out long length)
        {
            length = 0;
            return false;
        }
    }
}
