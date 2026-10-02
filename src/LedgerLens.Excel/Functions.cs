using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using LedgerLens.Core;

namespace LedgerLens.Excel
{
    public static class Functions
    {
        [ExcelFunction(Name = "LL.METRIC", Description = "Reported annual financials in USD millions; DilutedEPS in USD/share. Use LL.STATUS to inspect freshness.", Category = "LedgerLens")]
        public static object Metric(string ticker, string metric, string period, double revision = 0) =>
            ExcelAsyncUtil.RunTaskWithCancellation("LL.METRIC", new object[] { ticker, metric, period, revision }, ct => MetricAsync(ticker, metric, period, revision, ct));
        private static async Task<object> MetricAsync(string ticker, string metric, string period, double revision, CancellationToken cancellation)
        {
            try { var result = await HostRuntime.Client.FactAsync(new FactKey(ticker, metric, period), revision, cancellation).ConfigureAwait(false); return (double)result.Fact.Value; }
            catch (ArgumentException e) { HostRuntime.RecordError(e.Message); return ExcelError.ExcelErrorValue; }
            catch (OperationCanceledException) { HostRuntime.FormulaCanceled(); return ExcelError.ExcelErrorNA; }
            catch (Exception e) { HostRuntime.RecordError(e.Message); return ExcelError.ExcelErrorNA; }
        }
        [ExcelFunction(Name = "LL.SOURCE", Description = "Original SEC filing index for a reported value.", Category = "LedgerLens")]
        public static object Source(string ticker, string metric, string period, double revision = 0) =>
            ExcelAsyncUtil.RunTaskWithCancellation("LL.SOURCE", new object[] { ticker, metric, period, revision }, ct => SourceAsync(ticker, metric, period, revision, ct));
        private static async Task<object> SourceAsync(string ticker, string metric, string period, double revision, CancellationToken cancellation)
        {
            try { return (await HostRuntime.Client.FactAsync(new FactKey(ticker, metric, period), revision, cancellation).ConfigureAwait(false)).Fact.SourceUrl; }
            catch (OperationCanceledException) { HostRuntime.FormulaCanceled(); return ExcelError.ExcelErrorNA; }
            catch (Exception e) { HostRuntime.RecordError(e.Message); return ExcelError.ExcelErrorNA; }
        }
        [ExcelFunction(Name = "LL.STATUS", Description = "Freshness, units, filing date and source state. Historical snapshots are not live market prices.", Category = "LedgerLens")]
        public static object Status(string ticker, string metric, string period, double revision = 0) =>
            ExcelAsyncUtil.RunTaskWithCancellation("LL.STATUS", new object[] { ticker, metric, period, revision }, ct => StatusAsync(ticker, metric, period, revision, ct));
        private static async Task<string> StatusAsync(string ticker, string metric, string period, double revision, CancellationToken cancellation)
        {
            try { var result = await HostRuntime.Client.FactAsync(new FactKey(ticker, metric, period), revision, cancellation).ConfigureAwait(false); return result.Freshness.ToUpperInvariant() + " · " + result.Fact.Unit + " · filed " + result.Fact.Filed + " · " + result.Message; }
            catch (OperationCanceledException) { HostRuntime.FormulaCanceled(); return "Canceled"; }
            catch (Exception e) { return "Unavailable: " + e.Message; }
        }
        [ExcelFunction(Name = "LL.TABLE", Description = "Spill a company's nine sourced metrics: metric, value, units. USD millions except EPS.", Category = "LedgerLens")]
        public static object Table(string ticker, string period, double revision = 0) =>
            ExcelAsyncUtil.RunTaskWithCancellation("LL.TABLE", new object[] { ticker, period, revision }, ct => TableAsync(ticker, period, revision, ct));
        private static async Task<object> TableAsync(string ticker, string period, double revision, CancellationToken cancellation)
        {
            try
            {
                var tasks = new System.Collections.Generic.List<Task<FactResult>>();
                foreach (var metric in MetricCatalog.Labels.Keys) tasks.Add(HostRuntime.Client.FactAsync(new FactKey(ticker, metric, period), revision, cancellation));
                var facts = await Task.WhenAll(tasks).ConfigureAwait(false);
                var values = new object[facts.Length + 1, 3]; values[0, 0] = "Metric"; values[0, 1] = "Value"; values[0, 2] = "Units";
                for (var i = 0; i < facts.Length; i++) { values[i + 1, 0] = facts[i].Fact.Label; values[i + 1, 1] = (double)facts[i].Fact.Value; values[i + 1, 2] = facts[i].Fact.Unit; }
                return values;
            }
            catch (OperationCanceledException) { HostRuntime.FormulaCanceled(); return ExcelError.ExcelErrorNA; }
            catch (Exception e) { HostRuntime.RecordError(e.Message); return ExcelError.ExcelErrorNA; }
        }
        [ExcelFunction(Name = "LL.ASK", Description = "Explicit OpenAI research over supplied annual financial facts. May incur API charges; answers are cached.", Category = "LedgerLens")]
        public static object Ask(string ticker, string question) =>
            ExcelAsyncUtil.RunTaskWithCancellation("LL.ASK", new object[] { ticker, question }, ct => AskAsync(ticker, question, ct));
        private static async Task<object> AskAsync(string ticker, string question, CancellationToken cancellation)
        {
            try
            {
                var answer = await HostRuntime.Client.PostAsync<ResearchAnswer>("/api/research", new ResearchRequest { Ticker = ticker, Question = question, UseAi = true }, cancellation).ConfigureAwait(false);
                var text = new System.Text.StringBuilder(answer.Headline + "\n" + answer.Summary);
                foreach (var claim in answer.Claims) text.Append("\n• ").Append(claim.Text).Append(" [").Append(string.Join(", ", claim.SourceIds)).Append("]");
                text.Append("\nProvider: ").Append(answer.Provider).Append(". Use the research pane to inspect citations.");
                return text.ToString();
            }
            catch (OperationCanceledException) { HostRuntime.FormulaCanceled(); return ExcelError.ExcelErrorNA; }
            catch (Exception e) { HostRuntime.RecordError(e.Message); return "Research unavailable: " + e.Message; }
        }
        [ExcelFunction(Name = "LL.LIVE", Description = "Live service notifications over WebSocket. Replays are labeled; this is not a market-price feed.", Category = "LedgerLens")]
        public static IObservable<object> Live() => HostRuntime.Feed;
        [ExcelFunction(Name = "LL.VERSION", Description = "LedgerLens add-in version.", Category = "LedgerLens", IsThreadSafe = true)]
        public static string Version() => "LedgerLens 1.0.0";
        [ExcelFunction(Name = "LL.DIAGNOSTICS", Description = "Recent host diagnostics for integration validation.", IsHidden = true)]
        public static string Diagnostics() => string.Join("\n", HostRuntime.Errors);
        [ExcelFunction(Name = "LL.CANCELLATIONS", Description = "Canceled formula waiters for integration validation.", IsHidden = true, IsThreadSafe = true)]
        public static double Cancellations() => HostRuntime.CanceledFormulas;
    }
}
