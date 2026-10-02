using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using LedgerLens.Core;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LedgerLens.Excel
{
    [ComVisible(true)]
    public sealed class ResearchPane : UserControl
    {
        private readonly WebView2 browser = new WebView2();
        private readonly int windowId;
        private readonly CancellationTokenSource lifetime = new CancellationTokenSource();
        private string desiredTab = "overview";
        private bool ready;
        private bool initialized;
        private int disposalStarted;
        public ResearchPane(int windowId)
        {
            this.windowId = windowId;
            BackColor = System.Drawing.Color.FromArgb(245, 247, 247);
            browser.Dock = DockStyle.Fill; Controls.Add(browser);
            Load += async (_, __) => await InitializeAsync();
        }
        private async Task InitializeAsync()
        {
            if (initialized || IsDisposed) return; initialized = true;
            try
            {
                var profile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LedgerLens", "WebView2");
                var environment = await CoreWebView2Environment.CreateAsync(null, profile);
                if (IsDisposed) return;
                await browser.EnsureCoreWebView2Async(environment);
                if (IsDisposed) return;
                browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                browser.CoreWebView2.Settings.AreDevToolsEnabled = false;
                browser.CoreWebView2.Settings.IsStatusBarEnabled = false;
                browser.CoreWebView2.Settings.AreHostObjectsAllowed = false;
                browser.CoreWebView2.WebMessageReceived += HandleMessage;
                browser.CoreWebView2.NavigationStarting += (_, args) => { if (!IsLocal(args.Uri)) args.Cancel = true; };
                browser.CoreWebView2.NewWindowRequested += (_, args) =>
                {
                    args.Handled = true;
                    if (Uri.TryCreate(args.Uri, UriKind.Absolute, out var uri) && uri.Scheme == "https" && uri.Host == "www.sec.gov") Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
                };
                browser.CoreWebView2.PermissionRequested += (_, args) => args.State = CoreWebView2PermissionState.Deny;
                browser.CoreWebView2.Navigate(HostRuntime.Client.Endpoint.BaseUrl);
                HostRuntime.RecordActivity("Research pane: browser initialized");
            }
            catch (Exception e)
            {
                HostRuntime.RecordError("Research pane: " + e.Message);
                if (!IsDisposed) { Controls.Clear(); Controls.Add(new Label { Dock = DockStyle.Fill, Text = "The research pane could not start.\r\n\r\nRun Start LedgerLens again and check that the Microsoft Edge WebView2 Runtime is installed.", Padding = new Padding(24), AutoSize = false }); }
            }
        }
        private bool IsLocal(string uri) => Uri.TryCreate(uri, UriKind.Absolute, out var parsed) && parsed.GetLeftPart(UriPartial.Authority) == HostRuntime.Client.Endpoint.BaseUrl;
        private async void HandleMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            if (!IsLocal(args.Source) || IsDisposed) return;
            string? id = null;
            try
            {
                var message = JObject.Parse(args.WebMessageAsJson);
                id = message.Value<string>("id"); var command = message.Value<string>("command"); var payload = message["payload"] as JObject ?? new JObject();
                if (command == "ready") { ready = true; Post(new { type = "bootstrap", token = HostRuntime.Client.Endpoint.Token }); Post(new { type = "navigate", tab = desiredTab }); return; }
                object result;
                switch (command)
                {
                    case "insertFormula": result = await InWorkbook(() => HostRuntime.Actions.InsertFormula(payload.Value<string>("ticker") ?? "", payload.Value<string>("metric") ?? "", payload.Value<string>("period") ?? "")); break;
                    case "previewRefresh": result = await HostRuntime.Actions.PreviewAsync(windowId, lifetime.Token); break;
                    case "applyRefresh": result = await InWorkbook(() => HostRuntime.Actions.Apply(payload.Value<string>("planId") ?? "")); break;
                    case "rollbackRefresh": result = await InWorkbook(HostRuntime.Actions.Rollback); break;
                    case "recalculate": result = await InWorkbook(HostRuntime.Actions.Recalculate); break;
                    case "saveResearch": var answer = payload["answer"]?.ToObject<ResearchAnswer>() ?? throw new ArgumentException("Missing research answer."); result = await InWorkbook(() => HostRuntime.Actions.SaveResearch(answer)); break;
                    default: throw new ArgumentException("Unknown workbook action.");
                }
                Post(new { id, result });
            }
            catch (OperationCanceledException) when (Volatile.Read(ref disposalStarted) != 0) { }
            catch (Exception e) { HostRuntime.RecordError(e.Message); Post(new { id, error = e.Message }); }
        }
        private Task<T> InWorkbook<T>(Func<T> action) => HostRuntime.OnExcelThread(() => { WorkbookActions.RequireWindow(windowId); return action(); }, lifetime.Token);
        internal void NavigateTo(string tab) { desiredTab = tab; if (ready) Post(new { type = "navigate", tab }); }
        internal void ShowSource(FinancialFact fact) { if (ready) Post(new { type = "source", fact }); }
        private void Post(object message)
        {
            // Excel callbacks do not guarantee a WinForms synchronization context.
            // An await may therefore resume on the pool even though the event began on the UI thread.
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action(() => Post(message))); }
                catch (InvalidOperationException) when (IsDisposed || !IsHandleCreated) { }
                return;
            }
            if (IsDisposed || browser.IsDisposed || browser.CoreWebView2 == null) return;
            browser.CoreWebView2.PostWebMessageAsJson(JsonConvert.SerializeObject(message, new JsonSerializerSettings { ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver() }));
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && Interlocked.Exchange(ref disposalStarted, 1) == 0)
            {
                lifetime.Cancel();
                if (browser.CoreWebView2 != null) browser.CoreWebView2.WebMessageReceived -= HandleMessage;
                browser.Dispose(); lifetime.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
