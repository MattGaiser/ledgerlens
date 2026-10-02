using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using LedgerLens.Core;
using Xl = Microsoft.Office.Interop.Excel;

namespace LedgerLens.Excel
{
    public sealed class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                HostRuntime.Start();
            }
            catch (Exception e) { HostRuntime.RecordError(e.Message); }
            ExcelIntegration.RegisterUnhandledExceptionHandler(e => { HostRuntime.RecordError((e as Exception)?.Message ?? "Unexpected add-in error."); return ExcelError.ExcelErrorValue; });
        }
        public void AutoClose() => HostRuntime.Stop();
    }
    internal static class HostRuntime
    {
        private static SessionClient? client;
        private static FeedObservable? feed;
        private static Xl.Application? application;
        private static string? logPath;
        private static readonly object logGate = new object();
        private static long canceledFormulas;
        internal static long CanceledFormulas => Interlocked.Read(ref canceledFormulas);
        internal static void FormulaCanceled() => Interlocked.Increment(ref canceledFormulas);
        private static readonly ConcurrentQueue<string> errors = new ConcurrentQueue<string>();
        private static readonly System.Collections.Generic.Dictionary<int, PaneHandle> panes = new System.Collections.Generic.Dictionary<int, PaneHandle>();
        internal static SessionClient Client => client ?? throw new InvalidOperationException("Start LedgerLens using the included launcher, then reopen the add-in.");
        internal static FeedObservable Feed => feed ?? throw new InvalidOperationException("Research notification service is not connected.");
        internal static WorkbookActions Actions { get; } = new WorkbookActions();
        internal static void Start()
        {
            if (client != null)
                return;
            var root = Environment.GetEnvironmentVariable("LEDGERLENS_ROOT");
            if (string.IsNullOrWhiteSpace(root))
            {
                var dir = new DirectoryInfo(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!);
                while (dir != null && !File.Exists(Path.Combine(dir.FullName, "data", "financials.json")))
                    dir = dir.Parent;
                root = dir?.FullName;
            }
            if (root == null)
                throw new InvalidOperationException("LedgerLens evidence folder could not be found.");
            logPath = Path.Combine(root, ".runtime", "excel.log");
            client = new SessionClient(root);
            feed = new FeedObservable(client.Endpoint);
            application = (Xl.Application)ExcelDnaUtil.Application;
            application.WorkbookBeforeClose += OnWorkbookBeforeClose;
        }
        private static void OnWorkbookBeforeClose(Xl.Workbook book, ref bool cancel)
        {
            Actions.Forget(book);
            using (var scope = new ComScope())
            {
                var windows = scope.Own(book.Windows);
                for (var index = 1; index <= windows.Count; index++)
                {
                    var window = scope.Own(windows[index]);
                    if (panes.TryGetValue(window.Hwnd, out var pane))
                    {
                        pane.Dispose();
                        panes.Remove(window.Hwnd);
                    }
                }
            }
        }
        internal static void ShowPane(string tab = "overview")
        {
            Start();
            using (var scope = new ComScope())
            {
                var app = (Xl.Application)ExcelDnaUtil.Application;
                var window = scope.Own(app.ActiveWindow);
                if (window == null)
                    throw new InvalidOperationException("Open a workbook first.");
                if (!panes.TryGetValue(window.Hwnd, out var handle))
                {
                    var control = new ResearchPane(window.Hwnd);
                    var pane = CustomTaskPaneFactory.CreateCustomTaskPane(control, "LedgerLens", window);
                    using (var graphics = control.CreateGraphics())
                        pane.Width = (int)Math.Round(430 * graphics.DpiX / 96.0);
                    pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    handle = new PaneHandle(pane, control);
                    panes[window.Hwnd] = handle;
                }
                handle.Pane.Visible = true;
                handle.Control.NavigateTo(tab);
            }
        }
        internal static Task<T> OnExcelThread<T>(Func<T> action, CancellationToken cancellation = default)
        {
            return QueuedAction.Run(callback => ExcelAsyncUtil.QueueAsMacro(() => callback()), action, cancellation);
        }
        internal static void RecordError(string error)
        {
            RecordActivity("Error: " + error);
            errors.Enqueue(DateTimeOffset.UtcNow.ToString("o") + " " + error);
            while (errors.Count > 50)
                errors.TryDequeue(out _);
        }
        internal static void RecordActivity(string message)
        {
            if (logPath == null)
                return;
            try
            {
                lock (logGate)
                    File.AppendAllText(logPath, DateTimeOffset.UtcNow.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        internal static string[] Errors => errors.ToArray();
        internal static void Stop()
        {
            RecordActivity("Host: disconnecting");
            if (application != null)
            {
                application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                application = null;
            }
            foreach (var pane in panes.Values)
            {
                try
                {
                    pane.Dispose();
                }
                catch (Exception e) { RecordError("Pane cleanup: " + e.GetType().Name); }
            }
            panes.Clear();
            Actions.Clear();
            feed?.Dispose();
            feed = null;
            client?.Dispose();
            client = null;
            RecordActivity("Host: disconnected");
        }
        private sealed class PaneHandle : IDisposable
        {
            public CustomTaskPane Pane
            {
                get;
            }
            public ResearchPane Control
            {
                get;
            }
            public PaneHandle(CustomTaskPane pane, ResearchPane control)
            {
                Pane = pane;
                Control = control;
            }
            private bool disposed;
            public void Dispose()
            {
                if (disposed)
                    return;
                disposed = true;
                try
                {
                    Pane.Delete();
                }
                catch (System.Runtime.InteropServices.COMException) { }
                finally { Control.Dispose(); }
            }
        }
    }
    internal sealed class ComScope : IDisposable
    {
        private readonly System.Collections.Generic.List<object> owned = new System.Collections.Generic.List<object>();
        public T Own<T>(T item)
        {
            if (item != null && System.Runtime.InteropServices.Marshal.IsComObject(item))
                owned.Add(item);
            return item;
        }
        public void Dispose()
        {
            // These RCWs belong to Excel's main-thread apartment. Manually releasing
            // one can invalidate references shared by Excel-DNA or another callback.
            // https://excel-dna.net/docs/guides-basic/excel-programming-interfaces/using-the-excel-com-automation-interfaces/
            owned.Clear();
        }
    }
}
