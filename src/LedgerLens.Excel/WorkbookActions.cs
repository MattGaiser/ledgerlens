using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using LedgerLens.Core;
using ExcelDna.Integration;
using Xl = Microsoft.Office.Interop.Excel;

namespace LedgerLens.Excel
{
    internal sealed class WorkbookActions
    {
        private readonly Dictionary<string, RefreshPlan> previews = new Dictionary<string, RefreshPlan>();
        private readonly Dictionary<string, AppliedTransaction> applied = new Dictionary<string, AppliedTransaction>();
        private readonly WorkbookSessions sessions = new WorkbookSessions();
        private sealed class AppliedTransaction
        {
            internal RefreshPlan Plan
            {
                get;
            }
            internal (string Address, string Before, string After)[] Writes
            {
                get;
            }
            internal AppliedTransaction(RefreshPlan plan, (string Address, string Before, string After)[] writes)
            {
                Plan = plan;
                Writes = writes;
            }
        }
        internal static Xl.Application App => (Xl.Application)ExcelDnaUtil.Application;

        public async Task<RefreshPlan> PreviewAsync(int windowId, CancellationToken cancellation)
        {
            var capture = await HostRuntime.OnExcelThread(() => { RequireWindow(windowId); return Capture(); }, cancellation).ConfigureAwait(false);
            var plan = await HostRuntime.Client.PostAsync<RefreshPlan>("/api/refresh/preview", new
            {
                workbookId = capture.WorkbookId,
                cells = capture.Cells,
                dependencies = capture.Dependencies
            }, cancellation).ConfigureAwait(false);
            await HostRuntime.OnExcelThread(() =>
            {
                RequireWindow(windowId);
                using (var scope = new ComScope())
                {
                    var book = scope.Own(App.ActiveWorkbook);
                    if (book == null || Identity(book, scope) != plan.WorkbookId)
                        throw new InvalidOperationException("The active workbook changed while the preview loaded. Preview again in the intended workbook.");
                    previews[plan.WorkbookId] = plan;
                    return true;
                }
            }, cancellation).ConfigureAwait(false);
            return plan;
        }
        public object Apply(string planId)
        {
            using (var scope = new ComScope())
            {
                var book = scope.Own(App.ActiveWorkbook) ?? throw new InvalidOperationException("Open your model first.");
                var id = Identity(book, scope);
                if (!previews.TryGetValue(id, out var plan) || plan.Id != planId)
                    throw new InvalidOperationException("The preview is no longer valid. Preview the model again.");
                if (DateTimeOffset.UtcNow - plan.CreatedAt > TimeSpan.FromMinutes(10))
                    throw new InvalidOperationException("The preview expired. Preview the model again.");
                if (book.ReadOnly)
                    throw new InvalidOperationException("This workbook is read-only. Save an editable copy first.");
                var conflicts = RefreshPlanner.Conflicts(plan, id, address => ReadContent(book, address, scope));
                if (conflicts.Length != 0)
                    throw new InvalidOperationException("Nothing was changed. Cells edited since preview: " + string.Join(", ", conflicts) + ". Preview again.");
                var changes = plan.Changes.Select(c => (c.Address, c.ExpectedContent, RefreshPlanner.EncodeNumber(c.NewValue))).ToList();
                changes.Add(("_LedgerLens!B3", plan.Dependencies["_LedgerLens!B3"], plan.Dependencies["Model!B4"]));
                changes.AddRange(AuditWrites(book, plan, scope));
                var writes = changes.ToArray();
                WriteTransaction(book, writes, scope);
                previews.Remove(id);
                applied[id] = new AppliedTransaction(plan, writes);
                Recalculate();
                return new
                {
                    changed = plan.Changes.Length,
                    planId = plan.Id
                };
            }
        }
        public object Rollback()
        {
            using (var scope = new ComScope())
            {
                var book = scope.Own(App.ActiveWorkbook) ?? throw new InvalidOperationException("Open your model first.");
                var id = Identity(book, scope);
                if (!applied.TryGetValue(id, out var transaction))
                    throw new InvalidOperationException("There is no model update to undo in this workbook session.");
                var plan = transaction.Plan;
                if (book.ReadOnly)
                    throw new InvalidOperationException("This workbook is read-only.");
                var conflicts = RefreshPlanner.RollbackConflicts(plan, address => ReadContent(book, address, scope));
                if (ReadContent(book, "_LedgerLens!B3", scope) != plan.Dependencies["Model!B4"])
                    throw new InvalidOperationException("Nothing was changed. The imported company marker was edited after the update.");
                if (conflicts.Length != 0)
                    throw new InvalidOperationException("Nothing was changed. Undo would overwrite later edits in " + string.Join(", ", conflicts) + ".");
                var auditConflicts = transaction.Writes.Where(c => ReadContent(book, c.Address, scope) != c.After).Select(c => c.Address).ToArray();
                if (auditConflicts.Length != 0)
                    throw new InvalidOperationException("Nothing was changed. The import audit or model was edited after the update: " + string.Join(", ", auditConflicts.Take(5)));
                WriteTransaction(book, transaction.Writes.Select(c => (c.Address, c.After, c.Before)).ToArray(), scope);
                applied.Remove(id);
                previews.Remove(id);
                Recalculate();
                return new
                {
                    changed = plan.Changes.Length
                };
            }
        }
        private static (string Address, string Before, string After)[] AuditWrites(Xl.Workbook book, RefreshPlan plan, ComScope scope)
        {
            var sourceSheet = scope.Own((Xl.Worksheet)scope.Own(book.Worksheets)["Sources"]);
            var used = scope.Own(sourceSheet.UsedRange);
            var firstRow = used.Row + scope.Own(used.Rows).Count + 3;
            if (firstRow + plan.Changes.Length + 2 > 1048576)
                throw new InvalidOperationException("The source audit worksheet is full. Archive it before importing more values.");
            var rows = new List<object[]> { new object[] { "REVIEWED IMPORT " + plan.Id, plan.CreatedAt.ToString("u", CultureInfo.InvariantCulture) }, new object[] { "Model cell", "Ticker", "Metric", "Fiscal year", "Value", "Units", "Period start", "Period end", "Filed", "SEC accession", "Evidence ID", "SEC filing", "Raw value", "XBRL concept", "Acquired UTC" } };
            foreach (var change in plan.Changes)
            {
                var f = change.Source;
                rows.Add(new object[] { change.Address, f.Ticker, f.Label, f.Period, f.Value, f.Unit, f.Start ?? "Instant", f.End, f.Filed, f.Accession, f.SourceId, f.SourceUrl, f.RawValue, f.Concept, f.AcquiredAt.ToString("u", CultureInfo.InvariantCulture) });
            }
            var result = new List<(string Address, string Before, string After)>();
            for (var r = 0; r < rows.Count; r++)
                for (var c = 0; c < rows[r].Length; c++)
                {
                    var address = "Sources!" + (char)('A' + c) + (firstRow + r);
                    var after = rows[r][c] is decimal number ? RefreshPlanner.EncodeNumber(number) : "s:" + Convert.ToString(rows[r][c], CultureInfo.InvariantCulture);
                    result.Add((address, ReadContent(book, address, scope), after));
                }
            return result.ToArray();
        }
        private static void WriteTransaction(Xl.Workbook book, (string Address, string Before, string After)[] changes, ComScope scope)
        {
            var ranges = changes.Select(c => Range(book, c.Address, scope)).ToArray();
            foreach (var range in ranges)
            {
                var sheet = scope.Own((Xl.Worksheet)range.Worksheet);
                if (sheet.ProtectContents || (bool)range.MergeCells)
                    throw new InvalidOperationException("Tracked historical cells must be unprotected and unmerged. Nothing was changed.");
            }
            using (new ExcelStateGuard(App))
            {
                var completed = 0;
                try
                {
                    for (; completed < changes.Length; completed++)
                        SetContent(ranges[completed], changes[completed].After);
                }
                catch
                {
                    var failedRestores = new List<string>();
                    for (var i = Math.Min(completed, changes.Length - 1); i >= 0; i--)
                    {
                        try
                        {
                            SetContent(ranges[i], changes[i].Before);
                        }
                        catch { failedRestores.Add(changes[i].Address); }
                    }
                    if (failedRestores.Count > 0)
                        throw new InvalidOperationException("Excel interrupted the update and these cells could not be restored: " + string.Join(", ", failedRestores) + ". Reopen the last saved workbook.");
                    throw;
                }
            }
        }
        public object InsertFormula(string ticker, string metric, string period)
        {
            var key = new FactKey(ticker, metric, period);
            using (var scope = new ComScope())
            {
                var cell = scope.Own(App.Selection as Xl.Range) ?? throw new InvalidOperationException("Select an empty worksheet cell first.");
                if (Convert.ToDouble(cell.CountLarge, CultureInfo.InvariantCulture) != 1 || cell.Value2 != null || (bool)cell.HasFormula || (bool)cell.MergeCells)
                    throw new InvalidOperationException("Select one empty, unmerged cell. Existing content is never overwritten by Insert.");
                var sheet = scope.Own((Xl.Worksheet)cell.Worksheet);
                var book = scope.Own((Xl.Workbook)sheet.Parent);
                if (book.ReadOnly)
                    throw new InvalidOperationException("This workbook is read-only. Save an editable copy first.");
                if (sheet.ProtectContents)
                    throw new InvalidOperationException("The selected worksheet is protected.");
                cell.Formula = "=LL.METRIC(\"" + key.Ticker + "\",\"" + key.Metric + "\",\"" + key.Period + "\")";
                cell.NumberFormat = key.Metric == "DilutedEPS" ? "0.00" : "#,##0;[Red](#,##0);–";
                return new
                {
                    address = cell.Address[false, false]
                };
            }
        }
        public object Recalculate()
        {
            HostRuntime.Client.ClearCache();
            using (var scope = new ComScope())
            {
                var book = scope.Own(App.ActiveWorkbook);
                if (book != null)
                {
                    try
                    {
                        var hidden = scope.Own((Xl.Worksheet)scope.Own(book.Worksheets)["_LedgerLens"]);
                        var revision = scope.Own(hidden.Range["B2"]);
                        revision.Value2 = Convert.ToDouble(revision.Value2 ?? 0, CultureInfo.InvariantCulture) + 1;
                    }
                    catch (System.Runtime.InteropServices.COMException) { }
                }
            }
            App.CalculateFull();
            return new
            {
                recalculated = true
            };
        }
        public object SaveResearch(ResearchAnswer answer)
        {
            ResearchValidation.Validate(answer, HostRuntime.Client.Snapshot.Facts);
            ResearchValidation.ValidateExport(answer);
            using (var scope = new ComScope())
            {
                var book = scope.Own(App.ActiveWorkbook) ?? throw new InvalidOperationException("Open a workbook first.");
                if (book.ReadOnly || book.ProtectStructure)
                    throw new InvalidOperationException("Use an editable workbook with unprotected structure.");
                var sheets = scope.Own(book.Worksheets);
                var sheet = scope.Own((Xl.Worksheet)sheets.Add(After: scope.Own(sheets[sheets.Count])));
                sheet.Name = "Research " + DateTime.Now.ToString("HHmmss") + " " + Guid.NewGuid().ToString("N").Substring(0, 4);
                var rows = new List<object[]> { new object[] { "LEDGERLENS RESEARCH", "" }, new object[] { answer.Headline, answer.Provider }, new object[] { answer.Summary, answer.Model }, new object[] { "", "" }, new object[] { "Claim", "Evidence IDs" } };
                rows.AddRange(answer.Claims.Select(c => new object[] { c.Text, string.Join("; ", c.SourceIds) }));
                rows.Add(new object[] { "CAVEATS", "" });
                rows.AddRange(answer.Caveats.Select(c => new object[] { c, "" }));
                rows.Add(new object[] { "SOURCE", "SEC filing" });
                rows.AddRange(answer.Sources.Select(f => new object[] { f.SourceId, f.SourceUrl }));
                var values = new object[rows.Count, 2];
                for (var i = 0; i < rows.Count; i++)
                    for (var j = 0; j < 2; j++)
                        values[i, j] = "'" + Convert.ToString(rows[i][j], CultureInfo.InvariantCulture);
                var range = scope.Own(sheet.Range["A1", "B" + rows.Count]);
                range.NumberFormat = "@";
                range.Value2 = values;
                range.WrapText = true;
                scope.Own(sheet.Range["A:A"]).ColumnWidth = 95;
                scope.Own(sheet.Range["B:B"]).ColumnWidth = 75;
                scope.Own(range.Rows).AutoFit();
                WorkbookBuilder.Title(sheet, "A1:B1", scope);
                return new
                {
                    sheet = sheet.Name
                };
            }
        }
        public static void RequireWindow(int windowId)
        {
            var window = App.ActiveWindow;
            if (window == null || window.Hwnd != windowId)
                throw new InvalidOperationException("Activate the workbook containing this pane, then try again. No other workbook was changed.");
        }
        private (string WorkbookId, ModelCell[] Cells, Dictionary<string, string> Dependencies) Capture()
        {
            using (var scope = new ComScope())
            {
                var book = scope.Own(App.ActiveWorkbook) ?? throw new InvalidOperationException("Open the LedgerLens model workbook first.");
                var id = Identity(book, scope);
                var cells = new List<ModelCell>();
                var model = scope.Own((Xl.Worksheet)scope.Own(book.Worksheets)["Model"]);
                var ticker = Convert.ToString(scope.Own(model.Range["B4"]).Value2, CultureInfo.InvariantCulture) ?? "MSFT";
                var period = Convert.ToString(scope.Own(model.Range["E8"]).Value2, CultureInfo.InvariantCulture) ?? "";
                var dependencies = new Dictionary<string, string>();
                foreach (var address in new[] { "Model!B4", "Model!E8", "_LedgerLens!B3" })
                    dependencies[address] = ReadContent(book, address, scope);
                var metrics = MetricCatalog.Labels.Keys.ToArray();
                for (var i = 0; i < metrics.Length; i++)
                {
                    var address = "Model!E" + (10 + i);
                    var cell = Range(book, address, scope);
                    if ((bool)cell.HasFormula)
                        throw new InvalidOperationException("Historical import cells must contain reported values or be empty. Move the formula in " + address + " outside E10:E18 before importing. Nothing was changed.");
                    var metricAddress = "Model!B" + (10 + i);
                    dependencies[metricAddress] = ReadContent(book, metricAddress, scope);
                    var metric = Convert.ToString(Range(book, metricAddress, scope).Value2, CultureInfo.InvariantCulture) ?? "";
                    cells.Add(new ModelCell { Address = address, Content = Encode(cell), Ticker = ticker, Metric = metric, Period = period });
                    foreach (var column in new[] { "F", "G" })
                    {
                        var forecast = Range(book, "Model!" + column + (10 + i), scope);
                        cells.Add(new ModelCell { Address = "Model!" + column + (10 + i), Content = Encode(forecast), HasFormula = (bool)forecast.HasFormula, IsAnalystInput = true });
                    }
                }
                foreach (var address in new[] { "Model!F5", "Model!G5", "Model!F6", "Model!G6" })
                    cells.Add(new ModelCell { Address = address, Content = ReadContent(book, address, scope), IsAnalystInput = true });
                return (id, cells.ToArray(), dependencies);
            }
        }
        internal string Identity(Xl.Workbook book, ComScope scope)
        {
            try
            {
                var sheet = scope.Own((Xl.Worksheet)scope.Own(book.Worksheets)["_LedgerLens"]);
                if (Convert.ToString(scope.Own(sheet.Range["A1"]).Value2, CultureInfo.InvariantCulture) != "LedgerLens/1")
                    throw new InvalidOperationException("This workbook does not have a LedgerLens model map.");
                var documentId = Convert.ToString(scope.Own(sheet.Range["B1"]).Value2, CultureInfo.InvariantCulture) ?? "";
                return sessions.GetId(book, documentId);
            }
            catch (System.Runtime.InteropServices.COMException) { throw new InvalidOperationException("Open the included LedgerLens model, or create one from the LedgerLens ribbon."); }
        }
        private static Xl.Range Range(Xl.Workbook book, string address, ComScope scope)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(address, @"^(?:Model!(?:B4|E8|B1[0-8]|[EFG](?:[5-6]|1[0-8]))|_LedgerLens!B3|Sources![A-O][1-9][0-9]{0,6})$"))
                throw new InvalidOperationException("The address is outside the tracked model area.");
            var parts = address.Split('!');
            return scope.Own(scope.Own((Xl.Worksheet)scope.Own(book.Worksheets)[parts[0]]).Range[parts[1]]);
        }
        private static string ReadContent(Xl.Workbook book, string address, ComScope scope) => Encode(Range(book, address, scope));
        private static string Encode(Xl.Range cell)
        {
            if ((bool)cell.HasFormula)
                return "f:" + Convert.ToString(cell.Formula, CultureInfo.InvariantCulture);
            var value = cell.Value2;
            if (value == null)
                return "e:";
            if (value is double || value is int || value is decimal)
                return RefreshPlanner.EncodeNumber(Convert.ToDecimal(value, CultureInfo.InvariantCulture));
            return "s:" + Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        private static void SetContent(Xl.Range cell, string content)
        {
            if (content.StartsWith("n:", StringComparison.Ordinal))
                cell.Value2 = (double)decimal.Parse(content.Substring(2), CultureInfo.InvariantCulture);
            else if (content == "e:")
                cell.ClearContents();
            else if (content.StartsWith("f:", StringComparison.Ordinal))
                cell.Formula = content.Substring(2);
            else
                cell.Value2 = "'" + content.Substring(2);
        }
        public void Forget(Xl.Workbook book)
        {
            if (sessions.Remove(book, out var id))
            {
                previews.Remove(id);
                applied.Remove(id);
            }
        }
        public void Clear()
        {
            previews.Clear();
            applied.Clear();
            sessions.Clear();
        }
    }
    internal sealed class ExcelStateGuard : IDisposable
    {
        private readonly Xl.Application app;
        private readonly bool updating;
        private readonly bool events;
        private readonly Xl.XlCalculation calculation;
        public ExcelStateGuard(Xl.Application app)
        {
            this.app = app;
            updating = app.ScreenUpdating;
            events = app.EnableEvents;
            calculation = app.Calculation;
            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Xl.XlCalculation.xlCalculationManual;
            }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            try
            {
                app.Calculation = calculation;
            }
            finally { try { app.EnableEvents = events; } finally { app.ScreenUpdating = updating; } }
        }
    }
}
