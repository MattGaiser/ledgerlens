using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using LedgerLens.Core;
using Xl = Microsoft.Office.Interop.Excel;

namespace LedgerLens.Excel
{
    internal static class WorkbookBuilder
    {
        private const int Navy = 3944212;
        private static readonly int Teal = ColorTranslator.ToOle(Color.FromArgb(18, 125, 114));
        private static readonly int Pale = ColorTranslator.ToOle(Color.FromArgb(236, 244, 240));
        private static readonly int Muted = ColorTranslator.ToOle(Color.FromArgb(100, 118, 128));
        private static readonly int Blue = ColorTranslator.ToOle(Color.FromArgb(48, 100, 180));
        public static void Create()
        {
            var app = WorkbookActions.App;
            using (var scope = new ComScope())
            {
                var books = scope.Own(app.Workbooks); var book = scope.Own(books.Add(Xl.XlWBATemplate.xlWBATWorksheet));
                try
                {
                  using (new ExcelStateGuard(app))
                  {
                    HostRuntime.RecordActivity("Workbook: initializing sheets");
                    var sheets = scope.Own(book.Worksheets);
                    var dashboard = scope.Own((Xl.Worksheet)sheets[1]); dashboard.Name = "Dashboard";
                    var model = AddSheet(sheets, "Model", scope);
                    var sources = AddSheet(sheets, "Sources", scope);
                    var guide = AddSheet(sheets, "Guide", scope);
                    var control = AddSheet(sheets, "_LedgerLens", scope);
                    foreach (var sheet in new[] { dashboard, model, sources, guide, control })
                    {
                        var cells = scope.Own(sheet.Cells); var font = scope.Own(cells.Font); font.Name = "Aptos"; font.Size = 11; font.Color = Navy;
                        scope.Own(sheet.Columns).ColumnWidth = 14;
                        sheet.Activate(); var window = scope.Own(app.ActiveWindow); window.DisplayGridlines = false; window.Zoom = 90;
                    }
                    Cell(control, "A1", "LedgerLens/1", scope); Cell(control, "B1", Guid.NewGuid().ToString("N"), scope);
                    Cell(control, "A2", "Formula revision", scope); Cell(control, "B2", 0, scope);
                    Cell(control, "A3", "Imported company", scope);
                    var names = scope.Own(book.Names); scope.Own(names.Add("_LL_REVISION", "='_LedgerLens'!$B$2"));
                    scope.Own(names.Add("_LL_IMPORTED_TICKER", "='_LedgerLens'!$B$3"));
                    control.Visible = Xl.XlSheetVisibility.xlSheetVeryHidden;
                    HostRuntime.RecordActivity("Workbook: dashboard"); BuildDashboard(dashboard, scope);
                    HostRuntime.RecordActivity("Workbook: operating model"); BuildModel(model, scope);
                    HostRuntime.RecordActivity("Workbook: evidence register"); BuildSources(sources, scope);
                    HostRuntime.RecordActivity("Workbook: guide"); BuildGuide(guide, scope);
                    dashboard.Activate(); scope.Own(dashboard.Range["A1"]).Select();
                    // Workbook title and sheets carry the independent demo identity; no proprietary content is used.
                    book.CheckCompatibility = false;
                  }
                  HostRuntime.RecordActivity("Workbook: ready");
                }
                catch (Exception e) { HostRuntime.RecordError("Workbook creation: " + e); book.Close(false); throw; }
            }
        }
        private static Xl.Worksheet AddSheet(Xl.Sheets sheets, string name, ComScope scope)
        { var last = scope.Own(sheets[sheets.Count]); var sheet = scope.Own((Xl.Worksheet)sheets.Add(After: last)); sheet.Name = name; return sheet; }
        private static void BuildDashboard(Xl.Worksheet sheet, ComScope scope)
        {
            scope.Own(sheet.Range["A:A"]).ColumnWidth = 3;
            scope.Own(sheet.Range["B:B"]).ColumnWidth = 23;
            scope.Own(sheet.Range["C:H"]).ColumnWidth = 16;
            Merged(sheet, "B2:H3", "LedgerLens  /  Analyst workspace", scope); Title(sheet, "B2:H3", scope);
            Merged(sheet, "B5:H5", "REPORTED FUNDAMENTALS, WITH THE EVIDENCE IN REACH", scope); Style(sheet.Range["B5:H5"], null, Teal, true, 10, scope);
            Merged(sheet, "B6:H6", "Public SEC filings · USD millions unless noted · FY2025 is historical, not a live market quote", scope); Style(sheet.Range["B6:H6"], null, Muted, false, 10, scope);
            var headers = new object[,] { { "Company", "Ticker", "Revenue", "Revenue YoY", "Operating margin", "OCF less capex", "Diluted EPS" } };
            scope.Own(sheet.Range["B9:H9"]).Value2 = headers; Header(sheet, "B9:H9", scope);
            var companies = HostRuntime.Client.Snapshot.Companies;
            for (var i = 0; i < companies.Length; i++)
            {
                var row = 10 + i;
                Cell(sheet, "B" + row, companies[i].Name, scope); Cell(sheet, "C" + row, companies[i].Ticker, scope);
                Formula(sheet, "D" + row, "=LL.METRIC(C" + row + ",\"Revenue\",\"FY2025\",_LL_REVISION)", scope);
                Formula(sheet, "E" + row, "=IFERROR(D" + row + "/LL.METRIC(C" + row + ",\"Revenue\",\"FY2024\",_LL_REVISION)-1,NA())", scope);
                Formula(sheet, "F" + row, "=IFERROR(LL.METRIC(C" + row + ",\"OperatingIncome\",\"FY2025\",_LL_REVISION)/D" + row + ",NA())", scope);
                Formula(sheet, "G" + row, "=LL.METRIC(C" + row + ",\"OperatingCashFlow\",\"FY2025\",_LL_REVISION)-LL.METRIC(C" + row + ",\"CapitalExpenditure\",\"FY2025\",_LL_REVISION)", scope);
                Formula(sheet, "H" + row, "=LL.METRIC(C" + row + ",\"DilutedEPS\",\"FY2025\",_LL_REVISION)", scope);
            }
            scope.Own(sheet.Range["D10:G12"]).NumberFormat = "#,##0;[Red](#,##0);–";
            scope.Own(sheet.Range["E10:F12"]).NumberFormat = "0.0%"; scope.Own(sheet.Range["H10:H12"]).NumberFormat = "$0.00";
            Style(sheet.Range["B10:H12"], Pale, Navy, false, 12, scope); scope.Own(sheet.Range["10:12"]).RowHeight = 34;
            Merged(sheet, "B15:H15", "THREE YEARS OF REPORTED REVENUE", scope); Style(sheet.Range["B15:H15"], null, Teal, true, 10, scope);
            scope.Own(sheet.Range["B17:E17"]).Value2 = new object[,] { { "Company", "FY2023", "FY2024", "FY2025" } }; Header(sheet, "B17:E17", scope);
            for (var i = 0; i < companies.Length; i++)
            {
                var row = 18 + i; Cell(sheet, "B" + row, companies[i].Ticker, scope);
                foreach (var column in new[] { "C", "D", "E" }) Formula(sheet, column + row, "=LL.METRIC($B" + row + ",\"Revenue\"," + column + "$17,_LL_REVISION)", scope);
            }
            scope.Own(sheet.Range["C18:E20"]).NumberFormat = "#,##0";
            Merged(sheet, "G17:H17", "WORKFLOW", scope); Style(sheet.Range["G17:H17"], null, Teal, true, 10, scope);
            Merged(sheet, "G18:H19", "1  Open Research pane\n2  Inspect a sourced fact", scope);
            Merged(sheet, "G20:H22", "3  Model → Review updates\n4  Adjust the blue assumptions\n5  Test offline recovery", scope);
            scope.Own(sheet.Range["G18:H22"]).WrapText = true; Style(sheet.Range["G18:H22"], null, Muted, false, 10, scope);
            var charts = scope.Own((Xl.ChartObjects)sheet.ChartObjects());
            var chartObject = scope.Own(charts.Add(32, 445, 675, 255)); var chart = scope.Own(chartObject.Chart);
            chart.ChartType = Xl.XlChartType.xlColumnClustered; chart.SetSourceData(scope.Own(sheet.Range["B17:E20"]), Xl.XlRowCol.xlRows);
            chart.HasTitle = true; scope.Own(chart.ChartTitle).Text = "Revenue trajectory  |  USD millions";
            chart.ChartStyle = 13; chart.HasLegend = true;
            Merged(sheet, "B43:H43", "LIVE WORKSPACE ACTIVITY", scope); Style(sheet.Range["B43:H43"], null, Teal, true, 10, scope);
            Merged(sheet, "B44:H45", "", scope); Formula(sheet, "B44", "=LL.LIVE()", scope); scope.Own(sheet.Range["B44:H45"]).WrapText = true; Style(sheet.Range["B44:H45"], Pale, Navy, false, 10, scope);
            Merged(sheet, "B47:H48", "Comparability note: Microsoft ends its fiscal year in June, Apple in September, and NVIDIA in January. Capital expenditure concepts differ by issuer. OCF less capex is derived, not a standardized GAAP line item.", scope); scope.Own(sheet.Range["B47:H48"]).WrapText = true; Style(sheet.Range["B47:H48"], null, Muted, false, 9, scope);
        }
        private static void BuildModel(Xl.Worksheet sheet, ComScope scope)
        {
            scope.Own(sheet.Range["A:A"]).ColumnWidth = 30; scope.Own(sheet.Range["B:B"]).ColumnWidth = 23; scope.Own(sheet.Range["C:G"]).ColumnWidth = 18;
            Merged(sheet, "A1:G2", "LedgerLens  /  Operating model", scope); Title(sheet, "A1:G2", scope);
            Cell(sheet, "A4", "Company", scope); Cell(sheet, "B4", "MSFT", scope); Style(sheet.Range["B4"], Pale, Blue, true, 12, scope);
            Cell(sheet, "E4", "ANALYST ASSUMPTIONS", scope); Style(sheet.Range["E4:G4"], null, Teal, true, 10, scope);
            Cell(sheet, "E5", "Revenue growth", scope); Cell(sheet, "F5", .12, scope); Cell(sheet, "G5", .10, scope);
            Cell(sheet, "E6", "Operating margin", scope); Cell(sheet, "F6", .45, scope); Cell(sheet, "G6", .45, scope);
            scope.Own(sheet.Range["F5:G6"]).NumberFormat = "0.0%"; Style(sheet.Range["F5:G6"], ColorTranslator.ToOle(Color.FromArgb(235, 242, 255)), Blue, true, 11, scope);
            var validation = scope.Own(scope.Own(sheet.Range["B4"]).Validation); validation.Add(Xl.XlDVType.xlValidateList, Xl.XlDVAlertStyle.xlValidAlertStop, Xl.XlFormatConditionOperator.xlBetween, "MSFT,AAPL,NVDA");
            scope.Own(sheet.Range["A8:G8"]).Value2 = new object[,] { { "Reported financials", "Metric ID", "FY2023", "FY2024", "FY2025", "FY2026E", "FY2027E" } }; Header(sheet, "A8:G8", scope);
            Cell(sheet, "C9", "Sourced formula", scope); Cell(sheet, "D9", "Sourced formula", scope); Cell(sheet, "E9", "Reviewed import", scope); Cell(sheet, "F9", "Forecast", scope); Cell(sheet, "G9", "Forecast", scope); Style(sheet.Range["C9:G9"], null, Muted, false, 9, scope);
            Formula(sheet, "E9", "=IF(_LL_IMPORTED_TICKER=\"\",\"Awaiting import\",\"Imported: \"&_LL_IMPORTED_TICKER)", scope);
            var metrics = MetricCatalog.Labels.ToArray();
            for (var i = 0; i < metrics.Length; i++)
            {
                var row = 10 + i; Cell(sheet, "A" + row, metrics[i].Value, scope); Cell(sheet, "B" + row, metrics[i].Key, scope);
                foreach (var column in new[] { "C", "D" }) Formula(sheet, column + row, "=LL.METRIC($B$4,$B" + row + "," + column + "$8,_LL_REVISION)", scope);
                var f = row == 10 ? "E10*(1+$F$5)" : row == 12 ? "F10*$F$6" : "E" + row + "*(F10/$E$10)";
                var g = row == 10 ? "F10*(1+$G$5)" : row == 12 ? "G10*$G$6" : "F" + row + "*(G10/$F$10)";
                Formula(sheet, "F" + row, "=IF(OR($B$4<>_LL_IMPORTED_TICKER,$E$10=\"\",$E$10=0),\"\"," + f + ")", scope);
                Formula(sheet, "G" + row, "=IF(OR($F$10=\"\",$F$10=0),\"\"," + g + ")", scope);
            }
            scope.Own(sheet.Range["C10:G18"]).NumberFormat = "#,##0;[Red](#,##0);–"; scope.Own(sheet.Range["C16:G16"]).NumberFormat = "0.00";
            Style(sheet.Range["E10:E18"], Pale, Navy, false, 11, scope); Style(sheet.Range["F10:G18"], ColorTranslator.ToOle(Color.FromArgb(246, 248, 252)), Navy, false, 11, scope);
            scope.Own(sheet.Range["10:18"]).RowHeight = 27; Style(sheet.Range["B10:B18"], null, Muted, false, 9, scope);
            Merged(sheet, "A21:G22", "", scope); Formula(sheet, "A21", "=IF(AND($B$4=_LL_IMPORTED_TICKER,COUNTBLANK(E10:E18)=0),\"Reported FY2025 values imported for \"&$B$4&\". Blue inputs drive the illustrative forecasts.\",\"Review required: FY2025 inputs are empty or belong to another company. Forecasts are paused. Open LedgerLens > Review updates.\")", scope); scope.Own(sheet.Range["A21:G22"]).WrapText = true; Style(sheet.Range["A21:G22"], Pale, Teal, false, 11, scope);
            Merged(sheet, "A24:G25", "Model mechanics: revenue uses the blue growth assumptions; operating income uses the blue margin assumptions. Other lines scale with revenue for demonstration. These forecasts are illustrative analyst scenarios, not company guidance or investment recommendations.", scope); scope.Own(sheet.Range["A24:G25"]).WrapText = true; Style(sheet.Range["A24:G25"], null, Muted, false, 10, scope);
            Merged(sheet, "A27:G28", "Refresh safeguards: existing formulas and analyst assumptions are preserved. Edits made after preview stop the whole update. Undo checks for later edits before restoring the previous values. Changing the company requires reviewing a new import.", scope); scope.Own(sheet.Range["A27:G28"]).WrapText = true; Style(sheet.Range["A27:G28"], null, Muted, false, 10, scope);
            Freeze(sheet, 9, 2, scope);
        }
        private static void BuildSources(Xl.Worksheet sheet, ComScope scope)
        {
            var facts = HostRuntime.Client.Snapshot.Facts;
            Merged(sheet, "A1:J2", "LedgerLens  /  Evidence register", scope); Title(sheet, "A1:J2", scope);
            Merged(sheet, "A3:J3", "Bundled SEC evidence. Reviewed model imports append a full source audit below this table, including raw values and acquisition dates.", scope); Style(sheet.Range["A3:J3"], null, Muted, false, 10, scope);
            var values = new object[facts.Length + 1, 10];
            var headers = new[] { "Ticker", "Metric", "Fiscal year", "Value", "Units", "Period start", "Period end", "Filed", "SEC filing", "XBRL concept" };
            for (var col = 0; col < headers.Length; col++) values[0, col] = headers[col];
            for (var row = 0; row < facts.Length; row++)
            { var f = facts[row]; var fields = new object[] { f.Ticker, f.Label, f.Period, (double)f.Value, f.Unit, f.Start ?? "Instant", f.End, f.Filed, f.SourceUrl, f.Concept }; for (var col = 0; col < fields.Length; col++) values[row + 1, col] = fields[col]; }
            var range = scope.Own(sheet.Range["A5", "J" + (facts.Length + 5)]); range.Value2 = values;
            var tables = scope.Own(sheet.ListObjects); var table = scope.Own(tables.Add(Xl.XlListObjectSourceType.xlSrcRange, range, Type.Missing, Xl.XlYesNoGuess.xlYes)); table.Name = "EvidenceRegister"; table.TableStyle = "TableStyleMedium2";
            scope.Own(sheet.Range["A:A"]).ColumnWidth = 10; scope.Own(sheet.Range["B:B"]).ColumnWidth = 24; scope.Own(sheet.Range["C:H"]).ColumnWidth = 16; scope.Own(sheet.Range["I:I"]).ColumnWidth = 60; scope.Own(sheet.Range["J:J"]).ColumnWidth = 50;
            Freeze(sheet, 5, 0, scope);
        }
        private static void BuildGuide(Xl.Worksheet sheet, ComScope scope)
        {
            scope.Own(sheet.Range["A:A"]).ColumnWidth = 5; scope.Own(sheet.Range["B:B"]).ColumnWidth = 28; scope.Own(sheet.Range["C:G"]).ColumnWidth = 18;
            Merged(sheet, "B2:G3", "LedgerLens  /  Start here", scope); Title(sheet, "B2:G3", scope);
            var instructions = new[]
            {
                ("01  Explore", "Dashboard compares three companies with asynchronous C# financial formulas. Open the LedgerLens ribbon > Research pane. Select a numeric fact to inspect its SEC source."),
                ("02  Research", "Ask a question about the selected company's annual financials. OpenAI answers cite supplied source IDs. Uncheck Use OpenAI for deterministic calculated observations. No AlphaSense account is used."),
                ("03  Review", "On the Model sheet, select your company. Open Review updates and preview FY2025 reported values before applying. The green historical-input column is intentionally empty in a new model."),
                ("04  Model", "Edit the blue growth and margin assumptions. Forecast formulas remain intact during refresh. A changed cell after preview stops the update; Undo refuses to overwrite subsequent edits."),
                ("05  Recover", "Open Health and test Offline, Slow, or Provider outage. Inspect cache/coalescing counters and the live notification feed. Replay notifications are explicitly simulated; reported facts never change synthetically."),
                ("Formula library", "LL.METRIC(ticker, metric, period, revision): number. LL.SOURCE: SEC filing. LL.STATUS: freshness and units. LL.TABLE: spill nine metrics. LL.ASK: explicit cached AI research. LL.LIVE: service notifications."),
                ("Units & periods", "All currency metrics are USD millions; diluted EPS is USD/share. The dataset covers fiscal FY2023–FY2025. Issuers' calendars and capex concepts differ. All acquisition and reporting dates are in the evidence register."),
                ("Boundaries", "Independent engineering demonstration, not an AlphaSense product. Forecasts are illustrative. Source-ID checks do not establish the truth of every AI interpretation. Review the original filings and calculations."),
                ("Compatibility", "Validated on this Windows Microsoft 365 x64 installation. The native XLL targets .NET Framework 4.8. Dynamic arrays need a supporting Excel version. A separate Office.js bridge is provided for future cross-platform deployment.")
            };
            for (var i = 0; i < instructions.Length; i++)
            {
                var row = 6 + i * 4;
                Merged(sheet, "B" + row + ":B" + (row + 2), instructions[i].Item1, scope); Style(sheet.Range["B" + row + ":B" + (row + 2)], null, Teal, true, 11, scope);
                Merged(sheet, "C" + row + ":G" + (row + 2), instructions[i].Item2, scope); var body = scope.Own(sheet.Range["C" + row + ":G" + (row + 2)]); body.WrapText = true; body.VerticalAlignment = Xl.XlVAlign.xlVAlignTop;
            }
        }
        internal static void Title(Xl.Worksheet sheet, string address, ComScope scope) => Style(sheet.Range[address], Navy, ColorTranslator.ToOle(Color.White), true, 22, scope);
        private static void Header(Xl.Worksheet sheet, string address, ComScope scope) { Style(sheet.Range[address], Teal, ColorTranslator.ToOle(Color.White), true, 10, scope); scope.Own(sheet.Range[address]).RowHeight = 28; }
        private static void Cell(Xl.Worksheet sheet, string address, object value, ComScope scope) => scope.Own(sheet.Range[address]).Value2 = value;
        private static void Formula(Xl.Worksheet sheet, string address, string formula, ComScope scope) => scope.Own(sheet.Range[address]).Formula = formula;
        private static void Merged(Xl.Worksheet sheet, string address, string value, ComScope scope) { var range = scope.Own(sheet.Range[address]); range.Merge(); range.Value2 = value; range.VerticalAlignment = Xl.XlVAlign.xlVAlignCenter; }
        private static void Style(Xl.Range range, int? background, int foreground, bool bold, int size, ComScope scope)
        { scope.Own(range); if (background.HasValue) scope.Own(range.Interior).Color = background.Value; var font = scope.Own(range.Font); font.Color = foreground; font.Bold = bold; font.Size = size; }
        private static void Freeze(Xl.Worksheet sheet, int rows, int columns, ComScope scope)
        {
            // Excel can reject window preferences during COM-driven macro execution.
            // Worksheet construction must not depend on view or printer availability.
            try { sheet.Activate(); var window = scope.Own(WorkbookActions.App.ActiveWindow); window.SplitRow = rows; window.SplitColumn = columns; window.FreezePanes = true; }
            catch (System.Runtime.InteropServices.COMException) { }
        }
    }
}
