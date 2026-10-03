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
            var books = app.Workbooks;
            var book = books.Add(Xl.XlWBATemplate.xlWBATWorksheet);
            try
            {
                using (new ExcelStateGuard(app))
                {
                    HostRuntime.RecordActivity("Workbook: initializing sheets");
                    var sheets = book.Worksheets;
                    var dashboard = (Xl.Worksheet)sheets[1];
                    dashboard.Name = "Dashboard";
                    var model = AddSheet(sheets, "Model");
                    var sources = AddSheet(sheets, "Sources");
                    var guide = AddSheet(sheets, "Guide");
                    var control = AddSheet(sheets, "_LedgerLens");
                    foreach (var sheet in new[] { dashboard, model, sources, guide, control })
                    {
                        var cells = sheet.Cells;
                        var font = cells.Font;
                        font.Name = "Aptos";
                        font.Size = 11;
                        font.Color = Navy;
                        sheet.Columns.ColumnWidth = 14;
                        sheet.Activate();
                        var window = app.ActiveWindow;
                        window.DisplayGridlines = false;
                        window.Zoom = 90;
                    }
                    Cell(control, "A1", "LedgerLens/1");
                    Cell(control, "B1", Guid.NewGuid().ToString("N"));
                    Cell(control, "A2", "Formula revision");
                    Cell(control, "B2", 0);
                    Cell(control, "A3", "Imported company");
                    var names = book.Names;
                    names.Add("_LL_REVISION", "='_LedgerLens'!$B$2");
                    names.Add("_LL_IMPORTED_TICKER", "='_LedgerLens'!$B$3");
                    control.Visible = Xl.XlSheetVisibility.xlSheetVeryHidden;
                    HostRuntime.RecordActivity("Workbook: dashboard");
                    BuildDashboard(dashboard);
                    HostRuntime.RecordActivity("Workbook: operating model");
                    BuildModel(model);
                    HostRuntime.RecordActivity("Workbook: evidence register");
                    BuildSources(sources);
                    HostRuntime.RecordActivity("Workbook: guide");
                    BuildGuide(guide);
                    dashboard.Activate();
                    dashboard.Range["A1"].Select();
                    // Workbook title and sheets carry the independent demo identity; no proprietary content is used.
                    book.CheckCompatibility = false;
                }
                HostRuntime.RecordActivity("Workbook: ready");
            }
            catch (Exception e) { HostRuntime.RecordError("Workbook creation: " + e); book.Close(false); throw; }
        }
        private static Xl.Worksheet AddSheet(Xl.Sheets sheets, string name)
        {
            var last = sheets[sheets.Count];
            var sheet = (Xl.Worksheet)sheets.Add(After: last);
            sheet.Name = name;
            return sheet;
        }
        private static void BuildDashboard(Xl.Worksheet sheet)
        {
            sheet.Range["A:A"].ColumnWidth = 3;
            sheet.Range["B:B"].ColumnWidth = 23;
            sheet.Range["C:H"].ColumnWidth = 16;
            Merged(sheet, "B2:H3", "LedgerLens  /  Analyst workspace");
            Title(sheet, "B2:H3");
            Merged(sheet, "B5:H5", "REPORTED FUNDAMENTALS, WITH THE EVIDENCE IN REACH");
            Style(sheet.Range["B5:H5"], null, Teal, true, 10);
            Merged(sheet, "B6:H6", "Public SEC filings · USD millions unless noted · FY2025 is historical, not a live market quote");
            Style(sheet.Range["B6:H6"], null, Muted, false, 10);
            var headers = new object[,] { { "Company", "Ticker", "Revenue", "Revenue YoY", "Operating margin", "OCF less capex", "Diluted EPS" } };
            sheet.Range["B9:H9"].Value2 = headers;
            Header(sheet, "B9:H9");
            var companies = HostRuntime.Client.Snapshot.Companies;
            for (var i = 0; i < companies.Length; i++)
            {
                var row = 10 + i;
                Cell(sheet, "B" + row, companies[i].Name);
                Cell(sheet, "C" + row, companies[i].Ticker);
                Formula(sheet, "D" + row, "=LL.METRIC(C" + row + ",\"Revenue\",\"FY2025\",_LL_REVISION)");
                Formula(sheet, "E" + row, "=IFERROR(D" + row + "/LL.METRIC(C" + row + ",\"Revenue\",\"FY2024\",_LL_REVISION)-1,NA())");
                Formula(sheet, "F" + row, "=IFERROR(LL.METRIC(C" + row + ",\"OperatingIncome\",\"FY2025\",_LL_REVISION)/D" + row + ",NA())");
                Formula(sheet, "G" + row, "=LL.METRIC(C" + row + ",\"OperatingCashFlow\",\"FY2025\",_LL_REVISION)-LL.METRIC(C" + row + ",\"CapitalExpenditure\",\"FY2025\",_LL_REVISION)");
                Formula(sheet, "H" + row, "=LL.METRIC(C" + row + ",\"DilutedEPS\",\"FY2025\",_LL_REVISION)");
            }
            sheet.Range["D10:G12"].NumberFormat = "#,##0;[Red](#,##0);–";
            sheet.Range["E10:F12"].NumberFormat = "0.0%";
            sheet.Range["H10:H12"].NumberFormat = "$0.00";
            Style(sheet.Range["B10:H12"], Pale, Navy, false, 12);
            sheet.Range["10:12"].RowHeight = 34;
            Merged(sheet, "B15:H15", "THREE YEARS OF REPORTED REVENUE");
            Style(sheet.Range["B15:H15"], null, Teal, true, 10);
            sheet.Range["B17:E17"].Value2 = new object[,] { { "Company", "FY2023", "FY2024", "FY2025" } };
            Header(sheet, "B17:E17");
            for (var i = 0; i < companies.Length; i++)
            {
                var row = 18 + i;
                Cell(sheet, "B" + row, companies[i].Ticker);
                foreach (var column in new[] { "C", "D", "E" })
                    Formula(sheet, column + row, "=LL.METRIC($B" + row + ",\"Revenue\"," + column + "$17,_LL_REVISION)");
            }
            sheet.Range["C18:E20"].NumberFormat = "#,##0";
            Merged(sheet, "G17:H17", "WORKFLOW");
            Style(sheet.Range["G17:H17"], null, Teal, true, 10);
            Merged(sheet, "G18:H19", "1  Open Research pane\n2  Inspect a sourced fact");
            Merged(sheet, "G20:H22", "3  Model → Review updates\n4  Adjust the blue assumptions\n5  Test offline recovery");
            sheet.Range["G18:H22"].WrapText = true;
            Style(sheet.Range["G18:H22"], null, Muted, false, 10);
            var charts = (Xl.ChartObjects)sheet.ChartObjects();
            var chartObject = charts.Add(32, 445, 675, 255);
            var chart = chartObject.Chart;
            chart.ChartType = Xl.XlChartType.xlColumnClustered;
            chart.SetSourceData(sheet.Range["B17:E20"], Xl.XlRowCol.xlRows);
            chart.HasTitle = true;
            chart.ChartTitle.Text = "Revenue trajectory  |  USD millions";
            chart.ChartStyle = 13;
            chart.HasLegend = true;
            Merged(sheet, "B43:H43", "LIVE WORKSPACE ACTIVITY");
            Style(sheet.Range["B43:H43"], null, Teal, true, 10);
            Merged(sheet, "B44:H45", "");
            Formula(sheet, "B44", "=LL.LIVE()");
            sheet.Range["B44:H45"].WrapText = true;
            Style(sheet.Range["B44:H45"], Pale, Navy, false, 10);
            Merged(sheet, "B47:H48", "Comparability note: Microsoft ends its fiscal year in June, Apple in September, and NVIDIA in January. Capital expenditure concepts differ by issuer. OCF less capex is derived, not a standardized GAAP line item.");
            sheet.Range["B47:H48"].WrapText = true;
            Style(sheet.Range["B47:H48"], null, Muted, false, 9);
        }
        private static void BuildModel(Xl.Worksheet sheet)
        {
            sheet.Range["A:A"].ColumnWidth = 30;
            sheet.Range["B:B"].ColumnWidth = 23;
            sheet.Range["C:G"].ColumnWidth = 18;
            Merged(sheet, "A1:G2", "LedgerLens  /  Operating model");
            Title(sheet, "A1:G2");
            Cell(sheet, "A4", "Company");
            Cell(sheet, "B4", "MSFT");
            Style(sheet.Range["B4"], Pale, Blue, true, 12);
            Cell(sheet, "E4", "ANALYST ASSUMPTIONS");
            Style(sheet.Range["E4:G4"], null, Teal, true, 10);
            Cell(sheet, "E5", "Revenue growth");
            Cell(sheet, "F5", .12);
            Cell(sheet, "G5", .10);
            Cell(sheet, "E6", "Operating margin");
            Cell(sheet, "F6", .45);
            Cell(sheet, "G6", .45);
            sheet.Range["F5:G6"].NumberFormat = "0.0%";
            Style(sheet.Range["F5:G6"], ColorTranslator.ToOle(Color.FromArgb(235, 242, 255)), Blue, true, 11);
            var validation = sheet.Range["B4"].Validation;
            validation.Add(Xl.XlDVType.xlValidateList, Xl.XlDVAlertStyle.xlValidAlertStop, Xl.XlFormatConditionOperator.xlBetween, "MSFT,AAPL,NVDA");
            sheet.Range["A8:G8"].Value2 = new object[,] { { "Reported financials", "Metric ID", "FY2023", "FY2024", "FY2025", "FY2026E", "FY2027E" } };
            Header(sheet, "A8:G8");
            Cell(sheet, "C9", "Sourced formula");
            Cell(sheet, "D9", "Sourced formula");
            Cell(sheet, "E9", "Reviewed import");
            Cell(sheet, "F9", "Forecast");
            Cell(sheet, "G9", "Forecast");
            Style(sheet.Range["C9:G9"], null, Muted, false, 9);
            Formula(sheet, "E9", "=IF(_LL_IMPORTED_TICKER=\"\",\"Awaiting import\",\"Imported: \"&_LL_IMPORTED_TICKER)");
            var metrics = MetricCatalog.Labels.ToArray();
            for (var i = 0; i < metrics.Length; i++)
            {
                var row = 10 + i;
                Cell(sheet, "A" + row, metrics[i].Value);
                Cell(sheet, "B" + row, metrics[i].Key);
                foreach (var column in new[] { "C", "D" })
                    Formula(sheet, column + row, "=LL.METRIC($B$4,$B" + row + "," + column + "$8,_LL_REVISION)");
                var f = row == 10 ? "E10*(1+$F$5)" : row == 12 ? "F10*$F$6" : "E" + row + "*(F10/$E$10)";
                var g = row == 10 ? "F10*(1+$G$5)" : row == 12 ? "G10*$G$6" : "F" + row + "*(G10/$F$10)";
                Formula(sheet, "F" + row, "=IF(OR($B$4<>_LL_IMPORTED_TICKER,$E$10=\"\",$E$10=0),\"\"," + f + ")");
                Formula(sheet, "G" + row, "=IF(OR($F$10=\"\",$F$10=0),\"\"," + g + ")");
            }
            sheet.Range["C10:G18"].NumberFormat = "#,##0;[Red](#,##0);–";
            sheet.Range["C16:G16"].NumberFormat = "0.00";
            Style(sheet.Range["E10:E18"], Pale, Navy, false, 11);
            Style(sheet.Range["F10:G18"], ColorTranslator.ToOle(Color.FromArgb(246, 248, 252)), Navy, false, 11);
            sheet.Range["10:18"].RowHeight = 27;
            Style(sheet.Range["B10:B18"], null, Muted, false, 9);
            Merged(sheet, "A21:G22", "");
            Formula(sheet, "A21", "=IF(AND($B$4=_LL_IMPORTED_TICKER,COUNTBLANK(E10:E18)=0),\"Reported FY2025 values imported for \"&$B$4&\". Blue inputs drive the illustrative forecasts.\",\"Review required: FY2025 inputs are empty or belong to another company. Forecasts are paused. Open LedgerLens > Review updates.\")");
            sheet.Range["A21:G22"].WrapText = true;
            Style(sheet.Range["A21:G22"], Pale, Teal, false, 11);
            Merged(sheet, "A24:G25", "Model mechanics: revenue uses the blue growth assumptions; operating income uses the blue margin assumptions. Other lines scale with revenue for demonstration. These forecasts are illustrative analyst scenarios, not company guidance or investment recommendations.");
            sheet.Range["A24:G25"].WrapText = true;
            Style(sheet.Range["A24:G25"], null, Muted, false, 10);
            Merged(sheet, "A27:G28", "Refresh safeguards: existing formulas and analyst assumptions are preserved. Edits made after preview stop the whole update. Undo checks for later edits before restoring the previous values. Changing the company requires reviewing a new import.");
            sheet.Range["A27:G28"].WrapText = true;
            Style(sheet.Range["A27:G28"], null, Muted, false, 10);
            Freeze(sheet, 9, 2);
        }
        private static void BuildSources(Xl.Worksheet sheet)
        {
            var facts = HostRuntime.Client.Snapshot.Facts;
            Merged(sheet, "A1:J2", "LedgerLens  /  Evidence register");
            Title(sheet, "A1:J2");
            Merged(sheet, "A3:J3", "Bundled SEC evidence. Reviewed model imports append a full source audit below this table, including raw values and acquisition dates.");
            Style(sheet.Range["A3:J3"], null, Muted, false, 10);
            var values = new object[facts.Length + 1, 10];
            var headers = new[] { "Ticker", "Metric", "Fiscal year", "Value", "Units", "Period start", "Period end", "Filed", "SEC filing", "XBRL concept" };
            for (var col = 0; col < headers.Length; col++)
                values[0, col] = headers[col];
            for (var row = 0; row < facts.Length; row++)
            {
                var f = facts[row];
                var fields = new object[] { f.Ticker, f.Label, f.Period, (double)f.Value, f.Unit, f.Start ?? "Instant", f.End, f.Filed, f.SourceUrl, f.Concept };
                for (var col = 0; col < fields.Length; col++)
                    values[row + 1, col] = fields[col];
            }
            var range = sheet.Range["A5", "J" + (facts.Length + 5)];
            range.Value2 = values;
            var tables = sheet.ListObjects;
            var table = tables.Add(Xl.XlListObjectSourceType.xlSrcRange, range, Type.Missing, Xl.XlYesNoGuess.xlYes);
            table.Name = "EvidenceRegister";
            table.TableStyle = "TableStyleMedium2";
            sheet.Range["A:A"].ColumnWidth = 10;
            sheet.Range["B:B"].ColumnWidth = 24;
            sheet.Range["C:H"].ColumnWidth = 16;
            sheet.Range["I:I"].ColumnWidth = 60;
            sheet.Range["J:J"].ColumnWidth = 50;
            Freeze(sheet, 5, 0);
        }
        private static void BuildGuide(Xl.Worksheet sheet)
        {
            sheet.Range["A:A"].ColumnWidth = 5;
            sheet.Range["B:B"].ColumnWidth = 28;
            sheet.Range["C:G"].ColumnWidth = 18;
            Merged(sheet, "B2:G3", "LedgerLens  /  Start here");
            Title(sheet, "B2:G3");
            var instructions = new[]
            {
                ("01  Explore", "Dashboard compares three companies with asynchronous C# financial formulas. Open the LedgerLens ribbon > Research pane. Select a numeric fact to inspect its SEC source."),
                ("02  Research", "Ask a question about the selected company's annual financials. OpenAI answers cite supplied source IDs. Uncheck Use OpenAI for deterministic calculated observations. No AlphaSense account is used."),
                ("03  Review", "On the Model sheet, select your company. Open Review updates and preview FY2025 reported values before applying. The green historical-input column is intentionally empty in a new model."),
                ("04  Model", "Edit the blue growth and margin assumptions. Forecast formulas remain intact during refresh. A changed cell after preview stops the update; Undo refuses to overwrite subsequent edits."),
                ("05  Connection", "Open Health to switch offline, inspect SEC request and retry counts, or test notification delivery. Saved facts remain available offline. Reconnect and use Sync SEC to check for revised filings."),
                ("Formula library", "LL.METRIC(ticker, metric, period, revision): number. LL.SOURCE: SEC filing. LL.STATUS: freshness and units. LL.TABLE: spill nine metrics. LL.ASK: explicit cached AI research. LL.LIVE: service notifications."),
                ("Units & periods", "All currency metrics are USD millions; diluted EPS is USD/share. The dataset covers fiscal FY2023–FY2025. Issuers' calendars and capex concepts differ. All acquisition and reporting dates are in the evidence register."),
                ("Boundaries", "Independent engineering demonstration, not an AlphaSense product. Forecasts are illustrative. Source-ID checks do not establish the truth of every AI interpretation. Review the original filings and calculations."),
                ("Compatibility", "Validated on this Windows Microsoft 365 x64 installation. The native XLL targets .NET Framework 4.8. Dynamic arrays need a supporting Excel version. A separate Office.js bridge is provided for future cross-platform deployment.")
            };
            for (var i = 0; i < instructions.Length; i++)
            {
                var row = 6 + i * 4;
                Merged(sheet, "B" + row + ":B" + (row + 2), instructions[i].Item1);
                Style(sheet.Range["B" + row + ":B" + (row + 2)], null, Teal, true, 11);
                Merged(sheet, "C" + row + ":G" + (row + 2), instructions[i].Item2);
                var body = sheet.Range["C" + row + ":G" + (row + 2)];
                body.WrapText = true;
                body.VerticalAlignment = Xl.XlVAlign.xlVAlignTop;
            }
        }
        internal static void Title(Xl.Worksheet sheet, string address) => Style(sheet.Range[address], Navy, ColorTranslator.ToOle(Color.White), true, 22);
        private static void Header(Xl.Worksheet sheet, string address)
        {
            Style(sheet.Range[address], Teal, ColorTranslator.ToOle(Color.White), true, 10);
            sheet.Range[address].RowHeight = 28;
        }
        private static void Cell(Xl.Worksheet sheet, string address, object value) => sheet.Range[address].Value2 = value;
        private static void Formula(Xl.Worksheet sheet, string address, string formula) => sheet.Range[address].Formula = formula;
        private static void Merged(Xl.Worksheet sheet, string address, string value)
        {
            var range = sheet.Range[address];
            range.Merge();
            range.Value2 = value;
            range.VerticalAlignment = Xl.XlVAlign.xlVAlignCenter;
        }
        private static void Style(Xl.Range range, int? background, int foreground, bool bold, int size)
        {
            if (background.HasValue)
                range.Interior.Color = background.Value;
            var font = range.Font;
            font.Color = foreground;
            font.Bold = bold;
            font.Size = size;
        }
        private static void Freeze(Xl.Worksheet sheet, int rows, int columns)
        {
            // Excel can reject window preferences during COM-driven macro execution.
            // Worksheet construction must not depend on view or printer availability.
            try
            {
                sheet.Activate();
                var window = WorkbookActions.App.ActiveWindow;
                window.SplitRow = rows;
                window.SplitColumn = columns;
                window.FreezePanes = true;
            }
            catch (System.Runtime.InteropServices.COMException) { }
        }
    }
}
