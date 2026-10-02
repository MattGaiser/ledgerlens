using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace LedgerLens.Excel
{
    [ComVisible(true)]
    public sealed class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='LedgerLensTab' label='LedgerLens'><group id='WorkspaceGroup' label='Research workspace'><button id='OpenWorkspace' label='Research pane' size='large' imageMso='ResearchPane' onAction='OpenWorkspace'/><button id='NewModel' label='Create analyst model' size='large' imageMso='TableInsert' onAction='NewModel'/></group><group id='ModelGroup' label='Model control'><button id='ReviewChanges' label='Review updates' size='large' imageMso='ReviewAcceptChange' onAction='ReviewChanges'/><button id='RecalculateModel' label='Refresh formulas' imageMso='CalculateNow' onAction='RecalculateModel'/><button id='UndoUpdate' label='Undo model update' imageMso='Undo' onAction='UndoUpdate'/></group><group id='ReliabilityGroup' label='Reliability'><button id='Health' label='Workspace health' size='large' imageMso='ShowAllProperties' onAction='Health'/></group></tab></tabs></ribbon></customUI>";
        public void OpenWorkspace(IRibbonControl control) => Safe(() => HostRuntime.ShowPane());
        public void NewModel(IRibbonControl control) => Safe(() => WorkbookBuilder.Create());
        public void ReviewChanges(IRibbonControl control) => Safe(() => HostRuntime.ShowPane("refresh"));
        public void RecalculateModel(IRibbonControl control) => Safe(() => HostRuntime.Actions.Recalculate());
        public void UndoUpdate(IRibbonControl control) => Safe(() => HostRuntime.Actions.Rollback());
        public void Health(IRibbonControl control) => Safe(() => HostRuntime.ShowPane("health"));
        private static void Safe(Action action) { try { action(); } catch (Exception e) { HostRuntime.RecordError(e.Message); System.Windows.Forms.MessageBox.Show(e.Message, "LedgerLens", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); } }
    }

    public static class Commands
    {
        [ExcelCommand(Name = "LL.OPEN", Description = "Open the LedgerLens research pane.")]
        public static void Open() => HostRuntime.ShowPane();
        [ExcelCommand(Name = "LL.NEW", Description = "Create the LedgerLens analyst workbook.")]
        public static void New() => WorkbookBuilder.Create();
        [ExcelCommand(Name = "LL.HEALTH", Description = "Inspect workspace health and failure scenarios.")]
        public static void Health() => HostRuntime.ShowPane("health");
        [ExcelCommand(Name = "LL.REFRESH", Description = "Refresh formula revisions in the active LedgerLens workbook.")]
        public static void Refresh() => HostRuntime.Actions.Recalculate();
        [ExcelCommand(Name = "LL.DISCONNECT", Description = "Release this add-in's resources before automated unloading.")]
        public static void Disconnect() => HostRuntime.Stop();
    }
}
