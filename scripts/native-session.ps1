param([int]$DebugPort = 9223, [int]$LifetimeSeconds = 900, [string]$AddInDirectory = '')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$sessionDir = Join-Path $root '.runtime\native-test'
[void](New-Item -ItemType Directory -Path $sessionDir -Force)
foreach ($name in @('command.json','response.json','session.json')) { $oldFile=Join-Path $sessionDir $name; if(Test-Path -LiteralPath $oldFile) { Remove-Item -LiteralPath $oldFile } }
$env:LEDGERLENS_ROOT = $root
# Enabled only in this integration harness, never by the product launcher.
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = '--remote-debugging-port=' + $DebugPort
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class NativeSessionWindow {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
$app=$null; $book=$null; $blank=$null; $books=$null; $addin=$null; $addinList=$null; $pendingBook=$null; $clone=$null; $excelProcessId=0
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $true; $app.DisplayAlerts = $false
    $books=$app.Workbooks; $blank=$books.Add()
    [void][NativeSessionWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$excelProcessId)
    if (-not $AddInDirectory) { $AddInDirectory = Join-Path $root 'src\LedgerLens.Excel\bin\Release\net48' }
    $addinList=$app.AddIns
    # Excel may auto-load an earlier LedgerLens build. Unload it only in this
    # newly created test instance before registering the candidate assembly.
    for($index=1;$index -le $addinList.Count;$index++){
        $existing=$addinList.Item($index)
        try {
            if($existing.Installed -and $existing.Name -like 'LedgerLens*.xll' -and $existing.FullName.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){$existing.Installed=$false}
        } finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($existing)}
    }
    $addin=$addinList.Add((Join-Path $AddInDirectory 'LedgerLens.Excel-AddIn64.xll'),$false)
    $addin.Installed=$true
    if(-not $addin.Installed){throw 'XLL installation failed.'}
    [void]$app.Run('LL.NEW')
    $book=$books.Item($books.Count)
    if ($book.Worksheets.Count -ne 5) { throw 'Model creation failed.' }
    $book.Activate()
    $app.WindowState=-4137
    [void]$app.Run('LL.OPEN')
    @{pid=$excelProcessId;book=$book.Name;hwnd=$app.Hwnd;debugPort=$DebugPort;ready=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $sessionDir 'session.json') -Encoding UTF8
    Write-Output ('Native test session ready: Excel PID ' + $excelProcessId)
    $lastId=''; $timer=[Diagnostics.Stopwatch]::StartNew()
    while ($timer.Elapsed.TotalSeconds -lt $LifetimeSeconds) {
        $commandPath=Join-Path $sessionDir 'command.json'
        if (Test-Path -LiteralPath $commandPath) {
            $request=Get-Content -LiteralPath $commandPath -Raw | ConvertFrom-Json
            if ($request.id -ne $lastId) {
                $lastId=$request.id; $response=@{id=$lastId}
                try {
                    switch ($request.action) {
                        'get' { $sheet=$book.Worksheets.Item($request.sheet); $cell=$sheet.Range($request.address); $response.result=@{value=$cell.Value2;formula=$cell.Formula;hasFormula=$cell.HasFormula}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
                        'set' { $sheet=$book.Worksheets.Item($request.sheet); $cell=$sheet.Range($request.address); if($request.PSObject.Properties.Name -contains 'formula'){$cell.Formula=$request.formula}elseif($null -eq $request.value){$cell.ClearContents()}elseif($request.value -is [ValueType]){$cell.Value2=[double]$request.value}else{$cell.Value2=[string]$request.value}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $app.Calculate(); $response.result=$true }
                        'select' { $book.Activate(); $sheet=$book.Worksheets.Item($request.sheet); $sheet.Activate(); $cell=$sheet.Range($request.address); $cell.Select(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet); $response.result=$true }
                        'activateBlank' { $blank.Activate(); $response.result=$true }
                        'activateModel' { $book.Activate(); $response.result=$true }
                        'protect' { $sheet=$book.Worksheets.Item($request.sheet);if($request.enabled){$sheet.Protect()}else{$sheet.Unprotect()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet);$response.result=$true }
                        'openClone' { $clonePath=Join-Path $sessionDir ('copy-'+[Guid]::NewGuid().ToString('N')+'.xlsx');$book.SaveCopyAs($clonePath);$clone=$books.Open($clonePath,0,[bool]$request.readOnly);$clone.Activate();[void]$app.Run('LL.OPEN');$response.result=$clone.Name }
                        'readClone' { $sheet=$clone.Worksheets.Item('Model');$cell=$sheet.Range('E10');$response.result=$cell.Value2;[void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
                        'closeClone' { $clone.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($clone);$clone=$null;$book.Activate();$response.result=$true }
                        'calculate' { $app.CalculateFull(); $response.result=$true }
                        'spill' { $sheet=$book.Worksheets.Item('Model');$cell=$sheet.Range('I14');$cell.Formula2='=LL.TABLE("MSFT","FY2025",5)';$app.Calculate();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet);$response.result=$true }
                        'bulkFormula' { $sheet=$book.Worksheets.Add(); $sheet.Name='Performance'; $cell=$sheet.Range('A1:A2000'); $watch=[Diagnostics.Stopwatch]::StartNew(); $cell.Formula=$request.formula; $app.Calculate(); $response.result=@{writeAndCalculateMs=$watch.ElapsedMilliseconds;cells=2000}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
                        'cycle' { $previous=$book; [void]$app.Run('LL.NEW'); $book=$books.Item($books.Count); $book.Activate(); [void]$app.Run('LL.OPEN'); $previous.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($previous); $previous=$null; $response.result=$true }
                        'memory' { $currentProcess=Get-Process -Id $excelProcessId; $response.result=@{privateMb=$currentProcess.PrivateMemorySize64/1MB;workingSetMb=$currentProcess.WorkingSet64/1MB;handles=$currentProcess.HandleCount} }
                        'openPending' { $pendingBook=$books.Add(); $sheet=$pendingBook.Worksheets.Item(1);$cell=$sheet.Range('A1');$cell.Formula='=LL.METRIC("AAPL","Revenue","FY2025",271828)';$app.Calculate();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet);$response.result=$true }
                        'closePending' { $pendingBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pendingBook);$pendingBook=$null;$book.Activate();$response.result=$true }
                        'cancellations' { $sheet=$book.Worksheets.Item('Model');$cell=$sheet.Range('J1');$cell.Formula='=LL.CANCELLATIONS()';$cell.Dirty();$cell.Calculate();$response.result=$cell.Value2;[void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
                        'status' { $response.result=@{ready=$app.Ready;updating=$app.ScreenUpdating;events=$app.EnableEvents;calculation=$app.Calculation;sheetCount=$book.Worksheets.Count} }
                        'save' { $outputPath=Join-Path $root ('artifacts\native-validated-' + [Guid]::NewGuid().ToString('N') + '.xlsx'); $book.SaveCopyAs($outputPath); $response.result=$outputPath }
                        'stop' { $response.result=$true; $timer.Stop() }
                        default { throw 'Unsupported test action.' }
                    }
                } catch { $response.error=$_.Exception.Message }
                $response | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $sessionDir 'response.json') -Encoding UTF8
                if (-not $timer.IsRunning) { break }
            }
        }
        Start-Sleep -Milliseconds 100
    }
} finally {
    if ($clone) {try{$clone.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($clone)}catch{}}
    if ($pendingBook) {try{$pendingBook.Close($false)}catch{}}
    if ($app) { try { [void]$app.Run('LL.DISCONNECT') } catch {} }
    if ($book) { try {$book.Close($false)}catch{} }
    if ($blank) { try {$blank.Close($false)}catch{} }
    if ($addin) { try {$addin.Installed=$false}catch{} }
    if ($app) { try {$app.Quit()}catch{} }
    foreach ($com in @($book,$blank,$books,$addin,$addinList,$app)) { if ($com -and [Runtime.InteropServices.Marshal]::IsComObject($com)) {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)} }
    $book=$null; $blank=$null; $books=$null; $addin=$null; $addinList=$null; $app=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Output ('Native test session closed: Excel PID ' + $excelProcessId)
}
