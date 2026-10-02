param([switch]$Browser, [switch]$NoExcel, [switch]$OfficePreview, [switch]$Validate, [ValidateRange(1024,65534)][int]$Port = 17843)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$runtime = Join-Path $root '.runtime'
[void](New-Item -ItemType Directory -Path $runtime -Force)
# Session credentials and WebView data are private to the current Windows account.
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().User
$acl = [IO.Directory]::GetAccessControl($runtime,[Security.AccessControl.AccessControlSections]::Access)
$acl.SetAccessRuleProtection($true,$false)
$acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($identity,'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
[IO.Directory]::SetAccessControl($runtime,$acl)
$servicePath = Join-Path $root 'service\LedgerLens.Service.exe'
$nativeDir = Join-Path $root 'native'
if (-not (Test-Path -LiteralPath $servicePath)) {
    $servicePath = Join-Path $root 'artifacts\build\service\LedgerLens.Service.exe'
    $nativeDir = Join-Path $root 'artifacts\build\native'
}
if (-not (Test-Path -LiteralPath $servicePath)) { throw 'Build first with .\build.ps1, or use the extracted Windows release.' }
$endpointPath = Join-Path $runtime 'endpoint.json'
$endpoint = $null
if (Test-Path -LiteralPath $endpointPath) {
    try {
        $candidate = Get-Content -LiteralPath $endpointPath -Raw | ConvertFrom-Json
        $process = Get-Process -Id $candidate.ProcessId -ErrorAction Stop
        if ($process.Path -eq $servicePath -and $candidate.BaseUrl -eq "http://127.0.0.1:$Port") {
            $null = Invoke-RestMethod -Uri ($candidate.BaseUrl+'/api/diagnostics') -Headers @{Authorization='Bearer '+$candidate.Token} -TimeoutSec 3
            $endpoint = $candidate
        }
    } catch { }
}
if (-not $endpoint) {
    $env:LEDGERLENS_ROOT = $root
    $env:LEDGERLENS_PORT = [string]$Port
    $env:LEDGERLENS_HTTPS = if ($OfficePreview) { '1' } else { '0' }
    $service = Start-Process -FilePath $servicePath -WorkingDirectory $root -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runtime 'service.log') -RedirectStandardError (Join-Path $runtime 'service-error.log')
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ($service.HasExited) { throw 'Service exited during startup. Check .runtime\service-error.log and confirm the port is available.' }
        try {
            $candidate = Get-Content -LiteralPath $endpointPath -Raw | ConvertFrom-Json
            if ($candidate.ProcessId -eq $service.Id) {
                $null = Invoke-RestMethod -Uri ($candidate.BaseUrl+'/health') -TimeoutSec 2
                $endpoint = $candidate
                break
            }
        } catch { }
        Start-Sleep -Milliseconds 150
    }
    if (-not $endpoint) { throw 'Service did not become ready within 20 seconds.' }
}
Write-Output 'LedgerLens research service is ready.'
if ($OfficePreview) {
    if (-not $endpoint.OfficeUrl) { throw 'Stop the existing service and relaunch with -OfficePreview to enable HTTPS.' }
    [xml]$manifest = Get-Content -LiteralPath (Join-Path $root 'office\manifest.xml') -Raw
    foreach ($node in $manifest.SelectNodes('//*[@DefaultValue]')) {
        if ($node.DefaultValue.StartsWith('https://127.0.0.1:17844')) { $node.DefaultValue = $node.DefaultValue.Replace('https://127.0.0.1:17844',$endpoint.OfficeUrl) }
    }
    $source = $manifest.SelectSingleNode('//*[local-name()="SourceLocation"]')
    $source.DefaultValue = $endpoint.OfficeUrl+'/?office=1#session='+$endpoint.Token
    $manifest.Save((Join-Path $runtime 'office-session-manifest.xml'))
    Write-Output 'Office preview manifest created at .runtime\office-session-manifest.xml. See docs\OFFICE-PREVIEW.md.'
}
if ($Browser) { Start-Process ($endpoint.BaseUrl+'/#session='+$endpoint.Token) }
if ($NoExcel) { return }

function Get-FormulaErrorCount($Workbook) {
    $worksheets=$null; $errorCount=0
    try {
        $worksheets=$Workbook.Worksheets
        for($index=1; $index -le $worksheets.Count; $index++) {
            $worksheet=$null; $used=$null; $errors=$null
            try {
                $worksheet=$worksheets.Item($index)
                $used=$worksheet.UsedRange
                try { $errors=$used.SpecialCells(-4123,16) }
                catch {
                    # SpecialCells raises Excel error 1004 when no matching cells exist.
                    $failure=$_.Exception
                    while($failure.InnerException){$failure=$failure.InnerException}
                    if($failure.HResult -ne -2146827284){throw}
                }
                if($errors){$errorCount+=$errors.Count}
            } finally {
                foreach($com in @($errors,$used,$worksheet)) {
                    if($com -and [Runtime.InteropServices.Marshal]::IsComObject($com)){
                        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)
                    }
                }
            }
        }
        return $errorCount
    } finally {
        if($worksheets){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheets)}
    }
}

$app=$null; $books=$null; $blank=$null; $book=$null; $addin=$null; $addins=$null; $sheet=$null; $cell=$null; $success=$false
$env:LEDGERLENS_ROOT = $root
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible=$true; $app.DisplayAlerts=$false
    $books=$app.Workbooks; $blank=$books.Add()
    $reader = New-Object IO.BinaryReader([IO.File]::OpenRead((Join-Path $app.Path 'EXCEL.EXE')))
    try { $reader.BaseStream.Position=0x3c; $offset=$reader.ReadInt32(); $reader.BaseStream.Position=$offset+4; $machine=$reader.ReadUInt16() } finally { $reader.Dispose() }
    $xllName = if ($machine -eq 0x8664) { 'LedgerLens.Excel-AddIn64.xll' } elseif ($machine -eq 0x14c) { 'LedgerLens.Excel-AddIn.xll' } else { throw 'This native package supports x64 or x86 Excel on Windows.' }
    $xll=Join-Path $nativeDir $xllName
    if (-not (Test-Path -LiteralPath $xll)) { throw 'The native add-in is missing. Keep the entire release folder together.' }
    $addins=$app.AddIns; $addin=$addins.Add($xll,$false); $addin.Installed=$true
    $productVersion=[string]$app.Run('LL.VERSION')
    $assemblyVersion=[Reflection.AssemblyName]::GetAssemblyName((Join-Path $nativeDir 'LedgerLens.Core.dll')).Version
    $expectedVersion='LedgerLens '+$assemblyVersion.ToString(3)
    if($productVersion -ne $expectedVersion){throw 'Excel loaded a different LedgerLens version. Save and close LedgerLens workbooks before upgrading.'}
    [void]$app.Run('LL.NEW')
    $book=$books.Item($books.Count)
    if ($book.Worksheets.Count -ne 5) { throw 'The analyst workbook could not be created. See .runtime\excel.log.' }
    $sheet=$book.Worksheets.Item('Dashboard'); $cell=$sheet.Range('D10')
    $deadline=[DateTime]::UtcNow.AddSeconds(25)
    do {
        $app.Calculate()
        $revenue=$cell.Value2
        $formulaErrors=Get-FormulaErrorCount $book
        $calculationDone=$app.CalculationState -eq 0
        if($revenue -eq 281724 -and $formulaErrors -eq 0 -and $calculationDone){break}
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($revenue -ne 281724 -or $formulaErrors -ne 0 -or -not $calculationDone) {
        throw 'The initial workbook formulas did not all resolve within 25 seconds. See .runtime\excel.log.'
    }
    $outputDir=Join-Path $root 'workbooks'; [void](New-Item -ItemType Directory -Path $outputDir -Force)
    $outputPath=Join-Path $outputDir ('LedgerLens-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,6)+'.xlsx')
    $book.SaveCopyAs($outputPath)
    $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book=$books.Open($outputPath); $blank.Close($false); $book.Activate()
    $app.DisplayAlerts=$true; $app.UserControl=$true
    [void]$app.Run('LL.OPEN')
    if ($Validate) {
        Start-Sleep -Milliseconds 2000
        [ordered]@{checkedAt=[DateTime]::UtcNow.ToString('o');status='PASS';productVersion=$productVersion;excelVersion=$app.Version;excelBuild=$app.Build;architecture=$machine;worksheets=$book.Worksheets.Count;revenue=$revenue;resolvedFormulaErrors=$formulaErrors;workbook=$outputPath;addin=$xll} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $runtime 'launch-validation.json') -Encoding UTF8
        Write-Output 'PASS: native add-in loaded, five worksheets created, all initial formulas resolved, saved workbook reopened, research pane opened.'
    } else {
        Write-Output ('Opened '+$outputPath)
    }
    $success=$true
} finally {
    if ($Validate -or -not $success) {
        if($app){try{[void]$app.Run('LL.DISCONNECT')}catch{}}
        if($book){try{$book.Close($false)}catch{}}
        if($blank){try{$blank.Close($false)}catch{}}
        if($addin){try{$addin.Installed=$false}catch{}}
        if($app){try{$app.Quit()}catch{}}
    }
    foreach($com in @($cell,$sheet,$book,$blank,$books,$addin,$addins,$app)) {
        if($com -and [Runtime.InteropServices.Marshal]::IsComObject($com)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)}
    }
    $cell=$null; $sheet=$null; $book=$null; $blank=$null; $books=$null; $addin=$null; $addins=$null; $app=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
