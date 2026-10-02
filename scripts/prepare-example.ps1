$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
$report=Get-Content -LiteralPath (Join-Path $root 'artifacts\validation\native-results.json') -Raw | ConvertFrom-Json
if($report.status -ne 'PASS'){throw 'Native tests must pass before preparing the example.'}
$bookPath=Join-Path $root 'artifacts\LedgerLens-Analyst-Model-1.0.1.xlsx'
Copy-Item -LiteralPath $report.measurements.workbook -Destination $bookPath -Force
$zip=[IO.Compression.ZipFile]::Open($bookPath,[IO.Compression.ZipArchiveMode]::Update)
$errors=@();$formulas=0
try {
    foreach($entry in $zip.Entries){
        if($entry.FullName -match '^xl/worksheets/sheet\d+\.xml$'){
            $reader=New-Object IO.StreamReader($entry.Open());try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
            $formulas+=$xml.SelectNodes('//*[local-name()="f"]').Count
            foreach($cell in $xml.SelectNodes('//*[local-name()="c" and @t="e"]')){$errors+=($entry.FullName+':'+$cell.r+'='+$cell.v)}
        }
    }
    $entry=$zip.GetEntry('xl/workbook.xml');$reader=New-Object IO.StreamReader($entry.Open());try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
    $sheets=@($xml.SelectNodes('//*[local-name()="sheet"]') | ForEach-Object name)
    foreach($view in $xml.SelectNodes('//*[local-name()="workbookView"]')){$view.SetAttribute('activeTab','0');$view.SetAttribute('firstSheet','0')}
    $entry.Delete();$entry=$zip.CreateEntry('xl/workbook.xml');$writer=New-Object IO.StreamWriter($entry.Open(),(New-Object Text.UTF8Encoding($false)));try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
    $external=@($zip.Entries | Where-Object FullName -like 'xl/externalLinks/*' | ForEach-Object FullName)
    [ordered]@{checkedAt=[DateTime]::UtcNow.ToString('o');status=$(if($errors.Count -eq 0){'PASS'}else{'FAIL'});worksheets=$sheets;formulas=$formulas;cachedErrors=$errors;externalLinks=$external;activeSheet='Dashboard'} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'artifacts\validation\workbook-inspection.json') -Encoding UTF8
    if($errors.Count){throw ('Saved workbook has cached formula errors: '+($errors -join ', '))}
    Write-Output ('PASS: example workbook contains '+$sheets.Count+' sheets, '+$formulas+' formulas, and no cached formula errors.')
} finally {$zip.Dispose()}
