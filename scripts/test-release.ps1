param([string]$Archive, [ValidateRange(1024,65534)][int]$Port=17953)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
if(-not $Archive){
    $version=(Get-Content -LiteralPath (Join-Path $root 'package.json') -Raw | ConvertFrom-Json).version
    $Archive=Join-Path $root ('dist\LedgerLens-'+$version+'-windows.zip')
}
$Archive=(Resolve-Path -LiteralPath $Archive).Path
$folder=Join-Path $root ('artifacts\release test '+[Guid]::NewGuid().ToString('N'))
Expand-Archive -LiteralPath $Archive -DestinationPath $folder
$release=Join-Path $folder ([IO.Path]::GetFileNameWithoutExtension($Archive))
try {
    & (Join-Path $release 'Start-LedgerLens.ps1') -Validate -Port $Port
    $proof=Get-Content -LiteralPath (Join-Path $release '.runtime\launch-validation.json') -Raw | ConvertFrom-Json
    if($proof.status -ne 'PASS'){throw 'The launcher did not report success.'}
    Add-Type -AssemblyName System.IO.Compression,System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($proof.workbook)
    $errors=@(); $formulas=0
    try {
        foreach($entry in $zip.Entries | Where-Object {$_.FullName -match '^xl/worksheets/sheet\d+\.xml$'}) {
            $reader=New-Object IO.StreamReader($entry.Open())
            try {[xml]$xml=$reader.ReadToEnd()} finally {$reader.Dispose()}
            $formulas+=$xml.SelectNodes('//*[local-name()="f"]').Count
            foreach($cell in $xml.SelectNodes('//*[local-name()="c"][@t="e"]')){
                $errors+=($entry.FullName+':'+$cell.r+'='+$cell.v)
            }
        }
    } finally {$zip.Dispose()}
    if($errors.Count){throw ('The fresh launch saved formula errors: '+($errors -join ', '))}
    if($formulas -lt 50){throw 'The saved workbook does not contain the expected analyst formulas.'}
    $proof | Add-Member -NotePropertyName archive -NotePropertyValue ([IO.Path]::GetFileName($Archive))
    $proof | Add-Member -NotePropertyName workflow -NotePropertyValue 'Fresh ZIP extraction in a path containing spaces; native launch, save/reopen, pane initialization and saved XLSX inspection.'
    $proof | Add-Member -NotePropertyName savedFormulas -NotePropertyValue $formulas
    $proof | Add-Member -NotePropertyName savedFormulaErrors -NotePropertyValue $errors.Count
    $evidence=Join-Path $root 'artifacts\validation'
    [void](New-Item -ItemType Directory -Path $evidence -Force)
    $proof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $evidence 'release-launch.json') -Encoding UTF8
    Write-Output ('PASS: fresh release saved '+$formulas+' formulas with no cached errors.')
} finally {
    if(Test-Path -LiteralPath (Join-Path $release '.runtime\endpoint.json')){
        & (Join-Path $release 'Stop-LedgerLens.ps1')
    }
}
