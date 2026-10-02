param([switch]$SelfContained, [string]$OutputRoot = '')
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$sdk = Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
if (-not (Test-Path -LiteralPath $sdk)) { $sdk = (Get-Command dotnet -ErrorAction Stop).Source }
if (-not $OutputRoot) { $OutputRoot = Join-Path $root 'artifacts\build' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
[void](New-Item -ItemType Directory -Path $OutputRoot -Force)
Push-Location $root
try {
    & $sdk restore .\LedgerLens.sln --locked-mode --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Locked dependency restore failed.' }
    & $sdk build .\src\LedgerLens.Excel\LedgerLens.Excel.csproj -c Release --no-restore --nologo -o (Join-Path $OutputRoot 'native')
    if ($LASTEXITCODE -ne 0) { throw 'Native add-in build failed.' }
    $selfContainedValue = if ($SelfContained) { 'true' } else { 'false' }
    & $sdk publish .\src\LedgerLens.Service\LedgerLens.Service.csproj -c Release -r win-x64 --self-contained $selfContainedValue --nologo -o (Join-Path $OutputRoot 'service')
    if ($LASTEXITCODE -ne 0) { throw 'Research service publish failed.' }
    Write-Output ('Built LedgerLens in ' + $OutputRoot)
} finally { Pop-Location }
