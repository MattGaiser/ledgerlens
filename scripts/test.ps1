param([switch]$Native)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    [void](New-Item -ItemType Directory -Path 'artifacts\validation' -Force)
    if($Native){
        $sessionPath=Join-Path $root '.runtime\native-test\session.json'
        if(Test-Path -LiteralPath $sessionPath){Remove-Item -LiteralPath $sessionPath}
        $helper=Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+(Join-Path $PSScriptRoot 'native-session.ps1')+'"'),'-AddInDirectory',('"'+(Join-Path $root 'artifacts\build\native')+'"'),'-LifetimeSeconds','600') -RedirectStandardOutput (Join-Path $root '.runtime\native-harness.log') -RedirectStandardError (Join-Path $root '.runtime\native-harness-error.log')
        try {
            $deadline=[DateTime]::UtcNow.AddSeconds(45)
            while(-not (Test-Path -LiteralPath $sessionPath)){
                if($helper.HasExited -or [DateTime]::UtcNow -gt $deadline){throw 'Native harness failed to start. Inspect .runtime\native-harness-error.log.'}
                Start-Sleep -Milliseconds 200
            }
            & node scripts/test-native.mjs
            if($LASTEXITCODE -ne 0){throw 'Native integration tests failed.'}
        } finally {
            @{id=[Guid]::NewGuid().ToString();action='stop'} | ConvertTo-Json | Set-Content -LiteralPath '.runtime\native-test\command.json' -Encoding UTF8
            if(-not $helper.WaitForExit(15000)){Write-Warning 'The test helper is still closing its Excel instance. Inspect its log; do not terminate unrelated Excel processes.'}
        }
        return
    }
    $sdk=Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    & $sdk test tests/LedgerLens.Tests/LedgerLens.Tests.csproj -c Release --nologo --logger 'trx;LogFileName=unit-results.trx' --results-directory artifacts/validation/unit
    if($LASTEXITCODE -ne 0){throw '.NET tests failed.'}
    & node --test tests/office-bridge.test.mjs
    if($LASTEXITCODE -ne 0){throw 'Office bridge tests failed.'}
    & node scripts/verify-sources.mjs
    if($LASTEXITCODE -ne 0){throw 'SEC source reconciliation failed. Use the source package with data/raw included.'}
    & node scripts/test-service.mjs
    if($LASTEXITCODE -ne 0){throw 'Service tests failed.'}
    # PowerShell 5 can turn a native stderr warning into a terminating error
    # when output is redirected. These tools report success via their exit code.
    try { $ErrorActionPreference='Continue'; & (Join-Path $root 'node_modules\.bin\playwright.cmd') test; $code=$LASTEXITCODE }
    finally { $ErrorActionPreference='Stop' }
    if($code -ne 0){throw 'Browser tests failed.'}
    try { $ErrorActionPreference='Continue'; & (Join-Path $root 'node_modules\.bin\office-addin-manifest.cmd') validate office/manifest.xml; $code=$LASTEXITCODE }
    finally { $ErrorActionPreference='Stop' }
    if($code -ne 0){throw 'Office manifest validation failed.'}
} finally {Pop-Location}
