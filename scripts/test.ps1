param([switch]$Native, [switch]$OriginalSources)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    [void](New-Item -ItemType Directory -Path 'artifacts\test-results' -Force)
    $null=& (Join-Path $PSScriptRoot 'project-version.ps1')
    if($Native){
        $sessionPath=Join-Path $root '.runtime\native-test\session.json'
        $networkReady=Join-Path $root '.runtime\native-test\network-ready.json'
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $sessionPath) -Force)
        if(Test-Path -LiteralPath $sessionPath){Remove-Item -LiteralPath $sessionPath}
        if(Test-Path -LiteralPath $networkReady){Remove-Item -LiteralPath $networkReady}
        $proxy=Start-Process node.exe -WindowStyle Hidden -PassThru -WorkingDirectory $root -ArgumentList @(('"'+(Join-Path $root 'tests\fixtures\native-network.mjs')+'"')) -RedirectStandardOutput (Join-Path $root '.runtime\network-fixture.log') -RedirectStandardError (Join-Path $root '.runtime\network-fixture-error.log')
        $helper=$null
        try {
            $deadline=[DateTime]::UtcNow.AddSeconds(15)
            while(-not(Test-Path -LiteralPath $networkReady)){
                if($proxy.HasExited -or [DateTime]::UtcNow -gt $deadline){throw 'Native network fixture failed to start.'}
                Start-Sleep -Milliseconds 100
            }
            $helper=Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+(Join-Path $PSScriptRoot 'native-session.ps1')+'"'),'-AddInDirectory',('"'+(Join-Path $root 'artifacts\build\native')+'"'),'-ServiceRoot',('"'+(Join-Path $root '.runtime\native-test\root')+'"'),'-LifetimeSeconds','600') -RedirectStandardOutput (Join-Path $root '.runtime\native-harness.log') -RedirectStandardError (Join-Path $root '.runtime\native-harness-error.log')
            $deadline=[DateTime]::UtcNow.AddSeconds(45)
            while(-not (Test-Path -LiteralPath $sessionPath)){
                if($helper.HasExited -or [DateTime]::UtcNow -gt $deadline){throw 'Native harness failed to start. Inspect .runtime\native-harness-error.log.'}
                Start-Sleep -Milliseconds 200
            }
            & node scripts/test-native.mjs
            if($LASTEXITCODE -ne 0){throw 'Native integration tests failed.'}
        } finally {
            if($helper){
                @{id=[Guid]::NewGuid().ToString();action='stop'} | ConvertTo-Json | Set-Content -LiteralPath '.runtime\native-test\command.json' -Encoding UTF8
                if(-not $helper.WaitForExit(15000)){Write-Warning 'The test helper is still closing its Excel instance. Inspect its log; do not terminate unrelated Excel processes.'}
            }
            [IO.File]::WriteAllText((Join-Path $root '.runtime\native-test\network.json'),'{"stop":true}')
            if(-not $proxy.WaitForExit(5000)){Stop-Process -InputObject $proxy}
        }
        return
    }
    $sdk=Join-Path $env:ProgramFiles 'dotnet\dotnet.exe'
    & $sdk format whitespace LedgerLens.sln --no-restore --verify-no-changes --verbosity minimal
    if($LASTEXITCODE -ne 0){throw 'C# formatting check failed.'}
    & (Join-Path $root 'node_modules\.bin\prettier.cmd') --check web 'scripts/*.mjs' 'tests/**/*.js' 'tests/**/*.mjs' playwright.config.js
    if($LASTEXITCODE -ne 0){throw 'Web formatting check failed.'}
    & $sdk test tests/LedgerLens.Tests/LedgerLens.Tests.csproj -c Release --nologo --logger 'trx;LogFileName=unit-results.trx' --results-directory artifacts/test-results/unit
    if($LASTEXITCODE -ne 0){throw '.NET tests failed.'}
    & node --test tests/office-bridge.test.mjs | Tee-Object -FilePath artifacts/test-results/office-adapter.tap
    if($LASTEXITCODE -ne 0){throw 'Office bridge tests failed.'}
    if($OriginalSources){
        & node scripts/verify-sources.mjs
        if($LASTEXITCODE -ne 0){throw 'Original SEC source reconciliation failed. See docs/DATA.md for the optional archived inputs.'}
    }
    & node scripts/test-service.mjs
    if($LASTEXITCODE -ne 0){throw 'Service tests failed.'}
    # PowerShell 5 can turn a native stderr warning into a terminating error
    # when output is redirected. These tools report success via their exit code.
    try { $ErrorActionPreference='Continue'; & (Join-Path $root 'node_modules\.bin\playwright.cmd') test 2>&1 | ForEach-Object { $_.ToString() }; $code=$LASTEXITCODE }
    finally { $ErrorActionPreference='Stop' }
    if($code -ne 0){throw 'Browser tests failed.'}
    try { $ErrorActionPreference='Continue'; & (Join-Path $root 'node_modules\.bin\office-addin-manifest.cmd') validate office/manifest.xml 2>&1 | ForEach-Object { $_.ToString() } | Tee-Object -FilePath artifacts/test-results/office-manifest.txt; $code=$LASTEXITCODE }
    finally { $ErrorActionPreference='Stop' }
    if($code -ne 0){throw 'Office manifest validation failed.'}
} finally {Pop-Location}
