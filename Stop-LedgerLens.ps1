$ErrorActionPreference='Stop'
$endpointPath=Join-Path $PSScriptRoot '.runtime\endpoint.json'
if (-not (Test-Path -LiteralPath $endpointPath)) { Write-Output 'No LedgerLens session found.'; return }
$endpoint=Get-Content -LiteralPath $endpointPath -Raw | ConvertFrom-Json
$process=Get-Process -Id $endpoint.ProcessId -ErrorAction SilentlyContinue
if (-not $process) { Write-Output 'LedgerLens service is already stopped.'; return }
$rootPath=[IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')+'\'
if ($process.ProcessName -ne 'LedgerLens.Service' -or -not $process.Path.StartsWith($rootPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Service identity did not match this LedgerLens folder.' }
$uri=[Uri]$endpoint.BaseUrl
if($uri.Host -ne '127.0.0.1' -or $uri.Scheme -ne 'http'){throw 'Invalid loopback endpoint.'}
$null=Invoke-RestMethod -Uri ($endpoint.BaseUrl+'/api/shutdown') -Method Post -Headers @{Authorization='Bearer '+$endpoint.Token} -TimeoutSec 5
if(-not $process.WaitForExit(5000)){throw 'The service is still draining requests. Retry after closing the LedgerLens pane.'}
Write-Output 'LedgerLens service stopped. Your Excel workbooks remain open.'
