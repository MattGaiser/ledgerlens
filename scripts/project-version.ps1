$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
[xml]$properties = Get-Content -LiteralPath (Join-Path $root 'Directory.Build.props') -Raw
$version = [string]$properties.Project.PropertyGroup.Version
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw 'Directory.Build.props must define a three-part release version.' }
$package = Get-Content -LiteralPath (Join-Path $root 'package.json') -Raw | ConvertFrom-Json
# PowerShell 5 cannot deserialize npm's empty-string root package key.
$lockVersions = & node -e "const fs=require('node:fs');const lock=JSON.parse(fs.readFileSync(process.argv[1],'utf8'));process.stdout.write(JSON.stringify([lock.version,lock.packages[''].version]));" (Join-Path $root 'package-lock.json')
if ($LASTEXITCODE -ne 0) { throw 'Could not read versions from the npm lockfile.' }
$lockVersions = $lockVersions | ConvertFrom-Json
[xml]$office = Get-Content -LiteralPath (Join-Path $root 'office\manifest.xml') -Raw
$officeVersion = $office.SelectSingleNode('//*[local-name()="Version"]').InnerText
if ($package.version -ne $version -or $lockVersions[0] -ne $version -or $lockVersions[1] -ne $version -or $officeVersion -ne ($version+'.0')) {
    throw 'Release versions disagree across .NET, npm, the dependency lockfile, and the Office manifest.'
}
Write-Output $version
