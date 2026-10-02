param([ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$projectVersion=& (Join-Path $PSScriptRoot 'project-version.ps1')
if($Version -and $Version -ne $projectVersion){throw 'The requested package version does not match the project version.'}
$Version=$projectVersion
foreach($relative in @('native\LedgerLens.Core.dll','service\LedgerLens.Service.dll')){
    $built=Join-Path $root ('artifacts\build\'+$relative)
    if(-not(Test-Path -LiteralPath $built)){throw 'Run build.ps1 -SelfContained first.'}
    if([Reflection.AssemblyName]::GetAssemblyName($built).Version.ToString(3) -ne $Version){
        throw 'Built assemblies do not match the package version. Rebuild before packaging.'
    }
}
# The source archive includes historical validation evidence. Preserve its
# timestamps when filling reports not produced by a local rerun.
$priorEvidence=Join-Path $root 'validation'
$evidence=Join-Path $root 'artifacts\validation'
[void](New-Item -ItemType Directory -Path $evidence -Force)
if(Test-Path -LiteralPath $priorEvidence){
    foreach($item in Get-ChildItem -LiteralPath $priorEvidence -File){
        $target=Join-Path $evidence $item.Name
        if(-not(Test-Path -LiteralPath $target)){Copy-Item -LiteralPath $item.FullName -Destination $target}
    }
}
& (Join-Path $PSScriptRoot 'prepare-example.ps1')
Push-Location $root
try { & node scripts/report-validation.mjs; if($LASTEXITCODE -ne 0){throw 'Validation report generation failed.'} }
finally {Pop-Location}
$dist=Join-Path $root 'dist'
[void](New-Item -ItemType Directory -Path $dist -Force)
$stage=Join-Path $root ('artifacts\package-'+[Guid]::NewGuid().ToString('N'))
$release=Join-Path $stage ('LedgerLens-'+$Version+'-windows')
$source=Join-Path $stage ('LedgerLens-'+$Version+'-source')
[void](New-Item -ItemType Directory -Path $release,$source -Force)
foreach($name in @('native','service')){
    $built=Join-Path $root ('artifacts\build\'+$name)
    if(-not(Test-Path -LiteralPath $built)){throw 'Run build.ps1 -SelfContained first.'}
    $target=Join-Path $release $name
    [void](New-Item -ItemType Directory -Path $target)
    foreach($item in Get-ChildItem -LiteralPath $built){
        # Full unpacked native distribution: packed XLLs are deliberately excluded.
        if($item.Name -ne 'publish' -and $item.Extension -ne '.pdb'){Copy-Item -LiteralPath $item.FullName -Destination $target -Recurse}
    }
}
if(-not(Test-Path -LiteralPath (Join-Path $release 'service\coreclr.dll'))){throw 'Release service must be self-contained.'}
foreach($name in @('README.md','Start LedgerLens.cmd','Start-LedgerLens.ps1','Stop-LedgerLens.ps1','docs','web','office','licenses','THIRD-PARTY-NOTICES.md')){
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination $release -Recurse
}
[void](New-Item -ItemType Directory -Path (Join-Path $release 'data'))
Copy-Item -LiteralPath (Join-Path $root 'data\financials.json') -Destination (Join-Path $release 'data')
[void](New-Item -ItemType Directory -Path (Join-Path $release 'examples'),(Join-Path $release 'validation'))
Copy-Item -LiteralPath (Join-Path $root ('artifacts\LedgerLens-Analyst-Model-'+$Version+'.xlsx')) -Destination (Join-Path $release 'examples\LedgerLens-Analyst-Model.xlsx')
foreach($name in @('native-results.json','service-results.json','browser-results.json','source-reconciliation.json','live-openai.json','live-sec-sync.json','dependency-audit.txt','npm-audit.json','office-adapter.tap','office-manifest.txt','workbook-inspection.json')){
    Copy-Item -LiteralPath (Join-Path $root ('artifacts\validation\'+$name)) -Destination (Join-Path $release 'validation')
}
Copy-Item -LiteralPath (Join-Path $root 'artifacts\validation\unit\unit-results.trx') -Destination (Join-Path $release 'validation')
foreach($review in @('quality-review','second-review')){
    $reviewEvidence=Join-Path $root ('artifacts\validation\'+$review)
    if(Test-Path -LiteralPath $reviewEvidence){Copy-Item -LiteralPath $reviewEvidence -Destination (Join-Path $release ('validation\'+$review)) -Recurse}
}
$launchProof=Join-Path $root 'artifacts\validation\release-launch.json'
if(Test-Path -LiteralPath $launchProof){Copy-Item -LiteralPath $launchProof -Destination (Join-Path $release 'validation')}
# Explicit source allowlist also works from the source ZIP, without a .git directory.
$files=@('.editorconfig','.prettierrc.json','.gitattributes','.gitignore','README.md','THIRD-PARTY-NOTICES.md','Directory.Build.props','global.json','LedgerLens.sln','build.ps1','Start-LedgerLens.ps1','Stop-LedgerLens.ps1','Start LedgerLens.cmd','package.json','package-lock.json','playwright.config.js','data\financials.json')
foreach($directory in @('src','tests','scripts','docs','web','office','licenses')){
    $files+=@(Get-ChildItem -LiteralPath (Join-Path $root $directory) -File -Recurse | Where-Object {$_.FullName -notmatch '[\\/](bin|obj|EnvironmentCheck)[\\/]' -and $_.Extension -ne '.user'} | ForEach-Object {$_.FullName.Substring($root.Length+1)})
}
foreach($relative in $files){
    $destination=Join-Path $source $relative
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force)
    Copy-Item -LiteralPath (Join-Path $root $relative) -Destination $destination
}
Copy-Item -LiteralPath (Join-Path $root 'data\raw') -Destination (Join-Path $source 'data\raw') -Recurse
Copy-Item -LiteralPath (Join-Path $release 'validation') -Destination $source -Recurse
foreach($folder in @($release,$source)){
    $manifest=@(Get-ChildItem -LiteralPath $folder -File -Recurse | ForEach-Object { [ordered]@{path=$_.FullName.Substring($folder.Length+1).Replace('\','/');sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant();bytes=$_.Length} })
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $folder 'SHA256SUMS.json') -Encoding UTF8
}
& node (Join-Path $PSScriptRoot 'scan-release.mjs') $release $source
if($LASTEXITCODE -ne 0){throw 'Release credential scan failed.'}
foreach($folder in @($release,$source)){
    $archive=Join-Path $dist ((Split-Path -Leaf $folder)+'.zip')
    Compress-Archive -LiteralPath $folder -DestinationPath $archive -CompressionLevel Optimal -Force
    Write-Output $archive
}
Write-Output ('Staging folder: '+$stage)
