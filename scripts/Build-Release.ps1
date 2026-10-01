param(
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version = '1.0.5'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Split-Path -Parent $PSScriptRoot)).Path
$staging = Join-Path $root 'installer\staging'
$publish = Join-Path $root 'installer\publish'
if (-not $publish.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Publish directory must stay inside the repository.'
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) { throw 'Visual Studio 2022 was not found.' }
$vsPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsPath) { throw 'MSBuild was not found.' }
$msbuild = Join-Path $vsPath 'MSBuild\Current\Bin\MSBuild.exe'
$officeTargets = Join-Path $vsPath 'MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets'
if (-not (Test-Path $officeTargets)) {
    throw 'Install the Visual Studio “Office/SharePoint development” workload first.'
}

$iscc = (Get-Command iscc.exe -ErrorAction SilentlyContinue).Source
if (-not $iscc) { $iscc = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe' }
if (-not (Test-Path $iscc)) { throw 'Inno Setup 6 was not found.' }

$wix = Get-ChildItem (Join-Path ${env:ProgramFiles(x86)} 'WiX Toolset v3.*\bin\candle.exe') -ErrorAction SilentlyContinue |
    Sort-Object FullName -Descending | Select-Object -First 1
if (-not $wix) { throw 'WiX Toolset v3 was not found.' }
$wixBin = Split-Path $wix.FullName

Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $staging -Force | Out-Null
Get-ChildItem $publish -Force | Where-Object { $_.Name -ne '.gitkeep' } | Remove-Item -Recurse -Force
$cert = $null
try {
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject 'CN=WordLatexVSTO Local Build' `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy Exportable `
        -NotAfter (Get-Date).AddYears(2)

    & $msbuild (Join-Path $root 'WordLatexVSTO.sln') /restore /t:Rebuild /m `
        /p:Configuration=Release '/p:Platform=Any CPU' /p:SignManifests=true `
        "/p:ApplicationVersion=$Version.0" `
        "/p:ManifestCertificateThumbprint=$($cert.Thumbprint)"
    if ($LASTEXITCODE -ne 0) { throw 'Solution build failed.' }

    & (Join-Path $root 'tests\AfterMathCore.Tests\bin\Release\AfterMathCore.Tests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }

    $bin = Join-Path $root 'src\WordLatexAddin\bin\Release'
    Get-ChildItem $bin -File | Where-Object {
        $_.Extension -in @('.dll', '.config', '.manifest', '.vsto')
    } | Copy-Item -Destination $staging -Force
    foreach ($required in @('WordLatexVSTOAddin.vsto', 'WordLatexVSTOAddin.dll.manifest', 'WordLatexVSTOAddin.dll')) {
        if (-not (Test-Path (Join-Path $staging $required))) { throw "The flat MSI layout is missing $required." }
    }
    [xml]$deploymentManifest = Get-Content (Join-Path $staging 'WordLatexVSTOAddin.vsto') -Raw
    $applicationCodebase = $deploymentManifest.assembly.dependency.dependentAssembly.codebase
    if ($applicationCodebase -ne 'WordLatexVSTOAddin.dll.manifest') {
        throw "Unexpected MSI manifest codebase: $applicationCodebase"
    }

    New-Item -ItemType Directory -Path (Join-Path $staging 'docs') -Force | Out-Null
    Copy-Item (Join-Path $root 'docs\INSTALL.md'), (Join-Path $root 'docs\USER_GUIDE.md') (Join-Path $staging 'docs') -Force
    Copy-Item (Join-Path $root 'README.md'), (Join-Path $root 'LICENSE'), (Join-Path $root 'THIRD_PARTY_NOTICES.md') $staging -Force

    $harvest = Join-Path $root 'installer\setup\HarvestedFiles.wxs'
    $wixObj = Join-Path $root 'installer\setup\obj'
    Remove-Item $wixObj -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $wixObj -Force | Out-Null
    & (Join-Path $wixBin 'heat.exe') dir $staging -nologo -cg PublishedFiles -dr INSTALLFOLDER `
        -gg -g1 -sfrag -srd -sreg -var var.PublishDir -out $harvest
    if ($LASTEXITCODE -ne 0) { throw 'WiX file harvesting failed.' }
    & (Join-Path $wixBin 'candle.exe') -nologo -arch x64 "-dProductVersion=$Version" `
        -out (Join-Path $wixObj 'Product.wixobj') (Join-Path $root 'installer\setup\Product.wxs')
    if ($LASTEXITCODE -ne 0) { throw 'WiX product compilation failed.' }
    & (Join-Path $wixBin 'candle.exe') -nologo -arch x64 "-dPublishDir=$staging" `
        -out (Join-Path $wixObj 'HarvestedFiles.wixobj') $harvest
    if ($LASTEXITCODE -ne 0) { throw 'WiX payload compilation failed.' }
    & (Join-Path $wixBin 'light.exe') -nologo -out (Join-Path $publish 'WordLatexVSTO.msi') `
        (Join-Path $wixObj 'Product.wixobj') (Join-Path $wixObj 'HarvestedFiles.wixobj')
    if ($LASTEXITCODE -ne 0) { throw 'MSI package build failed.' }

    & $iscc "/DMyAppVersion=$Version" (Join-Path $root 'installer\setup\WordLatexVSTO.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }

    $setup = Join-Path $publish 'WordLatexVSTO_Setup.exe'
    if (-not (Test-Path $setup)) { throw 'Installer output was not found.' }
    Write-Host "Created $setup"
}
finally {
    if ($cert) { Remove-Item ("Cert:\CurrentUser\My\" + $cert.Thumbprint) -Force -ErrorAction SilentlyContinue }
}
