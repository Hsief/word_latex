param(
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Split-Path -Parent $PSScriptRoot)).Path
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
        /p:Configuration=Release '/p:Platform=Any CPU' /p:SignManifests=false
    if ($LASTEXITCODE -ne 0) { throw 'Solution build failed.' }

    & (Join-Path $root 'tests\AfterMathCore.Tests\bin\Release\AfterMathCore.Tests.exe')
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }

    & $msbuild (Join-Path $root 'src\WordLatexAddin\WordLatexAddin.csproj') /t:Publish /m `
        /p:Configuration=Release /p:Platform=AnyCPU "/p:PublishUrl=$publish\" `
        "/p:ApplicationVersion=$Version.0" /p:SignManifests=true `
        "/p:ManifestCertificateThumbprint=$($cert.Thumbprint)"
    if ($LASTEXITCODE -ne 0) { throw 'VSTO publish failed.' }

    & $iscc "/DMyAppVersion=$Version" (Join-Path $root 'installer\setup\WordLatexVSTO.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }

    $setup = Join-Path $publish 'WordLatexVSTO_Setup.exe'
    if (-not (Test-Path $setup)) { throw 'Installer output was not found.' }
    Write-Host "Created $setup"
}
finally {
    if ($cert) { Remove-Item ("Cert:\CurrentUser\My\" + $cert.Thumbprint) -Force -ErrorAction SilentlyContinue }
}
