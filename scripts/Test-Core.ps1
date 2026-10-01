param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root "tests\AfterMathCore.Tests\bin\$Configuration"
New-Item -ItemType Directory -Path $output -Force | Out-Null

$compiler = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$framework = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319"
$coreSources = Get-ChildItem (Join-Path $root "src\AfterMathCore") -Filter *.cs -Recurse | ForEach-Object FullName
$testSource = Join-Path $root "tests\AfterMathCore.Tests\Program.cs"
$scannerSource = Join-Path $root "src\WordLatexAddin\LatexScanner.cs"
$coreDll = Join-Path $output "AfterMathCore.dll"
$testExe = Join-Path $output "AfterMathCore.Tests.exe"
$xmlLinq = Join-Path $framework "System.Xml.Linq.dll"

& $compiler /nologo /target:library /optimize+ /out:$coreDll /reference:$xmlLinq $coreSources
if ($LASTEXITCODE -ne 0) { throw "AfterMathCore compilation failed." }

& $compiler /nologo /target:exe /optimize+ /out:$testExe /reference:$coreDll /reference:$xmlLinq $testSource $scannerSource
if ($LASTEXITCODE -ne 0) { throw "Test compilation failed." }

& $testExe
if ($LASTEXITCODE -ne 0) { throw "AfterMathCore tests failed." }
