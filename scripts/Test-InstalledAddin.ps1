param(
    [string]$OutputDocument
)

$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw '请先保存文档并关闭所有 Word 窗口，再运行安装验证。'
}

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $addIn = $word.COMAddIns.Item('WordLatexVSTOAddin')
    if (-not $addIn.Connect) { $addIn.Connect = $true }
    if (-not $addIn.Connect) { throw 'Word 未能加载 WordLatexVSTOAddin。' }

    $document = $word.Documents.Add()
    $word.Selection.TypeText('$E=mc^2$')
    $document.Range(0, $document.Content.End - 1).Select()

    $automation = $addIn.Object
    if ($null -eq $automation) { throw '插件已加载，但未公开自动化验证接口。' }
    $converted = $automation.ConvertCurrentSelection()
    if ($converted -ne 1) { throw "预期转换 1 个公式，实际转换 $converted 个。" }
    if ($document.OMaths.Count -ne 1) { throw "预期生成 1 个 Word 原生公式，实际为 $($document.OMaths.Count) 个。" }

    if ($OutputDocument) {
        $target = [System.IO.Path]::GetFullPath($OutputDocument)
        $parent = Split-Path -Parent $target
        if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $document.SaveAs2($target, 16)
    }

    [pscustomobject]@{
        AddInConnected = $addIn.Connect
        Converted = $converted
        NativeOMathCount = $document.OMaths.Count
        OutputDocument = if ($OutputDocument) { [System.IO.Path]::GetFullPath($OutputDocument) } else { $null }
    }
}
finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
    if ($document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
