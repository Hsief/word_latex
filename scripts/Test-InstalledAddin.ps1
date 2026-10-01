param(
    [string]$OutputDocument
)

$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close all Word windows before running the installed add-in test.'
}

$word = $null
$document = $null
$wordProcess = $null
try {
    $wordPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot' -ErrorAction Stop).Path
    $wordProcess = Start-Process -FilePath (Join-Path $wordPath 'WINWORD.EXE') -ArgumentList '/w' -PassThru
    for ($attempt = 0; $attempt -lt 30 -and $null -eq $word; $attempt++) {
        Start-Sleep -Milliseconds 500
        try { $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
        catch { }
    }
    if ($null -eq $word) { throw 'Word started, but the normal desktop instance was not available.' }
    $word.Visible = $false
    $addIn = $null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $addIn = $word.COMAddIns.Item('WordLatexVSTOAddin')
        if ($addIn.Connect) { break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $addIn.Connect) { throw 'Word did not load WordLatexVSTOAddin during normal startup.' }

    $document = $word.ActiveDocument
    if ($null -eq $document) { $document = $word.Documents.Add() }
    $document.Content.Text = ''
    $word.Selection.TypeText('$E=mc^2$')
    $document.Range(0, $document.Content.End - 1).Select()

    $automation = $addIn.Object
    if ($null -eq $automation) { throw 'The add-in loaded but did not expose its verification API.' }
    if (-not $automation.RibbonCustomUiRequested) { throw 'Word loaded the add-in but never requested its Ribbon XML.' }
    if (-not $automation.RibbonLoaded) { throw 'Word requested the Ribbon XML, but Office rejected it before the onLoad callback.' }
    $converted = $automation.ConvertCurrentSelection()
    if ($converted -ne 1) { throw "Expected one conversion; received $converted." }
    if ($document.OMaths.Count -ne 1) { throw "Expected one native Word equation; received $($document.OMaths.Count)." }

    if ($OutputDocument) {
        $target = [System.IO.Path]::GetFullPath($OutputDocument)
        $parent = Split-Path -Parent $target
        if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $document.SaveAs2($target, 16)
    }

    [pscustomobject]@{
        AddInConnected = $addIn.Connect
        RibbonCustomUiRequested = $automation.RibbonCustomUiRequested
        RibbonLoaded = $automation.RibbonLoaded
        RequestedRibbonId = $automation.RequestedRibbonId
        Converted = $converted
        NativeOMathCount = $document.OMaths.Count
        OutputDocument = if ($OutputDocument) { [System.IO.Path]::GetFullPath($OutputDocument) } else { $null }
    }
}
finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
    elseif ($wordProcess -and -not $wordProcess.HasExited) {
        [void]$wordProcess.CloseMainWindow()
        if (-not $wordProcess.WaitForExit(5000)) { $wordProcess.Kill() }
    }
    if ($document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
