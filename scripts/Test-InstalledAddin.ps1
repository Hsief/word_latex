param(
    [string]$OutputDocument
)

# This integration test requires an installed add-in and real desktop Word.
# It creates and modifies only its own new document, never ActiveDocument.
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Run this test in Windows PowerShell 5.1; PowerShell 7 does not expose Marshal.GetActiveObject.'
}
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close all Word windows before running the installed add-in test.'
}
if ($OutputDocument -and (Test-Path -LiteralPath $OutputDocument)) {
    throw 'The test output already exists; choose a new path instead of overwriting a document.'
}

function Add-TestText {
    param($Document, [string]$Text)
    $tail = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
    $tail.InsertBefore($Text)
}

function Get-NonEquationText {
    param($Document)
    $bounds = @()
    for ($index = 1; $index -le $Document.OMaths.Count; $index++) {
        $range = $Document.OMaths.Item($index).Range
        $bounds += [pscustomobject]@{ Start = [int]$range.Start; End = [int]$range.End }
    }
    $builder = New-Object System.Text.StringBuilder
    $cursor = $Document.Content.Start
    foreach ($bound in ($bounds | Sort-Object Start)) {
        if ($bound.Start -lt $cursor) { throw 'The test encountered overlapping equation ranges.' }
        [void]$builder.Append([string]$Document.Range($cursor, $bound.Start).Text)
        $cursor = $bound.End
    }
    [void]$builder.Append([string]$Document.Range($cursor, $Document.Content.End).Text)
    return $builder.ToString()
}

function Get-EquationXml {
    param($Equation)
    [xml]$xml = $Equation.Range.WordOpenXML
    $namespaces = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $namespaces.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
    $math = $xml.SelectSingleNode('//m:oMath', $namespaces)
    if ($null -eq $math) { throw 'A native equation has no OMML content.' }
    return $math.OuterXml
}

function Assert-EqualText {
    param([string]$Expected, [string]$Actual, [string]$Label)
    if (-not [string]::Equals($Expected, $Actual, [StringComparison]::Ordinal)) {
        $difference = 0
        while ($difference -lt [Math]::Min($Expected.Length, $Actual.Length) -and
               $Expected[$difference] -ceq $Actual[$difference]) { $difference++ }
        throw "$Label changed at character $difference (expected length $($Expected.Length), actual length $($Actual.Length))."
    }
}

function Invoke-TestConversion {
    param($Word, $Document, $Automation)
    $Document.Activate()
    $Document.Content.Select()
    if ($Word.Selection.Range.Document.FullName -cne $Document.FullName) {
        throw 'The test document lost the active selection; conversion was not requested.'
    }
    return $Automation.ConvertCurrentSelection()
}

$root = Split-Path -Parent $PSScriptRoot
$fixturePath = Join-Path $root 'tests\fixtures\voxel-equations.txt'
$fixture = [IO.File]::ReadAllText($fixturePath, [Text.Encoding]::UTF8)
# Word paragraph marks are CR, even when the checked-in fixture uses LF.
$fixture = $fixture.Replace("`r`n", "`n").Replace("`n", "`r")
$gridFormula = '$$' + "`r" +
    'D=\lceil L_z/\Delta_z\rceil,\quad H=\lceil L_y/\Delta_y\rceil,\quad W=\lceil L_x/\Delta_x\rceil' +
    "`r" + '$$'
$featureFormula = '\[' + "`r" +
    '\mathbf{f}_i=\left[I_i,\frac{x_i-x_v}{\Delta_x},\frac{y_i-y_v}{\Delta_y},\frac{z_i-z_v}{\Delta_z}\right]' +
    "`r" + '\]'
# Expected ranges are fixed literals, independent of the add-in's scanner.
$formulaSources = @(
    '$L_x$', '$L_y$', '$L_z$', '\(\Delta_x\)', '\(\Delta_y\)', '\(\Delta_z\)',
    '$D$', '$H$', '$W$', $gridFormula, '$\mathcal{V}_t$', '$v$',
    '$\mathcal{P}_v$', '\((x_v,y_v,z_v)\)', '$I_i$', $featureFormula,
    '$a_1$', '\(\beta\)', '$p=1$'
)

$word = $null
$document = $null
$wordProcess = $null
try {
    $wordPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot' -ErrorAction Stop).Path
    $wordProcess = Start-Process -FilePath (Join-Path $wordPath 'WINWORD.EXE') -WindowStyle Hidden -PassThru
    for ($attempt = 0; $attempt -lt 30 -and $null -eq $word; $attempt++) {
        Start-Sleep -Milliseconds 500
        try { $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
        catch { }
    }
    if ($null -eq $word) { throw 'Word started, but the normal desktop instance was not available.' }
    $addIn = $null
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $addIn = $word.COMAddIns.Item('WordLatexVSTOAddin')
        if ($addIn.Connect) { break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $addIn.Connect) { throw 'Word did not load WordLatexVSTOAddin during normal startup.' }

    # Never use/clear ActiveDocument: a user may open a document while Word starts.
    $automation = $addIn.Object
    if ($null -eq $automation) { throw 'The add-in loaded but did not expose its verification API.' }
    if (-not $automation.RibbonCustomUiRequested) { throw 'Word loaded the add-in but never requested its Ribbon XML.' }
    if (-not $automation.RibbonLoaded) { throw 'Word requested the Ribbon XML, but Office rejected it before the onLoad callback.' }
    $word.Visible = $true
    if (-not $automation.ActivateRibbonTab()) { throw "Office loaded the Ribbon but could not activate the custom tab: $($automation.RibbonActivationError)" }

    $document = $word.Documents.Add()
    $document.Content.Text = "E=mc^2`r"
    $nativeRange = $document.Range(0, 6)
    $nativeMathRange = $document.OMaths.Add($nativeRange)
    $nativeMathRange.OMaths.Item(1).BuildUp()
    if ($document.OMaths.Count -ne 1) { throw 'Could not create the pre-existing native equation.' }
    $existingMathXml = Get-EquationXml ($document.OMaths.Item(1))

    Add-TestText $document "`rExisting field result: "
    $fieldRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $field = $document.Fields.Add($fieldRange, -1, '= 2 + 3', $false)
    [void]$field.Update()
    if ($field.Result.Text.Trim() -ne '5') { throw 'Could not create the field used by the integration test.' }
    Add-TestText $document " must stay unchanged.`r"
    Add-TestText $document $fixture

    $tableRange = $document.Range($document.Content.End - 1, $document.Content.End - 1)
    $table = $document.Tables.Add($tableRange, 2, 2)
    $table.Cell(1, 1).Range.Text = 'Cell one: $a_1$; preserve this text.'
    $table.Cell(1, 2).Range.Text = 'Cell two: \(\beta\); preserve this text.'
    $table.Cell(2, 1).Range.Text = 'Cell three: $p=1$; preserve this text.'
    $table.Cell(2, 2).Range.Text = 'Cell four has no equation.'

    $fieldCode = [string]$field.Code.Text
    $fieldResult = [string]$field.Result.Text
    $expectedPlainText = Get-NonEquationText $document
    foreach ($formula in $formulaSources) {
        $first = $expectedPlainText.IndexOf($formula, [StringComparison]::Ordinal)
        if ($first -lt 0 -or $expectedPlainText.IndexOf($formula, $first + $formula.Length, [StringComparison]::Ordinal) -ge 0) {
            throw "Fixture must contain exactly one occurrence of formula: $formula"
        }
        $expectedPlainText = $expectedPlainText.Replace($formula, '')
    }
    $expectedConversions = $formulaSources.Count
    $converted = Invoke-TestConversion $word $document $automation
    if ($converted -ne $expectedConversions) {
        throw "Expected $expectedConversions conversions; received $converted."
    }
    if ($document.OMaths.Count -ne $expectedConversions + 1) {
        throw "Expected $($expectedConversions + 1) native equations; received $($document.OMaths.Count)."
    }
    Assert-EqualText $expectedPlainText (Get-NonEquationText $document) 'Ordinary text and paragraph/cell boundaries'
    if ($document.Content.Text -match '\$|\\') { throw 'LaTeX delimiters or command backslashes remain in the converted document.' }
    Assert-EqualText $existingMathXml (Get-EquationXml ($document.OMaths.Item(1))) 'Pre-existing native equation'
    if ($document.Fields.Count -ne 1) { throw 'The field count changed.' }
    Assert-EqualText $fieldCode ([string]$document.Fields.Item(1).Code.Text) 'Field code'
    Assert-EqualText $fieldResult ([string]$document.Fields.Item(1).Result.Text) 'Field result'
    if ($document.Tables.Count -ne 1 -or $document.Tables.Item(1).Rows.Count -ne 2 -or $document.Tables.Item(1).Columns.Count -ne 2) {
        throw 'The table structure changed.'
    }

    $textAfterFirstPass = [string]$document.Content.Text
    $mathAfterFirstPass = @()
    for ($index = 1; $index -le $document.OMaths.Count; $index++) {
        $mathAfterFirstPass += Get-EquationXml ($document.OMaths.Item($index))
    }
    $convertedAgain = Invoke-TestConversion $word $document $automation
    if ($convertedAgain -ne 0) { throw "The second conversion changed $convertedAgain equations; expected zero." }
    Assert-EqualText $textAfterFirstPass ([string]$document.Content.Text) 'Second-pass document text'
    if ($document.OMaths.Count -ne $mathAfterFirstPass.Count) { throw 'The second conversion changed the equation count.' }
    for ($index = 1; $index -le $document.OMaths.Count; $index++) {
        Assert-EqualText ($mathAfterFirstPass[$index - 1]) (Get-EquationXml ($document.OMaths.Item($index))) "Second-pass equation $index"
    }

    if ($OutputDocument) {
        $target = [System.IO.Path]::GetFullPath($OutputDocument)
        if (Test-Path -LiteralPath $target) { throw 'The test output now exists; it will not be overwritten.' }
        $parent = Split-Path -Parent $target
        if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $document.SaveAs2($target, 16)
    }

    [pscustomobject]@{
        AddInConnected = $addIn.Connect
        RibbonCustomUiRequested = $automation.RibbonCustomUiRequested
        RibbonLoaded = $automation.RibbonLoaded
        RibbonTabActivated = $true
        RequestedRibbonId = $automation.RequestedRibbonId
        Converted = $converted
        ConvertedOnSecondPass = $convertedAgain
        NativeOMathCount = $document.OMaths.Count
        OrdinaryTextPreserved = $true
        ExistingEquationPreserved = $true
        FieldAndTablePreserved = $true
        OutputDocument = if ($OutputDocument) { [System.IO.Path]::GetFullPath($OutputDocument) } else { $null }
    }
}
finally {
    if ($document) { $document.Close(0) }
    if ($word -and $word.Documents.Count -eq 0) { $word.Quit() }
    # If COM attachment failed or another document appeared, leave Word alone.
    # Do not CloseMainWindow/Kill: it could be a user's unsaved document.
    if ($document) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
