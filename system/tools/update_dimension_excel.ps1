<#
.SYNOPSIS
    COM-free writer that saves one product's or carton's dimensions into the live
    "Products Physical Dimension - FG" workbook by editing the underlying Zip/XML
    (OOXML) package directly, then triggers a headless cache rebuild.

.DESCRIPTION
    No Excel process is ever launched. The workbook is staged to system\staging,
    edited as a plain zip archive via System.IO.Compression + System.Xml, then
    atomically copied back over the original. A raw exclusive-open probe aborts
    in under a second if the file is locked (Excel or OneDrive), with no retry
    loop. Every write is preceded by a timestamped backup, same as before.

    This intentionally reproduces the write semantics of the previous COM-based
    writer (literal values into Item number/Unit/Width/Depth/Height, or
    carton_id/carton_name/dims/volume), not the workbook's own array-formula
    auto-fill template - typing into those template cells already overwrites
    their formula in normal use, so this writer does the same thing on purpose.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File update_dimension_excel.ps1 -Type product -Id NPI-2026-001 -LengthMm 225 -WidthMm 130 -HeightMm 55 -TrackId TRK-ABCD1234
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File update_dimension_excel.ps1 -Type carton -Id CTN-NPI-01 -Name "Master Carton NPI" -LengthMm 400 -WidthMm 300 -HeightMm 250
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('product', 'carton')][string]$Type,
    [Parameter(Mandatory = $true)][string]$Id,
    [string]$Name = '',
    [Parameter(Mandatory = $true)][double]$LengthMm,
    [Parameter(Mandatory = $true)][double]$WidthMm,
    [Parameter(Mandatory = $true)][double]$HeightMm,
    # Optional. Product only: written to the 'Gross weight' (kg) column of the EA row on the Dimension sheet.
    # 0 (the default) means "not supplied" and leaves the cell exactly as it was.
    [double]$GrossWeightKg = 0,
    [string]$TrackId = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ([string]::IsNullOrWhiteSpace($TrackId)) {
    $TrackId = 'TRK-' + [Guid]::NewGuid().ToString().Substring(0, 8).ToUpperInvariant()
}

$NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NS_XML = 'http://www.w3.org/XML/1998/namespace'
$NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

# ---------------------------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------------------------
$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$SystemDir = Split-Path -Parent $ScriptRoot            # .../system
$AppRoot = Split-Path -Parent $SystemDir                # .../WVN_Master_Carton_Fit_Tool
$ToolsParent = Split-Path -Parent $AppRoot               # .../MasterCarton fit check tool
$WeeklyBomDir = Join-Path $ToolsParent '1. Weekly BOM'
$CacheStatePath = Join-Path $SystemDir 'generated\cache_state.ini'
$LogDir = Join-Path $SystemDir 'logs'
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogPath = Join-Path $LogDir 'excel_writeback.log'
$StagingDir = Join-Path $SystemDir 'staging'
if (-not (Test-Path -LiteralPath $StagingDir)) { New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null }

function Write-WbLog {
    param([string]$Message)
    $line = '{0} [{1}] {2}' -f ([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')), $TrackId, $Message
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Read-IniFile {
    param([string]$FilePath)
    $ini = @{}
    if (-not (Test-Path -LiteralPath $FilePath)) { return $ini }
    $section = ''
    foreach ($line in Get-Content -LiteralPath $FilePath) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith(';') -or $t.StartsWith('#')) { continue }
        if ($t.StartsWith('[') -and $t.EndsWith(']')) { $section = $t.Substring(1, $t.Length - 2).Trim(); $ini[$section] = @{}; continue }
        if ($t -match '^([^=]+)=(.*)$') {
            $k = $matches[1].Trim(); $v = $matches[2].Trim()
            if ($section) { $ini[$section][$k] = $v }
        }
    }
    return $ini
}

function Send-Result {
    <# The listener parses the LAST stdout line as JSON, so this must be the only thing
       this script ever writes to the output stream. track_id rides on every response so
       a failure in the UI can be traced straight back to this exact run in the log. #>
    param([bool]$Ok, [hashtable]$Fields)
    $obj = [ordered]@{ ok = $Ok; track_id = $TrackId }
    foreach ($k in $Fields.Keys) { $obj[$k] = $Fields[$k] }
    Write-Output ($obj | ConvertTo-Json -Compress -Depth 6)
}

# ---------------------------------------------------------------------------------------------
# OOXML helpers - all operate on plain System.Xml.XmlDocument / System.IO.Compression types.
# ---------------------------------------------------------------------------------------------
function Get-ColIndexFromRef {
    param([string]$Ref)
    if ($Ref -notmatch '^([A-Z]+)\d+$') { throw "Bad cell reference: $Ref" }
    $letters = $matches[1]
    $idx = 0
    foreach ($ch in $letters.ToCharArray()) { $idx = $idx * 26 + ([int]$ch - [int][char]'A' + 1) }
    return $idx
}

function Get-ColumnLetter {
    param([int]$Index)
    $letters = ''
    $n = $Index
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $letters = [char](65 + $rem) + $letters
        $n = [int](($n - $rem - 1) / 26)
    }
    return $letters
}

function Read-ZipXml {
    param($Zip, [string]$EntryPath)
    $entry = $Zip.GetEntry($EntryPath)
    if (-not $entry) { throw "Zip entry not found: $EntryPath" }
    $doc = New-Object Xml.XmlDocument
    $stream = $entry.Open()
    try { $doc.Load($stream) } finally { $stream.Dispose() }
    return $doc
}

function Test-ZipEntry {
    param($Zip, [string]$EntryPath)
    return [bool]$Zip.GetEntry($EntryPath)
}

function Write-ZipXml {
    param($Zip, [string]$EntryPath, [Xml.XmlDocument]$Doc)
    $oldEntry = $Zip.GetEntry($EntryPath)
    if ($oldEntry) { $oldEntry.Delete() }
    $newEntry = $Zip.CreateEntry($EntryPath, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $newEntry.Open()
    try {
        $settings = New-Object Xml.XmlWriterSettings
        $settings.Encoding = New-Object Text.UTF8Encoding($false)   # UTF-8, no BOM
        $settings.OmitXmlDeclaration = $false
        $writer = [Xml.XmlWriter]::Create($stream, $settings)
        try { $Doc.Save($writer) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Remove-ZipEntryIfPresent {
    param($Zip, [string]$EntryPath)
    $entry = $Zip.GetEntry($EntryPath)
    if ($entry) { $entry.Delete() }
}

function Get-SiText {
    <# A shared-string <si> holds either one <t> directly or several <r><t> runs. #>
    param($SiEl)
    $direct = $SiEl['t']
    if ($direct) { return $direct.InnerText }
    $sb = New-Object Text.StringBuilder
    foreach ($run in $SiEl.ChildNodes) {
        if ($run.LocalName -eq 'r') {
            $rt = $run['t']
            if ($rt) { [void]$sb.Append($rt.InnerText) }
        }
    }
    return $sb.ToString()
}

function Resolve-CellText {
    <# Handles every cell-type variant that actually appears in this workbook: shared
       string (t="s"), inline string (t="inlineStr"), formula-cached string (t="str"),
       and plain numeric (no t attribute). A missing cell reads as blank, same as Excel. #>
    param($CellEl, [Xml.XmlDocument]$SstDoc)
    if (-not $CellEl) { return '' }
    $t = $CellEl.GetAttribute('t')
    if ($t -eq 's') {
        $vEl = $CellEl['v']
        if (-not $vEl -or [string]::IsNullOrWhiteSpace($vEl.InnerText)) { return '' }
        $idx = [int]$vEl.InnerText
        $siNodes = $SstDoc.DocumentElement.ChildNodes
        if ($idx -lt 0 -or $idx -ge $siNodes.Count) { return '' }
        return (Get-SiText $siNodes[$idx])
    }
    elseif ($t -eq 'inlineStr') {
        $isEl = $CellEl['is']
        if (-not $isEl) { return '' }
        return (Get-SiText $isEl)
    }
    else {
        $vEl = $CellEl['v']
        if (-not $vEl) { return '' }
        return $vEl.InnerText
    }
}

function Add-SharedString {
    param([Xml.XmlDocument]$SstDoc, [string]$Text)
    $sst = $SstDoc.DocumentElement
    $si = $SstDoc.CreateElement('si', $NS_MAIN)
    $t = $SstDoc.CreateElement('t', $NS_MAIN)
    $t.InnerText = $Text
    if ($Text -ne $Text.Trim() -or $Text -eq '') { $t.SetAttribute('space', $NS_XML, 'preserve') }
    [void]$si.AppendChild($t)
    $idx = $sst.ChildNodes.Count
    [void]$sst.AppendChild($si)
    $countAttr = $sst.GetAttribute('count'); $uniqueAttr = $sst.GetAttribute('uniqueCount')
    $countVal = 0; $uniqueVal = 0
    [void][int]::TryParse($countAttr, [ref]$countVal); [void][int]::TryParse($uniqueAttr, [ref]$uniqueVal)
    $sst.SetAttribute('count', ([string]($countVal + 1)))
    $sst.SetAttribute('uniqueCount', ([string]($uniqueVal + 1)))
    return $idx
}

function Insert-CellInOrder {
    param($RowEl, $NewCellEl, [int]$TargetIdx)
    foreach ($existing in @($RowEl.ChildNodes)) {
        if ($existing.LocalName -ne 'c') { continue }
        if ((Get-ColIndexFromRef $existing.GetAttribute('r')) -gt $TargetIdx) {
            [void]$RowEl.InsertBefore($NewCellEl, $existing)
            return
        }
    }
    [void]$RowEl.AppendChild($NewCellEl)
}

function Get-OrCreateCell {
    param([Xml.XmlDocument]$Doc, $RowEl, [string]$Ref, [string]$StyleFromRef)
    foreach ($c in $RowEl.ChildNodes) {
        if ($c.LocalName -eq 'c' -and $c.GetAttribute('r') -eq $Ref) { return $c }
    }
    $c = $Doc.CreateElement('c', $NS_MAIN)
    $c.SetAttribute('r', $Ref)
    if ($StyleFromRef) {
        foreach ($existing in $RowEl.ChildNodes) {
            if ($existing.LocalName -eq 'c' -and $existing.GetAttribute('r') -eq $StyleFromRef) {
                $s = $existing.GetAttribute('s')
                if ($s) { $c.SetAttribute('s', $s) }
                break
            }
        }
    }
    Insert-CellInOrder $RowEl $c (Get-ColIndexFromRef $Ref)
    return $c
}

function Clear-CellValue {
    <# Removes t/f/v/is so the cell reads as genuinely blank - used for the CA row's
       placeholder dimension cells and to strip any leftover template formula. #>
    param($CellEl)
    if (-not $CellEl) { return }
    $CellEl.RemoveAttribute('t')
    foreach ($child in @($CellEl.ChildNodes)) { [void]$CellEl.RemoveChild($child) }
}

function Set-CellStringValue {
    param([Xml.XmlDocument]$Doc, $CellEl, [Xml.XmlDocument]$SstDoc, [string]$Text)
    Clear-CellValue $CellEl
    $CellEl.SetAttribute('t', 's')
    $sIdx = Add-SharedString $SstDoc $Text
    $v = $Doc.CreateElement('v', $NS_MAIN)
    $v.InnerText = [string]$sIdx
    [void]$CellEl.AppendChild($v)
}

function Set-CellNumberValue {
    param([Xml.XmlDocument]$Doc, $CellEl, [double]$Number)
    Clear-CellValue $CellEl
    $v = $Doc.CreateElement('v', $NS_MAIN)
    $v.InnerText = $Number.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
    [void]$CellEl.AppendChild($v)
}

function Set-CellFormulaValue {
    param([Xml.XmlDocument]$Doc, $CellEl, [string]$Formula, [double]$CachedValue)
    Clear-CellValue $CellEl
    $f = $Doc.CreateElement('f', $NS_MAIN)
    $f.InnerText = $Formula
    [void]$CellEl.AppendChild($f)
    $v = $Doc.CreateElement('v', $NS_MAIN)
    $v.InnerText = $CachedValue.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
    [void]$CellEl.AppendChild($v)
}

function Get-OrCreateRow {
    param([Xml.XmlDocument]$Doc, $SheetDataEl, [hashtable]$RowsByNumber, [int]$RowNum, [int]$SpanCols)
    if ($RowsByNumber.ContainsKey($RowNum)) { return $RowsByNumber[$RowNum] }
    $r = $Doc.CreateElement('row', $NS_MAIN)
    $r.SetAttribute('r', [string]$RowNum)
    $r.SetAttribute('spans', "1:$SpanCols")
    $inserted = $false
    foreach ($existing in @($SheetDataEl.ChildNodes)) {
        if ($existing.LocalName -ne 'row') { continue }
        if ([int]$existing.GetAttribute('r') -gt $RowNum) {
            [void]$SheetDataEl.InsertBefore($r, $existing)
            $inserted = $true
            break
        }
    }
    if (-not $inserted) { [void]$SheetDataEl.AppendChild($r) }
    $RowsByNumber[$RowNum] = $r
    return $r
}

function Resolve-RelativePart {
    <# OPC relationship Targets are relative to the directory of the part that OWNS the
       .rels file, not to the package root - "../tables/table2.xml" from a relationship
       declared for xl/worksheets/sheet6.xml resolves to xl/tables/table2.xml, not
       xl/../tables/table2.xml. Uri combining resolves ".." correctly; naive string
       prepending does not. #>
    param([string]$OwningPart, [string]$RelativeTarget)
    if ($RelativeTarget.StartsWith('/')) { return $RelativeTarget.TrimStart('/') }
    $baseUri = New-Object Uri("file:///$OwningPart")
    $resolved = New-Object Uri($baseUri, $RelativeTarget)
    return [Uri]::UnescapeDataString($resolved.AbsolutePath).TrimStart('/')
}

function Resolve-SheetPart {
    <# Sheet name -> zip part path, resolved through workbook.xml + its .rels every run.
       Never hardcode "sheetN.xml" - the mapping is workbook-specific and can change. #>
    param($Zip, [string]$SheetName)
    $wbDoc = Read-ZipXml $Zip 'xl/workbook.xml'
    $nsmgr = New-Object Xml.XmlNamespaceManager($wbDoc.NameTable)
    $nsmgr.AddNamespace('m', $NS_MAIN); $nsmgr.AddNamespace('r', $NS_R)
    $sheetNode = $wbDoc.SelectSingleNode("//m:sheets/m:sheet[@name='$SheetName']", $nsmgr)
    if (-not $sheetNode) { throw "Sheet '$SheetName' was not found in workbook.xml." }
    $rId = $sheetNode.GetAttribute('id', $NS_R)

    $relsDoc = Read-ZipXml $Zip 'xl/_rels/workbook.xml.rels'
    $nsmgr2 = New-Object Xml.XmlNamespaceManager($relsDoc.NameTable)
    $nsmgr2.AddNamespace('p', $NS_RELS)
    $relNode = $relsDoc.SelectSingleNode("//p:Relationship[@Id='$rId']", $nsmgr2)
    if (-not $relNode) { throw "Relationship '$rId' for sheet '$SheetName' was not found." }
    return Resolve-RelativePart 'xl/workbook.xml' $relNode.GetAttribute('Target')
}

function Remove-CalcChain {
    <# Deleting calcChain.xml after adding/changing formula cells is the standard,
       safe move every OOXML writer makes - Excel silently rebuilds it. Leaving a
       stale one behind, or a dangling reference to a deleted part, is what actually
       risks a "we found a problem with some content" repair prompt. #>
    param($Zip)
    if (-not (Test-ZipEntry $Zip 'xl/calcChain.xml')) { return }
    Remove-ZipEntryIfPresent $Zip 'xl/calcChain.xml'

    $ctDoc = Read-ZipXml $Zip '[Content_Types].xml'
    $nsmgr = New-Object Xml.XmlNamespaceManager($ctDoc.NameTable)
    $nsmgr.AddNamespace('ct', $NS_CT)
    $override = $ctDoc.SelectSingleNode("//ct:Override[@PartName='/xl/calcChain.xml']", $nsmgr)
    if ($override) { [void]$override.ParentNode.RemoveChild($override) }
    Write-ZipXml $Zip '[Content_Types].xml' $ctDoc

    $relsDoc = Read-ZipXml $Zip 'xl/_rels/workbook.xml.rels'
    $nsmgr2 = New-Object Xml.XmlNamespaceManager($relsDoc.NameTable)
    $nsmgr2.AddNamespace('p', $NS_RELS)
    $relNode = $relsDoc.SelectSingleNode("//p:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain']", $nsmgr2)
    if ($relNode) { [void]$relNode.ParentNode.RemoveChild($relNode) }
    Write-ZipXml $Zip 'xl/_rels/workbook.xml.rels' $relsDoc
}

function Set-FullCalcOnLoad {
    <# Belt-and-suspenders alongside the cached <v> values this script always writes:
       forces Excel to recompute every formula on next open regardless. #>
    param($Zip)
    $wbDoc = Read-ZipXml $Zip 'xl/workbook.xml'
    $nsmgr = New-Object Xml.XmlNamespaceManager($wbDoc.NameTable)
    $nsmgr.AddNamespace('m', $NS_MAIN)
    $calcPr = $wbDoc.SelectSingleNode('//m:calcPr', $nsmgr)
    if (-not $calcPr) {
        $calcPr = $wbDoc.CreateElement('calcPr', $NS_MAIN)
        [void]$wbDoc.DocumentElement.AppendChild($calcPr)
    }
    $calcPr.SetAttribute('fullCalcOnLoad', '1')
    Write-ZipXml $Zip 'xl/workbook.xml' $wbDoc
}

function Get-HeaderMap {
    <# First-occurrence-wins by header text, exactly like the previous COM version's
       Resolve-Col - this sheet has duplicate/blank headers so order matters. #>
    param($SheetDataEl, [Xml.XmlDocument]$SstDoc)
    $headerRow = $null
    foreach ($r in $SheetDataEl.ChildNodes) {
        if ($r.LocalName -eq 'row' -and $r.GetAttribute('r') -eq '1') { $headerRow = $r; break }
    }
    if (-not $headerRow) { throw 'Header row (row 1) was not found.' }
    $map = @{}
    foreach ($c in $headerRow.ChildNodes) {
        if ($c.LocalName -ne 'c') { continue }
        $text = (Resolve-CellText $c $SstDoc).Trim()
        if ($text -and -not $map.ContainsKey($text)) { $map[$text] = Get-ColIndexFromRef $c.GetAttribute('r') }
    }
    return $map
}

function Resolve-Col {
    param([hashtable]$HeaderMap, [string]$HeaderName, [string]$SheetName)
    if (-not $HeaderMap.ContainsKey($HeaderName)) { throw "Column '$HeaderName' was not found in sheet '$SheetName'." }
    return $HeaderMap[$HeaderName]
}

# ---------------------------------------------------------------------------------------------
# 1. Resolve the target workbook - never hardcoded, per the current session's dated filename.
# ---------------------------------------------------------------------------------------------
$targetPath = $null
$ini = Read-IniFile -FilePath $CacheStatePath
if ($ini.ContainsKey('DimensionWorkbook') -and $ini['DimensionWorkbook'].ContainsKey('Path')) {
    $p = $ini['DimensionWorkbook']['Path']
    if ($p -and (Test-Path -LiteralPath $p)) { $targetPath = $p }
}
if (-not $targetPath -and (Test-Path -LiteralPath $WeeklyBomDir)) {
    $candidate = Get-ChildItem -LiteralPath $WeeklyBomDir -Filter 'Products Physical Dimension - FG *.xlsx' -File |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($candidate) { $targetPath = $candidate.FullName }
}
if (-not $targetPath) {
    Write-WbLog 'FAILED: could not resolve the dimension workbook path.'
    Send-Result -Ok $false -Fields @{ error = 'Could not find the dimension workbook. Checked cache_state.ini and the 1. Weekly BOM folder.' }
    exit 1
}

# ---------------------------------------------------------------------------------------------
# 2. Fail-fast lock check - a raw exclusive-open probe, no retry loop. Aborts in under a
#    second if Excel or OneDrive has the file open, before any backup/staging copy runs.
# ---------------------------------------------------------------------------------------------
$lockSw = [Diagnostics.Stopwatch]::StartNew()
try {
    $probe = [IO.File]::Open($targetPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $probe.Close(); $probe.Dispose()
}
catch {
    $lockSw.Stop()
    Write-WbLog ("FAILED: workbook is locked (checked in {0}ms): {1}" -f $lockSw.ElapsedMilliseconds, $_.Exception.Message)
    Send-Result -Ok $false -Fields @{
        error    = 'The workbook is open elsewhere (Excel or OneDrive sync). Close it and try again.'
        locked   = $true
        check_ms = $lockSw.ElapsedMilliseconds
    }
    exit 1
}
Write-WbLog ("Lock check passed in {0}ms." -f $lockSw.ElapsedMilliseconds)

# ---------------------------------------------------------------------------------------------
# 3. Backup before any write.
# ---------------------------------------------------------------------------------------------
$backupDir = Join-Path $WeeklyBomDir '_backups'
if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
$stamp = [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')
$backupName = '{0}.{1}.bak.xlsx' -f ([IO.Path]::GetFileNameWithoutExtension($targetPath)), $stamp
$backupPath = Join-Path $backupDir $backupName
try {
    Copy-Item -LiteralPath $targetPath -Destination $backupPath -Force
    Write-WbLog "Backup created: $backupPath"
}
catch {
    Write-WbLog "FAILED: backup copy failed: $($_.Exception.Message)"
    Send-Result -Ok $false -Fields @{ error = "Could not create a backup copy before writing, so nothing was changed. ($($_.Exception.Message))" }
    exit 1
}

# ---------------------------------------------------------------------------------------------
# 4. Stage a working copy and edit it as a zip archive. The original is only ever touched
#    once, at the very end, via an atomic copy of the fully-edited staged file.
# ---------------------------------------------------------------------------------------------
$stagedPath = Join-Path $StagingDir ("staged_dimension_update.$TrackId.xlsx")
Copy-Item -LiteralPath $targetPath -Destination $stagedPath -Force

$writeSw = [Diagnostics.Stopwatch]::StartNew()
$zip = $null
$row = 0
$caRow = 0
$plRow = 0
$isNew = $false
$sheetName = $null
try {
    $zip = [IO.Compression.ZipFile]::Open($stagedPath, [IO.Compression.ZipArchiveMode]::Update)

    $sheetName = if ($Type -eq 'product') { 'Dimension' } else { 'Mastercarton_Dimensions' }
    $sheetPart = Resolve-SheetPart $zip $sheetName
    $sheetDoc = Read-ZipXml $zip $sheetPart
    $sstDoc = Read-ZipXml $zip 'xl/sharedStrings.xml'
    $nsmgr = New-Object Xml.XmlNamespaceManager($sheetDoc.NameTable)
    $nsmgr.AddNamespace('m', $NS_MAIN)
    $sheetDataEl = $sheetDoc.SelectSingleNode('//m:sheetData', $nsmgr)
    if (-not $sheetDataEl) { throw "sheetData element not found in $sheetPart." }

    $headerMap = Get-HeaderMap $sheetDataEl $sstDoc
    $formulaWritten = $false

    if ($Type -eq 'product') {
        $idCol = Resolve-Col $headerMap 'Item number' $sheetName
        $unitCol = Resolve-Col $headerMap 'Unit' $sheetName
        $widthCol = Resolve-Col $headerMap 'Width' $sheetName
        $depthCol = Resolve-Col $headerMap 'Depth' $sheetName
        $heightCol = Resolve-Col $headerMap 'Height' $sheetName
        $volCol = 0
        if ($headerMap.ContainsKey('Volume')) { $volCol = $headerMap['Volume'] }
        # First 'Gross weight' header wins (same rule as every other column on this sheet), which is the kg
        # column beside Unit - the later 'Gross weight' headers belong to the pallet helper block.
        $weightCol = 0
        if ($GrossWeightKg -gt 0) { $weightCol = Resolve-Col $headerMap 'Gross weight' $sheetName }
        $eaWeightKg = [Math]::Round($GrossWeightKg, 3)
        $idLetter = Get-ColumnLetter $idCol
        $unitLetter = Get-ColumnLetter $unitCol
        $maxCol = ($headerMap.Values | Measure-Object -Maximum).Maximum

        # Single pass over every existing row: find the true last data row (a row is only
        # "used" if its Item-number cell has an actual value - the sheet's array-formula
        # template rows extend far beyond real data and must not be mistaken for it) and,
        # in the same pass, look for an existing EA row for this exact ID.
        $rowsByNumber = @{}
        $lastDataRow = 0
        $matchRow = 0
        foreach ($rowEl in $sheetDataEl.ChildNodes) {
            if ($rowEl.LocalName -ne 'row') { continue }
            $rn = [int]$rowEl.GetAttribute('r')
            $rowsByNumber[$rn] = $rowEl
            $idCell = $null; $unitCell = $null
            $idRef = "$idLetter$rn"; $unitRef = "$unitLetter$rn"
            foreach ($c in $rowEl.ChildNodes) {
                if ($c.LocalName -ne 'c') { continue }
                $ref = $c.GetAttribute('r')
                if ($ref -eq $idRef) { $idCell = $c }
                elseif ($ref -eq $unitRef) { $unitCell = $c }
            }
            $idText = (Resolve-CellText $idCell $sstDoc).Trim()
            if ($idText) {
                if ($rn -gt $lastDataRow) { $lastDataRow = $rn }
                if ($idText -eq $Id.Trim()) {
                    $unitText = (Resolve-CellText $unitCell $sstDoc).Trim().ToUpperInvariant()
                    if ($unitText -eq 'EA') { $matchRow = $rn }
                }
            }
        }
        if ($lastDataRow -lt 1) { $lastDataRow = 1 }
        $row = $matchRow

        $eaWidthCm = [Math]::Round($WidthMm / 10.0, 3)
        $eaDepthCm = [Math]::Round($LengthMm / 10.0, 3)
        $eaHeightCm = [Math]::Round($HeightMm / 10.0, 3)

        if ($row -eq 0) {
            # New FG: write a full CA/EA/PL block, matching every other item's layout on
            # this sheet. CA and PL keep their placeholder shape - only EA is fully real.
            $isNew = $true
            $caRow = $lastDataRow + 1
            $row = $lastDataRow + 2
            $plRow = $lastDataRow + 3

            $caRowEl = Get-OrCreateRow $sheetDoc $sheetDataEl $rowsByNumber $caRow $maxCol
            $eaRowEl = Get-OrCreateRow $sheetDoc $sheetDataEl $rowsByNumber $row $maxCol
            $plRowEl = Get-OrCreateRow $sheetDoc $sheetDataEl $rowsByNumber $plRow $maxCol

            $caId = Get-OrCreateCell $sheetDoc $caRowEl "$idLetter$caRow" $null
            Set-CellStringValue $sheetDoc $caId $sstDoc $Id
            $caUnit = Get-OrCreateCell $sheetDoc $caRowEl "$unitLetter$caRow" $null
            Set-CellStringValue $sheetDoc $caUnit $sstDoc 'CA'
            foreach ($col in @($widthCol, $depthCol, $heightCol)) {
                Clear-CellValue (Get-OrCreateCell $sheetDoc $caRowEl "$(Get-ColumnLetter $col)$caRow" $null)
            }
            if ($volCol) { Clear-CellValue (Get-OrCreateCell $sheetDoc $caRowEl "$(Get-ColumnLetter $volCol)$caRow" $null) }

            $eaId = Get-OrCreateCell $sheetDoc $eaRowEl "$idLetter$row" $null
            Set-CellStringValue $sheetDoc $eaId $sstDoc $Id
            $eaUnit = Get-OrCreateCell $sheetDoc $eaRowEl "$unitLetter$row" $null
            Set-CellStringValue $sheetDoc $eaUnit $sstDoc 'EA'
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $widthCol)$row" $null) $eaWidthCm
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $depthCol)$row" $null) $eaDepthCm
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $heightCol)$row" $null) $eaHeightCm
            if ($weightCol) { Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $weightCol)$row" $null) $eaWeightKg }
            if ($volCol) {
                $volLetter = Get-ColumnLetter $volCol
                $wL = Get-ColumnLetter $widthCol; $dL = Get-ColumnLetter $depthCol; $hL = Get-ColumnLetter $heightCol
                $volFormula = "$wL$row*$dL$row*$hL$row"
                $volCached = $eaWidthCm * $eaDepthCm * $eaHeightCm
                Set-CellFormulaValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$volLetter$row" $null) $volFormula $volCached
                $formulaWritten = $true
            }

            $plId = Get-OrCreateCell $sheetDoc $plRowEl "$idLetter$plRow" $null
            Set-CellStringValue $sheetDoc $plId $sstDoc $Id
            $plUnit = Get-OrCreateCell $sheetDoc $plRowEl "$unitLetter$plRow" $null
            Set-CellStringValue $sheetDoc $plUnit $sstDoc 'PL'
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $plRowEl "$(Get-ColumnLetter $widthCol)$plRow" $null) 120
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $plRowEl "$(Get-ColumnLetter $depthCol)$plRow" $null) 100
            # Height has no default per spec - left blank, so Volume is left blank too rather
            # than showing a misleading zero.
            Clear-CellValue (Get-OrCreateCell $sheetDoc $plRowEl "$(Get-ColumnLetter $heightCol)$plRow" $null)
            if ($volCol) { Clear-CellValue (Get-OrCreateCell $sheetDoc $plRowEl "$(Get-ColumnLetter $volCol)$plRow" $null) }
        }
        else {
            # FG already has an EA row: update dimensions (and Volume) in place only.
            # CA and PL, wherever they sit, are never touched.
            $eaRowEl = $rowsByNumber[$row]
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $widthCol)$row" $null) $eaWidthCm
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $depthCol)$row" $null) $eaDepthCm
            Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $heightCol)$row" $null) $eaHeightCm
            if ($weightCol) { Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$(Get-ColumnLetter $weightCol)$row" $null) $eaWeightKg }
            if ($volCol) {
                $volLetter = Get-ColumnLetter $volCol
                $wL = Get-ColumnLetter $widthCol; $dL = Get-ColumnLetter $depthCol; $hL = Get-ColumnLetter $heightCol
                $volFormula = "$wL$row*$dL$row*$hL$row"
                $volCached = $eaWidthCm * $eaDepthCm * $eaHeightCm
                Set-CellFormulaValue $sheetDoc (Get-OrCreateCell $sheetDoc $eaRowEl "$volLetter$row" $null) $volFormula $volCached
                $formulaWritten = $true
            }
        }

        $newDimRef = "A1:$(Get-ColumnLetter $maxCol)$([Math]::Max($plRow, $lastDataRow))"
        $dimEl = $sheetDoc.SelectSingleNode('//m:dimension', $nsmgr)
        if ($dimEl -and $isNew) { $dimEl.SetAttribute('ref', $newDimRef) }
    }
    else {
        $idCol = Resolve-Col $headerMap 'carton_id' $sheetName
        $nameCol = Resolve-Col $headerMap 'carton_name' $sheetName
        $widthCol = Resolve-Col $headerMap 'carton_width_cm' $sheetName
        $depthCol = Resolve-Col $headerMap 'carton_depth_cm' $sheetName
        $heightCol = Resolve-Col $headerMap 'carton_height_cm' $sheetName
        $volCol = Resolve-Col $headerMap 'carton_volume_mm3' $sheetName
        $lifecycleCol = Resolve-Col $headerMap 'carton_lifecycle' $sheetName
        $idLetter = Get-ColumnLetter $idCol
        $maxCol = ($headerMap.Values | Measure-Object -Maximum).Maximum

        $rowsByNumber = @{}
        $lastDataRow = 0
        $matchRow = 0
        foreach ($rowEl in $sheetDataEl.ChildNodes) {
            if ($rowEl.LocalName -ne 'row') { continue }
            $rn = [int]$rowEl.GetAttribute('r')
            $rowsByNumber[$rn] = $rowEl
            $idCell = $null
            $idRef = "$idLetter$rn"
            foreach ($c in $rowEl.ChildNodes) {
                if ($c.LocalName -eq 'c' -and $c.GetAttribute('r') -eq $idRef) { $idCell = $c; break }
            }
            $idText = (Resolve-CellText $idCell $sstDoc).Trim()
            if ($idText) {
                if ($rn -gt $lastDataRow) { $lastDataRow = $rn }
                if ($idText -eq $Id.Trim()) { $matchRow = $rn }
            }
        }
        if ($lastDataRow -lt 1) { $lastDataRow = 1 }
        $row = $matchRow
        if ($row -eq 0) { $isNew = $true; $row = $lastDataRow + 1 }

        $wCm = [Math]::Round($WidthMm / 10.0, 3)
        $dCm = [Math]::Round($LengthMm / 10.0, 3)
        $hCm = [Math]::Round($HeightMm / 10.0, 3)

        $rowEl = Get-OrCreateRow $sheetDoc $sheetDataEl $rowsByNumber $row $maxCol
        Set-CellStringValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$idLetter$row" $null) $sstDoc $Id
        if ($Name) { Set-CellStringValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $nameCol)$row" $null) $sstDoc $Name }
        Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $widthCol)$row" $null) $wCm
        Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $depthCol)$row" $null) $dCm
        Set-CellNumberValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $heightCol)$row" $null) $hCm

        $wL = Get-ColumnLetter $widthCol; $dL = Get-ColumnLetter $depthCol; $hL = Get-ColumnLetter $heightCol
        $volFormula = "($wL$row*10)*($dL$row*10)*($hL$row*10)"
        $volCached = ($wCm * 10) * ($dCm * 10) * ($hCm * 10)
        Set-CellFormulaValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $volCol)$row" $null) $volFormula $volCached
        $formulaWritten = $true

        Set-CellStringValue $sheetDoc (Get-OrCreateCell $sheetDoc $rowEl "$(Get-ColumnLetter $lifecycleCol)$row" $null) $sstDoc 'Active'

        if ($isNew) {
            $dimEl = $sheetDoc.SelectSingleNode('//m:dimension', $nsmgr)
            if ($dimEl) { $dimEl.SetAttribute('ref', "A1:$(Get-ColumnLetter $maxCol)$row") }

            # Extend the Table's own ref range too, or Excel will flag the file for repair
            # on open. Only needed when the new row falls outside the table's current range.
            $sheetRelsPart = ($sheetPart -replace '(.*/)([^/]+)$', '$1_rels/$2') + '.rels'
            if (Test-ZipEntry $zip $sheetRelsPart) {
                $sheetRelsDoc = Read-ZipXml $zip $sheetRelsPart
                $nsmgr3 = New-Object Xml.XmlNamespaceManager($sheetRelsDoc.NameTable)
                $nsmgr3.AddNamespace('p', $NS_RELS)
                $tableRel = $sheetRelsDoc.SelectSingleNode("//p:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/table']", $nsmgr3)
                if ($tableRel) {
                    $tablePart = Resolve-RelativePart $sheetPart $tableRel.GetAttribute('Target')
                    $tableDoc = Read-ZipXml $zip $tablePart
                    $tableEl = $tableDoc.DocumentElement
                    $currentRef = $tableEl.GetAttribute('ref')
                    if ($currentRef -match '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') {
                        $refStartCol = $matches[1]; $refStartRow = [int]$matches[2]; $refEndCol = $matches[3]; $refEndRow = [int]$matches[4]
                        if ($row -gt $refEndRow) {
                            $tableEl.SetAttribute('ref', "${refStartCol}${refStartRow}:${refEndCol}${row}")
                            Write-ZipXml $zip $tablePart $tableDoc
                            Write-WbLog "Extended table ref to ${refStartCol}${refStartRow}:${refEndCol}${row}"
                        }
                    }
                }
            }
        }
    }

    Write-ZipXml $zip $sheetPart $sheetDoc
    Write-ZipXml $zip 'xl/sharedStrings.xml' $sstDoc

    if ($formulaWritten) {
        Remove-CalcChain $zip
        Set-FullCalcOnLoad $zip
    }

    $zip.Dispose()
    $zip = $null
    $writeSw.Stop()
    Write-WbLog ("Staged {0} row {1} (id={2}, isNew={3}) in sheet '{4}' - {5}ms" -f $Type, $row, $Id, $isNew, $sheetName, $writeSw.ElapsedMilliseconds)
}
catch {
    $errMsg = $_.Exception.Message
    Write-WbLog ("FAILED: {0} (line {1}: {2})" -f $errMsg, $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim())
    if ($zip) { try { $zip.Dispose() } catch { } }
    try { Remove-Item -LiteralPath $stagedPath -Force -ErrorAction SilentlyContinue } catch { }
    Send-Result -Ok $false -Fields @{ error = $errMsg; backup_path = $backupPath }
    exit 1
}

# ---------------------------------------------------------------------------------------------
# 5. Atomic swap - the original is only ever touched here, with a single file copy of the
#    fully-edited staged file. If anything above failed, this line never runs.
# ---------------------------------------------------------------------------------------------
try {
    [IO.File]::Copy($stagedPath, $targetPath, $true)
    Write-WbLog "Atomic swap complete: $targetPath"
}
catch {
    Write-WbLog "FAILED: atomic swap failed: $($_.Exception.Message)"
    Send-Result -Ok $false -Fields @{ error = "The edit succeeded but could not be swapped into place: $($_.Exception.Message)"; backup_path = $backupPath }
    exit 1
}
finally {
    try { Remove-Item -LiteralPath $stagedPath -Force -ErrorAction SilentlyContinue } catch { }
}

# ---------------------------------------------------------------------------------------------
# 6. Headless cache rebuild - fire-and-forget, not awaited. Even at Zip/XML speed the CSV
#    rebuild itself still re-scans the ~90k-row BOM workbook (~40s), so this stays async.
# ---------------------------------------------------------------------------------------------
$cacheRebuildStarted = $false
$rebuildError = ''
try {
    $launcherPath = Join-Path $SystemDir 'launcher.ps1'
    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$launcherPath`"", '--manifests-only'
    ) | Out-Null
    $cacheRebuildStarted = $true
    Write-WbLog 'Cache rebuild (--manifests-only) started in the background.'
}
catch {
    $rebuildError = $_.Exception.Message
    Write-WbLog "Could not start the cache rebuild: $rebuildError"
}

Send-Result -Ok $true -Fields @{
    type                  = $Type
    id                    = $Id
    row                   = $row
    ca_row                = $caRow
    pl_row                = $plRow
    is_new                = $isNew
    sheet                 = $sheetName
    backup_path           = $backupPath
    write_ms              = $writeSw.ElapsedMilliseconds
    gross_weight_kg       = $(if ($Type -eq 'product' -and $GrossWeightKg -gt 0) { [Math]::Round($GrossWeightKg, 3) } else { $null })
    cache_rebuild_started = $cacheRebuildStarted
    rebuild_error         = $rebuildError
}
