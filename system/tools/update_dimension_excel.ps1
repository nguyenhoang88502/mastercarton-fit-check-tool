<#
.SYNOPSIS
    Short-lived writer that saves one product's or carton's dimensions into the live
    "Products Physical Dimension - FG" workbook via Excel COM automation, then triggers
    a headless cache rebuild so the tool's CSVs reflect the change immediately.

.DESCRIPTION
    Spawned fresh per save, exits as soon as it is done, and never stays resident - the
    same short-lived model the http-listener already uses for CSVs. Every write is
    preceded by a timestamped backup copy of the workbook, and the Excel instance is
    always closed and released in a finally block so a mid-write failure cannot leave
    Excel running or the file locked.

    Column positions are resolved by header text on each run rather than hardcoded, so a
    reordered column does not silently write into the wrong field.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File update_dimension_excel.ps1 -Type product -Id NPI-2026-001 -LengthMm 225 -WidthMm 130 -HeightMm 55
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
    [Parameter(Mandatory = $true)][double]$HeightMm
)

$ErrorActionPreference = 'Stop'
$xlUp = -4162  # Excel.Constants.xlUp - avoids a COM interop reference just for one enum.

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

function Write-WbLog {
    param([string]$Message)
    $line = '{0} {1}' -f ([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')), $Message
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

function Get-ColumnValues {
    <# One COM call for the whole column instead of one per row - cell-by-cell .Text reads
       over a couple thousand rows measured ~30s on this workbook from marshaling overhead
       alone; a single Range.Value2 bulk read is near-instant. Returns a 0-indexed array
       aligned to row (StartRow + index). Value2 comes back as a scalar for a single-cell
       range and a 2D array for anything larger, so both shapes are normalized here. #>
    param($Worksheet, [int]$StartRow, [int]$EndRow, [int]$Col)
    if ($EndRow -lt $StartRow) { return @() }
    $range = $Worksheet.Range($Worksheet.Cells.Item($StartRow, $Col), $Worksheet.Cells.Item($EndRow, $Col))
    $raw = $range.Value2
    if ($EndRow -eq $StartRow) { return , $raw }
    $count = $EndRow - $StartRow + 1
    $result = New-Object object[] $count
    for ($i = 1; $i -le $count; $i++) { $result[$i - 1] = $raw[$i, 1] }
    return $result
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

function Send-Result {
    <# The listener parses the LAST stdout line as JSON, so this must be the only thing
       this script ever writes to the output stream. #>
    param([bool]$Ok, [hashtable]$Fields)
    $obj = [ordered]@{ ok = $Ok }
    foreach ($k in $Fields.Keys) { $obj[$k] = $Fields[$k] }
    Write-Output ($obj | ConvertTo-Json -Compress -Depth 6)
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
    Write-WbLog 'FAILED: could not resolve the dimension workbook path (cache_state.ini and 1. Weekly BOM folder scan both came up empty).'
    Send-Result -Ok $false -Fields @{ error = 'Could not find the dimension workbook. Checked cache_state.ini and the 1. Weekly BOM folder.' }
    exit 0
}
Write-WbLog "Resolved workbook: $targetPath"

# ---------------------------------------------------------------------------------------------
# 2. Backup before any write - this is the recovery path if anything below goes wrong.
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
    exit 0
}

# ---------------------------------------------------------------------------------------------
# 3. Excel COM write - always closed/released in the catch too, so a failure never leaves
#    an orphaned Excel process or a locked file behind.
# ---------------------------------------------------------------------------------------------
$excel = $null
$wb = $null
$row = 0
$isNew = $false
$sheetName = $null
try {
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Excel is not available on this machine for COM automation: $($_.Exception.Message)"
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    try {
        $wb = $excel.Workbooks.Open($targetPath, [Type]::Missing, $false)
    }
    catch {
        throw "Could not open the workbook for writing - it may already be open elsewhere. ($($_.Exception.Message))"
    }

    if ($wb.ReadOnly) {
        throw 'The workbook opened read-only, which means it is already open elsewhere. Close it there and try again.'
    }

    $sheetName = if ($Type -eq 'product') { 'Dimension' } else { 'Mastercarton_Dimensions' }
    $ws = $null
    foreach ($sheet in $wb.Sheets) { if ($sheet.Name -eq $sheetName) { $ws = $sheet; break } }
    if (-not $ws) { throw "Sheet '$sheetName' was not found in the workbook." }

    # Header-name -> column-index map read fresh from row 1 each run, so a column that gets
    # reordered later still resolves to the right place instead of silently writing wrong data.
    $headerLastCol = $ws.UsedRange.Columns.Count
    $headerMap = @{}
    for ($c = 1; $c -le $headerLastCol; $c++) {
        $h = [string]$ws.Cells.Item(1, $c).Text
        if (-not [string]::IsNullOrWhiteSpace($h) -and -not $headerMap.ContainsKey($h.Trim())) { $headerMap[$h.Trim()] = $c }
    }
    function Resolve-Col {
        param([string]$HeaderName)
        if (-not $headerMap.ContainsKey($HeaderName)) { throw "Column '$HeaderName' was not found in sheet '$sheetName'." }
        return $headerMap[$HeaderName]
    }

    $table = $null
    if ($ws.ListObjects.Count -ge 1) { $table = $ws.ListObjects.Item(1) }

    if ($Type -eq 'product') {
        $idCol = Resolve-Col 'Item number'
        $unitCol = Resolve-Col 'Unit'
        $widthCol = Resolve-Col 'Width'
        $depthCol = Resolve-Col 'Depth'
        $heightCol = Resolve-Col 'Height'

        # UsedRange over-counts on this sheet (trailing formatted-but-empty rows), so the real
        # last row comes from walking up the id column instead, same as Ctrl+Up in Excel.
        $lastDataRow = $ws.Cells.Item($ws.Rows.Count, $idCol).End($xlUp).Row
        if ($lastDataRow -lt 1) { $lastDataRow = 1 }

        $idValues = Get-ColumnValues $ws 2 $lastDataRow $idCol
        $unitValues = Get-ColumnValues $ws 2 $lastDataRow $unitCol
        for ($i = 0; $i -lt $idValues.Count; $i++) {
            $rowId = ([string]$idValues[$i]).Trim()
            $rowUnit = ([string]$unitValues[$i]).Trim().ToUpperInvariant()
            if ($rowId -eq $Id.Trim() -and $rowUnit -eq 'EA') { $row = 2 + $i; break }
        }

        if ($row -eq 0) {
            $isNew = $true
            if ($table) { $row = ($table.ListRows.Add()).Range.Row } else { $row = $lastDataRow + 1 }
        }

        # Plain default-property assignment, not .Value2 - alternating string/double writes
        # through .Value2 on a late-bound COM Range hits a known PowerShell binder quirk that
        # mis-casts the second type ("Unable to cast ... Double ... to ... String"), confirmed
        # against this exact workbook. The default indexer does not have that problem.
        $ws.Cells.Item($row, $idCol) = $Id
        $ws.Cells.Item($row, $unitCol) = 'EA'
        $ws.Cells.Item($row, $widthCol) = [Math]::Round($WidthMm / 10.0, 3)
        $ws.Cells.Item($row, $depthCol) = [Math]::Round($LengthMm / 10.0, 3)
        $ws.Cells.Item($row, $heightCol) = [Math]::Round($HeightMm / 10.0, 3)
    }
    else {
        $idCol = Resolve-Col 'carton_id'
        $nameCol = Resolve-Col 'carton_name'
        $widthCol = Resolve-Col 'carton_width_cm'
        $depthCol = Resolve-Col 'carton_depth_cm'
        $heightCol = Resolve-Col 'carton_height_cm'
        $volCol = Resolve-Col 'carton_volume_mm3'
        $lifecycleCol = Resolve-Col 'carton_lifecycle'

        $lastDataRow = $ws.Cells.Item($ws.Rows.Count, $idCol).End($xlUp).Row
        if ($lastDataRow -lt 1) { $lastDataRow = 1 }

        $idValues = Get-ColumnValues $ws 2 $lastDataRow $idCol
        for ($i = 0; $i -lt $idValues.Count; $i++) {
            $rowId = ([string]$idValues[$i]).Trim()
            if ($rowId -eq $Id.Trim()) { $row = 2 + $i; break }
        }

        if ($row -eq 0) {
            $isNew = $true
            # ListRows.Add() extends the table's own Ref range along with it, so a plain
            # "write past the last row" never leaves this carton sitting outside the table.
            if ($table) { $row = ($table.ListRows.Add()).Range.Row } else { $row = $lastDataRow + 1 }
        }

        $ws.Cells.Item($row, $idCol) = $Id
        if ($Name) { $ws.Cells.Item($row, $nameCol) = $Name }
        $ws.Cells.Item($row, $widthCol) = [Math]::Round($WidthMm / 10.0, 3)
        $ws.Cells.Item($row, $depthCol) = [Math]::Round($LengthMm / 10.0, 3)
        $ws.Cells.Item($row, $heightCol) = [Math]::Round($HeightMm / 10.0, 3)

        $wLetter = Get-ColumnLetter $widthCol
        $dLetter = Get-ColumnLetter $depthCol
        $hLetter = Get-ColumnLetter $heightCol
        $ws.Cells.Item($row, $volCol).Formula = "=($wLetter$row*10)*($dLetter$row*10)*($hLetter$row*10)"
        $ws.Cells.Item($row, $lifecycleCol) = 'Active'
    }

    $wb.Save()
    Write-WbLog ("Saved {0} row {1} (id={2}, isNew={3}) in sheet '{4}'" -f $Type, $row, $Id, $isNew, $sheetName)
}
catch {
    $errMsg = $_.Exception.Message
    Write-WbLog ("FAILED: {0} (line {1}: {2})" -f $errMsg, $_.InvocationInfo.ScriptLineNumber, $_.InvocationInfo.Line.Trim())
    try { if ($wb) { $wb.Close($false) } } catch { }
    try { if ($excel) { $excel.Quit() } } catch { }
    if ($wb) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { } }
    if ($excel) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    Send-Result -Ok $false -Fields @{ error = $errMsg; backup_path = $backupPath }
    exit 0
}

try { $wb.Close($false) } catch { }
try { $excel.Quit() } catch { }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()

# ---------------------------------------------------------------------------------------------
# 4. Headless cache rebuild - --manifests-only rebuilds the CSVs and skips the browser/listener
#    launch, so this never opens a second browser window or a second listener instance.
#
#    Fired in the background, NOT awaited: a full rebuild re-scans the ~90k-row BOM workbook
#    and measured ~40s on this machine. Waiting on it here would make every single Save click
#    hang for 40s. The Excel write above is already durable by the time this script returns;
#    the CSVs (and the tool's in-browser view) simply catch up a little after.
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
    type                   = $Type
    id                     = $Id
    row                    = $row
    is_new                 = $isNew
    sheet                  = $sheetName
    backup_path            = $backupPath
    cache_rebuild_started  = $cacheRebuildStarted
    rebuild_error          = $rebuildError
}
