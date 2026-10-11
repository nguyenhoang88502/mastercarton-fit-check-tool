<#
.SYNOPSIS
    Writes every save that is still waiting (pending_workbook_writes.csv) into the dimension
    workbook, one item at a time, through update_dimension_excel.ps1.

.DESCRIPTION
    Saving a product or carton in the tool is instant: the entry goes into the helper's CSV
    files and is queued here. The launcher runs this script right before it reads the
    workbooks, so the queued entries are in the workbook by the time it is parsed.

    An entry leaves the queue only when its write succeeded. If the workbook is open elsewhere
    the run stops at the first locked write and everything left stays queued for next time -
    nothing is lost and nothing is retried in a loop.

    The last line written to the output stream is one JSON object:
        { "ok": true, "total": 3, "written": 2, "failed": 1, "locked": true }

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File apply_pending_writes.ps1
#>
[CmdletBinding()]
param(
    [string]$PendingPath = '',
    [string]$WriterPath = ''
)

$ErrorActionPreference = 'Stop'

$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$SystemDir = Split-Path -Parent $ScriptRoot
if ([string]::IsNullOrWhiteSpace($PendingPath)) { $PendingPath = Join-Path $SystemDir 'http-listener\system\generated\pending_workbook_writes.csv' }
if ([string]::IsNullOrWhiteSpace($WriterPath)) { $WriterPath = Join-Path $ScriptRoot 'update_dimension_excel.ps1' }

$LogDir = Join-Path $SystemDir 'logs'
if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogPath = Join-Path $LogDir 'excel_writeback.log'
$Columns = @('Type', 'ID', 'Name', 'Length_mm', 'Width_mm', 'Height_mm', 'Gross_Weight_kg', 'Net_Weight_kg', 'Carton_Gross_Weight_kg', 'Carton_Net_Weight_kg', 'Queued_At', 'Attempts', 'Last_Error')

function Write-PendingLog {
    param([string]$Message)
    $line = '{0} [PENDING] {1}' -f ([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')), $Message
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 } catch { }
}

function Send-Summary {
    param([bool]$Ok, [int]$Total, [int]$Written, [int]$Failed, [bool]$Locked, [string]$ErrorText = '')
    $summary = [ordered]@{ ok = $Ok; total = $Total; written = $Written; failed = $Failed; locked = $Locked }
    if ($ErrorText) { $summary['error'] = $ErrorText }
    Write-Output ($summary | ConvertTo-Json -Compress)
}

function ConvertTo-CsvField {
    param([object]$Value)
    $s = ''
    if ($null -ne $Value) { $s = [string]$Value }
    return '"' + $s.Replace('"', '""') + '"'
}

function Save-Pending {
    <# Rewrites the queue with what is still waiting, or removes the file when nothing is. #>
    param($Rows)
    if (@($Rows).Count -eq 0) {
        Remove-Item -LiteralPath $PendingPath -Force -ErrorAction SilentlyContinue
        return
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $_ }) -join ','))
    [void]$sb.Append([Environment]::NewLine)
    foreach ($r in $Rows) {
        [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $r.$_ }) -join ','))
        [void]$sb.Append([Environment]::NewLine)
    }
    $tmp = "$PendingPath.tmp"
    [IO.File]::WriteAllText($tmp, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
    [IO.File]::Copy($tmp, $PendingPath, $true)
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $PendingPath)) { Send-Summary $true 0 0 0 $false; exit 0 }
if (-not (Test-Path -LiteralPath $WriterPath)) {
    Write-PendingLog "Writer script not found at $WriterPath; nothing was written."
    Send-Summary $false 0 0 0 $false 'Excel writer script is missing.'
    exit 1
}

$rows = @()
try { $rows = @(Import-Csv -LiteralPath $PendingPath -Encoding UTF8) }
catch {
    Write-PendingLog "Could not read the queue: $($_.Exception.Message)"
    Send-Summary $false 0 0 0 $false ("Could not read the queue: {0}" -f $_.Exception.Message)
    exit 1
}
if ($rows.Count -eq 0) { Save-Pending @(); Send-Summary $true 0 0 0 $false; exit 0 }

Write-PendingLog "Writing $($rows.Count) queued save(s) into the workbook."
$remaining = New-Object System.Collections.ArrayList
$written = 0
$locked = $false

foreach ($row in $rows) {
    if ($locked) { [void]$remaining.Add($row); continue }   # no point hammering a locked workbook

    $type = ([string]$row.Type).Trim().ToLowerInvariant()
    $id = ([string]$row.ID).Trim()
    $errorText = ''
    if (@('product', 'carton') -notcontains $type -or -not $id) { $errorText = 'Unreadable queue entry (type or id missing).' }
    else {
        # Array elements go straight to the process, so a space in a path or a name needs no quoting.
        $writerArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $WriterPath,
            '-Type', $type, '-Id', $id, '-LengthMm', [string]$row.Length_mm, '-WidthMm', [string]$row.Width_mm, '-HeightMm', [string]$row.Height_mm, '-NoRebuild')
        if (-not [string]::IsNullOrWhiteSpace($row.Name)) { $writerArgs += @('-Name', [string]$row.Name) }
        if ($type -eq 'product') {
            foreach ($pair in @(@('Gross_Weight_kg', '-GrossWeightKg'), @('Net_Weight_kg', '-NetWeightKg'), @('Carton_Gross_Weight_kg', '-CartonGrossWeightKg'), @('Carton_Net_Weight_kg', '-CartonNetWeightKg'))) {
                $value = [string]$row.($pair[0])
                if (-not [string]::IsNullOrWhiteSpace($value)) { $writerArgs += @($pair[1], $value) }
            }
        }
        try {
            $output = & powershell.exe @writerArgs 2>&1
            $jsonLine = ($output | Where-Object { $_ -match '^\{.*\}$' } | Select-Object -Last 1)
            if (-not $jsonLine) { $errorText = 'The Excel writer did not return a result.' }
            else {
                $result = $jsonLine | ConvertFrom-Json
                if (-not $result.ok) { $errorText = [string]$result.error }
            }
        }
        catch { $errorText = $_.Exception.Message }
    }

    if (-not $errorText) {
        $written++
        Write-PendingLog "$type $id written."
        continue
    }

    if ($errorText -match 'open elsewhere|locked|used by another process') { $locked = $true }
    $attempts = 0
    [void][int]::TryParse([string]$row.Attempts, [ref]$attempts)
    $row.Attempts = [string]($attempts + 1)
    $row.Last_Error = $errorText
    [void]$remaining.Add($row)
    Write-PendingLog "$type $id NOT written: $errorText"
}

try { Save-Pending $remaining }
catch {
    Write-PendingLog "Could not update the queue file: $($_.Exception.Message)"
    Send-Summary $false $rows.Count $written $remaining.Count $locked ("Could not update the queue file: {0}" -f $_.Exception.Message)
    exit 1
}

Write-PendingLog "Done: $written written, $($remaining.Count) still waiting."
Send-Summary $true $rows.Count $written $remaining.Count $locked
exit 0
