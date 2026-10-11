<#
.SYNOPSIS
    Short-lived HTTP listener for WVN Master Carton Fit Tool data entry.

.DESCRIPTION
    Spawns on demand, serves index.html, accepts JSON POSTs, appends validated rows to CSV,
    then exits and releases every handle. Nothing stays resident, so Excel and the main
    Fit Tool can open the same CSVs at any time.

    Stock Windows only: PowerShell 5.1 + System.Net.HttpListener. No Python, Node, .NET Core,
    IIS, or package managers.

    Timeout model: the listener exits after -IdleTimeoutSeconds with NO traffic, not after a
    fixed wall-clock window. A hard 30s cap would kill the server while the user was still
    typing into the form. The browser sends POST /api/ping every 10s while the page is open,
    so the process stays alive exactly as long as someone is using it and no longer.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File listen.ps1
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File listen.ps1 -Port 3100 -IdleTimeoutSeconds 60
#>
[CmdletBinding()]
param(
    [int]$Port = 3000,
    [int]$IdleTimeoutSeconds = 30,
    [int]$MaxRequests = 100,
    [string]$LogPath = '',
    [string]$DataPath = '',
    [string]$UiPath = ''
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------------------------
# $PSScriptRoot comes back empty in a few real situations (dot-sourcing, some OneDrive-synced
# paths). Resolving it defensively is why this script starts instead of dying before it can log.
$ScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ScriptRoot) -and $MyInvocation.MyCommand.Path) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($ScriptRoot)) { $ScriptRoot = (Get-Location).ProviderPath }

if ([string]::IsNullOrWhiteSpace($DataPath)) { $DataPath = Join-Path $ScriptRoot 'system\generated' }
if ([string]::IsNullOrWhiteSpace($LogPath)) { $LogPath = Join-Path $ScriptRoot 'system\logs' }
if ([string]::IsNullOrWhiteSpace($UiPath)) { $UiPath = Join-Path $ScriptRoot 'index.html' }

foreach ($dir in @($DataPath, $LogPath)) {
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}
$LogFile = Join-Path $LogPath 'listener.log'

# ---------------------------------------------------------------------------------------------
# Dataset definitions - column order is contractual, do not reorder
# ---------------------------------------------------------------------------------------------
$Datasets = [ordered]@{
    'fg-dimensions'    = @{
        File    = 'fg_dimensions.csv'
        Columns = @('FG_ID', 'Length_mm', 'Width_mm', 'Height_mm', 'Target_Qty', 'Source', 'Timestamp')
        Keys    = @('FG_ID')
    }
    'carton-dimensions' = @{
        File    = 'carton_dimensions.csv'
        Columns = @('Carton_ID', 'Length_mm', 'Width_mm', 'Height_mm', 'Dimension_Source', 'Timestamp')
        Keys    = @('Carton_ID')
    }
    'bom-assignments'  = @{
        File    = 'bom_carton_assignments.csv'
        Columns = @('FG_ID', 'Carton_ID', 'Source_BOM_Date', 'Confidence', 'Timestamp')
        Keys    = @('FG_ID', 'Carton_ID')
    }
}

# Saved pallet layouts are not rows in a CSV: each carton gets one JSON document (schema in
# docs\pallet_layout.schema.json) under <DataPath>\pallet_layouts\. One file per carton keeps a
# save atomic and independent of every other carton, and the files can be copied or backed up as-is.
$PalletLayoutSchemaVersion = 1
$PalletLayoutDir = Join-Path $DataPath 'pallet_layouts'
$PalletLayoutMaxLevels = 100
$PalletLayoutMaxBoxesPerLevel = 400
$PalletLayoutMaxBodyChars = 2000000

$SourceEnum = @('MANUAL_ENTRY', 'D365_IMPORT', 'MOBILE_APP', 'BATCH_IMPORT')
$DimensionSourceEnum = @('BOM_LINKED_CA_AVERAGE', 'MANUAL_MEASUREMENT', 'SUPPLIER_SPEC')
$ConfidenceEnum = @('HIGH', 'MEDIUM', 'LOW')
$IdPattern = '^[A-Za-z0-9_-]+$'

# ---------------------------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------------------------
function Get-UtcStamp { return [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ') }

function Write-TextAppend {
    <# Append text with retries. Excel takes an exclusive lock on an open CSV, so a transient
       IOException is expected rather than exceptional; only give up after several attempts. #>
    param([string]$Path, [string]$Text, [int]$Attempts = 4, [int]$DelayMs = 150)

    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $fs = New-Object System.IO.FileStream($Path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
            try { $fs.Write($bytes, 0, $bytes.Length); $fs.Flush() } finally { $fs.Dispose() }
            return
        }
        catch [System.IO.IOException] {
            if ($i -eq $Attempts) { throw }
            Start-Sleep -Milliseconds $DelayMs
        }
    }
}

function Write-Log {
    param([string]$Message)
    $line = '{0} {1}{2}' -f (Get-UtcStamp), $Message, [Environment]::NewLine
    try { Write-TextAppend -Path $LogFile -Text $line } catch { }
    Write-Host ('  {0}' -f $Message)
}

function ConvertTo-CsvField {
    param([object]$Value)
    $s = ''
    if ($null -ne $Value) { $s = [string]$Value }
    return '"' + $s.Replace('"', '""') + '"'
}

function Read-CsvSafe {
    <# Opened with FileShare.ReadWrite so a read never blocks anyone else, and so the listener
       never becomes the reason a file appears locked. #>
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    $text = ''
    $reader = $null
    try {
        $stream = New-Object System.IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    return @($text | ConvertFrom-Csv)
}

function Write-TextReplace {
    <# Whole-file rewrite via a temp file + move, so a crash mid-write cannot leave a
       half-written CSV. Same retry policy as the append path. #>
    param([string]$Path, [string]$Text, [int]$Attempts = 4, [int]$DelayMs = 150)

    $tmp = "$Path.tmp"
    [IO.File]::WriteAllText($tmp, $Text, (New-Object Text.UTF8Encoding($false)))
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            [IO.File]::Copy($tmp, $Path, $true)
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            return
        }
        catch [System.IO.IOException] {
            if ($i -eq $Attempts) {
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                throw
            }
            Start-Sleep -Milliseconds $DelayMs
        }
    }
}

function Write-CsvAll {
    <# Rewrite the whole file from $Rows. Rows may be PSCustomObjects read back from CSV or
       hashtables built from a request; $row.$column resolves against both. #>
    param([string]$Path, [string[]]$Columns, $Rows)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $_ }) -join ','))
    [void]$sb.Append([Environment]::NewLine)
    foreach ($r in $Rows) {
        [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $r.$_ }) -join ','))
        [void]$sb.Append([Environment]::NewLine)
    }
    Write-TextReplace -Path $Path -Text $sb.ToString()
}

function Add-CsvRecord {
    param([string]$Path, [string[]]$Columns, [hashtable]$Record)

    $needHeader = $true
    if (Test-Path -LiteralPath $Path) {
        if ((Get-Item -LiteralPath $Path).Length -gt 0) { $needHeader = $false }
    }

    $sb = New-Object System.Text.StringBuilder
    if ($needHeader) {
        [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $_ }) -join ','))
        [void]$sb.Append([Environment]::NewLine)
    }
    [void]$sb.Append((($Columns | ForEach-Object { ConvertTo-CsvField $Record[$_] }) -join ','))
    [void]$sb.Append([Environment]::NewLine)

    Write-TextAppend -Path $Path -Text $sb.ToString()
}

# ---------------------------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------------------------
function Get-Field {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function New-FieldError {
    param([string]$Field, [string]$Message)
    return [pscustomobject]@{ field = $Field; message = $Message }
}

function Test-Blank {
    param($Value)
    if ($null -eq $Value) { return $true }
    return [string]::IsNullOrWhiteSpace([string]$Value)
}

function Resolve-IdField {
    param($Value, [string]$Field, $Errors)
    if (Test-Blank $Value) { [void]$Errors.Add((New-FieldError $Field 'Required')); return $null }
    $s = ([string]$Value).Trim()
    if ($s.Length -gt 50) { [void]$Errors.Add((New-FieldError $Field 'Must be 50 characters or fewer')); return $null }
    if ($s -notmatch $IdPattern) {
        [void]$Errors.Add((New-FieldError $Field 'Only letters, digits, dash and underscore are allowed'))
        return $null
    }
    return $s
}

function Resolve-Dimension {
    param($Value, [string]$Field, $Errors)
    if (Test-Blank $Value) { [void]$Errors.Add((New-FieldError $Field 'Required')); return $null }
    $parsed = 0.0
    $ok = [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)
    if (-not $ok) { [void]$Errors.Add((New-FieldError $Field 'Must be numeric')); return $null }
    if ($parsed -lt 10 -or $parsed -gt 1000) { [void]$Errors.Add((New-FieldError $Field 'Must be 10-1000 mm')); return $null }
    return $parsed
}

function Resolve-TargetQty {
    param($Value, [string]$Field, $Errors)
    if (Test-Blank $Value) { return $null }   # optional
    $parsed = 0.0
    $ok = [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)
    if (-not $ok) { [void]$Errors.Add((New-FieldError $Field 'Must be numeric')); return $null }
    if ([Math]::Floor($parsed) -ne $parsed) { [void]$Errors.Add((New-FieldError $Field 'Must be a whole number')); return $null }
    if ($parsed -lt 1 -or $parsed -gt 10000) { [void]$Errors.Add((New-FieldError $Field 'Must be 1-10000')); return $null }
    return [int]$parsed
}

function Resolve-Enum {
    param($Value, [string]$Field, [string[]]$Allowed, [string]$Default, $Errors)
    if (Test-Blank $Value) {
        if ($Default) { return $Default }
        [void]$Errors.Add((New-FieldError $Field 'Required'))
        return $null
    }
    $s = ([string]$Value).Trim().ToUpperInvariant()
    if ($Allowed -notcontains $s) {
        [void]$Errors.Add((New-FieldError $Field ('Must be one of: {0}' -f ($Allowed -join ', '))))
        return $null
    }
    return $s
}

function Resolve-IsoDate {
    param($Value, [string]$Field, $Errors)
    if (Test-Blank $Value) { [void]$Errors.Add((New-FieldError $Field 'Required')); return $null }
    $s = ([string]$Value).Trim()
    $dt = [DateTime]::MinValue
    # TryParseExact rejects impossible calendar dates such as 2026-02-30, which a regex would pass.
    $ok = [DateTime]::TryParseExact($s, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dt)
    if (-not $ok) { [void]$Errors.Add((New-FieldError $Field 'Must be a valid date in YYYY-MM-DD format')); return $null }
    return $dt.ToString('yyyy-MM-dd')
}

# ---------------------------------------------------------------------------------------------
# Record builders - return @{ Ok; Errors; Record; Key; Conflict }
# ---------------------------------------------------------------------------------------------
function Build-FgRecord {
    param($Body)
    $errors = New-Object System.Collections.ArrayList
    $id = Resolve-IdField (Get-Field $Body 'fg_id') 'fg_id' $errors
    $len = Resolve-Dimension (Get-Field $Body 'length_mm') 'length_mm' $errors
    $wid = Resolve-Dimension (Get-Field $Body 'width_mm') 'width_mm' $errors
    $hei = Resolve-Dimension (Get-Field $Body 'height_mm') 'height_mm' $errors
    $qty = Resolve-TargetQty (Get-Field $Body 'target_qty') 'target_qty' $errors
    $src = Resolve-Enum (Get-Field $Body 'source') 'source' $SourceEnum 'MANUAL_ENTRY' $errors

    return @{
        Ok     = ($errors.Count -eq 0)
        Errors = $errors
        Key    = @{ FG_ID = $id }
        Record = @{
            FG_ID      = $id
            Length_mm  = $len
            Width_mm   = $wid
            Height_mm  = $hei
            Target_Qty = $qty
            Source     = $src
            Timestamp  = (Get-UtcStamp)
        }
    }
}

function Build-CartonRecord {
    param($Body)
    $errors = New-Object System.Collections.ArrayList
    $id = Resolve-IdField (Get-Field $Body 'carton_id') 'carton_id' $errors
    $len = Resolve-Dimension (Get-Field $Body 'length_mm') 'length_mm' $errors
    $wid = Resolve-Dimension (Get-Field $Body 'width_mm') 'width_mm' $errors
    $hei = Resolve-Dimension (Get-Field $Body 'height_mm') 'height_mm' $errors
    $dsrc = Resolve-Enum (Get-Field $Body 'dimension_source') 'dimension_source' $DimensionSourceEnum '' $errors

    return @{
        Ok     = ($errors.Count -eq 0)
        Errors = $errors
        Key    = @{ Carton_ID = $id }
        Record = @{
            Carton_ID        = $id
            Length_mm        = $len
            Width_mm         = $wid
            Height_mm        = $hei
            Dimension_Source = $dsrc
            Timestamp        = (Get-UtcStamp)
        }
    }
}

function Build-BomRecord {
    param($Body)
    $errors = New-Object System.Collections.ArrayList
    $fg = Resolve-IdField (Get-Field $Body 'fg_id') 'fg_id' $errors
    $carton = Resolve-IdField (Get-Field $Body 'carton_id') 'carton_id' $errors
    $date = Resolve-IsoDate (Get-Field $Body 'source_bom_date') 'source_bom_date' $errors
    $conf = Resolve-Enum (Get-Field $Body 'confidence') 'confidence' $ConfidenceEnum 'MEDIUM' $errors

    # Referential integrity: both parents must already exist, otherwise the assignment is an orphan.
    if ($errors.Count -eq 0) {
        $fgRows = Read-CsvSafe (Join-Path $DataPath $Datasets['fg-dimensions'].File)
        if (-not ($fgRows | Where-Object { $_.FG_ID -eq $fg })) {
            [void]$errors.Add((New-FieldError 'fg_id' ("FG_ID '{0}' not found in fg_dimensions.csv" -f $fg)))
        }
        $cartonRows = Read-CsvSafe (Join-Path $DataPath $Datasets['carton-dimensions'].File)
        if (-not ($cartonRows | Where-Object { $_.Carton_ID -eq $carton })) {
            [void]$errors.Add((New-FieldError 'carton_id' ("Carton_ID '{0}' not found in carton_dimensions.csv" -f $carton)))
        }
    }

    return @{
        Ok     = ($errors.Count -eq 0)
        Errors = $errors
        Key    = @{ FG_ID = $fg; Carton_ID = $carton }
        Record = @{
            FG_ID           = $fg
            Carton_ID       = $carton
            Source_BOM_Date = $date
            Confidence      = $conf
            Timestamp       = (Get-UtcStamp)
        }
    }
}

# ---------------------------------------------------------------------------------------------
# Insert pipeline - one record, returns @{ Status; Payload }
# ---------------------------------------------------------------------------------------------
function Invoke-Insert {
    param([string]$Dataset, $Body)

    $def = $Datasets[$Dataset]
    $csvPath = Join-Path $DataPath $def.File

    switch ($Dataset) {
        'fg-dimensions' { $built = Build-FgRecord $Body }
        'carton-dimensions' { $built = Build-CartonRecord $Body }
        'bom-assignments' { $built = Build-BomRecord $Body }
    }

    if (-not $built.Ok) {
        # Orphan-reference failures are integrity errors, still a 400 per the contract.
        return @{
            Status  = 400
            Payload = [ordered]@{
                success = $false
                status  = 400
                error   = 'Validation failed'
                details = @($built.Errors)
            }
        }
    }

    # Re-read per record so a batch sees rows written earlier within the same batch.
    $existing = @(Read-CsvSafe $csvPath)
    $isDuplicate = $false
    foreach ($r in $existing) {
        $match = $true
        foreach ($k in $def.Keys) {
            if ([string]$r.$k -ne [string]$built.Key[$k]) { $match = $false; break }
        }
        if ($match) { $isDuplicate = $true; break }
    }

    try {
        if ($isDuplicate) {
            # Re-entering an existing item REPLACES it. Someone re-measuring a product is
            # correcting it, so the newest measurement is the one that should win; making them
            # hunt down and delete the old row first would be the wrong default.
            $kept = New-Object System.Collections.ArrayList
            $written = $false
            foreach ($r in $existing) {
                $match = $true
                foreach ($k in $def.Keys) {
                    if ([string]$r.$k -ne [string]$built.Key[$k]) { $match = $false; break }
                }
                if ($match) {
                    # Replace in place the first time, drop any further duplicates.
                    if (-not $written) { [void]$kept.Add($built.Record); $written = $true }
                }
                else { [void]$kept.Add($r) }
            }
            Write-CsvAll -Path $csvPath -Columns $def.Columns -Rows $kept
        }
        else {
            Add-CsvRecord -Path $csvPath -Columns $def.Columns -Record $built.Record
        }
    }
    catch [System.IO.IOException] {
        return @{
            Status  = 503
            Payload = [ordered]@{
                success = $false
                status  = 503
                error   = ('CSV file locked. Close Excel and retry. ({0})' -f $def.File)
            }
        }
    }

    $primary = [string]$built.Key[$def.Keys[0]]
    $unixMs = [long]([DateTime]::UtcNow - [DateTime]'1970-01-01').TotalMilliseconds
    return @{
        Status  = 200
        Payload = [ordered]@{
            success     = $true
            record_id   = ('{0}-{1}' -f $primary.ToLowerInvariant(), $unixMs)
            replaced    = $isDuplicate
            timestamp   = $built.Record.Timestamp
            csv_updated = $true
            csv_path    = ('system/generated/{0}' -f $def.File)
        }
        Key      = $primary
        Replaced = $isDuplicate
    }
}

# ---------------------------------------------------------------------------------------------
# One product / carton dimension request, validated. Shared by the immediate Excel write and by
# the queue of writes that wait for the next workbook update.
# ---------------------------------------------------------------------------------------------
$DimensionWeightFields = @(
    @{ field = 'gross_weight_kg';        arg = '-GrossWeightKg';       column = 'Gross_Weight_kg' },
    @{ field = 'net_weight_kg';          arg = '-NetWeightKg';         column = 'Net_Weight_kg' },
    @{ field = 'carton_gross_weight_kg'; arg = '-CartonGrossWeightKg'; column = 'Carton_Gross_Weight_kg' },
    @{ field = 'carton_net_weight_kg';   arg = '-CartonNetWeightKg';   column = 'Carton_Net_Weight_kg' }
)

function Resolve-DimensionRequest {
    param($parsed)

    $dimErrors = New-Object System.Collections.ArrayList
    $dimType = Resolve-Enum (Get-Field $parsed 'type') 'type' @('PRODUCT', 'CARTON') '' $dimErrors
    $dimId = Resolve-IdField (Get-Field $parsed 'id') 'id' $dimErrors
    $dimLength = Resolve-Dimension (Get-Field $parsed 'length_mm') 'length_mm' $dimErrors
    $dimWidth = Resolve-Dimension (Get-Field $parsed 'width_mm') 'width_mm' $dimErrors
    $dimHeight = Resolve-Dimension (Get-Field $parsed 'height_mm') 'height_mm' $dimErrors

    # Optional weights (kg) - product saves only. The Dimension sheet keeps four per FG: FG gross / FG net on its
    # EA row and carton gross / carton net on its CA row. Blank means "leave the workbook cell alone";
    # anything supplied must be a sensible positive number and a net weight may not exceed its gross.
    $dimWeights = [ordered]@{}
    if ($dimType -eq 'PRODUCT') {
        foreach ($wf in $DimensionWeightFields) {
            $weightRaw = Get-Field $parsed $wf.field
            if (Test-Blank $weightRaw) { continue }
            $weightParsed = 0.0
            $weightOk = [double]::TryParse([string]$weightRaw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$weightParsed)
            if (-not $weightOk) { [void]$dimErrors.Add((New-FieldError $wf.field 'Must be numeric')) }
            elseif ($weightParsed -le 0 -or $weightParsed -gt 1000) { [void]$dimErrors.Add((New-FieldError $wf.field 'Must be greater than 0 and at most 1000 kg')) }
            else { $dimWeights[$wf.field] = $weightParsed }
        }
        if ($dimWeights.Contains('gross_weight_kg') -and $dimWeights.Contains('net_weight_kg') -and $dimWeights['net_weight_kg'] -gt $dimWeights['gross_weight_kg']) {
            [void]$dimErrors.Add((New-FieldError 'net_weight_kg' 'Net weight cannot exceed the gross weight'))
        }
        if ($dimWeights.Contains('carton_gross_weight_kg') -and $dimWeights.Contains('carton_net_weight_kg') -and $dimWeights['carton_net_weight_kg'] -gt $dimWeights['carton_gross_weight_kg']) {
            [void]$dimErrors.Add((New-FieldError 'carton_net_weight_kg' 'Net weight cannot exceed the gross weight'))
        }
    }

    $dimNameRaw = Get-Field $parsed 'name'
    $dimName = ''
    if (-not (Test-Blank $dimNameRaw)) {
        $dimName = ([string]$dimNameRaw).Trim()
        if ($dimName.Length -gt 100) { $dimName = $dimName.Substring(0, 100) }
    }

    return @{ Errors = $dimErrors; Type = $dimType; Id = $dimId; Length = $dimLength; Width = $dimWidth; Height = $dimHeight; Weights = $dimWeights; Name = $dimName }
}

# Saves made in the tool that still have to be written into the dimension workbook. The launcher
# writes them (tools\apply_pending_writes.ps1) right before it next reads the workbook.
$PendingWritesFile = 'pending_workbook_writes.csv'
$PendingWritesColumns = @('Type', 'ID', 'Name', 'Length_mm', 'Width_mm', 'Height_mm', 'Gross_Weight_kg', 'Net_Weight_kg', 'Carton_Gross_Weight_kg', 'Carton_Net_Weight_kg', 'Queued_At', 'Attempts', 'Last_Error')

function Format-InvariantNumber {
    param($Value)
    return ([double]$Value).ToString([Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------------------------------------
# Launcher preference kept in system\config.local.ini (per PC, overrides config.ini): how often
# the launcher re-reads the Excel workbooks.
# ---------------------------------------------------------------------------------------------
$UpdateDataModes = @('ask', 'always', 'daily', 'weekly', 'changed', 'skip')
$ConfigDir = Split-Path -Parent $ScriptRoot

function Get-IniValue {
    param([string]$Path, [string]$Section, [string]$Key)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $inSection = $false
    foreach ($line in [IO.File]::ReadAllLines($Path)) {
        $t = $line.Trim()
        if ($t -match '^\[(.+)\]$') { $inSection = ($Matches[1].Trim() -ieq $Section); continue }
        if ($inSection -and $t -notmatch '^[;#]' -and $t -match '^([^=]+)=(.*)$' -and $Matches[1].Trim() -ieq $Key) { return $Matches[2].Trim() }
    }
    return ''
}

function Set-IniValue {
    <# Rewrites one key and leaves every other line of the file exactly as it was. #>
    param([string]$Path, [string]$Section, [string]$Key, [string]$Value)
    $lines = New-Object System.Collections.Generic.List[string]
    if (Test-Path -LiteralPath $Path) { foreach ($line in [IO.File]::ReadAllLines($Path)) { $lines.Add($line) } }
    $sectionAt = -1; $sectionEnd = $lines.Count; $keyAt = -1; $inSection = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($t -match '^\[(.+)\]$') {
            if ($inSection) { $sectionEnd = $i; break }
            if ($Matches[1].Trim() -ieq $Section) { $inSection = $true; $sectionAt = $i }
            continue
        }
        if ($inSection -and $t -notmatch '^[;#]' -and $t -match '^([^=]+)=' -and $Matches[1].Trim() -ieq $Key) { $keyAt = $i; break }
    }
    $entry = '{0}={1}' -f $Key, $Value
    if ($keyAt -ge 0) { $lines[$keyAt] = $entry }
    elseif ($sectionAt -ge 0) {
        $at = $sectionEnd
        while ($at -gt $sectionAt + 1 -and [string]::IsNullOrWhiteSpace($lines[$at - 1])) { $at-- }
        $lines.Insert($at, $entry)
    }
    else {
        if ($lines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) { $lines.Add('') }
        $lines.Add('[{0}]' -f $Section)
        $lines.Add($entry)
    }
    Write-TextReplace -Path $Path -Text (($lines -join [Environment]::NewLine) + [Environment]::NewLine)
}

function Get-UpdateDataMode {
    foreach ($file in @('config.local.ini', 'config.ini')) {
        $value = (Get-IniValue (Join-Path $ConfigDir $file) 'Launcher' 'UpdateData').ToLowerInvariant()
        if ($UpdateDataModes -contains $value) { return @{ Mode = $value; Source = $file } }
    }
    return @{ Mode = 'ask'; Source = 'default' }
}

# ---------------------------------------------------------------------------------------------
# Saved pallet layouts - one JSON document per carton (docs\pallet_layout.schema.json)
# ---------------------------------------------------------------------------------------------
function ConvertTo-JsonString {
    <# A JSON string literal, escaped by hand. The stored document is assembled as text so it
       never goes through ConvertTo-Json, which in PowerShell 5.1 mangles nested arrays that
       came out of ConvertFrom-Json. Everything outside printable ASCII becomes \uXXXX. #>
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int]$ch
        if ($code -eq 34) { [void]$sb.Append('\"') }
        elseif ($code -eq 92) { [void]$sb.Append('\\') }
        elseif ($code -lt 32 -or $code -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $code)) }
        else { [void]$sb.Append($ch) }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Format-JsonNumber {
    <# Millimetres to one decimal, always with a '.' whatever the PC's regional settings are. #>
    param([double]$Value)
    $rounded = [Math]::Round($Value, 1, [MidpointRounding]::AwayFromZero)
    if ($rounded -eq 0) { return '0' }
    return $rounded.ToString('0.#', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-JsonNumber {
    <# $null unless the value is a real JSON number - "12" or true must not pass as 12 or 1. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if (-not ($Value -is [int] -or $Value -is [long] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [single])) { return $null }
    $d = [double]$Value
    if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return $null }
    return $d
}

function Get-PalletLayoutPath {
    <# Carton IDs are free text in the workbook, so anything outside A-Z 0-9 - _ is written
       as %XX. Upper-cased first: "ab-1" and "AB-1" are the same carton and share one file. #>
    param([string]$CartonId)
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in [Text.Encoding]::UTF8.GetBytes($CartonId.Trim().ToUpperInvariant())) {
        if (($b -ge 48 -and $b -le 57) -or ($b -ge 65 -and $b -le 90) -or $b -eq 45 -or $b -eq 95) { [void]$sb.Append([char]$b) }
        else { [void]$sb.Append(('%{0:X2}' -f $b)) }
    }
    return (Join-Path $PalletLayoutDir ($sb.ToString() + '.json'))
}

function Build-PalletLayout {
    <# Validates one layout and returns the canonical document text. Geometry is checked here
       as well as in the page: every carton must be the carton's own footprint (either way
       round), sit inside the pallet, and not overlap another carton on its level. #>
    param($Body)

    $errors = New-Object System.Collections.ArrayList
    $tol = 0.6

    if ($null -eq $Body -or $Body -is [System.Array] -or $Body -isnot [System.Management.Automation.PSCustomObject]) {
        [void]$errors.Add((New-FieldError 'body' 'Expected one JSON object'))
        return @{ Ok = $false; Errors = $errors }
    }

    $version = Get-JsonNumber (Get-Field $Body 'schema_version')
    if ($null -eq $version -or $version -ne $PalletLayoutSchemaVersion) {
        [void]$errors.Add((New-FieldError 'schema_version' ("schema_version must be {0}" -f $PalletLayoutSchemaVersion)))
    }

    $cartonId = ''
    $rawId = Get-Field $Body 'carton_id'
    if ($rawId -is [string]) { $cartonId = $rawId.Trim() }
    if ($cartonId.Length -lt 1 -or $cartonId.Length -gt 80 -or $cartonId -match '[\x00-\x1F]') {
        [void]$errors.Add((New-FieldError 'carton_id' 'carton_id is required (text, up to 80 characters)'))
    }

    $dims = @{}
    foreach ($spec in @(
            @('carton', 'length_mm', 1, 5000), @('carton', 'width_mm', 1, 5000), @('carton', 'height_mm', 1, 5000),
            @('pallet', 'length_mm', 100, 5000), @('pallet', 'width_mm', 100, 5000))) {
        $name = '{0}.{1}' -f $spec[0], $spec[1]
        $n = Get-JsonNumber (Get-Field (Get-Field $Body $spec[0]) $spec[1])
        if ($null -eq $n -or $n -lt $spec[2] -or $n -gt $spec[3]) {
            [void]$errors.Add((New-FieldError $name ("{0} must be a number between {1} and {2}" -f $name, $spec[2], $spec[3])))
        }
        else { $dims[$name] = $n }
    }

    $method = Get-Field $Body 'ship_method'
    if (-not ($method -is [string]) -or @('air', 'truck') -cnotcontains $method) {
        [void]$errors.Add((New-FieldError 'ship_method' "ship_method must be 'air' or 'truck'"))
    }

    $maxHeight = Get-JsonNumber (Get-Field $Body 'max_height_mm')
    if ($null -eq $maxHeight -or $maxHeight -lt 100 -or $maxHeight -gt 5000) {
        [void]$errors.Add((New-FieldError 'max_height_mm' 'max_height_mm must be a number between 100 and 5000'))
    }

    # Optional: how far (mm, per side) cartons were allowed to hang over the pallet edge when the
    # layout was built. Absent means 0 - the layout stays inside the deck.
    $overhang = 0.0
    $rawOverhang = Get-Field $Body 'overhang_mm'
    if ($null -ne $rawOverhang) {
        $overhangNum = Get-JsonNumber $rawOverhang
        if ($null -eq $overhangNum -or $overhangNum -lt 0 -or $overhangNum -gt 200) {
            [void]$errors.Add((New-FieldError 'overhang_mm' 'overhang_mm must be a number between 0 and 200'))
        }
        else { $overhang = $overhangNum }
    }

    $origin = 'manual'
    $rawOrigin = Get-Field $Body 'origin'
    if ($null -ne $rawOrigin) {
        if (-not ($rawOrigin -is [string]) -or @('manual', 'suggested') -cnotcontains $rawOrigin) {
            [void]$errors.Add((New-FieldError 'origin' "origin must be 'manual' or 'suggested'"))
        }
        else { $origin = $rawOrigin }
    }

    $sourceFg = ''
    $rawFg = Get-Field $Body 'source_fg_id'
    if ($null -ne $rawFg) {
        if (-not ($rawFg -is [string]) -or $rawFg.Trim().Length -gt 80 -or $rawFg -match '[\x00-\x1F]') {
            [void]$errors.Add((New-FieldError 'source_fg_id' 'source_fg_id must be text, up to 80 characters'))
        }
        else { $sourceFg = $rawFg.Trim() }
    }

    # Property access, not Get-Field: returning an array from a function unrolls it, and a
    # one-level layout would arrive here as that level instead of as a list of levels.
    $levels = $null
    $levelsProp = $Body.PSObject.Properties['levels']
    if ($null -ne $levelsProp) { $levels = $levelsProp.Value }
    if ($levels -isnot [System.Array] -or $levels.Count -lt 1 -or $levels.Count -gt $PalletLayoutMaxLevels) {
        [void]$errors.Add((New-FieldError 'levels' ("levels must be a list of 1 to {0} levels" -f $PalletLayoutMaxLevels)))
    }

    if ($errors.Count -gt 0) { return @{ Ok = $false; Errors = $errors } }

    $cl = $dims['carton.length_mm']; $cw = $dims['carton.width_mm']; $ch = $dims['carton.height_mm']
    $pl = $dims['pallet.length_mm']; $pw = $dims['pallet.width_mm']

    if ($levels.Count * $ch -gt $maxHeight + $tol) {
        [void]$errors.Add((New-FieldError 'levels' ("{0} levels of {1} mm are taller than max_height_mm {2}" -f $levels.Count, $ch, $maxHeight)))
        return @{ Ok = $false; Errors = $errors }
    }

    $levelText = New-Object System.Collections.ArrayList
    $cartonCount = 0
    for ($k = 0; $k -lt $levels.Count -and $errors.Count -lt 10; $k++) {
        $level = $levels[$k]
        $where = 'levels[{0}]' -f $k
        if ($level -isnot [System.Array] -or $level.Count -lt 1 -or $level.Count -gt $PalletLayoutMaxBoxesPerLevel) {
            [void]$errors.Add((New-FieldError $where ("each level must hold 1 to {0} cartons" -f $PalletLayoutMaxBoxesPerLevel)))
            continue
        }
        $rects = New-Object System.Collections.ArrayList
        $parts = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $level.Count -and $errors.Count -lt 10; $i++) {
            $box = $level[$i]
            $at = '{0}[{1}]' -f $where, $i
            $x = Get-JsonNumber (Get-Field $box 'x'); $y = Get-JsonNumber (Get-Field $box 'y')
            $l = Get-JsonNumber (Get-Field $box 'l'); $w = Get-JsonNumber (Get-Field $box 'w')
            if ($null -eq $x -or $null -eq $y -or $null -eq $l -or $null -eq $w) {
                [void]$errors.Add((New-FieldError $at 'a carton needs numeric x, y, l and w (mm)'))
                continue
            }
            $asIs = ([Math]::Abs($l - $cl) -le $tol) -and ([Math]::Abs($w - $cw) -le $tol)
            $turned = ([Math]::Abs($l - $cw) -le $tol) -and ([Math]::Abs($w - $cl) -le $tol)
            if (-not ($asIs -or $turned)) {
                [void]$errors.Add((New-FieldError $at ("l x w must be the carton footprint {0} x {1} or {1} x {0}" -f $cl, $cw)))
                continue
            }
            if ($x -lt (0 - $overhang - $tol) -or $y -lt (0 - $overhang - $tol) -or ($x + $l) -gt ($pl + $overhang + $tol) -or ($y + $w) -gt ($pw + $overhang + $tol)) {
                [void]$errors.Add((New-FieldError $at ("carton lies outside the {0} x {1} pallet (overhang allowed: {2} mm)" -f $pl, $pw, $overhang)))
                continue
            }
            $clash = $false
            foreach ($o in $rects) {
                $ox = [Math]::Min($x + $l, $o[0] + $o[2]) - [Math]::Max($x, $o[0])
                $oy = [Math]::Min($y + $w, $o[1] + $o[3]) - [Math]::Max($y, $o[1])
                if ($ox -gt 1.0 -and $oy -gt 1.0) { $clash = $true; break }
            }
            if ($clash) {
                [void]$errors.Add((New-FieldError $at 'carton overlaps another carton on the same level'))
                continue
            }
            [void]$rects.Add(@($x, $y, $l, $w))
            [void]$parts.Add(('{{"x":{0},"y":{1},"l":{2},"w":{3}}}' -f (Format-JsonNumber $x), (Format-JsonNumber $y), (Format-JsonNumber $l), (Format-JsonNumber $w)))
        }
        $cartonCount += $rects.Count
        [void]$levelText.Add('    [' + ($parts -join ',') + ']')
    }

    if ($errors.Count -gt 0) { return @{ Ok = $false; Errors = $errors } }

    $savedAt = Get-UtcStamp
    $nl = [Environment]::NewLine
    $lines = @(
        '{',
        '  "schema": "wvn.pallet-layout",',
        ('  "schema_version": {0},' -f $PalletLayoutSchemaVersion),
        ('  "carton_id": {0},' -f (ConvertTo-JsonString $cartonId)),
        ('  "carton": {{ "length_mm": {0}, "width_mm": {1}, "height_mm": {2} }},' -f (Format-JsonNumber $cl), (Format-JsonNumber $cw), (Format-JsonNumber $ch)),
        ('  "pallet": {{ "length_mm": {0}, "width_mm": {1} }},' -f (Format-JsonNumber $pl), (Format-JsonNumber $pw)),
        ('  "ship_method": {0},' -f (ConvertTo-JsonString $method)),
        ('  "max_height_mm": {0},' -f (Format-JsonNumber $maxHeight)),
        ('  "overhang_mm": {0},' -f (Format-JsonNumber $overhang)),
        ('  "level_count": {0},' -f $levels.Count),
        ('  "carton_count": {0},' -f $cartonCount),
        ('  "origin": {0},' -f (ConvertTo-JsonString $origin)),
        ('  "source_fg_id": {0},' -f (ConvertTo-JsonString $sourceFg)),
        ('  "saved_at": {0},' -f (ConvertTo-JsonString $savedAt)),
        ('  "saved_by": {0},' -f (ConvertTo-JsonString ([string]$env:USERNAME))),
        '  "levels": [',
        ($levelText -join (',' + $nl)),
        '  ]',
        '}'
    )
    return @{
        Ok          = $true
        Json        = ($lines -join $nl)
        CartonId    = $cartonId
        LevelCount  = $levels.Count
        CartonCount = $cartonCount
        SavedAt     = $savedAt
    }
}

function Read-PalletLayouts {
    <# Every stored document as text. A file someone edited by hand into invalid JSON is
       skipped (and logged) rather than allowed to break the whole response. #>
    $docs = New-Object System.Collections.ArrayList
    if (Test-Path -LiteralPath $PalletLayoutDir) {
        foreach ($file in @(Get-ChildItem -LiteralPath $PalletLayoutDir -Filter '*.json' -File | Sort-Object Name)) {
            try {
                $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8).Trim()
                if ([string]::IsNullOrWhiteSpace($text)) { continue }
                $null = $text | ConvertFrom-Json
                [void]$docs.Add($text)
            }
            catch { Write-Log ('[WARN] Skipped unreadable pallet layout {0}: {1}' -f $file.Name, $_.Exception.Message) }
        }
    }
    return , $docs
}

# ---------------------------------------------------------------------------------------------
# HTTP plumbing
# ---------------------------------------------------------------------------------------------
function Send-Json {
    param($Response, [int]$StatusCode, $Payload)
    $json = $Payload | ConvertTo-Json -Depth 10
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = 'application/json; charset=utf-8'
    $Response.ContentLength64 = $bytes.Length
    $Response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Send-Bytes {
    param($Response, [int]$StatusCode, [byte[]]$Bytes, [string]$ContentType)
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.ContentLength64 = $Bytes.Length
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
}

# ---------------------------------------------------------------------------------------------
# Start
# ---------------------------------------------------------------------------------------------
$listener = New-Object System.Net.HttpListener
# "localhost" is the one prefix HTTP.SYS lets a non-admin bind. "+" or "*" would need an
# elevated netsh urlacl, which we cannot assume on a factory PC.
$prefix = 'http://localhost:{0}/' -f $Port
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
}
catch {
    Write-Host ''
    Write-Host ('  Could not listen on {0}' -f $prefix) -ForegroundColor Red
    Write-Host '  Port ' -NoNewline -ForegroundColor Red
    Write-Host ('{0} is already in use. Another instance is probably still running.' -f $Port) -ForegroundColor Red
    Write-Host '  Close it and retry, or start this script with a different -Port.' -ForegroundColor Red
    Write-Log ('[ERROR] Could not bind {0}: {1}' -f $prefix, $_.Exception.Message)
    exit 1
}

Write-Host ''
Write-Host ('  Listening on {0}' -f $prefix) -ForegroundColor Green
Write-Host ('  Data   -> {0}' -f $DataPath)
Write-Host ('  Logs   -> {0}' -f $LogFile)
Write-Host ('  Exits after {0}s idle or {1} data requests. Press Ctrl+C to stop now.' -f $IdleTimeoutSeconds, $MaxRequests)
Write-Host ''
Write-Log ('[INFO] Listening on {0} (idle timeout {1}s, max {2} requests)' -f $prefix, $IdleTimeoutSeconds, $MaxRequests)

$requestCount = 0
$shutdownRequested = $false
$lastActivity = Get-Date
$pending = $null
$exitReason = 'idle timeout'
# 0 = normal idle window; >0 = the page signalled it was closing, so use this shorter one.
$graceSeconds = 0

try {
    while ($true) {
        if ($shutdownRequested) { $exitReason = 'client requested shutdown'; break }
        if ($MaxRequests -gt 0 -and $requestCount -ge $MaxRequests) { $exitReason = 'request limit reached'; break }
        $idleLimit = $IdleTimeoutSeconds
        if ($graceSeconds -gt 0) { $idleLimit = $graceSeconds }
        if ($idleLimit -gt 0 -and ((Get-Date) - $lastActivity).TotalSeconds -ge $idleLimit) {
            if ($graceSeconds -gt 0) { $exitReason = 'browser closed' } else { $exitReason = 'idle timeout' }
            break
        }

        # Hold ONE pending accept. Re-issuing BeginGetContext on every tick would stack
        # abandoned async operations for as long as the process lived.
        if ($null -eq $pending) { $pending = $listener.BeginGetContext($null, $null) }
        if (-not $pending.AsyncWaitHandle.WaitOne(400)) { continue }

        $context = $listener.EndGetContext($pending)
        $pending = $null
        $lastActivity = Get-Date
        # Any request at all means someone is still using the tool, so cancel a pending
        # release countdown. This has to happen here rather than further down: the ping and
        # health handlers 'continue' out of the loop body before any later statement runs.
        $graceSeconds = 0

        $req = $context.Request
        $res = $context.Response
        $method = $req.HttpMethod
        $path = $req.Url.AbsolutePath
        if ($path.Length -gt 1) { $path = $path.TrimEnd('/') }

        try {
            $res.Headers['Access-Control-Allow-Origin'] = '*'
            $res.Headers['Access-Control-Allow-Headers'] = 'Content-Type'
            $res.Headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
            $res.Headers['Cache-Control'] = 'no-store'

            if ($method -eq 'OPTIONS') {
                $res.StatusCode = 204
                continue
            }

            $body = ''
            if ($req.HasEntityBody) {
                $enc = $req.ContentEncoding
                if ($null -eq $enc) { $enc = [Text.Encoding]::UTF8 }
                $sr = New-Object System.IO.StreamReader($req.InputStream, $enc)
                try { $body = $sr.ReadToEnd() } finally { $sr.Dispose() }
            }

            # ---- static UI -------------------------------------------------------------
            if ($method -eq 'GET' -and ($path -eq '/' -or $path -eq '/index.html')) {
                if (Test-Path -LiteralPath $UiPath) {
                    Send-Bytes $res 200 ([IO.File]::ReadAllBytes($UiPath)) 'text/html; charset=utf-8'
                }
                else {
                    Send-Json $res 404 ([ordered]@{ success = $false; status = 404; error = "index.html not found at $UiPath" })
                }
                continue
            }
            if ($method -eq 'GET' -and $path -eq '/favicon.ico') { $res.StatusCode = 204; continue }

            # ---- lifecycle -------------------------------------------------------------
            if ($path -eq '/api/health' -and $method -eq 'GET') {
                Send-Json $res 200 ([ordered]@{
                        success   = $true
                        status    = 'listening'
                        port      = $Port
                        requests  = $requestCount
                        data_path = $DataPath
                        timestamp = (Get-UtcStamp)
                    })
                continue
            }
            if ($path -eq '/api/ping') {
                # Keep-alive only. Deliberately does NOT increment $requestCount, or an idle
                # page would burn through -MaxRequests and shut the listener down.
                Send-Json $res 200 ([ordered]@{ success = $true; pong = (Get-UtcStamp) })
                continue
            }
            if ($path -eq '/api/shutdown') {
                Send-Json $res 200 ([ordered]@{ success = $true; message = 'Shutting down' })
                $shutdownRequested = $true
                Write-Log '[INFO] Shutdown requested by client'
                continue
            }
            if ($path -eq '/api/release') {
                # The page is going away - but "going away" also fires on F5 and on
                # in-page navigation. Exiting immediately would kill the listener out from
                # under a simple refresh, so shorten the idle window instead. If the page
                # comes back its keep-alive ping lands well inside the grace period and the
                # cancellation below restores the normal timeout.
                Send-Json $res 200 ([ordered]@{ success = $true; message = 'Releasing after grace period' })
                $graceSeconds = 5
                Write-Log '[INFO] Page released; exiting in 5s unless it reconnects'
                continue
            }

            # ---- excel write-back --------------------------------------------------------
            if ($path -eq '/api/save-dimension') {
                # Minted once per request and threaded through to the writer via -TrackId, so
                # the writer's own log lines use the SAME id rather than minting a second one -
                # one id traces a failure across the UI toast, this log, and excel_writeback.log.
                $trackId = 'TRK-' + [Guid]::NewGuid().ToString().Substring(0, 8).ToUpperInvariant()

                if ($method -ne 'POST') {
                    Send-Json $res 405 ([ordered]@{ success = $false; status = 405; error = 'Method not allowed'; track_id = $trackId })
                    continue
                }

                $requestCount++

                if ([string]::IsNullOrWhiteSpace($body)) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Request body is empty'; track_id = $trackId })
                    Write-Log ('[POST] {0} [400] [{1}] empty body' -f $path, $trackId)
                    continue
                }

                try { $parsed = $body | ConvertFrom-Json }
                catch {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Invalid JSON'; detail = $_.Exception.Message; track_id = $trackId })
                    Write-Log ('[POST] {0} [400] [{1}] malformed JSON' -f $path, $trackId)
                    continue
                }

                $dim = Resolve-DimensionRequest $parsed
                $dimErrors = $dim.Errors; $dimType = $dim.Type; $dimId = $dim.Id; $dimName = $dim.Name
                $dimLength = $dim.Length; $dimWidth = $dim.Width; $dimHeight = $dim.Height; $dimWeights = $dim.Weights
                $weightFields = $DimensionWeightFields

                if ($dimErrors.Count -gt 0) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Validation failed'; details = @($dimErrors); track_id = $trackId })
                    Write-Log ('[POST] {0} [400] [{1}] validation failed' -f $path, $trackId)
                    continue
                }

                $writerType = $dimType.ToLowerInvariant()
                $writerScript = Join-Path (Split-Path -Parent $ScriptRoot) 'tools\update_dimension_excel.ps1'
                if (-not (Test-Path -LiteralPath $writerScript)) {
                    Send-Json $res 500 ([ordered]@{ success = $false; status = 500; error = 'Excel writer script is missing.'; track_id = $trackId })
                    Write-Log ('[POST] {0} [500] [{1}] writer script missing' -f $path, $trackId)
                    continue
                }

                # Array elements passed straight through to the process, not a joined command
                # line, so a space anywhere in this project's own path never needs quoting here.
                $writerArgs = @(
                    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $writerScript,
                    '-Type', $writerType, '-Id', $dimId, '-LengthMm', $dimLength, '-WidthMm', $dimWidth, '-HeightMm', $dimHeight,
                    '-TrackId', $trackId
                )
                if ($dimName) { $writerArgs += @('-Name', $dimName) }
                foreach ($wf in $weightFields) {
                    if ($dimWeights.Contains($wf.field)) { $writerArgs += @($wf.arg, ([double]$dimWeights[$wf.field]).ToString([Globalization.CultureInfo]::InvariantCulture)) }
                }

                $writerOutput = & powershell.exe @writerArgs 2>&1
                $jsonLine = ($writerOutput | Where-Object { $_ -match '^\{.*\}$' } | Select-Object -Last 1)

                if (-not $jsonLine) {
                    Send-Json $res 502 ([ordered]@{
                            success = $false; status = 502
                            error   = 'The Excel writer did not return a result.'
                            detail  = (($writerOutput | Select-Object -Last 5) -join ' | ')
                            track_id = $trackId
                        })
                    Write-Log ('[POST] {0} [502] [{1}] writer produced no JSON output' -f $path, $trackId)
                    continue
                }

                try { $writerResult = $jsonLine | ConvertFrom-Json }
                catch {
                    Send-Json $res 502 ([ordered]@{ success = $false; status = 502; error = 'Could not parse the Excel writer result.'; detail = $jsonLine; track_id = $trackId })
                    Write-Log ('[POST] {0} [502] [{1}] writer output not valid JSON' -f $path, $trackId)
                    continue
                }

                if ($writerResult.ok) {
                    Send-Json $res 200 ([ordered]@{
                            success               = $true
                            status                = 'success'
                            message               = 'Saved to Excel and cache updated'
                            row                   = $writerResult.row
                            is_new                = $writerResult.is_new
                            sheet                 = $writerResult.sheet
                            backup_path           = $writerResult.backup_path
                            cache_rebuild_started = $writerResult.cache_rebuild_started
                            write_ms              = $writerResult.write_ms
                            track_id              = $trackId
                        })
                    Write-Log ('[POST] {0} [200] [{1}] {2} {3} saved to Excel (row {4}, {5}ms)' -f $path, $trackId, $writerType, $dimId, $writerResult.row, $writerResult.write_ms)
                }
                else {
                    $lockStatus = 502
                    if ($writerResult.error -match 'open elsewhere|locked|used by another process') { $lockStatus = 503 }
                    Send-Json $res $lockStatus ([ordered]@{
                            success     = $false
                            status      = $lockStatus
                            error       = [string]$writerResult.error
                            backup_path = $writerResult.backup_path
                            track_id    = $trackId
                        })
                    Write-Log ('[POST] {0} [{1}] [{2}] excel write failed: {3}' -f $path, $lockStatus, $trackId, $writerResult.error)
                }
                continue
            }

            # ---- saves waiting for the workbook -----------------------------------------
            if ($path -eq '/api/pending-writes' -and $method -eq 'GET') {
                $rows = @(Read-CsvSafe (Join-Path $DataPath $PendingWritesFile))
                Send-Json $res 200 ([ordered]@{ success = $true; count = $rows.Count; rows = $rows })
                continue
            }
            if ($path -eq '/api/queue-dimension') {
                if ($method -ne 'POST') {
                    Send-Json $res 405 ([ordered]@{ success = $false; status = 405; error = 'Method not allowed' })
                    continue
                }
                $requestCount++
                if ([string]::IsNullOrWhiteSpace($body)) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Request body is empty' })
                    continue
                }
                try { $parsed = $body | ConvertFrom-Json }
                catch {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Invalid JSON'; detail = $_.Exception.Message })
                    Write-Log ('[POST] {0} [400] malformed JSON' -f $path)
                    continue
                }
                $dim = Resolve-DimensionRequest $parsed
                if ($dim.Errors.Count -gt 0) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Validation failed'; details = @($dim.Errors) })
                    Write-Log ('[POST] {0} [400] validation failed' -f $path)
                    continue
                }
                $record = @{
                    Type = $dim.Type.ToLowerInvariant(); ID = $dim.Id; Name = $dim.Name
                    Length_mm = (Format-InvariantNumber $dim.Length); Width_mm = (Format-InvariantNumber $dim.Width); Height_mm = (Format-InvariantNumber $dim.Height)
                    Queued_At = (Get-UtcStamp); Attempts = '0'; Last_Error = ''
                }
                foreach ($wf in $DimensionWeightFields) {
                    $record[$wf.column] = ''
                    if ($dim.Weights.Contains($wf.field)) { $record[$wf.column] = (Format-InvariantNumber $dim.Weights[$wf.field]) }
                }
                $pendingPath = Join-Path $DataPath $PendingWritesFile
                try {
                    # A newer save of the same item replaces the one still waiting.
                    $kept = New-Object System.Collections.ArrayList
                    $replacedPending = $false
                    foreach ($r in @(Read-CsvSafe $pendingPath)) {
                        if (([string]$r.Type -ieq $record.Type) -and ([string]$r.ID -ieq $record.ID)) { $replacedPending = $true } else { [void]$kept.Add($r) }
                    }
                    [void]$kept.Add($record)
                    Write-CsvAll -Path $pendingPath -Columns $PendingWritesColumns -Rows $kept
                }
                catch [System.IO.IOException] {
                    Send-Json $res 503 ([ordered]@{ success = $false; status = 503; error = ('CSV file locked. Close Excel and retry. ({0})' -f $PendingWritesFile) })
                    Write-Log ('[POST] {0} [503] csv locked' -f $path)
                    continue
                }
                Send-Json $res 200 ([ordered]@{ success = $true; queued = $true; replaced = $replacedPending; pending = $kept.Count; type = $record.Type; id = $record.ID })
                Write-Log ('[POST] {0} [200] {1} {2} queued for the workbook ({3} waiting)' -f $path, $record.Type, $record.ID, $kept.Count)
                continue
            }

            # ---- launcher preference: how often the workbooks are re-read ----------------
            if ($path -eq '/api/launcher-settings') {
                if ($method -eq 'GET') {
                    $current = Get-UpdateDataMode
                    Send-Json $res 200 ([ordered]@{ success = $true; update_data = $current.Mode; source = $current.Source; options = $UpdateDataModes })
                    continue
                }
                if ($method -ne 'POST') {
                    Send-Json $res 405 ([ordered]@{ success = $false; status = 405; error = 'Method not allowed' })
                    continue
                }
                $requestCount++
                $parsed = $null
                try { $parsed = $body | ConvertFrom-Json } catch { }
                $requested = [string](Get-Field $parsed 'update_data')
                if ($UpdateDataModes -cnotcontains $requested) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = ('update_data must be one of: {0}' -f ($UpdateDataModes -join ', ')) })
                    Write-Log ('[POST] {0} [400] bad update_data' -f $path)
                    continue
                }
                try { Set-IniValue (Join-Path $ConfigDir 'config.local.ini') 'Launcher' 'UpdateData' $requested }
                catch {
                    Send-Json $res 503 ([ordered]@{ success = $false; status = 503; error = ('Could not write config.local.ini: {0}' -f $_.Exception.Message) })
                    Write-Log ('[POST] {0} [503] {1}' -f $path, $_.Exception.Message)
                    continue
                }
                Send-Json $res 200 ([ordered]@{ success = $true; update_data = $requested; source = 'config.local.ini' })
                Write-Log ('[POST] {0} [200] UpdateData={1}' -f $path, $requested)
                continue
            }

            # ---- saved pallet layouts ---------------------------------------------------
            if ($path -eq '/api/pallet-layouts') {
                $jsonType = 'application/json; charset=utf-8'

                if ($method -eq 'GET') {
                    $docs = Read-PalletLayouts
                    $text = '{{"success":true,"schema_version":{0},"count":{1},"layouts":[{2}]}}' -f $PalletLayoutSchemaVersion, $docs.Count, ($docs -join ',')
                    Send-Bytes $res 200 ([Text.Encoding]::UTF8.GetBytes($text)) $jsonType
                    Write-Log ('[GET] {0} [200] {1} layouts' -f $path, $docs.Count)
                    continue
                }

                if ($method -ne 'POST') {
                    Send-Json $res 405 ([ordered]@{ success = $false; status = 405; error = 'Method not allowed' })
                    continue
                }

                $requestCount++

                if ([string]::IsNullOrWhiteSpace($body)) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Request body is empty' })
                    Write-Log ('[POST] {0} [400] empty body' -f $path)
                    continue
                }
                if ($body.Length -gt $PalletLayoutMaxBodyChars) {
                    Send-Json $res 413 ([ordered]@{ success = $false; status = 413; error = 'Layout is too large' })
                    Write-Log ('[POST] {0} [413] {1} chars' -f $path, $body.Length)
                    continue
                }

                try { $parsed = $body | ConvertFrom-Json }
                catch {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Invalid JSON'; detail = $_.Exception.Message })
                    Write-Log ('[POST] {0} [400] malformed JSON' -f $path)
                    continue
                }

                $built = Build-PalletLayout $parsed
                if (-not $built.Ok) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Validation failed'; details = @($built.Errors) })
                    Write-Log ('[POST] {0} [400] validation failed: {1}' -f $path, (($built.Errors | ForEach-Object { $_.field }) -join ', '))
                    continue
                }

                $layoutPath = Get-PalletLayoutPath $built.CartonId
                $replaced = Test-Path -LiteralPath $layoutPath
                try {
                    if (-not (Test-Path -LiteralPath $PalletLayoutDir)) { New-Item -ItemType Directory -Path $PalletLayoutDir -Force | Out-Null }
                    Write-TextReplace -Path $layoutPath -Text ($built.Json + [Environment]::NewLine)
                }
                catch [System.IO.IOException] {
                    Send-Json $res 503 ([ordered]@{ success = $false; status = 503; error = 'The layout file is locked. Close whatever has it open and retry.' })
                    Write-Log ('[POST] {0} [503] layout file locked' -f $path)
                    continue
                }

                $replacedText = 'false'
                if ($replaced) { $replacedText = 'true' }
                $text = '{{"success":true,"replaced":{0},"layout":{1}}}' -f $replacedText, $built.Json
                Send-Bytes $res 200 ([Text.Encoding]::UTF8.GetBytes($text)) $jsonType
                $verb = 'saved'
                if ($replaced) { $verb = 'replaced' }
                Write-Log ('[POST] {0} [200] {1} {2} ({3} levels, {4} cartons)' -f $path, $built.CartonId, $verb, $built.LevelCount, $built.CartonCount)
                continue
            }

            # ---- data ------------------------------------------------------------------
            if ($path -like '/api/data/*') {
                $dataset = $path.Substring('/api/data/'.Length)
                if (-not $Datasets.Contains($dataset)) {
                    Send-Json $res 404 ([ordered]@{ success = $false; status = 404; error = "Unknown dataset '$dataset'" })
                    Write-Log ('[{0}] {1} [404] unknown dataset' -f $method, $path)
                    continue
                }
                $def = $Datasets[$dataset]
                $csvPath = Join-Path $DataPath $def.File

                if ($method -eq 'GET') {
                    $rows = @(Read-CsvSafe $csvPath)
                    Send-Json $res 200 ([ordered]@{ success = $true; count = $rows.Count; rows = $rows })
                    Write-Log ('[GET] {0} [200] {1} rows' -f $path, $rows.Count)
                    continue
                }

                if ($method -ne 'POST') {
                    Send-Json $res 405 ([ordered]@{ success = $false; status = 405; error = 'Method not allowed' })
                    continue
                }

                $requestCount++

                if ([string]::IsNullOrWhiteSpace($body)) {
                    Send-Json $res 400 ([ordered]@{ success = $false; status = 400; error = 'Request body is empty' })
                    Write-Log ('[POST] {0} [400] empty body' -f $path)
                    continue
                }

                try { $parsed = $body | ConvertFrom-Json }
                catch {
                    Send-Json $res 400 ([ordered]@{
                            success = $false; status = 400
                            error   = 'Invalid JSON'
                            detail  = $_.Exception.Message
                        })
                    Write-Log ('[POST] {0} [400] malformed JSON' -f $path)
                    continue
                }

                # A JSON array means batch mode; a single object is one record.
                if ($parsed -is [System.Array]) {
                    $results = New-Object System.Collections.ArrayList
                    $inserted = 0
                    $failed = 0
                    for ($i = 0; $i -lt $parsed.Count; $i++) {
                        $r = Invoke-Insert -Dataset $dataset -Body $parsed[$i]
                        $entry = [ordered]@{ index = $i; status = $r.Status; success = ($r.Status -eq 200) }
                        foreach ($p in $r.Payload.Keys) {
                            if ($p -ne 'success') { $entry[$p] = $r.Payload[$p] }
                        }
                        [void]$results.Add([pscustomobject]$entry)
                        if ($r.Status -eq 200) { $inserted++ } else { $failed++ }
                    }
                    # success = "the batch was processed", not "every row inserted". A batch
                    # where 2 of 3 rows land is a partial success the client reports per row;
                    # flagging the whole response false made it render as a hard failure.
                    Send-Json $res 200 ([ordered]@{
                            success  = $true
                            inserted = $inserted
                            failed   = $failed
                            results  = @($results)
                        })
                    Write-Log ('[POST] {0} [200] batch: {1} inserted, {2} failed' -f $path, $inserted, $failed)
                    continue
                }

                $result = Invoke-Insert -Dataset $dataset -Body $parsed
                Send-Json $res $result.Status $result.Payload
                if ($result.Status -eq 200) {
                    $verb = 'inserted'
                    if ($result.Replaced) { $verb = 'replaced (newer data kept)' }
                    Write-Log ('[POST] {0} [200] {1} {2}' -f $path, $result.Key, $verb)
                }
                elseif ($result.Status -eq 503) {
                    Write-Log ('[POST] {0} [503] csv locked' -f $path)
                }
                else {
                    $fields = ($result.Payload.details | ForEach-Object { $_.field }) -join ', '
                    Write-Log ('[POST] {0} [400] validation failed: {1}' -f $path, $fields)
                }
                continue
            }

            Send-Json $res 404 ([ordered]@{ success = $false; status = 404; error = "No route for $method $path" })
            Write-Log ('[{0}] {1} [404] no route' -f $method, $path)
        }
        catch {
            # Never let one bad request kill the listener.
            try {
                Send-Json $res 500 ([ordered]@{ success = $false; status = 500; error = 'Internal error'; detail = $_.Exception.Message })
            }
            catch { }
            Write-Log ('[{0}] {1} [500] {2}' -f $method, $path, $_.Exception.Message)
        }
        finally {
            try { $res.OutputStream.Close() } catch { }
            try { $res.Close() } catch { }
        }
    }
}
finally {
    if ($pending) { try { $listener.Abort() } catch { } }
    try { $listener.Stop() } catch { }
    try { $listener.Close() } catch { }
    Write-Log ('[INFO] Listener shutting down ({0})' -f $exitReason)
    Write-Log ('[INFO] Processed {0} data requests' -f $requestCount)
    Write-Log '[INFO] All handles released'
    Write-Host ''
    Write-Host ('  Stopped: {0}. {1} data requests handled. CSVs are unlocked.' -f $exitReason, $requestCount) -ForegroundColor Yellow
}
