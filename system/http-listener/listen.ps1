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

                $dimErrors = New-Object System.Collections.ArrayList
                $dimType = Resolve-Enum (Get-Field $parsed 'type') 'type' @('PRODUCT', 'CARTON') '' $dimErrors
                $dimId = Resolve-IdField (Get-Field $parsed 'id') 'id' $dimErrors
                $dimLength = Resolve-Dimension (Get-Field $parsed 'length_mm') 'length_mm' $dimErrors
                $dimWidth = Resolve-Dimension (Get-Field $parsed 'width_mm') 'width_mm' $dimErrors
                $dimHeight = Resolve-Dimension (Get-Field $parsed 'height_mm') 'height_mm' $dimErrors

                # Optional gross weight (kg) of one finished good - product saves only. Blank means "leave the
                # workbook cell alone"; anything supplied must be a sensible positive number.
                $dimWeight = $null
                $weightRaw = Get-Field $parsed 'gross_weight_kg'
                if ($dimType -eq 'PRODUCT' -and -not (Test-Blank $weightRaw)) {
                    $weightParsed = 0.0
                    $weightOk = [double]::TryParse([string]$weightRaw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$weightParsed)
                    if (-not $weightOk) { [void]$dimErrors.Add((New-FieldError 'gross_weight_kg' 'Must be numeric')) }
                    elseif ($weightParsed -le 0 -or $weightParsed -gt 1000) { [void]$dimErrors.Add((New-FieldError 'gross_weight_kg' 'Must be greater than 0 and at most 1000 kg')) }
                    else { $dimWeight = $weightParsed }
                }

                $dimNameRaw = Get-Field $parsed 'name'
                $dimName = ''
                if (-not (Test-Blank $dimNameRaw)) {
                    $dimName = ([string]$dimNameRaw).Trim()
                    if ($dimName.Length -gt 100) { $dimName = $dimName.Substring(0, 100) }
                }

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
                if ($null -ne $dimWeight) { $writerArgs += @('-GrossWeightKg', $dimWeight.ToString([Globalization.CultureInfo]::InvariantCulture)) }

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
                    if ($writerResult.error -match 'already open elsewhere|locked') { $lockStatus = 503 }
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
