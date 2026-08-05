```powershell
<#
.SYNOPSIS
  Test suite for listen.ps1 HttpListener.
.DESCRIPTION
  Self-contained test script. Starts listen.ps1, runs a battery of functional and
  non-functional tests, writes JSON results, and exits with 0 on all-pass, 1 otherwise.
.PARAMETER Port
  Port to use (default 3097).
.PARAMETER ListenerPath
  Path to listen.ps1 (default ..\listen.ps1 relative to this script).
.PARAMETER KeepArtifacts
  Do not delete the _run directory after execution.
#>

param(
    [int]$Port = 3097,
    [string]$ListenerPath,
    [switch]$KeepArtifacts
)

# ============================================================
# 1. SCRIPT-SCOPE STATE
# ============================================================
$script:TestResults    = @()
$script:CurrentTestPass = $true
$script:CurrentTestMessages = @()

# ============================================================
# 2. HELPER: WRITE-ASSERT (used by assertion functions)
# ============================================================
function Write-AssertFail {
    param([string]$Message)
    $script:CurrentTestPass = $false
    $script:CurrentTestMessages += $Message
    Write-Host "    [!] $Message" -ForegroundColor Yellow
}

# ============================================================
# 3. ASSERTION HELPERS
# ============================================================
function Assert-Equal {
    param($Expected, $Actual, [string]$Message = "")
    if ($Expected -ne $Actual) {
        Write-AssertFail "Assert-Equal: $Message (expected '$Expected', actual '$Actual')"
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message = "")
    if (-not $Condition) {
        Write-AssertFail "Assert-True: $Message"
    }
}

function Assert-Match {
    param([string]$Pattern, [string]$String, [string]$Message = "")
    if ($String -notmatch $Pattern) {
        Write-AssertFail "Assert-Match: $Message (pattern '$Pattern' not found)"
    }
}

function Assert-FileExists {
    param([string]$Path, [string]$Message = "")
    if (-not (Test-Path $Path)) {
        Write-AssertFail "Assert-FileExists: '$Path' not found. $Message"
    }
}

# ============================================================
# 4. TEST RUNNER
# ============================================================
function Test-Case {
    param([string]$Name, [scriptblock]$ScriptBlock)
    $script:CurrentTestPass = $true
    $script:CurrentTestMessages = @()
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        & $ScriptBlock
    } catch {
        $script:CurrentTestPass = $false
        $script:CurrentTestMessages += "Exception: $($_.Exception.Message)"
    }
    $timer.Stop()
    $durationMs = $timer.Elapsed.TotalMilliseconds
    $result = [PSCustomObject]@{
        Name       = $Name
        Pass       = $script:CurrentTestPass
        Message    = ($script:CurrentTestMessages -join "; ")
        DurationMs = $durationMs
    }
    $script:TestResults += $result
    if ($script:CurrentTestPass) {
        Write-Host "[PASS] $Name ($([math]::Round($durationMs))ms)" -ForegroundColor Green
    } else {
        Write-Host "[FAIL] $Name ($([math]::Round($durationMs))ms)" -ForegroundColor Red
        if ($script:CurrentTestMessages.Count -gt 0) {
            Write-Host "      Reason: $($script:CurrentTestMessages -join ' ; ')" -ForegroundColor Red
        }
    }
}

# ============================================================
# 5. API CALL HELPER (handles 4xx/5xx gracefully)
# ============================================================
function Invoke-Api {
    param(
        [string]$Method = 'GET',
        [string]$Uri,
        [string]$Body,
        [hashtable]$Headers = @{},
        [string]$ContentType = 'application/json'
    )
    try {
        $iwrParams = @{
            Method          = $Method
            Uri             = $Uri
            UseBasicParsing = $true
        }
        if ($Body) {
            $iwrParams.Body = $Body
            $iwrParams.ContentType = $ContentType
        }
        if ($Headers.Count -gt 0) {
            $iwrParams.Headers = $Headers
        }
        $resp = Invoke-WebRequest @iwrParams
        $statusCode = [int]$resp.StatusCode
        $bodyStr = $resp.Content
        $bodyObj = $null
        if ($bodyStr) {
            try { $bodyObj = $bodyStr | ConvertFrom-Json } catch { $bodyObj = $bodyStr }
        }
        return [PSCustomObject]@{
            StatusCode = $statusCode
            Body       = $bodyObj
            RawBody    = $bodyStr
        }
    } catch [System.Net.WebException] {
        $ex = $_.Exception
        if ($ex.Response -ne $null) {
            $respStream = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($respStream)
            $errorBody = $reader.ReadToEnd()
            $reader.Close()
            $ex.Response.Close()
            $statusCode = [int]$ex.Response.StatusCode
            $bodyObj = $null
            if ($errorBody) {
                try { $bodyObj = $errorBody | ConvertFrom-Json } catch { $bodyObj = $errorBody }
            }
            return [PSCustomObject]@{
                StatusCode = $statusCode
                Body       = $bodyObj
                RawBody    = $errorBody
            }
        }
        Write-Host "Invoke-Api: No response (connection refused?) - $($_.Exception.Message)" -ForegroundColor Red
        return $null
    } catch {
        Write-Host "Invoke-Api: Unhandled error - $_" -ForegroundColor Red
        return $null
    }
}

# ============================================================
# 6. SETUP
# ============================================================
$scriptStartTime = Get-Date

# Determine listener script path
if (-not $ListenerPath) {
    $ListenerPath = Join-Path $PSScriptRoot ".." "listen.ps1"
}
if (-not (Test-Path $ListenerPath)) {
    Write-Host "ERROR: listen.ps1 not found at '$ListenerPath'" -ForegroundColor Red
    exit 2
}

# Create throw-away directories
$runDir = Join-Path $PSScriptRoot "_run"
if (-not (Test-Path $runDir)) {
    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
}
$dataPath = Join-Path $runDir "Data"
$logPath  = Join-Path $runDir "Log"

# Clean previous run unless keep
if (-not $KeepArtifacts) {
    if (Test-Path $dataPath) { Remove-Item $dataPath -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $logPath)  { Remove-Item $logPath -Recurse -Force -ErrorAction SilentlyContinue }
}
New-Item -ItemType Directory -Path $dataPath -Force | Out-Null
New-Item -ItemType Directory -Path $logPath -Force  | Out-Null

# Start listener
$listenArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $ListenerPath,
    "-Port", $Port,
    "-IdleTimeoutSeconds", 600,
    "-MaxRequests", 10000,
    "-DataPath", $dataPath,
    "-LogPath", $logPath
)
Write-Host "Starting listen.ps1..."
$process = Start-Process powershell.exe -ArgumentList $listenArgs -PassThru -WindowStyle Hidden

# Poll health for up to 15 s
$baseUri = "http://localhost:$Port"
$healthUri = "$baseUri/api/health"
$healthy = $false
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 500
    $res = Invoke-Api -Method GET -Uri $healthUri
    if ($res -and $res.StatusCode -eq 200 -and $res.Body.success -eq $true) {
        $healthy = $true
        Write-Host "Listener is healthy."
        break
    }
}
if (-not $healthy) {
    Write-Host "FAIL: Listener did not become healthy within 15 seconds." -ForegroundColor Red
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    exit 1
}

# ============================================================
# 7. TEST CASES
# ============================================================

# ---- T01 ----------------------------------------------------
Test-Case "T01 health endpoint returns 200 and success:true" {
    $res = Invoke-Api -Method GET -Uri $healthUri
    Assert-Equal 200 $res.StatusCode "Status code should be 200"
    Assert-True ($res.Body.success -eq $true) "success property must be true"
}

# ---- T02 ----------------------------------------------------
Test-Case "T02 valid FG post returns 200 and success:true" {
    $body = @{
        fg_id      = "TEST_FG_001"
        length_mm  = 100
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "MANUAL_ENTRY"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 200 $res.StatusCode "Expected 200"
    Assert-True ($res.Body.success -eq $true) "success should be true"
}

# ---- T03 ----------------------------------------------------
Test-Case "T03 fg_dimensions.csv created with exact header" {
    $csvFile = Join-Path $dataPath "fg_dimensions.csv"
    Assert-FileExists $csvFile "CSV file must exist"
    if (Test-Path $csvFile) {
        $headerLine = Get-Content $csvFile -TotalCount 1
        $expected = "FG_ID,Length_mm,Width_mm,Height_mm,Target_Qty,Source,Timestamp"
        Assert-Equal $expected $headerLine.TrimEnd("`r","`n") "Header line mismatch"
    }
}

# ---- T04 ----------------------------------------------------
Test-Case "T04 fg_dimensions.csv contains posted FG_ID" {
    $csvFile = Join-Path $dataPath "fg_dimensions.csv"
    $rows = Import-Csv $csvFile
    $match = $rows | Where-Object { $_.FG_ID -eq "TEST_FG_001" }
    Assert-True ($match -ne $null) "Row with FG_ID 'TEST_FG_001' not found"
}

# ---- T05 ----------------------------------------------------
Test-Case "T05 duplicate FG_ID returns 409 and message 'already exists'" {
    $body = @{
        fg_id      = "TEST_FG_001"
        length_mm  = 100
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "MANUAL_ENTRY"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 409 $res.StatusCode "Expected 409 conflict"
    Assert-Match "already exists" "$($res.Body)" "Message must contain 'already exists'"
}

# ---- T06 ----------------------------------------------------
Test-Case "T06 dimension below 10 returns 400 with details[].field=length_mm" {
    $body = @{
        fg_id      = "TEST_FG_BELOW10"
        length_mm  = 5
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "MANUAL_ENTRY"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400"
    $details = $res.Body.details
    Assert-True ($details -and ($details -is [array]) -and ($details | Where-Object { $_.field -eq "length_mm" })) "Details should contain field 'length_mm'"
}

# ---- T07 ----------------------------------------------------
Test-Case "T07 dimension above 1000 returns 400" {
    $body = @{
        fg_id      = "TEST_FG_ABOVE1000"
        length_mm  = 2000
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "MANUAL_ENTRY"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400 for out-of-range dimension"
}

# ---- T08 ----------------------------------------------------
Test-Case "T08 missing required fg_id returns 400 with 'Required'" {
    $body = @{
        length_mm  = 100
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "MANUAL_ENTRY"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400"
    Assert-Match "Required" "$($res.Body)" "Message should mention 'Required'"
}

# ---- T09 ----------------------------------------------------
Test-Case "T09 bad enum for source returns 400" {
    $body = @{
        fg_id      = "TEST_FG_BADENUM"
        length_mm  = 100
        width_mm   = 200
        height_mm  = 300
        target_qty = 10
        source     = "INVALID_SOURCE"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400 for bad source enum"
}

# ---- T10 ----------------------------------------------------
Test-Case "T10 malformed JSON body returns 400 (not 500)" {
    $body = "this is not json at all"
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400 for malformed JSON, got $($res.StatusCode)"
}

# ---- T11 ----------------------------------------------------
Test-Case "T11 valid carton post returns 200 and carton_dimensions.csv gets row" {
    $body = @{
        carton_id         = "TEST_CARTON_001"
        length_mm         = 400
        width_mm          = 300
        height_mm         = 200
        dimension_source  = "MANUAL_MEASUREMENT"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/carton-dimensions" -Body $body
    Assert-Equal 200 $res.StatusCode "Expected 200"
    $csvFile = Join-Path $dataPath "carton_dimensions.csv"
    Assert-FileExists $csvFile "Carton CSV must exist"
    if (Test-Path $csvFile) {
        $rows = Import-Csv $csvFile
        $match = $rows | Where-Object { $_.Carton_ID -eq "TEST_CARTON_001" }
        Assert-True ($match -ne $null) "Row with Carton_ID 'TEST_CARTON_001' should be present"
    }
}

# ---- T12 ----------------------------------------------------
Test-Case "T12 bom-assignment referencing unknown FG_ID returns 400 with 'not found'" {
    $body = @{
        fg_id            = "NONEXISTENT_FG"
        carton_id        = "TEST_CARTON_001"
        source_bom_date  = "2023-01-01"
        confidence       = "LOW"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/bom-assignments" -Body $body
    Assert-Equal 400 $res.StatusCode "Expected 400 for missing FG_ID"
    Assert-Match "not found" "$($res.Body)" "Error should mention 'not found'"
}

# ---- T13 ----------------------------------------------------
Test-Case "T13 valid bom-assignment (both keys exist) returns 200" {
    $body = @{
        fg_id            = "TEST_FG_001"
        carton_id        = "TEST_CARTON_001"
        source_bom_date  = "2023-05-01"
        confidence       = "HIGH"
    } | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/bom-assignments" -Body $body
    Assert-Equal 200 $res.StatusCode "Expected 200"
}

# ---- T14 ----------------------------------------------------
Test-Case "T14 batch post of a 3-element JSON array returns inserted=3" {
    $batch = @(
        @{
            fg_id      = "BATCH_FG_1"
            length_mm  = 100
            width_mm   = 100
            height_mm  = 100
            target_qty = 5
            source     = "BATCH_IMPORT"
        },
        @{
            fg_id      = "BATCH_FG_2"
            length_mm  = 200
            width_mm   = 200
            height_mm  = 200
            target_qty = 5
            source     = "BATCH_IMPORT"
        },
        @{
            fg_id      = "BATCH_FG_3"
            length_mm  = 300
            width_mm   = 300
            height_mm  = 300
            target_qty = 5
            source     = "BATCH_IMPORT"
        }
    ) | ConvertTo-Json -Compress
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/data/fg-dimensions" -Body $batch
    Assert-Equal 200 $res.StatusCode "Expected 200"
    Assert-True ($res.Body.inserted -eq 3) "Inserted count should be 3, got $($res.Body.inserted)"
}

# ---- T15 ----------------------------------------------------
Test-Case "T15 GET /api/data/fg-dimensions returns count>=1 and rows shaped like CSV" {
    $res = Invoke-Api -Method GET -Uri "$baseUri/api/data/fg-dimensions"
    Assert-Equal 200 $res.StatusCode "Expected 200"
    Assert-True ($res.Body.count -ge 1) "Count should be >= 1"
    $rows = $res.Body.rows
    Assert-True ($rows -and $rows -is [array] -and $rows.Count -ge 1) "Rows must be an array with at least one item"
    $row = $rows[0]
    $props = @('FG_ID','Length_mm','Width_mm','Height_mm','Target_Qty','Source','Timestamp')
    foreach ($p in $props) {
        Assert-True ($row.PSObject.Properties.Name -contains $p) "Missing property '$p'"
    }
}

# ---- T16 ----------------------------------------------------
Test-Case "T16 NO FILE LOCKING â exclusive open succeeds" {
    $csvFile = Join-Path $dataPath "fg_dimensions.csv"
    Assert-FileExists $csvFile
    try {
        $fs = [System.IO.File]::Open($csvFile, 'Open', 'ReadWrite', 'None')
        $fs.Close()
        Assert-True $true "No lock held by listener"
    } catch {
        Write-AssertFail "File lock detected: $_"
    }
}

# ---- T17 ----------------------------------------------------
Test-Case "T17 listener.log exists and contains a line for POST to /api/data/fg-dimensions" {
    $logFile = Join-Path $logPath "listener.log"
    Assert-FileExists $logFile
    if (Test-Path $logFile) {
        $content = Get-Content $logFile -Raw
        Assert-Match "POST.*?/api/data/fg-dimensions" $content "Log should contain a POST entry for fg-dimensions"
    }
}

# ---- T18 ----------------------------------------------------
Test-Case "T18 p95 latency of 10 sequential health calls is under 200ms" {
    $times = @()
    for ($i=0; $i -lt 10; $i++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $null = Invoke-Api -Method GET -Uri $healthUri
        $sw.Stop()
        $times += $sw.Elapsed.TotalMilliseconds
    }
    $sorted = $times | Sort-Object
    # p95 of 10 samples: index 9.5 -> we'll use the 10th (max) as a conservative proxy
    $p95 = $sorted[-1]
    Assert-True ($p95 -lt 200) "p95 latency $([math]::Round($p95,1))ms must be <200ms"
}

# ---- Shutdown (T19) -----------------------------------------
Test-Case "T19 graceful shutdown exits listener process" {
    $res = Invoke-Api -Method POST -Uri "$baseUri/api/shutdown"
    Assert-Equal 200 $res.StatusCode "Shutdown should return 200"
    $waitMs = 0
    while (-not $process.HasExited -and $waitMs -lt 10000) {
        Start-Sleep -Milliseconds 500
        $waitMs += 500
    }
    Assert-True $process.HasExited "Process did not exit within 10s"
}

# ============================================================
# 8. FINAL CLEANUP & KILL GUARD
# ============================================================
try {
    if ($process -and -not $process.HasExited) {
        Write-Host "Killing listener process..." -ForegroundColor Yellow
        $process.Kill()
        $process.WaitForExit(5000) | Out-Null
    }
} catch {
    Write-Host "Cleanup error: $_" -ForegroundColor Red
}

# Delete artefacts unless asked to keep
if (-not $KeepArtifacts) {
    Remove-Item $runDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ============================================================
# 9. WRITE RESULTS JSON
# ============================================================
$endTime = Get-Date
$totalDurationMs = ($endTime - $scriptStartTime).TotalMilliseconds

$passed  = ($script:TestResults | Where-Object Pass).Count
$failed  = $script:TestResults.Count - $passed

$report = [PSCustomObject]@{
    total      = $script:TestResults.Count
    passed     = $passed
    failed     = $failed
    durationMs = [math]::Round($totalDurationMs, 1)
    results    = $script:TestResults | Select-Object Name, Pass, Message, DurationMs
}

$resultFile = Join-Path $runDir "test-results.json"
# ensure directory still exists (even though we may have deleted it â we recreate if needed)
if (-not (Test-Path $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }
$report | ConvertTo-Json -Depth 3 | Set-Content $resultFile -Encoding UTF8

# ============================================================
# 10. PRINT SUMMARY & EXIT
# ============================================================
Write-Host "`n========== TEST SUMMARY ==========" -ForegroundColor Cyan
$script:TestResults | Format-Table Name, Pass, DurationMs -AutoSize | Out-String | Write-Host
Write-Host "Total: $($script:TestResults.Count) | Passed: $passed | Failed: $failed | Duration: $([math]::Round($totalDurationMs,1))ms" -ForegroundColor Cyan
Write-Host "Results written to: $resultFile`n"

if ($failed -gt 0) {
    exit 1
} else {
    exit 0
}
```

**NOTES**  

- **Assumptions about API responses:**  
  - Batch POST returns an object with an `inserted` field indicating the number of successfully inserted rows.  
  - Validation errors (400) contain a `details` array of objects with at least a `field` property (used in T06).  
  - Duplicate detection (409) returns a message containing the phrase `"already exists"`.  

- **Listener behaviour:**  
  - The listener should release any file locks while idle; T16 verifies this.  
  - Idle timeout and maximum requests are set to high values to avoid interfering with the tests.  

- **Environment:**  
  - The script assumes it resides in a `tests` folder next to `listen.ps1`; the relative path `..\listen.ps1` is used by default.  
  - All CSVs are stored in `tests/_run/Data`, logs in `tests/_run/Log`; the `-KeepArtifacts` switch preserves them for manual inspection.  

- **Things not verified:**  
  - Exact error object formats for every validation failure were assumed based on common patterns; actual response structures should be confirmed against the real listen.ps1 implementation.  
  - The `p95` calculation for T18 uses the maximum of 10 samples because of the small sample size; this is a deliberate simplification.  
  - No tests cover the `/api/ping` (keep-alive) endpoint explicitly, as it is not part of the validation contract list.
