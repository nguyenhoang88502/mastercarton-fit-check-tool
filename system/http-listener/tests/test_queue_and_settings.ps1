<#
.SYNOPSIS
  Tests for the queue of workbook writes (POST /api/queue-dimension, GET /api/pending-writes)
  and the launcher preference route (GET/POST /api/launcher-settings) of listen.ps1.
.DESCRIPTION
  Runs a COPY of listen.ps1 from tests\_run\queue\system\http-listener, so the config files
  it reads and writes are throwaway copies and the real system\config.local.ini is never
  touched. No workbook is involved: writing the queue into Excel is tools\apply_pending_writes.ps1,
  which needs a real workbook and is exercised separately.
  Exits 0 when every check passes, 1 otherwise.
#>
$ErrorActionPreference = 'Stop'
$listenerSource = Join-Path (Join-Path $PSScriptRoot '..') 'listen.ps1'
$sandbox = Join-Path (Join-Path $PSScriptRoot '_run') 'queue'
if (Test-Path $sandbox) { Remove-Item $sandbox -Recurse -Force }
$sys = Join-Path $sandbox 'system'
$data = Join-Path $sys 'http-listener\system\generated'
New-Item -ItemType Directory -Force -Path $data | Out-Null
Copy-Item -LiteralPath $listenerSource -Destination (Join-Path $sys 'http-listener')
Set-Content -LiteralPath (Join-Path $sys 'config.ini') -Value "[Search]`r`nMaxDepth=6`r`n`r`n[Launcher]`r`nForceRebuild=false`r`n; a comment that must survive`r`nUpdateData=ask" -Encoding ASCII
Set-Content -LiteralPath (Join-Path $sys 'config.local.ini') -Value "[Launcher]`r`nPreferredBrowser=chrome" -Encoding ASCII

$port = 3195
$proc = Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$sys\http-listener\listen.ps1`"", '-Port', $port, '-IdleTimeoutSeconds', '40')
$base = "http://localhost:$port"
$up = $false
for ($i = 0; $i -lt 40; $i++) { try { if ((Invoke-RestMethod "$base/api/health" -TimeoutSec 2).success) { $up = $true; break } } catch { Start-Sleep -Milliseconds 300 } }
if (-not $up) { throw 'listener did not start' }

function Call($method, $path, $json) {
    try {
        $p = @{ Uri = "$base$path"; Method = $method; UseBasicParsing = $true; TimeoutSec = 15 }
        if ($json) { $p.Body = $json; $p.ContentType = 'application/json' }
        $r = Invoke-WebRequest @p
        return @{ code = [int]$r.StatusCode; body = ($r.Content | ConvertFrom-Json) }
    }
    catch {
        $resp = $_.Exception.Response; if ($null -eq $resp) { throw }
        $st = $resp.GetResponseStream(); try { $st.Position = 0 } catch { }; $sr = New-Object IO.StreamReader($st); $t = $sr.ReadToEnd(); $sr.Dispose()
        $b = $null; try { $b = $t | ConvertFrom-Json } catch { }
        return @{ code = [int]$resp.StatusCode; body = $b; text = $t }
    }
}
$pass = 0; $fail = 0
function Check($name, $cond, $info = '') { if ($cond) { $script:pass++; Write-Output "PASS  $name" } else { $script:fail++; Write-Output "FAIL  $name  $info" } }

try {
    # ---- launcher preference
    $r = Call GET '/api/launcher-settings'
    Check 'GET reads ask from config.ini' ($r.body.update_data -eq 'ask' -and $r.body.source -eq 'config.ini')
    $r = Call POST '/api/launcher-settings' '{"update_data":"daily"}'
    Check 'POST daily -> 200' ($r.code -eq 200 -and $r.body.update_data -eq 'daily')
    $local = Get-Content -LiteralPath (Join-Path $sys 'config.local.ini')
    Check 'config.local.ini keeps its other key' (($local -contains 'PreferredBrowser=chrome') -and ($local -contains 'UpdateData=daily') -and ($local[0] -eq '[Launcher]')) ($local -join ' | ')
    [void](Call POST '/api/launcher-settings' '{"update_data":"changed"}')
    $local = Get-Content -LiteralPath (Join-Path $sys 'config.local.ini')
    Check 'a second POST rewrites the same line' ((@($local | Where-Object { $_ -like 'UpdateData=*' }).Count -eq 1) -and ($local -contains 'UpdateData=changed')) ($local -join ' | ')
    $r = Call GET '/api/launcher-settings'
    Check 'GET now reads config.local.ini' ($r.body.update_data -eq 'changed' -and $r.body.source -eq 'config.local.ini')
    Check 'config.ini is left alone' ((Get-Content -LiteralPath (Join-Path $sys 'config.ini')) -contains 'UpdateData=ask')
    foreach ($mode in @('ask', 'always', 'weekly', 'skip')) {
        $r = Call POST '/api/launcher-settings' ('{"update_data":"' + $mode + '"}')
        Check "POST $mode accepted" ($r.code -eq 200)
    }
    $r = Call POST '/api/launcher-settings' '{"update_data":"hourly"}'
    Check 'unknown value -> 400' ($r.code -eq 400)
    $r = Call POST '/api/launcher-settings' '{"update_data":"Daily"}'
    Check 'value is case-sensitive -> 400' ($r.code -eq 400)

    # ---- queue of workbook writes
    $r = Call GET '/api/pending-writes'
    Check 'queue starts empty' ($r.body.success -and $r.body.count -eq 0)
    $r = Call POST '/api/queue-dimension' '{"type":"product","id":"NPI-QUEUE-001","length_mm":225.5,"width_mm":130,"height_mm":55,"gross_weight_kg":0.42,"net_weight_kg":0.35,"carton_gross_weight_kg":9.8,"carton_net_weight_kg":8.4}'
    Check 'queue a product' ($r.code -eq 200 -and $r.body.queued -and $r.body.pending -eq 1) $r.text
    $r = Call POST '/api/queue-dimension' '{"type":"carton","id":"CTN-QUEUE-01","name":"Queue, \"test\" carton","length_mm":400,"width_mm":300,"height_mm":250}'
    Check 'queue a carton' ($r.code -eq 200 -and $r.body.pending -eq 2) $r.text
    $r = Call POST '/api/queue-dimension' '{"type":"product","id":"npi-queue-001","length_mm":226,"width_mm":130,"height_mm":55}'
    Check 'the same item again replaces the waiting one' ($r.code -eq 200 -and $r.body.replaced -and $r.body.pending -eq 2) $r.text
    $rows = @((Call GET '/api/pending-writes').body.rows)
    $product = $rows | Where-Object { $_.Type -eq 'product' }
    $carton = $rows | Where-Object { $_.Type -eq 'carton' }
    Check 'queue holds the newer product, weights cleared' ($rows.Count -eq 2 -and $product.Length_mm -eq '226' -and $product.Gross_Weight_kg -eq '') ($rows | ConvertTo-Json -Compress)
    Check 'carton name survives commas and quotes' ($carton.Name -eq 'Queue, "test" carton') $carton.Name
    Check 'numbers are stored with a dot' ((Get-Content -LiteralPath (Join-Path $data 'pending_workbook_writes.csv') -Raw) -notmatch '225,5')
    $r = Call POST '/api/queue-dimension' '{"type":"product","id":"BAD","length_mm":226,"width_mm":130,"height_mm":55,"gross_weight_kg":0.2,"net_weight_kg":0.9}'
    Check 'net above gross -> 400' ($r.code -eq 400)
    $r = Call POST '/api/queue-dimension' '{"type":"pallet","id":"X","length_mm":226,"width_mm":130,"height_mm":55}'
    Check 'unknown type -> 400' ($r.code -eq 400)
    $r = Call POST '/api/queue-dimension' '{"type":"product","id":"X"}'
    Check 'missing dimensions -> 400' ($r.code -eq 400)
    $r = Call GET '/api/queue-dimension'
    Check 'GET on the queue route -> 405' ($r.code -eq 405)
    Check 'rejected requests added nothing' ((Call GET '/api/pending-writes').body.count -eq 2)
}
finally {
    try { Invoke-RestMethod "$base/api/shutdown" -Method Post -TimeoutSec 5 | Out-Null } catch { }
    Start-Sleep -Milliseconds 700
    if (-not $proc.HasExited) { $proc.Kill() }
}
Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) { exit 1 } else { exit 0 }
