<#
.SYNOPSIS
  Tests for the saved pallet layout routes of listen.ps1 (GET/POST /api/pallet-layouts).
.DESCRIPTION
  Starts listen.ps1 on its own port with a throwaway data folder under tests\_run, saves,
  re-reads and deliberately breaks layouts, and exits 0 when every check passes, 1 otherwise.
#>
$ErrorActionPreference = 'Stop'
$listenerPath = Join-Path (Join-Path $PSScriptRoot '..') 'listen.ps1'
$sandbox = Join-Path (Join-Path $PSScriptRoot '_run') 'pallet_layouts'
if (Test-Path $sandbox) { Remove-Item $sandbox -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $sandbox 'data'), (Join-Path $sandbox 'logs') -Force | Out-Null
$port = 3197
$proc = Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$listenerPath`"", '-Port', $port, '-IdleTimeoutSeconds', '40', '-DataPath', "`"$sandbox\data`"", '-LogPath', "`"$sandbox\logs`"")
$base = "http://localhost:$port"
$up = $false
for ($i = 0; $i -lt 40; $i++) { try { $h = Invoke-RestMethod "$base/api/health" -TimeoutSec 2; if ($h.success) { $up = $true; break } } catch { Start-Sleep -Milliseconds 300 } }
if (-not $up) { throw 'listener did not start' }

function Post($json) {
    try {
        $r = Invoke-WebRequest "$base/api/pallet-layouts" -Method Post -Body $json -ContentType 'application/json' -UseBasicParsing -TimeoutSec 15
        return @{ code = [int]$r.StatusCode; text = $r.Content }
    }
    catch {
        $resp = $_.Exception.Response
        if ($null -eq $resp) { throw }
        $st = $resp.GetResponseStream(); try { $st.Position = 0 } catch { }; $sr = New-Object IO.StreamReader($st); $t = $sr.ReadToEnd(); $sr.Dispose()
        return @{ code = [int]$resp.StatusCode; text = $t }
    }
}
$pass = 0; $fail = 0
function Check($name, $cond, $info = '') { if ($cond) { $script:pass++; Write-Output "PASS  $name" } else { $script:fail++; Write-Output "FAIL  $name  $info" } }

$good = '{"schema_version":1,"carton_id":"1008609","carton":{"length_mm":400,"width_mm":300,"height_mm":250},"pallet":{"length_mm":1200,"width_mm":1000},"ship_method":"truck","max_height_mm":2200,"origin":"suggested","source_fg_id":"1000299","levels":[[{"x":0,"y":0,"l":400,"w":300},{"x":400,"y":0,"l":400,"w":300},{"x":0,"y":300,"l":300,"w":400}],[{"x":0.5,"y":0,"l":300,"w":400}]]}'
$r = Post $good
Check 'valid save -> 200' ($r.code -eq 200) $r.text
$o = $r.text | ConvertFrom-Json
Check 'response success + not replaced' ($o.success -eq $true -and $o.replaced -eq $false)
Check 'layout echoed with counts' ($o.layout.level_count -eq 2 -and $o.layout.carton_count -eq 4 -and $o.layout.schema -eq 'wvn.pallet-layout') $r.text
Check 'levels stay nested (1-box level is still a list)' ($o.layout.levels.Count -eq 2 -and $o.layout.levels[1].Count -eq 1 -and $o.layout.levels[1][0].x -eq 0.5)
Check 'saved_at + saved_by stamped' ($o.layout.saved_at -match '^\d{4}-\d\d-\d\dT' -and $o.layout.saved_by -eq $env:USERNAME)
Check 'file written' (Test-Path (Join-Path $sandbox 'data\pallet_layouts\1008609.json'))

$r = Post $good
Check 'second save replaces' ((($r.text | ConvertFrom-Json).replaced) -eq $true)

# single level, single box (the array-unrolling trap)
$one = '{"schema_version":1,"carton_id":"ab 1/x","carton":{"length_mm":600,"width_mm":400,"height_mm":300},"pallet":{"length_mm":1200,"width_mm":1000},"ship_method":"air","max_height_mm":1600,"levels":[[{"x":0,"y":0,"l":600,"w":400}]]}'
$r = Post $one
Check 'one level / one box -> 200' ($r.code -eq 200) $r.text
Check 'odd id -> escaped file name' (Test-Path -LiteralPath (Join-Path $sandbox 'data\pallet_layouts\AB%201%2FX.json')) ((Get-ChildItem (Join-Path $sandbox 'data\pallet_layouts')).Name -join ' | ')
$o = $r.text | ConvertFrom-Json
Check 'origin defaults to manual' ($o.layout.origin -eq 'manual' -and $o.layout.carton_id -eq 'ab 1/x')

$all = Invoke-RestMethod "$base/api/pallet-layouts" -TimeoutSec 10
Check 'GET lists both' ($all.success -and $all.count -eq 2 -and $all.layouts.Count -eq 2) ($all | ConvertTo-Json -Depth 3 -Compress)
Check 'GET keeps level nesting' (($all.layouts | Where-Object { $_.carton_id -eq '1008609' }).levels[0].Count -eq 3)

function Bad($name, $json, $field) { $r = Post $json; Check $name ($r.code -eq 400 -and $r.text -match [regex]::Escape($field)) ("{0} {1}" -f $r.code, $r.text) }
Bad 'wrong schema_version' ($good -replace '"schema_version":1', '"schema_version":2') 'schema_version'
Bad 'missing carton_id' ($good -replace '"carton_id":"1008609",', '') 'carton_id'
Bad 'string dimension' ($good -replace '"height_mm":250', '"height_mm":"250"') 'carton.height_mm'
Bad 'bad ship_method' ($good -replace '"truck"', '"sea"') 'ship_method'
Bad 'empty levels' ($good -replace '"levels":\[.*\]\]', '"levels":[]') 'levels'
Bad 'empty level' ($good -replace ',\[\{"x":0.5[^\]]*\]', ',[]') 'levels[1]'
Bad 'wrong footprint' ($good -replace '"x":0.5,"y":0,"l":300,"w":400', '"x":0.5,"y":0,"l":350,"w":400') 'footprint'
Bad 'outside pallet' ($good -replace '"x":0.5,"y":0,"l":300,"w":400', '"x":950,"y":0,"l":300,"w":400') 'outside'
Bad 'overlap' ($good -replace '"x":400,"y":0,"l":400,"w":300', '"x":200,"y":0,"l":400,"w":300') 'overlaps'
Bad 'too tall' ($good -replace '"max_height_mm":2200', '"max_height_mm":400') 'taller'
Bad 'array body' ('[' + $good + ']') 'body'

# overhang allowance: a carton may hang over the edge by up to overhang_mm, and no further
$over = '{"schema_version":1,"carton_id":"OVER-1","carton":{"length_mm":400,"width_mm":300,"height_mm":250},"pallet":{"length_mm":1200,"width_mm":1000},"ship_method":"truck","max_height_mm":2200,"overhang_mm":30,"levels":[[{"x":-30,"y":-12.5,"l":400,"w":300},{"x":830,"y":725,"l":400,"w":300}]]}'
$r = Post $over
Check 'overhang within the allowance -> 200' ($r.code -eq 200 -and ($r.text | ConvertFrom-Json).layout.overhang_mm -eq 30) $r.text
Bad 'overhang beyond the allowance' ($over -replace '"x":-30', '"x":-45') 'outside'
Bad 'same layout without an allowance' ($over -replace '"overhang_mm":30,', '') 'outside'
Bad 'overhang_mm out of range' ($over -replace '"overhang_mm":30', '"overhang_mm":500') 'overhang_mm'
$r = Post '{not json'
Check 'malformed JSON -> 400' ($r.code -eq 400 -and $r.text -match 'Invalid JSON')

# a hand-broken file must not poison GET
Set-Content -LiteralPath (Join-Path $sandbox 'data\pallet_layouts\BROKEN.json') -Value '{ oops' -Encoding ASCII
$all = Invoke-RestMethod "$base/api/pallet-layouts" -TimeoutSec 10
Check 'broken file skipped' ($all.count -eq 3)

# existing routes untouched
$h = Invoke-RestMethod "$base/api/health"
Check 'health still fine' ($h.success -eq $true)

try { Invoke-RestMethod "$base/api/shutdown" -Method Post -TimeoutSec 5 | Out-Null } catch { }
Start-Sleep -Milliseconds 800
if (-not $proc.HasExited) { $proc.Kill() }
Write-Output ''
Write-Output ("{0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) { exit 1 } else { exit 0 }
