$ErrorActionPreference = 'Stop'
$repo = "C:\Users\nguyenhoang88502\Desktop\MasterCarton fit check tool\WVN_Master_Carton_Fit_Tool"
$dist = Join-Path $repo 'dist'
$name = 'WVN_Master_Carton_Fit_Tool'
$stage = Join-Path $dist $name

if (Test-Path -LiteralPath $dist) { Remove-Item -LiteralPath $dist -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

# --- root files the user actually touches -------------------------------------------------
foreach ($f in @('Open_WVN_Master_Carton_Fit_Tool.cmd', 'Quick_Start.txt', 'Whats_New.txt', 'README.md', 'HANDOVER.md')) {
    Copy-Item -LiteralPath (Join-Path $repo $f) -Destination $stage
}

# --- the tool itself ------------------------------------------------------------------------
New-Item -ItemType Directory -Path (Join-Path $stage 'system\generated'), (Join-Path $stage 'system\logs') -Force | Out-Null
foreach ($f in @('launcher.ps1', 'config.ini', 'WVN_Master_Carton_Fit_Tool.html')) {
    Copy-Item -LiteralPath (Join-Path $repo "system\$f") -Destination (Join-Path $stage 'system')
}
# Ship the prepared data so the tool works on first open, before any workbook rebuild.
# -Path, not -LiteralPath: LiteralPath treats the * as part of the name and silently copies nothing.
Copy-Item -Path (Join-Path $repo 'system\generated\*') -Destination (Join-Path $stage 'system\generated') -Force
$copied = @(Get-ChildItem (Join-Path $stage 'system\generated') -File)
if ($copied.Count -lt 5) { throw "Prepared data did not copy: only $($copied.Count) file(s) in system\generated." }

# Vendored library the HTML loads at runtime via <script src="./vendor/...">, so it has to ship.
New-Item -ItemType Directory -Path (Join-Path $stage 'system\vendor\binpackingjs') -Force | Out-Null
Copy-Item -Path (Join-Path $repo 'system\vendor\binpackingjs\*') -Destination (Join-Path $stage 'system\vendor\binpackingjs') -Force

# --- the saving helper ----------------------------------------------------------------------
# Lives under system\ so it stays out of the way: the only thing anyone opening the folder
# should see is the .cmd and the two text files. Only listen.ps1 ships. The README, the data
# contract, the test suite and the build harness are all engineering material and stay in the
# repository - nobody receiving this package needs to know how the saving works.
New-Item -ItemType Directory -Path (Join-Path $stage 'system\http-listener\system\generated'), (Join-Path $stage 'system\http-listener\system\logs') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repo 'system\http-listener\listen.ps1') -Destination (Join-Path $stage 'system\http-listener')

# listen.ps1 shells out to this at runtime to write Save Product/Save Carton into the source
# Excel workbook, so unlike the rest of system\tools (DeepSeek prompts, build harness) this one
# script is a real runtime dependency and has to ship even though its neighbors do not.
New-Item -ItemType Directory -Path (Join-Path $stage 'system\tools') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repo 'system\tools\update_dimension_excel.ps1') -Destination (Join-Path $stage 'system\tools')

# Keep the empty folders in the zip so nothing has to be created by hand.
foreach ($keep in @('system\logs', 'system\http-listener\system\generated', 'system\http-listener\system\logs')) {
    Set-Content -LiteralPath (Join-Path $stage "$keep\.keep") -Value '' -Encoding ASCII
}

$zip = Join-Path $dist "$name.zip"
Compress-Archive -Path $stage -DestinationPath $zip -CompressionLevel Optimal

Write-Output "--- package contents ---"
Get-ChildItem $stage -Recurse -File | ForEach-Object { $_.FullName.Substring($stage.Length + 1) } | Sort-Object
Write-Output ""
Write-Output ("zip: {0} ({1:N0} KB)" -f $zip, ((Get-Item $zip).Length / 1KB))
