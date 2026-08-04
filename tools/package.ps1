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

# --- the saving helper ----------------------------------------------------------------------
# tools/ (build harness) and tests/ stay out of the package: they are engineering material,
# not something anyone opening this folder needs.
New-Item -ItemType Directory -Path (Join-Path $stage 'http-listener\system\generated'), (Join-Path $stage 'http-listener\system\logs'), (Join-Path $stage 'http-listener\docs') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repo 'http-listener\listen.ps1') -Destination (Join-Path $stage 'http-listener')
Copy-Item -LiteralPath (Join-Path $repo 'http-listener\README.md') -Destination (Join-Path $stage 'http-listener')
Copy-Item -LiteralPath (Join-Path $repo 'http-listener\docs\data-contract.md') -Destination (Join-Path $stage 'http-listener\docs')

# Keep the empty folders in the zip so nothing has to be created by hand.
foreach ($keep in @('system\logs', 'http-listener\system\generated', 'http-listener\system\logs')) {
    Set-Content -LiteralPath (Join-Path $stage "$keep\.keep") -Value '' -Encoding ASCII
}

$zip = Join-Path $dist "$name.zip"
Compress-Archive -Path $stage -DestinationPath $zip -CompressionLevel Optimal

Write-Output "--- package contents ---"
Get-ChildItem $stage -Recurse -File | ForEach-Object { $_.FullName.Substring($stage.Length + 1) } | Sort-Object
Write-Output ""
Write-Output ("zip: {0} ({1:N0} KB)" -f $zip, ((Get-Item $zip).Length / 1KB))
