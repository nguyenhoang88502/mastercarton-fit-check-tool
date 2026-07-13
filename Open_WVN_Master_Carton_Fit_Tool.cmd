@echo off
setlocal

rem Build a manifest from Excel workbooks beside this launcher, newest first.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$root='%~dp0'; $files=@(); if(Test-Path -LiteralPath $root){$files=Get-ChildItem -LiteralPath $root -Filter *.xlsx -File | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | ForEach-Object {[pscustomobject]@{path=$_.Name;name=$_.Name;modifiedUtc=$_.LastWriteTimeUtc.ToString('o')}}}; $json=ConvertTo-Json -Compress -Depth 3 -InputObject @($files); Set-Content -Encoding UTF8 -LiteralPath (Join-Path $root 'WVN Master Carton Workbook Manifest.json') -Value $json"

rem Find the weekly BOM folder (sibling first, then Downloads, then local) and write newest-first manifest entries.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$root=(Resolve-Path -LiteralPath '%~dp0').Path; $parent=Split-Path -Parent $root; $downloads=Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\1. Weekly BOM'; $dirs=@((Join-Path $parent '1. Weekly BOM'),$downloads,(Join-Path $root '1. Weekly BOM')); $dir=$dirs | Where-Object {Test-Path -LiteralPath $_} | Select-Object -First 1; $files=@(); if($dir){$files=Get-ChildItem -LiteralPath $dir -Filter *.xlsx -File | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | ForEach-Object {[pscustomobject]@{path=([System.Uri]::new($_.FullName).AbsoluteUri);name=$_.Name;modifiedUtc=$_.LastWriteTimeUtc.ToString('o')}}}; $json=ConvertTo-Json -Compress -Depth 3 -InputObject @($files); Set-Content -Encoding UTF8 -LiteralPath (Join-Path $root 'Weekly BOM Manifest.json') -Value $json"

if not exist "%~dp0WVN_Master_Carton_Fit_Tool.html" (
  echo WVN_Master_Carton_Fit_Tool.html was not found beside this launcher.
  pause
  exit /b 1
)

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
  start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%~dp0WVN_Master_Carton_Fit_Tool.html"
  exit /b
)
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
  start "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%~dp0WVN_Master_Carton_Fit_Tool.html"
  exit /b
)
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
  start "" "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%~dp0WVN_Master_Carton_Fit_Tool.html"
  exit /b
)
if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
  start "" "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%~dp0WVN_Master_Carton_Fit_Tool.html"
  exit /b
)

echo Chrome or Edge was not found.
pause
