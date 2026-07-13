@echo off
setlocal

rem Build a manifest from Excel workbooks beside this launcher, newest first.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$root='%~dp0'; $files=@(); if(Test-Path -LiteralPath $root){$files=Get-ChildItem -LiteralPath $root -Filter *.xlsx -File | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | ForEach-Object {[pscustomobject]@{path=$_.Name;name=$_.Name;modifiedUtc=$_.LastWriteTimeUtc.ToString('o')}}}; $json=ConvertTo-Json -Compress -Depth 3 -InputObject @($files); Set-Content -Encoding UTF8 -LiteralPath (Join-Path $root 'WVN Master Carton Workbook Manifest.json') -Value $json"

rem Find the authoritative weekly BOM folder and write absolute local file URLs, newest first.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$root=(Resolve-Path -LiteralPath '%~dp0').Path; $parent=Split-Path -Parent $root; $locations=@((Join-Path $parent '1. Weekly BOM'),(Join-Path $env:USERPROFILE 'Downloads\1. Weekly BOM'),(Join-Path $root '1. Weekly BOM')); $dir=$locations | Where-Object {Test-Path -LiteralPath $_} | Select-Object -First 1; $months=@{JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12}; $files=@(); if($dir){$files=Get-ChildItem -LiteralPath $dir -Filter *.xlsx -File | Where-Object {$_.Name -notlike '~$*'} | ForEach-Object {$period=[datetime]::MinValue; if($_.BaseName -match '(?i)\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\b'){$period=[datetime]::new($_.LastWriteTimeUtc.Year,$months[$Matches[1].ToUpperInvariant()],[int]$Matches[2],0,0,0,[DateTimeKind]::Utc)}; [pscustomobject]@{file=$_;period=$period}} | Sort-Object @{Expression='period';Descending=$true},@{Expression={$_.file.LastWriteTimeUtc};Descending=$true} | ForEach-Object {[pscustomobject]@{path=([System.Uri]$_.file.FullName).AbsoluteUri;name=$_.file.Name;modifiedUtc=$_.file.LastWriteTimeUtc.ToString('o')}}}; $json=ConvertTo-Json -Compress -Depth 3 -InputObject @($files); Set-Content -Encoding UTF8 -LiteralPath (Join-Path $root 'Weekly BOM Manifest.json') -Value $json"

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
