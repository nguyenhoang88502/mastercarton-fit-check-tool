@echo off
setlocal EnableDelayedExpansion

set "TOOL_DIR=%~dp0"
set "TOOL_DIR=%TOOL_DIR:~0,-1%"
set "HTML_FILE=%TOOL_DIR%\WVN_Master_Carton_Fit_Tool.html"

echo ========================================
echo  WVN Master Carton Fit Tool Launcher
echo ========================================
echo Tool directory: %TOOL_DIR%

rem ---------------------------------------------------------------------------
rem 1. Dimension workbook manifest
rem ---------------------------------------------------------------------------
echo.
echo [Dimension workbook search]
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$toolDir='%TOOL_DIR%';" ^
  "$namePrefs=@(" ^
  "  @{Pattern='Products Physical Dimension - FG *.xlsx';       Label='preferred dimension workbook'}," ^
  "  @{Pattern='*Products Physical Dimension*.xlsx';             Label='Products Physical Dimension variant'}," ^
  "  @{Pattern='*Physical Dimension*.xlsx';                      Label='physical dimension workbook'}," ^
  "  @{Pattern='*Dimension - FG*.xlsx';                          Label='dimension FG workbook'}," ^
  "  @{Pattern='carton_tool_input.xlsx';                         Label='legacy fallback'}," ^
  "  @{Pattern='*.xlsx';                                         Label='any workbook (last resort)'}" ^
  ");" ^
  "$locations=@();" ^
  "if(Test-Path $toolDir){$locations+=@($toolDir)};" ^
  "if(Test-Path (Join-Path $toolDir '1. Weekly BOM')){$locations+=@(Join-Path $toolDir '1. Weekly BOM')};" ^
  "$onedrive=$env:OneDriveCommercial;" ^
  "if(!$onedrive){$onedrive=$env:OneDrive};" ^
  "if(!$onedrive){$onedrive=Join-Path $env:USERPROFILE 'OneDrive - Wahl Clipper Corporation'};" ^
  "if($onedrive -and (Test-Path $onedrive)){" ^
  "  $bomPath=Join-Path $onedrive 'NPI Information - Documents\General\Project Management\5. Tools for NPI\1. Weekly BOM';" ^
  "  if(Test-Path $bomPath){$locations+=@($bomPath)};" ^
  "  $toolPath=Join-Path $onedrive 'NPI Information - Documents\General\Project Management\5. Tools for NPI\MasterCarton fit check tool';" ^
  "  if(Test-Path $toolPath){$locations+=@($toolPath)}" ^
  "};" ^
  "if(Test-Path (Join-Path $env:USERPROFILE 'Downloads')){$locations+=@(Join-Path $env:USERPROFILE 'Downloads')};" ^
  "if(Test-Path (Join-Path $env:USERPROFILE 'Desktop')){$locations+=@(Join-Path $env:USERPROFILE 'Desktop')};" ^
  "$results=@();" ^
  "foreach($loc in $locations){" ^
  "  if(!(Test-Path $loc)){continue};" ^
  "  Write-Host ('  Scanning: '+$loc);" ^
  "  foreach($entry in $namePrefs){" ^
  "    $found=Get-ChildItem -LiteralPath $loc -Filter $entry.Pattern -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1;" ^
  "    if($found){" ^
  "      $uri=([System.Uri]$found.FullName).AbsoluteUri;" ^
  "      Write-Host ('    MATCHED: '+$found.Name+' ('+$entry.Label+')');" ^
  "      $results+=@([pscustomobject]@{path=$uri;name=$found.Name;modifiedUtc=$found.LastWriteTimeUtc.ToString('o')});" ^
  "      break" ^
  "    }" ^
  "  };" ^
  "  if($results.Count -gt 0){break}" ^
  "};" ^
  "if($results.Count -eq 0){" ^
  "  Write-Host '    WARNING: No dimension workbook found in any search location.';" ^
  "  Write-Host '    You can still select or drag-drop a workbook manually in the browser.'" ^
  "} else {" ^
  "  Write-Host ('    Selected: '+$results[0].name)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 3 -InputObject @($results);" ^
  "$manifestPath=Join-Path $toolDir 'WVN Master Carton Workbook Manifest.json';" ^
  "Set-Content -Encoding UTF8 -LiteralPath $manifestPath -Value $json;" ^
  "Write-Host ('  Manifest written: '+$manifestPath)"

rem ---------------------------------------------------------------------------
rem 2. Weekly BOM manifest
rem ---------------------------------------------------------------------------
echo.
echo [Weekly BOM search]
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$toolDir='%TOOL_DIR%';" ^
  "$locations=@();" ^
  "if(Test-Path (Join-Path $toolDir '1. Weekly BOM')){$locations+=@(Join-Path $toolDir '1. Weekly BOM')};" ^
  "$parent=Split-Path -Parent $toolDir;" ^
  "$siblingBom=Join-Path $parent '1. Weekly BOM';" ^
  "if(Test-Path $siblingBom){$locations+=@($siblingBom)};" ^
  "$onedrive=$env:OneDriveCommercial;" ^
  "if(!$onedrive){$onedrive=$env:OneDrive};" ^
  "if(!$onedrive){$onedrive=Join-Path $env:USERPROFILE 'OneDrive - Wahl Clipper Corporation'};" ^
  "if($onedrive -and (Test-Path $onedrive)){" ^
  "  $bomPath=Join-Path $onedrive 'NPI Information - Documents\General\Project Management\5. Tools for NPI\1. Weekly BOM';" ^
  "  if(Test-Path $bomPath){$locations+=@($bomPath)}" ^
  "};" ^
  "$downloadsBom=Join-Path $env:USERPROFILE 'Downloads\1. Weekly BOM';" ^
  "if(Test-Path $downloadsBom){$locations+=@($downloadsBom)};" ^
  "$desktopBom=Join-Path $env:USERPROFILE 'Desktop\1. Weekly BOM';" ^
  "if(Test-Path $desktopBom){$locations+=@($desktopBom)};" ^
  "if(Test-Path $toolDir){$locations+=@($toolDir)};" ^
  "$months=@{JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12};" ^
  "$files=@();" ^
  "foreach($loc in $locations){" ^
  "  if(!(Test-Path $loc)){continue};" ^
  "  Write-Host ('  Scanning: '+$loc);" ^
  "  $inDir=Get-ChildItem -LiteralPath $loc -Filter '*.xlsx' -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*' -and $_.Name -like '*BOM*'};" ^
  "  foreach($f in $inDir){" ^
  "    $period=[datetime]::MinValue;" ^
  "    if($f.BaseName -match '(?i)\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\b'){" ^
  "      $period=[datetime]::new($f.LastWriteTimeUtc.Year,$months[$Matches[1].ToUpperInvariant()],[int]$Matches[2],0,0,0,[DateTimeKind]::Utc)" ^
  "    };" ^
  "    $files+=@([pscustomobject]@{path=([System.Uri]$f.FullName).AbsoluteUri;name=$f.Name;modifiedUtc=$f.LastWriteTimeUtc.ToString('o');period=$period})" ^
  "  }" ^
  "};" ^
  "$files=$files | Sort-Object @{Expression='period';Descending=$true},@{Expression='modifiedUtc';Descending=$true} | ForEach-Object {[pscustomobject]@{path=$_.path;name=$_.name;modifiedUtc=$_.modifiedUtc}};" ^
  "if($files.Count -eq 0){" ^
  "  Write-Host '    WARNING: No Weekly BOM files found in any search location.';" ^
  "  Write-Host '    You can still open a BOM manually in the browser (Workbook panel).'" ^
  "} else {" ^
  "  Write-Host ('    Selected latest BOM: '+$files[0].name)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 3 -InputObject @($files);" ^
  "$manifestPath=Join-Path $toolDir 'Weekly BOM Manifest.json';" ^
  "Set-Content -Encoding UTF8 -LiteralPath $manifestPath -Value $json;" ^
  "Write-Host ('  Manifest written: '+$manifestPath)"

rem ---------------------------------------------------------------------------
rem 3. Launch browser
rem ---------------------------------------------------------------------------
if not exist "%HTML_FILE%" (
  echo.
  echo ERROR: WVN_Master_Carton_Fit_Tool.html was not found beside this launcher.
  pause
  exit /b 1
)

echo.
echo [Browser]
set "CHROME="

rem Check Chrome via LOCALAPPDATA (user install, most common)
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
)
rem Check Chrome via Program Files (system install)
if not defined CHROME if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
)
if not defined CHROME if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
)

if defined CHROME (
  echo   Selected: Chrome ^(%CHROME%^)
  start "" "!CHROME!" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%HTML_FILE%"
  exit /b
)

rem Fallback to Edge
set "EDGE="
if exist "%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe" (
  set "EDGE=%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"
)
if not defined EDGE if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
  set "EDGE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
)
if not defined EDGE if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
  set "EDGE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
)

if defined EDGE (
  echo   Selected: Edge ^(!EDGE!^)
  start "" "!EDGE!" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%HTML_FILE%"
  exit /b
)

echo   ERROR: Neither Chrome nor Edge was found.
echo   Checked: LOCALAPPDATA, ProgramFiles, ProgramFiles(x86^)
pause
exit /b 1
