@echo off
setlocal EnableDelayedExpansion

set "TOOL_DIR=%~dp0"
set "TOOL_DIR=%TOOL_DIR:~0,-1%"
set "HTML_FILE=%TOOL_DIR%\WVN_Master_Carton_Fit_Tool.html"

rem ---------------------------------------------------------------------------
rem Detect user root from %USERPROFILE%
rem ---------------------------------------------------------------------------
set "USER_ROOT=%USERPROFILE%"

echo ========================================
echo  WVN Master Carton Fit Tool Launcher
echo ========================================
echo TOOL_DIR      : %TOOL_DIR%
echo USER_ROOT     : %USER_ROOT%

rem ---------------------------------------------------------------------------
rem Build fixed company NPI paths (username is the only variable)
rem ---------------------------------------------------------------------------
set "NPI_TOOLS_DIR=%USER_ROOT%\OneDrive - Wahl Clipper Corporation (1)\NPI Information - Documents\General\Project Management\5. Tools for NPI"
set "WEEKLY_BOM_DIR=%NPI_TOOLS_DIR%\1. Weekly BOM"
set "MASTER_CARTON_TOOL_DIR=%NPI_TOOLS_DIR%\MasterCarton fit check tool"

echo NPI_TOOLS_DIR : %NPI_TOOLS_DIR%
echo WEEKLY_BOM_DIR exists:
if exist "%WEEKLY_BOM_DIR%" (echo   YES - %WEEKLY_BOM_DIR%) else (echo   NO  - %WEEKLY_BOM_DIR%)
echo MASTER_CARTON_TOOL_DIR exists:
if exist "%MASTER_CARTON_TOOL_DIR%" (echo   YES - %MASTER_CARTON_TOOL_DIR%) else (echo   NO  - %MASTER_CARTON_TOOL_DIR%)

rem ---------------------------------------------------------------------------
rem 1. Dimension workbook manifest
rem ---------------------------------------------------------------------------
echo.
echo ========================================
echo  Dimension workbook search
echo ========================================
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$toolDir='%TOOL_DIR%';" ^
  "$weeklyBomDir='%WEEKLY_BOM_DIR%';" ^
  "$masterCartonToolDir='%MASTER_CARTON_TOOL_DIR%';" ^
  "$namePrefs=@(" ^
  "  @{Pattern='Products Physical Dimension - FG 1.Apr.25.xlsx'; Label='preferred dimension workbook'}," ^
  "  @{Pattern='*Products Physical Dimension - FG *.xlsx';    Label='Products Physical Dimension (date variant)'}," ^
  "  @{Pattern='*Products Physical Dimension*.xlsx';           Label='Products Physical Dimension variant'}," ^
  "  @{Pattern='*Physical Dimension*.xlsx';                    Label='physical dimension workbook'}," ^
  "  @{Pattern='*Dimension*FG*.xlsx';                          Label='dimension FG workbook'}," ^
  "  @{Pattern='carton_tool_input.xlsx';                       Label='legacy fallback'}" ^
  ");" ^
  "$locations=@(" ^
  "  @{Path=$weeklyBomDir;          Label='Weekly BOM folder'}," ^
  "  @{Path=$masterCartonToolDir;   Label='MasterCarton tool folder'}," ^
  "  @{Path=$toolDir;              Label='tool directory'}" ^
  ");" ^
  "$results=@();" ^
  "foreach($loc in $locations){" ^
  "  if(!(Test-Path $loc.Path)){ Write-Host ('  SKIP (not found): '+$loc.Path+' ('+$loc.Label+')'); continue };" ^
  "  Write-Host ('  Scanning: '+$loc.Path+' ('+$loc.Label+')');" ^
  "  foreach($pref in $namePrefs){" ^
  "    $found=Get-ChildItem -LiteralPath $loc.Path -Filter $pref.Pattern -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1;" ^
  "    if($found){" ^
  "      $uri=([System.Uri]$found.FullName).AbsoluteUri;" ^
  "      $entry=@{name=$found.Name;path=$found.FullName;url=$uri;modifiedUtc=$found.LastWriteTimeUtc.ToString('o');sourceType='dimension'};" ^
  "      Write-Host ('    MATCHED: '+$found.Name+' ('+$pref.Label+')');" ^
  "      Write-Host ('    path : '+$found.FullName);" ^
  "      Write-Host ('    url  : '+$uri);" ^
  "      $results+=@($entry);" ^
  "      break" ^
  "    }" ^
  "  };" ^
  "  if($results.Count -gt 0){break}" ^
  "};" ^
  "if($results.Count -eq 0){" ^
  "  Write-Host '    WARNING: No dimension workbook found in any search location.';" ^
  "  Write-Host '    You can still select or drag-drop a workbook manually in the browser.'" ^
  "} else {" ^
  "  Write-Host ('  Selected dimension workbook: '+$results[0].name)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 4 -InputObject @($results);" ^
  "$manifestPath=Join-Path $toolDir 'WVN Master Carton Workbook Manifest.json';" ^
  "Set-Content -Encoding UTF8 -LiteralPath $manifestPath -Value $json;" ^
  "Write-Host ('  Manifest: '+$manifestPath)"

rem ---------------------------------------------------------------------------
rem 2. Weekly BOM manifest
rem ---------------------------------------------------------------------------
echo.
echo ========================================
echo  Weekly BOM search
echo ========================================
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$toolDir='%TOOL_DIR%';" ^
  "$weeklyBomDir='%WEEKLY_BOM_DIR%';" ^
  "$locations=@(" ^
  "  @{Path=$weeklyBomDir;                Label='Weekly BOM folder'}," ^
  "  @{Path=(Join-Path $toolDir '1. Weekly BOM'); Label='local 1. Weekly BOM'}," ^
  "  @{Path=$toolDir;                    Label='tool directory'}" ^
  ");" ^
  "$months=@{JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12};" ^
  "$files=@();" ^
  "foreach($loc in $locations){" ^
  "  if(!(Test-Path $loc.Path)){ Write-Host ('  SKIP (not found): '+$loc.Path+' ('+$loc.Label+')'); continue };" ^
  "  Write-Host ('  Scanning: '+$loc.Path+' ('+$loc.Label+')');" ^
  "  $inDir=Get-ChildItem -LiteralPath $loc.Path -Filter 'BOM ^& FORMULA * D365.xlsx' -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*'};" ^
  "  foreach($f in $inDir){" ^
  "    $period=[datetime]::MinValue;" ^
  "    if($f.BaseName -match '(?i)\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\b'){" ^
  "      $period=[datetime]::new($f.LastWriteTimeUtc.Year,$months[$Matches[1].ToUpperInvariant()],[int]$Matches[2],0,0,0,[DateTimeKind]::Utc)" ^
  "    };" ^
  "    $uri=([System.Uri]$f.FullName).AbsoluteUri;" ^
  "    $files+=@(@{name=$f.Name;path=$f.FullName;url=$uri;modifiedUtc=$f.LastWriteTimeUtc.ToString('o');sourceType='weekly_bom';period=$period})" ^
  "  }" ^
  "};" ^
  "$files=$files | Sort-Object @{Expression='period';Descending=$true},@{Expression='modifiedUtc';Descending=$true} | ForEach-Object {[pscustomobject]@{name=$_.name;path=$_.path;url=$_.url;modifiedUtc=$_.modifiedUtc;sourceType=$_.sourceType}};" ^
  "if($files.Count -eq 0){" ^
  "  Write-Host '  WARNING: No Weekly BOM files found.';" ^
  "  Write-Host '  You can still open a BOM manually in the browser (Workbook panel).'" ^
  "} else {" ^
  "  Write-Host ('  Latest BOM (by period): '+$files[0].name);" ^
  "  Write-Host ('  path : '+$files[0].path);" ^
  "  Write-Host ('  url  : '+$files[0].url);" ^
  "  Write-Host ('  Total BOM candidates  : '+$files.Count)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 4 -InputObject @($files);" ^
  "$manifestPath=Join-Path $toolDir 'Weekly BOM Manifest.json';" ^
  "Set-Content -Encoding UTF8 -LiteralPath $manifestPath -Value $json;" ^
  "Write-Host ('  Manifest: '+$manifestPath)"

rem ---------------------------------------------------------------------------
rem 3. Launch browser
rem ---------------------------------------------------------------------------
if not exist "%HTML_FILE%" (
  echo.
  echo ========================================
  echo  ERROR: HTML file not found
  echo ========================================
  echo %HTML_FILE%
  pause
  exit /b 1
)

echo.
echo ========================================
echo  Browser launch
echo ========================================
set "CHROME="

if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
)
if not defined CHROME if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
)
if not defined CHROME if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
  set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
)

if defined CHROME (
  echo   Browser: Chrome
  start "" "!CHROME!" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%HTML_FILE%"
  exit /b
)

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
  echo   Browser: Edge
  start "" "!EDGE!" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%HTML_FILE%"
  exit /b
)

echo   ERROR: Neither Chrome nor Edge was found.
echo   Checked: LOCALAPPDATA, ProgramFiles, ProgramFiles(x86)
pause
exit /b 1
