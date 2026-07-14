@echo off
setlocal EnableDelayedExpansion

set "TOOL_DIR=%~dp0"
set "TOOL_DIR=%TOOL_DIR:~0,-1%"
set "HTML_FILE=%TOOL_DIR%\WVN_Master_Carton_Fit_Tool.html"

echo ========================================
echo  WVN Master Carton Fit Tool Launcher
echo ========================================
echo Tool directory: %TOOL_DIR%
echo HTML: %HTML_FILE%

rem ---------------------------------------------------------------------------
rem Discover OneDrive folders (may have suffix (1), (2), etc.)
rem ---------------------------------------------------------------------------
echo.
echo [OneDrive discovery]
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$matches=@(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -like 'OneDrive - Wahl Clipper Corporation*'} | ForEach-Object {$_.FullName});" ^
  "if($matches.Count -gt 0){foreach($m in $matches){Write-Host ('  Found: '+$m)}}else{Write-Host '  WARNING: No OneDrive - Wahl Clipper Corporation folder found.'}"

rem ---------------------------------------------------------------------------
rem 1. Dimension workbook manifest
rem ---------------------------------------------------------------------------
echo.
echo [Dimension workbook search]
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$toolDir='%TOOL_DIR%';" ^
  "$namePrefs=@(" ^
  "  @{Pattern='Products Physical Dimension - FG 1.Apr.25.xlsx'; Label='preferred dimension workbook'}," ^
  "  @{Pattern='*Products Physical Dimension - FG *.xlsx';    Label='Products Physical Dimension (date variant)'}," ^
  "  @{Pattern='*Products Physical Dimension*.xlsx';           Label='Products Physical Dimension variant'}," ^
  "  @{Pattern='*Physical Dimension*.xlsx';                    Label='physical dimension workbook'}," ^
  "  @{Pattern='*Dimension*FG*.xlsx';                          Label='dimension FG workbook'}," ^
  "  @{Pattern='carton_tool_input.xlsx';                       Label='legacy fallback'}" ^
  ");" ^
  "$locations=@();" ^
  "if(Test-Path $toolDir){$locations+=@{Path=$toolDir;Label='tool directory'}};" ^
  "$localBom=Join-Path $toolDir '1. Weekly BOM';" ^
  "if(Test-Path $localBom){$locations+=@{Path=$localBom;Label='local 1. Weekly BOM'}};" ^
  "$odMatches=@(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -like 'OneDrive - Wahl Clipper Corporation*'});" ^
  "foreach($od in $odMatches){" ^
  "  $bomPath=Join-Path $od.FullName 'NPI Information - Documents\General\Project Management\5. Tools for NPI\1. Weekly BOM';" ^
  "  if(Test-Path $bomPath){$locations+=@{Path=$bomPath;Label=('OneDrive NPI Weekly BOM ('+$od.Name+')')}};" ^
  "  $toolPath=Join-Path $od.FullName 'NPI Information - Documents\General\Project Management\5. Tools for NPI\MasterCarton fit check tool';" ^
  "  if(Test-Path $toolPath){$locations+=@{Path=$toolPath;Label=('OneDrive NPI tool folder ('+$od.Name+')')}}" ^
  "};" ^
  "if(Test-Path (Join-Path $env:USERPROFILE 'Downloads')){$locations+=@{Path=(Join-Path $env:USERPROFILE 'Downloads');Label='Downloads'}};" ^
  "if(Test-Path (Join-Path $env:USERPROFILE 'Desktop')){$locations+=@{Path=(Join-Path $env:USERPROFILE 'Desktop');Label='Desktop'}};" ^
  "$results=@();" ^
  "foreach($loc in $locations){" ^
  "  Write-Host ('  Scanning: '+$loc.Path+' ('+$loc.Label+')');" ^
  "  foreach($pref in $namePrefs){" ^
  "    $found=Get-ChildItem -LiteralPath $loc.Path -Filter $pref.Pattern -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*'} | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1;" ^
  "    if($found){" ^
  "      $uri=([System.Uri]$found.FullName).AbsoluteUri;" ^
  "      $entry=@{name=$found.Name;path=$found.FullName;url=$uri;lastWriteTimeUtc=$found.LastWriteTimeUtc.ToString('o');sourceType='dimension'};" ^
  "      Write-Host ('    MATCHED: '+$found.Name+' ('+$pref.Label+')');" ^
  "      Write-Host ('    path: '+$found.FullName);" ^
  "      Write-Host ('    url : '+$uri);" ^
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
  "  Write-Host ('    Selected: '+$results[0].name)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 4 -InputObject @($results);" ^
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
  "$localBom=Join-Path $toolDir '1. Weekly BOM';" ^
  "if(Test-Path $localBom){$locations+=@{Path=$localBom;Label='local 1. Weekly BOM'}};" ^
  "$odMatches=@(Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -like 'OneDrive - Wahl Clipper Corporation*'});" ^
  "foreach($od in $odMatches){" ^
  "  $bomPath=Join-Path $od.FullName 'NPI Information - Documents\General\Project Management\5. Tools for NPI\1. Weekly BOM';" ^
  "  if(Test-Path $bomPath){$locations+=@{Path=$bomPath;Label=('OneDrive NPI Weekly BOM ('+$od.Name+')')}}" ^
  "};" ^
  "$downloadsBom=Join-Path $env:USERPROFILE 'Downloads\1. Weekly BOM';" ^
  "if(Test-Path $downloadsBom){$locations+=@{Path=$downloadsBom;Label='Downloads\1. Weekly BOM'}};" ^
  "$desktopBom=Join-Path $env:USERPROFILE 'Desktop\1. Weekly BOM';" ^
  "if(Test-Path $desktopBom){$locations+=@{Path=$desktopBom;Label='Desktop\1. Weekly BOM'}};" ^
  "if(Test-Path $toolDir){$locations+=@{Path=$toolDir;Label='tool directory'}};" ^
  "$months=@{JAN=1;FEB=2;MAR=3;APR=4;MAY=5;JUN=6;JUL=7;AUG=8;SEP=9;OCT=10;NOV=11;DEC=12};" ^
  "$files=@();" ^
  "foreach($loc in $locations){" ^
  "  Write-Host ('  Scanning: '+$loc.Path+' ('+$loc.Label+')');" ^
  "  $inDir=Get-ChildItem -LiteralPath $loc.Path -Filter '*.xlsx' -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -notlike '~$*' -and $_.Name -like '*BOM*'};" ^
  "  foreach($f in $inDir){" ^
  "    $period=[datetime]::MinValue;" ^
  "    if($f.BaseName -match '(?i)\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{1,2})\b'){" ^
  "      $period=[datetime]::new($f.LastWriteTimeUtc.Year,$months[$Matches[1].ToUpperInvariant()],[int]$Matches[2],0,0,0,[DateTimeKind]::Utc)" ^
  "    };" ^
  "    $uri=([System.Uri]$f.FullName).AbsoluteUri;" ^
  "    $files+=@(@{name=$f.Name;path=$f.FullName;url=$uri;lastWriteTimeUtc=$f.LastWriteTimeUtc.ToString('o');sourceType='weekly_bom';period=$period})" ^
  "  }" ^
  "};" ^
  "$files=$files | Sort-Object @{Expression='period';Descending=$true},@{Expression='lastWriteTimeUtc';Descending=$true} | ForEach-Object {[pscustomobject]@{name=$_.name;path=$_.path;url=$_.url;lastWriteTimeUtc=$_.lastWriteTimeUtc;sourceType=$_.sourceType}};" ^
  "if($files.Count -eq 0){" ^
  "  Write-Host '    WARNING: No Weekly BOM files found in any search location.';" ^
  "  Write-Host '    You can still open a BOM manually in the browser (Workbook panel).'" ^
  "} else {" ^
  "  Write-Host ('    Selected latest BOM: '+$files[0].name);" ^
  "  Write-Host ('    path: '+$files[0].path);" ^
  "  Write-Host ('    url : '+$files[0].url)" ^
  "};" ^
  "$json=ConvertTo-Json -Compress -Depth 4 -InputObject @($files);" ^
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
  echo   Selected: Chrome ^(!CHROME!^)
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
  echo   Selected: Edge ^(!EDGE!^)
  start "" "!EDGE!" --user-data-dir="%TEMP%\WVN-Master-Carton-Tool-%RANDOM%" --allow-file-access-from-files "%HTML_FILE%"
  exit /b
)

echo   ERROR: Neither Chrome nor Edge was found.
echo   Checked: LOCALAPPDATA, ProgramFiles, ProgramFiles(x86^)
pause
exit /b 1
