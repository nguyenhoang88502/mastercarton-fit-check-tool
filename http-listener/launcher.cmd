@echo off
REM ==============================================================================
REM  WVN Master Carton Fit Tool - HTTP Listener launcher
REM
REM  Double-click this file. It starts the short-lived PowerShell listener in its
REM  own window, waits until the listener actually answers, then opens the data
REM  entry UI in your default browser.
REM
REM  Nothing is installed. The listener exits on its own once the page is closed
REM  or after a short idle period, so the CSV files are never left locked.
REM ==============================================================================

setlocal enabledelayedexpansion

set "APPDIR=%~dp0"
if "%APPDIR:~-1%"=="\" set "APPDIR=%APPDIR:~0,-1%"

set "PORT=3000"
if not "%~1"=="" set "PORT=%~1"

set "URL=http://localhost:%PORT%/"

echo.
echo   WVN Master Carton Fit Tool - data entry listener
echo   ------------------------------------------------
echo   Folder : %APPDIR%
echo   Address: %URL%
echo.

REM --- make sure the working folders exist --------------------------------------
if not exist "%APPDIR%\system\generated" mkdir "%APPDIR%\system\generated"
if not exist "%APPDIR%\system\logs"      mkdir "%APPDIR%\system\logs"

REM --- refuse to start a second copy on the same port ----------------------------
netstat -ano -p TCP | findstr /R /C:"LISTENING" | findstr /C:":%PORT% " >nul
if not errorlevel 1 (
    echo   Port %PORT% is already in use - a listener is probably already running.
    echo   Opening the existing one instead.
    echo.
    start "" "%URL%"
    timeout /t 3 /nobreak >nul
    exit /b 0
)

REM --- start the listener in its own window --------------------------------------
echo   Starting listener...
start "WVN Carton HTTP Listener" powershell -NoProfile -ExecutionPolicy Bypass -File "%APPDIR%\listen.ps1" -Port %PORT%

REM --- wait until it actually answers, rather than guessing with a fixed sleep ----
set "READY="
for /L %%i in (1,1,20) do (
    if not defined READY (
        powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -Uri 'http://localhost:%PORT%/api/health' -UseBasicParsing -TimeoutSec 2; if ($r.StatusCode -eq 200) { exit 0 } else { exit 1 } } catch { exit 1 }" >nul 2>&1
        if not errorlevel 1 set "READY=1"
        if not defined READY ping -n 2 127.0.0.1 >nul
    )
)

if not defined READY (
    echo.
    echo   The listener did not come up within about 20 seconds.
    echo   Check the listener window for an error message, and see
    echo   system\logs\listener.log
    echo.
    pause
    exit /b 1
)

echo   Listener is up. Opening the browser...
start "" "%URL%"

echo.
echo   Done. Leave the listener window open while you enter data.
echo   It closes itself shortly after you close the browser tab.
echo.
timeout /t 4 /nobreak >nul
exit /b 0
