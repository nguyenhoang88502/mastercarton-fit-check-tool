@echo off
setlocal

set "APP_ROOT=%~dp0"
set "LAUNCHER=%APP_ROOT%system\launcher.ps1"

if not exist "%LAUNCHER%" (
    echo ERROR: launcher.ps1 was not found.
    echo Expected:
    echo %LAUNCHER%
    pause
    exit /b 1
)

echo Preparing carton data and opening WVN Master Carton Fit Tool...
powershell.exe ^
  -NoProfile ^
  -ExecutionPolicy Bypass ^
  -File "%LAUNCHER%" %*

set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Launcher preparation failed. Review the newest per-run log in system\logs.
    pause
)

exit /b %EXIT_CODE%
