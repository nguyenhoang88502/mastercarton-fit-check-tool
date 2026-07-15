@echo off
setlocal
set "ARGS=%*"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1" %ARGS%
