@echo off
setlocal
set "ARGS=%*"
echo Preparing carton data and opening WVN Master Carton Fit Tool...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1" %ARGS%
