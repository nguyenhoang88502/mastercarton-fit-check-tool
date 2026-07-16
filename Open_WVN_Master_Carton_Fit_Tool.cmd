@echo off
setlocal
pushd "%~dp0"
echo Preparing carton data and opening WVN Master Carton Fit Tool...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1" %*
if errorlevel 1 (
  echo.
  echo Launcher preparation failed. Review WVN_Master_Carton_Fit_Tool_launcher.log.
  pause
)
popd
