@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Device Audit Tool - Win7 Compatibility Check
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" (
  echo [FAIL] Khong tim thay Windows PowerShell.
  pause
  exit /b 2
)
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Check_Win7_Compatibility.ps1"
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
