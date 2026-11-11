@echo off
setlocal EnableExtensions
rem Survey launcher - Administrator required
rem Works with native Windows PowerShell 2.0+ (Win7, Win8/8.1, Win10/11).
rem Set 0 below for slower PCs; keeps all data-export behavior unchanged.
set "AUDIT_PARALLEL=1"
rem Gioi han buoc LICENSE khi chay song song (giay). Mac dinh 40.
set "AUDIT_LICENSE_TIMEOUT_SEC=40"
rem Gioi han rieng truy van Windows WMI (30 giay) - SLMGR fallback 12s.
set "AUDIT_WINDOWS_TIMEOUT_SEC=30"
rem Console font height in pixels (20/22/24/26/28). 0=keep current font.
set "AUDIT_FONT_PX=26"
cd /d "%~dp0"
title Device Audit Tool
color 0B
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
rem Prefer 64-bit host if invoked from a 32-bit shell on 64-bit Windows.
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" (
  color 0C
  echo [ERROR] Windows PowerShell not found. Contact IT.
  pause
  exit /b 2
)
rem Hard gate before Survey.ps1. No UAC bypass or silent elevation.
"%PS%" -NoProfile -NonInteractive -Command "$id=[System.Security.Principal.WindowsIdentity]::GetCurrent();$p=New-Object System.Security.Principal.WindowsPrincipal($id);if($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)){exit 0}else{exit 5}" >nul 2>&1
if errorlevel 1 (
  color 0E
  echo.
  echo ============================================================
  echo  DEVICE AUDIT TOOL
  echo ============================================================
  echo  Yeu cau chay bang quyen Administrator.
  echo  KHONG co buoc khao sat nao duoc thuc hien.
  echo.
  echo  Huong dan: Nhan chuot phai 1_Run_Survey.bat
  echo             Chon "Run as administrator".
  echo  Neu may yeu cau mat khau, lien he bo phan IT.
  echo ============================================================
  echo.
  pause
  exit /b 5
)
echo.
echo.
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Survey.ps1"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  color 0E
  echo.
  echo [WARN] Survey did not finish normally. Exit code: %RC%
  echo [INFO] Check KhaoSat\Log if a log was created.
)
echo.
pause
exit /b %RC%
