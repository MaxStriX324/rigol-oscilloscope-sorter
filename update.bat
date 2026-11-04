@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo ERROR: Windows PowerShell is required.
  pause
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\update.ps1" -Repository "MaxStriX324/rigol-oscilloscope-sorter" -Branch "main" -Installer "install.cmd"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" echo Update did not complete. Review the error above.
pause
exit /b %RC%
