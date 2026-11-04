@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
title Rigol Sorter - One-time pass
REM Default: folder next to this launcher; optional argument overrides it.
set "RIGOL_ROOT=%~dp0"
REM set "RIGOL_ROOT=C:\RigolData\DHO814"
if not "%~1"=="" set "RIGOL_ROOT=%~1"
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" -c "import openpyxl" >nul 2>nul
    if not errorlevel 1 goto :venv_ready
)
echo Python environment is missing or broken. Repairing...
call "%~dp0install.cmd" /quiet
if errorlevel 1 (
    echo ERROR: Rigol Python environment could not be repaired.
    pause
    exit /b 3
)
:venv_ready
if not exist "%~dp0rigol_sort.py" (
    echo ERROR: rigol_sort.py not found.
    pause
    exit /b 2
)
pushd "%RIGOL_ROOT%"
if errorlevel 1 (
    echo ERROR: Cannot access folder: %RIGOL_ROOT%
    pause
    exit /b 1
)
echo Processing eligible Rigol captures in %CD%
"%~dp0.venv\Scripts\python.exe" -u "%~dp0rigol_sort.py" --min-age-seconds 60
set "RC=%ERRORLEVEL%"
popd
echo.
if "%RC%"=="0" (echo Completed.) else (echo Sorting failed. Check _logs\sort_log.txt.)
pause
exit /b %RC%
