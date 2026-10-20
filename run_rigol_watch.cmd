@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
title Rigol Sorter - Watch (60 seconds)

REM 1. Default capture folder: this script's folder.
REM 2. To keep the program outside the Windows shared capture folder,
REM    replace the line below with your path, e.g. C:\RigolData\DHO814
set "RIGOL_ROOT=%~dp0"
REM set "RIGOL_ROOT=C:\RigolData\DHO814"
REM 3. Optional: run_rigol_watch.cmd "C:\RigolData\DHO924S"
if not "%~1"=="" set "RIGOL_ROOT=%~1"

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo ERROR: Python environment not installed. Run install.cmd first.
    pause
    exit /b 3
)
if not exist "%~dp0rigol_sort.py" (
    echo ERROR: rigol_sort.py is missing next to this CMD.
    pause
    exit /b 2
)
REM pushd supports local folders and UNC shares via temporary drive mapping.
pushd "%RIGOL_ROOT%"
if errorlevel 1 (
    echo ERROR: Cannot access Rigol capture folder: %RIGOL_ROOT%
    pause
    exit /b 1
)
echo ================================================
echo Rigol watcher: %CD%
echo Waiting: 60 seconds after last modification.
echo Polling: every 5 seconds. Ctrl+C to stop.
echo ================================================
"%~dp0.venv\Scripts\python.exe" -u "%~dp0rigol_sort.py" --watch --min-age-seconds 60 --poll-seconds 5
set "RC=%ERRORLEVEL%"
popd
echo.
echo Sorter stopped. Exit code %RC%.
pause
exit /b %RC%
