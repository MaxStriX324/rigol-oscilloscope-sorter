@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
title Rigol Sorter - Install Python and openpyxl

REM Install dependencies into an isolated .venv next to this script.
REM Do not require admin rights: Python is installed for current Windows user.
pushd "%~dp0"
if errorlevel 1 (
  echo ERROR: Cannot access installer directory.
  pause
  exit /b 1
)

echo ================================================
echo    RIGOL SORTER - PYTHON AND LIBRARY INSTALL
 echo ================================================
echo.

call :ensure_venv
if not errorlevel 1 goto :packages

echo Python 3 not found. Attempting Windows Package Manager...
where winget >nul 2>nul
if errorlevel 1 goto :python_org
winget install --id Python.Python.3.13 --exact --source winget --scope user --silent --accept-package-agreements --accept-source-agreements
call :ensure_venv
if not errorlevel 1 goto :packages

:python_org
echo.
echo Attempting signed 64-bit official Python installer...
REM Python 3.13.7 is a known official release; newer installed Python is preferred.
set "PYTHON_SETUP=%TEMP%\rigol-python-3.13.7-amd64.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest 'https://www.python.org/ftp/python/3.13.7/python-3.13.7-amd64.exe' -OutFile '%PYTHON_SETUP%'; if ((Get-AuthenticodeSignature '%PYTHON_SETUP%').Status -ne 'Valid') {throw 'Invalid Python installer signature'}"
if errorlevel 1 goto :failed
start /wait "" "%PYTHON_SETUP%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 Include_test=0
if errorlevel 1 goto :failed
call :ensure_venv
if errorlevel 1 goto :failed

:packages
echo.
echo Installing requirements.txt in isolated environment...
"%~dp0.venv\Scripts\python.exe" -m pip install --disable-pip-version-check --upgrade pip
if errorlevel 1 goto :failed
"%~dp0.venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 goto :failed
"%~dp0.venv\Scripts\python.exe" -c "import sys,openpyxl;print('Python:',sys.version.split()[0]);print('openpyxl:',openpyxl.__version__)"
if errorlevel 1 goto :failed
echo.
echo INSTALLATION COMPLETE!
echo Start run_rigol_watch.cmd for automatic sorting.
popd
pause
exit /b 0

:ensure_venv
if exist "%~dp0.venv\Scripts\python.exe" exit /b 0
if exist "%LocalAppData%\Programs\Python\Python314\python.exe" (
  "%LocalAppData%\Programs\Python\Python314\python.exe" -m venv "%~dp0.venv" >nul 2>nul
  if not errorlevel 1 exit /b 0
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
  "%LocalAppData%\Programs\Python\Python313\python.exe" -m venv "%~dp0.venv" >nul 2>nul
  if not errorlevel 1 exit /b 0
)
py -3 -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>nul
if not errorlevel 1 (
  py -3 -m venv "%~dp0.venv" >nul 2>nul
  if not errorlevel 1 exit /b 0
)
python -c "import sys; assert sys.version_info >= (3, 10)" >nul 2>nul
if not errorlevel 1 (
  python -m venv "%~dp0.venv" >nul 2>nul
  if not errorlevel 1 exit /b 0
)
exit /b 1

:failed
echo.
echo INSTALLATION FAILED. Check network access and try again.
echo You can also install Python manually from:
echo https://www.python.org/downloads/windows/
echo Then rerun install.cmd.
popd
pause
exit /b 1
