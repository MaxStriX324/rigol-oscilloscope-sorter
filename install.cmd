@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "VENV=%~dp0.venv"
set "VENV_PY=%VENV%\Scripts\python.exe"
set "NO_PAUSE=0"
if /i "%~1"=="/quiet" set "NO_PAUSE=1"
title Rigol Sorter - Install Python and openpyxl
pushd "%~dp0"
if errorlevel 1 goto :failed_folder

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
set "PYTHON_SETUP=%TEMP%\rigol-python-3.13.7-amd64.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest 'https://www.python.org/ftp/python/3.13.7/python-3.13.7-amd64.exe' -OutFile '%PYTHON_SETUP%'; if ((Get-AuthenticodeSignature '%PYTHON_SETUP%').Status -ne 'Valid') {throw 'Invalid Python installer signature'}"
if errorlevel 1 goto :failed_python
start /wait "" "%PYTHON_SETUP%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1 Include_test=0
if errorlevel 1 goto :failed_python
call :ensure_venv
if errorlevel 1 goto :failed_python

:packages
if not exist "%~dp0requirements.txt" goto :failed_requirements
echo Checking required Python packages...
"%VENV_PY%" -m pip --version >nul 2>nul
if errorlevel 1 (
  "%VENV_PY%" -m ensurepip --upgrade
  if errorlevel 1 goto :failed_packages
)
"%VENV_PY%" -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 goto :failed_packages
"%VENV_PY%" -c "import sys,openpyxl;print('Python:',sys.version.split()[0]);print('openpyxl:',openpyxl.__version__)"
if errorlevel 1 goto :failed_packages
echo.
echo INSTALLATION COMPLETE! Run run_rigol_watch.cmd.
goto :success

:ensure_venv
if exist "%VENV_PY%" (
  "%VENV_PY%" -c "import sys; sys.exit(0 if sys.prefix != sys.base_prefix and sys.version_info >= (3, 10) else 1)" >nul 2>nul
  if not errorlevel 1 exit /b 0
)
if exist "%VENV%" (
  echo Broken or partial Python environment found. Recreating it...
  rmdir /s /q "%VENV%"
  if exist "%VENV%" exit /b 1
)
call :try_python py -3.14
if not errorlevel 1 exit /b 0
call :try_python py -3.13
if not errorlevel 1 exit /b 0
call :try_python py -3.12
if not errorlevel 1 exit /b 0
call :try_python py -3
if not errorlevel 1 exit /b 0
call :try_python python
if not errorlevel 1 exit /b 0
if exist "%LocalAppData%\Programs\Python\Python314\python.exe" (
  call :try_python "%LocalAppData%\Programs\Python\Python314\python.exe"
  if not errorlevel 1 exit /b 0
)
if exist "%LocalAppData%\Programs\Python\Python313\python.exe" (
  call :try_python "%LocalAppData%\Programs\Python\Python313\python.exe"
  if not errorlevel 1 exit /b 0
)
exit /b 1

:try_python
if exist "%VENV%" (
  rmdir /s /q "%VENV%" >nul 2>nul
  if exist "%VENV%" exit /b 1
)
%* -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if errorlevel 1 exit /b 1
%* -m venv "%VENV%" >nul 2>nul
if errorlevel 1 exit /b 1
"%VENV_PY%" -c "import sys; sys.exit(0 if sys.prefix != sys.base_prefix else 1)" >nul 2>nul
exit /b %errorlevel%

:failed_folder
echo ERROR: Cannot open the script directory.
goto :failed_no_pop
:failed_requirements
echo ERROR: requirements.txt was not found. Download the whole repository.
goto :failed
:failed_packages
echo ERROR: Python is working but package installation/import failed.
echo Check the pip error above; it may be a network or permissions issue.
goto :failed
:failed_python
echo ERROR: Could not install a working Python 3.
echo Manual alternative: https://www.python.org/downloads/windows/
goto :failed
:failed
popd
:failed_no_pop
echo INSTALLATION FAILED.
if "%NO_PAUSE%"=="0" pause
exit /b 1
:success
popd
if "%NO_PAUSE%"=="0" pause
exit /b 0
