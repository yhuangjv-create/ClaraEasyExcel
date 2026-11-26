@echo off
REM ====================================================================
REM ClaraEasyExcel - One-click launcher for Windows
REM   This .bat only discovers python, then hands off to run.vbs
REM   which launches pythonw.exe with NO console window.
REM ====================================================================
cd /d "%~dp0"

set "PYTHON_CMD="

if exist ".venv\Scripts\pythonw.exe" goto use_venv_w
if exist ".venv\Scripts\python.exe" goto use_venv_py

where pythonw.exe >nul 2>nul
if %ERRORLEVEL%==0 goto use_sys_w

where python.exe >nul 2>nul
if %ERRORLEVEL%==0 goto use_sys_py

echo [ClaraEasyExcel] ERROR: No Python found on your system.
echo Please install Python 3.11+ from https://www.python.org/downloads/
echo Then run once:  python -m venv .venv ^&^& pip install -e .
pause
exit /b 1

:use_venv_w
set "PYTHON_CMD=.venv\Scripts\pythonw.exe"
goto run_vbs

:use_venv_py
set "PYTHON_CMD=.venv\Scripts\python.exe"
goto run_vbs

:use_sys_w
set "PYTHON_CMD=pythonw.exe"
goto run_vbs

:use_sys_py
set "PYTHON_CMD=python.exe"
goto run_vbs

REM --------------------------------------------------------------------
REM hand off to run.vbs, then exit IMMEDIATELY so the CMD window closes
REM --------------------------------------------------------------------
:run_vbs
start "" wscript "%~dp0run.vbs" "%PYTHON_CMD%"
exit /b 0
