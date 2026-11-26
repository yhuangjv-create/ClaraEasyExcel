@echo off
REM ====================================================================
REM ClaraEasyExcel - One-click launcher for Windows
REM Auto-detects Python, prefers pythonw.exe (no console window)
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
goto run

:use_venv_py
set "PYTHON_CMD=.venv\Scripts\python.exe"
goto run

:use_sys_w
set "PYTHON_CMD=pythonw.exe"
goto run

:use_sys_py
set "PYTHON_CMD=python.exe"
goto run

:run
"%PYTHON_CMD%" entry.py
exit /b %ERRORLEVEL%
