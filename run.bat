@echo off
REM ============================================================
REM  ClaraEasyExcel — One-click launcher
REM  Auto-detects venv, prefers pythonw (no console window),
REM  falls back to system python if needed.
REM ============================================================

REM always run from the directory where this .bat lives
cd /d "%~dp0"

REM --- pick the best Python ---
set "PYTHON_CMD="

if exist ".venv\Scripts\pythonw.exe" (
    set "PYTHON_CMD=.venv\Scripts\pythonw.exe"
    goto :run
)
if exist ".venv\Scripts\python.exe" (
    set "PYTHON_CMD=.venv\Scripts\python.exe"
    goto :run
)

where pythonw.exe >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PYTHON_CMD=pythonw.exe"
    goto :run
)

where python.exe >nul 2>nul
if %ERRORLEVEL%==0 (
    set "PYTHON_CMD=python.exe"
    goto :run
)

REM no Python found at all
echo [ClaraEasyExcel] ERROR: No Python found.
echo   Please install Python 3.11+ from https://www.python.org/downloads/
echo   or run: python -m venv .venv ^&^& pip install -e .
pause
exit /b 1

REM --- run ---
:run
"%PYTHON_CMD%" entry.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ClaraEasyExcel] exited with code %ERRORLEVEL%.
    echo   Re-running with visible console to help debug...
    python entry.py
)
