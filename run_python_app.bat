@echo off
setlocal EnableExtensions

set "PYTHON_CMD="

py -3 --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    python --version >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    )
)

if "%PYTHON_CMD%"=="" (
    echo Python was not found.
    echo Run build_exe_auto.bat first, or install Python manually.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating local virtual environment...
    %PYTHON_CMD% -m venv .venv
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)

".venv\Scripts\python.exe" app.py
