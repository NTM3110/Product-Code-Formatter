@echo off
setlocal
cd /d "%~dp0"
echo Product Code Formatter license admin - Vietmax unified workflow / ProductCodeFormatter_v27
if not exist ".venv\Scripts\python.exe" (
    echo Missing .venv. Trying system Python launcher...
    py -3 license_server_admin.py
    if errorlevel 1 (
        pause
        exit /b 1
    )
    exit /b 0
)
".venv\Scripts\python.exe" license_server_admin.py
if errorlevel 1 (
    pause
    exit /b 1
)
endlocal
