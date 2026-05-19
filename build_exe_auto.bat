@echo off
setlocal EnableExtensions EnableDelayedExpansion

echo ============================================
echo Product Code Formatter - Auto Build EXE
echo ============================================
echo.

REM --------------------------------------------------
REM Find Python
REM --------------------------------------------------
set "PYTHON_CMD="

py -3 --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :python_found
)

python --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo Python was not found.
echo Trying to install Python automatically using winget...
echo.

winget --version >nul 2>&1
if not %errorlevel%==0 (
    echo ERROR: winget is not available on this computer.
    echo.
    echo Please install Python manually from:
    echo https://www.python.org/downloads/windows/
    echo.
    echo During installation, tick:
    echo [x] Add python.exe to PATH
    echo.
    pause
    exit /b 1
)

winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements

echo.
echo Python installation attempted.
echo Trying to find Python again...
echo.

py -3 --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :python_found
)

python --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo ERROR: Python still cannot be found.
echo Close this CMD window, open a new CMD window, then run this file again.
pause
exit /b 1


:python_found
echo Using Python command: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

REM --------------------------------------------------
REM Create virtual environment
REM --------------------------------------------------
echo Creating local virtual environment...
%PYTHON_CMD% -m venv .venv
if not %errorlevel%==0 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
)

set "VENV_PY=.venv\Scripts\python.exe"

echo.
echo Upgrading pip...
"%VENV_PY%" -m pip install --upgrade pip
if not %errorlevel%==0 (
    echo ERROR: Failed to upgrade pip.
    pause
    exit /b 1
)

echo.
echo Installing required packages...
"%VENV_PY%" -m pip install -r requirements.txt
if not %errorlevel%==0 (
    echo ERROR: Failed to install requirements.
    pause
    exit /b 1
)

echo.
echo Building Angular frontend...
cd frontend
call npm install
if not %errorlevel%==0 (
    echo ERROR: Failed to install frontend dependencies.
    cd ..
    pause
    exit /b 1
)
call npm run build
if not %errorlevel%==0 (
    echo ERROR: Failed to build frontend.
    cd ..
    pause
    exit /b 1
)
cd ..

echo.
echo Preparing static folder...
if exist static rmdir /s /q static
mkdir static
xcopy /E /I /Y "frontend\dist\frontend\browser\*" "static\"

echo.
echo Closing any running instances of the application...
taskkill /F /IM ProductCodeFormatter.exe /T >nul 2>&1

echo.
echo Building EXE...
"%VENV_PY%" -m PyInstaller --noconsole --onefile --icon "app_icon.ico" --add-data "static;static" --add-data "app_icon.ico;." --name ProductCodeFormatter app.py
if not %errorlevel%==0 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo Copying EXE to deploy folder...
if not exist deploy mkdir deploy
copy /Y "dist\ProductCodeFormatter.exe" "deploy\ProductCodeFormatter.exe" >nul
if not %errorlevel%==0 (
    echo ERROR: Failed to copy EXE to deploy folder.
    pause
    exit /b 1
)
if exist "deploy\app_icon.ico" del /q "deploy\app_icon.ico"

echo.
echo Creating Desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$desktop=[Environment]::GetFolderPath('Desktop'); $shortcut=Join-Path $desktop 'Product Code Formatter.lnk'; $target=(Resolve-Path 'deploy\ProductCodeFormatter.exe').Path; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=$target; $link.WorkingDirectory=(Resolve-Path 'deploy').Path; $link.IconLocation=$target; $link.Save()"
if not %errorlevel%==0 (
    echo ERROR: Failed to create Desktop shortcut.
    pause
    exit /b 1
)

echo.
echo ============================================
echo DONE
echo Your EXE is here:
echo deploy\ProductCodeFormatter.exe
echo ============================================
echo.
pause
