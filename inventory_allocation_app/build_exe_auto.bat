@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  py -m venv .venv
)
set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
if exist "..\.venv\Scripts\python.exe" set "PYTHON_EXE=%CD%\..\.venv\Scripts\python.exe"
"%PYTHON_EXE%" -m pip install -r requirements.txt

set "PYI_TOKEN=%RANDOM%%RANDOM%"
set "PYI_ROOT=%CD%\.pyinstaller_tmp"
set "PYI_WORKPATH=%PYI_ROOT%\build-%PYI_TOKEN%"
set "PYI_DISTPATH=%PYI_ROOT%\dist-%PYI_TOKEN%"
if not exist "%PYI_ROOT%" mkdir "%PYI_ROOT%"

"%PYTHON_EXE%" -m PyInstaller --clean --workpath "%PYI_WORKPATH%" --distpath "%PYI_DISTPATH%" --noconsole --onefile --icon "assets\black_coffee.ico" --add-data "static;static" --add-data "assets;assets" --name InventoryAllocator app.py
if errorlevel 1 exit /b 1

if not exist "deploy" mkdir "deploy"
copy /y "%PYI_DISTPATH%\InventoryAllocator.exe" "deploy\InventoryAllocator.exe" >nul
set "DEPLOY_EXE=%CD%\deploy\InventoryAllocator.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$target=$env:DEPLOY_EXE; $desktop=[Environment]::GetFolderPath('Desktop'); $shortcut=Join-Path $desktop 'Inventory Allocator.lnk'; $shell=New-Object -ComObject WScript.Shell; $link=$shell.CreateShortcut($shortcut); $link.TargetPath=$target; $link.WorkingDirectory=Split-Path $target; $link.IconLocation=$target + ',0'; $link.Description='Inventory Allocator'; $link.Save()"
echo Built deploy\InventoryAllocator.exe
echo Created desktop shortcut "Inventory Allocator"
