import base64
import os
import socket
import subprocess
import sys
import threading
import time
from pathlib import Path

import uvicorn

from app import ICON_PATH
from web_api import app as fastapi_app, find_free_port

APP_TITLE = "Product Code Formatter"
DESKTOP_SHORTCUT_NAME = "Product Code Formatter.lnk"
DESKTOP_SHORTCUT_ALIASES = [
    "Product Code Formatter.lnk",
    "Product Code Formatter Web.lnk",
    "ProductCodeFormatter.lnk",
    "ProductCodeFormatterWeb.lnk",
]


class DesktopApi:
    def save_file(self, filename: str, data_base64: str) -> dict:
        import webview

        window = webview.windows[0] if webview.windows else None
        if window is None:
            return {"saved": False, "error": "Không tìm thấy cửa sổ để mở hộp thoại lưu file."}
        requested_name = Path(filename).name
        requested_suffix = Path(requested_name).suffix.lower()
        if requested_suffix == ".json":
            file_types = ("JSON config (*.json)", "All files (*.*)")
            default_suffix = ".json"
        elif requested_suffix == ".xls":
            file_types = ("Excel workbook (*.xls)", "All files (*.*)")
            default_suffix = ".xls"
        elif requested_suffix == ".xlsm":
            file_types = ("Excel macro workbook (*.xlsm)", "All files (*.*)")
            default_suffix = ".xlsm"
        elif requested_suffix == ".zip":
            file_types = ("ZIP archive (*.zip)", "All files (*.*)")
            default_suffix = ".zip"
        else:
            file_types = ("Excel workbook (*.xlsx)", "All files (*.*)")
            default_suffix = ".xlsx"
        paths = window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=requested_name,
            file_types=file_types,
        )
        if not paths:
            return {"saved": False, "cancelled": True}
        target = Path(paths[0])
        if not target.suffix:
            target = target.with_suffix(default_suffix)
        elif default_suffix == ".xls" and target.suffix.lower() != ".xls":
            target = target.with_suffix(".xls")
        target.write_bytes(base64.b64decode(data_base64))
        return {"saved": True, "path": str(target)}


def ps_quote(value: str | Path) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def current_exe_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def ensure_desktop_shortcut() -> None:
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return
    exe_path = current_exe_path()
    shortcut_names = "@(" + ",".join(ps_quote(name) for name in DESKTOP_SHORTCUT_ALIASES) + ")"
    script = f"""
$ErrorActionPreference = 'SilentlyContinue'
$ExePath = {ps_quote(exe_path)}
$ShortcutName = {ps_quote(DESKTOP_SHORTCUT_NAME)}
$ShortcutAliases = {shortcut_names}
$Shell = New-Object -ComObject WScript.Shell
$desktopPaths = @(
  [Environment]::GetFolderPath('Desktop'),
  [Environment]::GetFolderPath('CommonDesktopDirectory')
) | Where-Object {{ $_ -and (Test-Path $_) }} | Select-Object -Unique

foreach ($desktop in $desktopPaths) {{
  foreach ($name in $ShortcutAliases) {{
    $shortcutPath = Join-Path $desktop $name
    if (Test-Path $shortcutPath) {{
      try {{
        Remove-Item -LiteralPath $shortcutPath -Force
      }} catch {{}}
    }}
  }}
}}

$userDesktop = [Environment]::GetFolderPath('Desktop')
if ($userDesktop -and (Test-Path $userDesktop)) {{
  $shortcutPath = Join-Path $userDesktop $ShortcutName
  $link = $Shell.CreateShortcut($shortcutPath)
  $link.TargetPath = $ExePath
  $link.WorkingDirectory = Split-Path -Parent $ExePath
  $link.IconLocation = $ExePath + ',0'
  $link.Description = 'Product Code Formatter'
  $link.Save()
}}

$ie4uinit = Join-Path $env:WINDIR 'System32\\ie4uinit.exe'
if (Test-Path $ie4uinit) {{
  Start-Process -FilePath $ie4uinit -ArgumentList '-show' -WindowStyle Hidden
}}
""".strip()
    startupinfo = None
    creationflags = 0
    if hasattr(subprocess, "STARTUPINFO"):
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            text=True,
            capture_output=True,
            timeout=8,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except Exception:
        pass


def wait_for_server(port: int, timeout: float = 20.0) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.25):
                return
        except OSError:
            time.sleep(0.15)
    raise RuntimeError("FastAPI server did not start in time.")


def run_api_server(port: int) -> None:
    config = uvicorn.Config(fastapi_app, host="127.0.0.1", port=port, log_level="warning", access_log=False)
    server = uvicorn.Server(config)
    server.run()


def main() -> None:
    import webview

    ensure_desktop_shortcut()
    port = int(os.environ.get("PRODUCT_CODE_FORMATTER_PORT") or find_free_port())
    thread = threading.Thread(target=run_api_server, args=(port,), daemon=True)
    thread.start()
    wait_for_server(port)

    window_options = {
        "title": "Product Code Formatter - Vietmax",
        "url": f"http://127.0.0.1:{port}",
        "width": 1400,
        "height": 900,
        "min_size": (1024, 700),
    }
    if ICON_PATH.exists():
        window_options["background_color"] = "#eef4f8"
    webview.create_window(**window_options, js_api=DesktopApi())
    webview.start(debug=False)
    os._exit(0)


if __name__ == "__main__":
    main()
