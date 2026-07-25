import base64
import os
import socket
import shutil
import sys
import threading
import time
from pathlib import Path

def run_velopack_startup_hooks() -> None:
    try:
        import velopack
    except ImportError:
        return
    velopack.App().run()


if __name__ == "__main__":
    run_velopack_startup_hooks()


import uvicorn

from app import ICON_PATH
from web_api import app as fastapi_app, close_workflow_runtime, find_free_port

APP_TITLE = os.environ.get("PRODUCT_CODE_FORMATTER_APP_TITLE") or "ProductCodeFormatter"
DESKTOP_SHORTCUT_NAME = "ProductCodeFormatter.lnk"
DESKTOP_SHORTCUT_ALIASES = [
    "Product Code Formatter.lnk",
    "Product Code Formatter Web.lnk",
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


def current_exe_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def is_velopack_install() -> bool:
    if os.name != "nt" or not getattr(sys, "frozen", False):
        return False
    current_dir = current_exe_path().parent
    return (current_dir / "sq.version").exists() or (current_dir.parent / "sq.version").exists()


def cleanup_legacy_install_artifacts() -> None:
    if not is_velopack_install():
        return
    local_app_data = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    app_data_dir = local_app_data / "ProductCodeFormatter"
    legacy_exe = app_data_dir / "ProductCodeFormatter.exe"
    if legacy_exe.exists() and legacy_exe.resolve() != current_exe_path():
        try:
            legacy_exe.unlink()
        except OSError:
            pass
    shutil.rmtree(app_data_dir / "updates", ignore_errors=True)

    shortcut_roots = [
        Path.home() / "Desktop",
        Path(os.environ.get("PUBLIC") or "C:/Users/Public") / "Desktop",
        Path(os.environ.get("APPDATA") or "") / "Microsoft/Windows/Start Menu/Programs",
        Path(os.environ.get("PROGRAMDATA") or "C:/ProgramData") / "Microsoft/Windows/Start Menu/Programs",
    ]
    for root in shortcut_roots:
        if not root.exists():
            continue
        for name in DESKTOP_SHORTCUT_ALIASES:
            try:
                (root / name).unlink(missing_ok=True)
            except OSError:
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

    cleanup_legacy_install_artifacts()
    port = int(os.environ.get("PRODUCT_CODE_FORMATTER_PORT") or find_free_port())
    thread = threading.Thread(target=run_api_server, args=(port,), daemon=True)
    thread.start()
    wait_for_server(port)

    frontend_url = (os.environ.get("PRODUCT_CODE_FORMATTER_FRONTEND_URL") or "").strip()
    window_options = {
        "title": APP_TITLE,
        "url": frontend_url or f"http://127.0.0.1:{port}",
        "width": 1400,
        "height": 900,
        "min_size": (1024, 700),
    }
    if ICON_PATH.exists():
        window_options["background_color"] = "#eef4f8"
    webview.create_window(**window_options, js_api=DesktopApi())
    try:
        webview.start(debug=False)
    finally:
        close_workflow_runtime()
    os._exit(0)


if __name__ == "__main__":
    main()
