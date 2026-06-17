import base64
import os
import socket
import threading
import time
from pathlib import Path

import uvicorn

from app import ICON_PATH
from web_api import app as fastapi_app, find_free_port


class DesktopApi:
    def save_file(self, filename: str, data_base64: str) -> dict:
        import webview

        window = webview.windows[0] if webview.windows else None
        if window is None:
            return {"saved": False, "error": "Không tìm thấy cửa sổ để mở hộp thoại lưu file."}
        paths = window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=Path(filename).name,
            file_types=("Excel workbook (*.xlsx)", "All files (*.*)"),
        )
        if not paths:
            return {"saved": False, "cancelled": True}
        target = Path(paths[0])
        if target.suffix.lower() != ".xlsx":
            target = target.with_suffix(".xlsx")
        target.write_bytes(base64.b64decode(data_base64))
        return {"saved": True, "path": str(target)}


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
