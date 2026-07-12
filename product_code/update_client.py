"""Velopack update integration for the Windows desktop application."""

import os
import re
from urllib.parse import urljoin

from product_code.release_version import CHANNEL


DEFAULT_UPDATE_SERVER_URL = os.environ.get(
    "PRODUCT_CODE_FORMATTER_UPDATE_SERVER",
    "http://192.168.1.210:8080",
)
UPDATE_FEED_PATH = "/updates/product-code-formatter"


def normalize_update_server_url(value=""):
    url = str(value or DEFAULT_UPDATE_SERVER_URL).strip().rstrip("/")
    if url and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url):
        url = f"http://{url}"
    return url


def velopack_feed_url(server_url=""):
    base = normalize_update_server_url(server_url)
    return urljoin(f"{base}/", UPDATE_FEED_PATH.lstrip("/"))


def version_tuple(value):
    numbers = re.findall(r"\d+", str(value or ""))
    return tuple(int(item) for item in numbers[:4]) or (0,)


def _update_manager(server_url=""):
    try:
        import velopack
    except ImportError as exc:
        raise ValueError("Bản ứng dụng này chưa chứa Velopack. Hãy cài bằng ProductCodeFormatter-Setup.exe một lần.") from exc
    feed_url = velopack_feed_url(server_url)
    try:
        channel = CHANNEL if CHANNEL in {"test", "stable"} else None
        options = velopack.UpdateOptions(False, 10, channel) if channel else None
        source = velopack.HttpSource(feed_url)
        return velopack.UpdateManager(source, options), feed_url
    except RuntimeError as exc:
        raise ValueError(
            "Bản đang chạy chưa được cài bằng ProductCodeFormatter-Setup.exe. "
            "Hãy chạy Setup một lần; config và license hiện tại vẫn được giữ nguyên."
        ) from exc


def _asset_payload(asset):
    if asset is None:
        return {}
    return {
        "version": str(asset.Version or ""),
        "filename": str(asset.FileName or ""),
        "sha256": str(asset.SHA256 or ""),
        "size": int(asset.Size or 0),
        "notes": str(asset.NotesMarkdown or ""),
        "package_id": str(asset.PackageId or ""),
        "asset_type": str(asset.Type or ""),
    }


def check_for_update(current_version, server_url=""):
    manager, feed_url = _update_manager(server_url)
    installed_version = str(manager.get_current_version() or current_version or "")
    update_info = manager.check_for_updates()
    if update_info is None:
        return {
            "current_version": installed_version,
            "version": installed_version,
            "available": False,
            "server_url": normalize_update_server_url(server_url),
            "feed_url": feed_url,
            "installed": True,
            "notes": "",
        }
    target = update_info.TargetFullRelease
    return {
        "current_version": installed_version,
        "available": True,
        "server_url": normalize_update_server_url(server_url),
        "feed_url": feed_url,
        "installed": True,
        **_asset_payload(target),
    }


def prepare_velopack_update(server_url="", progress_callback=None):
    manager, feed_url = _update_manager(server_url)
    update_info = manager.check_for_updates()
    if update_info is None:
        raise ValueError("Ứng dụng đang dùng phiên bản mới nhất; không có package cần cập nhật.")
    try:
        manager.download_updates(update_info, progress_callback)
    except Exception as exc:
        raise ValueError(f"Không tải hoặc xác minh được package Velopack từ {feed_url}: {exc}") from exc
    return manager, update_info, _asset_payload(update_info.TargetFullRelease)


def schedule_velopack_restart(manager, update_info):
    try:
        manager.wait_exit_then_apply_updates(update_info, silent=False, restart=True)
    except Exception as exc:
        raise ValueError(f"Không khởi động được Velopack updater: {exc}") from exc
    return {
        "scheduled": True,
        "version": str(update_info.TargetFullRelease.Version or ""),
    }