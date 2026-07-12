"""Validate and publish Velopack release bundles on the Raspberry Pi."""

import hashlib
import json
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path, PurePosixPath


CHANNEL_PACKAGE_IDS = {
    "stable": "ProductCodeFormatter.App",
    "test": "ProductCodeFormatter.App.Test",
}
MAX_BUNDLE_BYTES = 1_500_000_000
MAX_EXPANDED_BYTES = 2_000_000_000
_VERSION_RE = re.compile(r"^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$")


class ReleaseBundleError(ValueError):
    pass


def release_feed_dir(release_dir):
    return Path(release_dir) / "product-code-formatter"


def _version_key(value):
    text = str(value or "")
    main, separator, suffix = text.partition("-")
    numbers = [int(item) for item in re.findall(r"\d+", main)[:4]]
    numbers.extend([0] * (4 - len(numbers)))
    return (*numbers, 1 if not separator else 0, suffix)


def _manifest_assets(payload):
    assets = payload.get("Assets") if isinstance(payload, dict) else None
    if not isinstance(assets, list):
        raise ReleaseBundleError("Feed Velopack không có danh sách Assets hợp lệ.")
    return [item for item in assets if isinstance(item, dict)]


def _latest_full_asset(payload):
    full_assets = [
        asset
        for asset in _manifest_assets(payload)
        if str(asset.get("Type") or "").casefold() == "full"
    ]
    if not full_assets:
        raise ReleaseBundleError("Feed Velopack không có full package.")
    return max(full_assets, key=lambda asset: _version_key(asset.get("Version")))


def _read_manifest(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ReleaseBundleError(f"Không đọc được feed {Path(path).name}: {exc}") from exc


def _sha256_stream(stream):
    digest = hashlib.sha256()
    while True:
        chunk = stream.read(1024 * 1024)
        if not chunk:
            break
        digest.update(chunk)
    return digest.hexdigest().upper()


def _safe_member_name(info):
    name = str(info.filename or "").replace("\\", "/")
    path = PurePosixPath(name)
    if info.is_dir() or not name or path.is_absolute() or len(path.parts) != 1 or path.name != name:
        raise ReleaseBundleError(f"Bundle có đường dẫn không hợp lệ: {name or '(trống)'}")
    if info.flag_bits & 0x1:
        raise ReleaseBundleError(f"Bundle không được chứa file mã hóa: {name}")
    return name


def _allowed_name(name):
    return bool(
        re.fullmatch(r"releases\.(?:stable|test)\.json", name)
        or re.fullmatch(r"assets\.(?:stable|test)\.json", name)
        or re.fullmatch(r"RELEASES-(?:stable|test)", name)
        or re.fullmatch(r"[A-Za-z0-9._-]+\.nupkg", name)
        or re.fullmatch(r"[A-Za-z0-9._-]+-Setup\.exe", name)
    )


def _validate_bundle(archive, feed_dir):
    members = {}
    expanded_size = 0
    for info in archive.infolist():
        name = _safe_member_name(info)
        if not _allowed_name(name):
            raise ReleaseBundleError(f"Bundle chứa file không thuộc feed Velopack: {name}")
        if name in members:
            raise ReleaseBundleError(f"Bundle có file trùng tên: {name}")
        expanded_size += int(info.file_size or 0)
        if expanded_size > MAX_EXPANDED_BYTES:
            raise ReleaseBundleError("Bundle sau giải nén vượt quá giới hạn 2 GB.")
        members[name] = info

    manifests = [name for name in members if re.fullmatch(r"releases\.(stable|test)\.json", name)]
    if len(manifests) != 1:
        raise ReleaseBundleError("Bundle phải có đúng một file releases.stable.json hoặc releases.test.json.")

    manifest_name = manifests[0]
    channel = manifest_name.split(".")[1]
    package_id = CHANNEL_PACKAGE_IDS[channel]
    expected_metadata = {
        f"releases.{channel}.json",
        f"assets.{channel}.json",
        f"RELEASES-{channel}",
    }
    expected_setup = f"{package_id}-{channel}-Setup.exe"

    for name in members:
        if name.startswith(("releases.", "assets.", "RELEASES-")) and name not in expected_metadata:
            raise ReleaseBundleError(f"Bundle trộn metadata của channel khác: {name}")
        if name.endswith(".nupkg") and (
            not name.startswith(f"{package_id}-") or f"-{channel}-" not in name
        ):
            raise ReleaseBundleError(f"Package không đúng app/channel {channel}: {name}")
        if name.endswith("-Setup.exe") and name != expected_setup:
            raise ReleaseBundleError(f"Setup không đúng app/channel {channel}: {name}")

    with archive.open(members[manifest_name]) as stream:
        try:
            payload = json.loads(stream.read().decode("utf-8-sig"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ReleaseBundleError(f"{manifest_name} không phải JSON hợp lệ: {exc}") from exc

    assets = _manifest_assets(payload)
    for asset in assets:
        if str(asset.get("PackageId") or "") != package_id:
            raise ReleaseBundleError(
                f"Feed {channel} chứa PackageId không hợp lệ: {asset.get('PackageId')}"
            )
        filename = Path(str(asset.get("FileName") or "")).name
        if not filename:
            raise ReleaseBundleError("Feed có asset thiếu FileName.")
        if filename not in members and not (feed_dir / filename).is_file():
            raise ReleaseBundleError(f"Feed tham chiếu package chưa có trên server: {filename}")

    latest = _latest_full_asset(payload)
    version = str(latest.get("Version") or "")
    if not _VERSION_RE.fullmatch(version):
        raise ReleaseBundleError(f"Version Velopack không hợp lệ: {version}")
    package_name = Path(str(latest.get("FileName") or "")).name
    if package_name not in members:
        raise ReleaseBundleError(f"Bundle thiếu full package mới nhất: {package_name}")
    if expected_setup not in members:
        raise ReleaseBundleError(f"Bundle thiếu bộ cài {expected_setup}.")

    package_info = members[package_name]
    expected_size = int(latest.get("Size") or 0)
    if expected_size and expected_size != package_info.file_size:
        raise ReleaseBundleError(
            f"Sai kích thước {package_name}: feed={expected_size}, bundle={package_info.file_size}."
        )
    expected_sha = str(latest.get("SHA256") or "").strip().upper()
    if expected_sha:
        with archive.open(package_info) as stream:
            actual_sha = _sha256_stream(stream)
        if actual_sha != expected_sha:
            raise ReleaseBundleError(
                f"Sai SHA-256 của {package_name}: feed={expected_sha}, bundle={actual_sha}."
            )

    return {
        "channel": channel,
        "package_id": package_id,
        "version": version,
        "notes": str(latest.get("NotesMarkdown") or ""),
        "package": package_name,
        "setup": expected_setup,
        "sha256": expected_sha,
        "size": package_info.file_size,
        "manifest": manifest_name,
        "members": members,
    }


def install_release_bundle(bundle_path, release_dir):
    bundle_path = Path(bundle_path)
    if not bundle_path.is_file():
        raise ReleaseBundleError("Không tìm thấy release bundle ZIP.")
    if bundle_path.stat().st_size > MAX_BUNDLE_BYTES:
        raise ReleaseBundleError("Release bundle vượt quá giới hạn 1.5 GB.")

    feed_dir = release_feed_dir(release_dir)
    feed_dir.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(bundle_path) as archive:
            release = _validate_bundle(archive, feed_dir)
            with tempfile.TemporaryDirectory(prefix=".velopack-stage-", dir=str(feed_dir.parent)) as temp_name:
                stage_dir = Path(temp_name)
                for name, info in release["members"].items():
                    target = stage_dir / name
                    with archive.open(info) as source, target.open("wb") as output:
                        shutil.copyfileobj(source, output, length=1024 * 1024)

                manifest_name = release["manifest"]
                publish_names = [name for name in release["members"] if name != manifest_name]
                publish_names.append(manifest_name)
                for name in publish_names:
                    source = stage_dir / name
                    target = feed_dir / name
                    temporary = feed_dir / f".{name}.{os.getpid()}.tmp"
                    shutil.copyfile(source, temporary)
                    os.replace(temporary, target)
    except ReleaseBundleError:
        raise
    except (OSError, RuntimeError, zipfile.BadZipFile, zipfile.LargeZipFile) as exc:
        raise ReleaseBundleError(f"Không đọc hoặc publish được release bundle ZIP: {exc}") from exc

    return {key: value for key, value in release.items() if key != "members"}
def channel_release_status(release_dir, channel):
    if channel not in CHANNEL_PACKAGE_IDS:
        raise ReleaseBundleError(f"Channel không hợp lệ: {channel}")
    feed_dir = release_feed_dir(release_dir)
    manifest_path = feed_dir / f"releases.{channel}.json"
    if not manifest_path.is_file():
        return {
            "channel": channel,
            "available": False,
            "package_id": CHANNEL_PACKAGE_IDS[channel],
        }
    payload = _read_manifest(manifest_path)
    latest = _latest_full_asset(payload)
    setup_name = f"{CHANNEL_PACKAGE_IDS[channel]}-{channel}-Setup.exe"
    return {
        "channel": channel,
        "available": True,
        "package_id": str(latest.get("PackageId") or ""),
        "version": str(latest.get("Version") or ""),
        "notes": str(latest.get("NotesMarkdown") or ""),
        "package": Path(str(latest.get("FileName") or "")).name,
        "sha256": str(latest.get("SHA256") or ""),
        "size": int(latest.get("Size") or 0),
        "setup": setup_name,
        "setup_available": (feed_dir / setup_name).is_file(),
        "feed_url": "/updates/product-code-formatter",
    }


def all_release_status(release_dir):
    return {
        "releases": [
            channel_release_status(release_dir, channel)
            for channel in ("test", "stable")
        ]
    }
