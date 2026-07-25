# -*- mode: python ; coding: utf-8 -*-

import json
import os
from pathlib import Path
import sys

from product_code.release_version import NOTES as DEFAULT_NOTES, VERSION as DEFAULT_VERSION


build_version = os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_VERSION", DEFAULT_VERSION)
build_notes = os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_NOTES", DEFAULT_NOTES)
build_channel = os.environ.get("PRODUCT_CODE_FORMATTER_BUILD_CHANNEL", "dev")
metadata_dir = Path("build") / "release_metadata"
metadata_dir.mkdir(parents=True, exist_ok=True)
metadata_path = metadata_dir / "release_metadata.json"
metadata_path.write_text(
    json.dumps(
        {"version": build_version, "notes": build_notes, "channel": build_channel},
        ensure_ascii=False,
        indent=2,
    ),
    encoding="utf-8",
)

datas = [
    ("app_icon.ico", "."),
    ("mau HD ban ra.xlsx", "."),
    ("mau_mua_vao_up.xlsx", "."),
    ("mau_ban_ra_up.xlsx", "."),
    (str(metadata_path), "."),
]
react_dist = Path("react_frontend") / "dist"
if react_dist.exists():
    datas.append((str(react_dist), "react_frontend/dist"))

pathex = []
venv_site_packages = Path(".venv") / "Lib" / "site-packages"
venv_root = Path(".venv").resolve()
if venv_site_packages.exists() and Path(sys.prefix).resolve() != venv_root:
    pathex.append(str(venv_site_packages))

python_root = Path(sys.base_prefix)
runtime_binaries = []
for runtime_name in ("vcruntime140.dll", "vcruntime140_1.dll"):
    runtime_path = python_root / runtime_name
    if runtime_path.exists():
        runtime_binaries.append((str(runtime_path), "."))

a = Analysis(
    ["web_desktop_app.py"],
    pathex=pathex,
    binaries=runtime_binaries,
    datas=datas,
    hiddenimports=[
        "velopack",
        "webview",
        "webview.platforms.edgechromium",
        "uvicorn",
        "uvicorn.logging",
        "uvicorn.loops",
        "uvicorn.loops.auto",
        "uvicorn.protocols",
        "uvicorn.protocols.http",
        "uvicorn.protocols.http.auto",
        "uvicorn.protocols.websockets",
        "uvicorn.protocols.websockets.auto",
        "fastapi",
        "fastapi.middleware.cors",
        "pydantic",
        "pandas",
        "openpyxl",
        "xlrd",
        "xlwt",
        "numpy",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ProductCodeFormatter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["app_icon.ico"],
)
