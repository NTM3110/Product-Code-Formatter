# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

datas = [('app_icon.ico', '.'), ('mau HD ban ra.xlsx', '.'), ('mau_mua_vao_up.xlsx', '.'), ('mau_ban_ra_up.xlsx', '.')]
react_dist = Path('react_frontend') / 'dist'
if react_dist.exists():
    datas.append((str(react_dist), 'react_frontend/dist'))

a = Analysis(
    ['web_desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['webview', 'webview.platforms.edgechromium', 'uvicorn', 'uvicorn.logging', 'uvicorn.loops', 'uvicorn.loops.auto', 'uvicorn.protocols', 'uvicorn.protocols.http', 'uvicorn.protocols.http.auto', 'uvicorn.protocols.websockets', 'uvicorn.protocols.websockets.auto', 'fastapi', 'fastapi.middleware.cors', 'pydantic', 'pandas', 'openpyxl', 'xlrd', 'xlwt', 'numpy'],
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
    name='ProductCodeFormatterWeb',
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
    icon=['app_icon.ico'],
)
