"""Excel IO helpers shared by FastAPI routes.

This keeps legacy .xls/.xlsx conversion and diagnostic logging out of route
modules while preserving the existing workbook-reading behavior from app.py.
"""

import time
from io import BytesIO
from pathlib import Path

import pandas as pd

from app import UPLOAD_DIR, raw_text, read_workbook


DIAGNOSTIC_LOG_PATH = UPLOAD_DIR.parent / "diagnostics.log"


def diagnostic_log(message: str) -> None:
    try:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        DIAGNOSTIC_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with DIAGNOSTIC_LOG_PATH.open("a", encoding="utf-8") as fh:
            fh.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


def is_zip_excel_bytes(content: bytes) -> bool:
    return bytes(content or b"").startswith(b"PK")


def dataframe_to_openpyxl_bytes(sheet_name: str, df: pd.DataFrame) -> bytes:
    stream = BytesIO()
    safe_sheet = (raw_text(sheet_name) or "Sheet1")[:31]
    with pd.ExcelWriter(stream, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=safe_sheet, index=False, header=False)
    return stream.getvalue()


def workbook_content_for_openpyxl(path: Path) -> bytes:
    content = path.read_bytes()
    if is_zip_excel_bytes(content):
        return content
    sheet, df = read_workbook(path)
    diagnostic_log(f"convert workbook for openpyxl path={path.name} rows={len(df)} cols={df.shape[1]}")
    return dataframe_to_openpyxl_bytes(sheet, df)


def uploaded_workbook_content_for_openpyxl(content: bytes, filename: str = "") -> bytes:
    content = bytes(content or b"")
    if not content or is_zip_excel_bytes(content):
        return content
    suffix = Path(filename or "").suffix.lower()
    legacy_xls = content.startswith(b"\xd0\xcf\x11\xe0") or (suffix == ".xls" and not is_zip_excel_bytes(content))
    engine = "xlrd" if legacy_xls else "openpyxl"
    with pd.ExcelFile(BytesIO(content), engine=engine) as xl:
        sheet = xl.sheet_names[0]
        df = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=object)
    diagnostic_log(f"convert uploaded workbook for openpyxl filename={filename or '-'} rows={len(df)} cols={df.shape[1]}")
    return dataframe_to_openpyxl_bytes(sheet, df)
