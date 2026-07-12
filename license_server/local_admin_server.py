"""ARM-native license and release server for Raspberry Pi.

This replaces the x86-only Keygen dependency for the LAN deployment. License
validation intentionally stays online: the Pi owns the database and the
Windows client receives only the metadata needed to run the selected profile.
"""

import hashlib
import json
import re
import shutil
import os
import secrets
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from release_store import (
    MAX_BUNDLE_BYTES,
    ReleaseBundleError,
    all_release_status,
    install_release_bundle,
    release_feed_dir,
)


APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
RELEASE_DIR = Path(os.environ.get("RELEASE_DIR", str(APP_DIR / "releases")))
DB_PATH = Path(os.environ.get("LICENSE_DB_PATH", "/data/license-data/licenses.sqlite3"))
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "").strip()
SERVER_MODE = os.environ.get("LICENSE_SERVER_MODE", "").strip()
ACCOUNT_ID = os.environ.get("LICENSE_ACCOUNT_ID", "6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001")
if SERVER_MODE != "local-pi-v1":
    raise RuntimeError(
        "Sai hoặc thiếu LICENSE_SERVER_MODE. Hãy chạy ./setup_pi.sh để tạo .env local-pi-v1."
    )
if not ADMIN_TOKEN or ADMIN_TOKEN.startswith("change-"):
    raise RuntimeError(
        "ADMIN_TOKEN chưa được cấu hình an toàn. Hãy chạy ./setup_pi.sh để tạo token."
    )
DEFAULT_PROFILE_OPTIONS = [
    {"code": "son_phuong", "label": "Sơn Phương"},
    {"code": "cao_thanh", "label": "Cao Thành"},
    {"code": "quang_thinh", "label": "Quang Thịnh"},
    {"code": "vietmax", "label": "Vietmax"},
    {"code": "ho_guom", "label": "Hồ Gươm"},
    {"code": "viet_hung", "label": "Việt Hưng"},
]
MANIFEST_PATH = RELEASE_DIR / "manifest.json"
UPDATE_FEED_DIR = release_feed_dir(RELEASE_DIR)

def release_metadata_from_filename(filename):
    stem = Path(str(filename or "")).stem
    match = re.search(r"(?:^|[_-])v(\d+[.]\d+[.]\d+(?:[-+][A-Za-z0-9.-]+)?)", stem, re.IGNORECASE)
    if not match:
        raise ValueError("Không đọc được version từ tên file. Hãy đặt tên dạng ProductCodeFormatter_v0.3.5.exe.")
    version = match.group(1)
    suffix = stem[match.end():].lstrip("_-").replace("_", " ").replace("-", " ").strip()
    notes = f"ProductCodeFormatter {version}" + (f" - {suffix}" if suffix else "")
    return version, notes

app = FastAPI(title="Product Code Formatter ARM License Server", version="1.0.0")
RELEASE_DIR.mkdir(parents=True, exist_ok=True)
UPDATE_FEED_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH.parent.mkdir(parents=True, exist_ok=True)


class LicenseCreate(BaseModel):
    key: str = ""
    name: str = ""
    expiry: str = ""
    allowed_profiles: list[str] = Field(default_factory=list)
    allowed_companies: list[str] = Field(default_factory=list)


class LicenseEdit(LicenseCreate):
    suspended: bool | None = None
    allowed_companies: list[str] | None = None


class ProfileOptionCreate(BaseModel):
    code: str
    label: str


class ProfileOptionEdit(ProfileOptionCreate):
    pass


def now_text():
    return datetime.now(timezone.utc).isoformat()


def connect():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with connect() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS licenses (
                id TEXT PRIMARY KEY,
                license_key TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL DEFAULT '',
                expiry TEXT NOT NULL DEFAULT '',
                suspended INTEGER NOT NULL DEFAULT 0,
                allowed_profiles TEXT NOT NULL DEFAULT '[]',
                allowed_companies TEXT NOT NULL DEFAULT '[]',
                product_code TEXT NOT NULL DEFAULT 'Product Code Formatter',
                application TEXT NOT NULL DEFAULT 'Product Code Formatter',
                machine_fingerprint TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        db.execute("""
            CREATE TABLE IF NOT EXISTS profile_options (
                code TEXT PRIMARY KEY,
                label TEXT NOT NULL,
                builtin INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)
        for option in DEFAULT_PROFILE_OPTIONS:
            db.execute(
                "INSERT OR IGNORE INTO profile_options (code, label, builtin, created_at) VALUES (?, ?, 1, ?)",
                (option["code"], option["label"], now_text()),
            )


def clean_list(values):
    if isinstance(values, str):
        values = values.replace(";", ",").replace("\n", ",").split(",")
    return list(dict.fromkeys(str(item).strip() for item in (values or []) if str(item).strip()))


def license_key():
    parts = [secrets.token_hex(3).upper() for _ in range(4)]
    return "-".join(parts) + "-V3"


def row_json(row):
    if row is None:
        return None
    return {
        "id": row["id"],
        "key": row["license_key"],
        "name": row["name"],
        "status": "SUSPENDED" if row["suspended"] else "ACTIVE",
        "expiry": row["expiry"],
        "suspended": bool(row["suspended"]),
        "allowed_profiles": json.loads(row["allowed_profiles"] or "[]"),
        "allowed_companies": json.loads(row["allowed_companies"] or "[]"),
        "product_code": row["product_code"],
        "application": row["application"],
        "machine_fingerprint": row["machine_fingerprint"],
    }


def row_by_key(key):
    with connect() as db:
        return db.execute("SELECT * FROM licenses WHERE license_key = ?", (key,)).fetchone()


def row_by_id(identifier):
    with connect() as db:
        return db.execute("SELECT * FROM licenses WHERE id = ?", (identifier,)).fetchone()


def keygen_shape(row):
    item = row_json(row)
    return {
        "type": "licenses",
        "id": item["id"],
        "attributes": {
            "key": item["key"],
            "name": item["name"],
            "status": item["status"],
            "expiry": item["expiry"] or None,
            "suspended": item["suspended"],
            "metadata": {
                "product_code": item["product_code"],
                "application": item["application"],
                "allowed_profiles": item["allowed_profiles"],
                "allowed_companies": [],
            },
        },
    }


def require_admin(authorization: str):
    if not ADMIN_TOKEN or authorization.strip() != f"Bearer {ADMIN_TOKEN}":
        raise HTTPException(status_code=401, detail="Cần ADMIN_TOKEN để quản trị license.")


def license_payload(row, valid=True, code="VALID", detail=""):
    return {"meta": {"valid": valid, "code": code, "detail": detail}, "data": keygen_shape(row)}


def validate_expiry(value):
    if not value:
        return False
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")) <= datetime.now(timezone.utc)
    except ValueError:
        return False


@app.post("/v1/accounts/{account}/licenses/actions/validate-key")
def validate_key(account: str, payload: dict, authorization: str = Header(default="")):
    meta = payload.get("meta") if isinstance(payload, dict) else {}
    key = str((meta or {}).get("key") or "").strip()
    if not key and authorization.lower().startswith("license "):
        key = authorization[8:].strip()
    row = row_by_key(key)
    if row is None:
        raise HTTPException(status_code=404, detail="License không tồn tại.")
    if row["suspended"]:
        return license_payload(row, False, "SUSPENDED", "License đã bị khóa.")
    if validate_expiry(row["expiry"]):
        return license_payload(row, False, "EXPIRED", "License đã hết hạn.")
    fingerprint = str(((meta or {}).get("scope") or {}).get("fingerprint") or "").strip()
    stored = row["machine_fingerprint"]
    if stored and fingerprint and stored != fingerprint:
        return license_payload(row, False, "MACHINE_SCOPE_MISMATCH", "License đã gắn với máy khác.")
    if fingerprint and not stored:
        with connect() as db:
            db.execute("UPDATE licenses SET machine_fingerprint = ?, updated_at = ? WHERE id = ?", (fingerprint, now_text(), row["id"]))
    return license_payload(row, True)


@app.post("/v1/accounts/{account}/machines")
def activate_machine(account: str, payload: dict, authorization: str = Header(default="")):
    key = authorization[8:].strip() if authorization.lower().startswith("license ") else ""
    row = row_by_key(key)
    if row is None:
        raise HTTPException(status_code=404, detail="License không tồn tại.")
    fingerprint = str((((payload.get("data") or {}).get("attributes") or {}).get("fingerprint") or "").strip())
    with connect() as db:
        db.execute("UPDATE licenses SET machine_fingerprint = ?, updated_at = ? WHERE id = ?", (fingerprint, now_text(), row["id"]))
    return {"data": {"type": "machines", "id": str(uuid.uuid4()), "attributes": {"fingerprint": fingerprint}}}


@app.get("/api/admin/profiles")
def list_profiles(authorization: str = Header(default="")):
    require_admin(authorization)
    with connect() as db:
        rows = db.execute("SELECT code, label, builtin, created_at FROM profile_options ORDER BY builtin DESC, label COLLATE NOCASE").fetchall()
    return {
        "profiles": [
            {"code": row["code"], "label": row["label"], "builtin": bool(row["builtin"]), "allowed_profile": row["code"], "created_at": row["created_at"]}
            for row in rows
        ]
    }


@app.post("/api/admin/profiles")
def create_profile(payload: ProfileOptionCreate, authorization: str = Header(default="")):
    require_admin(authorization)
    code = re.sub(r"\s+", "_", payload.code.strip().casefold())
    label = payload.label.strip()
    if not re.fullmatch(r"[a-z][a-z0-9_]*", code):
        raise HTTPException(status_code=400, detail="Profile code phải dùng chữ thường, số và dấu gạch dưới; ví dụ ho_guom.")
    if not label:
        raise HTTPException(status_code=400, detail="Tên profile không được để trống.")
    try:
        with connect() as db:
            db.execute(
                "INSERT INTO profile_options (code, label, builtin, created_at) VALUES (?, ?, 0, ?)",
                (code, label, now_text()),
            )
    except sqlite3.IntegrityError as exc:
        raise HTTPException(status_code=409, detail=f"Profile code đã tồn tại: {code}") from exc
    return {"profile": {"code": code, "label": label, "builtin": False, "allowed_profile": code}}


@app.patch("/api/admin/profiles/{code}")
def edit_profile(code: str, payload: ProfileOptionEdit, authorization: str = Header(default="")):
    require_admin(authorization)
    code = code.strip().casefold()
    label = payload.label.strip()
    if not label:
        raise HTTPException(status_code=400, detail="Tên profile không được để trống.")
    with connect() as db:
        row = db.execute("SELECT builtin FROM profile_options WHERE code = ?", (code,)).fetchone()
        if row is None:
            raise HTTPException(status_code=404, detail="Không tìm thấy profile.")
        db.execute("UPDATE profile_options SET label = ? WHERE code = ?", (label, code))
    return {"profile": {"code": code, "label": label, "builtin": bool(row["builtin"]), "allowed_profile": code}}


@app.delete("/api/admin/profiles/{code}")
def delete_profile(code: str, authorization: str = Header(default="")):
    require_admin(authorization)
    code = code.strip().casefold()
    with connect() as db:
        row = db.execute("SELECT builtin FROM profile_options WHERE code = ?", (code,)).fetchone()
        if row is None:
            raise HTTPException(status_code=404, detail="Không tìm thấy profile.")
        if row["builtin"]:
            raise HTTPException(status_code=400, detail="Không thể xóa profile mặc định.")
        db.execute("DELETE FROM profile_options WHERE code = ?", (code,))
    return {"deleted": True, "code": code}

@app.get("/api/admin/licenses")
def list_licenses(authorization: str = Header(default="")):
    require_admin(authorization)
    with connect() as db:
        rows = db.execute("SELECT * FROM licenses ORDER BY created_at DESC").fetchall()
    return {"licenses": [row_json(row) for row in rows]}


@app.post("/api/admin/licenses")
def create_license(payload: LicenseCreate, authorization: str = Header(default="")):
    require_admin(authorization)
    key = payload.key.strip() or license_key()
    identifier = str(uuid.uuid4())
    timestamp = now_text()
    try:
        with connect() as db:
            db.execute("""INSERT INTO licenses (id, license_key, name, expiry, allowed_profiles, allowed_companies, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", (identifier, key, payload.name.strip(), payload.expiry.strip(), json.dumps(clean_list(payload.allowed_profiles)), json.dumps(clean_list(payload.allowed_companies)), timestamp, timestamp))
    except sqlite3.IntegrityError as exc:
        raise HTTPException(status_code=409, detail="License key đã tồn tại.") from exc
    return {"license": row_json(row_by_id(identifier))}


@app.patch("/api/admin/licenses/{license_id}")
def edit_license(license_id: str, payload: LicenseEdit, authorization: str = Header(default="")):
    require_admin(authorization)
    current = row_by_id(license_id)
    if current is None:
        raise HTTPException(status_code=404, detail="License không tồn tại.")
    suspended = current["suspended"] if payload.suspended is None else int(payload.suspended)
    with connect() as db:
        companies = current["allowed_companies"] if payload.allowed_companies is None else json.dumps(clean_list(payload.allowed_companies))
        db.execute("""UPDATE licenses SET name = ?, expiry = ?, suspended = ?, allowed_profiles = ?, allowed_companies = ?, updated_at = ? WHERE id = ?""", (payload.name.strip(), payload.expiry.strip(), suspended, json.dumps(clean_list(payload.allowed_profiles)), companies, now_text(), license_id))
    return {"license": row_json(row_by_id(license_id))}


@app.delete("/api/admin/licenses/{license_id}")
def delete_license(license_id: str, authorization: str = Header(default="")):
    require_admin(authorization)
    with connect() as db:
        result = db.execute("DELETE FROM licenses WHERE id = ?", (license_id,))
    if result.rowcount == 0:
        raise HTTPException(status_code=404, detail="License không tồn tại.")
    return {"deleted": True, "id": license_id}


@app.post("/api/admin/releases")
async def publish_release_bundle(file: UploadFile = File(...), authorization: str = Header(default="")):
    require_admin(authorization)
    if not str(file.filename or "").lower().endswith(".zip"):
        raise HTTPException(
            status_code=400,
            detail="Release phải là bundle ZIP do build_release.ps1 tạo.",
        )
    temporary = RELEASE_DIR / f".velopack-upload-{uuid.uuid4().hex}.zip"
    total_size = 0
    try:
        with temporary.open("wb") as output:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                total_size += len(chunk)
                if total_size > MAX_BUNDLE_BYTES:
                    raise HTTPException(status_code=413, detail="Release bundle vượt quá giới hạn 1.5 GB.")
                output.write(chunk)
        try:
            release = install_release_bundle(temporary, RELEASE_DIR)
        except ReleaseBundleError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {
            "release": release,
            "message": f"Đã publish {release['channel']} {release['version']} lên feed Velopack.",
        }
    finally:
        temporary.unlink(missing_ok=True)


@app.get("/api/admin/releases/status")
def release_status(authorization: str = Header(default="")):
    require_admin(authorization)
    return all_release_status(RELEASE_DIR)

@app.get("/api/update/manifest")
def update_manifest(platform: str = "windows-x64"):
    if not MANIFEST_PATH.exists():
        raise HTTPException(status_code=404, detail="Chưa có release nào trên server.")
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    if manifest.get("platform") and manifest["platform"] != platform:
        raise HTTPException(status_code=404, detail="Không có release cho platform này.")
    return manifest


@app.get("/downloads/{filename}")
def download_release(filename: str):
    target = (RELEASE_DIR / Path(filename).name).resolve()
    if target.parent != RELEASE_DIR.resolve() or not target.exists():
        raise HTTPException(status_code=404, detail="Không tìm thấy release.")
    return FileResponse(target, filename=target.name, media_type="application/octet-stream")


@app.get("/health")
def health():
    return {"ok": True, "service": "local-license-server", "mode": SERVER_MODE, "config_version": 1, "account_id": ACCOUNT_ID, "update_feed": "/updates/product-code-formatter"}

@app.get("/", response_class=HTMLResponse)
def index():
    return (STATIC_DIR / "local_admin_release.html").read_text(encoding="utf-8")


init_db()
app.mount("/updates/product-code-formatter", StaticFiles(directory=UPDATE_FEED_DIR), name="product-code-formatter-updates")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host=os.environ.get("ADMIN_BIND_HOST", "0.0.0.0"), port=int(os.environ.get("ADMIN_PORT", "8080")))
