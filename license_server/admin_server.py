"""Raspberry Pi LAN admin and release server.

The browser talks to this service. The Keygen admin token stays on the Pi and
is never shipped in ProductCodeFormatter.exe.
"""

import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field


APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
RELEASE_DIR = Path(os.environ.get("RELEASE_DIR", str(APP_DIR / "releases")))
MANIFEST_PATH = RELEASE_DIR / "manifest.json"
KEYGEN_URL = os.environ.get("KEYGEN_API_URL", "http://keygen:3000").rstrip("/")
KEYGEN_ACCOUNT_ID = os.environ.get("KEYGEN_ACCOUNT_ID", "6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "").strip()

app = FastAPI(title="Product Code Formatter LAN Admin", version="1.0.0")
RELEASE_DIR.mkdir(parents=True, exist_ok=True)


class LicenseEdit(BaseModel):
    name: str | None = None
    expiry: str | None = None
    suspended: bool | None = None
    allowed_profiles: list[str] = Field(default_factory=list)
    allowed_companies: list[str] = Field(default_factory=list)


def keygen_request(method, path, payload=None):
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=503, detail="ADMIN_TOKEN chưa được cấu hình trên Raspberry Pi.")
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    request = Request(
        f"{KEYGEN_URL}/v1/accounts/{KEYGEN_ACCOUNT_ID}/{path.lstrip('/')}",
        data=body,
        method=method,
        headers={
            "Accept": "application/vnd.api+json",
            "Content-Type": "application/vnd.api+json",
            "Authorization": f"Bearer {ADMIN_TOKEN}",
        },
    )
    try:
        with urlopen(request, timeout=15) as response:
            raw = response.read()
            return json.loads(raw.decode("utf-8")) if raw else {}
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise HTTPException(status_code=exc.code, detail=detail or str(exc.reason)) from exc
    except URLError as exc:
        raise HTTPException(status_code=502, detail=f"Không kết nối được Keygen: {exc.reason}") from exc


def require_admin(authorization: str = ""):
    expected = f"Bearer {ADMIN_TOKEN}" if ADMIN_TOKEN else ""
    if not expected or authorization.strip() != expected:
        raise HTTPException(status_code=401, detail="Cần ADMIN_TOKEN để quản trị license.")


def public_license(item):
    attributes = item.get("attributes") or {}
    metadata = attributes.get("metadata") or {}
    return {
        "id": item.get("id", ""),
        "key": attributes.get("key", ""),
        "name": attributes.get("name") or "",
        "status": attributes.get("status", ""),
        "expiry": attributes.get("expiry") or "",
        "suspended": bool(attributes.get("suspended")),
        "allowed_profiles": metadata.get("allowed_profiles") or [],
        "allowed_companies": metadata.get("allowed_companies") or [],
        "metadata": metadata,
    }


@app.get("/health")
def health():
    return {"status": "ok", "service": "product-code-formatter-admin"}


@app.get("/api/update/manifest")
def update_manifest(platform: str = "windows-x64"):
    if not MANIFEST_PATH.exists():
        raise HTTPException(status_code=404, detail="Chưa có release nào trên server.")
    payload = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    if payload.get("platform") and payload.get("platform") != platform:
        raise HTTPException(status_code=404, detail="Không có release cho platform này.")
    return payload


@app.get("/downloads/{filename}")
def download_release(filename: str):
    target = (RELEASE_DIR / Path(filename).name).resolve()
    if target.parent != RELEASE_DIR.resolve() or not target.exists():
        raise HTTPException(status_code=404, detail="Không tìm thấy release.")
    return FileResponse(target, filename=target.name, media_type="application/octet-stream")


@app.get("/api/admin/licenses")
def list_licenses(authorization: str = Header(default="")):
    require_admin(authorization)
    payload = keygen_request("GET", "licenses?" + urlencode({"limit": 100}))
    return {"licenses": [public_license(item) for item in payload.get("data", [])]}


@app.patch("/api/admin/licenses/{license_id}")
def edit_license(license_id: str, payload: LicenseEdit, authorization: str = Header(default="")):
    require_admin(authorization)
    attributes = {
        "metadata": {
            "allowed_profiles": [str(value).strip() for value in payload.allowed_profiles if str(value).strip()],
            "allowed_companies": [str(value).strip() for value in payload.allowed_companies if str(value).strip()],
        },
    }
    if payload.name is not None:
        attributes["name"] = payload.name.strip()
    if payload.expiry is not None:
        attributes["expiry"] = payload.expiry.strip() or None
    if payload.suspended is not None:
        attributes["suspended"] = payload.suspended
    result = keygen_request("PATCH", f"licenses/{license_id}", {"data": {"type": "licenses", "attributes": attributes}})
    return {"license": public_license(result.get("data") or {})}


@app.delete("/api/admin/licenses/{license_id}")
def delete_license(license_id: str, authorization: str = Header(default="")):
    require_admin(authorization)
    keygen_request("DELETE", f"licenses/{license_id}")
    return {"deleted": True, "id": license_id}


@app.get("/", response_class=HTMLResponse)
def index():
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host=os.environ.get("ADMIN_BIND_HOST", "0.0.0.0"), port=int(os.environ.get("ADMIN_PORT", "8080")))
