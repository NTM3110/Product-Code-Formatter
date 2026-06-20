import json
import os
import re
import sys
import time
import uuid
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


APP_NAME = "Inventory Allocator"
PRODUCT_CODE = "inventory-allocator"
PRODUCT_GROUP = "InventoryAllocator"
LOCAL_KEYGEN_HOSTS = {"localhost", "127.0.0.1", "::1"}
KEYGEN_ACTIVATION_REQUIRED_CODES = {
    "NO_MACHINE",
    "NO_MACHINES",
    "FINGERPRINT_SCOPE_MISMATCH",
    "MACHINE_SCOPE_MISMATCH",
}


def app_data_dir():
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home())
    path = base / PRODUCT_GROUP
    path.mkdir(parents=True, exist_ok=True)
    return path


CONFIG_PATH = app_data_dir() / "inventory_allocator_config.json"


def empty_license_config():
    return {
        "server_url": "",
        "account_id": "",
        "license_key": "",
        "machine_fingerprint": "",
        "machine_id": "",
        "activated": False,
        "last_validated_at": "",
        "status": "",
        "product_code": PRODUCT_CODE,
    }


def default_config():
    return {
        "app_name": APP_NAME,
        "product_code": PRODUCT_CODE,
        "license": empty_license_config(),
    }


def normalize_config(value):
    value = value if isinstance(value, dict) else {}
    result = default_config()
    license_value = value.get("license") if isinstance(value.get("license"), dict) else {}
    license_value = license_value if isinstance(license_value, dict) else {}
    license_cfg = empty_license_config()
    for key in ["server_url", "account_id", "license_key", "machine_fingerprint", "machine_id", "last_validated_at", "status"]:
        license_cfg[key] = str(license_value.get(key) or "").strip()
    license_cfg["activated"] = bool(license_value.get("activated"))
    license_cfg["product_code"] = PRODUCT_CODE
    result["license"] = license_cfg
    return result


def load_config():
    if not CONFIG_PATH.exists():
        return default_config()
    try:
        return normalize_config(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        return default_config()


def save_config(config):
    normalized = normalize_config(config)
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")
    return normalized


def normalize_rule_key(value):
    text = str(value or "").strip().casefold()
    text = re.sub(r"[^a-z0-9_-]+", "-", text)
    return text.strip("-")


def local_machine_fingerprint():
    return f"{os.environ.get('COMPUTERNAME') or 'windows'}-{uuid.getnode():012x}".upper()


def keygen_url(server_url, account_id, path):
    base = str(server_url or "").strip().rstrip("/")
    account = str(account_id or "").strip().strip("/")
    if not base or not account:
        raise ValueError("Cần nhập địa chỉ server và account của Keygen.")
    parsed = urlparse(base)
    if parsed.scheme != "https" and parsed.hostname not in LOCAL_KEYGEN_HOSTS:
        raise ValueError("License server phải dùng HTTPS. Chỉ localhost được phép dùng HTTP để thử nghiệm.")
    return f"{base}/v1/accounts/{account}/{path.lstrip('/')}"


def keygen_is_local_http_url(url):
    parsed = urlparse(str(url or ""))
    return parsed.scheme == "http" and parsed.hostname in LOCAL_KEYGEN_HOSTS


def keygen_request(method, url, payload=None, license_key=None, timeout=10):
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    headers = {
        "Accept": "application/vnd.api+json",
        "Content-Type": "application/vnd.api+json",
    }
    if keygen_is_local_http_url(url):
        headers["X-Forwarded-Proto"] = "https"
    if license_key:
        headers["Authorization"] = f"License {license_key}"
    request_obj = Request(url, data=data, headers=headers, method=method)
    try:
        with urlopen(request_obj, timeout=timeout) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise ValueError(f"Keygen trả về lỗi {exc.code}: {detail or exc.reason}") from exc
    except URLError as exc:
        raise ValueError(f"Không kết nối được license server: {exc.reason}") from exc
    return json.loads(body) if body else {}


def keygen_license_metadata(response):
    data = response.get("data") if isinstance(response, dict) else {}
    attributes = data.get("attributes") if isinstance(data, dict) else {}
    metadata = attributes.get("metadata") if isinstance(attributes, dict) else {}
    return metadata if isinstance(metadata, dict) else {}


def metadata_matches_inventory_product(metadata):
    if not isinstance(metadata, dict):
        return False
    candidates = [
        metadata.get("product_code"),
        metadata.get("productCode"),
        metadata.get("product"),
        metadata.get("application"),
        metadata.get("app"),
    ]
    normalized = {normalize_rule_key(item) for item in candidates if item}
    return PRODUCT_CODE in normalized or normalize_rule_key(PRODUCT_GROUP) in normalized


def keygen_validate_license(server_url, account_id, license_key, fingerprint=None, timeout=10):
    key = str(license_key or "").strip()
    if not key:
        raise ValueError("Cần nhập license key.")
    fingerprint = str(fingerprint or local_machine_fingerprint()).strip()
    payload = {"meta": {"key": key, "scope": {"fingerprint": fingerprint}}}
    response = keygen_request("POST", keygen_url(server_url, account_id, "licenses/actions/validate-key"), payload, timeout=timeout)
    meta = response.get("meta") if isinstance(response, dict) else {}
    meta = meta if isinstance(meta, dict) else {}
    data = response.get("data") if isinstance(response, dict) else {}
    data = data if isinstance(data, dict) else {}
    attributes = data.get("attributes") if isinstance(data.get("attributes"), dict) else {}
    metadata = keygen_license_metadata(response)
    product_ok = metadata_matches_inventory_product(metadata)
    return {
        "valid": meta.get("valid") is not False and product_ok,
        "code": str(meta.get("code") or ""),
        "detail": str(meta.get("detail") or ""),
        "license_id": str(data.get("id") or ""),
        "status": str((attributes or {}).get("status") or meta.get("code") or "valid"),
        "metadata": metadata,
        "product_ok": product_ok,
        "raw": response,
    }


def keygen_activate_machine(server_url, account_id, license_key, license_id, fingerprint=None, timeout=10):
    fingerprint = str(fingerprint or local_machine_fingerprint()).strip()
    license_id = str(license_id or "").strip()
    if not license_id:
        raise ValueError("Không tìm thấy license id để kích hoạt máy.")
    payload = {
        "data": {
            "type": "machines",
            "attributes": {
                "fingerprint": fingerprint,
                "name": os.environ.get("COMPUTERNAME") or PRODUCT_GROUP,
                "platform": sys.platform,
            },
            "relationships": {"license": {"data": {"type": "licenses", "id": license_id}}},
        }
    }
    try:
        response = keygen_request("POST", keygen_url(server_url, account_id, "machines"), payload, license_key=license_key, timeout=timeout)
    except ValueError as exc:
        if "409" not in str(exc):
            raise
        return ""
    data = response.get("data") if isinstance(response, dict) else {}
    return str(data.get("id") or "") if isinstance(data, dict) else ""


def activate_keygen_license(server_url, account_id, license_key, timeout=10):
    fingerprint = local_machine_fingerprint()
    validation = keygen_validate_license(server_url, account_id, license_key, fingerprint, timeout)
    if not validation["product_ok"]:
        raise ValueError("License này không thuộc Inventory Allocator.")
    machine_id = ""
    if not validation["valid"]:
        if validation["code"] not in KEYGEN_ACTIVATION_REQUIRED_CODES:
            raise ValueError(validation["detail"] or validation["code"] or "License không hợp lệ.")
        machine_id = keygen_activate_machine(server_url, account_id, license_key, validation["license_id"], fingerprint, timeout)
        validation = keygen_validate_license(server_url, account_id, license_key, fingerprint, timeout)
    if not validation["valid"]:
        raise ValueError(validation["detail"] or validation["code"] or "License chưa hợp lệ sau khi kích hoạt máy.")
    return {
        "server_url": str(server_url or "").strip().rstrip("/"),
        "account_id": str(account_id or "").strip(),
        "license_key": str(license_key or "").strip(),
        "machine_fingerprint": fingerprint,
        "machine_id": machine_id,
        "activated": True,
        "last_validated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "status": validation["status"],
        "product_code": PRODUCT_CODE,
    }
