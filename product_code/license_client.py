"""Keygen client and machine-fingerprint helpers for Product Code Formatter."""

import ipaddress
import json
import os
import re
import sys
import time
import unicodedata
import uuid
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen


LOCAL_KEYGEN_HOSTS = {"localhost", "127.0.0.1", "::1"}
KEYGEN_ACTIVATION_REQUIRED_CODES = {
    "NO_MACHINE",
    "NO_MACHINES",
    "FINGERPRINT_SCOPE_MISMATCH",
    "MACHINE_SCOPE_MISMATCH",
}
DEFAULT_LICENSE_SERVER_IP = os.environ.get("PRODUCT_CODE_FORMATTER_LICENSE_SERVER_IP", "192.168.1.210")
DEFAULT_KEYGEN_SERVER_URL = os.environ.get("PRODUCT_CODE_FORMATTER_KEYGEN_SERVER", f"http://{DEFAULT_LICENSE_SERVER_IP}:8080")
DEFAULT_KEYGEN_ACCOUNT_ID = os.environ.get("PRODUCT_CODE_FORMATTER_KEYGEN_ACCOUNT", "6f1f56e8-3b6f-4a86-9a31-9e0e7f62c001")
DEFAULT_KEYGEN_PUBLIC_HOST = os.environ.get("PRODUCT_CODE_FORMATTER_KEYGEN_HOST", DEFAULT_LICENSE_SERVER_IP)


def raw_text(value):
    return "" if value is None else str(value).strip()


def rm_accents(text):
    text = str(text or "")
    text = text.replace("Đ", "D").replace("đ", "d").replace("Ä", "D").replace("Ä‘", "d").replace("Ã„Â", "D").replace("Ã„â€˜", "d")
    text = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def normalize_rule_key(text):
    text = rm_accents(str(text or "")).casefold().strip()
    return re.sub(r"\s+", " ", text)


def normalize_string_list(value):
    if isinstance(value, str):
        items = re.split(r"[,;\n]+", value)
    elif isinstance(value, list):
        items = value
    else:
        items = []
    result = []
    seen = set()
    for item in items:
        text = str(item or "").strip()
        key = normalize_rule_key(text)
        if not text or not key or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def current_machine_name():
    name = os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or "windows"
    return str(name or "windows").strip() or "windows"


def legacy_machine_fingerprint():
    return f"{current_machine_name()}-{uuid.getnode():012x}".upper()


def windows_machine_guid():
    if sys.platform != "win32":
        return ""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
    except Exception:
        return ""
    guid = re.sub(r"[^A-Za-z0-9-]+", "", raw_text(value)).strip("-")
    return guid.upper()


def local_machine_fingerprint():
    machine_guid = windows_machine_guid()
    if machine_guid:
        return f"{current_machine_name()}-WIN-{machine_guid}".upper()
    return legacy_machine_fingerprint()


def keygen_license_metadata(response):
    data = response.get("data") if isinstance(response, dict) else {}
    attributes = data.get("attributes") if isinstance(data, dict) else {}
    metadata = attributes.get("metadata") if isinstance(attributes, dict) else {}
    return metadata if isinstance(metadata, dict) else {}


def normalize_keygen_server_url(server_url):
    base = str(server_url or DEFAULT_KEYGEN_SERVER_URL).strip().rstrip("/")
    if base and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", base):
        base = f"http://{base}"
    return base


def normalize_keygen_account_id(account_id):
    return DEFAULT_KEYGEN_ACCOUNT_ID


def keygen_allows_http_host(hostname):
    host = str(hostname or "").strip().lower()
    if host in LOCAL_KEYGEN_HOSTS:
        return True
    try:
        address = ipaddress.ip_address(host)
    except ValueError:
        return host.endswith(".local") or "." not in host
    return address.is_private or address.is_loopback or address.is_link_local


def keygen_is_local_http_url(url):
    parsed = urlparse(str(url or ""))
    return parsed.scheme == "http" and keygen_allows_http_host(parsed.hostname)


def keygen_host_header(url):
    parsed = urlparse(str(url or ""))
    if parsed.scheme != "http" or not keygen_allows_http_host(parsed.hostname):
        return ""
    public_host = str(DEFAULT_KEYGEN_PUBLIC_HOST or "").strip()
    if not public_host:
        return ""
    public_parsed = urlparse(public_host if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", public_host) else f"http://{public_host}")
    host = public_parsed.hostname or public_host.strip("/")
    port = public_parsed.port or parsed.port
    return f"{host}:{port}" if port else host


def keygen_forbidden_hint(url):
    parsed = urlparse(str(url or ""))
    host = str(parsed.hostname or "").strip().lower()
    if host in {"127.0.0.1", "localhost", "::1"}:
        return (
            " Keygen dang chan host nay. Hay nhap IP that cua may license server "
            "hoac kiem tra license_server/.env co KEYGEN_HOST=license-server.local."
        )
    return (
        " Kiem tra License server co dung KEYGEN_HOST trong license_server/.env "
        "va Account co dung KEYGEN_ACCOUNT_ID khong."
    )


def keygen_url(server_url, account_id, path):
    base = normalize_keygen_server_url(server_url)
    account = normalize_keygen_account_id(account_id)
    if not base or not account:
        raise ValueError("Can nhap dia chi server va account cua Keygen.")
    parsed = urlparse(base)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("License server phai bat dau bang http:// hoac https://.")
    if parsed.scheme == "http" and not keygen_allows_http_host(parsed.hostname):
        raise ValueError("License server HTTP chi duoc phep cho localhost hoac may trong mang LAN.")
    return f"{base}/v1/accounts/{account}/{path.lstrip('/')}"


def keygen_request(method, url, payload=None, license_key=None, timeout=10, opener=None):
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    headers = {
        "Accept": "application/vnd.api+json",
        "Content-Type": "application/vnd.api+json",
    }
    if keygen_is_local_http_url(url):
        headers["X-Forwarded-Proto"] = "https"
        host_header = keygen_host_header(url)
        if host_header:
            headers["Host"] = host_header
    if license_key:
        headers["Authorization"] = f"License {license_key}"
    request_obj = Request(url, data=data, headers=headers, method=method)
    transport = opener or urlopen
    try:
        with transport(request_obj, timeout=timeout) as response:
            body = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        hint = keygen_forbidden_hint(url) if exc.code == 403 else ""
        raise ValueError(f"Keygen tra ve loi {exc.code}: {detail or exc.reason}.{hint}") from exc
    except URLError as exc:
        raise ValueError(f"Khong ket noi duoc license server: {exc.reason}") from exc
    return json.loads(body) if body else {}


def extract_allowed_companies(metadata):
    if not isinstance(metadata, dict):
        return []
    for key in ["allowed_companies", "allowedCompanies", "companies", "allowed_company_mst", "allowedCompanyMst", "allowed_mst", "allowedMst", "mst", "msts"]:
        if key in metadata:
            return normalize_string_list(metadata.get(key))
    return []


def extract_allowed_profiles(metadata):
    if not isinstance(metadata, dict):
        return []
    for key in ["allowed_profiles", "allowedProfiles", "profiles", "company_profiles", "companyProfiles"]:
        if key in metadata:
            return normalize_string_list(metadata.get(key))
    return []


def extract_supported_profiles(metadata):
    if not isinstance(metadata, dict):
        return []
    for key in ["supported_profiles", "supportedProfiles", "available_profiles", "availableProfiles"]:
        if key in metadata:
            return normalize_string_list(metadata.get(key))
    return []


def keygen_validate_license(server_url, account_id, license_key, fingerprint=None, timeout=10, requester=None, fingerprint_func=None):
    key = str(license_key or "").strip()
    if not key:
        raise ValueError("Can nhap license key.")
    make_fingerprint = fingerprint_func or local_machine_fingerprint
    fingerprint = str(fingerprint or make_fingerprint()).strip()
    payload = {"meta": {"key": key, "scope": {"fingerprint": fingerprint}}}
    request_func = requester or keygen_request
    response = request_func("POST", keygen_url(server_url, account_id, "licenses/actions/validate-key"), payload, timeout=timeout)
    meta = response.get("meta") if isinstance(response, dict) else {}
    data = response.get("data") if isinstance(response, dict) else {}
    attributes = data.get("attributes") if isinstance(data, dict) else {}
    metadata = keygen_license_metadata(response)
    return {
        "valid": meta.get("valid") is not False,
        "code": str(meta.get("code") or ""),
        "detail": str(meta.get("detail") or ""),
        "license_id": str(data.get("id") or "") if isinstance(data, dict) else "",
        "status": str((attributes or {}).get("status") or meta.get("code") or "valid"),
        "product_code": raw_text(metadata.get("product_code") or metadata.get("productCode")),
        "application": raw_text(metadata.get("application")),
        "allowed_companies": extract_allowed_companies(metadata),
        "allowed_profiles": extract_allowed_profiles(metadata),
        "supported_profiles": extract_supported_profiles(metadata),
        "raw": response,
    }


def keygen_activate_machine(server_url, account_id, license_key, license_id, fingerprint=None, timeout=10, requester=None, fingerprint_func=None):
    make_fingerprint = fingerprint_func or local_machine_fingerprint
    fingerprint = str(fingerprint or make_fingerprint()).strip()
    license_id = str(license_id or "").strip()
    if not license_id:
        raise ValueError("Khong tim thay license id de kich hoat may.")
    payload = {
        "data": {
            "type": "machines",
            "attributes": {
                "fingerprint": fingerprint,
                "name": os.environ.get("COMPUTERNAME") or "ProductCodeFormatter",
                "platform": sys.platform,
            },
            "relationships": {"license": {"data": {"type": "licenses", "id": license_id}}},
        }
    }
    try:
        request_func = requester or keygen_request
        response = request_func("POST", keygen_url(server_url, account_id, "machines"), payload, license_key=license_key, timeout=timeout)
    except ValueError as exc:
        if "409" not in str(exc):
            raise
        return ""
    data = response.get("data") if isinstance(response, dict) else {}
    return str(data.get("id") or "") if isinstance(data, dict) else ""


def activate_keygen_license(server_url, account_id, license_key, timeout=10, requester=None, fingerprint_func=None):
    make_fingerprint = fingerprint_func or local_machine_fingerprint
    fingerprint = make_fingerprint()
    request_func = requester or keygen_request
    validation = keygen_validate_license(server_url, account_id, license_key, fingerprint, timeout, requester=request_func, fingerprint_func=make_fingerprint)
    machine_id = ""
    if not validation["valid"]:
        if validation["code"] not in KEYGEN_ACTIVATION_REQUIRED_CODES:
            raise ValueError(validation["detail"] or validation["code"] or "License khong hop le.")
        machine_id = keygen_activate_machine(server_url, account_id, license_key, validation["license_id"], fingerprint, timeout, requester=request_func, fingerprint_func=make_fingerprint)
        validation = keygen_validate_license(server_url, account_id, license_key, fingerprint, timeout, requester=request_func, fingerprint_func=make_fingerprint)
    if not validation["valid"]:
        raise ValueError(validation["detail"] or validation["code"] or "License chua hop le sau khi kich hoat may.")
    return {
        "server_url": normalize_keygen_server_url(server_url),
        "account_id": normalize_keygen_account_id(account_id),
        "license_key": str(license_key or "").strip(),
        "machine_fingerprint": fingerprint,
        "machine_id": machine_id,
        "activated": True,
        "last_validated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "status": validation["status"],
        "product_code": validation["product_code"],
        "application": validation["application"],
        "allowed_companies": validation["allowed_companies"],
        "allowed_profiles": validation["allowed_profiles"],
        "supported_profiles": validation["supported_profiles"],
    }


def public_license_status(config, allows_profile_func, local_activation_func):
    """Return the UI-safe license shape used by the React shell."""
    license_cfg = (config or {}).get("license") or {}
    allowed_profiles = license_cfg.get("allowed_profiles") or []
    activated = local_activation_func(license_cfg)
    return {
        "activated": activated,
        "status": str(license_cfg.get("status") or ("Kích hoạt thành công" if activated else "Chưa kích hoạt")),
        "allowed_profiles": allowed_profiles,
        "allowed_companies": license_cfg.get("allowed_companies") or [],
        "supported_profiles": license_cfg.get("supported_profiles") or [],
        "product_code": str(license_cfg.get("product_code") or ""),
        "application": str(license_cfg.get("application") or ""),
        "vietmax_allowed": activated and (
            allows_profile_func("vietmax", allowed_profiles)
            or allows_profile_func("vietmax_mua_vao", allowed_profiles)
            or allows_profile_func("vietmax_ban_ra", allowed_profiles)
        ),
        "server_url": str(license_cfg.get("server_url") or ""),
        "account_id": str(license_cfg.get("account_id") or ""),
    }
