import json
import os
import re
import subprocess
import sys
import threading
import time
import unicodedata
import uuid
import webbrowser
import zipfile
from collections import Counter
from copy import copy
from io import BytesIO
from pathlib import Path
from urllib.request import urlopen

import numpy as np
import pandas as pd
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask.json.provider import DefaultJSONProvider

APP_VERSION = "0.1"

class CustomJSONProvider(DefaultJSONProvider):
    def default(self, o):
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        if isinstance(o, np.ndarray):
            return o.tolist()
        if isinstance(o, pd.Timestamp):
            return str(o)
        return super().default(o)


if getattr(sys, "frozen", False):
    BASE_DIR = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "ProductCodeFormatter"
    STATIC_DIR = Path(getattr(sys, "_MEIPASS", str(BASE_DIR))) / "static"
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", str(BASE_DIR)))
else:
    BASE_DIR = Path(__file__).resolve().parent
    STATIC_DIR = BASE_DIR / "static"
    RESOURCE_DIR = BASE_DIR

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
CONFIG_PATH = BASE_DIR / "product_code_config.json"
ICON_PATH = RESOURCE_DIR / "app_icon.ico"
BASE_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder=None)
app.json_provider_class = CustomJSONProvider
app.json = CustomJSONProvider(app)

if not getattr(sys, "frozen", False):
    try:
        from flask_cors import CORS

        CORS(app)
    except ImportError:
        pass


PROFILE_LABELS = {
    "son_phuong": "S\u01a1n Ph\u01b0\u01a1ng",
    "cao_thanh": "Cao Th\u00e0nh",
    "quang_thinh": "Quang Th\u1ecbnh",
    "vietmax": "Vietmax",
}

PROFILE_ALIASES = {
    "quang_thinh_1": "quang_thinh",
    "quang_thinh_2": "quang_thinh",
}

MAX_CODE_LENGTH = 50
DIAMETER_CHARS = "\u03a6\u03c6\u03d5\u00d8\u00f8\u2205\u2300\u0424\u0444\uff06"
DEFAULT_INVOICE_STATUS_COL = "AJ"
DEFAULT_INVOICE_STATUS_SKIP_VALUES = [
    "Hóa đơn đã bị điều chỉnh",
    "Hóa đơn bị thay thế",
    "Hóa đơn đã bị thay thế",
    "Hóa đơn đã bị hủy",
]
IGNORED_INVOICE_STATUSES = {
    "hoa don da bi dieu chinh",
    "hoa don bi thay the",
    "hoa don da bi thay the",
    "hoa don da bi huy",
}


def empty_profile_config(profile_key_name=None):
    return {
        "prefixes": {},
        "selected_products": {},
        "removed_companies": {},
        "word_rules": {"\u0111en": "DEN", "t\u00f4n": "TON"},
        "first_word_rules": {},
        "repeated_phrase_removals": ["inox"] if profile_key_name == "cao_thanh" else [],
        "price_group_rules": {},
        "price_range_rules": {},
        "price_adjust_all_percent": 0,
        "manual_code_overrides": {},
        "include_company_prefix": True,
        "output_path": "",
        "columns": {},
    }


def default_config():
    return {
        "app_version": APP_VERSION,
        "selected_profile": "son_phuong",
        "profiles": {key: empty_profile_config(key) for key in PROFILE_LABELS},
        "columns": {
            "company_col": "F",
            "mst_col": "G",
            "address_col": "H",
            "product_col": "M",
            "qty_col": "O",
            "price_col": "",
            "output_col": "L",
            "invoice_status_col": DEFAULT_INVOICE_STATUS_COL,
            "invoice_status_skip_values": DEFAULT_INVOICE_STATUS_SKIP_VALUES[:],
        },
    }


def profile_key(value):
    key = PROFILE_ALIASES.get(value, value)
    return key if key in PROFILE_LABELS else "son_phuong"


def normalize_removed_companies(value):
    if isinstance(value, dict):
        return {str(mst): bool(enabled) for mst, enabled in value.items() if enabled}
    if isinstance(value, list):
        return {str(mst): True for mst in value if str(mst).strip()}
    return {}


def normalize_products_map(value):
    result = {}
    if not isinstance(value, dict):
        return result
    for key, code in value.items():
        key = str(key)
        if "|||" in key:
            result[key] = str(code or "")
        elif "|" in key:
            mst, product = key.split("|", 1)
            result[product_key(mst, product)] = str(code or "")
    return result


def legacy_word_rules(profile_key_name, profile):
    formula_options = profile.get("formula_options") or {}
    keep_words = formula_options.get(f"{profile_key_name}_keep_words")
    if keep_words is None and profile_key_name == "quang_thinh":
        keep_words = formula_options.get("son_phuong_keep_words")
    return keep_words


def normalize_price_group_rules(value, products=None):
    result = {}
    products = products or {}
    if not isinstance(value, dict):
        return result
    for key, rule in value.items():
        if not isinstance(rule, dict):
            continue
        normalized_key = str(key)
        if "|||" not in normalized_key and "|" in normalized_key:
            mst, product = normalized_key.split("|", 1)
            normalized_key = product_key(mst, product)
        try:
            min_price = float(rule.get("min_price"))
            max_price = float(rule.get("max_price"))
            percent = float(rule.get("percent", 8))
        except Exception:
            continue
        if min_price > max_price:
            min_price, max_price = max_price, min_price
        base_code = rule.get("base_code") or products.get(normalized_key) or ""
        groups = []
        for group in rule.get("groups") or []:
            if not isinstance(group, dict):
                continue
            try:
                group_min = float(group.get("min_price"))
                group_max = float(group.get("max_price"))
                adjust_percent = float(group.get("adjust_percent", 0))
            except Exception:
                continue
            if group_min > group_max:
                group_min, group_max = group_max, group_min
            groups.append({
                "index": int(group.get("index") or len(groups) + 1),
                "label": str(group.get("label") or f"Nhóm {len(groups) + 1}"),
                "min_price": group_min,
                "max_price": group_max,
                "average_price": float(group.get("average_price")) if group.get("average_price") is not None else None,
                "adjust_percent": adjust_percent,
            })
        result[normalized_key] = {
            "base_code": str(base_code),
            "min_price": min_price,
            "max_price": max_price,
            "percent": percent,
            "groups": groups,
        }
    return result


def normalize_phrase_list(value):
    result = []
    seen = set()
    if not isinstance(value, list):
        return result
    for item in value:
        phrase = str(item or "").strip()
        key = normalize_rule_key(phrase)
        if not phrase or not key or key in seen:
            continue
        seen.add(key)
        result.append(phrase)
    return result


def price_ranges_from_groups(price_group_rules):
    ranges = {}
    for rule in price_group_rules.values():
        base_code = str(rule.get("base_code") or "").strip()
        if not base_code:
            continue
        ranges[base_code] = {
            "min_price": rule.get("min_price"),
            "max_price": rule.get("max_price"),
            "percent": rule.get("percent", 8),
        }
    return normalize_price_range_rules(ranges)


def legacy_price_adjust_all_percent(profile, price_groups, price_ranges):
    if "price_adjust_all_percent" in profile:
        return float(profile.get("price_adjust_all_percent") or 0)
    margins = set()
    rule_sets = [price_ranges.values(), price_groups.values()]
    for rules in rule_sets:
        for rule in rules:
            for group in rule.get("groups") or []:
                try:
                    margins.add(float(group.get("adjust_percent") or 0))
                except Exception:
                    continue
    return margins.pop() if len(margins) == 1 else 0


def normalize_profile_config(profile_key_name, profile):
    defaults = empty_profile_config(profile_key_name)
    profile = profile if isinstance(profile, dict) else {}
    prefixes = profile.get("prefixes") or profile.get("companies") or {}
    products = normalize_products_map(profile.get("manual_code_overrides") or profile.get("products") or {})
    word_rules_source = profile.get("word_rules")
    first_word_rules_source = profile.get("first_word_rules") or {}
    if word_rules_source is None:
        word_rules_source = legacy_word_rules(profile_key_name, profile)
    if word_rules_source is None:
        word_rules_source = defaults["word_rules"]
    price_groups = normalize_price_group_rules(profile.get("price_group_rules") or {}, products)
    price_ranges = normalize_price_range_rules(profile.get("price_range_rules") or {})
    if not price_ranges:
        price_ranges = price_ranges_from_groups(price_groups)
    return {
        "prefixes": {str(mst): str(prefix) for mst, prefix in dict(prefixes).items()} if isinstance(prefixes, dict) else {},
        "selected_products": dict(profile.get("selected_products") or {}),
        "removed_companies": normalize_removed_companies(profile.get("removed_companies")),
        "word_rules": normalize_word_rules(word_rules_source),
        "first_word_rules": normalize_word_rules(first_word_rules_source),
        "repeated_phrase_removals": normalize_phrase_list(profile.get("repeated_phrase_removals", defaults["repeated_phrase_removals"])),
        "price_group_rules": price_groups,
        "price_range_rules": price_ranges,
        "price_adjust_all_percent": legacy_price_adjust_all_percent(profile, price_groups, price_ranges),
        "manual_code_overrides": products,
        "include_company_prefix": profile.get("include_company_prefix") is not False,
        "output_path": str(profile.get("output_path") or ""),
        "columns": dict(profile.get("columns") or {}) if isinstance(profile.get("columns"), dict) else {},
    }


def normalize_config(data):
    cfg = default_config()
    cfg["app_version"] = APP_VERSION
    if isinstance(data, dict):
        selected = data.get("selected_profile") or data.get("active_profile") or data.get("format_rule")
        cfg["selected_profile"] = profile_key(selected) if selected else cfg["selected_profile"]
        cfg["columns"].update(data.get("columns") or {})
        cfg["columns"].setdefault("invoice_status_col", DEFAULT_INVOICE_STATUS_COL)
        if not isinstance(cfg["columns"].get("invoice_status_skip_values"), list):
            cfg["columns"]["invoice_status_skip_values"] = DEFAULT_INVOICE_STATUS_SKIP_VALUES[:]
        profiles = data.get("profiles") or {}
        if not isinstance(profiles, dict):
            profiles = {}
        if not profiles and any(key in data for key in ["companies", "products", "removed_companies", "formula_options", "price_group_rules"]):
            profiles = {cfg["selected_profile"]: data}
        for alias, target in PROFILE_ALIASES.items():
            if alias in profiles and target not in profiles:
                profiles[target] = profiles[alias]
        for key in PROFILE_LABELS:
            cfg["profiles"][key].update(normalize_profile_config(key, profiles.get(key) or {}))
    if cfg["selected_profile"] not in PROFILE_LABELS:
        cfg["selected_profile"] = "son_phuong"
    cfg["columns"].setdefault("invoice_status_col", DEFAULT_INVOICE_STATUS_COL)
    if not isinstance(cfg["columns"].get("invoice_status_skip_values"), list):
        cfg["columns"]["invoice_status_skip_values"] = DEFAULT_INVOICE_STATUS_SKIP_VALUES[:]
    return cfg


def load_config():
    if not CONFIG_PATH.exists():
        return default_config()
    try:
        return normalize_config(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except Exception:
        return default_config()


def save_config(cfg):
    cfg = normalize_config(cfg)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return cfg


def normalize_word_rules(value):
    if isinstance(value, list):
        result = {}
        for item in value:
            if isinstance(item, dict):
                word = str(item.get("word") or "").strip()
                output = str(item.get("output") or "").strip()
            else:
                word = str(item or "").strip()
                output = ""
            if word:
                result[word] = normalize_token(output or word)
        return result
    if isinstance(value, dict):
        result = {}
        for word, output in value.items():
            word = "đen" if str(word).upper() == "DEN" else "tôn" if str(word).upper() == "TON" else str(word).strip()
            if word:
                result[word] = normalize_token(str(output or word))
        return result
    return {}


def normalize_price_range_rules(value):
    result = {}
    if not isinstance(value, dict):
        return result
    for key, rule in value.items():
        if not isinstance(rule, dict):
            continue
        try:
            min_price = float(rule.get("min_price"))
            max_price = float(rule.get("max_price"))
            percent = float(rule.get("percent", 8))
        except Exception:
            continue
        if min_price > max_price:
            min_price, max_price = max_price, min_price
        groups = []
        for group in rule.get("groups") or []:
            if not isinstance(group, dict):
                continue
            try:
                group_min = float(group.get("min_price"))
                group_max = float(group.get("max_price"))
                adjust_percent = float(group.get("adjust_percent", 0))
            except Exception:
                continue
            if group_min > group_max:
                group_min, group_max = group_max, group_min
            groups.append({
                "index": int(group.get("index") or len(groups) + 1),
                "label": str(group.get("label") or f"Nhóm {len(groups) + 1}"),
                "min_price": group_min,
                "max_price": group_max,
                "average_price": float(group.get("average_price")) if group.get("average_price") is not None else None,
                "adjust_percent": adjust_percent,
            })
        result[str(key)] = {
            "min_price": min_price,
            "max_price": max_price,
            "percent": percent,
            "groups": groups,
        }
    return result


@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify(error=str(e)), 500


def excel_col_to_index(col):
    col = str(col).strip().upper()
    if not re.fullmatch(r"[A-Z]+", col):
        raise ValueError(f"Invalid Excel column: {col}")
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n - 1


def index_to_excel_col(idx):
    idx += 1
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def rm_accents(text):
    text = str(text)
    text = text.replace("Đ", "D").replace("đ", "d").replace("Ä", "D").replace("Ä‘", "d")
    for ch in DIAMETER_CHARS:
        text = text.replace(ch, "F")
    text = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def normalize_token(text):
    text = rm_accents(text).upper()
    text = re.sub(r"[^A-Z0-9.]+", "", text)
    return text


def normalize_code_token(text, keep_slash=False, keep_hyphen=False):
    text = rm_accents(text).upper()
    allowed = "A-Z0-9."
    if keep_slash:
        allowed += "/"
    if keep_hyphen:
        allowed += "-"
    return re.sub(rf"[^{allowed}]+", "", text)


def is_upper_code_token(text):
    raw = rm_accents(str(text or "")).strip()
    return bool(re.fullmatch(r"[A-Z0-9./-]+", raw) and re.search(r"[A-Z]", raw))


def normalize_rule_key(text):
    text = rm_accents(str(text or "")).casefold().strip()
    return re.sub(r"\s+", " ", text)


VIETNAM_LOCATION_PHRASES = sorted(
    [
        tuple(normalize_token(word) for word in name.split())
        for name in [
            "An Giang", "Bà Rịa Vũng Tàu", "Bắc Giang", "Bắc Kạn", "Bạc Liêu", "Bắc Ninh", "Bến Tre",
            "Bình Dương", "Bình Định", "Bình Phước", "Bình Thuận", "Cà Mau", "Cao Bằng", "Cần Thơ",
            "Đà Nẵng", "Đắk Lắk", "Đắk Nông", "Điện Biên", "Đồng Nai", "Đồng Tháp", "Gia Lai", "Hà Giang",
            "Hà Nam", "Hà Nội", "Hà Tĩnh", "Hải Dương", "Hải Phòng", "Hậu Giang", "Hòa Bình", "Hồ Chí Minh",
            "Hưng Yên", "Khánh Hòa", "Kiên Giang", "Kon Tum", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai",
            "Long An", "Nam Định", "Nghệ An", "Ninh Bình", "Ninh Thuận", "Phú Thọ", "Phú Yên", "Quảng Bình",
            "Quảng Nam", "Quảng Ngãi", "Quảng Ninh", "Quảng Trị", "Sóc Trăng", "Sơn La", "Tây Ninh",
            "Thái Bình", "Thái Nguyên", "Thanh Hóa", "Thành phố Huế", "Thừa Thiên Huế", "Huế", "Tiền Giang",
            "Trà Vinh", "Tuyên Quang", "Vĩnh Long", "Vĩnh Phúc", "Yên Bái",
        ]
    ],
    key=len,
    reverse=True,
)


def remove_company_location_phrases(tokens):
    result = []
    index = 0
    while index < len(tokens):
        phrase = next(
            (items for items in VIETNAM_LOCATION_PHRASES if tuple(tokens[index:index + len(items)]) == items),
            None,
        )
        if phrase:
            index += len(phrase)
        else:
            result.append(tokens[index])
            index += 1
    return result


def suggest_prefix(company):
    words = re.sub(r"[^A-Za-z0-9À-ỹĐđ ]+", " ", str(company)).split()
    skip = {"CONG", "TY", "TNHH", "TM", "DV", "CP", "CO", "LTD", "MTV", "THUONG", "MAI"}
    meaningful = remove_company_location_phrases([normalize_token(w) for w in words if normalize_token(w)])
    meaningful = [word for word in meaningful if word not in skip]
    tail_words = meaningful[-2:]
    prefix = "".join(w[:1] for w in tail_words)
    return prefix or normalize_token(company)[:2]


def normalize_sep(text):
    return re.sub(r"(?<=\d)\s*([xX*])\s*(?=\d)", "x", str(text).strip())


def norm_num(value, is_last_numeric=False):
    s = str(value).replace(",", ".")
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(f).rstrip("0").rstrip(".")
    except Exception:
        return s


def extract_dimensions(name):
    text = normalize_sep(name)
    m = re.search(r"(\d+(?:[.,]\d+)?)x(\d+(?:[.,]\d+)?)(?:x(\d+(?:[.,]\d+)?))?", text, re.I)
    if not m:
        return ""
    parts = [p for p in m.groups() if p]
    return "x".join(norm_num(p) for p in parts)


def has_number(text):
    return bool(re.search(r"\d", str(text)))


def simple_number_text(text):
    text = normalize_sep(text)
    text = rm_accents(text).upper()
    text = re.sub(r"\b(MM|M|CM|C)\b", "", text)
    text = re.sub(r"[^A-Z0-9.x,]+", "", text)
    return text.replace(",", ".")


def code_words(name):
    text = str(name or "")
    for ch in DIAMETER_CHARS:
        text = text.replace(ch, "F")
    text = normalize_sep(text)
    return [w for w in re.split(r"\s+", text.strip()) if w]


def trim_code(value):
    value = re.sub(r"\.+", ".", str(value or "")).strip(". ")
    return value[:MAX_CODE_LENGTH].rstrip(".")


def is_volume_token(token):
    token = str(token or "").strip()
    return bool(re.fullmatch(r"\d+(?:[,.]\d+)?[lL]", token))


def is_dimension_token(token):
    token = normalize_sep(token)
    return bool(re.search(r"\d", token)) or token.lower() == "x"


def son_phuong_dimension_pair(words, idx):
    cur = normalize_token(words[idx])
    nxt = normalize_token(words[idx + 1]) if idx + 1 < len(words) else ""
    if cur == "C" and re.fullmatch(r"\d+(?:MM)?", nxt):
        return "Cx" + re.sub(r"MM$", "", nxt)
    return ""


def should_process_qty(v):
    if pd.isna(v):
        return False
    s = str(v).strip()
    if not s:
        return False
    try:
        return float(s.replace(",", ".")) > 0
    except Exception:
        return False


def parse_price(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s:
        if s.rfind(".") > s.rfind(","):
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) == 3 and len(parts) > 1:
            s = "".join(parts)
        else:
            s = s.replace(",", ".")
    else:
        s = s
    try:
        return float(s)
    except Exception:
        return None


def raw_text(value):
    return "" if pd.isna(value) else str(value).strip()


def normalized_invoice_status(value):
    return rm_accents(raw_text(value)).casefold()


def normalized_invoice_status_set(values):
    if not isinstance(values, list):
        return IGNORED_INVOICE_STATUSES
    selected = {normalized_invoice_status(value) for value in values if raw_text(value)}
    return selected


def ignored_invoice_status(value, skip_statuses=None):
    status = rm_accents(raw_text(value)).casefold()
    return status in normalized_invoice_status_set(skip_statuses)


def parse_quantity(value):
    parsed = parse_price(value)
    return parsed if parsed is not None and parsed > 0 else 1


def fmt_price(value):
    if value is None:
        return ""
    if abs(value - int(value)) < 0.000001:
        return f"{int(value):,}"
    return f"{value:,.2f}".rstrip("0").rstrip(".")


def cell(df, row, col):
    return "" if col < 0 or col >= df.shape[1] else df.iat[row, col]


def word_piece(token, word_rules, keep_numeric=True, keep_liter=False, default_len=1, preserve_upper_code=False, keep_slash=False, keep_hyphen=False):
    key = normalize_rule_key(token)
    rule_key = next((rule for rule in word_rules if normalize_rule_key(rule) == key), None)
    if rule_key is not None:
        return normalize_token(word_rules[rule_key])
    compact = normalize_code_token(token, keep_slash=keep_slash, keep_hyphen=keep_hyphen)
    if preserve_upper_code and is_upper_code_token(token):
        return compact
    if keep_numeric and has_number(token):
        return compact
    if keep_liter and re.search(r"\d+\s*[lL]\b|[lL]\s*$", str(token)):
        return compact
    return compact[:default_len]


def phrase_rule_piece(words, start, word_rules):
    if not word_rules:
        return None
    max_len = max((len(code_words(rule)) for rule in word_rules), default=0)
    max_len = min(max_len, len(words) - start)
    for length in range(max_len, 0, -1):
        phrase = " ".join(words[start:start + length])
        rule_key = next((rule for rule in word_rules if normalize_rule_key(rule) == normalize_rule_key(phrase)), None)
        if rule_key is not None:
            return normalize_token(word_rules[rule_key]), length
    return None


def word_pieces(words, word_rules, fallback):
    parts = []
    i = 0
    while i < len(words):
        matched = phrase_rule_piece(words, i, word_rules)
        if matched:
            part, length = matched
            parts.append(part)
            i += length
            continue
        parts.append(fallback(words[i]))
        i += 1
    return parts


def remove_repeated_phrases(words, repeated_phrases):
    phrase_words = [code_words(phrase) for phrase in normalize_phrase_list(repeated_phrases)]
    phrase_words = [items for items in phrase_words if items]
    if not phrase_words:
        return words
    phrase_words.sort(key=len, reverse=True)
    seen = set()
    result = []
    i = 0
    while i < len(words):
        matched = None
        for items in phrase_words:
            length = len(items)
            if length > len(words) - i:
                continue
            phrase = " ".join(words[i:i + length])
            if normalize_rule_key(phrase) == normalize_rule_key(" ".join(items)):
                matched = (phrase, length)
                break
        if matched:
            phrase, length = matched
            key = normalize_rule_key(phrase)
            if key not in seen:
                result.extend(words[i:i + length])
                seen.add(key)
            i += length
        else:
            result.append(words[i])
            i += 1
    return result


def normalize_cao_thanh_inox_grade_words(words):
    result = []
    for word in words:
        current = normalize_code_token(word)
        previous = normalize_code_token(result[-1]) if result else ""
        if current == "304" and previous == "INOX":
            continue
        result.append(word)
    return result


def normalize_cao_thanh_con_reducer_words(words):
    if not words:
        return words
    first = normalize_code_token(words[0])
    last = normalize_code_token(words[-1], keep_slash=True)
    if first != "CON" or not re.search(r"\d+(?:[.,]\d+)?/\d+", last):
        return words
    if any(normalize_code_token(word) == "THU" for word in words):
        return words
    return [words[0], "thu", *words[1:]]


def make_product_part(profile, product, word_rules, first_word_rules=None, repeated_phrase_removals=None):
    first_word_rules = first_word_rules or {}
    name = "" if pd.isna(product) else str(product).strip()

    if profile == "cao_thanh":
        name = re.sub(r"\([^)]*\)", " ", name)
        words = remove_repeated_phrases(code_words(name), repeated_phrase_removals)
        words = normalize_cao_thanh_inox_grade_words(words)
        words = normalize_cao_thanh_con_reducer_words(words)
        first = word_pieces(
            words[:2],
            first_word_rules,
            lambda w: word_piece(w, {}, keep_numeric=True, default_len=len(str(w)), preserve_upper_code=True, keep_slash=True),
        )
        rest = word_pieces(
            words[2:],
            word_rules,
            lambda w: word_piece(w, {}, keep_numeric=True, preserve_upper_code=True, keep_slash=True),
        )
        return "".join(first + rest)

    words = remove_repeated_phrases(code_words(name), repeated_phrase_removals)

    if profile == "vietmax":
        parts = []
        for index, word in enumerate(words):
            parts.append(word_piece(word, {}, keep_numeric=True, default_len=(len(str(word)) if index == 0 else 2), preserve_upper_code=True, keep_hyphen=True))
        return "".join(p for p in parts if p)

    if profile == "quang_thinh":
        filtered = []
        for w in words:
            if normalize_token(w) == "HANG":
                break
            if normalize_token(w) == "SON":
                continue
            filtered.append(w)
        parts = []
        i = 0
        while i < len(filtered):
            matched = phrase_rule_piece(filtered, i, word_rules)
            if matched:
                part, length = matched
                parts.append(part)
                i += length
                continue
            w = filtered[i]
            if is_volume_token(w):
                parts.append(normalize_token(w))
            else:
                parts.append(word_piece(w, {}, keep_numeric=True, default_len=2))
            i += 1
        return ".".join(p for p in parts if p)

    parts = []
    i = 0
    while i < len(words):
        paired = son_phuong_dimension_pair(words, i)
        if paired:
            parts.append(paired)
            i += 2
            continue
        w = words[i]
        if normalize_token(w) == "X":
            parts.append("x")
        elif is_dimension_token(w):
            parts.append(normalize_token(w))
        else:
            matched = phrase_rule_piece(words, i, word_rules)
            if matched:
                part, length = matched
                parts.append(part)
                i += length
                continue
            parts.append(word_piece(w, {}, keep_numeric=False))
        i += 1
    return "".join(parts)


def make_code(mst, product, qty, prefix_map, profile, word_rules, first_word_rules=None, require_qty=True, include_company_prefix=True, repeated_phrase_removals=None):
    if require_qty and not should_process_qty(qty):
        return ""
    mst = "" if pd.isna(mst) else str(mst).strip()
    if include_company_prefix and (not mst or mst not in prefix_map):
        return ""
    body = make_product_part(profile, product, word_rules, first_word_rules, repeated_phrase_removals)
    if not include_company_prefix:
        return trim_code(body)
    prefix = normalize_token(prefix_map[mst])
    return trim_code(f"{prefix}.{body}")


def raw_price_group(price, rule):
    min_price = float(rule.get("min_price") or 0)
    percent = float(rule.get("percent") or 8)
    if price is None or min_price <= 0 or percent <= 0:
        return None
    step = min_price * percent / 100
    if step <= 0:
        return None
    if price < min_price:
        return 1
    return int((price - min_price) // step) + 1


def price_group_suffix(price, rule, occupied_groups=None):
    group = raw_price_group(price, rule)
    if group is None:
        return ""
    if occupied_groups:
        ordered = sorted(g for g in occupied_groups if g is not None)
        if group in ordered:
            group = ordered.index(group) + 1
    return f".{group:03d}"


def merge_price_ranges(old_rules, new_rules):
    merged = normalize_price_range_rules(old_rules or {})
    for key, rule in normalize_price_range_rules(new_rules or {}).items():
        if key in merged:
            merged[key]["min_price"] = min(merged[key]["min_price"], rule["min_price"])
            merged[key]["max_price"] = max(merged[key]["max_price"], rule["max_price"])
            merged[key]["percent"] = rule.get("percent") or merged[key].get("percent") or 8
            merged[key]["groups"] = rule.get("groups") or merged[key].get("groups") or []
        else:
            merged[key] = rule
    return merged


def product_key(mst, product):
    return f"{mst}|||{product}"


def read_workbook(path):
    with pd.ExcelFile(path) as xl:
        sheet = xl.sheet_names[0]
        df = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=object)
    return sheet, df


def company_rows(df, company_col, mst_col, product_col, qty_col=None, price_col=None, address_col=None, invoice_status_col=DEFAULT_INVOICE_STATUS_COL, invoice_status_skip_values=None):
    ci, mi, pi = map(excel_col_to_index, [company_col, mst_col, product_col])
    qi = excel_col_to_index(qty_col) if str(qty_col or "").strip() else None
    indexes = [ci, mi, pi]
    if qi is not None:
        indexes.append(qi)
    pri = excel_col_to_index(price_col) if price_col else None
    if pri is not None:
        indexes.append(pri)
    ai = excel_col_to_index(address_col) if address_col else None
    if ai is not None:
        indexes.append(ai)
    status_col = str(invoice_status_col or "").strip().upper()
    status_index = excel_col_to_index(status_col) if status_col else None
    if status_index is not None and df.shape[1] > status_index:
        indexes.append(status_index)
    if df.shape[1] <= max(indexes):
        raise ValueError("Selected columns exceed the number of columns in the sheet.")
    rows = []
    for i in range(len(df)):
        if status_index is not None and df.shape[1] > status_index and ignored_invoice_status(cell(df, i, status_index), invoice_status_skip_values):
            continue
        if qi is not None and not should_process_qty(cell(df, i, qi)):
            continue
        mst = "" if pd.isna(cell(df, i, mi)) else str(cell(df, i, mi)).strip()
        product = "" if pd.isna(cell(df, i, pi)) else str(cell(df, i, pi)).strip()
        if not mst or not product:
            continue
        quantity = parse_quantity(cell(df, i, qi)) if qi is not None else 1
        price = parse_price(cell(df, i, pri)) if pri is not None else None
        unit_index = pi + 1
        amount_index = pri + 1 if pri is not None else -1
        amount = parse_price(cell(df, i, amount_index)) if amount_index >= 0 else None
        if (amount is None or amount <= 0) and price is not None:
            amount = price * quantity
        rows.append({
            "excel_row": i + 1,
            "stt": raw_text(cell(df, i, 0)),
            "invoice_no": raw_text(cell(df, i, 2)),
            "invoice_date": raw_text(cell(df, i, 3)),
            "mst": mst,
            "company": "" if pd.isna(cell(df, i, ci)) else str(cell(df, i, ci)).strip(),
            "address": "" if ai is None or pd.isna(cell(df, i, ai)) else str(cell(df, i, ai)).strip(),
            "product": product,
            "unit": raw_text(cell(df, i, unit_index)),
            "quantity": quantity,
            "price": price,
            "amount": amount,
        })
    return rows


def choose_company(names):
    names = [n for n in names if str(n).strip()]
    return Counter(names).most_common(1)[0][0] if names else ""


def unique_values(values):
    seen = set()
    result = []
    for v in values:
        v = "" if pd.isna(v) else str(v).strip()
        if not v or v in seen:
            continue
        seen.add(v)
        result.append(v)
    return result


def analyze(path, company_col, mst_col, address_col, product_col, qty_col, price_col, profile_cfg, invoice_status_col=DEFAULT_INVOICE_STATUS_COL, invoice_status_skip_values=None):
    _, df = read_workbook(path)
    rows = company_rows(df, company_col, mst_col, product_col, qty_col, price_col, address_col, invoice_status_col, invoice_status_skip_values)
    by = {}
    addresses = {}
    products = {}
    for r in rows:
        by.setdefault(r["mst"], []).append(r["company"])
        if r["address"]:
            addresses.setdefault(r["mst"], []).append(r["address"])
        products.setdefault(r["mst"], {}).setdefault(r["product"], []).append(r)

    mst_company = {mst: choose_company(names) for mst, names in by.items()}
    suggestions = {mst: suggest_prefix(comp) for mst, comp in mst_company.items()}
    cnt = Counter(s for s in suggestions.values() if s)
    companies_data = []
    saved_prefixes = profile_cfg.get("prefixes") or {}
    saved_selected = profile_cfg.get("selected_products") or {}

    for mst in mst_company:
        suggested = saved_prefixes.get(mst) or suggestions[mst]
        product_items = []
        for name, product_rows in products.get(mst, {}).items():
            prices = [r["price"] for r in product_rows if r["price"] is not None]
            price_rows = [{
                "price": r["price"],
                "quantity": r["quantity"],
                "amount": r["amount"],
                "excelRow": r["excel_row"],
                "stt": r["stt"],
                "invoiceNo": r["invoice_no"],
                "invoiceDate": r["invoice_date"],
                "unit": r["unit"],
                "name": r["product"],
            } for r in product_rows if r["price"] is not None]
            product_items.append({
                "name": name,
                "count": len(product_rows),
                "minPrice": min(prices) if prices else None,
                "maxPrice": max(prices) if prices else None,
                "priceCount": len(set(prices)),
                "priceRows": price_rows,
            })
        skipped_list = saved_selected.get(mst)
        skipped_set = set(skipped_list) if isinstance(skipped_list, list) else set()
        companies_data.append({
            "mst": str(mst),
            "company": str(mst_company[mst]),
            "address": unique_values(addresses.get(mst, []))[0] if addresses.get(mst) else "",
            "addresses": unique_values(addresses.get(mst, [])),
            "all_names": [str(n) for n in unique_values(by[mst])],
            "all_products": product_items,
            "selected_product_names": [p["name"] for p in product_items if p["name"] not in skipped_set],
            "count": int(len(by[mst])),
            "default_prefix": str(suggestions[mst]),
            "suggested": str(suggested),
            "value": str(suggested),
            "needs_manual": (not suggested) or cnt[suggestions[mst]] > 1,
            "status": "OK" if suggested and cnt[suggestions[mst]] <= 1 else "Need manual check",
        })

    companies_data.sort(key=lambda x: (x["mst"], x["company"]))
    for idx, item in enumerate(companies_data):
        item["safe_id"] = str(idx)
    return {"rows_to_process": int(len(rows)), "company_count": int(len(companies_data)), "companies": companies_data}


def resolve_output_path(original, requested_path):
    default_name = f"{Path(original).stem}_formatted.xlsx"
    requested_path = str(requested_path or "").strip().strip('"')
    if not requested_path:
        return OUTPUT_DIR / default_name
    p = Path(requested_path)
    if p.suffix.lower() in {".xlsx", ".xlsm"}:
        p.parent.mkdir(parents=True, exist_ok=True)
        return p
    p.mkdir(parents=True, exist_ok=True)
    return p / default_name


def validate_payload(data):
    all_mst = [str(x).strip() for x in data.get("all_mst", [])]
    selected = set(str(x).strip() for x in data.get("process_mst", []))
    include_company_prefix = data.get("include_company_prefix") is not False
    safe = {}
    for item in data.get("mst_safe_id", []):
        if "|||" in item:
            mst, sid = item.split("|||", 1)
            safe[mst] = sid
    prefix_map = {}
    selected_products = {}
    used = {}
    errors = []
    for mst in all_mst:
        if not mst or mst not in selected:
            continue
        sid = safe.get(mst, "")
        if include_company_prefix:
            prefix = str(data.get(f"prefix_{sid}", "") or "").strip().upper()
            if not prefix:
                errors.append(f"{mst}: prefix is empty.")
                continue
            if not re.fullmatch(r"[A-Z0-9]{1,20}", prefix):
                errors.append(f"{mst}: prefix must use only A-Z or 0-9.")
                continue
            if prefix in used and used[prefix] != mst:
                errors.append(f"{mst}: prefix {prefix} duplicates MST {used[prefix]}.")
                continue
            used[prefix] = mst
            prefix_map[mst] = prefix
        selected_products[mst] = set(data.get(f"selected_products_{sid}", []))
    if errors:
        raise ValueError("\n".join(errors))
    return prefix_map, selected_products


def process_workbook(path, out, data):
    sheet, df = read_workbook(path)
    company_col = data.get("company_col", "F").upper()
    mst_col = data.get("mst_col", "G").upper()
    product_col = data.get("product_col", "N").upper()
    qty_col = str(data.get("qty_col", "P") or "").upper()
    price_col = str(data.get("price_col", "R") or "").upper()
    output_col = data.get("output_col", "M").upper()
    invoice_status_col = str(data.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL) or "").upper()
    invoice_status_skip_values = data.get("invoice_status_skip_values")
    ci, mi, pi, oi = map(excel_col_to_index, [company_col, mst_col, product_col, output_col])
    pri = excel_col_to_index(price_col) if price_col else None
    qi = excel_col_to_index(qty_col) if qty_col else None
    indexes = [ci, mi, pi, oi]
    if pri is not None:
        indexes.append(pri)
    if qi is not None:
        indexes.append(qi)
    if df.shape[1] <= max(indexes):
        raise ValueError("Selected columns exceed the number of columns in the sheet.")

    profile = data.get("profile", "son_phuong")
    word_rules = data.get("word_rules") or {}
    first_word_rules = data.get("first_word_rules") or {}
    repeated_phrase_removals = normalize_phrase_list(data.get("repeated_phrase_removals") or [])
    include_company_prefix = data.get("include_company_prefix") is not False
    price_rules = data.get("price_group_rules") or {}
    price_range_rules = data.get("price_range_rules") or {}
    manual_code_overrides = data.get("manual_code_overrides") or {}
    prefix_map, selected_products = validate_payload(data)
    rows = company_rows(df, company_col, mst_col, product_col, qty_col, price_col, invoice_status_col=invoice_status_col, invoice_status_skip_values=invoice_status_skip_values)

    occupied = {}
    for r in rows:
        mst = r["mst"]
        prod = r["product"]
        if (include_company_prefix and mst not in prefix_map) or prod not in selected_products.get(mst, set()):
            continue
        key = product_key(mst, prod)
        base_code = str(manual_code_overrides.get(key) or "").strip()
        if not base_code:
            base_code = make_code(mst, prod, 1, prefix_map, profile, word_rules, first_word_rules, require_qty=False, include_company_prefix=include_company_prefix, repeated_phrase_removals=repeated_phrase_removals)
        rule = price_rules.get(key) or price_range_rules.get(base_code)
        if not rule:
            continue
        group = raw_price_group(r["price"], rule)
        if group is not None:
            occupied.setdefault(key, set()).add(group)

    processed_row_indexes = set()
    for row in rows:
        i = row["excel_row"] - 1
        mst = row["mst"]
        prod = row["product"]
        if (include_company_prefix and mst not in prefix_map) or prod not in selected_products.get(mst, set()):
            continue
        key = product_key(mst, prod)
        qty = cell(df, i, qi) if qi is not None else 1
        code = str(manual_code_overrides.get(key) or "").strip()
        if not code:
            code = make_code(mst, prod, qty, prefix_map, profile, word_rules, first_word_rules, require_qty=(qi is not None), include_company_prefix=include_company_prefix, repeated_phrase_removals=repeated_phrase_removals)
        rule = price_rules.get(key) or price_range_rules.get(code)
        if code and rule and pri is not None:
            code += price_group_suffix(parse_price(cell(df, i, pri)), rule, occupied.get(key))
        if code:
            df.iat[i, oi] = code
            processed_row_indexes.add(i)

    header_index = None
    for i in range(len(df)):
        output_header = rm_accents(raw_text(cell(df, i, oi))).upper()
        product_header = rm_accents(raw_text(cell(df, i, pi))).upper()
        qty_header = rm_accents(raw_text(cell(df, i, qi))).upper() if qi is not None else ""
        if "MA VT" in output_header or ("TEN" in product_header and "HANG" in product_header) or "SO LUONG" in qty_header:
            header_index = i
            break
    if header_index is None:
        header_index = min(processed_row_indexes, default=len(df)) - 1

    keep_indexes = [
        i for i in range(len(df))
        if i <= header_index or i in processed_row_indexes
    ]
    df = df.iloc[keep_indexes].reset_index(drop=True)

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet, index=False, header=False)
    return df


def up_ban_ra_tax_values(value):
    rate = parse_price(value)
    if rate is None:
        return "", ""
    percent = rate * 100 if abs(rate) <= 1 else rate
    rounded = int(round(percent))
    if abs(percent - rounded) < 0.000001:
        return f"{rounded:02d}", rounded
    return f"{percent:02g}", percent


def up_ban_ra_template_sheet(workbook):
    for ws in workbook.worksheets:
        if rm_accents(ws.title).casefold() == "up ban ra":
            return ws
    return workbook.worksheets[0]


def create_up_ban_ra_workbook(processed_df):
    from openpyxl import load_workbook

    template_path = RESOURCE_DIR / "mau HD ban ra.xlsx"
    if not template_path.exists():
        raise ValueError("Không tìm thấy file mẫu mau HD ban ra.xlsx để tạo file UP bán ra.")

    wb = load_workbook(template_path)
    ws = up_ban_ra_template_sheet(wb)
    ws.title = "UP Bán ra"
    for other in list(wb.worksheets):
        if other is not ws:
            wb.remove(other)

    max_column = ws.max_column
    data_styles = [copy(ws.cell(2, col)._style) for col in range(1, max_column + 1)]
    data_number_formats = [ws.cell(2, col).number_format for col in range(1, max_column + 1)]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    source_indexes = {name: excel_col_to_index(name) for name in ["C", "D", "J", "L", "O", "P", "Q", "R", "W"]}
    output_row = 2
    for row in range(len(processed_df)):
        code = raw_text(cell(processed_df, row, source_indexes["L"]))
        invoice_no = raw_text(cell(processed_df, row, source_indexes["C"]))
        quantity = parse_price(cell(processed_df, row, source_indexes["O"]))
        if not code or not invoice_no or quantity is None or rm_accents(code).casefold() == "ma vt":
            continue

        tax_code, tax_rate = up_ban_ra_tax_values(cell(processed_df, row, source_indexes["R"]))
        values = {
            "A": cell(processed_df, row, source_indexes["J"]),
            "E": cell(processed_df, row, source_indexes["C"]),
            "F": cell(processed_df, row, source_indexes["D"]),
            "H": cell(processed_df, row, source_indexes["O"]),
            "O": cell(processed_df, row, source_indexes["P"]),
            "P": cell(processed_df, row, source_indexes["Q"]),
            "V": tax_code,
            "W": tax_rate,
            "Y": cell(processed_df, row, source_indexes["W"]),
            "Z": "1311",
            "AA": "",
            "AB": "",
            "AC": "632",
            "AE": "33311",
            "AG": cell(processed_df, row, source_indexes["L"]),
            "AH": "Xuất bán hàng",
            "AK": 1,
            "AQ": "CTY",
            "AS": 1,
        }
        for col in range(1, max_column + 1):
            output_cell = ws.cell(output_row, col)
            output_cell._style = copy(data_styles[col - 1])
            output_cell.number_format = data_number_formats[col - 1]
        for column, value in values.items():
            ws[f"{column}{output_row}"] = "" if pd.isna(value) else value
        output_row += 1

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


def up_ban_ra_output_path(formatted_path):
    return formatted_path.with_name(f"{formatted_path.stem}_UP_ban_ra.xlsx")


def process_zip_stream(formatted_path, up_path):
    stream = BytesIO()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(formatted_path, formatted_path.name)
        archive.write(up_path, up_path.name)
    stream.seek(0)
    return stream


def preview_data(df):
    preview_indexes = list(range(df.shape[1]))
    view = df.iloc[:8, preview_indexes].copy()
    view.columns = [index_to_excel_col(i) for i in preview_indexes]
    view = view.fillna("")
    return [{col: str(val) if val != "" else "" for col, val in row.items()} for _, row in view.iterrows()]


def invoice_status_options(df, invoice_status_col=DEFAULT_INVOICE_STATUS_COL, skip_statuses=None):
    status_col = str(invoice_status_col or "").strip().upper()
    if not status_col:
        return []
    status_index = excel_col_to_index(status_col)
    if df.shape[1] <= status_index:
        return []
    counts = Counter()
    display_values = {}
    for i in range(len(df)):
        value = raw_text(cell(df, i, status_index))
        if not value:
            continue
        normalized = normalized_invoice_status(value)
        if normalized == "trang thai hoa don":
            continue
        counts[normalized] += 1
        display_values.setdefault(normalized, value)
    selected = normalized_invoice_status_set(skip_statuses)
    return [
        {
            "value": display_values[key],
            "count": count,
            "skip": key in selected,
        }
        for key, count in counts.most_common()
    ]


def safe_excel_sheet_name(value, used_names):
    name = re.sub(r"[:\\/?*\[\]]", " ", str(value or "Sheet")).strip() or "Sheet"
    name = name[:31]
    base = name
    counter = 2
    while name in used_names:
        suffix = f" {counter}"
        name = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1
    used_names.add(name)
    return name


def safe_excel_filename(value):
    stem = Path(str(value or "bao_cao_ban_hang.xlsx")).stem or "bao_cao_ban_hang"
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", stem).strip(" .") or "bao_cao_ban_hang"
    return f"{stem}.xlsx"


def coerce_excel_value(value):
    if value is None:
        return ""
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value) if isinstance(value, (float, np.floating)) else int(value)
    return str(value)


def make_excel_workbook(sheets):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()
    header_fill = PatternFill("solid", fgColor="EAF4FF")
    total_fill = PatternFill("solid", fgColor="F7FBFF")
    bold_font = Font(bold=True)

    for sheet in sheets:
        rows = sheet.get("rows") if isinstance(sheet, dict) else []
        rows = rows if isinstance(rows, list) else []
        ws = wb.create_sheet(safe_excel_sheet_name(sheet.get("name") if isinstance(sheet, dict) else "Sheet", used_names))
        headers = sheet.get("headers") if isinstance(sheet, dict) and isinstance(sheet.get("headers"), list) else []
        headers = [str(header) for header in headers if str(header).strip()]
        if not headers:
            headers = list(rows[0].keys()) if rows and isinstance(rows[0], dict) else []
        if not headers:
            ws.append(["Không có dữ liệu"])
            continue

        ws.append(headers)
        for cell in ws[1]:
            cell.font = bold_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in rows:
            values = [coerce_excel_value(row.get(header)) if isinstance(row, dict) else "" for header in headers]
            ws.append(values)
            if str(row.get(headers[0], "") if isinstance(row, dict) else "").upper() == "TỔNG CỘNG":
                for cell in ws[ws.max_row]:
                    cell.font = bold_font
                    cell.fill = total_fill

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_index, header in enumerate(headers, start=1):
            letter = get_column_letter(column_index)
            width = min(max(len(str(header)) + 4, 12), 36)
            for cell in ws[letter][1: min(ws.max_row, 50)]:
                width = min(max(width, len(str(cell.value or "")) + 2), 36)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
            ws.column_dimensions[letter].width = width

    if not wb.worksheets:
        wb.create_sheet("Bao cao")
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream


@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(load_config())


@app.route("/api/config", methods=["POST"])
def set_config():
    return jsonify(save_config(request.get_json() or {}))


@app.route("/api/config/profile/<profile>", methods=["POST"])
def import_profile_config(profile):
    target = profile_key(profile)
    incoming = normalize_config(request.get_json() or {})
    source = target if target in incoming.get("profiles", {}) else incoming.get("selected_profile")
    source = source if source in PROFILE_LABELS else target
    current = load_config()
    current["selected_profile"] = target
    current["columns"].update(incoming.get("columns") or {})
    current["profiles"][target] = incoming["profiles"].get(source, empty_profile_config(target))
    return jsonify(save_config(current))


@app.route("/api/mapping", methods=["POST"])
def mapping():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded."}), 400
    original = f.filename
    ext = Path(original).suffix.lower()
    if ext not in [".xlsx", ".xlsm"]:
        return jsonify({"error": "Please upload .xlsx or .xlsm only."}), 400
    saved = f"{uuid.uuid4().hex}{ext}"
    path = UPLOAD_DIR / saved
    f.save(path)
    try:
        _, df = read_workbook(path)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    cols = []
    for idx in range(df.shape[1]):
        letter = index_to_excel_col(idx)
        samples = []
        for r in range(min(6, len(df))):
            v = df.iat[r, idx]
            if not pd.isna(v) and str(v).strip():
                samples.append(str(v).strip())
        label = letter + ((" - " + " | ".join(samples[:2])[:45]) if samples else "")
        cols.append({"letter": letter, "label": label})
    return jsonify({
        "original_name": original,
        "saved_name": saved,
        "columns": cols,
        "preview": preview_data(df),
        "invoice_statuses": invoice_status_options(df, DEFAULT_INVOICE_STATUS_COL, DEFAULT_INVOICE_STATUS_SKIP_VALUES),
    })


@app.route("/api/invoice_statuses", methods=["POST"])
def invoice_statuses():
    data = request.get_json() or {}
    saved = data.get("saved_name", "")
    path = UPLOAD_DIR / saved
    if not path.exists():
        return jsonify({"error": "Uploaded file was not found. Please upload again."}), 400
    try:
        _, df = read_workbook(path)
        statuses = invoice_status_options(
            df,
            str(data.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL) or "").upper(),
            data.get("invoice_status_skip_values"),
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"invoice_statuses": statuses})


@app.route("/api/check", methods=["POST"])
def check():
    data = request.get_json() or {}
    saved = data.get("saved_name", "")
    profile = data.get("profile", "son_phuong")
    cfg = load_config()["profiles"].get(profile, empty_profile_config())
    path = UPLOAD_DIR / saved
    if not path.exists():
        return jsonify({"error": "Uploaded file was not found. Please upload again."}), 400
    try:
        result = analyze(
            path,
            data.get("company_col", "F").upper(),
            data.get("mst_col", "G").upper(),
            data.get("address_col", "H").upper(),
            data.get("product_col", "N").upper(),
            data.get("qty_col", "P").upper(),
            data.get("price_col", "R").upper(),
            cfg,
            str(data.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL) or "").upper(),
            data.get("invoice_status_skip_values"),
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    result.update({
        "original_name": data.get("original_name", "output.xlsx"),
        "saved_name": saved,
    })
    return jsonify(result)


@app.route("/api/export_price_report", methods=["POST"])
def export_price_report():
    data = request.get_json() or {}
    sheets = data.get("sheets") or []
    if not isinstance(sheets, list) or not sheets:
        return jsonify({"error": "Không có dữ liệu để xuất Excel."}), 400
    try:
        workbook = make_excel_workbook(sheets)
        filename = safe_excel_filename(data.get("filename") or "bao_cao_ban_hang.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return send_file(
        workbook,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/process", methods=["POST"])
def process():
    data = request.get_json() or {}
    saved = data.get("saved_name", "")
    original = data.get("original_name", "output.xlsx")
    path = UPLOAD_DIR / saved
    if not path.exists():
        return jsonify({"error": "Uploaded file was not found. Please upload again."}), 400
    try:
        out = resolve_output_path(original, data.get("output_path", ""))
        processed_df = process_workbook(path, out, data)
        up_stream = create_up_ban_ra_workbook(processed_df)
        up_out = up_ban_ra_output_path(out)
        up_out.write_bytes(up_stream.getvalue())
        cfg = load_config()
        profile = data.get("profile", cfg.get("selected_profile", "son_phuong"))
        cfg["selected_profile"] = profile
        cfg["columns"].update({k: data.get(k, v) for k, v in cfg["columns"].items()})
        cfg["profiles"].setdefault(profile, empty_profile_config(profile))
        old_profile = cfg["profiles"].get(profile) or empty_profile_config(profile)
        merged_price_ranges = merge_price_ranges(old_profile.get("price_range_rules"), data.get("price_range_rules", {}))
        cfg["profiles"][profile].update({
            "prefixes": data.get("prefixes", {}),
            "selected_products": data.get("skipped_products_map", {}),
            "removed_companies": data.get("removed_companies", old_profile.get("removed_companies", {})),
            "word_rules": data.get("word_rules", {}),
            "first_word_rules": data.get("first_word_rules", {}),
            "repeated_phrase_removals": normalize_phrase_list(data.get("repeated_phrase_removals", old_profile.get("repeated_phrase_removals", []))),
            "price_group_rules": data.get("price_group_rules", {}),
            "price_range_rules": merged_price_ranges,
            "price_adjust_all_percent": float(data.get("price_adjust_all_percent") or 0),
            "manual_code_overrides": data.get("manual_code_overrides", {}),
            "include_company_prefix": data.get("include_company_prefix") is not False,
            "output_path": data.get("output_path", ""),
        })
        save_config(cfg)
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    zip_name = f"{Path(original).stem}_ket_qua_xu_ly.zip"
    return send_file(
        process_zip_stream(out, up_out),
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_frontend(path):
    if STATIC_DIR.exists():
        file_path = STATIC_DIR / path
        if path and file_path.exists() and file_path.is_file():
            return send_from_directory(str(STATIC_DIR), path)
        return send_from_directory(str(STATIC_DIR), "index.html")
    return jsonify({"status": "ok"})


def open_browser():
    for _ in range(60):
        try:
            with urlopen("http://127.0.0.1:5000/api/config", timeout=0.5):
                break
        except Exception:
            time.sleep(0.5)
    webbrowser.open("http://localhost:5000")


def run_tray_icon():
    try:
        import pystray
        from PIL import Image
    except Exception:
        return
    if not ICON_PATH.exists():
        return

    def open_app(icon, item):
        open_browser()

    def quit_app(icon, item):
        icon.stop()
        os._exit(0)

    image = Image.open(ICON_PATH)
    menu = pystray.Menu(
        pystray.MenuItem("Mở Product Code Formatter", open_app),
        pystray.MenuItem("Thoát", quit_app),
    )
    pystray.Icon("ProductCodeFormatter", image, "Product Code Formatter", menu).run()


def close_existing_instances():
    current_pid = str(os.getpid())
    try:
        output = subprocess.check_output(
            ["netstat", "-ano", "-p", "tcp"],
            text=True,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        return
    for line in output.splitlines():
        parts = line.split()
        if len(parts) < 5 or parts[0].lower() != "tcp":
            continue
        local_address = parts[1]
        state = parts[3].lower()
        pid = parts[4]
        if not local_address.endswith(":5000") or state != "listening":
            continue
        if pid and pid != current_pid:
            subprocess.run(["taskkill", "/F", "/PID", pid, "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def main():
    if getattr(sys, "frozen", False):
        close_existing_instances()
        threading.Thread(target=run_tray_icon, daemon=True).start()
        threading.Timer(1.0, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()

