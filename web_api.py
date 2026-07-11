import atexit
import json
import mimetypes
import os
import re
import shutil
import socket
import threading
import time
import uuid
import webbrowser
import zipfile
from io import BytesIO
from pathlib import Path

import pandas as pd
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app import (
    APP_DATA_DIR,
    DEFAULT_INVOICE_STATUS_COL,
    DEFAULT_INVOICE_STATUS_SKIP_VALUES,
    ICON_PATH,
    OUTPUT_DIR,
    TEMPLATE_DIR,
    UPLOAD_DIR,
    XLS_MEDIA_TYPE,
    VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES,
    VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
    VIETMAX_PHASE_PURCHASE,
    VIETMAX_PHASE_SALES,
    MAX_CODE_LENGTH,
    VIETMAX_PROFILE,
    analyze,
    activate_keygen_license,
    build_vietmax_ban_ra_purchase_matches,
    build_vietmax_khh_exact_purchase_matches,
    load_system_default_form_mappings,
    apply_word_rules_to_words,
    cell,
    code_words,
    effective_processing_profile,
    empty_profile_config,
    excel_col_to_index,
    index_to_excel_col,
    invoice_status_options,
    license_allows_profile,
    license_has_local_activation,
    load_config,
    make_excel_workbook,
    make_product_part,
    merge_price_ranges,
    mixed_xls_workbook,
    apply_product_code_replacement,
    normalize_config,
    normalize_profile_config,
    normalize_product_code_replacements,
    normalize_vietmax_comparison_scope,
    normalize_inventory_pair_rules,
    normalize_inventory_pairs,
    normalize_missing_mst_prefix_strategy,
    normalize_phrase_list,
    normalize_vietmax_phase,
    normalized_header_label,
    openpyxl_workbook_to_xls_stream,
    preview_data,
    process_workbook,
    product_key,
    profile_key,
    raw_text,
    read_workbook,
    remove_repeated_phrases,
    resolve_output_path,
    sanitize_product_code,
    save_config,
    should_process_qty,
    fast_import_sheet_rows,
    fast_merge_mapping_forms,
    fast_import_multi_sheet_workbook,
    validate_vietmax_processed_purchase_workbook,
    validate_fast_import_processed_dataframe,
    vietmax_default_source_columns,
    vietmax_ban_ra_match_key,
    vietmax_ban_ra_sales_products_from_workbook,
    vietmax_company_identity_key,
    vietmax_product_review_rows,
    vietmax_purchase_match_export_rows,
    vietmax_purchase_products_from_workbook,
)

from inventory_allocation_app.app import (
    OUTPUT_DIR as INVENTORY_OUTPUT_DIR,
    get_analysis_job as get_inventory_analysis_job,
    run_analysis_job as run_inventory_analysis_job,
    update_analysis_job as update_inventory_analysis_job,
)
from product_code.excel_io import (
    diagnostic_log,
    uploaded_workbook_content_for_openpyxl,
    workbook_content_for_openpyxl,
)
from product_code.license_client import public_license_status as build_public_license_status
from product_code.workflow_runtime import (
    WorkflowFailure,
    WorkflowJobManager,
    WorkflowSessionStore,
    file_sha256,
    stable_signature,
)
from workflows.estimate_extractor.logic import analyze_estimate_workbook, create_estimate_output_workbook, list_estimate_workbook_sheets

APP_DIR = Path(__file__).resolve().parent
REACT_DIST_DIR = APP_DIR / "react_frontend" / "dist"
PROGRESS_LOCK = threading.Lock()
PROGRESS_JOBS: dict[str, dict] = {}
PROGRESS_TTL_SECONDS = 900
WORKFLOW_SESSION_STORE = WorkflowSessionStore(APP_DATA_DIR / "sessions", UPLOAD_DIR)
WORKFLOW_JOB_MANAGER = WorkflowJobManager(max_workers=2)


def close_workflow_runtime():
    WORKFLOW_JOB_MANAGER.shutdown()
    WORKFLOW_SESSION_STORE.close()


atexit.register(close_workflow_runtime)


PREFIX_STRATEGIES = ("last_2_words", "last_3_mst", "2_words_mst", "all_name_words")


def has_config_value(value) -> bool:
    if isinstance(value, dict):
        return bool(value)
    if isinstance(value, list):
        return bool(value)
    if isinstance(value, str):
        return bool(value.strip())
    return value is not None


def keep_existing_when_empty(payload: dict, profile_cfg: dict, key: str, default=None):
    if key not in payload:
        return profile_cfg.get(key, default)
    value = payload.get(key)
    if has_config_value(value) or not has_config_value(profile_cfg.get(key)):
        return value
    return profile_cfg.get(key, default)


def keep_form_mapping_presets(payload: dict, profile_cfg: dict):
    if "form_mapping_presets" not in payload:
        return profile_cfg.get("form_mapping_presets", [])
    value = payload.get("form_mapping_presets")
    if not isinstance(value, list):
        return profile_cfg.get("form_mapping_presets", [])
    if payload.get("replace_form_mapping_presets") is True:
        return value

    existing = profile_cfg.get("form_mapping_presets")
    existing = existing if isinstance(existing, list) else []
    incoming_ids = {
        str(item.get("id") or "").strip()
        for item in value
        if isinstance(item, dict) and str(item.get("id") or "").strip()
    }
    preserved = [
        item for item in existing
        if isinstance(item, dict) and str(item.get("id") or "").strip() not in incoming_ids
    ]
    return value + preserved


def has_company_selection_payload(payload: dict) -> bool:
    if has_config_value(payload.get("all_mst")) or has_config_value(payload.get("mst_safe_id")):
        return True
    return any(str(key).startswith("selected_products_") for key in payload)


def keep_company_config_when_unloaded(payload: dict, profile_cfg: dict, payload_key: str, profile_key_name: str, default=None):
    if has_company_selection_payload(payload):
        return payload.get(payload_key, default)
    if has_config_value(payload.get(payload_key)):
        return payload.get(payload_key)
    return profile_cfg.get(profile_key_name, default)


def has_meaningful_profile_config(profile_cfg: dict) -> bool:
    if not isinstance(profile_cfg, dict):
        return False
    collection_keys = (
        "prefixes",
        "selected_products",
        "removed_companies",
        "first_word_rules",
        "repeated_phrase_removals",
        "price_group_rules",
        "price_range_rules",
        "manual_code_overrides",
        "product_code_replacements",
        "product_review_merges",
        "vietmax_mua_vao_internal_merges",
        "vietmax_ban_ra_sales_internal_merges",
        "vietmax_ban_ra_purchase_match_rules",
        "inventory_pairs",
        "inventory_pair_rules",
        "inventory_allocation_config",
        "prefix_strategy_values",
        "processing_groups",
        "company_group_assignments",
        "form_mapping_presets",
        "columns",
    )
    if any(has_config_value(profile_cfg.get(key)) for key in collection_keys):
        return True
    if profile_cfg.get("include_company_prefix") is False:
        return True
    if profile_cfg.get("use_default_inventory_pair") is True:
        return True
    if raw_text(profile_cfg.get("default_inventory_pair_id")):
        return True
    if raw_text(profile_cfg.get("prefix_strategy") or "last_2_words") != "last_2_words":
        return True
    for key, default in (
        ("prefix_mst_digits", 3),
        ("prefix_name_words", 2),
        ("prefix_name_chars", 1),
    ):
        try:
            if int(profile_cfg.get(key, default)) != default:
                return True
        except Exception:
            return True
    if raw_text(profile_cfg.get("prefix_missing_mst_strategy") or "all_name_words") != "all_name_words":
        return True
    return False


def is_default_word_rules(value) -> bool:
    if not isinstance(value, dict) or len(value) != 2:
        return False
    normalized = {raw_text(key).casefold(): raw_text(val).upper() for key, val in value.items()}
    return normalized.get("đen") == "DEN" and normalized.get("tôn") == "TON"


def should_use_merged_profile_value(key: str, value, existing) -> bool:
    if isinstance(value, list):
        return bool(value) or not has_config_value(existing)
    if isinstance(value, dict):
        if key == "word_rules" and is_default_word_rules(value) and has_config_value(existing) and not is_default_word_rules(existing):
            return False
        return bool(value) or not has_config_value(existing)
    if isinstance(value, str):
        return bool(value.strip()) or not has_config_value(existing)
    return value is not None


def merge_profile_configs(*configs):
    result = {}
    for config in configs:
        if not isinstance(config, dict):
            continue
        for key, value in config.items():
            if key == "scopes":
                continue
            if should_use_merged_profile_value(key, value, result.get(key)):
                result[key] = value
    return result


def scoped_profile_config_from_profiles(profiles: dict, profile: str, phase: str):
    profile = profile_key(profile)
    phase = normalize_vietmax_phase(phase)
    profile_cfg = profiles.get(profile) if isinstance(profiles, dict) else {}
    profile_cfg = profile_cfg if isinstance(profile_cfg, dict) else {}
    scopes = profile_cfg.get("scopes") if isinstance(profile_cfg.get("scopes"), dict) else {}
    scoped = scopes.get(phase)
    scoped_cfg = scoped if has_meaningful_profile_config(scoped) else None
    if profile == VIETMAX_PROFILE:
        legacy_key = effective_processing_profile(profile, phase)
        legacy_cfg = profiles.get(legacy_key) if isinstance(profiles, dict) else {}
        legacy_cfg = legacy_cfg if has_meaningful_profile_config(legacy_cfg) else None
        return merge_profile_configs(profile_cfg, legacy_cfg, scoped_cfg) or empty_profile_config(profile)
    if scoped_cfg:
        return merge_profile_configs(profile_cfg, scoped_cfg)
    return profile_cfg or empty_profile_config(profile)


def scoped_profile_config(cfg: dict, profile: str, phase: str):
    profiles = cfg.get("profiles", {}) if isinstance(cfg.get("profiles"), dict) else {}
    return scoped_profile_config_from_profiles(profiles, profile, phase)


def ensure_profile_scope(cfg: dict, profile: str, phase: str):
    profile = profile_key(profile)
    phase = normalize_vietmax_phase(phase)
    cfg.setdefault("profiles", {}).setdefault(profile, empty_profile_config(profile))
    root = cfg["profiles"].get(profile) or empty_profile_config(profile)
    if not isinstance(root, dict):
        root = empty_profile_config(profile)
    root.setdefault("scopes", {})
    root["scopes"].setdefault(phase, empty_profile_config(profile, include_scopes=False))
    cfg["profiles"][profile] = root
    return root, root["scopes"][phase]


def sync_vietmax_legacy_profile(cfg: dict, profile: str, phase: str, profile_cfg: dict):
    if profile_key(profile) != VIETMAX_PROFILE:
        return
    legacy_key = effective_processing_profile(VIETMAX_PROFILE, phase)
    legacy_cfg = cfg.setdefault("profiles", {}).get(legacy_key)
    if not isinstance(legacy_cfg, dict):
        legacy_cfg = empty_profile_config(legacy_key)
    legacy_cfg.update(profile_cfg)
    cfg["profiles"][legacy_key] = legacy_cfg


def normalize_prefix_strategy_values(raw):
    values = {strategy: {} for strategy in PREFIX_STRATEGIES}
    if not isinstance(raw, dict):
        return values
    for strategy in PREFIX_STRATEGIES:
        strategy_values = raw.get(strategy)
        if not isinstance(strategy_values, dict):
            continue
        values[strategy] = {
            str(key): str(value).strip().upper()
            for key, value in strategy_values.items()
            if key and value is not None
        }
    return values


VIETMAX_CONFIG_PROFILE_KEYS = {
    "prefixes",
    "selected_products",
    "removed_companies",
    "word_rules",
    "first_word_rules",
    "repeated_phrase_removals",
    "manual_code_overrides",
    "product_code_replacements",
    "product_review_merges",
    "vietmax_mua_vao_internal_merges",
    "vietmax_ban_ra_sales_internal_merges",
    "vietmax_ban_ra_purchase_match_rules",
    "inventory_pairs",
    "use_default_inventory_pair",
    "default_inventory_pair_id",
    "inventory_pair_rules",
    "inventory_allocation_config",
    "include_company_prefix",
    "prefix_strategy",
    "prefix_mst_digits",
    "prefix_name_words",
    "prefix_name_chars",
    "prefix_missing_mst_strategy",
    "prefix_strategy_values",
    "processing_groups",
    "company_group_assignments",
    "form_mapping_presets",
    "columns",
}


def vietmax_config_storage_profile(phase: str) -> tuple[str, str]:
    normalized_phase = VIETMAX_PHASE_SALES if raw_text(phase) == VIETMAX_PHASE_SALES else VIETMAX_PHASE_PURCHASE
    return normalized_phase, effective_processing_profile(VIETMAX_PROFILE, normalized_phase)


def dict_value(value):
    return value if isinstance(value, dict) else {}


def looks_like_vietmax_profile_config(value) -> bool:
    return isinstance(value, dict) and any(key in value for key in VIETMAX_CONFIG_PROFILE_KEYS)


def vietmax_profile_from_process_payload(payload: dict, phase: str) -> dict:
    if not isinstance(payload, dict):
        return {}
    profile_cfg = {}
    direct_keys = [
        "prefixes",
        "removed_companies",
        "include_company_prefix",
        "prefix_strategy",
        "prefix_mst_digits",
        "prefix_name_words",
        "prefix_name_chars",
        "prefix_missing_mst_strategy",
        "prefix_strategy_values",
        "processing_groups",
        "company_group_assignments",
        "form_mapping_presets",
        "columns",
        "word_rules",
        "first_word_rules",
        "repeated_phrase_removals",
        "manual_code_overrides",
        "product_code_replacements",
        "product_review_merges",
        "inventory_pairs",
        "use_default_inventory_pair",
        "default_inventory_pair_id",
        "inventory_pair_rules",
        "inventory_allocation_config",
    ]
    for key in direct_keys:
        if key in payload:
            profile_cfg[key] = payload.get(key)
    if "skipped_products_map" in payload:
        profile_cfg["selected_products"] = payload.get("skipped_products_map") or {}
    elif "selected_products" in payload:
        profile_cfg["selected_products"] = payload.get("selected_products") or {}
    if phase == VIETMAX_PHASE_SALES:
        if "vietmax_ban_ra_sales_internal_merges" in payload:
            profile_cfg["vietmax_ban_ra_sales_internal_merges"] = payload.get("vietmax_ban_ra_sales_internal_merges") or []
        if "vietmax_ban_ra_purchase_match_rules" in payload:
            profile_cfg["vietmax_ban_ra_purchase_match_rules"] = payload.get("vietmax_ban_ra_purchase_match_rules") or []
    else:
        if "vietmax_mua_vao_internal_merges" in payload:
            profile_cfg["vietmax_mua_vao_internal_merges"] = payload.get("vietmax_mua_vao_internal_merges") or []
    if "prefix_strategy_values" in profile_cfg:
        profile_cfg["prefix_strategy_values"] = normalize_prefix_strategy_values(profile_cfg.get("prefix_strategy_values") or {})
    if "inventory_pairs" in profile_cfg:
        profile_cfg["inventory_pairs"] = normalize_inventory_pairs(profile_cfg.get("inventory_pairs") or [])
    if "inventory_pair_rules" in profile_cfg:
        profile_cfg["inventory_pair_rules"] = normalize_inventory_pair_rules(profile_cfg.get("inventory_pair_rules") or [])
    if "repeated_phrase_removals" in profile_cfg:
        profile_cfg["repeated_phrase_removals"] = normalize_phrase_list(profile_cfg.get("repeated_phrase_removals") or [])
    if "product_code_replacements" in profile_cfg:
        profile_cfg["product_code_replacements"] = normalize_product_code_replacements(profile_cfg.get("product_code_replacements") or {})
    return profile_cfg


def extract_vietmax_import_profile(payload: dict, phase: str) -> tuple[str, dict]:
    if not isinstance(payload, dict):
        raise HTTPException(status_code=400, detail="File cấu hình không đúng định dạng JSON object.")
    normalized_phase, storage_profile = vietmax_config_storage_profile(phase)
    incoming_phase = raw_text(payload.get("phase"))
    if incoming_phase in {VIETMAX_PHASE_PURCHASE, VIETMAX_PHASE_SALES} and incoming_phase != normalized_phase:
        raise HTTPException(status_code=400, detail=f"File cấu hình là {incoming_phase}, không khớp với phase đang nhập {normalized_phase}.")
    incoming_storage_profile = raw_text(payload.get("storage_profile"))
    if incoming_storage_profile in {"vietmax_mua_vao", "vietmax_ban_ra"} and incoming_storage_profile != storage_profile:
        raise HTTPException(status_code=400, detail=f"File cấu hình thuộc {incoming_storage_profile}, không khớp với {storage_profile}.")

    snapshot = dict_value(payload.get("saved_config_file_snapshot"))
    snapshot_profiles = dict_value(snapshot.get("profiles"))
    direct_profiles = dict_value(payload.get("profiles"))
    all_saved_vietmax = dict_value(payload.get("all_saved_vietmax_configs"))

    imported = empty_profile_config(storage_profile)
    for candidate in (
        snapshot_profiles.get(storage_profile),
        direct_profiles.get(storage_profile),
        all_saved_vietmax.get(storage_profile),
        payload.get("saved_profile_config"),
    ):
        if isinstance(candidate, dict):
            imported.update(candidate)
    if looks_like_vietmax_profile_config(payload):
        imported.update({key: payload.get(key) for key in VIETMAX_CONFIG_PROFILE_KEYS if key in payload})

    process_payload = dict_value(payload.get("process_payload"))
    imported.update(vietmax_profile_from_process_payload(process_payload, normalized_phase))
    if not process_payload:
        imported.update(vietmax_profile_from_process_payload(payload, normalized_phase))
    return storage_profile, imported


def review_merge_key(product, company="", mst="", company_key="", comparison_scope=VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES):
    product_part = vietmax_ban_ra_match_key(product)
    if not product_part:
        return ""
    if comparison_scope == VIETMAX_COMPARISON_SCOPE_SAME_COMPANY:
        identity = raw_text(company_key) or vietmax_company_identity_key(company, mst)
        return f"{identity}|||{product_part}" if identity else ""
    return product_part


def review_merge_group(item):
    if not isinstance(item, dict):
        return "other"
    group = raw_text(item.get("review_group"))
    if group:
        return group
    return "dimension_diff" if item.get("dimension_only") else "other"


def apply_saved_review_choices(rows, saved_merges, comparison_scope):
    if not isinstance(saved_merges, list) or not saved_merges:
        return rows
    scope = normalize_vietmax_comparison_scope(comparison_scope)
    forward = {}
    reverse = {}
    for item in saved_merges:
        if not isinstance(item, dict):
            continue
        item_scope = normalize_vietmax_comparison_scope(item.get("comparison_scope") or VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES)
        if item_scope != scope:
            continue
        group = review_merge_group(item)
        left_key = review_merge_key(item.get("product"), item.get("company"), item.get("mst"), item.get("company_key"), item_scope)
        right_key = review_merge_key(item.get("similar_product"), item.get("similar_company"), item.get("similar_mst"), item.get("similar_company_key"), item_scope)
        if left_key and right_key:
            forward[(item_scope, group, left_key, right_key)] = item
            reverse[(item_scope, group, right_key, left_key)] = item
    if not forward and not reverse:
        return rows
    restored = []
    for row in rows:
        next_row = dict(row)
        group = review_merge_group(row)
        left_key = review_merge_key(row.get("product"), row.get("company"), row.get("mst"), row.get("company_key"), scope)
        right_key = review_merge_key(row.get("similar_product"), row.get("similar_company"), row.get("similar_mst"), row.get("similar_company_key"), scope)
        forward_key = (scope, group, left_key, right_key)
        reverse_key = (scope, group, left_key, right_key)
        if left_key and right_key and forward_key in forward:
            saved = forward[forward_key]
            saved_confirmed = saved.get("confirmed") is True
            next_row.update({
                "confirmed": saved_confirmed,
                "code_choice": raw_text(saved.get("code_choice")) or ("similar" if saved_confirmed else raw_text(row.get("code_choice")) or "current"),
            })
            if raw_text(saved.get("split_code")):
                next_row["split_code"] = raw_text(saved.get("split_code"))
            if raw_text(saved.get("similar_split_code")):
                next_row["similar_split_code"] = raw_text(saved.get("similar_split_code"))
        elif left_key and right_key and reverse_key in reverse:
            saved = reverse[reverse_key]
            saved_confirmed = saved.get("confirmed") is True
            next_row.update({
                "confirmed": saved_confirmed,
                "code_choice": raw_text(saved.get("code_choice")) or ("current" if saved_confirmed else raw_text(row.get("code_choice")) or "current"),
            })
            if raw_text(saved.get("split_code")):
                next_row["similar_split_code"] = raw_text(saved.get("split_code"))
            if raw_text(saved.get("similar_split_code")):
                next_row["split_code"] = raw_text(saved.get("similar_split_code"))
        restored.append(next_row)
    return restored


class VietmaxReviewRequest(BaseModel):
    saved_name: str
    phase: str = VIETMAX_PHASE_PURCHASE
    comparison_scope: str = VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES
    product_col: str = "M"
    qty_col: str = "O"
    invoice_status_col: str = DEFAULT_INVOICE_STATUS_COL
    invoice_status_skip_values: list[str] | None = None
    price_col: str = "P"
    require_existing_code: bool = False
    word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None
    products: list[dict] | None = None
    operation_id: str = ""
    allow_same_code_split: bool = False


class VietmaxAnalyzeRequest(BaseModel):
    saved_name: str
    phase: str = VIETMAX_PHASE_PURCHASE
    company_col: str = "F"
    mst_col: str = "G"
    address_col: str = "H"
    product_col: str = "M"
    qty_col: str = "O"
    price_col: str = "P"
    invoice_status_col: str = DEFAULT_INVOICE_STATUS_COL
    invoice_status_skip_values: list[str] | None = None


class GenericAnalyzeRequest(BaseModel):
    saved_name: str
    original_name: str = "output.xlsx"
    profile: str = "son_phuong"
    vietmax_phase: str = VIETMAX_PHASE_PURCHASE
    company_col: str = "F"
    mst_col: str = "G"
    address_col: str = "H"
    product_col: str = "M"
    qty_col: str = "O"
    price_col: str = ""
    invoice_status_col: str = DEFAULT_INVOICE_STATUS_COL
    invoice_status_skip_values: list[str] | None = None
    operation_id: str = ""


class GenericReviewRequest(BaseModel):
    saved_name: str
    profile: str = "son_phuong"
    vietmax_phase: str = VIETMAX_PHASE_PURCHASE
    comparison_scope: str = VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES
    word_rules: dict[str, str] | None = None
    first_word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None
    products: list[dict] | None = None
    operation_id: str = ""


class VietmaxProductPreviewRequest(BaseModel):
    phase: str = VIETMAX_PHASE_PURCHASE
    products: list[str]
    word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None
    product_code_replacements: dict[str, str] | None = None


class GenericProductPreviewRequest(BaseModel):
    profile: str = "son_phuong"
    products: list[str]
    word_rules: dict[str, str] | None = None
    first_word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None
    product_code_replacements: dict[str, str] | None = None


class VietmaxSalesMatchRequest(BaseModel):
    sales_saved_name: str
    purchase_saved_name: str
    comparison_scope: str = VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES
    sales_company_col: str = "I"
    sales_mst_col: str = "J"
    product_col: str = "M"
    qty_col: str = "O"
    invoice_status_col: str = DEFAULT_INVOICE_STATUS_COL
    invoice_status_skip_values: list[str] | None = None
    sales_price_col: str = "P"
    purchase_price_col: str = "P"
    require_existing_purchase_code: bool = True
    operation_id: str = ""


class VietmaxProcessedFileStatsRequest(BaseModel):
    saved_name: str
    phase: str = VIETMAX_PHASE_PURCHASE


class VietmaxFastImportPackageRequest(BaseModel):
    profile: str = VIETMAX_PROFILE
    purchase_saved_name: str = ""
    sales_saved_name: str = ""
    purchase_original_saved_name: str = ""
    sales_original_saved_name: str = ""
    purchase_form_mapping_presets: list[dict] | None = None
    sales_form_mapping_presets: list[dict] | None = None
    purchase_company_group_assignments: dict | None = None
    sales_company_group_assignments: dict | None = None
    operation_id: str = ""


class VietmaxExportMatchesRequest(BaseModel):
    matches: list[dict]
    filename: str = "vietmax_khop_mua_ban.xls"


class VietmaxProcessRequest(BaseModel):
    saved_name: str
    original_name: str = "output.xlsx"
    payload: dict
    operation_id: str = ""
    cache_only: bool = False


class WorkflowProcessJobRequest(BaseModel):
    saved_name: str
    original_name: str = "output.xlsx"
    payload: dict
    processor: str = "vietmax"
    retry: bool = False


class WorkflowJsonJobRequest(BaseModel):
    kind: str
    payload: dict
    retry: bool = False


class LicenseActivationRequest(BaseModel):
    server_url: str = ""
    account_id: str = ""
    license_key: str


class EstimateWorkbookRequest(BaseModel):
    saved_name: str
    original_name: str = "du_toan.xlsx"
    bid_sheet_index: int | None = None
    detail_sheet_index: int | None = None
    bid_header_row: int | None = None
    detail_header_row: int | None = None
    bid_columns: dict | None = None
    detail_columns: dict | None = None


def estimate_warning_payload(result) -> dict:
    fields = [
        "identity_mismatches",
        "calculation_mismatches",
        "unclassified_rows",
        "thvt_mismatches",
        "thvt_key_mismatches",
        "thvt_missing_rows",
        "thvt_extra_rows",
    ]
    return {field: getattr(result, field, [])[:50] for field in fields}


def estimate_analysis_payload(result) -> dict:
    return {
        "summary": result.summary(),
        "warnings": estimate_warning_payload(result),
        "sheet_names": list(getattr(result, "sheet_names", []) or []),
    }


def safe_xlsx_download_name(original_name: str, suffix: str) -> str:
    stem = Path(original_name or "du_toan").stem or "du_toan"
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", stem).strip("._")[:80] or "du_toan"
    return f"{safe_stem}{suffix}.xlsx"


def create_app() -> FastAPI:
    api = FastAPI(title="ProductCodeFormatter API", version="0.1.0")
    api.add_middleware(
        CORSMiddleware,
        allow_origins=["http://127.0.0.1:5173", "http://localhost:5173"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=["X-Processed-Saved-Name", "X-Up-Saved-Name", "X-Estimate-Summary"],
    )

    @api.get("/api/health")
    def health() -> dict[str, str]:
        return {"status": "ok", "profile": "vietmax"}

    @api.get("/api/session/current")
    def current_workflow_session() -> dict:
        return WORKFLOW_SESSION_STORE.snapshot()

    @api.post("/api/session/close")
    def close_workflow_session() -> dict:
        WORKFLOW_SESSION_STORE.close()
        return {"status": "closed", "session_id": WORKFLOW_SESSION_STORE.session_id}

    @api.get("/api/jobs/{job_id}")
    def workflow_job(job_id: str) -> dict:
        job = WORKFLOW_JOB_MANAGER.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Không tìm thấy tiến trình xử lý.")
        return job

    @api.get("/api/artifacts/{artifact_id}")
    def workflow_artifact(artifact_id: str) -> FileResponse:
        artifact = WORKFLOW_SESSION_STORE.artifact(artifact_id)
        if not artifact or not artifact.get("valid"):
            raise HTTPException(status_code=404, detail="File cache không còn hợp lệ trong phiên làm việc này.")
        path = uploaded_path(artifact.get("saved_name"))
        filename = artifact.get("original_name") or path.name
        return FileResponse(path, filename=filename, media_type=mimetypes.guess_type(filename)[0] or "application/octet-stream")

    @api.post("/api/workflow/process-jobs")
    def start_workflow_process_job(payload: WorkflowProcessJobRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        process_payload = dict(payload.payload or {})
        phase = normalize_vietmax_phase(process_payload.get("vietmax_phase"))
        profile = profile_key(process_payload.get("profile") or VIETMAX_PROFILE)
        signature = stable_signature({
            "processor_version": 2,
            "processor": payload.processor,
            "profile": profile,
            "phase": phase,
            "source_sha256": file_sha256(path),
            "payload": process_payload,
        })
        kind = f"processed:{profile}:{phase}"

        def runner(progress):
            cached = WORKFLOW_SESSION_STORE.artifact_by_signature(kind, signature)
            if cached:
                return {"artifact": cached, "processed_saved_name": cached["saved_name"], "reused": True}
            return run_workflow_process_job(
                path,
                payload.original_name,
                process_payload,
                payload.processor,
                kind,
                signature,
                progress,
            )

        return WORKFLOW_JOB_MANAGER.start(
            kind,
            signature,
            runner,
            retry=payload.retry,
            context={"profile": profile, "phase": phase, "source": payload.original_name},
        )

    @api.post("/api/workflow/json-jobs")
    def start_workflow_json_job(payload: WorkflowJsonJobRequest) -> dict:
        kind = raw_text(payload.kind)
        request_payload = dict(payload.payload or {})
        request_payload.pop("operation_id", None)
        source_hashes = {}
        for key in ("saved_name", "purchase_saved_name", "sales_saved_name"):
            saved_name = raw_text(request_payload.get(key))
            if not saved_name:
                continue
            source_hashes[key] = file_sha256(uploaded_path(saved_name))
        signature = stable_signature({
            "operation_version": 2,
            "kind": kind,
            "payload": request_payload,
            "source_hashes": source_hashes,
        })

        def runner(progress):
            progress(0, 1, "Đang xử lý")
            try:
                if kind == "vietmax-review":
                    result = vietmax_review(VietmaxReviewRequest(**request_payload))
                elif kind == "generic-review":
                    result = generic_review(GenericReviewRequest(**request_payload))
                elif kind == "sales-match":
                    result = vietmax_sales_match(VietmaxSalesMatchRequest(**request_payload))
                elif kind == "fast-export":
                    response = create_vietmax_fast_import_package(VietmaxFastImportPackageRequest(**request_payload))
                    saved_name = f"fast_{uuid.uuid4().hex}.xls"
                    target = UPLOAD_DIR / saved_name
                    target.write_bytes(bytes(response.body))
                    artifact_kind = f"export:fast:{profile_key(request_payload.get('profile') or VIETMAX_PROFILE)}"
                    artifact = WORKFLOW_SESSION_STORE.register_file(
                        saved_name,
                        kind=artifact_kind,
                        original_name="vietmax_fast_import.xls",
                        signature=signature,
                        supersede_kind=True,
                    )
                    result = {"artifact": artifact}
                else:
                    raise WorkflowFailure(
                        "UNKNOWN_JOB_KIND",
                        f"Không hỗ trợ loại tiến trình: {kind}",
                        field="kind",
                        retryable=False,
                    )
            except HTTPException as exc:
                detail = exc.detail
                if isinstance(detail, dict):
                    raise WorkflowFailure(
                        detail.get("code") or "WORKFLOW_REQUEST_FAILED",
                        detail.get("message") or str(detail),
                        stage=detail.get("stage") or "",
                        field=detail.get("field") or "",
                        details=detail.get("details") or {},
                        retryable=detail.get("retryable", True),
                    ) from exc
                raise WorkflowFailure("WORKFLOW_REQUEST_FAILED", str(detail), retryable=True) from exc
            progress(1, 1, "Đã hoàn tất")
            return json_safe(result)

        return WORKFLOW_JOB_MANAGER.start(
            kind,
            signature,
            runner,
            retry=payload.retry,
            context={"kind": kind},
        )

    @api.get("/api/progress/{operation_id}")
    def operation_progress(operation_id: str) -> dict:
        return get_progress_job(operation_id)

    @api.get("/api/config")
    def get_config() -> dict:
        return save_config(load_config())

    @api.post("/api/config")
    def set_config(payload: dict) -> dict:
        if not isinstance(payload, dict):
            payload = {}
        if "license" not in payload:
            current = load_config()
            payload = dict(payload)
            payload["license"] = current.get("license") or {}
        return save_config(payload)

    @api.post("/api/config/profile/{profile}")
    def import_profile_config(profile: str, payload: dict) -> dict:
        profile = profile_key(profile)
        incoming = normalize_config(payload or {})
        current = load_config()
        incoming_profiles = incoming.get("profiles") if isinstance(incoming.get("profiles"), dict) else {}
        source_profile = profile if profile in incoming_profiles else incoming.get("selected_profile")
        source_profile = source_profile if source_profile in incoming_profiles else profile
        current["selected_profile"] = profile
        current.setdefault("columns", {}).update(incoming.get("columns") or {})
        current.setdefault("profiles", {})[profile] = incoming_profiles.get(source_profile, empty_profile_config(profile))
        return save_config(current)

    @api.get("/api/license/status")
    def license_status() -> dict:
        return public_license_status(load_config())

    @api.post("/api/license/activate")
    def activate_license(payload: LicenseActivationRequest) -> dict:
        try:
            license_cfg = activate_keygen_license(payload.server_url, payload.account_id, payload.license_key)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        cfg = load_config()
        cfg["license"] = license_cfg
        saved = save_config(cfg)
        return public_license_status(saved)

    @api.post("/api/license/reload")
    def reload_license() -> dict:
        current = load_config()
        license_cfg = current.get("license") or {}
        server_url = str(license_cfg.get("server_url") or "").strip()
        account_id = str(license_cfg.get("account_id") or "").strip()
        license_key = str(license_cfg.get("license_key") or "").strip()
        if not server_url or not account_id or not license_key:
            raise HTTPException(status_code=400, detail="Chưa có đủ License server, Account, hoặc License key để tải lại.")
        try:
            refreshed = activate_keygen_license(server_url, account_id, license_key)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        current["license"] = refreshed
        return public_license_status(save_config(current))

    @api.post("/api/files/upload")
    async def upload_excel(file: UploadFile = File(...), purpose: str = Form("source")) -> dict:
        original = file.filename or ""
        ext = Path(original).suffix.lower()
        if ext not in {".xls", ".xlsx", ".xlsm"}:
            raise HTTPException(status_code=400, detail="Please upload .xls, .xlsx or .xlsm only.")
        saved_name = f"{uuid.uuid4().hex}{ext}"
        path = UPLOAD_DIR / saved_name
        path.write_bytes(await file.read())
        try:
            _, df = read_workbook(path)
        except Exception as exc:
            path.unlink(missing_ok=True)
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        summary = workbook_summary(df, original, saved_name)
        safe_purpose = re.sub(r"[^a-z0-9_-]+", "-", raw_text(purpose).lower()).strip("-") or "source"
        artifact = WORKFLOW_SESSION_STORE.register_file(
            saved_name,
            kind=f"source:{safe_purpose}",
            original_name=original,
            metadata={"summary": summary},
        )
        return {
            **summary,
            "artifact_id": artifact["artifact_id"],
            "session_id": WORKFLOW_SESSION_STORE.session_id,
        }

    @api.post("/api/templates/upload")
    async def upload_form_template(file: UploadFile = File(...)) -> dict:
        original = file.filename or ""
        ext = Path(original).suffix.lower()
        if ext not in {".xls", ".xlsx", ".xlsm"}:
            raise HTTPException(status_code=400, detail="Chỉ hỗ trợ file mẫu Excel .xls, .xlsx hoặc .xlsm.")
        saved_name = f"template_{uuid.uuid4().hex}{ext}"
        path = TEMPLATE_DIR / saved_name
        path.write_bytes(await file.read())
        try:
            _, df = read_workbook(path)
        except Exception as exc:
            path.unlink(missing_ok=True)
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {**workbook_summary(df, original, saved_name), "persistent_template": True}

    @api.post("/api/estimate/upload")
    async def upload_estimate_workbook(file: UploadFile = File(...)) -> dict:
        original = file.filename or ""
        ext = Path(original).suffix.lower()
        if ext not in {".xls", ".xlsx", ".xlsm"}:
            raise HTTPException(status_code=400, detail="Chỉ hỗ trợ file Excel .xls, .xlsx hoặc .xlsm.")
        saved_name = f"estimate_{uuid.uuid4().hex}{ext}"
        path = UPLOAD_DIR / saved_name
        path.write_bytes(await file.read())
        try:
            sheet_info = list_estimate_workbook_sheets(str(path))
        except Exception as exc:
            path.unlink(missing_ok=True)
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        artifact = WORKFLOW_SESSION_STORE.register_file(saved_name, kind="estimate-source", original_name=original)
        return json_safe({
            "original_name": original,
            "saved_name": saved_name,
            "artifact_id": artifact["artifact_id"],
            "session_id": WORKFLOW_SESSION_STORE.session_id,
            "size": path.stat().st_size,
            **sheet_info,
        })

    @api.post("/api/estimate/analyze")
    def analyze_estimate(payload: EstimateWorkbookRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        try:
            analysis = analyze_estimate_workbook(
                str(path),
                bid_sheet_index=payload.bid_sheet_index,
                detail_sheet_index=payload.detail_sheet_index,
                thvt_sheet_index=None,
                bid_header_row=payload.bid_header_row,
                detail_header_row=payload.detail_header_row,
                bid_columns=payload.bid_columns,
                detail_columns=payload.detail_columns,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return json_safe(estimate_analysis_payload(analysis))

    @api.post("/api/estimate/export")
    def export_estimate(payload: EstimateWorkbookRequest) -> Response:
        path = uploaded_path(payload.saved_name)
        try:
            content, summary = create_estimate_output_workbook(
                str(path),
                bid_sheet_index=payload.bid_sheet_index,
                detail_sheet_index=payload.detail_sheet_index,
                bid_header_row=payload.bid_header_row,
                detail_header_row=payload.detail_header_row,
                bid_columns=payload.bid_columns,
                detail_columns=payload.detail_columns,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        filename = safe_xlsx_download_name(payload.original_name, "_boc_tach")
        return Response(
            content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "X-Estimate-Summary": json.dumps(summary, ensure_ascii=False),
            },
        )

    @api.get("/api/files/download/{saved_name}")
    def download_uploaded_or_cached_file(saved_name: str) -> FileResponse:
        path = uploaded_path(saved_name)
        return FileResponse(path, filename=path.name, media_type=mimetypes.guess_type(path.name)[0] or "application/octet-stream")

    @api.post("/api/mapping")
    async def mapping(file: UploadFile = File(...)) -> dict:
        return await upload_excel(file)

    @api.get("/api/vietmax/format-mapping-defaults")
    def vietmax_format_mapping_defaults() -> dict:
        return json_safe({
            "source_columns": {
                "purchase": vietmax_default_source_columns(VIETMAX_PHASE_PURCHASE),
                "sales": vietmax_default_source_columns(VIETMAX_PHASE_SALES),
            },
            "form_mapping_presets": load_system_default_form_mappings(),
        })

    @api.post("/api/invoice_statuses")
    def invoice_statuses(payload: dict) -> dict:
        path = uploaded_path(str(payload.get("saved_name", "")))
        try:
            _, df = read_workbook(path)
            statuses = invoice_status_options(
                df,
                str(payload.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL) or "").upper(),
                payload.get("invoice_status_skip_values"),
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return {"invoice_statuses": statuses}

    @api.post("/api/check")
    def check(payload: GenericAnalyzeRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        profile = profile_key(payload.profile)
        cfg = load_config()
        profile_cfg = scoped_profile_config(cfg, profile, payload.vietmax_phase)
        try:
            result = analyze(
                path,
                payload.company_col,
                payload.mst_col,
                payload.address_col,
                payload.product_col,
                payload.qty_col,
                payload.price_col,
                profile_cfg,
                invoice_status_col=payload.invoice_status_col,
                invoice_status_skip_values=payload.invoice_status_skip_values,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        removed = profile_cfg.get("removed_companies") or {}
        for company in result.get("companies", []):
            mst = raw_text(company.get("mst"))
            company_id = raw_text(company.get("company_id")) or mst
            process = not bool(removed.get(company_id) or (mst and removed.get(mst)))
            company["process"] = process
            company["pending_process"] = process
            company["committed_prefix"] = company.get("value") or ""
        result.update({
            "original_name": payload.original_name,
            "saved_name": payload.saved_name,
            "manual_code_overrides": profile_cfg.get("manual_code_overrides") or {},
            "product_code_replacements": profile_cfg.get("product_code_replacements") or {},
            "word_rules": profile_cfg.get("word_rules") or {},
            "first_word_rules": profile_cfg.get("first_word_rules") or {},
            "repeated_phrase_removals": profile_cfg.get("repeated_phrase_removals") or [],
            "inventory_pairs": profile_cfg.get("inventory_pairs") or [],
            "use_default_inventory_pair": bool(profile_cfg.get("use_default_inventory_pair")),
            "default_inventory_pair_id": profile_cfg.get("default_inventory_pair_id") or "",
            "inventory_pair_rules": profile_cfg.get("inventory_pair_rules") or [],
            "inventory_allocation_config": profile_cfg.get("inventory_allocation_config") or {},
            "include_company_prefix": profile_cfg.get("include_company_prefix") is not False,
            "prefix_strategy": profile_cfg.get("prefix_strategy") or "last_2_words",
            "prefix_mst_digits": profile_cfg.get("prefix_mst_digits") or 3,
            "prefix_name_words": profile_cfg.get("prefix_name_words") or 2,
            "prefix_name_chars": profile_cfg.get("prefix_name_chars") or 1,
            "prefix_missing_mst_strategy": profile_cfg.get("prefix_missing_mst_strategy") or "all_name_words",
            "prefix_strategy_values": profile_cfg.get("prefix_strategy_values") or {},
            "processing_groups": profile_cfg.get("processing_groups") or [],
            "company_group_assignments": profile_cfg.get("company_group_assignments") or {},
            "form_mapping_presets": profile_cfg.get("form_mapping_presets") or [],
            "product_review_merges": profile_cfg.get("product_review_merges") or [],
            "price_range_rules": profile_cfg.get("price_range_rules") or {},
            "price_adjust_all_percent": profile_cfg.get("price_adjust_all_percent") or 0,
            "columns": profile_cfg.get("columns") or {},
        })
        return json_safe(result)

    @api.post("/api/vietmax/review")
    def vietmax_review(payload: VietmaxReviewRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        operation_id = raw_text(payload.operation_id)
        update_progress_job(operation_id, 0, 1, "Đang chuẩn bị review Mã VT", status="running")
        try:
            def progress(done, total, label):
                update_progress_job(operation_id, done, total, label, status="running")

            if payload.phase == VIETMAX_PHASE_SALES:
                products = payload.products if payload.products is not None else vietmax_ban_ra_sales_products_from_workbook(
                    path,
                    payload.product_col,
                    payload.qty_col,
                    payload.invoice_status_col,
                    payload.invoice_status_skip_values,
                    progress_callback=progress,
                    comparison_scope=payload.comparison_scope,
                    price_col=payload.price_col,
                )
                rows = vietmax_product_review_rows(
                    enrich_vietmax_review_products(products, "sales", payload),
                    "sales_product",
                    "sales_unit",
                    comparison_scope=payload.comparison_scope,
                    progress_callback=progress,
                    allow_same_code_split=payload.allow_same_code_split,
                )
            else:
                products = payload.products if payload.products is not None else vietmax_purchase_products_from_workbook(
                    path,
                    price_col=payload.price_col,
                    progress_callback=progress,
                    comparison_scope=payload.comparison_scope,
                    require_existing_code=payload.require_existing_code,
                )
                rows = vietmax_product_review_rows(enrich_vietmax_review_products(products, "purchase", payload), "purchase_product", "purchase_unit", comparison_scope=payload.comparison_scope, progress_callback=progress)
            saved_profile = vietmax_profile_config(VIETMAX_PHASE_SALES if payload.phase == VIETMAX_PHASE_SALES else VIETMAX_PHASE_PURCHASE)
            saved_merges = saved_profile.get("vietmax_ban_ra_sales_internal_merges") if payload.phase == VIETMAX_PHASE_SALES else saved_profile.get("vietmax_mua_vao_internal_merges")
            rows = apply_saved_review_choices(rows, saved_merges or [], payload.comparison_scope)
        except Exception as exc:
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, 1, 1, "Đã tạo xong danh sách review Mã VT", status="complete")
        return {"products": json_safe(products), "review_rows": json_safe(rows)}

    @api.post("/api/review")
    def generic_review(payload: GenericReviewRequest) -> dict:
        operation_id = raw_text(payload.operation_id)
        update_progress_job(operation_id, 0, 1, "Dang chuan bi review Ma VT", status="running")
        try:
            profile = profile_key(payload.profile)

            def progress(done, total, label):
                update_progress_job(operation_id, done, total, label, status="running")

            products = payload.products or []
            rows = vietmax_product_review_rows(
                enrich_generic_review_products(products, profile, payload),
                "purchase_product",
                "purchase_unit",
                comparison_scope=payload.comparison_scope,
                progress_callback=progress,
            )
            profile_cfg = scoped_profile_config(load_config(), profile, payload.vietmax_phase)
            rows = apply_saved_review_choices(rows, profile_cfg.get("product_review_merges") or [], payload.comparison_scope)
        except Exception as exc:
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, 1, 1, "Da tao xong danh sach review Ma VT", status="complete")
        return {"products": json_safe(products), "review_rows": json_safe(rows)}

    @api.post("/api/vietmax/analyze")
    def vietmax_analyze(payload: VietmaxAnalyzeRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        try:
            phase = VIETMAX_PHASE_SALES if payload.phase == VIETMAX_PHASE_SALES else VIETMAX_PHASE_PURCHASE
            profile_cfg = vietmax_profile_config(phase)
            result = analyze(
                path,
                payload.company_col,
                payload.mst_col,
                payload.address_col,
                payload.product_col,
                payload.qty_col,
                payload.price_col,
                profile_cfg,
                invoice_status_col=payload.invoice_status_col,
                invoice_status_skip_values=payload.invoice_status_skip_values,
            )
            removed = profile_cfg.get("removed_companies") or {}
            for company in result.get("companies", []):
                mst = raw_text(company.get("mst"))
                company_id = raw_text(company.get("company_id")) or mst
                process = not bool(removed.get(company_id) or (mst and removed.get(mst)))
                company["process"] = process
                company["pending_process"] = process
            result["manual_code_overrides"] = profile_cfg.get("manual_code_overrides") or {}
            result["product_code_replacements"] = profile_cfg.get("product_code_replacements") or {}
            result["word_rules"] = profile_cfg.get("word_rules") or {}
            result["repeated_phrase_removals"] = profile_cfg.get("repeated_phrase_removals") or []
            result["inventory_pairs"] = profile_cfg.get("inventory_pairs") or []
            result["use_default_inventory_pair"] = bool(profile_cfg.get("use_default_inventory_pair"))
            result["default_inventory_pair_id"] = profile_cfg.get("default_inventory_pair_id") or ""
            result["inventory_pair_rules"] = profile_cfg.get("inventory_pair_rules") or []
            result["inventory_allocation_config"] = profile_cfg.get("inventory_allocation_config") or {}
            result["sales_match_rules"] = profile_cfg.get("vietmax_ban_ra_purchase_match_rules") or []
            result["vietmax_mua_vao_internal_merges"] = profile_cfg.get("vietmax_mua_vao_internal_merges") or []
            result["vietmax_ban_ra_sales_internal_merges"] = profile_cfg.get("vietmax_ban_ra_sales_internal_merges") or []
            result["include_company_prefix"] = profile_cfg.get("include_company_prefix") is not False
            result["prefix_strategy"] = profile_cfg.get("prefix_strategy") or "last_2_words"
            result["prefix_mst_digits"] = profile_cfg.get("prefix_mst_digits") or 3
            result["prefix_name_words"] = profile_cfg.get("prefix_name_words") or 2
            result["prefix_name_chars"] = profile_cfg.get("prefix_name_chars") or 1
            result["prefix_missing_mst_strategy"] = profile_cfg.get("prefix_missing_mst_strategy") or "all_name_words"
            result["prefix_strategy_values"] = profile_cfg.get("prefix_strategy_values") or {}
            result["processing_groups"] = profile_cfg.get("processing_groups") or []
            result["company_group_assignments"] = profile_cfg.get("company_group_assignments") or {}
            result["form_mapping_presets"] = profile_cfg.get("form_mapping_presets") or []
            return json_safe(result)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @api.post("/api/export_price_report")
    def export_price_report(payload: dict) -> Response:
        sheets = payload.get("sheets") or []
        if not isinstance(sheets, list) or not sheets:
            raise HTTPException(status_code=400, detail="Không có dữ liệu để xuất Excel.")
        try:
            stream = make_excel_workbook(sheets)
            filename = safe_download_name(str(payload.get("filename") or "bao_cao_ban_hang.xlsx"))
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return Response(
            stream.getvalue(),
            media_type=XLS_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    @api.post("/api/process")
    def process_generic(payload: dict):
        saved_name = str(payload.get("saved_name", ""))
        original_name = str(payload.get("original_name") or "output.xlsx")
        path = uploaded_path(saved_name)
        process_payload = dict(payload)
        operation_id = raw_text(process_payload.get("operation_id"))
        cache_only = bool(process_payload.get("cache_only"))
        started_at = time.perf_counter()
        profile_for_log = profile_key(process_payload.get("profile", "son_phuong"))
        phase_for_log = normalize_vietmax_phase(process_payload.get("vietmax_phase"))
        update_progress_job(operation_id, 0, 5, "Đang chuẩn bị tạo file", status="running")

        def progress(done: int, total: int, label: str) -> None:
            update_progress_job(operation_id, done, total, label, status="running")

        try:
            out = resolve_output_path(original_name, "")
            diagnostic_log(
                "generic/process start "
                f"operation={operation_id or '-'} profile={profile_for_log} phase={phase_for_log} "
                f"cache_only={cache_only} export_form_mappings={bool(process_payload.get('export_form_mappings'))} "
                f"input={path} output={out}"
            )
            processed_df = process_workbook(path, out, process_payload, progress_callback=progress)
            diagnostic_log(
                "generic/process workbook_done "
                f"operation={operation_id or '-'} profile={profile_for_log} phase={phase_for_log} "
                f"elapsed={time.perf_counter() - started_at:.2f}s rows={len(processed_df.index)} cols={len(processed_df.columns)}"
            )
            cfg = load_config()
            profile = profile_key(process_payload.get("profile", cfg.get("selected_profile", "son_phuong")))
            phase = normalize_vietmax_phase(process_payload.get("vietmax_phase"))
            uses_price_rules = profile == "cao_thanh"
            cfg["selected_profile"] = profile
            _, profile_cfg = ensure_profile_scope(cfg, profile, phase)
            old_profile = profile_cfg or empty_profile_config(profile, include_scopes=False)
            merged_price_ranges = merge_price_ranges(old_profile.get("price_range_rules"), process_payload.get("price_range_rules", {})) if uses_price_rules else {}
            columns = {
                "company_col": process_payload.get("company_col", ""),
                "mst_col": process_payload.get("mst_col", ""),
                "address_col": process_payload.get("address_col", ""),
                "product_col": process_payload.get("product_col", ""),
                "qty_col": process_payload.get("qty_col", ""),
                "price_col": process_payload.get("price_col", ""),
                "output_col": process_payload.get("output_col", ""),
                "invoice_status_col": process_payload.get("invoice_status_col", DEFAULT_INVOICE_STATUS_COL),
                "invoice_status_skip_values": process_payload.get("invoice_status_skip_values", DEFAULT_INVOICE_STATUS_SKIP_VALUES),
            }
            cfg.setdefault("columns", {}).update(columns)
            profile_cfg.update({
                "prefixes": keep_company_config_when_unloaded(process_payload, old_profile, "prefixes", "prefixes", {}),
                "selected_products": keep_company_config_when_unloaded(process_payload, old_profile, "skipped_products_map", "selected_products", {}),
                "removed_companies": keep_company_config_when_unloaded(process_payload, old_profile, "removed_companies", "removed_companies", {}),
                "word_rules": keep_existing_when_empty(process_payload, old_profile, "word_rules", {}),
                "first_word_rules": keep_existing_when_empty(process_payload, old_profile, "first_word_rules", {}),
                "repeated_phrase_removals": normalize_phrase_list(keep_existing_when_empty(process_payload, old_profile, "repeated_phrase_removals", [])),
                "price_group_rules": process_payload.get("price_group_rules", {}) if uses_price_rules else {},
                "price_range_rules": merged_price_ranges,
                "price_adjust_all_percent": float(process_payload.get("price_adjust_all_percent") or 0) if uses_price_rules else 0,
                "manual_code_overrides": keep_company_config_when_unloaded(process_payload, old_profile, "manual_code_overrides", "manual_code_overrides", {}),
                "product_code_replacements": normalize_product_code_replacements(keep_existing_when_empty(process_payload, old_profile, "product_code_replacements", {})),
                "product_review_merges": process_payload.get("product_review_merges", old_profile.get("product_review_merges", [])),
                "vietmax_mua_vao_internal_merges": process_payload.get("vietmax_mua_vao_internal_merges", old_profile.get("vietmax_mua_vao_internal_merges", [])),
                "vietmax_ban_ra_sales_internal_merges": process_payload.get("vietmax_ban_ra_sales_internal_merges", old_profile.get("vietmax_ban_ra_sales_internal_merges", [])),
                "vietmax_ban_ra_purchase_match_rules": process_payload.get("vietmax_ban_ra_purchase_match_rules", old_profile.get("vietmax_ban_ra_purchase_match_rules", [])),
                "inventory_pairs": normalize_inventory_pairs(keep_existing_when_empty(process_payload, old_profile, "inventory_pairs", [])),
                "use_default_inventory_pair": bool(process_payload.get("use_default_inventory_pair", old_profile.get("use_default_inventory_pair", False))),
                "default_inventory_pair_id": str(process_payload.get("default_inventory_pair_id", old_profile.get("default_inventory_pair_id", "")) or "").strip(),
                "inventory_pair_rules": normalize_inventory_pair_rules(keep_existing_when_empty(process_payload, old_profile, "inventory_pair_rules", [])),
                "inventory_allocation_config": keep_existing_when_empty(process_payload, old_profile, "inventory_allocation_config", {}),
                "include_company_prefix": process_payload.get("include_company_prefix") is not False,
                "prefix_strategy": process_payload.get("prefix_strategy", old_profile.get("prefix_strategy", "last_2_words")),
                "prefix_mst_digits": process_payload.get("prefix_mst_digits", old_profile.get("prefix_mst_digits", 3)),
                "prefix_name_words": process_payload.get("prefix_name_words", old_profile.get("prefix_name_words", 2)),
                "prefix_name_chars": process_payload.get("prefix_name_chars", old_profile.get("prefix_name_chars", 1)),
                "prefix_missing_mst_strategy": normalize_missing_mst_prefix_strategy(process_payload.get("prefix_missing_mst_strategy", old_profile.get("prefix_missing_mst_strategy", "all_name_words"))),
                "prefix_strategy_values": normalize_prefix_strategy_values(keep_company_config_when_unloaded(process_payload, old_profile, "prefix_strategy_values", "prefix_strategy_values", {})),
                "processing_groups": keep_existing_when_empty(process_payload, old_profile, "processing_groups", []),
                "company_group_assignments": keep_company_config_when_unloaded(process_payload, old_profile, "company_group_assignments", "company_group_assignments", {}),
                "form_mapping_presets": keep_form_mapping_presets(process_payload, old_profile),
                "columns": columns,
            })
            sync_vietmax_legacy_profile(cfg, profile, phase, profile_cfg)
            save_config(cfg)
            if process_payload.get("export_form_mappings"):
                progress(3, 5, "Đang dựng workbook form mapping từ FDI đã xử lý")
                input_phase = generic_profile_form_input_phase(profile)
                if "form_mapping_presets" in process_payload:
                    form_mapping_presets = process_payload.get("form_mapping_presets") if isinstance(process_payload.get("form_mapping_presets"), list) else []
                else:
                    saved_form_mapping_presets = profile_cfg.get("form_mapping_presets")
                    form_mapping_presets = saved_form_mapping_presets if isinstance(saved_form_mapping_presets, list) else load_system_default_form_mappings(input_phase)
                final_stream = generic_form_mapping_export_workbook_stream(
                    processed_df,
                    form_mapping_presets=form_mapping_presets,
                    company_group_assignments=process_payload.get("company_group_assignments") or profile_cfg.get("company_group_assignments"),
                    profile=profile,
                    progress_callback=progress,
                )
                safe_stem = re.sub(r'[^A-Za-z0-9_.-]+', "_", Path(original_name or out.name).stem).strip("._") or "ket_qua"
                filename = f"{safe_stem}_fast.xls"
                if cache_only:
                    cached_saved_name = cache_workbook_stream(final_stream, f"{profile}_fast")
                    update_progress_job(operation_id, 5, 5, "Đã cache workbook form mapping", status="complete")
                    diagnostic_log(f"generic/process cache complete operation={operation_id or '-'} profile={profile} elapsed={time.perf_counter() - started_at:.2f}s saved={cached_saved_name}")
                    return {"processedSavedName": cached_saved_name}
                update_progress_job(operation_id, 5, 5, "Đã tạo workbook form mapping", status="complete")
                return Response(
                    final_stream.getvalue(),
                    media_type=XLS_MEDIA_TYPE,
                    headers={"Content-Disposition": f'attachment; filename="{filename}"'},
                )
            if cache_only:
                cached_saved_name = cache_processed_workbook(out)
                update_progress_job(operation_id, 5, 5, "Đã cache file kết quả", status="complete")
                diagnostic_log(f"generic/process cache complete operation={operation_id or '-'} profile={profile} elapsed={time.perf_counter() - started_at:.2f}s saved={cached_saved_name}")
                return {"processedSavedName": cached_saved_name}
        except Exception as exc:
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            diagnostic_log(
                "generic/process error "
                f"operation={operation_id or '-'} profile={profile_for_log} phase={phase_for_log} "
                f"elapsed={time.perf_counter() - started_at:.2f}s error={exc}"
            )
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, 5, 5, "Đã tạo file kết quả", status="complete")
        return Response(
            out.read_bytes(),
            media_type=XLS_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{out.name}"'},
        )

    @api.post("/api/vietmax/product-preview")
    def vietmax_product_preview(payload: VietmaxProductPreviewRequest) -> dict:
        profile = "vietmax_ban_ra" if payload.phase == VIETMAX_PHASE_SALES else "vietmax_mua_vao"
        word_rules = payload.word_rules or {}
        repeated = payload.repeated_phrase_removals or []
        replacements = normalize_product_code_replacements(payload.product_code_replacements or {})
        codes = {}
        for product in payload.products:
            name = raw_text(product)
            if name:
                codes[name] = apply_product_code_replacement(make_product_part(profile, name, word_rules, repeated_phrase_removals=repeated), replacements)
        return {"codes": codes}

    @api.post("/api/product-preview")
    def generic_product_preview(payload: GenericProductPreviewRequest) -> dict:
        profile = profile_key(payload.profile)
        word_rules = payload.word_rules or {}
        first_word_rules = payload.first_word_rules or {}
        repeated = payload.repeated_phrase_removals or []
        replacements = normalize_product_code_replacements(payload.product_code_replacements or {})
        codes = {}
        for product in payload.products:
            name = raw_text(product)
            if name:
                codes[name] = apply_product_code_replacement(make_product_part(profile, name, word_rules, first_word_rules, repeated), replacements)
        return {"codes": codes}

    @api.post("/api/product-code-replacements/import")
    async def import_product_code_replacements(file: UploadFile = File(...)) -> dict:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="File import đang trống.")
        try:
            workbook_content = uploaded_workbook_content_for_openpyxl(content, file.filename)
            df = pd.read_excel(BytesIO(workbook_content), sheet_name=0, header=None, dtype=object, engine="openpyxl")
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Không đọc được file import đổi mã VT: {exc}") from exc
        replacements: dict[str, str] = {}
        skipped_header = False
        for _, row in df.iterrows():
            old_raw = raw_text(row.iloc[0]) if len(row) > 0 else ""
            new_raw = raw_text(row.iloc[1]) if len(row) > 1 else ""
            old_norm = re.sub(r"\s+", " ", raw_text(old_raw)).strip().lower()
            new_norm = re.sub(r"\s+", " ", raw_text(new_raw)).strip().lower()
            if not skipped_header and (
                ("gốc" in old_norm or "goc" in old_norm or "old" in old_norm or "original" in old_norm)
                and ("mới" in new_norm or "moi" in new_norm or "new" in new_norm)
            ):
                skipped_header = True
                continue
            old_code = sanitize_product_code(old_raw)
            new_code = sanitize_product_code(new_raw)
            if old_code and new_code:
                replacements[old_code] = new_code
        normalized = normalize_product_code_replacements(replacements)
        return {"product_code_replacements": normalized, "count": len(normalized)}

    @api.post("/api/vietmax/sales-match")
    def vietmax_sales_match(payload: VietmaxSalesMatchRequest) -> dict:
        sales_path = uploaded_path(payload.sales_saved_name)
        purchase_path = uploaded_path(payload.purchase_saved_name)
        operation_id = raw_text(payload.operation_id)
        update_progress_job(operation_id, 0, 1, "Đang chuẩn bị khớp mua vào / bán ra", status="running")
        try:
            def progress(done, total, label):
                update_progress_job(operation_id, done, total, label, status="running")

            sales_products = vietmax_ban_ra_sales_products_from_workbook(
                sales_path,
                payload.product_col,
                payload.qty_col,
                payload.invoice_status_col,
                payload.invoice_status_skip_values,
                progress_callback=progress,
                comparison_scope=payload.comparison_scope,
                price_col=payload.sales_price_col,
                company_col=payload.sales_company_col,
                mst_col=payload.sales_mst_col,
            )
            purchase_products = vietmax_purchase_products_from_workbook(
                purchase_path,
                price_col=payload.purchase_price_col,
                progress_callback=progress,
                comparison_scope=payload.comparison_scope,
                require_existing_code=payload.require_existing_purchase_code,
            )
            update_progress_job(operation_id, 0, max(1, len(sales_products)), "Đang khớp chính xác KVT/152", status="running")
            exact_matches = build_vietmax_khh_exact_purchase_matches(sales_products, purchase_products, payload.comparison_scope)
            update_progress_job(operation_id, 0, max(1, len(sales_products)), "Đang so khớp hàng bán/mua", status="running")
            fuzzy_matches = build_vietmax_ban_ra_purchase_matches(sales_products, purchase_products, progress_callback=progress, comparison_scope=payload.comparison_scope)
            matches = merge_exact_and_fuzzy_matches(exact_matches, fuzzy_matches)
            match_rules = vietmax_profile_config(VIETMAX_PHASE_SALES).get("vietmax_ban_ra_purchase_match_rules") or []
        except Exception as exc:
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, max(1, len(matches)), max(1, len(matches)), "Đã tạo xong danh sách khớp mua/bán", status="complete")
        return {
            "sales_products": json_safe(sales_products),
            "purchase_products": json_safe(purchase_products),
            "exact_matches": json_safe(exact_matches),
            "matches": json_safe(matches),
            "match_rules": json_safe(match_rules),
        }

    @api.post("/api/vietmax/validate-processed-purchase")
    def validate_processed_purchase(payload: dict) -> dict:
        path = uploaded_path(str(payload.get("saved_name", "")))
        try:
            return validate_vietmax_processed_purchase_workbook(path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @api.post("/api/vietmax/processed-file-stats")
    def processed_file_stats(payload: VietmaxProcessedFileStatsRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        try:
            return inspect_vietmax_processed_file(path, payload.phase)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @api.post("/api/vietmax/validate-fast-import-processed")
    def validate_fast_import_processed(payload: VietmaxProcessedFileStatsRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        label = "FDI mua vào đã xử lý" if payload.phase == VIETMAX_PHASE_PURCHASE else "FDI bán ra đã xử lý"
        try:
            _, df = read_workbook(path)
            return validate_fast_import_processed_dataframe(df, label)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @api.post("/api/vietmax/fast-import-package")
    def create_vietmax_fast_import_package(payload: VietmaxFastImportPackageRequest) -> Response:
        operation_id = raw_text(payload.operation_id)
        started_at = time.perf_counter()
        profile_for_log = profile_key(payload.profile)
        update_progress_job(operation_id, 0, 1, "Đang chuẩn bị tạo workbook FAST 5 sheet", status="running")
        try:
            diagnostic_log(
                "fast-import-package start "
                f"operation={operation_id or '-'} profile={profile_for_log} "
                f"purchase={raw_text(payload.purchase_saved_name) or '-'} sales={raw_text(payload.sales_saved_name) or '-'}"
            )
            purchase_df = None
            sales_df = None
            if raw_text(payload.purchase_saved_name):
                _, purchase_df = read_workbook(uploaded_path(payload.purchase_saved_name))
            if raw_text(payload.sales_saved_name):
                _, sales_df = read_workbook(uploaded_path(payload.sales_saved_name))
            if purchase_df is None and sales_df is None:
                raise ValueError("Cần có ít nhất một file FDI đã xử lý để tạo file import.")

            cfg = load_config()
            profiles = cfg.get("profiles", {}) if isinstance(cfg.get("profiles"), dict) else {}
            profile = profile_key(payload.profile)
            purchase_profile = scoped_profile_config_from_profiles(profiles, profile, VIETMAX_PHASE_PURCHASE)
            sales_profile = scoped_profile_config_from_profiles(profiles, profile, VIETMAX_PHASE_SALES)
            purchase_forms = payload.purchase_form_mapping_presets if payload.purchase_form_mapping_presets is not None else purchase_profile.get("form_mapping_presets")
            sales_forms = payload.sales_form_mapping_presets if payload.sales_form_mapping_presets is not None else sales_profile.get("form_mapping_presets")
            form_mapping_presets = fast_merge_mapping_forms(purchase_forms, sales_forms)
            validate_form_mapping_columns(form_mapping_presets)
            purchase_assignments = payload.purchase_company_group_assignments if payload.purchase_company_group_assignments is not None else purchase_profile.get("company_group_assignments")
            sales_assignments = payload.sales_company_group_assignments if payload.sales_company_group_assignments is not None else sales_profile.get("company_group_assignments")

            def progress(done, total, label):
                update_progress_job(operation_id, done, total, label, status="running")

            stream = fast_import_multi_sheet_workbook(
                purchase_df,
                sales_df,
                progress_callback=progress,
                form_mapping_presets=form_mapping_presets,
                purchase_company_group_assignments=purchase_assignments,
                sales_company_group_assignments=sales_assignments,
            )
            diagnostic_log(
                "fast-import-package workbook_done "
                f"operation={operation_id or '-'} profile={profile_for_log} elapsed={time.perf_counter() - started_at:.2f}s"
            )
        except WorkflowFailure as exc:
            exc.operation_id = exc.operation_id or operation_id
            diagnostic_log(
                "fast-import-package validation_error "
                f"operation={operation_id or '-'} profile={profile_for_log} elapsed={time.perf_counter() - started_at:.2f}s error={exc}"
            )
            update_progress_job(operation_id, 1, 1, exc.message, status="error")
            raise HTTPException(status_code=400, detail=exc.as_dict()) from exc
        except Exception as exc:
            diagnostic_log(
                "fast-import-package error "
                f"operation={operation_id or '-'} profile={profile_for_log} elapsed={time.perf_counter() - started_at:.2f}s error={exc}"
            )
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, 1, 1, "Đã tạo xong workbook FAST 5 sheet", status="complete")
        return Response(
            stream.getvalue(),
            media_type=XLS_MEDIA_TYPE,
            headers={"Content-Disposition": 'attachment; filename="vietmax_fast_import.xls"'},
        )

    @api.post("/api/vietmax/export-matches")
    def export_matches(payload: VietmaxExportMatchesRequest) -> Response:
        rows = vietmax_purchase_match_export_rows(payload.matches)
        headers = [
            "Dùng",
            "Hàng bán ra",
            "ĐVT bán",
            "Đơn giá bán",
            "Số HD bán",
            "Ngày HD bán",
            "Mã VT mua vào",
            "Hàng mua vào",
            "ĐVT mua",
            "Đơn giá mua",
            "Cảnh báo",
            "Quy đổi",
            "Khác biệt",
            "Độ giống",
        ]
        stream = make_excel_workbook([{"name": "Khớp mua bán", "headers": headers, "rows": [dict(zip(headers, row)) for row in rows]}])
        return Response(
            stream.getvalue(),
            media_type=XLS_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{safe_download_name(payload.filename)}"'},
        )


    @api.post("/api/vietmax/save-config")
    def save_vietmax_config(payload: dict) -> dict:
        cfg = load_config()
        profile = profile_key(payload.get("profile") or VIETMAX_PROFILE)
        phase = normalize_vietmax_phase(payload.get("vietmax_phase"))
        cfg["selected_profile"] = profile
        cfg.setdefault("columns", {}).update(payload.get("columns") or {})
        _, profile_cfg = ensure_profile_scope(cfg, profile, phase)
        previous_profile_signature = stable_signature(profile_cfg)
        profile_cfg.update({
            "prefixes": keep_company_config_when_unloaded(payload, profile_cfg, "prefixes", "prefixes", {}),
            "selected_products": keep_company_config_when_unloaded(payload, profile_cfg, "skipped_products_map", "selected_products", {}),
            "removed_companies": keep_company_config_when_unloaded(payload, profile_cfg, "removed_companies", "removed_companies", {}),
            "include_company_prefix": payload.get("include_company_prefix") is not False,
            "prefix_strategy": payload.get("prefix_strategy", profile_cfg.get("prefix_strategy", "last_2_words")),
            "prefix_mst_digits": payload.get("prefix_mst_digits", profile_cfg.get("prefix_mst_digits", 3)),
            "prefix_name_words": payload.get("prefix_name_words", profile_cfg.get("prefix_name_words", 2)),
            "prefix_name_chars": payload.get("prefix_name_chars", profile_cfg.get("prefix_name_chars", 1)),
            "prefix_missing_mst_strategy": normalize_missing_mst_prefix_strategy(payload.get("prefix_missing_mst_strategy", profile_cfg.get("prefix_missing_mst_strategy", "all_name_words"))),
            "prefix_strategy_values": normalize_prefix_strategy_values(keep_company_config_when_unloaded(payload, profile_cfg, "prefix_strategy_values", "prefix_strategy_values", {})),
            "processing_groups": keep_existing_when_empty(payload, profile_cfg, "processing_groups", []),
            "company_group_assignments": keep_company_config_when_unloaded(payload, profile_cfg, "company_group_assignments", "company_group_assignments", {}),
            "form_mapping_presets": keep_form_mapping_presets(payload, profile_cfg),
            "columns": payload.get("columns", {}),
            "word_rules": keep_existing_when_empty(payload, profile_cfg, "word_rules", {}),
            "first_word_rules": keep_existing_when_empty(payload, profile_cfg, "first_word_rules", {}),
            "repeated_phrase_removals": normalize_phrase_list(keep_existing_when_empty(payload, profile_cfg, "repeated_phrase_removals", [])),
            "manual_code_overrides": keep_company_config_when_unloaded(payload, profile_cfg, "manual_code_overrides", "manual_code_overrides", {}),
            "product_code_replacements": normalize_product_code_replacements(keep_existing_when_empty(payload, profile_cfg, "product_code_replacements", {})),
            "product_review_merges": payload.get("product_review_merges", profile_cfg.get("product_review_merges", [])),
            "vietmax_mua_vao_internal_merges": payload.get("vietmax_mua_vao_internal_merges", profile_cfg.get("vietmax_mua_vao_internal_merges", [])),
            "vietmax_ban_ra_sales_internal_merges": payload.get("vietmax_ban_ra_sales_internal_merges", profile_cfg.get("vietmax_ban_ra_sales_internal_merges", [])),
            "vietmax_ban_ra_purchase_match_rules": payload.get("vietmax_ban_ra_purchase_match_rules", profile_cfg.get("vietmax_ban_ra_purchase_match_rules", [])),
            "price_range_rules": merge_price_ranges(profile_cfg.get("price_range_rules"), payload.get("price_range_rules", {})) if profile == "cao_thanh" else profile_cfg.get("price_range_rules", {}),
            "price_adjust_all_percent": float(payload.get("price_adjust_all_percent", profile_cfg.get("price_adjust_all_percent", 0)) or 0) if profile == "cao_thanh" else profile_cfg.get("price_adjust_all_percent", 0),
            "inventory_pairs": normalize_inventory_pairs(keep_existing_when_empty(payload, profile_cfg, "inventory_pairs", [])),
            "use_default_inventory_pair": bool(payload.get("use_default_inventory_pair", profile_cfg.get("use_default_inventory_pair", False))),
            "default_inventory_pair_id": str(payload.get("default_inventory_pair_id", profile_cfg.get("default_inventory_pair_id", "")) or "").strip(),
            "inventory_pair_rules": normalize_inventory_pair_rules(keep_existing_when_empty(payload, profile_cfg, "inventory_pair_rules", [])),
            "inventory_allocation_config": keep_existing_when_empty(payload, profile_cfg, "inventory_allocation_config", {}),
        })
        if stable_signature(profile_cfg) != previous_profile_signature:
            invalidation_kinds = [f"processed:{profile}:{phase}"]
            if phase == VIETMAX_PHASE_PURCHASE:
                invalidation_kinds.append(f"processed:{profile}:{VIETMAX_PHASE_SALES}")
            WORKFLOW_SESSION_STORE.invalidate(kinds=invalidation_kinds)
        sync_vietmax_legacy_profile(cfg, profile, phase, profile_cfg)
        return save_config(cfg)

    @api.post("/api/vietmax/import-config/{phase}")
    def import_vietmax_config(phase: str, payload: dict) -> dict:
        storage_profile, imported_profile = extract_vietmax_import_profile(payload or {}, phase)
        cfg = load_config()
        cfg["selected_profile"] = VIETMAX_PROFILE
        normalized_phase = normalize_vietmax_phase(phase)
        _, scoped_profile = ensure_profile_scope(cfg, VIETMAX_PROFILE, normalized_phase)
        scoped_profile.update(normalize_profile_config(VIETMAX_PROFILE, imported_profile, include_scopes=False))

        snapshot = dict_value((payload or {}).get("saved_config_file_snapshot"))
        snapshot_columns = dict_value(snapshot.get("columns"))
        payload_columns = dict_value((payload or {}).get("columns"))
        profile_columns = dict_value(scoped_profile.get("columns"))
        if snapshot_columns or payload_columns or profile_columns:
            cfg.setdefault("columns", {}).update(snapshot_columns)
            cfg.setdefault("columns", {}).update(payload_columns)
            cfg.setdefault("columns", {}).update(profile_columns)

        saved = save_config(cfg)
        return {
            "status": "ok",
            "storage_profile": f"vietmax.scopes.{normalized_phase}",
            "legacy_source_profile": storage_profile,
            "imported_keys": sorted(imported_profile.keys()),
            "config": saved,
        }

    @api.post("/api/vietmax/process")
    def process_vietmax(payload: VietmaxProcessRequest):
        path = uploaded_path(payload.saved_name)
        process_payload = dict(payload.payload)
        process_payload["saved_name"] = payload.saved_name
        process_payload["original_name"] = payload.original_name
        operation_id = raw_text(payload.operation_id)
        started_at = time.perf_counter()
        update_progress_job(operation_id, 0, 1, "Đang chuẩn bị tạo file đã xử lý", status="running")
        def progress(done: int, total: int, label: str) -> None:
            update_progress_job(operation_id, done, total, label, status="running")
        try:
            processed_purchase_saved_name = raw_text(process_payload.get("vietmax_processed_purchase_saved_name"))
            if process_payload.get("vietmax_phase") == VIETMAX_PHASE_SALES and not process_payload.get("vietmax_ban_ra_purchase_match_rules"):
                profile = profile_key(process_payload.get("profile") or VIETMAX_PROFILE)
                profile_cfg = scoped_profile_config(load_config(), profile, process_payload.get("vietmax_phase"))
                saved_match_rules = profile_cfg.get("vietmax_ban_ra_purchase_match_rules") or []
                if saved_match_rules:
                    process_payload["vietmax_ban_ra_purchase_match_rules"] = saved_match_rules
            if processed_purchase_saved_name:
                process_payload["vietmax_processed_purchase_path"] = str(uploaded_path(processed_purchase_saved_name))
            phase = process_payload.get("vietmax_phase")
            cached_saved_name = ""
            if payload.cache_only and phase in {VIETMAX_PHASE_PURCHASE, VIETMAX_PHASE_SALES}:
                cached_saved_name = f"processed_{uuid.uuid4().hex}.xls"
                out = UPLOAD_DIR / cached_saved_name
            else:
                out = cache_only_output_path(payload.original_name) if payload.cache_only else resolve_output_path(payload.original_name, "")
            diagnostic_log(
                "vietmax/process start "
                f"operation={operation_id or '-'} phase={phase or '-'} "
                f"cache_only={payload.cache_only} input={path} output={out}"
            )
            process_workbook(path, out, process_payload, progress_callback=progress)
            diagnostic_log(
                "vietmax/process workbook_done "
                f"operation={operation_id or '-'} phase={phase or '-'} "
                f"elapsed={time.perf_counter() - started_at:.2f}s output={out}"
            )
            if not cached_saved_name:
                update_progress_job(operation_id, 1, 2, "Đang lưu cache file đã xử lý", status="running")
                cached_saved_name = cache_processed_workbook(out) if phase in {VIETMAX_PHASE_PURCHASE, VIETMAX_PHASE_SALES} else ""
            if payload.cache_only and out.parent != UPLOAD_DIR:
                try:
                    out.unlink(missing_ok=True)
                except OSError:
                    pass
        except Exception as exc:
            diagnostic_log(
                "vietmax/process error "
                f"operation={operation_id or '-'} elapsed={time.perf_counter() - started_at:.2f}s error={exc}"
            )
            update_progress_job(operation_id, 1, 1, str(exc), status="error")
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        update_progress_job(operation_id, 2, 2, "Đã tạo cache file đã xử lý", status="complete")
        diagnostic_log(
            "vietmax/process complete "
            f"operation={operation_id or '-'} elapsed={time.perf_counter() - started_at:.2f}s "
            f"cached={cached_saved_name or '-'}"
        )
        headers = {}
        if cached_saved_name:
            headers["X-Processed-Saved-Name"] = cached_saved_name
        if payload.cache_only:
            return JSONResponse(
                {
                    "processed_saved_name": cached_saved_name,
                    "filename": out.name,
                },
                headers=headers or None,
            )
        return FileResponse(out, filename=out.name, media_type=mimetypes.guess_type(out.name)[0] or "application/octet-stream", headers=headers or None)

    @api.post("/api/inventory-allocation/analyze-job")
    async def inventory_allocation_analyze_job(
        purchase_saved_name: str = Form(...),
        sales_saved_name: str = Form(...),
        sales_original_name: str = Form("inventory_allocation_result.xlsx"),
        mapping: str = Form("{}"),
        policy: str = Form("{}"),
        opening_file: UploadFile | None = File(None),
    ) -> dict:
        try:
            purchase_path = uploaded_path(purchase_saved_name)
            sales_path = uploaded_path(sales_saved_name)
            raw_mapping = json.loads(mapping or "{}")
            allocation_policy = json.loads(policy or "{}")
            opening_content = await opening_file.read() if opening_file and opening_file.filename else None
            if opening_content:
                opening_content = uploaded_workbook_content_for_openpyxl(opening_content, opening_file.filename)
            purchase_content = workbook_content_for_openpyxl(purchase_path)
            sales_content = workbook_content_for_openpyxl(sales_path)
            job_id = uuid.uuid4().hex
            update_inventory_analysis_job(job_id, status="queued", progress=0, done=0, total=0, label="Đã nhận file. Đang xếp hàng phân bổ tồn kho...")
            worker = threading.Thread(
                target=run_inventory_analysis_job,
                args=(
                    job_id,
                    purchase_content,
                    sales_content,
                    opening_content,
                    raw_mapping,
                    allocation_policy,
                    sales_original_name or sales_path.name,
                ),
                daemon=True,
            )
            worker.start()
            return {"analysis_job_id": job_id}
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @api.get("/api/inventory-allocation/analyze-job/{job_id}")
    def inventory_allocation_job_status(job_id: str) -> dict:
        if not re_fullmatch_hex(job_id):
            raise HTTPException(status_code=404, detail="Mã xử lý không hợp lệ.")
        job = get_inventory_analysis_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Không tìm thấy tiến trình phân bổ tồn kho.")
        return json_safe(job)

    @api.get("/api/inventory-allocation/download/{job_id}")
    def inventory_allocation_download(job_id: str) -> Response:
        if not re_fullmatch_hex(job_id):
            raise HTTPException(status_code=404, detail="Mã kết quả không hợp lệ.")
        matching = list(INVENTORY_OUTPUT_DIR.glob(f"{job_id}_*.xlsx"))
        if not matching:
            raise HTTPException(status_code=404, detail="Không tìm thấy file kết quả phân bổ tồn kho.")
        path = matching[0]
        filename = safe_download_name(path.name[len(job_id) + 1:])
        try:
            stream = openpyxl_workbook_to_xls_stream(path)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return Response(
            stream.getvalue(),
            media_type=XLS_MEDIA_TYPE,
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    if REACT_DIST_DIR.exists():
        api.mount("/assets", StaticFiles(directory=REACT_DIST_DIR / "assets"), name="assets")

        @api.get("/{path:path}")
        def serve_react(path: str) -> FileResponse:
            target = REACT_DIST_DIR / path
            if path and target.exists() and target.is_file():
                return FileResponse(target)
            return FileResponse(REACT_DIST_DIR / "index.html")

    return api


def vietmax_profile_config(phase: str) -> dict:
    cfg = load_config()
    return dict(scoped_profile_config(cfg, VIETMAX_PROFILE, phase))


def workbook_summary(df, original_name: str, saved_name: str) -> dict:
    columns = []
    for idx in range(df.shape[1]):
        letter = index_to_excel_col(idx)
        samples = []
        for row in range(min(6, len(df))):
            value = df.iat[row, idx]
            if not pd.isna(value) and str(value).strip():
                samples.append(str(value).strip())
        label = letter + ((" - " + " | ".join(samples[:2])[:45]) if samples else "")
        columns.append({"letter": letter, "label": label})
    return {
        "original_name": original_name,
        "saved_name": saved_name,
        "columns": columns,
        "preview": preview_data(df),
        "invoice_statuses": invoice_status_options(df, DEFAULT_INVOICE_STATUS_COL, DEFAULT_INVOICE_STATUS_SKIP_VALUES),
    }


def public_license_status(config: dict) -> dict:
    return build_public_license_status(config, license_allows_profile, license_has_local_activation)

def inspect_vietmax_processed_file(path: Path, phase: str) -> dict:
    _, df = read_workbook(path)
    if df.empty:
        raise ValueError("File đã xử lý không có dữ liệu.")
    normalized_phase = VIETMAX_PHASE_SALES if str(phase or "").strip().casefold() == VIETMAX_PHASE_SALES else VIETMAX_PHASE_PURCHASE
    columns = {
        VIETMAX_PHASE_PURCHASE: {"company": "F", "mst": "G", "product": "M", "qty": "O", "code": "L"},
        VIETMAX_PHASE_SALES: {"company": "I", "mst": "J", "product": "M", "qty": "O", "code": "L"},
    }[normalized_phase]
    company_index = excel_col_to_index(columns["company"])
    mst_index = excel_col_to_index(columns["mst"])
    product_index = excel_col_to_index(columns["product"])
    code_index = excel_col_to_index(columns["code"])
    qty_index = excel_col_to_index(columns["qty"])
    required = [company_index, mst_index, product_index, code_index]
    if df.shape[1] <= max(required):
        raise ValueError("File đã xử lý không đủ cột Vietmax mặc định để đọc thống kê.")
    if df.shape[1] <= qty_index:
        qty_index = None
    product_header_labels = {"ten hang", "ten hang hoa", "hang hoa", "ten vat tu"}
    code_header_labels = {"ma vt", "ma vat tu", "ma hang", "ma hang hoa"}
    company_keys = set()
    processed_company_keys = set()
    total_rows = 0
    processed_rows = 0
    for row_index in range(len(df)):
        product = raw_text(cell(df, row_index, product_index))
        if not product:
            continue
        product_label = normalized_header_label(product)
        code_label = normalized_header_label(cell(df, row_index, code_index))
        if product_label in product_header_labels or code_label in code_header_labels:
            continue
        if qty_index is not None and not should_process_qty(cell(df, row_index, qty_index)):
            continue
        company = raw_text(cell(df, row_index, company_index))
        mst = raw_text(cell(df, row_index, mst_index))
        if not company and not mst:
            continue
        company_key = (mst.casefold().strip(), company.casefold().strip())
        company_keys.add(company_key)
        total_rows += 1
        code = raw_text(cell(df, row_index, code_index))
        code_valid = bool(code and code not in {"0", "0.0"} and normalized_header_label(code) not in code_header_labels)
        if code_valid:
            processed_rows += 1
            processed_company_keys.add(company_key)
    return {
        "phase": normalized_phase,
        "company_count": len(company_keys),
        "processed_company_count": len(processed_company_keys),
        "product_row_count": total_rows,
        "processed_product_row_count": processed_rows,
        "company_col": columns["company"],
        "mst_col": columns["mst"],
        "product_col": columns["product"],
        "code_col": columns["code"],
    }


def validate_workflow_process_payload(path: Path, payload: dict, phase: str):
    pairs = normalize_inventory_pairs(payload.get("inventory_pairs") or [])
    active_companies = [raw_text(value) for value in (payload.get("process_mst") or []) if raw_text(value)]
    if active_companies and not pairs:
        raise WorkflowFailure(
            "INVENTORY_PAIR_REQUIRED",
            "Chưa có cặp Mã kho / TK vật tư cho nhóm vật tư. Hãy quay lại cấu hình nâng cao và thêm ít nhất một cặp trước khi tạo file.",
            stage="11" if phase == VIETMAX_PHASE_SALES else "5",
            field="inventory_pairs",
            details={"active_company_count": len(active_companies)},
            retryable=True,
        )
    incomplete_pairs = [
        pair for pair in pairs
        if not raw_text(pair.get("ma_kho")) or not raw_text(pair.get("tk_vat_tu"))
    ]
    if incomplete_pairs:
        raise WorkflowFailure(
            "INVENTORY_PAIR_INCOMPLETE",
            "Có cặp phân kho chưa nhập đủ Mã kho và TK vật tư.",
            stage="11" if phase == VIETMAX_PHASE_SALES else "5",
            field="inventory_pairs",
            details={"pair_ids": [pair.get("id") for pair in incomplete_pairs]},
            retryable=True,
        )
    try:
        _, df = read_workbook(path)
    except Exception as exc:
        raise WorkflowFailure(
            "SOURCE_WORKBOOK_UNREADABLE",
            str(exc),
            stage="11" if phase == VIETMAX_PHASE_SALES else "5",
            field="source_file",
            retryable=True,
        ) from exc
    required = {
        "company_col": payload.get("company_col"),
        "mst_col": payload.get("mst_col"),
        "product_col": payload.get("product_col"),
        "qty_col": payload.get("qty_col"),
    }
    invalid = []
    for field, column in required.items():
        try:
            index = excel_col_to_index(raw_text(column))
        except ValueError:
            invalid.append({"field": field, "column": raw_text(column), "reason": "invalid"})
            continue
        if index >= df.shape[1]:
            invalid.append({"field": field, "column": raw_text(column), "reason": "outside_sheet"})
    if invalid:
        raise WorkflowFailure(
            "SOURCE_COLUMN_INVALID",
            "Cấu hình cột không khớp với file nguồn. Hãy quay lại stage chọn cột và kiểm tra các cột được báo lỗi.",
            stage="7" if phase == VIETMAX_PHASE_SALES else "2",
            field="columns",
            details={"invalid_columns": invalid, "sheet_column_count": int(df.shape[1])},
            retryable=True,
        )


def validate_form_mapping_columns(forms):
    invalid = []
    for form in forms or []:
        if not isinstance(form, dict) or form.get("enabled") is False:
            continue
        output_columns = form.get("output_columns")
        valid_columns = [
            column for column in (output_columns or [])
            if isinstance(column, dict) and raw_text(column.get("letter"))
        ]
        if not valid_columns:
            invalid.append({
                "form_id": raw_text(form.get("id")),
                "label": raw_text(form.get("label")),
            })
    if invalid:
        raise WorkflowFailure(
            "FORM_OUTPUT_COLUMNS_MISSING",
            "Một hoặc nhiều form mapping chưa có danh sách cột output. Hãy mở Form mapping và đọc cột từ file mẫu hoặc khôi phục form mặc định.",
            stage="0.5",
            field="form_mapping_presets",
            details={"forms": invalid},
            retryable=True,
        )

def run_workflow_process_job(path, original_name, process_payload, processor, kind, signature, progress):
    phase = normalize_vietmax_phase(process_payload.get("vietmax_phase"))
    profile = profile_key(process_payload.get("profile") or VIETMAX_PROFILE)
    validate_workflow_process_payload(path, process_payload, phase)
    progress(1, 5, "Đã kiểm tra file nguồn và cấu hình")
    processed_purchase_saved_name = raw_text(process_payload.get("vietmax_processed_purchase_saved_name"))
    if phase == VIETMAX_PHASE_SALES and processed_purchase_saved_name:
        try:
            process_payload["vietmax_processed_purchase_path"] = str(uploaded_path(processed_purchase_saved_name))
        except HTTPException as exc:
            raise WorkflowFailure(
                "PURCHASE_CACHE_MISSING",
                "Không tìm thấy cache mua vào đã xử lý. Hãy quay lại stage 5 và tạo lại file mua vào.",
                stage="5",
                field="processed_purchase",
                retryable=True,
            ) from exc
    saved_name = f"processed_{uuid.uuid4().hex}.xls"
    out = UPLOAD_DIR / saved_name
    started_at = time.perf_counter()
    try:
        processed_df = process_workbook(path, out, process_payload, progress_callback=progress)
        if processor == "generic" and process_payload.get("export_form_mappings"):
            forms = process_payload.get("form_mapping_presets")
            if not isinstance(forms, list):
                forms = scoped_profile_config(load_config(), profile, phase).get("form_mapping_presets") or []
            validate_form_mapping_columns(forms)
            stream = generic_form_mapping_export_workbook_stream(
                processed_df,
                form_mapping_presets=forms,
                company_group_assignments=process_payload.get("company_group_assignments") or {},
                profile=profile,
                progress_callback=progress,
            )
            out.write_bytes(stream.getvalue())
    except WorkflowFailure:
        out.unlink(missing_ok=True)
        raise
    except Exception as exc:
        out.unlink(missing_ok=True)
        raise WorkflowFailure(
            "WORKBOOK_PROCESSING_FAILED",
            str(exc),
            stage="11" if phase == VIETMAX_PHASE_SALES else "5",
            field="processing",
            details={"profile": profile, "phase": phase},
            retryable=True,
        ) from exc
    progress(4, 5, "Đang lưu cache file đã xử lý")
    output_name = Path(original_name or "output.xls").stem + "_fdi.xls"
    metadata = {
        "profile": profile,
        "phase": phase,
        "elapsed_seconds": round(time.perf_counter() - started_at, 3),
        "rows": int(len(processed_df.index)),
        "columns": int(len(processed_df.columns)),
    }
    artifact = WORKFLOW_SESSION_STORE.register_file(
        saved_name,
        kind=kind,
        original_name=output_name,
        signature=signature,
        metadata=metadata,
        supersede_kind=True,
    )
    progress(5, 5, "Đã tạo cache file đã xử lý")
    return {
        "artifact": artifact,
        "processed_saved_name": saved_name,
        "reused": False,
        "stats": metadata,
    }


def uploaded_path(saved_name: str) -> Path:
    path = UPLOAD_DIR / Path(saved_name).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Uploaded file was not found. Please upload again.")
    return path


def cache_only_output_path(original_name: str) -> Path:
    stem = Path(original_name or "processed").stem or "processed"
    safe_stem = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in stem).strip("._")[:80] or "processed"
    return OUTPUT_DIR / f"{safe_stem}_cache_{uuid.uuid4().hex}.xls"


def cache_processed_workbook(path: Path) -> str:
    saved_name = f"processed_{uuid.uuid4().hex}{path.suffix or '.xls'}"
    target = UPLOAD_DIR / saved_name
    shutil.copyfile(path, target)
    return saved_name


def cache_workbook_stream(stream, prefix="processed") -> str:
    saved_name = f"{prefix}_{uuid.uuid4().hex}.xls"
    target = UPLOAD_DIR / saved_name
    stream.seek(0)
    target.write_bytes(stream.getvalue())
    return saved_name


def dataframe_raw_rows(df: pd.DataFrame) -> list[list]:
    return [
        [cell(df, row_index, col_index) for col_index in range(df.shape[1])]
        for row_index in range(len(df))
    ]


def generic_profile_form_input_phase(profile):
    return VIETMAX_PHASE_SALES if profile_key(profile) == "cao_thanh" else VIETMAX_PHASE_PURCHASE


def form_scope_matches_available_phase(form, available_phase):
    scope = raw_text(form.get("scope") or form.get("input_phase") or "both")
    return scope in {"", "both", available_phase}


def single_phase_form_presets(forms, available_phase):
    source_forms = forms if isinstance(forms, list) else load_system_default_form_mappings(available_phase)
    result = []
    for form in source_forms:
        if not isinstance(form, dict) or form.get("enabled") is False:
            continue
        if not form_scope_matches_available_phase(form, available_phase):
            continue
        next_form = dict(form)
        next_form["scope"] = available_phase if raw_text(next_form.get("scope")) in {"", "both"} else raw_text(next_form.get("scope"))
        next_form["input_phase"] = available_phase if raw_text(next_form.get("input_phase")) in {"", "both"} else raw_text(next_form.get("input_phase"))
        mappings = []
        for rule in next_form.get("mappings") or []:
            if not isinstance(rule, dict):
                continue
            source_phase = raw_text(rule.get("source_phase") or available_phase)
            if source_phase not in {"", "both", available_phase}:
                continue
            next_rule = dict(rule)
            if source_phase in {"", "both"}:
                next_rule["source_phase"] = available_phase
            mappings.append(next_rule)
        next_form["mappings"] = mappings
        result.append(next_form)
    return result


def generic_form_mapping_export_workbook_stream(processed_df, form_mapping_presets, company_group_assignments, profile, progress_callback=None):
    def progress(done, total, label):
        if progress_callback:
            progress_callback(done, total, label)

    available_phase = generic_profile_form_input_phase(profile)
    forms = single_phase_form_presets(form_mapping_presets, available_phase)
    progress(1, 4, "Đang chuẩn bị sheet FDI đã xử lý")
    raw_sheets = [("FDI_da_xu_ly", dataframe_raw_rows(processed_df))]
    progress(2, 4, "Đang dựng các sheet form mapping")
    purchase_df = processed_df if available_phase == VIETMAX_PHASE_PURCHASE else None
    sales_df = processed_df if available_phase == VIETMAX_PHASE_SALES else None
    purchase_assignments = company_group_assignments if available_phase == VIETMAX_PHASE_PURCHASE else None
    sales_assignments = company_group_assignments if available_phase == VIETMAX_PHASE_SALES else None
    form_sheets = fast_import_sheet_rows(
        purchase_df,
        sales_df,
        progress_callback=None,
        form_mapping_presets=forms,
        purchase_company_group_assignments=purchase_assignments,
        sales_company_group_assignments=sales_assignments,
        require_inventory_columns=False,
        include_duplicate_report=False,
        use_default_forms_when_empty=False,
    )
    tabular_sheets = dict(form_sheets)
    if not tabular_sheets:
        raise ValueError("Chưa có form mapping nào đang hoạt động để xuất file.")
    progress(4, 4, "Đang đóng gói workbook form mapping")
    return mixed_xls_workbook(raw_sheets=raw_sheets, tabular_sheets=tabular_sheets)


def cleanup_progress_jobs(now=None):
    now = now or time.time()
    expired = [key for key, value in PROGRESS_JOBS.items() if now - float(value.get("updated_at", now)) > PROGRESS_TTL_SECONDS]
    for key in expired:
        PROGRESS_JOBS.pop(key, None)


def update_progress_job(operation_id: str, done: int, total: int, label: str, status: str = "running") -> None:
    if not operation_id:
        return
    safe_total = max(1, int(total or 1))
    safe_done = max(0, min(int(done or 0), safe_total))
    percent = int(round((safe_done / safe_total) * 100))
    now = time.time()
    with PROGRESS_LOCK:
        cleanup_progress_jobs(now)
        PROGRESS_JOBS[operation_id] = {
            "operation_id": operation_id,
            "status": status,
            "done": safe_done,
            "total": safe_total,
            "percent": percent,
            "label": raw_text(label),
            "updated_at": now,
        }


def get_progress_job(operation_id: str) -> dict:
    with PROGRESS_LOCK:
        cleanup_progress_jobs()
        job = dict(PROGRESS_JOBS.get(operation_id) or {})
    if job:
        return {key: value for key, value in job.items() if key != "updated_at"}
    return {"operation_id": operation_id, "status": "missing", "done": 0, "total": 1, "percent": 0, "label": "Chưa có tiến trình."}


def json_safe(value):
    return json.loads(json.dumps(value, default=raw_text, ensure_ascii=False))


def review_code_from_config(profile: str, product: str, provided_code: str, word_rules: dict, first_word_rules=None, repeated_phrase_removals=None, provided_code_is_user_config=False) -> tuple[str, bool]:
    provided_code = sanitize_product_code(provided_code)
    repeated_phrase_removals = repeated_phrase_removals or []
    if provided_code:
        return provided_code, bool(provided_code_is_user_config)
    generated = sanitize_product_code(make_product_part(profile, product, word_rules or {}, first_word_rules or {}, repeated_phrase_removals))
    default_generated = sanitize_product_code(make_product_part(profile, product, {}, {}, []))
    return generated, bool(generated and generated != default_generated)


def enrich_vietmax_review_products(products, side: str, payload: VietmaxReviewRequest):
    profile = "vietmax_ban_ra" if side == "sales" else "vietmax_mua_vao"
    word_rules = payload.word_rules or {}
    repeated = payload.repeated_phrase_removals or []
    enriched = []
    for item in products or []:
        row = dict(item)
        product_field = "sales_product" if side == "sales" else "purchase_product"
        code_field = "sales_code" if side == "sales" else "purchase_code"
        company_field = "sales_company" if side == "sales" else "purchase_company"
        mst_field = "sales_mst" if side == "sales" else "purchase_mst"
        product = raw_text(row.get(product_field))
        review_words = remove_repeated_phrases(code_words(product), repeated)
        if word_rules:
            review_words = apply_word_rules_to_words(review_words, word_rules)
        review_product = " ".join(review_words) if review_words else product
        provided_code = raw_text(row.get(code_field) or row.get("code"))
        provided_code_is_user_config = bool(row.get("product_key") or row.get("phase") or row.get("company_index") is not None or row.get("product_index") is not None)
        code, code_from_config = review_code_from_config(profile, product, provided_code, word_rules, repeated_phrase_removals=repeated, provided_code_is_user_config=provided_code_is_user_config)
        row[f"original_{product_field}"] = product
        row[product_field] = review_product
        row[code_field] = code
        row[f"{code_field}_from_user_config"] = code_from_config
        row["code"] = code
        row["code_from_user_config"] = code_from_config
        row["product_key"] = product_key(raw_text(row.get(mst_field)), product)
        if company_field in row:
            row.setdefault("company", row.get(company_field))
        if mst_field in row:
            row.setdefault("mst", row.get(mst_field))
        enriched.append(row)
    return enriched


def enrich_generic_review_products(products, profile: str, payload: GenericReviewRequest):
    word_rules = payload.word_rules or {}
    first_word_rules = payload.first_word_rules or {}
    repeated = payload.repeated_phrase_removals or []
    enriched = []
    for item in products or []:
        row = dict(item)
        product = raw_text(row.get("purchase_product") or row.get("sales_product") or row.get("product") or row.get("name"))
        company = raw_text(row.get("purchase_company") or row.get("sales_company") or row.get("company"))
        mst = raw_text(row.get("purchase_mst") or row.get("sales_mst") or row.get("mst"))
        review_words = remove_repeated_phrases(code_words(product), repeated)
        if word_rules:
            review_words = apply_word_rules_to_words(review_words, word_rules)
        review_product = " ".join(review_words) if review_words else product
        provided_code = raw_text(row.get("purchase_code") or row.get("code"))
        provided_code_is_user_config = bool(row.get("product_key") or row.get("phase") or row.get("company_index") is not None or row.get("product_index") is not None)
        code, code_from_config = review_code_from_config(profile, product, provided_code, word_rules, first_word_rules, repeated, provided_code_is_user_config)
        row["original_purchase_product"] = product
        row["purchase_product"] = review_product
        row["purchase_code"] = code
        row["purchase_code_from_user_config"] = code_from_config
        row["code"] = code
        row["code_from_user_config"] = code_from_config
        row["product_key"] = product_key(mst, product)
        row.setdefault("purchase_company", company)
        row.setdefault("purchase_mst", mst)
        row.setdefault("purchase_company_key", raw_text(row.get("company_key")) or vietmax_company_identity_key(company, mst))
        row.setdefault("company", company)
        row.setdefault("mst", mst)
        enriched.append(row)
    return enriched


def merge_exact_and_fuzzy_matches(exact_matches, fuzzy_matches):
    matches = []
    seen = set()
    for match in list(exact_matches) + list(fuzzy_matches):
        key = (raw_text(match.get("sales_product")).casefold(), raw_text(match.get("sales_company")).casefold(), raw_text(match.get("sales_mst")).casefold())
        if key in seen:
            continue
        seen.add(key)
        matches.append(match)
    return matches


def safe_download_name(value: str) -> str:
    name = Path(value or "export.xls").name
    return name if name.lower().endswith(".xls") else f"{Path(name).stem}.xls"


def re_fullmatch_hex(value: str) -> bool:
    return bool(re.fullmatch(r"[a-f0-9]{32}", value or ""))


def find_free_port(start: int = 8765) -> int:
    for port in range(start, start + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("No free local port found for ProductCodeFormatter.")


def open_browser_when_ready(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    for _ in range(80):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.25):
                webbrowser.open(url)
                return
        except OSError:
            time.sleep(0.25)
    webbrowser.open(url)


app = create_app()


def main() -> None:
    port = int(os.environ.get("PRODUCT_CODE_FORMATTER_PORT") or find_free_port())
    threading.Thread(target=open_browser_when_ready, args=(port,), daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")


if __name__ == "__main__":
    main()
