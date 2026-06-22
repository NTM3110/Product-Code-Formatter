import json
import mimetypes
import os
import re
import socket
import threading
import time
import uuid
import webbrowser
from pathlib import Path

import pandas as pd
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app import (
    DEFAULT_INVOICE_STATUS_COL,
    DEFAULT_INVOICE_STATUS_SKIP_VALUES,
    ICON_PATH,
    OUTPUT_DIR,
    UPLOAD_DIR,
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
    normalize_config,
    normalize_inventory_pair_rules,
    normalize_inventory_pairs,
    normalize_phrase_list,
    normalized_header_label,
    preview_data,
    process_zip_stream,
    process_workbook,
    product_key,
    profile_key,
    raw_text,
    read_workbook,
    remove_repeated_phrases,
    resolve_output_path,
    save_config,
    should_process_qty,
    up_ban_ra_output_path,
    create_up_ban_ra_workbook,
    validate_vietmax_processed_purchase_workbook,
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

APP_DIR = Path(__file__).resolve().parent
REACT_DIST_DIR = APP_DIR / "react_frontend" / "dist"
PROGRESS_LOCK = threading.Lock()
PROGRESS_JOBS: dict[str, dict] = {}
PROGRESS_TTL_SECONDS = 900


PREFIX_STRATEGIES = ("last_2_words", "last_3_mst", "2_words_mst")
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


def review_merge_key(product, company="", mst="", company_key="", comparison_scope=VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES):
    product_part = vietmax_ban_ra_match_key(product)
    if not product_part:
        return ""
    if comparison_scope == VIETMAX_COMPARISON_SCOPE_SAME_COMPANY:
        identity = raw_text(company_key) or vietmax_company_identity_key(company, mst)
        return f"{identity}|||{product_part}" if identity else ""
    return product_part


def apply_saved_review_choices(rows, saved_merges, comparison_scope):
    if not isinstance(saved_merges, list) or not saved_merges:
        return rows
    scope = comparison_scope if comparison_scope == VIETMAX_COMPARISON_SCOPE_SAME_COMPANY else VIETMAX_COMPARISON_SCOPE_ALL_COMPANIES
    forward = {}
    reverse = {}
    for item in saved_merges:
        if not isinstance(item, dict) or item.get("confirmed") is False:
            continue
        left_key = review_merge_key(item.get("product"), item.get("company"), item.get("mst"), item.get("company_key"), scope)
        right_key = review_merge_key(item.get("similar_product"), item.get("similar_company"), item.get("similar_mst"), item.get("similar_company_key"), scope)
        if left_key and right_key:
            forward[(left_key, right_key)] = item
            reverse[(right_key, left_key)] = item
    if not forward and not reverse:
        return rows
    restored = []
    for row in rows:
        next_row = dict(row)
        left_key = review_merge_key(row.get("product"), row.get("company"), row.get("mst"), row.get("company_key"), scope)
        right_key = review_merge_key(row.get("similar_product"), row.get("similar_company"), row.get("similar_mst"), row.get("similar_company_key"), scope)
        if left_key and right_key and (left_key, right_key) in forward:
            saved = forward[(left_key, right_key)]
            next_row.update({"confirmed": True, "code_choice": "similar"})
            if raw_text(saved.get("split_code")):
                next_row["split_code"] = raw_text(saved.get("split_code"))
            if raw_text(saved.get("similar_split_code")):
                next_row["similar_split_code"] = raw_text(saved.get("similar_split_code"))
        elif left_key and right_key and (left_key, right_key) in reverse:
            saved = reverse[(left_key, right_key)]
            next_row.update({"confirmed": True, "code_choice": "current"})
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
    company_col: str = "F"
    mst_col: str = "G"
    address_col: str = "H"
    product_col: str = "M"
    qty_col: str = "O"
    price_col: str = ""
    invoice_status_col: str = DEFAULT_INVOICE_STATUS_COL
    invoice_status_skip_values: list[str] | None = None
    operation_id: str = ""


class VietmaxProductPreviewRequest(BaseModel):
    phase: str = VIETMAX_PHASE_PURCHASE
    products: list[str]
    word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None


class GenericProductPreviewRequest(BaseModel):
    profile: str = "son_phuong"
    products: list[str]
    word_rules: dict[str, str] | None = None
    first_word_rules: dict[str, str] | None = None
    repeated_phrase_removals: list[str] | None = None


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


class VietmaxExportMatchesRequest(BaseModel):
    matches: list[dict]
    filename: str = "vietmax_khop_mua_ban.xlsx"


class VietmaxProcessRequest(BaseModel):
    saved_name: str
    original_name: str = "output.xlsx"
    payload: dict
    operation_id: str = ""


class LicenseActivationRequest(BaseModel):
    server_url: str
    account_id: str
    license_key: str


def create_app() -> FastAPI:
    api = FastAPI(title="ProductCodeFormatter API", version="0.1.0")
    api.add_middleware(
        CORSMiddleware,
        allow_origins=["http://127.0.0.1:5173", "http://localhost:5173"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
        expose_headers=["X-Processed-Saved-Name"],
    )

    @api.get("/api/health")
    def health() -> dict[str, str]:
        return {"status": "ok", "profile": "vietmax"}

    @api.get("/api/progress/{operation_id}")
    def operation_progress(operation_id: str) -> dict:
        return get_progress_job(operation_id)

    @api.get("/api/config")
    def get_config() -> dict:
        return load_config()

    @api.post("/api/config")
    def set_config(payload: dict) -> dict:
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
    async def upload_excel(file: UploadFile = File(...)) -> dict:
        original = file.filename or ""
        ext = Path(original).suffix.lower()
        if ext not in {".xlsx", ".xlsm"}:
            raise HTTPException(status_code=400, detail="Please upload .xlsx or .xlsm only.")
        saved_name = f"{uuid.uuid4().hex}{ext}"
        path = UPLOAD_DIR / saved_name
        path.write_bytes(await file.read())
        try:
            _, df = read_workbook(path)
        except Exception as exc:
            path.unlink(missing_ok=True)
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        return workbook_summary(df, original, saved_name)

    @api.get("/api/files/download/{saved_name}")
    def download_uploaded_or_cached_file(saved_name: str) -> FileResponse:
        path = uploaded_path(saved_name)
        return FileResponse(path, filename=path.name, media_type=mimetypes.guess_type(path.name)[0] or "application/octet-stream")

    @api.post("/api/mapping")
    async def mapping(file: UploadFile = File(...)) -> dict:
        return await upload_excel(file)

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
        profile_cfg = (cfg.get("profiles") or {}).get(profile) or empty_profile_config(profile)
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
            process = not bool(removed.get(mst))
            company["process"] = process
            company["pending_process"] = process
            company["committed_prefix"] = company.get("value") or ""
        result.update({
            "original_name": payload.original_name,
            "saved_name": payload.saved_name,
            "manual_code_overrides": profile_cfg.get("manual_code_overrides") or {},
            "word_rules": profile_cfg.get("word_rules") or {},
            "first_word_rules": profile_cfg.get("first_word_rules") or {},
            "repeated_phrase_removals": profile_cfg.get("repeated_phrase_removals") or [],
            "inventory_pairs": profile_cfg.get("inventory_pairs") or [],
            "use_default_inventory_pair": bool(profile_cfg.get("use_default_inventory_pair")),
            "default_inventory_pair_id": profile_cfg.get("default_inventory_pair_id") or "",
            "inventory_pair_rules": profile_cfg.get("inventory_pair_rules") or [],
            "include_company_prefix": profile_cfg.get("include_company_prefix") is not False,
            "prefix_strategy": profile_cfg.get("prefix_strategy") or "last_2_words",
            "prefix_mst_digits": profile_cfg.get("prefix_mst_digits") or 3,
            "prefix_strategy_values": profile_cfg.get("prefix_strategy_values") or {},
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
                rows = vietmax_product_review_rows(enrich_vietmax_review_products(products, "sales", payload), "sales_product", "sales_unit", comparison_scope=payload.comparison_scope, progress_callback=progress)
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
                process = not bool(removed.get(mst))
                company["process"] = process
                company["pending_process"] = process
            result["manual_code_overrides"] = profile_cfg.get("manual_code_overrides") or {}
            result["word_rules"] = profile_cfg.get("word_rules") or {}
            result["repeated_phrase_removals"] = profile_cfg.get("repeated_phrase_removals") or []
            result["inventory_pairs"] = profile_cfg.get("inventory_pairs") or []
            result["use_default_inventory_pair"] = bool(profile_cfg.get("use_default_inventory_pair"))
            result["default_inventory_pair_id"] = profile_cfg.get("default_inventory_pair_id") or ""
            result["inventory_pair_rules"] = profile_cfg.get("inventory_pair_rules") or []
            result["sales_match_rules"] = profile_cfg.get("vietmax_ban_ra_purchase_match_rules") or []
            result["vietmax_mua_vao_internal_merges"] = profile_cfg.get("vietmax_mua_vao_internal_merges") or []
            result["vietmax_ban_ra_sales_internal_merges"] = profile_cfg.get("vietmax_ban_ra_sales_internal_merges") or []
            result["include_company_prefix"] = profile_cfg.get("include_company_prefix") is not False
            result["prefix_strategy"] = profile_cfg.get("prefix_strategy") or "last_2_words"
            result["prefix_mst_digits"] = profile_cfg.get("prefix_mst_digits") or 3
            result["prefix_strategy_values"] = profile_cfg.get("prefix_strategy_values") or {}
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
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    @api.post("/api/process")
    def process_generic(payload: dict) -> Response:
        saved_name = str(payload.get("saved_name", ""))
        original_name = str(payload.get("original_name") or "output.xlsx")
        path = uploaded_path(saved_name)
        process_payload = dict(payload)
        try:
            out = resolve_output_path(original_name, process_payload.get("output_path", ""))
            processed_df = process_workbook(path, out, process_payload)
            companion_stream = create_up_ban_ra_workbook(processed_df)
            companion_out = up_ban_ra_output_path(out)
            companion_out.write_bytes(companion_stream.getvalue())
            cfg = load_config()
            profile = profile_key(process_payload.get("profile", cfg.get("selected_profile", "son_phuong")))
            storage_profile = effective_processing_profile(profile, process_payload.get("vietmax_phase")) if profile == VIETMAX_PROFILE else profile
            uses_price_rules = storage_profile == "cao_thanh"
            cfg["selected_profile"] = profile
            cfg.setdefault("profiles", {}).setdefault(storage_profile, empty_profile_config(storage_profile))
            old_profile = cfg["profiles"].get(storage_profile) or empty_profile_config(storage_profile)
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
            cfg["profiles"][storage_profile].update({
                "prefixes": process_payload.get("prefixes", {}),
                "selected_products": process_payload.get("skipped_products_map", {}),
                "removed_companies": process_payload.get("removed_companies", old_profile.get("removed_companies", {})),
                "word_rules": keep_existing_when_empty(process_payload, old_profile, "word_rules", {}),
                "first_word_rules": keep_existing_when_empty(process_payload, old_profile, "first_word_rules", {}),
                "repeated_phrase_removals": normalize_phrase_list(keep_existing_when_empty(process_payload, old_profile, "repeated_phrase_removals", [])),
                "price_group_rules": process_payload.get("price_group_rules", {}) if uses_price_rules else {},
                "price_range_rules": merged_price_ranges,
                "price_adjust_all_percent": float(process_payload.get("price_adjust_all_percent") or 0) if uses_price_rules else 0,
                "manual_code_overrides": process_payload.get("manual_code_overrides", {}),
                "vietmax_mua_vao_internal_merges": process_payload.get("vietmax_mua_vao_internal_merges", old_profile.get("vietmax_mua_vao_internal_merges", [])),
                "vietmax_ban_ra_sales_internal_merges": process_payload.get("vietmax_ban_ra_sales_internal_merges", old_profile.get("vietmax_ban_ra_sales_internal_merges", [])),
                "inventory_pairs": normalize_inventory_pairs(keep_existing_when_empty(process_payload, old_profile, "inventory_pairs", [])),
                "use_default_inventory_pair": bool(process_payload.get("use_default_inventory_pair", old_profile.get("use_default_inventory_pair", False))),
                "default_inventory_pair_id": str(process_payload.get("default_inventory_pair_id", old_profile.get("default_inventory_pair_id", "")) or "").strip(),
                "inventory_pair_rules": normalize_inventory_pair_rules(keep_existing_when_empty(process_payload, old_profile, "inventory_pair_rules", [])),
                "include_company_prefix": process_payload.get("include_company_prefix") is not False,
                "prefix_strategy": process_payload.get("prefix_strategy", old_profile.get("prefix_strategy", "last_2_words")),
                "prefix_mst_digits": process_payload.get("prefix_mst_digits", old_profile.get("prefix_mst_digits", 3)),
                "prefix_strategy_values": normalize_prefix_strategy_values(process_payload.get("prefix_strategy_values", old_profile.get("prefix_strategy_values", {}))),
                "output_path": process_payload.get("output_path", ""),
                "columns": columns,
            })
            save_config(cfg)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        zip_name = f"{Path(original_name).stem}_ket_qua_xu_ly.zip"
        return Response(
            process_zip_stream(out, companion_out).getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
        )

    @api.post("/api/vietmax/product-preview")
    def vietmax_product_preview(payload: VietmaxProductPreviewRequest) -> dict:
        profile = "vietmax_ban_ra" if payload.phase == VIETMAX_PHASE_SALES else "vietmax_mua_vao"
        word_rules = payload.word_rules or {}
        repeated = payload.repeated_phrase_removals or []
        codes = {}
        for product in payload.products:
            name = raw_text(product)
            if name:
                codes[name] = make_product_part(profile, name, word_rules, repeated_phrase_removals=repeated)[:MAX_CODE_LENGTH]
        return {"codes": codes}

    @api.post("/api/product-preview")
    def generic_product_preview(payload: GenericProductPreviewRequest) -> dict:
        profile = profile_key(payload.profile)
        word_rules = payload.word_rules or {}
        first_word_rules = payload.first_word_rules or {}
        repeated = payload.repeated_phrase_removals or []
        codes = {}
        for product in payload.products:
            name = raw_text(product)
            if name:
                codes[name] = make_product_part(profile, name, word_rules, first_word_rules, repeated)[:MAX_CODE_LENGTH]
        return {"codes": codes}

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
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{safe_download_name(payload.filename)}"'},
        )


    @api.post("/api/vietmax/save-config")
    def save_vietmax_config(payload: dict) -> dict:
        cfg = load_config()
        profile = profile_key(payload.get("profile") or VIETMAX_PROFILE)
        storage_profile = effective_processing_profile(profile, payload.get("vietmax_phase")) if profile == VIETMAX_PROFILE else profile
        cfg["selected_profile"] = profile
        cfg.setdefault("columns", {}).update(payload.get("columns") or {})
        cfg.setdefault("profiles", {}).setdefault(storage_profile, empty_profile_config(storage_profile))
        profile_cfg = cfg["profiles"].get(storage_profile) or empty_profile_config(storage_profile)
        profile_cfg.update({
            "prefixes": payload.get("prefixes", {}),
            "selected_products": payload.get("skipped_products_map", {}),
            "removed_companies": payload.get("removed_companies", {}),
            "include_company_prefix": payload.get("include_company_prefix") is not False,
            "prefix_strategy": payload.get("prefix_strategy", profile_cfg.get("prefix_strategy", "last_2_words")),
            "prefix_mst_digits": payload.get("prefix_mst_digits", profile_cfg.get("prefix_mst_digits", 3)),
            "prefix_strategy_values": normalize_prefix_strategy_values(payload.get("prefix_strategy_values", profile_cfg.get("prefix_strategy_values", {}))),
            "output_path": payload.get("output_path", ""),
            "columns": payload.get("columns", {}),
            "word_rules": keep_existing_when_empty(payload, profile_cfg, "word_rules", {}),
            "first_word_rules": keep_existing_when_empty(payload, profile_cfg, "first_word_rules", {}),
            "repeated_phrase_removals": normalize_phrase_list(keep_existing_when_empty(payload, profile_cfg, "repeated_phrase_removals", [])),
            "manual_code_overrides": payload.get("manual_code_overrides", {}),
            "vietmax_mua_vao_internal_merges": payload.get("vietmax_mua_vao_internal_merges", profile_cfg.get("vietmax_mua_vao_internal_merges", [])),
            "vietmax_ban_ra_sales_internal_merges": payload.get("vietmax_ban_ra_sales_internal_merges", profile_cfg.get("vietmax_ban_ra_sales_internal_merges", [])),
            "vietmax_ban_ra_purchase_match_rules": payload.get("vietmax_ban_ra_purchase_match_rules", profile_cfg.get("vietmax_ban_ra_purchase_match_rules", [])),
            "inventory_pairs": normalize_inventory_pairs(keep_existing_when_empty(payload, profile_cfg, "inventory_pairs", [])),
            "use_default_inventory_pair": bool(payload.get("use_default_inventory_pair", profile_cfg.get("use_default_inventory_pair", False))),
            "default_inventory_pair_id": str(payload.get("default_inventory_pair_id", profile_cfg.get("default_inventory_pair_id", "")) or "").strip(),
            "inventory_pair_rules": normalize_inventory_pair_rules(keep_existing_when_empty(payload, profile_cfg, "inventory_pair_rules", [])),
        })
        cfg["profiles"][storage_profile] = profile_cfg
        return save_config(cfg)

    @api.post("/api/vietmax/process")
    def process_vietmax(payload: VietmaxProcessRequest) -> FileResponse:
        path = uploaded_path(payload.saved_name)
        process_payload = dict(payload.payload)
        process_payload["saved_name"] = payload.saved_name
        process_payload["original_name"] = payload.original_name
        try:
            processed_purchase_saved_name = raw_text(process_payload.get("vietmax_processed_purchase_saved_name"))
            if process_payload.get("vietmax_phase") == VIETMAX_PHASE_SALES and not process_payload.get("vietmax_ban_ra_purchase_match_rules"):
                profile = profile_key(process_payload.get("profile") or VIETMAX_PROFILE)
                storage_profile = effective_processing_profile(profile, process_payload.get("vietmax_phase")) if profile == VIETMAX_PROFILE else profile
                profile_cfg = (load_config().get("profiles", {}).get(storage_profile, {}) or {})
                saved_match_rules = profile_cfg.get("vietmax_ban_ra_purchase_match_rules") or []
                if saved_match_rules:
                    process_payload["vietmax_ban_ra_purchase_match_rules"] = saved_match_rules
            if processed_purchase_saved_name:
                process_payload["vietmax_processed_purchase_path"] = str(uploaded_path(processed_purchase_saved_name))
            out = resolve_output_path(payload.original_name, process_payload.get("output_path", ""))
            process_workbook(path, out, process_payload)
            cached_saved_name = cache_processed_workbook(out) if process_payload.get("vietmax_phase") in {VIETMAX_PHASE_PURCHASE, VIETMAX_PHASE_SALES} else ""
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
        headers = {"X-Processed-Saved-Name": cached_saved_name} if cached_saved_name else None
        return FileResponse(out, filename=out.name, media_type=mimetypes.guess_type(out.name)[0] or "application/octet-stream", headers=headers)

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
            job_id = uuid.uuid4().hex
            update_inventory_analysis_job(job_id, status="queued", progress=0, done=0, total=0, label="Đã nhận file. Đang xếp hàng phân bổ tồn kho...")
            worker = threading.Thread(
                target=run_inventory_analysis_job,
                args=(
                    job_id,
                    purchase_path.read_bytes(),
                    sales_path.read_bytes(),
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
    def inventory_allocation_download(job_id: str) -> FileResponse:
        if not re_fullmatch_hex(job_id):
            raise HTTPException(status_code=404, detail="Mã kết quả không hợp lệ.")
        matching = list(INVENTORY_OUTPUT_DIR.glob(f"{job_id}_*.xlsx"))
        if not matching:
            raise HTTPException(status_code=404, detail="Không tìm thấy file kết quả phân bổ tồn kho.")
        path = matching[0]
        filename = path.name[len(job_id) + 1:]
        return FileResponse(path, filename=filename, media_type=mimetypes.guess_type(path.name)[0] or "application/octet-stream")

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
    storage_profile = effective_processing_profile(VIETMAX_PROFILE, phase)
    return dict((cfg.get("profiles") or {}).get(storage_profile) or empty_profile_config(storage_profile))


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
    license_cfg = config.get("license") or {}
    allowed_profiles = license_cfg.get("allowed_profiles") or []
    activated = license_has_local_activation(license_cfg)
    return {
        "activated": activated,
        "status": str(license_cfg.get("status") or ("Kích hoạt thành công" if activated else "Chưa kích hoạt")),
        "allowed_profiles": allowed_profiles,
        "allowed_companies": license_cfg.get("allowed_companies") or [],
        "supported_profiles": license_cfg.get("supported_profiles") or [],
        "product_code": str(license_cfg.get("product_code") or ""),
        "application": str(license_cfg.get("application") or ""),
        "vietmax_allowed": activated and (
            license_allows_profile("vietmax", allowed_profiles)
            or license_allows_profile("vietmax_mua_vao", allowed_profiles)
            or license_allows_profile("vietmax_ban_ra", allowed_profiles)
        ),
        "server_url": str(license_cfg.get("server_url") or ""),
        "account_id": str(license_cfg.get("account_id") or ""),
    }


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


def uploaded_path(saved_name: str) -> Path:
    path = UPLOAD_DIR / Path(saved_name).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="Uploaded file was not found. Please upload again.")
    return path


def cache_processed_workbook(path: Path) -> str:
    saved_name = f"processed_{uuid.uuid4().hex}{path.suffix or '.xlsx'}"
    target = UPLOAD_DIR / saved_name
    target.write_bytes(path.read_bytes())
    return saved_name


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
        code = make_product_part(profile, product, word_rules, repeated_phrase_removals=repeated)[:MAX_CODE_LENGTH]
        row[f"original_{product_field}"] = product
        row[product_field] = review_product
        row[code_field] = code
        row["code"] = code
        row["product_key"] = product_key(raw_text(row.get(mst_field)), product)
        if company_field in row:
            row.setdefault("company", row.get(company_field))
        if mst_field in row:
            row.setdefault("mst", row.get(mst_field))
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
    name = Path(value or "export.xlsx").name
    return name if name.lower().endswith(".xlsx") else f"{Path(name).stem}.xlsx"


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
