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
    VIETMAX_PHASE_PURCHASE,
    VIETMAX_PHASE_SALES,
    MAX_CODE_LENGTH,
    VIETMAX_PROFILE,
    analyze,
    activate_keygen_license,
    build_vietmax_ban_ra_purchase_matches,
    build_vietmax_khh_exact_purchase_matches,
    apply_word_rules_to_words,
    code_words,
    effective_processing_profile,
    empty_profile_config,
    index_to_excel_col,
    invoice_status_options,
    license_allows_profile,
    license_has_local_activation,
    load_config,
    make_excel_workbook,
    make_product_part,
    preview_data,
    process_workbook,
    product_key,
    profile_key,
    raw_text,
    read_workbook,
    remove_repeated_phrases,
    resolve_output_path,
    save_config,
    validate_vietmax_processed_purchase_workbook,
    vietmax_ban_ra_sales_products_from_workbook,
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


class VietmaxProductPreviewRequest(BaseModel):
    phase: str = VIETMAX_PHASE_PURCHASE
    products: list[str]
    word_rules: dict[str, str] | None = None
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


class VietmaxExportMatchesRequest(BaseModel):
    matches: list[dict]
    filename: str = "vietmax_khop_mua_ban.xlsx"


class VietmaxProcessRequest(BaseModel):
    saved_name: str
    original_name: str = "output.xlsx"
    payload: dict


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
    )

    @api.get("/api/health")
    def health() -> dict[str, str]:
        return {"status": "ok", "profile": "vietmax"}

    @api.get("/api/config")
    def get_config() -> dict:
        return load_config()

    @api.post("/api/config")
    def set_config(payload: dict) -> dict:
        return save_config(payload)

    @api.post("/api/config/profile/{profile}")
    def import_profile_config(profile: str, payload: dict) -> dict:
        current = load_config()
        incoming_profiles = payload.get("profiles") if isinstance(payload.get("profiles"), dict) else {}
        source_profile = profile if profile in incoming_profiles else str(payload.get("selected_profile") or profile)
        current["selected_profile"] = profile
        if isinstance(payload.get("columns"), dict):
            current.setdefault("columns", {}).update(payload.get("columns") or {})
        if source_profile in incoming_profiles:
            current.setdefault("profiles", {})[profile] = incoming_profiles[source_profile]
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
    def check(payload: VietmaxAnalyzeRequest) -> dict:
        return vietmax_analyze(payload)

    @api.post("/api/vietmax/review")
    def vietmax_review(payload: VietmaxReviewRequest) -> dict:
        path = uploaded_path(payload.saved_name)
        try:
            if payload.phase == VIETMAX_PHASE_SALES:
                products = vietmax_ban_ra_sales_products_from_workbook(
                    path,
                    payload.product_col,
                    payload.qty_col,
                    payload.invoice_status_col,
                    payload.invoice_status_skip_values,
                    comparison_scope=payload.comparison_scope,
                    price_col=payload.price_col,
                )
                rows = vietmax_product_review_rows(enrich_vietmax_review_products(products, "sales", payload), "sales_product", "sales_unit", comparison_scope=payload.comparison_scope)
            else:
                products = payload.products if payload.products is not None else vietmax_purchase_products_from_workbook(
                    path,
                    price_col=payload.price_col,
                    comparison_scope=payload.comparison_scope,
                    require_existing_code=payload.require_existing_code,
                )
                rows = vietmax_product_review_rows(enrich_vietmax_review_products(products, "purchase", payload), "purchase_product", "purchase_unit", comparison_scope=payload.comparison_scope)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
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
    @api.post("/api/vietmax/sales-match")
    def vietmax_sales_match(payload: VietmaxSalesMatchRequest) -> dict:
        sales_path = uploaded_path(payload.sales_saved_name)
        purchase_path = uploaded_path(payload.purchase_saved_name)
        try:
            sales_products = vietmax_ban_ra_sales_products_from_workbook(
                sales_path,
                payload.product_col,
                payload.qty_col,
                payload.invoice_status_col,
                payload.invoice_status_skip_values,
                comparison_scope=payload.comparison_scope,
                price_col=payload.sales_price_col,
                company_col=payload.sales_company_col,
                mst_col=payload.sales_mst_col,
            )
            purchase_products = vietmax_purchase_products_from_workbook(
                purchase_path,
                price_col=payload.purchase_price_col,
                comparison_scope=payload.comparison_scope,
                require_existing_code=payload.require_existing_purchase_code,
            )
            exact_matches = build_vietmax_khh_exact_purchase_matches(sales_products, purchase_products, payload.comparison_scope)
            fuzzy_matches = build_vietmax_ban_ra_purchase_matches(sales_products, purchase_products, comparison_scope=payload.comparison_scope)
            matches = merge_exact_and_fuzzy_matches(exact_matches, fuzzy_matches)
            match_rules = vietmax_profile_config(VIETMAX_PHASE_SALES).get("vietmax_ban_ra_purchase_match_rules") or []
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
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
            "output_path": payload.get("output_path", ""),
            "columns": payload.get("columns", {}),
            "word_rules": payload.get("word_rules", {}),
            "first_word_rules": payload.get("first_word_rules", {}),
            "repeated_phrase_removals": payload.get("repeated_phrase_removals", []),
            "manual_code_overrides": payload.get("manual_code_overrides", {}),
            "vietmax_ban_ra_purchase_match_rules": payload.get("vietmax_ban_ra_purchase_match_rules", profile_cfg.get("vietmax_ban_ra_purchase_match_rules", [])),
            "inventory_pairs": payload.get("inventory_pairs", []),
            "use_default_inventory_pair": bool(payload.get("use_default_inventory_pair")),
            "default_inventory_pair_id": str(payload.get("default_inventory_pair_id") or "").strip(),
            "inventory_pair_rules": payload.get("inventory_pair_rules", []),
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
            update_inventory_analysis_job(job_id, status="queued", progress=0, label="Đã nhận file. Đang xếp hàng phân bổ tồn kho...")
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
        "vietmax_allowed": activated and (
            license_allows_profile("vietmax", allowed_profiles)
            or license_allows_profile("vietmax_mua_vao", allowed_profiles)
            or license_allows_profile("vietmax_ban_ra", allowed_profiles)
        ),
        "server_url": str(license_cfg.get("server_url") or ""),
        "account_id": str(license_cfg.get("account_id") or ""),
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
