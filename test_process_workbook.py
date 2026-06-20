import unittest
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook, load_workbook

import app
from app import build_vietmax_ban_ra_purchase_matches, create_up_ban_ra_workbook, default_config, license_allows_company, license_allows_profile, license_has_local_activation, make_product_part, normalize_config, process_workbook, profile_key, resolve_output_path, suggest_prefix, up_ban_ra_output_path, vietmax_ban_ra_sales_products_from_workbook, vietmax_product_review_rows, vietmax_purchase_match_export_rows, vietmax_purchase_products_from_workbook


class ProcessWorkbookTests(unittest.TestCase):
    def test_vietmax_profiles_are_available_by_default(self):
        config = default_config()
        self.assertEqual(profile_key("vietmax"), "vietmax")
        self.assertIn("vietmax", config["profiles"])
        self.assertIn("vietmax_mua_vao", config["profiles"])
        self.assertIn("vietmax_ban_ra", config["profiles"])

    def test_unified_vietmax_profile_config_is_canonical(self):
        config = normalize_config({
            "selected_profile": "vietmax",
            "profiles": {
                "vietmax": {
                    "manual_code_overrides": {"MST|||Hang A": "VM.A"},
                    "inventory_pairs": [{"id": "one", "ma_kho": "KHO", "tk_vat_tu": "152"}],
                }
            },
        })

        self.assertEqual(config["selected_profile"], "vietmax")
        self.assertEqual(config["profiles"]["vietmax"]["manual_code_overrides"], {"MST|||Hang A": "VM.A"})
        self.assertEqual(config["profiles"]["vietmax"]["inventory_pairs"][0]["id"], "one")

    def test_license_config_round_trips_allowed_profiles(self):
        config = normalize_config({
            "license": {
                "server_url": "http://license-server.local:3000/",
                "account_id": "acct",
                "license_key": "KEY",
                "activated": True,
                "allowed_profiles": "son_phuong; Cao Thành\nvietmax_ban_ra",
            }
        })

        self.assertEqual(config["license"]["server_url"], "http://license-server.local:3000/")
        self.assertTrue(config["license"]["activated"])
        self.assertEqual(config["license"]["allowed_profiles"], ["son_phuong", "Cao Thành", "vietmax_ban_ra"])

    def test_license_allowed_profiles_match_key_or_label(self):
        self.assertTrue(license_allows_profile("cao_thanh", ["Cao Thành"]))
        self.assertTrue(license_allows_profile("vietmax_ban_ra", ["vietmax_ban_ra"]))
        self.assertFalse(license_allows_profile("quang_thinh", ["son_phuong", "cao_thanh"]))

    def test_license_local_activation_trusts_same_machine_without_server(self):
        original_fingerprint = app.local_machine_fingerprint
        try:
            app.local_machine_fingerprint = lambda: "TEST-FINGERPRINT"
            self.assertTrue(license_has_local_activation({"activated": True, "machine_fingerprint": "TEST-FINGERPRINT"}))
            self.assertFalse(license_has_local_activation({"activated": True, "machine_fingerprint": "OTHER-FINGERPRINT"}))
            self.assertFalse(license_has_local_activation({"activated": False, "machine_fingerprint": "TEST-FINGERPRINT"}))
        finally:
            app.local_machine_fingerprint = original_fingerprint

    def test_keygen_metadata_accepts_camel_case_profile_keys(self):
        response = {
            "data": {
                "attributes": {
                    "metadata": {
                        "allowedProfiles": ["cao_thanh", "vietmax_ban_ra"]
                    }
                }
            }
        }
        metadata = app.keygen_license_metadata(response)
        self.assertEqual(app.extract_allowed_profiles(metadata), ["cao_thanh", "vietmax_ban_ra"])

    def test_keygen_company_metadata_does_not_limit_profiles(self):
        metadata = {"allowed_companies": ["BAC MAI"]}

        self.assertEqual(app.extract_allowed_companies(metadata), ["BAC MAI"])
        self.assertEqual(app.extract_allowed_profiles(metadata), [])

    def test_license_allowed_companies_match_mst_or_name(self):
        company = {"mst": "0303062904", "company": "Công ty A", "all_names": ["Tên khác"]}

        self.assertTrue(license_allows_company(company, ["0303062904"]))
        self.assertTrue(license_allows_company({"mst": "0303062904.0", "company": "X"}, ["03030629040"]))
        self.assertTrue(license_allows_company(company, ["cong ty a"]))
        self.assertFalse(license_allows_company(company, ["030857608"]))

    def test_suggest_prefix_skips_viet_nam_and_provinces(self):
        self.assertEqual(suggest_prefix("Công ty TNHH In Bao Bì Việt Nam"), "BB")
        self.assertEqual(suggest_prefix("Công ty TNHH In Bao Bì Hà Nội"), "BB")

    def test_keygen_url_allows_lan_http(self):
        self.assertTrue(app.keygen_url("http://license-server.local:3000", "acct", "licenses/actions/validate-key").startswith("http://license-server.local:3000/"))
        self.assertTrue(app.keygen_url("http://192.168.1.228:3000", "acct", "machines").startswith("http://192.168.1.228:3000/"))
        self.assertTrue(app.keygen_url("http://localhost:3000", "acct", "machines").startswith("http://localhost:3000/"))
        with self.assertRaises(ValueError):
            app.keygen_url("http://example.com", "acct", "machines")

    def test_keygen_request_marks_lan_http_as_forwarded_https(self):
        captured = []

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return b"{}"

        def fake_urlopen(request_obj, timeout=10):
            captured.append(request_obj)
            return FakeResponse()

        original_urlopen = app.urlopen
        try:
            app.urlopen = fake_urlopen
            app.keygen_request("POST", "http://license-server.local:3000/v1/accounts/acct/licenses/actions/validate-key", {})
        finally:
            app.urlopen = original_urlopen

        self.assertEqual(captured[0].get_header("X-forwarded-proto"), "https")

    def test_keygen_activation_creates_machine_then_revalidates(self):
        calls = []

        def fake_request(method, url, payload=None, license_key=None, timeout=10):
            calls.append((method, url, payload, license_key))
            if url.endswith("licenses/actions/validate-key") and len(calls) == 1:
                return {
                    "meta": {"valid": False, "code": "NO_MACHINE", "detail": "machine required"},
                    "data": {"id": "lic-1", "attributes": {"metadata": {"allowed_profiles": ["cao_thanh"]}}},
                }
            if url.endswith("machines"):
                self.assertEqual(payload["data"]["relationships"]["license"]["data"], {"type": "licenses", "id": "lic-1"})
                return {"data": {"id": "machine-1"}}
            return {
                "meta": {"valid": True, "code": "VALID"},
                "data": {"id": "lic-1", "attributes": {"status": "ACTIVE", "metadata": {"allowed_profiles": ["cao_thanh"]}}},
            }

        original_request = app.keygen_request
        original_fingerprint = app.local_machine_fingerprint
        try:
            app.keygen_request = fake_request
            app.local_machine_fingerprint = lambda: "TEST-FINGERPRINT"
            license_cfg = app.activate_keygen_license("https://license-server.local", "acct", "KEY")
        finally:
            app.keygen_request = original_request
            app.local_machine_fingerprint = original_fingerprint

        self.assertEqual(license_cfg["machine_id"], "machine-1")
        self.assertEqual(license_cfg["allowed_profiles"], ["cao_thanh"])
        self.assertEqual([call[0] for call in calls], ["POST", "POST", "POST"])

    def test_vietmax_product_code_matches_workbook_examples(self):
        examples = {
            "Bản nhôm CTP (D-L) 400x530x0.3 (50 tấm) 2 lớp": "BNCTP400x530x0.3",
            "Giấy Cacbon CF yellow 56/650*860_TL (R500)": "GCCFYELLOW56650*860_TL(R500)",
            "Màng tự dính SYNWK-F1840N": "MANGTUDINH",
            "Mực in - TK Mark V T Cyan (VN) - 2Kg (xanh)": "MUCINXANH",
            "Véc ni phủ bóng bề mặt OP Varnish (new)": "VECNI",
        }
        for product_name, expected in examples.items():
            with self.subTest(product_name=product_name):
                self.assertEqual(make_product_part("vietmax_mua_vao", product_name, {}), expected)

    def test_vietmax_product_code_falls_back_for_unknown_products(self):
        self.assertEqual(
            make_product_part("vietmax_mua_vao", "Pin SYNWK-F1840N dung lượng", {}),
            "PINSYNWK-F1840NDULU",
        )

    def test_vietmax_ban_ra_fallback_uses_first_word_then_initials(self):
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Pin SYNWK-F1840N 500ML dung lượng", {}),
            "PINSYNWK-F1840N500MLDL",
        )
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Giấy An Hòa PP 92/70", {}),
            "GAHPP92/70",
        )
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Giấy in offset 100gsm", {}),
            "GIAYINOFFSET100GSM",
        )
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Màng tự dính SYNWK-F1840N", {}),
            "MANGTUDINH",
        )
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Supercalifragilisticexpialidocious alpha beta gamma delta epsilon", {}),
            "SUPERCALIFRAGILISTICEXPIALIDOCIOUSABGDE"[:50],
        )

    def test_vietmax_product_code_uses_all_words_for_note_patterns(self):
        examples = {
            "Giấy Couche": "GIAYCOUCHE",
            "Giấy offset": "GIAYOFFSET",
            "Giấy Cacbon": "GIAYCACBON",
            "Giấy in BB58/92gsm": "GIAYINBB58/92GSM",
            "Giấy ivory ningbo 230gsm khổ 62cm": "GIVORYNINGBO230GSM62CM",
        }
        for product_name, expected in examples.items():
            with self.subTest(product_name=product_name):
                self.assertEqual(make_product_part("vietmax_mua_vao", product_name, {}), expected)

    def test_vietmax_product_code_uses_workbook_prefix_rules_for_variants(self):
        examples = {
            "Giấy An Hoà 95/60gsm/620x860": "GANHOA9560GSM620x860",
            "Giấy An Hoà 95/70gsm/620x860": "GANHOA9570GSM620x860",
            "Giấy Couche 250gsm (79x47) cm": "GCOUCHE250GSM79X47",
            "Giấy Cacbon CB white 56/610*860_TL (R500)": "GCCBWHITE56610*860_TL(R500)",
            "Giấy offset 100gsm khổ 62x86cm": "GOFFSET100GSM62X86CM",
            "Giấy Couche định lượng 80gsm, khổ 620x860mm": "GCOUCHE80GSM620x860MM",
            "Giấy Ivory định lượng 250gsm, khổ 790 mm": "GIVORY250GSM790MM",
            "Màng tự dính BLWK-Z0585MW": "MANGTUDINH",
            "Tấm bản in bằng nhôm CTP SR 530x400": "TAMBANNHOMCTPSR530x400",
        }
        for product_name, expected in examples.items():
            with self.subTest(product_name=product_name):
                self.assertEqual(make_product_part("vietmax_mua_vao", product_name, {}), expected)

    def test_vietmax_product_code_keeps_hl_prefix_marker(self):
        self.assertEqual(
            make_product_part("vietmax_mua_vao", "Tấm bản in bằng nhôm CTP HL 1030x800", {}),
            "TAMBANNHOMCTPHL1030x800",
        )

    def test_vietmax_mang_tu_dinh_keeps_customer_hardcode(self):
        self.assertEqual(make_product_part("vietmax_mua_vao", "Màng tự dính BLWK-Z0585MW", {}), "MANGTUDINH")
        self.assertEqual(make_product_part("vietmax_mua_vao", "Màng tự dính SYNWK-F1840N", {}), "MANGTUDINH")
        self.assertEqual(make_product_part("vietmax_ban_ra", "Màng tự dính SYNWK-F1840N", {}), "MANGTUDINH")

    def test_vietmax_duplex_g_m2_uses_single_uppercase_gm2(self):
        self.assertEqual(make_product_part("vietmax_mua_vao", "Giấy Duplex ĐL 300 g/m2", {}), "GDL300GM2")
        self.assertEqual(make_product_part("vietmax_mua_vao", "Giấy Duplex ĐL 400 g/m2", {}), "GDL400GM2")

    def test_vietmax_product_code_honors_custom_word_rules(self):
        self.assertEqual(
            make_product_part("vietmax_mua_vao", "Mực pantone 2009 C", {"pantone": "PANTONE"}),
            "MUCPANTONE2009C",
        )

    def test_vietmax_product_code_normalizes_formatting_only_tokens(self):
        self.assertEqual(
            make_product_part("vietmax_mua_vao", "Giấy couche hikote 120gsm khổ62x86cm", {}),
            make_product_part("vietmax_mua_vao", "Giấy Couche Hikote 120gsm khổ 62x86cm", {}),
        )
        self.assertEqual(
            make_product_part("vietmax_mua_vao", "Mực pantone2009C", {}),
            make_product_part("vietmax_mua_vao", "Mực pantone 2009 C", {}),
        )

    def test_vietmax_ban_ra_purchase_match_accepts_near_one_character_difference(self):
        matches = build_vietmax_ban_ra_purchase_matches(
            ["Giấy Couche 300 gsm"],
            [{"purchase_product": "Giấy Couche 300 gms", "purchase_code": "GC300", "purchase_row": 5}],
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["purchase_code"], "GC300")
        self.assertGreaterEqual(matches[0]["score"], 0.92)

    def test_vietmax_sales_products_include_invoice_metadata(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        source = outputs / f"_test_vietmax_sales_{uuid4().hex}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 3, "Số HD")
            sheet.cell(2, 4, "Ngày có hàng bán ra")
            sheet.cell(2, 13, "Tên hàng")
            sheet.cell(2, 14, "ĐVT")
            sheet.cell(2, 15, "Số lượng")
            sheet.cell(3, 3, "HD-001")
            sheet.cell(3, 4, "06/06/2026")
            sheet.cell(3, 13, "Giấy Couche 300 gsm")
            sheet.cell(3, 14, "Tờ")
            sheet.cell(3, 15, 2)
            sheet.cell(3, 16, 12345)
            workbook.save(source)
            workbook.close()

            products = vietmax_ban_ra_sales_products_from_workbook(source, price_col="P")
            matches = build_vietmax_ban_ra_purchase_matches(
                products,
                [{"purchase_product": "Giấy Couche 300 gsm", "purchase_code": "GC300", "purchase_row": 5, "purchase_unit": "Tờ"}],
            )

            self.assertEqual(products, [{"sales_product": "Giấy Couche 300 gsm", "sales_unit": "Tờ", "sales_price": "12345", "invoice_no": "HD-001", "invoice_date": "06/06/2026"}])
            self.assertEqual(matches[0]["invoice_no"], "HD-001")
            self.assertEqual(matches[0]["invoice_date"], "06/06/2026")
            self.assertEqual(matches[0]["sales_unit"], "Tờ")
            self.assertEqual(matches[0]["sales_price"], "12345")
            self.assertFalse(matches[0]["unit_mismatch"])
        finally:
            source.unlink(missing_ok=True)

    def test_vietmax_purchase_match_prefers_duplicate_rows_with_prices(self):
        sales_products = [
            {"sales_product": "Giấy Couche", "sales_unit": "Tờ", "sales_price": "", "invoice_no": "HD-001", "invoice_date": "06/06/2026"},
            {"sales_product": "Giấy Couche", "sales_unit": "Tờ", "sales_price": "10500", "invoice_no": "HD-002", "invoice_date": "07/06/2026"},
        ]
        purchase_products = [
            {"purchase_product": "Giấy Couche", "purchase_code": "GIAYCOUCHE", "purchase_row": 5, "purchase_unit": "Kg", "purchase_price": ""},
            {"purchase_product": "Giấy Couche", "purchase_code": "GIAYCOUCHE", "purchase_row": 6, "purchase_unit": "Kg", "purchase_price": "21000"},
        ]

        matches = build_vietmax_ban_ra_purchase_matches(sales_products, purchase_products)

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["sales_price"], "10500")
        self.assertEqual(matches[0]["purchase_price"], "21000")
        self.assertEqual(matches[0]["invoice_no"], "HD-002")

    def test_vietmax_exact_khh_match_keeps_sales_and_purchase_prices(self):
        matches = app.build_vietmax_khh_exact_purchase_matches(
            [{"sales_product": "Giấy Couche", "sales_unit": "Kg", "sales_price": "21,000", "invoice_no": "HD-001"}],
            [{"purchase_product": "Giấy Couche", "purchase_code": "GC", "purchase_unit": "kg", "purchase_price": "19,500"}],
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["sales_price"], "21,000")
        self.assertEqual(matches[0]["purchase_price"], "19,500")

    def test_vietmax_purchase_match_export_rows_match_display_columns(self):
        rows = vietmax_purchase_match_export_rows([
            {
                "confirmed": False,
                "sales_product": "Giấy Couche 300 gsm",
                "sales_unit": "Tờ",
                "invoice_no": "HD-001",
                "invoice_date": "06/06/2026",
                "purchase_code": "GC300",
                "purchase_product": "Giấy Couche 300 gms",
                "purchase_unit": "Ram",
                "unit_mismatch": True,
                "unit_warning": "Khác đơn vị tính",
                "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                "conversion_formula": "1 ram = 500 tờ",
                "purchase_row": 5,
                "score": 0.925,
            }
        ])

        self.assertEqual(rows[0], ["Dùng", "Hàng bán ra", "ĐVT bán ra", "Số HD", "Ngày có hàng bán ra", "Mã VT mua vào", "Hàng mua vào", "ĐVT mua vào", "Cảnh báo", "Quy đổi", "Khác biệt", "Độ giống"])
        self.assertEqual(rows[1], ["Không", "Giấy Couche 300 gsm", "Tờ", "HD-001", "06/06/2026", "GC300", "Giấy Couche 300 gms", "Ram", "Khác đơn vị tính", "1 ram = 500 tờ", "Giấy Couche 300 gsm -> Giấy Couche 300 gms", "92.5%"])

    def test_vietmax_purchase_match_flags_unit_mismatch(self):
        matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_unit": "Tờ"}],
            [{"purchase_product": "Giấy Couche 300 gsm", "purchase_code": "GC300", "purchase_row": 5, "purchase_unit": "Ram"}],
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["sales_unit"], "Tờ")
        self.assertEqual(matches[0]["purchase_unit"], "Ram")
        self.assertTrue(matches[0]["unit_mismatch"])
        self.assertEqual(matches[0]["unit_warning"], "Khác đơn vị tính")
        self.assertEqual(matches[0]["conversion_formula"], "")

    def test_vietmax_product_review_rows_show_near_matches_not_exact_duplicates(self):
        rows = vietmax_product_review_rows(
            [
                {"purchase_product": "Giấy Couche 300 gsm", "purchase_unit": "Tờ", "invoice_no": "MV-001", "invoice_date": "01/06/2026"},
                {"purchase_product": "Giấy Couche 300 gms", "purchase_unit": "Ram", "invoice_no": "MV-002", "invoice_date": "02/06/2026"},
                {"purchase_product": "Giấy Couche 300 gsm", "purchase_unit": "Tờ", "invoice_no": "MV-003", "invoice_date": "03/06/2026"},
                {"purchase_product": "Mực in xanh", "purchase_unit": "Hộp"},
            ],
            "purchase_product",
            "purchase_unit",
        )

        self.assertEqual([row["product"] for row in rows], ["Giấy Couche 300 gms"])
        self.assertEqual(rows[0]["invoice_no"], "MV-002")
        self.assertEqual(rows[0]["invoice_date"], "02/06/2026")
        self.assertEqual(rows[0]["similar_product"], "Giấy Couche 300 gsm")
        self.assertEqual(rows[0]["similar_invoice_no"], "MV-001")
        self.assertEqual(rows[0]["similar_invoice_date"], "01/06/2026")
        self.assertNotIn("Mực in xanh", [row["product"] for row in rows])

    def test_vietmax_product_review_rows_ignores_formatting_only_differences(self):
        rows = vietmax_product_review_rows(
            [
                {"purchase_product": "Giấy couche hikote 120gsm khổ62x86cm", "purchase_unit": "Kg"},
                {"purchase_product": "Giấy Couche Hikote 120gsm khổ 62 x 86cm", "purchase_unit": "Kg"},
                {"purchase_product": "Mực pantone2009C", "purchase_unit": "kg"},
                {"purchase_product": "Mực pantone 2009 C", "purchase_unit": "kg"},
            ],
            "purchase_product",
            "purchase_unit",
        )

        self.assertEqual(rows, [])

    def test_vietmax_product_review_rows_groups_obvious_size_or_volume_differences(self):
        rows = vietmax_product_review_rows(
            [
                {"purchase_product": "Bản nhôm CTP BOCICA 560x670x0.3 (50 tấm)", "purchase_unit": "Hộp"},
                {"purchase_product": "Bản nhôm CTP BOCICA 600x730x0.3 (50 tấm)", "purchase_unit": "Hộp"},
                {"purchase_product": "In decal Lau sàn Fineday huong ly 3.6 lít", "purchase_unit": "Chai"},
                {"purchase_product": "In decal Lau sàn Fineday huong ly 1.2 lít", "purchase_unit": "Chai"},
                {"purchase_product": "Tem nhãn A5", "purchase_unit": "Tờ"},
                {"purchase_product": "Tem nhãn A4", "purchase_unit": "Tờ"},
            ],
            "purchase_product",
            "purchase_unit",
        )

        self.assertGreaterEqual(len(rows), 2)
        self.assertTrue(all(row["review_group"] == "dimension_diff" for row in rows))
        self.assertTrue(all(row["dimension_only"] is True for row in rows))

    def test_vietmax_product_review_rows_keeps_hl_merge_candidate_separate(self):
        rows = vietmax_product_review_rows(
            [
                {"purchase_product": "Tấm bản in bằng nhôm CTP 1030 x 800", "purchase_unit": "Tấm", "code": "TAMBANNHOMCTP1030x800"},
                {"purchase_product": "Tấm bản in bằng nhôm CTP HL 1030 x 800", "purchase_unit": "Tấm", "code": "TAMBANNHOMCTPHL1030x800"},
            ],
            "purchase_product",
            "purchase_unit",
        )

        self.assertEqual(len(rows), 1)
        self.assertIn("HL", rows[0]["similar_product"] + rows[0]["product"])
        self.assertNotEqual(rows[0].get("code"), rows[0].get("similar_code"))

    def test_vietmax_product_review_rows_forces_special_purchase_pairs_into_other(self):
        rows = vietmax_product_review_rows(
            [
                {"purchase_product": "Giấy An Hòa PP 92/70", "purchase_unit": "Tờ", "code": "GAHPP9270"},
                {"purchase_product": "Giấy An Hòa 92/70gms/790x1090", "purchase_unit": "Tờ", "code": "GAH9270GMS790x1090"},
                {"purchase_product": "Giấy in offset 100gsm", "purchase_unit": "Tờ", "code": "GIAYINOFFSET100GSM"},
                {"purchase_product": "Giấy In Offset", "purchase_unit": "Tờ", "code": "GIAYINOFFSET"},
            ],
            "purchase_product",
            "purchase_unit",
        )

        pairs = {(row["product"], row["similar_product"]): row for row in rows}
        self.assertIn(("Giấy An Hòa PP 92/70", "Giấy An Hòa 92/70gms/790 x 1090"), pairs)
        self.assertIn(("Giấy in offset 100gsm", "Giấy In Offset"), pairs)
        self.assertTrue(all(row["review_group"] == "other" for row in pairs.values()))
        self.assertTrue(all(row["dimension_only"] is False for row in pairs.values()))
        self.assertTrue(all(row.get("forced_review") is True for row in pairs.values()))

    def test_vietmax_product_review_rows_same_company_scope_filters_near_matches(self):
        rows = vietmax_product_review_rows(
            [
                {"sales_product": "Giấy Couche 300 gsm", "sales_company": "Cong ty A", "sales_mst": "111"},
                {"sales_product": "Giấy Couche 300 gms", "sales_company": "Cong ty B", "sales_mst": "222"},
                {"sales_product": "Giấy Couche 300 gms", "sales_company": "Cong ty A", "sales_mst": "111"},
            ],
            "sales_product",
            "sales_unit",
            comparison_scope=app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
        )

        self.assertEqual(len(rows), 1)
        self.assertTrue(all(row["company_key"] == "mst:111" for row in rows))

    def test_vietmax_purchase_internal_merge_updates_processed_product_name(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_purchase_merge_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_purchase_merge_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Purchase"
            sheet.cell(2, 6, "Ten nguoi ban")
            sheet.cell(2, 7, "MST")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 14, "DVT")
            sheet.cell(2, 15, "So luong")
            sheet.cell(3, 6, "Cong ty A")
            sheet.cell(3, 7, "111")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Giấy An Hòa PP 92/70")
            sheet.cell(3, 14, "Tờ")
            sheet.cell(3, 15, 1)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "F",
                "mst_col": "G",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax",
                "vietmax_phase": app.VIETMAX_PHASE_PURCHASE,
                "include_company_prefix": False,
                "all_mst": ["111"],
                "process_mst": ["111"],
                "mst_safe_id": ["111|||0"],
                "selected_products_0": ["Giấy An Hòa PP 92/70"],
                "vietmax_mua_vao_internal_merges": [
                    {
                        "product": "Giấy An Hòa PP 92/70",
                        "similar_product": "Giấy An Hòa 92/70gms/790x1090",
                        "similar_code": "GAH9270GMS790x1090",
                        "confirmed": True,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Purchase"]
                self.assertEqual(result_sheet.cell(3, 12).value, "GAH9270GMS790x1090")
                self.assertEqual(result_sheet.cell(3, 13).value, "Giấy An Hòa 92/70gms/790x1090")
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_internal_purchase_merge_affects_cross_file_matching(self):
        purchase_products = [{"purchase_product": "Giấy Couche 300 gms", "purchase_code": "GC300", "purchase_unit": "Ram"}]
        merges = app.normalize_vietmax_internal_merges([
            {"product": "Giấy Couche 300 gms", "similar_product": "Giấy Couche 300 gsm", "confirmed": True},
        ])
        effective_purchase_products = app.apply_vietmax_internal_merges_to_products(
            purchase_products,
            "purchase_product",
            "purchase_unit",
            merges,
        )
        matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_unit": "Ram"}],
            effective_purchase_products,
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["purchase_code"], "GC300")
        self.assertEqual(matches[0]["purchase_product"], "Giấy Couche 300 gsm")

    def test_vietmax_internal_purchase_merge_uses_representative_code(self):
        purchase_products = [{"purchase_product": "Giấy Couche 300 gms", "purchase_code": "OLD300", "purchase_unit": "Ram"}]
        merges = app.normalize_vietmax_internal_merges([
            {
                "product": "Giấy Couche 300 gms",
                "similar_product": "Giấy Couche 300 gsm",
                "similar_code": "COMMON300",
                "confirmed": True,
            },
        ])
        effective_purchase_products = app.apply_vietmax_internal_merges_to_products(
            purchase_products,
            "purchase_product",
            "purchase_unit",
            merges,
        )
        matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_unit": "Ram"}],
            effective_purchase_products,
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["purchase_code"], "COMMON300")
        self.assertEqual(matches[0]["purchase_product"], "Giấy Couche 300 gsm")

    def test_unchecked_vietmax_internal_purchase_merge_is_ignored(self):
        purchase_products = [{"purchase_product": "Tên khác hoàn toàn", "purchase_code": "WRONG", "purchase_unit": "Ram"}]
        merges = app.normalize_vietmax_internal_merges([
            {"product": "Tên khác hoàn toàn", "similar_product": "Sản phẩm mục tiêu 123", "confirmed": False},
        ])
        effective_purchase_products = app.apply_vietmax_internal_merges_to_products(
            purchase_products,
            "purchase_product",
            "purchase_unit",
            merges,
        )
        matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Sản phẩm mục tiêu 123", "sales_unit": "Ram"}],
            effective_purchase_products,
        )

        self.assertEqual(matches, [])

    def test_vietmax_ban_ra_purchase_match_focus_list_bypasses_threshold(self):
        matches = build_vietmax_ban_ra_purchase_matches(
            ["Cabon 1 liên kích thước A4"],
            [{"purchase_product": "Tên mua vào khác hoàn toàn", "purchase_code": "FOCUS-CODE", "purchase_row": 8}],
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["purchase_code"], "FOCUS-CODE")
        self.assertTrue(matches[0]["focus"])
        self.assertLess(matches[0]["score"], 0.92)

    def test_vietmax_purchase_match_default_scope_matches_across_companies(self):
        matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_company": "Cong ty A", "sales_mst": "111"}],
            [{"purchase_product": "Giấy Couche 300 gsm", "purchase_code": "GC300", "purchase_company": "Cong ty B", "purchase_mst": "222"}],
        )

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["purchase_code"], "GC300")

    def test_vietmax_purchase_match_same_company_scope_filters_candidates(self):
        different_company_matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_company": "Cong ty A", "sales_mst": "111"}],
            [{"purchase_product": "Giấy Couche 300 gsm", "purchase_code": "WRONG", "purchase_company": "Cong ty B", "purchase_mst": "222"}],
            comparison_scope=app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
        )
        same_mst_matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Giấy Couche 300 gsm", "sales_company": "Ten ban ra", "sales_mst": "111"}],
            [{"purchase_product": "Giấy Couche 300 gsm", "purchase_code": "SAME-MST", "purchase_company": "Ten mua vao", "purchase_mst": "111"}],
            comparison_scope=app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
        )
        same_name_matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Màng tự dính SYNWK-F1840N", "sales_company": "Công ty In Bao Bì A"}],
            [{"purchase_product": "Màng tự dính SYNWK-F1840N", "purchase_code": "SAME-NAME", "purchase_company": "Cong ty In Bao Bi A"}],
            comparison_scope=app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
        )
        focus_wrong_company_matches = build_vietmax_ban_ra_purchase_matches(
            [{"sales_product": "Cabon 1 liên kích thước A4", "sales_company": "Cong ty A", "sales_mst": "111"}],
            [{"purchase_product": "Tên mua vào khác hoàn toàn", "purchase_code": "FOCUS-WRONG", "purchase_company": "Cong ty B", "purchase_mst": "222"}],
            comparison_scope=app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
        )

        self.assertEqual(different_company_matches, [])
        self.assertEqual(same_mst_matches[0]["purchase_code"], "SAME-MST")
        self.assertEqual(same_name_matches[0]["purchase_code"], "SAME-NAME")
        self.assertEqual(focus_wrong_company_matches, [])

    def test_vietmax_ban_ra_purchase_match_skips_print_service_rows(self):
        matches = build_vietmax_ban_ra_purchase_matches(
            ["In tờ rơi A4 giấy coucher", "Công in giấy couche", "Gia công in giấy Couche", "Giấy in couche"],
            [{"purchase_product": "Giấy in Couche", "purchase_code": "GIAYINCOUCHE", "purchase_row": 5}],
        )

        self.assertEqual([item["sales_product"] for item in matches], ["Giấy in couche"])
        self.assertEqual(matches[0]["purchase_code"], "GIAYINCOUCHE")

    def test_vietmax_purchase_products_generate_code_when_purchase_code_is_zero(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        source = outputs / f"_test_vietmax_purchase_{uuid4().hex}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 3, "Số HD")
            sheet.cell(2, 4, "Ngày HD")
            sheet.cell(2, 12, "Mã VT")
            sheet.cell(2, 13, "Tên hàng hóa, dịch vụ")
            sheet.cell(2, 14, "ĐVT")
            sheet.cell(3, 3, "MV-001")
            sheet.cell(3, 4, "06/06/2026")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Giấy Couche")
            sheet.cell(3, 14, "Ram")
            sheet.cell(3, 16, 67890)
            workbook.save(source)
            workbook.close()

            products = vietmax_purchase_products_from_workbook(source)

            self.assertEqual(products, [{"purchase_product": "Giấy Couche", "purchase_code": "GIAYCOUCHE", "purchase_row": 3, "purchase_unit": "Ram", "purchase_price": "67890", "invoice_no": "MV-001", "invoice_date": "06/06/2026", "purchase_invoice_no": "MV-001", "purchase_invoice_date": "06/06/2026"}])
        finally:
            source.unlink(missing_ok=True)

    def test_vietmax_purchase_products_read_price_by_header(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        source = outputs / f"_test_vietmax_purchase_price_{uuid4().hex}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 3, "Số HD")
            sheet.cell(2, 4, "Ngày HD")
            sheet.cell(2, 12, "Mã VT")
            sheet.cell(2, 13, "Tên hàng hóa, dịch vụ")
            sheet.cell(2, 14, "ĐVT")
            sheet.cell(2, 18, "Ðon giá")
            sheet.cell(3, 3, "MV-002")
            sheet.cell(3, 4, "07/06/2026")
            sheet.cell(3, 12, "GC300")
            sheet.cell(3, 13, "Giấy Couche 300 gsm")
            sheet.cell(3, 14, "Ram")
            sheet.cell(3, 16, 11111)
            sheet.cell(3, 18, 67890)
            workbook.save(source)
            workbook.close()

            products = vietmax_purchase_products_from_workbook(source)

            self.assertEqual(products[0]["purchase_price"], "67890")
        finally:
            source.unlink(missing_ok=True)

    def test_vietmax_purchase_products_read_selected_price_column(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        source = outputs / f"_test_vietmax_purchase_selected_price_{uuid4().hex}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 3, "Số HD")
            sheet.cell(2, 4, "Ngày HD")
            sheet.cell(2, 12, "Mã VT")
            sheet.cell(2, 13, "Tên hàng hóa, dịch vụ")
            sheet.cell(2, 14, "ĐVT")
            sheet.cell(2, 16, "Ðon giá")
            sheet.cell(3, 3, "MV-003")
            sheet.cell(3, 4, "08/06/2026")
            sheet.cell(3, 12, "GC300")
            sheet.cell(3, 13, "Giấy Couche 300 gsm")
            sheet.cell(3, 14, "Ram")
            sheet.cell(3, 16, 11111)
            sheet.cell(3, 18, 67890)
            workbook.save(source)
            workbook.close()

            products = vietmax_purchase_products_from_workbook(source, price_col="R")

            self.assertEqual(products[0]["purchase_price"], "67890")
        finally:
            source.unlink(missing_ok=True)

    def test_removes_invoice_rows_that_are_not_assigned_a_processed_code(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_process_source_{run_id}.xlsx"
        result = outputs / f"_test_process_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 13, "Ma VT")
            sheet.cell(2, 14, "Ten hang")
            sheet.cell(2, 16, "So luong")
            sheet.cell(2, 18, "Don gia")
            sheet.cell(3, 3, "HD-BO-QUA")
            sheet.cell(3, 13, 0)
            sheet.cell(4, 3, "HD-XU-LY")
            sheet.cell(4, 6, "Cong ty A")
            sheet.cell(4, 7, "MST")
            sheet.cell(4, 13, 0)
            sheet.cell(4, 14, "Hang A")
            sheet.cell(4, 16, 2)
            sheet.cell(4, 18, 100)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "F",
                "mst_col": "G",
                "product_col": "N",
                "qty_col": "P",
                "price_col": "R",
                "output_col": "M",
                "profile": "cao_thanh",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Hang A"],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.max_row, 3)
                self.assertEqual(result_sheet.cell(3, 3).value, "HD-XU-LY")
                self.assertNotIn(result_sheet.cell(3, 13).value, (None, "", 0, "0"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)


    def inventory_base_payload(self, selected_products):
        return {
            "company_col": "F",
            "mst_col": "G",
            "product_col": "N",
            "qty_col": "P",
            "output_col": "M",
            "profile": "son_phuong",
            "include_company_prefix": False,
            "all_mst": ["MST"],
            "process_mst": ["MST"],
            "mst_safe_id": ["MST|||0"],
            "selected_products_0": selected_products,
        }

    def run_inventory_case(self, rows, payload_updates=None, existing_inventory_headers=False):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_inventory_source_{run_id}.xlsx"
        result = outputs / f"_test_inventory_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 13, "Ma VT")
            sheet.cell(2, 14, "Ten hang")
            sheet.cell(2, 15, "Don vi")
            sheet.cell(2, 16, "So luong")
            if existing_inventory_headers:
                sheet.cell(2, 17, "TK vat tu")
                sheet.cell(2, 18, "Ma kho")
                sheet.cell(3, 17, "OLD-TK")
                sheet.cell(3, 18, "OLD-KHO")
            for offset, row in enumerate(rows, start=3):
                sheet.cell(offset, 3, f"HD-{offset}")
                sheet.cell(offset, 6, "Cong ty A")
                sheet.cell(offset, 7, "MST")
                sheet.cell(offset, 13, 0)
                sheet.cell(offset, 14, row["product"])
                sheet.cell(offset, 15, row.get("unit", "Cai"))
                sheet.cell(offset, 16, row.get("qty", 1))
            workbook.save(source)
            workbook.close()

            payload = self.inventory_base_payload([row["product"] for row in rows])
            payload.update(payload_updates or {})
            process_workbook(source, result, payload)

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                values = []
                for row_index in range(3, 3 + len(rows)):
                    values.append((result_sheet.cell(row_index, 17).value, result_sheet.cell(row_index, 18).value))
                headers = (result_sheet.cell(2, 17).value, result_sheet.cell(2, 18).value)
                return headers, values, result_sheet.max_column
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_inventory_default_pair_updates_existing_columns(self):
        headers, values, max_column = self.run_inventory_case(
            [{"product": "Hang A"}],
            {
                "inventory_pairs": [{"id": "default", "ma_kho": "KHO-1", "tk_vat_tu": "1561"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
            },
            existing_inventory_headers=True,
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("1561", "KHO-1")])
        self.assertEqual(max_column, 18)

    def test_inventory_rule_priority_wins_over_broad_rule(self):
        headers, values, _ = self.run_inventory_case(
            [{"product": "Gia cong in va hoan thien"}, {"product": "In tui giay"}],
            {
                "inventory_pairs": [
                    {"id": "ktp", "ma_kho": "KTP", "tk_vat_tu": "1551"},
                    {"id": "kgci", "ma_kho": "KGCI", "tk_vat_tu": "1552"},
                ],
                "inventory_pair_rules": [
                    {"source_col": "N", "operator": "contains", "value": "In", "pair_id": "ktp", "enabled": True, "priority": 1},
                    {"source_col": "N", "operator": "contains", "value": "Gia cong in", "pair_id": "kgci", "enabled": True, "priority": 10},
                ],
            },
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("1552", "KGCI"), ("1551", "KTP")])

    def test_inventory_unit_column_equals_rule(self):
        _, values, _ = self.run_inventory_case(
            [{"product": "Hang Kg", "unit": " kg "}, {"product": "Hang Cai", "unit": "Cai"}],
            {
                "inventory_pairs": [
                    {"id": "kg", "ma_kho": "KHO-KG", "tk_vat_tu": "152-KG"},
                    {"id": "other", "ma_kho": "KHO-OTHER", "tk_vat_tu": "152-OTHER"},
                ],
                "inventory_pair_rules": [
                    {"source_col": "O", "operator": "equals", "value": "KG", "pair_id": "kg", "enabled": True},
                ],
            },
        )

        self.assertEqual(values, [("152-KG", "KHO-KG"), (None, None)])

    def test_inventory_appends_missing_headers(self):
        headers, values, max_column = self.run_inventory_case(
            [{"product": "Hang A"}],
            {
                "inventory_pairs": [{"id": "default", "ma_kho": "KHO-APPEND", "tk_vat_tu": "1562"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
            },
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("1562", "KHO-APPEND")])
        self.assertEqual(max_column, 18)

    def test_inventory_single_pair_auto_defaults_all_processed_rows(self):
        headers, values, _ = self.run_inventory_case(
            [{"product": "Hang A"}, {"product": "Hang B"}],
            {
                "inventory_pairs": [{"id": "only", "ma_kho": "KHO-ONLY", "tk_vat_tu": "156-ONLY"}],
                "use_default_inventory_pair": False,
                "default_inventory_pair_id": "",
            },
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("156-ONLY", "KHO-ONLY"), ("156-ONLY", "KHO-ONLY")])

    def test_inventory_normalizes_reversed_saved_pair_values(self):
        headers, values, _ = self.run_inventory_case(
            [{"product": "Hang A"}],
            {
                "inventory_pairs": [{"id": "reversed", "ma_kho": "152", "tk_vat_tu": "KVT"}],
                "use_default_inventory_pair": False,
                "default_inventory_pair_id": "",
            },
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("152", "KVT")])


    def test_vietmax_ban_ra_purchase_match_converts_mismatched_unit_rows(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_convert_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_convert_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 14, "DVT")
            sheet.cell(2, 15, "So luong")
            sheet.cell(2, 16, "Don gia")
            sheet.cell(2, 17, "TK vat tu")
            sheet.cell(2, 18, "Ma kho")
            for row, quantity in [(3, 500), (4, 250)]:
                sheet.cell(row, 9, "Cong ty A")
                sheet.cell(row, 10, "MST")
                sheet.cell(row, 12, 0)
                sheet.cell(row, 13, "Giay in A4")
                sheet.cell(row, 14, "To")
                sheet.cell(row, 15, quantity)
                sheet.cell(row, 16, 100)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Giay in A4", "Giay my thuat"],
                "inventory_pairs": [{"id": "default", "ma_kho": "KVT", "tk_vat_tu": "156"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "vietmax_ban_ra_purchase_matches": [
                    {
                        "sales_product": "Giay in A4",
                        "purchase_code": "GIAY-RAM",
                        "sales_unit": "To",
                        "purchase_unit": "Ram",
                        "unit_mismatch": True,
                        "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                        "conversion_formula": "1 ram = 500 to",
                        "confirmed": True,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Giay in A4", {}))
                self.assertEqual(result_sheet.cell(4, 12).value, make_product_part("vietmax_ban_ra", "Giay in A4", {}))
                self.assertEqual(result_sheet.cell(3, 14).value, "Ram")
                self.assertEqual(result_sheet.cell(4, 14).value, "Ram")
                self.assertEqual(result_sheet.cell(3, 15).value, 1)
                self.assertEqual(result_sheet.cell(4, 15).value, 0.5)
                self.assertEqual(result_sheet.cell(3, 16).value, 50000)
                self.assertEqual(result_sheet.cell(4, 16).value, 50000)
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("152", "KHH"))
                self.assertEqual((result_sheet.cell(4, 17).value, result_sheet.cell(4, 18).value), ("152", "KHH"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_ban_ra_saved_purchase_match_rule_converts_with_exact_khh_match(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_saved_rule_source_{run_id}.xlsx"
        purchase = outputs / f"_test_vietmax_saved_rule_purchase_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_saved_rule_result_{run_id}.xlsx"
        try:
            sales_workbook = Workbook()
            sales_sheet = sales_workbook.active
            sales_sheet.title = "Invoices"
            sales_sheet.cell(2, 9, "Ten nguoi mua")
            sales_sheet.cell(2, 10, "MST nguoi mua")
            sales_sheet.cell(2, 12, "Ma VT")
            sales_sheet.cell(2, 13, "Ten hang")
            sales_sheet.cell(2, 14, "DVT")
            sales_sheet.cell(2, 15, "So luong")
            sales_sheet.cell(2, 16, "Don gia")
            sales_sheet.cell(2, 18, "Thue suat")
            sales_sheet.cell(2, 19, "Tien hang")
            sales_sheet.cell(2, 20, "Tien thue")
            sales_sheet.cell(2, 21, "Tong TT")
            sales_sheet.cell(3, 9, "Cong ty A")
            sales_sheet.cell(3, 10, "MST")
            sales_sheet.cell(3, 12, 0)
            sales_sheet.cell(3, 13, "Giay in A4")
            sales_sheet.cell(3, 14, "To")
            sales_sheet.cell(3, 15, 35)
            sales_sheet.cell(3, 16, 600)
            sales_sheet.cell(3, 18, 8)
            sales_sheet.cell(3, 19, 21000)
            sales_sheet.cell(3, 20, 1680)
            sales_sheet.cell(3, 21, 22680)
            sales_sheet.cell(4, 9, "Cong ty A")
            sales_sheet.cell(4, 10, "MST")
            sales_sheet.cell(4, 12, 0)
            sales_sheet.cell(4, 13, "Giay my thuat")
            sales_sheet.cell(4, 14, "Kg")
            sales_sheet.cell(4, 15, 0.5)
            sales_sheet.cell(4, 16, 35000)
            sales_sheet.cell(4, 18, 8)
            sales_sheet.cell(4, 19, 17500)
            sales_sheet.cell(4, 20, 1400)
            sales_sheet.cell(4, 21, 18900)
            sales_sheet.cell(5, 9, "Cong ty A")
            sales_sheet.cell(5, 10, "MST")
            sales_sheet.cell(5, 12, 0)
            sales_sheet.cell(5, 13, "Giay in A4")
            sales_sheet.cell(5, 14, "Kg")
            sales_sheet.cell(5, 15, 850)
            sales_sheet.cell(5, 16, 25000)
            sales_sheet.cell(5, 18, 8)
            sales_sheet.cell(5, 19, 21250000)
            sales_sheet.cell(5, 20, 1700000)
            sales_sheet.cell(5, 21, 22950000)
            sales_workbook.save(source)
            sales_workbook.close()

            purchase_workbook = Workbook()
            purchase_sheet = purchase_workbook.active
            purchase_sheet.title = "ProcessedPurchase"
            purchase_sheet.cell(2, 12, "Ma VT")
            purchase_sheet.cell(2, 13, "Ten hang")
            purchase_sheet.cell(2, 14, "DVT")
            purchase_sheet.cell(2, 16, "Don gia")
            purchase_sheet.cell(3, 12, "GIAY-KG")
            purchase_sheet.cell(3, 13, "Giay in A4")
            purchase_sheet.cell(3, 14, "Kg")
            purchase_sheet.cell(3, 16, 21000)
            purchase_sheet.cell(4, 12, "GIAYMYTHUAT")
            purchase_sheet.cell(4, 13, "Giay my thuat")
            purchase_sheet.cell(4, 14, "Xap")
            purchase_sheet.cell(4, 16, 1925000)
            purchase_workbook.save(purchase)
            purchase_workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Giay in A4", "Giay my thuat"],
                "inventory_pairs": [{"id": "default", "ma_kho": "KVT", "tk_vat_tu": "156"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "vietmax_processed_purchase_path": str(purchase),
                "vietmax_ban_ra_purchase_match_rules": [
                    {
                        "sales_product": "Giay in A4",
                        "purchase_code": "GIAY-KG",
                        "purchase_product": "Giay in A4",
                        "sales_unit": "To",
                        "purchase_unit": "Kg",
                        "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                        "conversion_formula": "35 To = 1 Kg",
                        "confirmed": True,
                    },
                    {
                        "sales_product": "Giay my thuat",
                        "purchase_code": "GIAYMYTHUAT",
                        "purchase_product": "Giay my thuat",
                        "sales_unit": "Kg",
                        "purchase_unit": "Xap",
                        "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                        "conversion_formula": "55 Kg = 1 Xap",
                        "confirmed": True,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 14).value, "Kg")
                self.assertEqual(result_sheet.cell(3, 15).value, 1)
                self.assertEqual(result_sheet.cell(3, 16).value, 21000)
                self.assertEqual(result_sheet.cell(3, 19).value, 21000)
                self.assertEqual(result_sheet.cell(3, 20).value, 1680)
                self.assertEqual(result_sheet.cell(3, 21).value, 22680)
                self.assertEqual((result_sheet.cell(3, 22).value, result_sheet.cell(3, 23).value), ("152", "KHH"))
                self.assertEqual(result_sheet.cell(4, 14).value, "Xap")
                self.assertAlmostEqual(result_sheet.cell(4, 15).value, 0.009091, places=6)
                self.assertAlmostEqual(result_sheet.cell(4, 16).value * result_sheet.cell(4, 15).value, 17500, places=2)
                self.assertEqual(result_sheet.cell(4, 19).value, 17500)
                self.assertEqual(result_sheet.cell(4, 20).value, 1400)
                self.assertEqual(result_sheet.cell(4, 21).value, 18900)
                self.assertEqual((result_sheet.cell(4, 22).value, result_sheet.cell(4, 23).value), ("152", "KHH"))
                self.assertEqual(result_sheet.cell(5, 14).value, "Kg")
                self.assertEqual(result_sheet.cell(5, 15).value, 850)
                self.assertEqual(result_sheet.cell(5, 16).value, 25000)
                self.assertEqual(result_sheet.cell(5, 19).value, 21250000)
                self.assertEqual(result_sheet.cell(5, 20).value, 1700000)
                self.assertEqual(result_sheet.cell(5, 21).value, 22950000)
                self.assertEqual((result_sheet.cell(5, 22).value, result_sheet.cell(5, 23).value), ("152", "KHH"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            purchase.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_ban_ra_unchecked_purchase_match_does_not_override_convert_or_force_inventory(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_unchecked_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_unchecked_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 14, "DVT")
            sheet.cell(2, 15, "So luong")
            sheet.cell(2, 17, "TK vat tu")
            sheet.cell(2, 18, "Ma kho")
            sheet.cell(3, 9, "Cong ty A")
            sheet.cell(3, 10, "MST")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Unchecked Product")
            sheet.cell(3, 14, "To")
            sheet.cell(3, 15, 500)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Unchecked Product"],
                "inventory_pairs": [{"id": "default", "ma_kho": "KVT", "tk_vat_tu": "156"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "vietmax_ban_ra_purchase_matches": [
                    {
                        "sales_product": "Unchecked Product",
                        "purchase_code": "PURCHASE-CODE",
                        "sales_unit": "To",
                        "purchase_unit": "Ram",
                        "unit_mismatch": True,
                        "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                        "conversion_formula": "1 ram = 500 to",
                        "confirmed": False,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertNotEqual(result_sheet.cell(3, 12).value, "PURCHASE-CODE")
                self.assertEqual(result_sheet.cell(3, 14).value, "To")
                self.assertEqual(result_sheet.cell(3, 15).value, 500)
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("156", "KVT"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_ban_ra_purchase_match_overrides_code_and_inventory(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_match_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_match_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 15, "So luong")
            sheet.cell(2, 17, "TK vat tu")
            sheet.cell(2, 18, "Ma kho")
            sheet.cell(3, 9, "Cong ty A")
            sheet.cell(3, 10, "MST")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Matched Product")
            sheet.cell(3, 15, 1)
            sheet.cell(4, 9, "Cong ty A")
            sheet.cell(4, 10, "MST")
            sheet.cell(4, 12, 0)
            sheet.cell(4, 13, "Other Product")
            sheet.cell(4, 15, 1)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Matched Product", "Other Product"],
                "inventory_pairs": [{"id": "default", "ma_kho": "KVT", "tk_vat_tu": "156"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "vietmax_ban_ra_purchase_matches": [
                    {"sales_product": "Matched Product", "purchase_code": "PURCHASE-CODE", "confirmed": True},
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Matched Product", {}))
                self.assertNotEqual(result_sheet.cell(4, 12).value, "PURCHASE-CODE")
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("152", "KHH"))
                self.assertEqual((result_sheet.cell(4, 17).value, result_sheet.cell(4, 18).value), ("156", "KVT"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_unified_vietmax_sales_uses_processed_purchase_for_khh_and_excludes_from_default(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        purchase = outputs / f"_test_vietmax_processed_purchase_{run_id}.xlsx"
        sales = outputs / f"_test_vietmax_sales_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_sales_result_{run_id}.xlsx"
        try:
            purchase_wb = Workbook()
            purchase_ws = purchase_wb.active
            purchase_ws.title = "Purchase"
            purchase_ws.cell(2, 12, "Ma VT")
            purchase_ws.cell(2, 13, "Ten hang")
            purchase_ws.cell(2, 14, "DVT")
            purchase_ws.cell(2, 17, "TK vat tu")
            purchase_ws.cell(2, 18, "Ma kho")
            purchase_ws.cell(3, 12, "KHH-PROC-001")
            purchase_ws.cell(3, 13, "Giay in A4")
            purchase_ws.cell(3, 14, "Ram")
            purchase_ws.cell(3, 17, "152")
            purchase_ws.cell(3, 18, "KHH")
            purchase_wb.save(purchase)
            purchase_wb.close()

            info = app.validate_vietmax_processed_purchase_workbook(purchase)
            self.assertEqual(info["code_col"], "L")
            self.assertEqual(info["valid_rows"], 1)

            sales_wb = Workbook()
            sales_ws = sales_wb.active
            sales_ws.title = "Sales"
            sales_ws.cell(2, 9, "Ten nguoi mua")
            sales_ws.cell(2, 10, "MST nguoi mua")
            sales_ws.cell(2, 12, "Ma VT")
            sales_ws.cell(2, 13, "Ten hang")
            sales_ws.cell(2, 14, "DVT")
            sales_ws.cell(2, 15, "So luong")
            sales_ws.cell(2, 17, "TK vat tu")
            sales_ws.cell(2, 18, "Ma kho")
            for row, product in [(3, "Giay in A4"), (4, "In decal tem")]:
                sales_ws.cell(row, 9, "Cong ty A")
                sales_ws.cell(row, 10, "MST")
                sales_ws.cell(row, 12, 0)
                sales_ws.cell(row, 13, product)
                sales_ws.cell(row, 14, "Ram")
                sales_ws.cell(row, 15, 1)
            sales_wb.save(sales)
            sales_wb.close()

            process_workbook(sales, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax",
                "vietmax_phase": app.VIETMAX_PHASE_SALES,
                "vietmax_processed_purchase_path": str(purchase),
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Giay in A4", "In decal tem"],
                "inventory_pairs": [
                    {"id": "ktp", "ma_kho": "KTP", "tk_vat_tu": "154"},
                    {"id": "default", "ma_kho": "KTP", "tk_vat_tu": "154"},
                ],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "inventory_pair_rules": [
                    {"enabled": True, "source_col": "M", "operator": "contains", "value": "In", "pair_id": "ktp"},
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Sales"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Giay in A4", {}))
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("152", "KHH"))
                self.assertNotEqual(result_sheet.cell(4, 12).value, "KHH-PROC-001")
                self.assertEqual((result_sheet.cell(4, 17).value, result_sheet.cell(4, 18).value), ("154", "KTP"))
            finally:
                output.close()
        finally:
            purchase.unlink(missing_ok=True)
            sales.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_unified_vietmax_sales_exact_khh_overrides_stale_manual_match(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        purchase = outputs / f"_test_vietmax_exact_priority_purchase_{run_id}.xlsx"
        sales = outputs / f"_test_vietmax_exact_priority_sales_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_exact_priority_result_{run_id}.xlsx"
        try:
            purchase_wb = Workbook()
            purchase_ws = purchase_wb.active
            purchase_ws.title = "Purchase"
            purchase_ws.cell(2, 12, "Ma VT")
            purchase_ws.cell(2, 13, "Ten hang")
            purchase_ws.cell(2, 14, "DVT")
            purchase_ws.cell(2, 17, "TK vat tu")
            purchase_ws.cell(2, 18, "Ma kho")
            purchase_ws.cell(3, 12, "KHH-EXACT")
            purchase_ws.cell(3, 13, "Giay in A4")
            purchase_ws.cell(3, 14, "Ram")
            purchase_ws.cell(3, 17, "152")
            purchase_ws.cell(3, 18, "KHH")
            purchase_wb.save(purchase)
            purchase_wb.close()

            sales_wb = Workbook()
            sales_ws = sales_wb.active
            sales_ws.title = "Sales"
            sales_ws.cell(2, 9, "Ten nguoi mua")
            sales_ws.cell(2, 10, "MST nguoi mua")
            sales_ws.cell(2, 12, "Ma VT")
            sales_ws.cell(2, 13, "Ten hang")
            sales_ws.cell(2, 14, "DVT")
            sales_ws.cell(2, 15, "So luong")
            sales_ws.cell(2, 17, "TK vat tu")
            sales_ws.cell(2, 18, "Ma kho")
            sales_ws.cell(3, 9, "Cong ty A")
            sales_ws.cell(3, 10, "MST")
            sales_ws.cell(3, 12, 0)
            sales_ws.cell(3, 13, "Giay in A4")
            sales_ws.cell(3, 14, "Ram")
            sales_ws.cell(3, 15, 1)
            sales_wb.save(sales)
            sales_wb.close()

            process_workbook(sales, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax",
                "vietmax_phase": app.VIETMAX_PHASE_SALES,
                "vietmax_processed_purchase_path": str(purchase),
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Giay in A4", "Giay my thuat"],
                "vietmax_ban_ra_purchase_matches": [
                    {"sales_product": "Giay in A4", "purchase_code": "STALE-MATCH", "confirmed": True},
                ],
                "inventory_pairs": [{"id": "default", "ma_kho": "KTP", "tk_vat_tu": "154"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Sales"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Giay in A4", {}))
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("152", "KHH"))
            finally:
                output.close()
        finally:
            purchase.unlink(missing_ok=True)
            sales.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_sales_internal_merge_affects_final_purchase_match_lookup(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_sales_internal_merge_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_sales_internal_merge_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 15, "So luong")
            sheet.cell(3, 9, "Cong ty A")
            sheet.cell(3, 10, "MST")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Giấy Couche 300 gms")
            sheet.cell(3, 15, 1)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Giấy Couche 300 gms"],
                "vietmax_ban_ra_sales_internal_merges": [
                    {"product": "Giấy Couche 300 gms", "similar_product": "Giấy Couche 300 gsm", "confirmed": True},
                ],
                "vietmax_ban_ra_purchase_matches": [
                    {"sales_product": "Giấy Couche 300 gsm", "purchase_code": "PURCHASE-CODE", "confirmed": True},
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Giấy Couche 300 gsm", {}))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_ban_ra_repeated_phrase_rule_affects_processed_excel(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_repeat_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_repeat_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 15, "So luong")
            sheet.cell(3, 9, "Cong ty A")
            sheet.cell(3, 10, "111")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Danh sách số 01 số 02 số 03")
            sheet.cell(3, 15, 1)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["111"],
                "process_mst": ["111"],
                "mst_safe_id": ["111|||0"],
                "selected_products_0": ["Danh sách số 01 số 02 số 03"],
                "repeated_phrase_removals": ["số"],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, "DANHSS010203")
                self.assertLessEqual(len(result_sheet.cell(3, 12).value), app.MAX_CODE_LENGTH)
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_sales_internal_merge_reuses_similar_product_code(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_sales_merge_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_sales_merge_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 15, "So luong")
            sheet.cell(3, 9, "Cong ty A")
            sheet.cell(3, 10, "111")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 13, "Giấy Couche 300 gms")
            sheet.cell(3, 15, 1)
            sheet.cell(4, 9, "Cong ty A")
            sheet.cell(4, 10, "111")
            sheet.cell(4, 12, 0)
            sheet.cell(4, 13, "Giấy Couche 300 gsm")
            sheet.cell(4, 15, 1)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["111"],
                "process_mst": ["111"],
                "mst_safe_id": ["111|||0"],
                "selected_products_0": ["Giấy Couche 300 gms", "Giấy Couche 300 gsm"],
                "comparison_scope": app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
                "vietmax_ban_ra_sales_internal_merges": [
                    {
                        "product": "Giấy Couche 300 gms",
                        "similar_product": "Giấy Couche 300 gsm",
                        "company": "Cong ty A",
                        "mst": "111",
                        "company_key": "111",
                        "comparison_scope": app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
                        "confirmed": True,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, result_sheet.cell(4, 12).value)
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Giấy Couche 300 gsm", {}))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)

    def test_vietmax_same_company_confirmed_match_affects_only_matching_company_rows(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_vietmax_same_company_source_{run_id}.xlsx"
        result = outputs / f"_test_vietmax_same_company_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 9, "Ten nguoi mua")
            sheet.cell(2, 10, "MST nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 13, "Ten hang")
            sheet.cell(2, 14, "DVT")
            sheet.cell(2, 15, "So luong")
            sheet.cell(2, 17, "TK vat tu")
            sheet.cell(2, 18, "Ma kho")
            for row, company, mst in [(3, "Cong ty A", "111"), (4, "Cong ty B", "222")]:
                sheet.cell(row, 9, company)
                sheet.cell(row, 10, mst)
                sheet.cell(row, 12, 0)
                sheet.cell(row, 13, "Matched Product")
                sheet.cell(row, 14, "To")
                sheet.cell(row, 15, 500)
            workbook.save(source)
            workbook.close()

            process_workbook(source, result, {
                "company_col": "I",
                "mst_col": "J",
                "product_col": "M",
                "qty_col": "O",
                "output_col": "L",
                "profile": "vietmax_ban_ra",
                "include_company_prefix": False,
                "all_mst": ["111", "222"],
                "process_mst": ["111", "222"],
                "mst_safe_id": ["111|||0", "222|||1"],
                "selected_products_0": ["Matched Product"],
                "selected_products_1": ["Matched Product"],
                "inventory_pairs": [{"id": "default", "ma_kho": "KVT", "tk_vat_tu": "156"}],
                "use_default_inventory_pair": True,
                "default_inventory_pair_id": "default",
                "comparison_scope": app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
                "vietmax_ban_ra_purchase_matches": [
                    {
                        "sales_product": "Matched Product",
                        "sales_company": "Cong ty A",
                        "sales_mst": "111",
                        "purchase_code": "PURCHASE-CODE",
                        "sales_unit": "To",
                        "purchase_unit": "Ram",
                        "unit_mismatch": True,
                        "conversion_mode": app.VIETMAX_CONVERSION_MODE_QTY_AND_UNIT,
                        "conversion_formula": "1 ram = 500 to",
                        "comparison_scope": app.VIETMAX_COMPARISON_SCOPE_SAME_COMPANY,
                        "confirmed": True,
                    },
                ],
            })

            output = load_workbook(result, data_only=True)
            try:
                result_sheet = output["Invoices"]
                self.assertEqual(result_sheet.cell(3, 12).value, make_product_part("vietmax_ban_ra", "Matched Product", {}))
                self.assertNotEqual(result_sheet.cell(4, 12).value, "PURCHASE-CODE")
                self.assertEqual(result_sheet.cell(3, 14).value, "Ram")
                self.assertEqual(result_sheet.cell(4, 14).value, "To")
                self.assertEqual(result_sheet.cell(3, 15).value, 1)
                self.assertEqual(result_sheet.cell(4, 15).value, 500)
                self.assertEqual((result_sheet.cell(3, 17).value, result_sheet.cell(3, 18).value), ("152", "KHH"))
                self.assertEqual((result_sheet.cell(4, 17).value, result_sheet.cell(4, 18).value), ("156", "KVT"))
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)


    def test_output_suffixes_are_fdi_and_nhap_kho(self):
        output = resolve_output_path("hoa_don.xlsx", "")
        self.assertEqual(output.name, "hoa_don_fdi.xlsx")
        self.assertEqual(up_ban_ra_output_path(output).name, "hoa_don_fdi_nhap_kho.xlsx")

    def test_selected_output_filename_is_normalized_to_fdi_suffix(self):
        output = resolve_output_path("hoa_don.xlsx", "outputs/custom_name.xlsx")
        self.assertEqual(output.name, "custom_name_fdi.xlsx")
        self.assertEqual(resolve_output_path("hoa_don.xlsx", "outputs/custom_name_fdi.xlsx").name, "custom_name_fdi.xlsx")

    def test_nhap_kho_workbook_uses_inventory_columns(self):
        outputs = Path(__file__).parent / "outputs"
        outputs.mkdir(exist_ok=True)
        run_id = uuid4().hex
        source = outputs / f"_test_nhap_kho_source_{run_id}.xlsx"
        result = outputs / f"_test_nhap_kho_result_{run_id}.xlsx"
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Invoices"
            sheet.cell(2, 3, "So hoa don")
            sheet.cell(2, 4, "Ngay hoa don")
            sheet.cell(2, 6, "Cong ty")
            sheet.cell(2, 7, "MST")
            sheet.cell(2, 10, "Nguoi mua")
            sheet.cell(2, 12, "Ma VT")
            sheet.cell(2, 14, "Ten hang")
            sheet.cell(2, 15, "So luong")
            sheet.cell(2, 16, "So luong xu ly")
            sheet.cell(2, 17, "Don gia")
            sheet.cell(2, 18, "Thue")
            sheet.cell(2, 23, "Ghi chu")
            sheet.cell(3, 3, "HD-1")
            sheet.cell(3, 4, "01/01/2026")
            sheet.cell(3, 6, "Cong ty A")
            sheet.cell(3, 7, "MST")
            sheet.cell(3, 10, "Nguoi mua A")
            sheet.cell(3, 12, 0)
            sheet.cell(3, 14, "Hang A")
            sheet.cell(3, 15, 2)
            sheet.cell(3, 16, 2)
            sheet.cell(3, 17, 100)
            sheet.cell(3, 18, 0.1)
            sheet.cell(3, 23, "Ghi chu")
            workbook.save(source)
            workbook.close()

            processed = process_workbook(source, result, {
                "company_col": "F",
                "mst_col": "G",
                "product_col": "N",
                "qty_col": "P",
                "output_col": "L",
                "profile": "son_phuong",
                "include_company_prefix": False,
                "all_mst": ["MST"],
                "process_mst": ["MST"],
                "mst_safe_id": ["MST|||0"],
                "selected_products_0": ["Hang A"],
                "inventory_pairs": [{"id": "only", "ma_kho": "KVT", "tk_vat_tu": "152"}],
            })
            stream = create_up_ban_ra_workbook(processed)
            output = load_workbook(stream, data_only=True)
            try:
                sheet = output.active
                self.assertEqual(sheet["AB2"].value, "152")
                self.assertEqual(sheet["AF2"].value, "KVT")
                self.assertNotEqual(sheet["AB2"].value, "KVT")
                self.assertNotEqual(sheet["AF2"].value, "152")
            finally:
                output.close()
        finally:
            source.unlink(missing_ok=True)
            result.unlink(missing_ok=True)



if __name__ == "__main__":
    unittest.main()
