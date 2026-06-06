import unittest
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook, load_workbook

from app import create_up_ban_ra_workbook, default_config, make_product_part, normalize_config, process_workbook, profile_key, resolve_output_path, up_ban_ra_output_path


class ProcessWorkbookTests(unittest.TestCase):
    def test_vietmax_profiles_are_available_by_default(self):
        config = default_config()
        self.assertEqual(profile_key("vietmax"), "vietmax_mua_vao")
        self.assertIn("vietmax_mua_vao", config["profiles"])
        self.assertIn("vietmax_ban_ra", config["profiles"])
        self.assertNotIn("vietmax", config["profiles"])

    def test_legacy_vietmax_profile_config_maps_to_mua_vao(self):
        config = normalize_config({
            "selected_profile": "vietmax",
            "profiles": {
                "vietmax": {
                    "manual_code_overrides": {"MST|||Hang A": "VM.A"},
                    "inventory_pairs": [{"id": "one", "ma_kho": "KHO", "tk_vat_tu": "152"}],
                }
            },
        })

        self.assertEqual(config["selected_profile"], "vietmax_mua_vao")
        self.assertEqual(config["profiles"]["vietmax_mua_vao"]["manual_code_overrides"], {"MST|||Hang A": "VM.A"})
        self.assertEqual(config["profiles"]["vietmax_mua_vao"]["inventory_pairs"][0]["id"], "one")

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

    def test_vietmax_ban_ra_fallback_keeps_uppercase_number_tokens(self):
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Pin SYNWK-F1840N 500ML dung lượng", {}),
            "PINSYNWK-F1840N500MLDULU",
        )
        self.assertEqual(
            make_product_part("vietmax_ban_ra", "Màng tự dính SYNWK-F1840N", {}),
            "MANGTUDISYNWK-F1840N",
        )

    def test_vietmax_product_code_uses_all_words_for_note_patterns(self):
        examples = {
            "Giấy Couche": "GIAYCOUCHE",
            "Giấy offset": "GIAYOFFSET",
            "Giấy Cacbon": "GIAYCACBON",
            "Giấy in BB58/92gsm": "GIAYINBB5892GSM",
            "Giấy ivory ningbo 230gsm khổ 62cm": "GIAYIVORYNINGBO230GSMKHO62CM",
        }
        for product_name, expected in examples.items():
            with self.subTest(product_name=product_name):
                self.assertEqual(make_product_part("vietmax_mua_vao", product_name, {}), expected)

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

    def test_inventory_first_rule_wins_and_multiple_rules_act_as_or(self):
        headers, values, _ = self.run_inventory_case(
            [{"product": "Paper special"}, {"product": "Ink special"}],
            {
                "inventory_pairs": [
                    {"id": "paper", "ma_kho": "KHO-PAPER", "tk_vat_tu": "152-P"},
                    {"id": "special", "ma_kho": "KHO-SPECIAL", "tk_vat_tu": "152-S"},
                ],
                "inventory_pair_rules": [
                    {"source_col": "N", "operator": "contains", "value": "paper", "pair_id": "paper", "enabled": True},
                    {"source_col": "N", "operator": "contains", "value": "special", "pair_id": "special", "enabled": True},
                ],
            },
        )

        self.assertEqual(headers, ("TK vật tư", "Mã kho"))
        self.assertEqual(values, [("152-P", "KHO-PAPER"), ("152-S", "KHO-SPECIAL")])

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
