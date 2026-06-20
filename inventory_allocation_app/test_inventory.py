import unittest
import time
from io import BytesIO
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from inventory_allocation_app import app as inventory_app
from inventory_allocation_app.app import DEFAULT_BAREM_MAP, allocate_stock, build_inventory_ledger, build_sales_report_rows_from_ledger, clean_mapping, create_output_workbook, find_sale_only_codes, generic_steel_sale_type, parse_barem_file, preview_workbook, read_lines, resolve_barem_weight, sales_summary_rows_for_ui, steel_coating, steel_kind, steel_profile_key


def line(kind, code, quantity, price=None, row_number=3, invoice_no="", invoice_date="", **extra):
    base = code.rsplit(".", 1)[0] if code.rsplit(".", 1)[-1].isdigit() else code
    suffix = int(code.rsplit(".", 1)[-1]) if code.rsplit(".", 1)[-1].isdigit() else None
    result = {
        "kind": kind,
        "row_number": row_number,
        "variant_code": code,
        "base_code": base,
        "suffix": suffix,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "product_name": "Hang A",
        "quantity": quantity,
        "unit_price": price,
    }
    result.update(extra)
    return result


def invoice_file(code, quantity, price):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Invoices"
    sheet.cell(2, 12, "Ma VT")
    sheet.cell(2, 13, "Ten hang")
    sheet.cell(2, 15, "So luong")
    sheet.cell(2, 16, "Don gia")
    sheet.cell(2, 3, "So hoa don")
    sheet.cell(2, 4, "Ngay hoa don")
    sheet.cell(3, 12, code)
    sheet.cell(3, 13, "Hang A")
    sheet.cell(3, 15, quantity)
    sheet.cell(3, 16, price)
    sheet.cell(3, 3, "HD1")
    sheet.cell(3, 4, "02/01/2025")
    content = BytesIO()
    workbook.save(content)
    content.seek(0)
    return content


def invoice_file_with_name(code, product_name, quantity, price):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Invoices"
    sheet.cell(2, 12, "Ma VT")
    sheet.cell(2, 13, "Ten hang")
    sheet.cell(2, 15, "So luong")
    sheet.cell(2, 16, "Don gia")
    sheet.cell(2, 3, "So hoa don")
    sheet.cell(2, 4, "Ngay hoa don")
    sheet.cell(3, 12, code)
    sheet.cell(3, 13, product_name)
    sheet.cell(3, 15, quantity)
    sheet.cell(3, 16, price)
    sheet.cell(3, 3, "HD1")
    sheet.cell(3, 4, "02/01/2026")
    content = BytesIO()
    workbook.save(content)
    content.seek(0)
    return content


def up_sales_file(code, quantity, price):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "UP Bán ra"
    sheet.cell(1, 5, "Số chứng từ\n(so_ct)")
    sheet.cell(1, 6, "Ngày chứng từ\n(ngay_ct)")
    sheet.cell(1, 8, "Số lượng:Q\n(so_luong)")
    sheet.cell(1, 15, "Giá bán:P0\n(gia2)")
    sheet.cell(1, 23, "Thuế suất\n(thue_suat)")
    sheet.cell(1, 25, "Tiền thuế:N0\n(tien_thue)")
    sheet.cell(1, 33, "Mã vật tư\n(ma_vt)")
    sheet.cell(1, 34, "Diễn giải\n(dien_giai)")
    sheet.cell(2, 5, "HDSP1")
    sheet.cell(2, 6, "02/01/2026")
    sheet.cell(2, 8, quantity)
    sheet.cell(2, 15, price)
    sheet.cell(2, 23, 10)
    sheet.cell(2, 25, quantity * price * 0.1)
    sheet.cell(2, 33, code)
    sheet.cell(2, 34, "Xuất bán hàng")
    content = BytesIO()
    workbook.save(content)
    content.seek(0)
    return content


class AllocationTests(unittest.TestCase):
    def test_sale_warehouse_without_inbound_keeps_zero_cost_and_zero_margin(self):
        purchase = [line("purchase", "PAPER", 10, 50, invoice_no="M1", invoice_date="01/01/2026", warehouse_code="KHH", warehouse_account="156")]
        sales = [line("sales", "PAPER", 3, 120, invoice_no="B1", invoice_date="02/01/2026", party_name="Khach A", warehouse_code="KGCI", warehouse_account="1552")]
        allocations, _, _, _ = allocate_stock([], purchase, sales, {"allow_negative_export": True})

        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)
        report_rows = build_sales_report_rows_from_ledger(ledger)
        kgci_rows = [row for row in report_rows if row["warehouse_code"] == "KGCI"]

        self.assertEqual(len(kgci_rows), 1)
        self.assertEqual(kgci_rows[0]["cost_amount"], 0)
        self.assertEqual(kgci_rows[0]["profit_amount"], 0)
        self.assertTrue(kgci_rows[0]["cost_missing"])
        summary = sales_summary_rows_for_ui(report_rows)[0]
        self.assertEqual(summary["margin_percent"], 0)
        self.assertEqual(summary["profit_amount"], 0)

    def test_son_phuong_detects_pipe_and_box_products(self):
        self.assertEqual(steel_kind("Ống tôn mạ CN 30x60x1.4x6.0"), "box")
        self.assertEqual(steel_kind("Thép hộp mạ kẽm 20x20x1.1x6000"), "box")
        self.assertEqual(steel_kind("Ống tôn mạ vuông 40x1.4x6.0"), "box")
        self.assertEqual(steel_kind("Ống thép tròn Φ26.65x1.4"), "pipe")
        self.assertEqual(steel_kind("Ống thép F59.9x1.4"), "pipe")
        self.assertEqual(steel_kind("Ống thép D21.2x1.2"), "pipe")
        self.assertEqual(steel_coating("Thép hộp đen 20x20x1.1"), "black")
        self.assertEqual(steel_coating("Thép hộp 20x20x1.1"), "galvanized")
        self.assertEqual(steel_kind("Thép cuộn cán nóng các loại", "ONGTRONGMA"), "unknown")
        self.assertEqual(generic_steel_sale_type("Thép ống hộp các loại"), "pipe_box")
        self.assertEqual(generic_steel_sale_type("Thép hộp các loại"), "box")
        self.assertEqual(generic_steel_sale_type("Thép ống các loại"), "pipe")

    def test_son_phuong_generic_sale_uses_allowed_steel_types_by_low_profit_then_stock(self):
        purchase = [
            line("purchase", "BOXA", 4, 90, product_name="Thép hộp 20x20x1.1x6000"),
            line("purchase", "BOXB", 12, 90, product_name="Thép hộp 30x60x1.4x6000"),
            line("purchase", "PIPEA", 20, 95, product_name="Ống thép tròn Φ26.65x1.4"),
        ]
        sales = [line("sales", "GENBOX", 10, 100, product_name="Thép hộp các loại")]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {
                "company_profile": "son_phuong",
                "max_loss_percent": 10,
                "max_profit_percent": 25,
                "son_phuong_split_counts": {"box": 2, "pipe": 2, "pipe_box": 2},
            },
        )

        used_codes = [item["variant_code"] for item in allocations[0]["used"]]
        self.assertEqual(used_codes, ["BOXB", "BOXA", "BOXB"])
        self.assertEqual(summary["material_quantity"], 10)
        self.assertEqual(summary["finished_quantity"], 0)
        self.assertNotIn("PIPEA", used_codes)

    def test_son_phuong_priority_uses_low_profit_before_loss(self):
        purchase = [
            line("purchase", "PROFIT_LOW", 100, 98, product_name="Thep hop 20x20x1.1"),
            line("purchase", "PROFIT_HIGH", 100, 80, product_name="Thep hop 30x60x1.4"),
            line("purchase", "LOSS_LOW", 100, 102, product_name="Thep hop 25x50x1.4"),
            line("purchase", "LOSS_HIGH", 100, 130, product_name="Thep hop 40x80x1.4"),
        ]
        sales = [line("sales", "GENBOX", 100, 100, product_name="Thep hop cac loai")]

        allocations, _, _, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong", "barem_tolerance_percent": 100},
            barem_map={
                "by_code": {"PROFITLOW": 10, "PROFITHIGH": 10, "LOSSLOW": 10, "LOSSHIGH": 10},
                "by_profile": {},
            },
        )

        used_codes = [item["variant_code"] for item in allocations[0]["used"]]
        used_codes = list(dict.fromkeys(used_codes))
        self.assertEqual(used_codes[:4], ["PROFIT_LOW", "PROFIT_HIGH", "LOSS_LOW", "LOSS_HIGH"])

    def test_son_phuong_shape_steel_matches_by_product_family_across_codes(self):
        purchase = [line("purchase", "TCC.THV50", 100, 90, product_name="Thép hình V50")]
        sales = [line("sales", "OTHER.CODE", 25, 100, product_name="Thép V 50")]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong"},
        )

        self.assertEqual(allocations[0]["used"][0]["variant_code"], "TCC.THV50")
        self.assertEqual(summary["material_quantity"], 25)
        self.assertEqual(summary["negative_export_quantity"], 0)

    def test_son_phuong_generic_sale_can_split_across_more_than_six_codes(self):
        purchase = [
            line("purchase", f"BOX{i}", 30, 95 + i, product_name=f"Thep hop {20 + i}x{20 + i}x1.1")
            for i in range(1, 9)
        ]
        sales = [line("sales", "GENBOX", 140, 100, product_name="Thep hop cac loai")]
        barem_map = {"by_code": {f"BOX{i}": 10 for i in range(1, 9)}, "by_profile": {}}

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong", "barem_tolerance_percent": 100},
            barem_map=barem_map,
        )

        used_codes = list(dict.fromkeys(item["variant_code"] for item in allocations[0]["used"]))
        self.assertGreater(len(used_codes), 6)
        self.assertEqual(sum(item["quantity"] for item in allocations[0]["used"]), 140)
        self.assertEqual(summary["negative_export_quantity"], 0)

    def test_son_phuong_pipe_generic_does_not_use_box_but_pipe_box_uses_both(self):
        purchase = [
            line("purchase", "BOXA", 12, 90, product_name="Thép hộp 20x20x1.1x6000"),
            line("purchase", "PIPEA", 12, 90, product_name="Ống thép tròn Φ26.65x1.4"),
        ]
        policy = {
            "company_profile": "son_phuong",
            "max_loss_percent": 10,
            "max_profit_percent": 25,
            "son_phuong_split_counts": {"box": 2, "pipe": 2, "pipe_box": 2},
        }

        pipe_allocations, _, _, _ = allocate_stock(
            [], purchase, [line("sales", "GENPIPE", 5, 100, product_name="Thép ống các loại")], policy
        )
        self.assertEqual([item["variant_code"] for item in pipe_allocations[0]["used"]], ["PIPEA"])

        pipe_box_allocations, _, _, _ = allocate_stock(
            [], purchase, [line("sales", "GENPIPEBOX", 10, 100, product_name="Thép ống hộp các loại")], policy
        )
        self.assertEqual(
            {item["variant_code"] for item in pipe_box_allocations[0]["used"]},
            {"BOXA", "PIPEA"},
        )

    def test_son_phuong_barem_plan_uses_only_barem_multiples_and_reports_missing_barem(self):
        purchase = [
            line("purchase", "BOXA", 100, 90, product_name="Thép hộp 20x20x1.1x6000"),
            line("purchase", "BOXB", 100, 91, product_name="Thép hộp 30x60x1.4x6000"),
            line("purchase", "PIPEA", 100, 92, product_name="Ống thép tròn Φ26.65x1.4"),
            line("purchase", "PIPEMISS", 100, 93, product_name="Ống thép tròn Φ59.9x1.4"),
        ]
        sales = [line("sales", "GENPIPEBOX", 100, 100, product_name="Thép ống hộp các loại")]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {
                "company_profile": "son_phuong",
                "barem_tolerance_percent": 5,
            },
            barem_map={"BOXA": 30, "BOXB": 30, "PIPEA": 30},
        )

        used = allocations[0]["used"]
        self.assertTrue(used)
        self.assertTrue(all(item["variant_code"] != "PIPEMISS" for item in used))
        remainders = [item["quantity"] % 30 for item in used]
        self.assertLessEqual(sum(1 for value in remainders if value > 0.000001), 1)
        self.assertEqual(summary["material_quantity"], 100)
        self.assertEqual(summary["finished_quantity"], 0)
        self.assertEqual(summary["missing_barem_count"], 1)
        self.assertEqual(summary["missing_barem_report"][0]["variant_code"], "PIPEMISS")

    def test_son_phuong_barem_plan_respects_min_max_per_code_per_sale_row(self):
        purchase = [
            line("purchase", "BOXA", 100, 90, product_name="Thep hop 20x20x1.1"),
            line("purchase", "BOXB", 100, 90, product_name="Thep hop 30x60x1.4"),
            line("purchase", "BOXC", 100, 90, product_name="Thep hop 25x50x1.4"),
        ]
        sales = [line("sales", "GENBOX", 90, 100, product_name="Thep hop cac loai")]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {
                "company_profile": "son_phuong",
                "generic_min_take_quantity": 20,
                "generic_max_take_quantity": 30,
            },
            barem_map={"BOXA": 10, "BOXB": 10, "BOXC": 10},
        )

        used = allocations[0]["used"]
        self.assertEqual(summary["material_quantity"], 90)
        self.assertTrue(used)
        self.assertTrue(all(item["quantity"] <= 30 for item in used))
        self.assertTrue(all(item["quantity"] >= 20 for item in used if not item.get("barem_remainder")))

    def test_purchase_classification_preview_returns_unknown_rows_for_filter(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.cell(2, 12, "Ma VT")
        sheet.cell(2, 13, "Ten hang")
        sheet.cell(2, 15, "So luong")
        sheet.cell(2, 16, "Don gia")
        for index in range(3, 256):
            sheet.cell(index, 12, f"BOX{index}")
            sheet.cell(index, 13, "Thep hop 20x20x1.1")
            sheet.cell(index, 15, 1)
            sheet.cell(index, 16, 10)
        sheet.cell(256, 12, "UNKNOWN")
        sheet.cell(256, 13, "Thep can nong cac loai")
        sheet.cell(256, 15, 1)
        sheet.cell(256, 16, 10)
        content = BytesIO()
        workbook.save(content)
        content.seek(0)

        response = inventory_app.app.test_client().post(
            "/api/purchase-classification-preview",
            data={
                "file": (content, "purchase.xlsx"),
                "mapping": "{}",
                "policy": '{"company_profile": "son_phuong"}',
            },
            content_type="multipart/form-data",
        )
        data = response.get_json()

        self.assertEqual(response.status_code, 200)
        self.assertTrue(data["limited"])
        self.assertEqual(data["unknown_rows"][0]["variant_code"], "UNKNOWN")

    def test_barem_file_applies_by_steel_type_and_dimension_across_company_codes(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Mã VT", "Tên hàng", "Kg / barem"])
        sheet.append(["SRC-A", "Ống thép tròn Φ26.65x1.4", 12.5])
        sheet.append(["SRC-B", "Thép hộp 20x20x1.1x6000", 7.2])
        content = BytesIO()
        workbook.save(content)
        content.seek(0)

        barem_map = parse_barem_file(content.getvalue())

        self.assertEqual(steel_profile_key("Ống thép F26.65x1.4", "OTHER-A"), "pipe|galvanized|26.65x1.4")
        self.assertEqual(steel_profile_key("Thép hộp 20x20x1.1", "OTHER-B"), "box|galvanized|20x20x1.1")
        purchase = [
            line("purchase", "OTHER-A", 100, 90, product_name="Ống thép F26.65x1.4"),
            line("purchase", "OTHER-B", 100, 91, product_name="Thép hộp 20x20x1.1"),
        ]
        self.assertEqual(resolve_barem_weight(barem_map, purchase[0]), 12.5)
        self.assertEqual(resolve_barem_weight(barem_map, purchase[1]), 7.2)
        sales = [line("sales", "GENPIPEBOX", 40, 100, product_name="Thép ống hộp các loại")]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong", "barem_tolerance_percent": 5},
            barem_map=barem_map,
        )

        remainder_rows = []
        for item in allocations[0]["used"]:
            barem = resolve_barem_weight(barem_map, {"variant_code": item["variant_code"], "product_name": next(row["product_name"] for row in purchase if row["variant_code"] == item["variant_code"])})
            remainder = item["quantity"] % barem
            if min(abs(remainder), abs(barem - remainder)) >= 0.000001:
                remainder_rows.append(item)
        self.assertLessEqual(len(remainder_rows), 1)
        self.assertTrue(all(item.get("barem_remainder") for item in remainder_rows))
        self.assertEqual(summary["material_quantity"], 40)
        self.assertEqual(summary["finished_quantity"], 0)
        self.assertEqual(summary["missing_barem_count"], 0)

    def test_builtin_barem_separates_black_and_galvanized_and_defaults_to_galvanized(self):
        galvanized = line("purchase", "PIPE-G", 100, 90, product_name="Ống thép D59.9x2")
        black = line("purchase", "PIPE-B", 100, 90, product_name="Ống thép đen D59.9x2")
        box = line("purchase", "BOX-G", 100, 90, product_name="Thép hộp 20x20x1.1")

        self.assertEqual(resolve_barem_weight(DEFAULT_BAREM_MAP, galvanized), 17.14)
        self.assertEqual(resolve_barem_weight(DEFAULT_BAREM_MAP, black), 17.13)
        self.assertEqual(resolve_barem_weight(DEFAULT_BAREM_MAP, box), 3.87)

        sales = [line("sales", "GENPIPE", 20, 100, product_name="Thép ống các loại")]
        allocations, _, _, _ = allocate_stock(
            [],
            [black, galvanized],
            sales,
            {"company_profile": "son_phuong", "barem_tolerance_percent": 5},
            barem_map=DEFAULT_BAREM_MAP,
        )
        self.assertEqual({item["variant_code"] for item in allocations[0]["used"]}, {"PIPE-G"})
        self.assertLessEqual(sum(1 for item in allocations[0]["used"] if item.get("barem_remainder")), 1)

    def test_son_phuong_shortage_stays_negative_in_khh_without_ktp(self):
        purchase = [line(
            "purchase", "PIPE-G", 20, 90,
            product_name="Ống thép D59.9x2",
            invoice_date="01/01/2025",
            invoice_date_iso="2025-01-01",
        )]
        sales = [line(
            "sales", "THEP-CAC-LOAI", 50, 100,
            product_name="Thép ống các loại",
            invoice_date="01/01/2025",
            invoice_date_iso="2025-01-01",
        )]

        allocations, _, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong"},
            barem_map=DEFAULT_BAREM_MAP,
        )
        ledger = build_inventory_ledger([], purchase, allocations, company_profile="son_phuong")

        self.assertEqual([warehouse["warehouse_code"] for warehouse in ledger["warehouses"]], ["KHH"])
        self.assertEqual(summary["finished_quantity"], 0)
        self.assertGreater(summary["negative_export_quantity"], 0)
        self.assertEqual(allocations[0]["used"][0]["ledger_variant_code"], "PIPE-G")
        self.assertTrue(allocations[0]["used"][-1]["negative_export"])

    def test_son_phuong_export_removes_old_ktp_sheets_and_writes_only_khh_ledger(self):
        purchase = [line(
            "purchase", "PIPE-G", 10, 90, invoice_no="MUA1", invoice_date="01/01/2025",
            invoice_date_iso="2025-01-01", product_name="Ong thep D59.9x2", unit_name="kg", line_amount=900,
        )]
        sales = [line(
            "sales", "GENPIPE", 12, 100, invoice_no="BAN1", invoice_date="02/01/2025",
            invoice_date_iso="2025-01-02", product_name="Thep ong cac loai", unit_name="kg", line_amount=1200,
        )]
        allocations, stock_rows, summary, _ = allocate_stock(
            [],
            purchase,
            sales,
            {"company_profile": "son_phuong"},
            barem_map={"PIPE-G": 5},
        )
        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales, company_profile="son_phuong")
        ledger["date_range"] = {"from": "2025-01-02", "to": "2025-01-02"}
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Invoices"
        sheet.cell(2, 12, "Ma VT")
        sheet.cell(2, 13, "Ten hang")
        sheet.cell(2, 15, "So luong")
        sheet.cell(2, 16, "Don gia")
        for old_name in ("SoChiTietKTP", "SoChiTietHH_TP", "TongHopNXT_KTP", "BaoCaoBH_KTP", "BangKeHDBH_KTP"):
            workbook.create_sheet(old_name)
        content = BytesIO()
        workbook.save(content)
        content.seek(0)

        result_content = create_output_workbook(
            content.getvalue(),
            {"sheet": "Invoices", "header_row": 2, "data_start_row": 3},
            allocations,
            stock_rows,
            summary,
            {"company_profile": "son_phuong"},
            purchase_lines=purchase,
            sales_lines=sales,
            ledger=ledger,
        )
        result = load_workbook(result_content)

        self.assertIn("SoChiTietKHH", result.sheetnames)
        self.assertNotIn("SoChiTietKTP", result.sheetnames)
        self.assertNotIn("SoChiTietHH_TP", result.sheetnames)
        self.assertIn("TongHopNXT_KHH", result.sheetnames)
        self.assertNotIn("TongHopNXT_KTP", result.sheetnames)
        self.assertIn("BaoCaoBH_KHH", result.sheetnames)
        self.assertNotIn("BaoCaoBH_KTP", result.sheetnames)
        self.assertIn("BangKeHDBH_KHH", result.sheetnames)
        self.assertNotIn("BangKeHDBH_KTP", result.sheetnames)

    def test_ambiguous_steel_report_does_not_treat_nong_as_ong(self):
        self.assertIsNone(inventory_app.steel_barem_ambiguity({
            "product_name": "Thép cuộn cán nóng các loại",
            "variant_code": "TH.TCCNCL",
        }))

    def test_default_barem_api_returns_four_reference_tables(self):
        response = inventory_app.app.test_client().get("/api/default-barem")
        data = response.get_json()

        self.assertEqual(response.status_code, 200)
        self.assertEqual(data["count"], len(data["rows"]))
        self.assertEqual({
            (row["kind"], row["coating"])
            for row in data["rows"]
        }, {
            ("box", "black"),
            ("pipe", "black"),
            ("pipe", "galvanized"),
            ("box", "galvanized"),
        })

    def test_uses_lowest_profit_split_code_first(self):
        purchase = [
            line("purchase", "A.001", 5, 90, invoice_no="MUA1"),
            line("purchase", "A.002", 10, 98, invoice_no="MUA2"),
        ]
        sales = [line("sales", "A", 12, 100)]

        allocations, stock, summary, _ = allocate_stock([], purchase, sales)

        used = allocations[0]["used"]
        self.assertEqual([item["variant_code"] for item in used], ["A.002", "A.001"])
        self.assertEqual([item["quantity"] for item in used], [10, 2])
        self.assertEqual([item["variant_code"] for item in allocations[0]["inventory_before"]], ["A.002", "A.001"])
        self.assertEqual(allocations[0]["inventory_after"][0]["variant_code"], "A.001")
        self.assertEqual(allocations[0]["inventory_after"][0]["quantity"], 3)
        self.assertEqual(summary["material_quantity"], 12)
        self.assertEqual(summary["finished_quantity"], 0)
        remaining = {item["variant_code"]: item["ending_quantity"] for item in stock}
        self.assertEqual(remaining, {"A.001": 3, "A.002": 0})
        self.assertTrue(all(item["opening_quantity"] == 0 for item in stock))

    def test_summarizes_same_purchase_code_with_weighted_average_cost(self):
        purchase = [
            line("purchase", "A.001", 2, 10, invoice_no="MUA1"),
            line("purchase", "A.001", 6, 20, invoice_no="MUA2"),
        ]
        sales = [line("sales", "A", 5, 30)]

        allocations, stock, _, _ = allocate_stock([], purchase, sales)

        used = allocations[0]["used"]
        self.assertEqual(len(used), 1)
        self.assertEqual(used[0]["variant_code"], "A.001")
        self.assertEqual(used[0]["quantity"], 5)
        self.assertEqual(used[0]["unit_cost"], 17.5)
        self.assertEqual(used[0]["summary_count"], 2)
        self.assertEqual(stock[0]["purchase_quantity"], 8)
        self.assertEqual(stock[0]["ending_quantity"], 3)

    def test_suffix_sales_rows_can_use_other_purchase_suffixes_under_same_base(self):
        purchase = [
            line("purchase", "A.001", 5, 90, invoice_no="MUA1"),
            line("purchase", "A.002", 10, 98, invoice_no="MUA2"),
        ]
        sales = [line("sales", "A.001", 6, 100)]

        allocations, stock, summary, _ = allocate_stock([], purchase, sales)

        self.assertEqual([item["variant_code"] for item in allocations[0]["used"]], ["A.002"])
        self.assertEqual(allocations[0]["material_quantity"], 6)
        self.assertEqual(allocations[0]["finished_quantity"], 0)
        remaining = {item["variant_code"]: item["ending_quantity"] for item in stock}
        self.assertEqual(remaining, {"A.001": 5, "A.002": 4})
        self.assertEqual(summary["material_quantity"], 6)

    def test_sends_shortage_to_finished_goods(self):
        opening = [line("opening", "B.001", 2, None)]
        purchase = [line("purchase", "B.002", 3, 60)]
        sales = [line("sales", "B", 8, 75)]

        allocations, _, summary, _ = allocate_stock(opening, purchase, sales)

        self.assertEqual(allocations[0]["material_quantity"], 5)
        self.assertEqual(allocations[0]["finished_quantity"], 3)
        self.assertEqual(summary["finished_quantity"], 3)

    def test_yen_thanh_shortage_exports_negative_in_ktp_not_khh(self):
        purchase = [line("purchase", "B.001", 3, 60)]
        sales = [line("sales", "B", 8, 75)]

        allocations, _, summary, _ = allocate_stock([], purchase, sales, {"allow_negative_export": True})
        ledger = build_inventory_ledger([], purchase, allocations)

        self.assertEqual(allocations[0]["material_quantity"], 3)
        self.assertEqual(allocations[0]["finished_quantity"], 5)
        self.assertEqual(summary["material_quantity"], 3)
        self.assertEqual(summary["finished_quantity"], 5)
        self.assertFalse(any(item.get("negative_export") for item in allocations[0]["used"]))
        ktp_group = ledger["warehouses"][1]["groups"][0]
        self.assertEqual(ktp_group["rows"][0]["type"], "sale")
        self.assertEqual(sum(row["qty_out"] for row in ktp_group["rows"]), 5)

    def test_future_purchase_does_not_cover_earlier_sale(self):
        purchase = [line(
            "purchase", "A.001", 500, 10, invoice_no="MUA42", invoice_date="30/06/2025",
            invoice_date_iso="2025-06-30",
        )]
        sales = [
            line(
                "sales", "A", 15, 20, invoice_no="BAN312", invoice_date="11/02/2025",
                invoice_date_iso="2025-02-11", row_number=3,
            ),
            line(
                "sales", "A", 20, 20, invoice_no="BAN2246", invoice_date="02/07/2025",
                invoice_date_iso="2025-07-02", row_number=4,
            ),
        ]

        allocations, _, summary, _ = allocate_stock([], purchase, sales)
        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)

        self.assertEqual(allocations[0]["material_quantity"], 0)
        self.assertEqual(allocations[0]["finished_quantity"], 15)
        self.assertEqual(allocations[1]["material_quantity"], 20)
        self.assertEqual(allocations[1]["finished_quantity"], 0)
        self.assertEqual(summary["material_quantity"], 20)
        self.assertEqual(summary["finished_quantity"], 15)
        khh_rows = ledger["warehouses"][0]["groups"][0]["rows"]
        ktp_rows = ledger["warehouses"][1]["groups"][0]["rows"]
        self.assertNotIn("11/02/2025", [row.get("date") for row in khh_rows if row["type"] == "sale"])
        self.assertEqual([row["date"] for row in ktp_rows], ["11/02/2025"])
        self.assertEqual(ktp_rows[0]["qty_in"], 0)
        self.assertEqual(ktp_rows[0]["qty_out"], 15)

    def test_yen_thanh_future_purchase_option_moves_purchase_before_sale(self):
        purchase = [line(
            "purchase", "A.001", 100, 10, invoice_no="MUA42", invoice_date="20/02/2025",
            invoice_date_iso="2025-02-20", row_number=42,
        )]
        sales = [line(
            "sales", "A", 30, 20, invoice_no="BAN312", invoice_date="11/02/2025",
            invoice_date_iso="2025-02-11", row_number=312,
        )]

        allocations, _, summary, _ = allocate_stock([], purchase, sales, {
            "company_profile": "yen_thanh",
            "allow_future_purchase_reorder": True,
            "future_purchase_window_days": 31,
        })
        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)
        khh_rows = ledger["warehouses"][0]["groups"][0]["rows"]

        self.assertEqual(allocations[0]["material_quantity"], 30)
        self.assertEqual(allocations[0]["finished_quantity"], 0)
        self.assertTrue(allocations[0]["used"][0]["future_purchase_reordered"])
        self.assertEqual(allocations[0]["used"][0]["future_reorder_days"], 9)
        self.assertEqual(summary["future_purchase_reorder_count"], 1)
        self.assertEqual(summary["future_purchase_reorder_quantity"], 30)
        self.assertEqual(khh_rows[0]["type"], "purchase_future_reorder")
        self.assertEqual(khh_rows[0]["date"], "11/02/2025")
        self.assertTrue(khh_rows[0]["future_purchase_reordered"])

    def test_yen_thanh_future_purchase_window_blocks_far_purchase(self):
        purchase = [line(
            "purchase", "A.001", 100, 10, invoice_no="MUA99", invoice_date="20/04/2025",
            invoice_date_iso="2025-04-20",
        )]
        sales = [line(
            "sales", "A", 30, 20, invoice_no="BAN312", invoice_date="11/02/2025",
            invoice_date_iso="2025-02-11",
        )]

        allocations, _, summary, _ = allocate_stock([], purchase, sales, {
            "company_profile": "yen_thanh",
            "allow_future_purchase_reorder": True,
            "future_purchase_window_days": 31,
        })

        self.assertEqual(allocations[0]["material_quantity"], 0)
        self.assertEqual(allocations[0]["finished_quantity"], 30)
        self.assertEqual(summary["future_purchase_reorder_count"], 0)

    def test_clears_fractional_rounding_remainder_when_stock_is_sufficient(self):
        purchase = [line("purchase", "R.001", 0.1, 90), line("purchase", "R.002", 0.2, 95)]
        sales = [line("sales", "R", 0.3, 100)]

        allocations, _, summary, _ = allocate_stock([], purchase, sales)

        self.assertEqual(allocations[0]["finished_quantity"], 0)
        self.assertEqual(summary["finished_quantity"], 0)

    def test_acceptance_range_excludes_excess_loss_and_excess_profit(self):
        purchase = [
            line("purchase", "A.001", 4, 120),
            line("purchase", "A.002", 4, 90),
            line("purchase", "A.003", 4, 60),
        ]
        sales = [line("sales", "A", 6, 100)]

        allocations, stock, summary, warnings = allocate_stock(
            [], purchase, sales, {"max_loss_percent": 10, "max_profit_percent": 25}
        )

        self.assertEqual([item["variant_code"] for item in allocations[0]["used"]], ["A.002"])
        self.assertEqual(allocations[0]["material_quantity"], 4)
        self.assertEqual(allocations[0]["finished_quantity"], 2)
        self.assertEqual({item["variant_code"] for item in allocations[0]["rejected"]}, {"A.001", "A.003"})
        self.assertEqual(summary["range_rejected_lines"], 1)
        self.assertTrue(warnings)
        remaining = {item["variant_code"]: item["ending_quantity"] for item in stock}
        self.assertEqual(remaining, {"A.001": 4, "A.002": 0, "A.003": 4})

    def test_loss_priority_uses_smallest_loss_before_larger_loss(self):
        purchase = [
            line("purchase", "A.001", 4, 105),
            line("purchase", "A.002", 4, 120),
        ]
        sales = [line("sales", "A", 5, 100)]

        allocations, _, _, _ = allocate_stock(
            [], purchase, sales, {"max_loss_percent": 30, "max_profit_percent": 25}
        )

        self.assertEqual([item["variant_code"] for item in allocations[0]["used"]], ["A.001", "A.002"])
        self.assertEqual([item["quantity"] for item in allocations[0]["used"]], [4, 1])

    def test_ktp_suffix_fallback_uses_best_purchase_suffix_even_when_rejected(self):
        purchase = [
            line("purchase", "A.001", 4, 105),
            line("purchase", "A.002", 4, 120),
        ]
        sales = [line("sales", "A", 3, 100)]

        allocations, _, _, _ = allocate_stock(
            [], purchase, sales, {"max_loss_percent": 1, "max_profit_percent": 25}
        )
        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)

        self.assertEqual(allocations[0]["material_quantity"], 0)
        self.assertEqual(allocations[0]["finished_quantity"], 3)
        self.assertEqual(allocations[0]["finished_variant_code"], "A.001")
        self.assertEqual(ledger["warehouses"][1]["groups"][0]["variant_code"], "A.001")

    def test_acceptance_range_rejects_missing_price_needed_for_validation(self):
        opening = [line("opening", "D.001", 3, None)]
        sales = [line("sales", "D", 2, 100)]

        allocations, _, _, _ = allocate_stock(
            opening, [], sales, {"max_loss_percent": 10, "max_profit_percent": 25}
        )

        self.assertEqual(allocations[0]["material_quantity"], 0)
        self.assertEqual(allocations[0]["finished_quantity"], 2)

    def test_lists_sales_codes_that_have_no_opening_or_purchase_stock(self):
        opening = [line("opening", "A.001", 1, 95)]
        purchase = [line("purchase", "B.002", 2, 90)]
        sales = [
            line("sales", "A", 1, 100, row_number=3),
            line("sales", "B", 2, 100, row_number=4),
            line("sales", "C", 3, 100, row_number=5),
            line("sales", "C", 4, 100, row_number=6),
        ]

        result = find_sale_only_codes(opening, purchase, sales)

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["variant_code"], "C")
        self.assertEqual(result[0]["quantity"], 7)
        self.assertEqual(result[0]["rows"], [5, 6])
        self.assertEqual(result[0]["opening_quantity"], 0)
        self.assertEqual(result[0]["purchase_quantity"], 0)

        _, stock_rows, _, _ = allocate_stock(opening, purchase, sales)
        missing_stock = next(row for row in stock_rows if row["variant_code"] == "C")
        self.assertEqual(missing_stock["opening_quantity"], 0)
        self.assertEqual(missing_stock["purchase_quantity"], 0)
        self.assertEqual(missing_stock["ending_quantity"], 0)

    def test_output_preserves_sales_sheet_and_adds_two_columns(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Invoices"
        sheet.cell(2, 13, "Ma VT")
        sheet.cell(2, 16, "So luong")
        sheet.cell(3, 13, "A")
        sheet.cell(3, 16, 4)
        content = BytesIO()
        workbook.save(content)
        allocation = line("sales", "A", 4, 100)
        allocation.update({"material_quantity": 3, "finished_quantity": 1, "used": []})
        mapping = {"sheet": "", "header_row": 2, "data_start_row": 3}

        result = create_output_workbook(content.getvalue(), mapping, [allocation], [], {
            "opening_quantity": 0,
            "purchase_quantity": 3,
            "sales_quantity": 4,
            "material_quantity": 3,
            "finished_quantity": 1,
            "material_percent": 75,
        })
        output = load_workbook(result)
        original = output["Invoices"]

        self.assertEqual(original.cell(2, 17).value, "SL khớp từ mua vào")
        self.assertEqual(original.cell(2, 18).value, "SL xuất theo kho bán ra")
        self.assertEqual(original.cell(3, 17).value, 3)
        self.assertEqual(original.cell(3, 18).value, 1)
        self.assertEqual(original.cell(2, 19).value, "Tồn mua vào trước khi bán")
        self.assertEqual(original.cell(2, 21).value, "Tồn mua vào sau khi bán")
        self.assertNotIn("PhanBoKho", output.sheetnames)
        self.assertNotIn("TonKhoHangHoa", output.sheetnames)
        self.assertIn("MaChiBanRaKhongTon", output.sheetnames)
        self.assertIn("MauMuaVao", output.sheetnames)
        self.assertIn("MauBanRa", output.sheetnames)

    def test_output_adds_purchase_and_sales_template_sheets(self):
        purchase = [
            line(
                "purchase", "A.001", 2, 70, invoice_no="MUA1", invoice_date="01/01/2025",
                party_tax_code="NCC001", party_name="Nha cung cap 1", line_amount=140,
            ),
            line(
                "purchase", "A.002", 3, 90, invoice_no="MUA2", invoice_date="02/01/2025",
                party_tax_code="NCC002", party_name="Nha cung cap 2", line_amount=270,
            ),
        ]
        sales = [line(
            "sales", "A", 6, 120, invoice_no="BAN1", invoice_date="03/01/2025",
            party_tax_code="KH001", party_name="Khach hang 1", line_amount=720,
            tax_rate_percent=10, tax_amount=72,
        )]
        allocations, stock, summary, _ = allocate_stock([], purchase, sales)

        result = create_output_workbook(
            invoice_file("A", 6, 120).getvalue(),
            clean_mapping({}, "sales"),
            allocations,
            stock,
            summary,
            purchase_lines=purchase,
            sales_lines=sales,
        )
        output = load_workbook(result)
        purchase_sheet = output["MauMuaVao"]
        sales_sheet = output["MauBanRa"]

        self.assertEqual([purchase_sheet.cell(1, column).value for column in range(1, 10)], [
            "MST", "Tên công ty", "Số HĐ", "Ngày HĐ", "Mã kho", "Mã VT", "Số lượng", "Đơn giá mua", "Thành tiền mua",
        ])
        self.assertEqual(purchase_sheet.cell(2, 5).value, "KHH")
        self.assertEqual(purchase_sheet.cell(2, 6).value, "A.001")
        self.assertEqual(purchase_sheet.cell(3, 9).value, 270)

        self.assertEqual(sales_sheet.cell(2, 12).value, "KHH")
        self.assertEqual(sales_sheet.cell(2, 13).value, "A.002")
        self.assertEqual(sales_sheet.cell(2, 5).value, 3)
        self.assertEqual(sales_sheet.cell(2, 7).value, 360)
        self.assertEqual(sales_sheet.cell(2, 8).value, 90)
        self.assertEqual(sales_sheet.cell(2, 11).value, 36)
        self.assertEqual(sales_sheet.cell(3, 12).value, "KHH")
        self.assertEqual(sales_sheet.cell(3, 13).value, "A.001")
        self.assertEqual(sales_sheet.cell(4, 12).value, "KTP")
        self.assertEqual(sales_sheet.cell(4, 13).value, "A.002")
        self.assertEqual(sales_sheet.cell(4, 8).value, 0)
        self.assertEqual(sales_sheet.cell(4, 9).value, 0)
        self.assertEqual(sales_sheet.cell(4, 11).value, 12)

    def test_inventory_ledger_splits_khh_and_ktp_rows(self):
        purchase = [line(
            "purchase", "A.001", 10, 5, invoice_no="MUA1", invoice_date="01/01/2025",
            invoice_date_iso="2025-01-01", party_name="Nha cung cap", unit_name="Cai", line_amount=50,
        )]
        sales = [line(
            "sales", "A", 12, 20, invoice_no="BAN1", invoice_date="02/01/2025",
            invoice_date_iso="2025-01-02", party_name="Khach hang", unit_name="Cai", line_amount=80,
        )]
        allocations, _, _, _ = allocate_stock([], purchase, sales)

        ledger = build_inventory_ledger([], purchase, allocations)
        khh = ledger["warehouses"][0]
        ktp = ledger["warehouses"][1]
        khh_group = next(group for group in khh["groups"] if group["variant_code"] == "A.001")
        ktp_group = ktp["groups"][0]
        rows = khh_group["rows"]

        self.assertEqual(khh["warehouse_code"], "KHH")
        self.assertEqual(ktp["warehouse_code"], "KTP")
        self.assertEqual(khh_group["variant_code"], "A.001")
        self.assertEqual(khh_group["unit_name"], "Cai")
        self.assertEqual(rows[0]["type"], "purchase")
        self.assertEqual(rows[0]["date"], "01/01/2025")
        self.assertEqual(rows[0]["qty_in"], 10)
        self.assertEqual(rows[0]["amount_in"], 50)
        self.assertEqual(rows[1]["type"], "sale")
        self.assertEqual(rows[1]["date"], "02/01/2025")
        self.assertEqual(rows[1]["account"], "6321")
        self.assertEqual(rows[1]["variant_code"], "A.001")
        self.assertEqual(rows[1]["qty_out"], 10)
        self.assertEqual(rows[1]["amount_out"], 50)
        self.assertEqual(ktp_group["variant_code"], "A.001")
        self.assertEqual(ktp_group["rows"][0]["type"], "sale")
        self.assertEqual(ktp_group["rows"][0]["qty_in"], 0)
        self.assertEqual(ktp_group["rows"][0]["qty_out"], 2)
        self.assertEqual(ktp_group["rows"][0]["unit_price"], 0)
        self.assertEqual(ktp_group["rows"][0]["amount_in"], 0)
        self.assertEqual(ktp_group["rows"][0]["amount_out"], 0)
        self.assertEqual(ktp_group["rows"][0]["account"], "6321")

    def test_inventory_ledger_routes_matched_purchase_into_sales_warehouse_for_cost(self):
        purchase = [line(
            "purchase", "PAPER", 10, 50, invoice_no="MUA1", invoice_date="01/01/2025",
            invoice_date_iso="2025-01-01", party_name="Nha cung cap", unit_name="Kg",
            line_amount=500, warehouse_code="KTP", warehouse_account="1551",
        )]
        sales = [line(
            "sales", "PAPER", 3, 120, invoice_no="BAN1", invoice_date="02/01/2025",
            invoice_date_iso="2025-01-02", party_name="Khach hang", unit_name="Kg",
            line_amount=360, warehouse_code="KHH", warehouse_account="152",
        )]
        allocations, _, _, _ = allocate_stock([], purchase, sales)

        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)
        khh = next(warehouse for warehouse in ledger["warehouses"] if warehouse["warehouse_code"] == "KHH")
        ktp = next(warehouse for warehouse in ledger["warehouses"] if warehouse["warehouse_code"] == "KTP")
        khh_group = khh["groups"][0]
        ktp_group = ktp["groups"][0]

        self.assertEqual(khh["account"], "152")
        self.assertEqual(sum(row["qty_in"] for row in khh_group["rows"]), 3)
        self.assertEqual(sum(row["qty_out"] for row in khh_group["rows"]), 3)
        sale_row = next(row for row in khh_group["rows"] if row["type"] == "sale")
        self.assertFalse(sale_row["cost_missing"])
        self.assertEqual(sale_row["unit_price"], 50)
        self.assertEqual(sale_row["amount_out"], 150)
        self.assertEqual(sum(row["qty_in"] for row in ktp_group["rows"]), 7)

    def test_inventory_ledger_moves_unmatched_khh_remainder_to_ktp(self):
        purchase = [line(
            "purchase", "PAPER", 10, 50, invoice_no="MUA1", invoice_date="05/01/2025",
            invoice_date_iso="2025-01-05", party_name="Nha cung cap", unit_name="Kg",
            line_amount=500, warehouse_code="KTP", warehouse_account="1551",
        )]
        sales = [line(
            "sales", "PAPER", 3, 120, invoice_no="BAN1", invoice_date="02/01/2025",
            invoice_date_iso="2025-01-02", party_name="Khach hang", unit_name="Kg",
            line_amount=360, warehouse_code="KHH", warehouse_account="152",
        )]
        allocations, _, _, _ = allocate_stock([], purchase, sales)

        ledger = build_inventory_ledger([], purchase, allocations, sales_lines=sales)
        warehouse_codes = [warehouse["warehouse_code"] for warehouse in ledger["warehouses"]]
        ktp = next(warehouse for warehouse in ledger["warehouses"] if warehouse["warehouse_code"] == "KTP")
        sale_row = next(row for group in ktp["groups"] for row in group["rows"] if row["type"] == "sale")

        self.assertNotIn("KHH", warehouse_codes)
        self.assertEqual(sale_row["qty_out"], 3)
        self.assertEqual(sale_row["unit_price"], 0)
        self.assertTrue(sale_row["cost_missing"])

    def test_inventory_ledger_puts_full_shortage_only_in_ktp(self):
        sales = [line(
            "sales", "NO_STOCK", 5, 20, invoice_no="BAN2", invoice_date="03/01/2025",
            invoice_date_iso="2025-01-03", party_name="Khach hang", unit_name="Cai", line_amount=100,
        )]
        allocations, _, _, _ = allocate_stock([], [], sales)

        ledger = build_inventory_ledger([], [], allocations)
        ktp = ledger["warehouses"][0]

        self.assertEqual(ktp["groups"][0]["variant_code"], "NO_STOCK")
        self.assertEqual([row["type"] for row in ktp["groups"][0]["rows"]], ["sale"])
        self.assertEqual(ktp["groups"][0]["rows"][0]["qty_in"], 0)
        self.assertEqual(ktp["groups"][0]["rows"][0]["qty_out"], 5)

    def test_sales_warehouse_from_processed_sales_drives_export_only_rows(self):
        sales = [line(
            "sales", "GC_ONLY", 7, 120, invoice_no="BAN-GC", invoice_date="04/01/2025",
            invoice_date_iso="2025-01-04", party_name="Khach gia cong", unit_name="Cai",
            line_amount=840, warehouse_code="KGCI", warehouse_account="1552",
        )]
        allocations, _, summary, _ = allocate_stock([], [], sales)
        ledger = build_inventory_ledger([], [], allocations, sales_lines=sales)
        report_rows = inventory_app.build_sales_report_rows(allocations, [], sales)

        self.assertEqual(summary["material_quantity"], 0)
        self.assertEqual(summary["finished_quantity"], 7)
        self.assertEqual([warehouse["warehouse_code"] for warehouse in ledger["warehouses"]], ["KGCI"])
        self.assertEqual(ledger["warehouses"][0]["account"], "1552")
        rows = ledger["warehouses"][0]["groups"][0]["rows"]
        self.assertEqual([row["type"] for row in rows], ["sale"])
        self.assertEqual(sum(row["qty_in"] for row in rows), 0)
        self.assertEqual(sum(row["qty_out"] for row in rows), 7)
        self.assertEqual(report_rows[0]["warehouse_code"], "KGCI")
        self.assertEqual(report_rows[0]["cost_amount"], 0)

    def test_ktp_uses_zero_cost_when_no_purchase_stock_exists(self):
        sales = [
            line(
                "sales", "KTP_ONLY", 2, 20, invoice_no="BAN1", invoice_date="03/01/2025",
                invoice_date_iso="2025-01-03", party_name="Khach hang", unit_name="Cai", line_amount=40,
                row_number=3,
            ),
            line(
                "sales", "KTP_ONLY", 6, 30, invoice_no="BAN2", invoice_date="04/01/2025",
                invoice_date_iso="2025-01-04", party_name="Khach hang", unit_name="Cai", line_amount=180,
                row_number=4,
            ),
        ]
        allocations, _, _, _ = allocate_stock([], [], sales)

        ledger = build_inventory_ledger([], [], allocations, sales_lines=sales)
        rows = ledger["warehouses"][0]["groups"][0]["rows"]

        self.assertEqual([row["type"] for row in rows], ["sale", "sale"])
        self.assertTrue(all(row["unit_price"] == 0 for row in rows))
        self.assertEqual(sum(row["qty_in"] for row in rows), 0)
        self.assertEqual(sum(row["qty_out"] for row in rows), 8)

    def test_output_reuses_existing_result_columns(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Invoices"
        sheet.cell(2, 13, "Ma VT")
        sheet.cell(2, 16, "So luong")
        sheet.cell(2, 17, "SL lấy từ kho hàng hóa")
        sheet.cell(2, 18, "SL lấy từ kho thành phẩm")
        content = BytesIO()
        workbook.save(content)
        allocation = line("sales", "A", 2, 100)
        allocation.update({"material_quantity": 2, "finished_quantity": 0, "used": []})
        result = create_output_workbook(
            content.getvalue(),
            {"sheet": "", "header_row": 2, "data_start_row": 3},
            [allocation],
            [],
            {
                "opening_quantity": 0, "purchase_quantity": 2, "sales_quantity": 2,
                "material_quantity": 2, "finished_quantity": 0, "material_percent": 100,
            },
        )
        output = load_workbook(result)

        self.assertEqual(output["Invoices"].max_column, 22)
        self.assertEqual(output["Invoices"].cell(3, 17).value, 2)

    def test_api_generates_downloadable_workbook(self):
        response = inventory_app.app.test_client().post(
            "/api/analyze",
            data={
                "purchase_file": (invoice_file("C.001", 3, 95), "mua.xlsx"),
                "sales_file": (invoice_file("C", 5, 100), "ban.xlsx"),
                "mapping": "{}",
                "policy": '{"max_loss_percent": 10, "max_profit_percent": 25}',
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["summary"]["material_quantity"], 3)
        self.assertEqual(payload["summary"]["finished_quantity"], 2)
        self.assertEqual(payload["sale_only_codes"], [])
        self.assertEqual(payload["allocations"][0]["used"][0]["variant_code"], "C.001")
        self.assertEqual(payload["allocations"][0]["used"][0]["sale_variant_code"], "C")
        self.assertEqual(payload["allocations"][0]["used"][0]["quantity"], 3)
        self.assertEqual(payload["allocations"][0]["invoice_no"], "HD1")
        self.assertEqual(payload["allocations"][0]["invoice_date"], "02/01/2025")
        self.assertEqual(payload["allocations"][0]["used"][0]["invoice_no"], "HD1")
        self.assertEqual(payload["allocations"][0]["used"][0]["invoice_date"], "02/01/2025")
        self.assertEqual(payload["allocations"][0]["sale_split_codes"], "C.001")
        self.assertNotIn("allocation_summaries", payload)
        self.assertTrue(payload["verification"])
        self.assertTrue(all(row["status"] == "OK" for row in payload["verification"]))

        download = inventory_app.app.test_client().get(f"/api/download/{payload['job_id']}")
        self.assertEqual(download.status_code, 200)
        result = load_workbook(BytesIO(download.data))
        self.assertIn("TongHopKho", result.sheetnames)
        self.assertIn("KiemTraDoiChieu", result.sheetnames)
        self.assertNotIn("TongHopNhapXuatTon", result.sheetnames)
        self.assertIn("TongHopNXT_KHH", result.sheetnames)
        self.assertIn("TongHopNXT_KTP", result.sheetnames)
        self.assertIn("BaoCaoBH_KHH", result.sheetnames)
        self.assertIn("BaoCaoBH_KTP", result.sheetnames)
        self.assertIn("BangKeHDBH_KHH", result.sheetnames)
        self.assertIn("BangKeHDBH_KTP", result.sheetnames)
        self.assertNotIn("PhanBoKho", result.sheetnames)
        self.assertNotIn("TonKhoHangHoa", result.sheetnames)
        self.assertIn("SoChiTietKHH", result.sheetnames)
        self.assertIn("SoChiTietKTP", result.sheetnames)
        self.assertIn("SoChiTietHH_TP", result.sheetnames)
        self.assertEqual(result["SoChiTietKHH"].cell(1, 1).value, "Ngày")
        self.assertEqual(result["SoChiTietHH_TP"].cell(1, 7).value, "Mã VT chi tiết")
        self.assertEqual(result["TongHopNXT_KHH"].cell(1, 1).value, "TỔNG HỢP NHẬP XUẤT TỒN")
        self.assertEqual(result["TongHopNXT_KTP"].cell(1, 1).value, "TỔNG HỢP NHẬP XUẤT TỒN")
        self.assertEqual(result["TongHopNXT_KHH"].cell(2, 1).value, "KHO: KHH - KHO VẬT TƯ, HÀNG HÓA")
        self.assertEqual(result["TongHopNXT_KTP"].cell(2, 1).value, "KHO: KTP - KHO THÀNH PHẨM")
        self.assertEqual(result["BaoCaoBH_KHH"].cell(1, 1).value, "BÁO CÁO TỔNG HỢP BÁN HÀNG")
        self.assertEqual(result["BaoCaoBH_KHH"].cell(5, 8).value, "TIỀN LÃI/LỖ")
        self.assertEqual(result["BaoCaoBH_KHH"].cell(5, 9).value, "% LÃI/LỖ")
        self.assertEqual(result["BaoCaoBH_KHH"].cell(7, 2).value, "Tổng cộng:")
        self.assertEqual(result["BangKeHDBH_KHH"].cell(1, 1).value, "BẢNG KÊ HOÁ ĐƠN BÁN HÀNG")
        verification_statuses = [
            result["KiemTraDoiChieu"].cell(row, 7).value
            for row in range(2, result["KiemTraDoiChieu"].max_row + 1)
        ]
        self.assertTrue(verification_statuses)
        self.assertTrue(all(status == "OK" for status in verification_statuses))
        download.close()

    def test_async_analyze_job_reports_progress_and_result(self):
        client = inventory_app.app.test_client()
        response = client.post(
            "/api/analyze-job",
            data={
                "purchase_file": (invoice_file("D.001", 3, 90), "mua.xlsx"),
                "sales_file": (invoice_file("D", 4, 120), "ban.xlsx"),
                "mapping": "{}",
                "policy": "{}",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        analysis_job_id = response.get_json()["analysis_job_id"]

        status = None
        for _ in range(30):
            poll = client.get(f"/api/analyze-job/{analysis_job_id}")
            self.assertEqual(poll.status_code, 200)
            status = poll.get_json()
            self.assertIn("progress", status)
            if status["status"] == "complete":
                break
            time.sleep(0.1)

        self.assertEqual(status["status"], "complete")
        self.assertEqual(status["progress"], 100)
        self.assertEqual(status["result"]["summary"]["material_quantity"], 3)
        self.assertEqual(status["result"]["summary"]["finished_quantity"], 1)

    def test_new_application_uses_separate_port(self):
        self.assertEqual(inventory_app.APP_PORT, 5082)

    def test_preview_returns_excel_letters_with_selected_header_labels(self):
        preview = preview_workbook(invoice_file("Z", 2, 100).getvalue(), header_row=2)

        columns = {column["letter"]: column["header"] for column in preview["columns"]}
        self.assertEqual(preview["header_row"], 2)
        self.assertEqual(columns["L"], "Ma VT")
        self.assertEqual(columns["D"], "Ngay hoa don")
        self.assertEqual(columns["O"], "So luong")
        self.assertEqual(columns["P"], "Don gia")

    def test_read_lines_auto_detects_up_sales_layout(self):
        rows = read_lines(up_sales_file("HP.OTONMC13X26X1.2X6.0", 2142, 18590.91).getvalue(), clean_mapping({}, "sales"), "sales", company_profile="son_phuong")[1]

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["variant_code"], "HP.OTONMC13X26X1.2X6.0")
        self.assertEqual(rows[0]["quantity"], 2142)
        self.assertEqual(rows[0]["unit_price"], 18590.91)
        self.assertEqual(rows[0]["invoice_no"], "HDSP1")
        self.assertEqual(rows[0]["invoice_date_iso"], "2026-01-02")

    def test_read_lines_normalizes_dates_to_display_and_iso(self):
        rows = read_lines(invoice_file("Z", 2, 100).getvalue(), clean_mapping({}, "sales"), "sales")[1]

        self.assertEqual(rows[0]["invoice_date"], "02/01/2025")
        self.assertEqual(rows[0]["invoice_date_iso"], "2025-01-02")

    def test_son_phuong_read_lines_strips_company_prefix_from_product_code(self):
        purchase = read_lines(
            invoice_file("HA.OTMT21X2X1X6.0", 2, 100).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1]
        sales = read_lines(
            invoice_file("CD.THMK150X150X2.5X6000", 3, 120).getvalue(),
            clean_mapping({}, "sales"),
            "sales",
            company_profile="son_phuong",
        )[1]

        self.assertEqual(purchase[0]["original_variant_code"], "HA.OTMT21X2X1X6.0")
        self.assertEqual(purchase[0]["variant_code"], "HA.OTMT21X2X1X6.0")
        self.assertEqual(sales[0]["original_variant_code"], "CD.THMK150X150X2.5X6000")
        self.assertEqual(sales[0]["variant_code"], "CD.THMK150X150X2.5X6000")

    def test_son_phuong_read_lines_normalizes_steel_profile_across_company_codes(self):
        purchase_a = read_lines(
            invoice_file_with_name("CD.THMK150X150X2.5X6000", "Thep hop ma kem 150x150x2.5x6000", 5, 18000).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1][0]
        purchase_b = read_lines(
            invoice_file_with_name("TPHY.THMK150X150X2.5X6000", "Hop ma kem 150x150x2.5x6000", 7, 18100).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1][0]
        pipe = read_lines(
            invoice_file_with_name("HA.OTMT212X21X6.0", "Ong thep ma tron 21,2 x 2,1 x 6.0", 3, 22000).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1][0]

        self.assertEqual(purchase_a["variant_code"], "CD.HOP.MK.150X150X2.5")
        self.assertEqual(purchase_b["variant_code"], "TPHY.HOP.MK.150X150X2.5")
        self.assertEqual(purchase_a["source_variant_code"], "THMK150X150X2.5X6000")
        self.assertEqual(pipe["variant_code"], "HA.ONG.MK.21.2X2.1")
        self.assertEqual(pipe["steel_profile_key"], "pipe|galvanized|21.2x2.1")

    def test_son_phuong_exact_sale_matches_purchase_by_normalized_profile(self):
        purchase = read_lines(
            invoice_file_with_name("CD.THMK150X150X2.5X6000", "Thep hop ma kem 150x150x2.5x6000", 5, 10000).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1]
        sales = read_lines(
            invoice_file_with_name("TP.THMK150X150X2.5X6000", "Hop ma kem 150x150x2.5x6000", 3, 12000).getvalue(),
            clean_mapping({}, "sales"),
            "sales",
            company_profile="son_phuong",
        )[1]

        allocations, _, summary, _ = allocate_stock(
            [], purchase, sales, {"company_profile": "son_phuong"}, barem_map=DEFAULT_BAREM_MAP
        )

        self.assertEqual(allocations[0]["used"][0]["variant_code"], "CD.HOP.MK.150X150X2.5")
        self.assertEqual(summary["negative_export_quantity"], 0)
        self.assertEqual(summary["material_quantity"], 3)

    def test_son_phuong_non_steel_sale_matches_purchase_by_normalized_product_family(self):
        purchase = read_lines(
            invoice_file_with_name("TH.BANMA14", "Thep ban ma 14mm", 20, 9000).getvalue(),
            clean_mapping({}, "purchase"),
            "purchase",
            company_profile="son_phuong",
        )[1]
        sales = read_lines(
            invoice_file_with_name("BANM14MM", "Thep ban ma 14mm", 5, 12000).getvalue(),
            clean_mapping({}, "sales"),
            "sales",
            company_profile="son_phuong",
        )[1]

        allocations, _, summary, _ = allocate_stock(
            [], purchase, sales, {"company_profile": "son_phuong"}, barem_map=DEFAULT_BAREM_MAP
        )
        report_rows = inventory_app.build_sales_report_rows(allocations, purchase, sales)

        self.assertEqual(allocations[0]["used"][0]["variant_code"], "TH.BANMA14")
        self.assertEqual(summary["negative_export_quantity"], 0)
        self.assertEqual(report_rows[0]["cost_amount"], 45000)
        self.assertEqual(report_rows[0]["profit_amount"], 15000)

    def test_ignores_placeholder_zero_code_rows_from_old_formatter_output(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.cell(2, 12, "Ma VT")
        sheet.cell(2, 15, "So luong")
        sheet.cell(3, 12, 0)
        sheet.cell(3, 15, 5)
        sheet.cell(4, 12, "A.001")
        sheet.cell(4, 15, 3)
        content = BytesIO()
        workbook.save(content)

        _, rows = read_lines(content.getvalue(), clean_mapping({}, "purchase"), "purchase")

        self.assertEqual([(row["variant_code"], row["quantity"]) for row in rows], [("A.001", 3.0)])

    def test_normalizes_vietnamese_item_codes_like_formatter_before_matching(self):
        purchase = invoice_file("Ống.001", 3, 95)
        sales = invoice_file("ống", 2, 100)
        purchase_rows = read_lines(purchase.getvalue(), clean_mapping({}, "purchase"), "purchase")[1]
        sales_rows = read_lines(sales.getvalue(), clean_mapping({}, "sales"), "sales")[1]

        allocations, _, _, _ = allocate_stock([], purchase_rows, sales_rows)

        self.assertEqual(purchase_rows[0]["variant_code"], "ONG.001")
        self.assertEqual(sales_rows[0]["variant_code"], "ONG")
        self.assertEqual(allocations[0]["used"][0]["variant_code"], "ONG.001")


if __name__ == "__main__":
    unittest.main()
