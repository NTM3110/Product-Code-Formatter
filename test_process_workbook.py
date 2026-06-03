import unittest
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook, load_workbook

from app import process_workbook


class ProcessWorkbookTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
