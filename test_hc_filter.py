from __future__ import annotations

import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import extract_orders
from extract_orders import is_hc_excel_file, run_extract


ORDER_HEADER = "\u4e9a\u9a6c\u900a\u8ba2\u5355\u53f7"
QUANTITY_HEADER = "\u6570\u91cf"


def make_order_workbook(path: Path, order_id: str, sku: str, quantity: int) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append([ORDER_HEADER, "SKU", QUANTITY_HEADER])
    worksheet.append([order_id, sku, quantity])
    workbook.save(path)
    workbook.close()


def make_zip(zip_path: Path, files: list[Path]) -> None:
    with zipfile.ZipFile(zip_path, "w") as archive:
        for file_path in files:
            archive.write(file_path, file_path.name)


class HcFilterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.base = Path(tempfile.mkdtemp(prefix="hc_filter_test_"))
        self.input_dir = self.base / "input"
        self.input_dir.mkdir()
        self.logs_dir = self.base / "logs"
        self.backups_dir = self.base / "backups"
        self.output_path = self.base / "summary.xlsx"

    def tearDown(self) -> None:
        shutil.rmtree(self.base, ignore_errors=True)

    def run_tool(self, *, dry_run: bool = False, enable_hc_filter: bool = False, input_mode: str = "archives") -> dict:
        return run_extract(
            str(self.input_dir),
            str(self.output_path),
            clear=True,
            dry_run=dry_run,
            workers=1,
            report_dir=str(self.logs_dir),
            backup_dir=str(self.backups_dir),
            enable_hc_filter=enable_hc_filter,
            input_mode=input_mode,
        )

    def test_hc_detection_uses_excel_file_name_only(self) -> None:
        for filename in ("HC.xlsx", "hc.xlsx", "Hc.xlsx", "hC.xlsx", "0507-HC-order.xlsx"):
            with self.subTest(filename=filename):
                self.assertTrue(is_hc_excel_file(Path(filename)))

        self.assertFalse(is_hc_excel_file(Path("HC-parent") / "normal_order.xlsx"))

    def test_hc_filter_enabled_copies_and_excludes_from_output(self) -> None:
        normal = self.base / "normal_order.xlsx"
        hc_file = self.base / "HC.xlsx"
        make_order_workbook(normal, "ORDER-NORMAL", "SKU-NORMAL", 2)
        make_order_workbook(hc_file, "ORDER-HC", "SKU-HC", 99)
        make_zip(self.input_dir / "orders.zip", [normal, hc_file])

        result = self.run_tool(enable_hc_filter=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 1)
        self.assertTrue(self.output_path.exists())
        self.assertTrue((self.input_dir / "HC" / "HC.xlsx").exists())
        self.assertEqual(result["stats"]["hc_file_count"], 1)
        self.assertEqual(result["stats"]["hc_copy_failed_count"], 0)
        self.assertIn("hc_file_count", result["stats"])
        self.assertIn("total_archives", result["stats"])
        self.assertEqual(len(result["hc_report_rows"]), 1)
        self.assertEqual(result["rows"][0]["亚马逊订单号"], "ORDER-NORMAL")

    def test_hc_filter_default_off_treats_hc_as_normal_excel_and_skips_multi_excel_unit(self) -> None:
        normal = self.base / "normal_order.xlsx"
        hc_file = self.base / "HC.xlsx"
        make_order_workbook(normal, "ORDER-NORMAL", "SKU-NORMAL", 2)
        make_order_workbook(hc_file, "ORDER-HC", "SKU-HC", 99)
        make_zip(self.input_dir / "orders.zip", [normal, hc_file])

        result = self.run_tool()

        self.assertFalse((self.input_dir / "HC").exists())
        self.assertEqual(result["stats"]["hc_file_count"], 0)
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertGreaterEqual(result["skipped_archives"], 1)
        skip_details = "\n".join(
            str(item) for item in result.get("exception_logs", []) + result.get("error_report_rows", [])
        )
        self.assertIn("多个正式 Excel", skip_details)
        self.assertIn("normal_order.xlsx", skip_details)
        self.assertIn("HC.xlsx", skip_details)

    def test_dry_run_does_not_create_hc_folder_or_output(self) -> None:
        normal = self.base / "normal_order.xlsx"
        hc_file = self.base / "hc.xlsx"
        make_order_workbook(normal, "ORDER-NORMAL", "SKU-NORMAL", 2)
        make_order_workbook(hc_file, "ORDER-HC", "SKU-HC", 99)
        make_zip(self.input_dir / "orders.zip", [normal, hc_file])

        result = self.run_tool(dry_run=True, enable_hc_filter=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertFalse((self.input_dir / "HC").exists())
        self.assertEqual(result["stats"]["hc_file_count"], 1)
        self.assertEqual(result["stats"]["hc_copy_failed_count"], 0)
        self.assertEqual(result["hc_report_rows"][0]["处理状态"], "dry-run 预计复制")

    def test_hc_copy_failure_still_excludes_hc_from_extraction(self) -> None:
        normal = self.base / "normal_order.xlsx"
        hc_file = self.base / "Hc.xlsx"
        make_order_workbook(normal, "ORDER-NORMAL", "SKU-NORMAL", 2)
        make_order_workbook(hc_file, "ORDER-HC", "SKU-HC", 99)
        make_zip(self.input_dir / "orders.zip", [normal, hc_file])
        original_copy2 = extract_orders.shutil.copy2

        def failing_copy(source: Path, target: Path):
            if Path(source).name == "Hc.xlsx":
                raise PermissionError("locked")
            return original_copy2(source, target)

        with patch("extract_orders.shutil.copy2", side_effect=failing_copy):
            result = self.run_tool(enable_hc_filter=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 1)
        self.assertEqual(result["stats"]["hc_file_count"], 1)
        self.assertEqual(result["stats"]["hc_copy_failed_count"], 1)
        self.assertEqual(result["rows"][0]["亚马逊订单号"], "ORDER-NORMAL")
        self.assertEqual(result["hc_report_rows"][0]["处理状态"], "复制失败，已排除")


if __name__ == "__main__":
    unittest.main()
