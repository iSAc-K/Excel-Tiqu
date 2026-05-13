from __future__ import annotations

import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

from extract_orders import run_extract


ORDER_HEADER = "\u4e9a\u9a6c\u900a\u8ba2\u5355\u53f7"
QUANTITY_HEADER = "\u6570\u91cf"
DATE_HEADER = "\u65e5\u671f"
SKIP_FOLDER_NAME = "\u672a\u5904\u7406\u538b\u7f29\u5305"
MODIFY_PREFIX = "\u4fee\u6539"


def make_order_workbook(path: Path, rows: list[tuple[str, str, object]], *, sheet_name: str = "Sheet1") -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    worksheet.append([ORDER_HEADER, "SKU", QUANTITY_HEADER])
    for order_id, sku, quantity in rows:
        worksheet.append([order_id, sku, quantity])
    workbook.save(path)
    workbook.close()


def make_zip(zip_path: Path, entries: list[tuple[Path, str]]) -> None:
    with zipfile.ZipFile(zip_path, "w") as archive:
        for source, archive_name in entries:
            archive.write(source, archive_name)


def read_summary_rows(output_path: Path) -> list[dict[str, object]]:
    workbook = load_workbook(output_path, data_only=True)
    rows: list[dict[str, object]] = []
    try:
        for worksheet in workbook.worksheets:
            headers = [worksheet.cell(row=1, column=col).value for col in range(1, 5)]
            if headers != [ORDER_HEADER, "SKU", QUANTITY_HEADER, DATE_HEADER]:
                continue
            for row_index in range(2, worksheet.max_row + 1):
                values = [worksheet.cell(row=row_index, column=col).value for col in range(1, 5)]
                if not any(value not in (None, "") for value in values):
                    continue
                rows.append(
                    {
                        "sheet": worksheet.title,
                        "order_id": values[0],
                        "sku": values[1],
                        "quantity": values[2],
                        "date": values[3],
                    }
                )
    finally:
        workbook.close()
    return rows


class CoreRegressionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.base = Path(tempfile.mkdtemp(prefix="core_regression_test_"))
        self.input_dir = self.base / "input"
        self.input_dir.mkdir()
        self.work_dir = self.base / "work"
        self.work_dir.mkdir()
        self.logs_dir = self.base / "logs"
        self.backups_dir = self.base / "backups"
        self.output_path = self.base / "summary.xlsx"

    def tearDown(self) -> None:
        shutil.rmtree(self.base, ignore_errors=True)

    def run_tool(
        self,
        *,
        dry_run: bool = False,
        skip_exact_duplicates: bool = True,
        clear: bool = True,
    ) -> dict:
        return run_extract(
            str(self.input_dir),
            str(self.output_path),
            clear=clear,
            dry_run=dry_run,
            workers=1,
            report_dir=str(self.logs_dir),
            backup_dir=str(self.backups_dir),
            skip_exact_duplicates=skip_exact_duplicates,
        )

    def test_fixture_helpers_create_readable_order_zip(self) -> None:
        workbook_path = self.work_dir / "fixture.xlsx"
        make_order_workbook(workbook_path, [("ORDER-FIXTURE", "SKU-FIXTURE", 1)])
        zip_path = self.input_dir / "fixture.zip"

        make_zip(zip_path, [(workbook_path, workbook_path.name)])

        self.assertTrue(workbook_path.exists())
        self.assertTrue(zip_path.exists())
        with zipfile.ZipFile(zip_path) as archive:
            self.assertEqual(archive.namelist(), [workbook_path.name])

    def test_dry_run_scans_and_reports_without_writing_output_or_backup(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-2pcs.xlsx"
        make_order_workbook(workbook_path, [("ORDER-DRY", "SKU-DRY", 2)])
        make_zip(self.input_dir / "dry_run.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(dry_run=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertFalse(self.backups_dir.exists())
        self.assertTrue(Path(result["process_report_path"]).exists())
        self.assertTrue(Path(result["log_file_path"]).exists())
        self.assertTrue(result["stats"]["dry_run"])

    def test_same_order_same_sku_merges_quantities_before_write(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-5pcs.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-MERGE", "SKU-MERGE", 2),
                ("ORDER-MERGE", "SKU-MERGE", 3),
            ],
        )
        make_zip(self.input_dir / "merge.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 2)
        self.assertEqual(result["written_rows"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["order_id"], "ORDER-MERGE")
        self.assertEqual(rows[0]["sku"], "SKU-MERGE")
        self.assertEqual(rows[0]["quantity"], 5)

    def test_same_order_different_sku_stays_as_two_normal_rows(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-3pcs.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-MULTI", "SKU-A", 1),
                ("ORDER-MULTI", "SKU-B", 2),
            ],
        )
        make_zip(self.input_dir / "different_sku.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 2)
        self.assertEqual(result["duplicate_count"], 0)
        rows = sorted(read_summary_rows(self.output_path), key=lambda item: str(item["sku"]))
        self.assertEqual([row["sku"] for row in rows], ["SKU-A", "SKU-B"])
        self.assertEqual([row["quantity"] for row in rows], [1, 2])

    def test_same_order_empty_sku_still_merges_when_order_id_exists(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-7pcs.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-EMPTY-SKU", "", 3),
                ("ORDER-EMPTY-SKU", "", 4),
            ],
        )
        make_zip(self.input_dir / "empty_sku.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 2)
        self.assertEqual(result["written_rows"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["order_id"], "ORDER-EMPTY-SKU")
        self.assertIn(rows[0]["sku"], ("", None))
        self.assertEqual(rows[0]["quantity"], 7)

    def test_exact_duplicate_rows_are_skipped_when_setting_is_enabled(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-2rows-2pcs.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("", "SKU-DUP", 1),
                ("", "SKU-DUP", 1),
            ],
        )
        make_zip(self.input_dir / "exact_duplicate_skip.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(skip_exact_duplicates=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 1)
        self.assertEqual(result["exact_duplicate_count"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 1)

    def test_exact_duplicate_rows_can_be_written_when_setting_is_disabled(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-2rows-2pcs.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("", "SKU-DUP", 1),
                ("", "SKU-DUP", 1),
            ],
        )
        make_zip(self.input_dir / "exact_duplicate_write.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(skip_exact_duplicates=False)

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 2)
        self.assertEqual(result["exact_duplicate_count"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 2)

    def test_archive_with_multiple_order_subfolders_processes_each_subfolder(self) -> None:
        first = self.work_dir / "0507-WZY-knife-1order-1pc.xlsx"
        second = self.work_dir / "0508-WZY-knife-1order-2pcs.xlsx"
        make_order_workbook(first, [("ORDER-SUB-1", "SKU-SUB-1", 1)])
        make_order_workbook(second, [("ORDER-SUB-2", "SKU-SUB-2", 2)])
        make_zip(
            self.input_dir / "multi_subfolders.zip",
            [
                (first, f"folder-a/{first.name}"),
                (second, f"folder-b/{second.name}"),
            ],
        )

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["processed_archives"], 1)
        self.assertEqual(result["written_rows"], 2)
        rows = sorted(read_summary_rows(self.output_path), key=lambda item: str(item["order_id"]))
        self.assertEqual([row["order_id"] for row in rows], ["ORDER-SUB-1", "ORDER-SUB-2"])
        self.assertTrue(Path(result["process_report_path"]).exists())

    def test_modify_archive_is_copied_to_skip_folder_and_excluded_from_output(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-1pc.xlsx"
        make_order_workbook(workbook_path, [("ORDER-SKIP", "SKU-SKIP", 1)])
        archive_name = f"{MODIFY_PREFIX}-order.zip"
        make_zip(self.input_dir / archive_name, [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertTrue((self.input_dir / SKIP_FOLDER_NAME / archive_name).exists())

    def test_modify_archive_dry_run_does_not_create_skip_folder(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-1pc.xlsx"
        make_order_workbook(workbook_path, [("ORDER-SKIP-DRY", "SKU-SKIP", 1)])
        archive_name = f"{MODIFY_PREFIX}-dry.zip"
        make_zip(self.input_dir / archive_name, [(workbook_path, workbook_path.name)])

        result = self.run_tool(dry_run=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse((self.input_dir / SKIP_FOLDER_NAME).exists())

    def test_reports_and_returned_paths_exist_after_run(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-knife-1order-1pc.xlsx"
        make_order_workbook(workbook_path, [("ORDER-REPORT", "SKU-REPORT", 1)])
        make_zip(self.input_dir / "reports.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertTrue(Path(result["process_report_path"]).exists())
        self.assertTrue(Path(result["debug_report_path"]).exists())
        self.assertTrue(Path(result["log_file_path"]).exists())


if __name__ == "__main__":
    unittest.main()
