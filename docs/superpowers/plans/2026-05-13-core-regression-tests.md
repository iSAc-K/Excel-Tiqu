# Core Regression Tests Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a focused regression test suite for the Excel order extractor's core business rules so future feature changes can be verified with one command.

**Architecture:** Keep production code unchanged unless a test exposes a real bug. Add one new `unittest` file beside the existing `test_hc_filter.py`, using temporary folders and generated `.xlsx` / `.zip` fixtures. Tests call the public `run_extract(...)` API and inspect returned stats, generated workbooks, copied files, and report paths.

**Tech Stack:** Python `unittest`, `tempfile`, `zipfile`, `openpyxl`, existing `extract_orders.run_extract`.

---

## File Structure

- Create: `test_core_regression.py`
  - Owns core regression fixtures and tests for Dry Run, merge, duplicate handling, skipped `修改` archives, multi-subfolder archives, and report artifact creation.
- Modify: none expected.
  - Only touch `extract_orders.py` if a test reveals an actual existing defect.
- Test: `test_core_regression.py`, plus keep `test_hc_filter.py` green.

## Task 1: Add Shared Regression Test Fixtures

**Files:**
- Create: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Create the test file with reusable helpers**

Create `test_core_regression.py` with this content:

```python
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
            if headers != [ORDER_HEADER, "SKU", QUANTITY_HEADER, "\u65e5\u671f"]:
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

    def run_tool(self, *, dry_run: bool = False, skip_exact_duplicates: bool = True) -> dict:
        return run_extract(
            str(self.input_dir),
            str(self.output_path),
            clear=True,
            dry_run=dry_run,
            workers=1,
            report_dir=str(self.logs_dir),
            backup_dir=str(self.backups_dir),
            skip_exact_duplicates=skip_exact_duplicates,
        )
```

- [ ] **Step 2: Run the empty fixture file**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.py -v
```

Expected: `Ran 0 tests` or an equivalent successful unittest discovery result.

## Task 2: Test Dry Run Side-Effect Boundaries

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add Dry Run test**

Add this method inside `CoreRegressionTests`:

```python
    def test_dry_run_scans_and_reports_without_writing_output_or_backup(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-2个.xlsx"
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
```

- [ ] **Step 2: Run the new failing-or-passing test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_dry_run_scans_and_reports_without_writing_output_or_backup -v
```

Expected: PASS. If it fails because Dry Run creates output or backup artifacts, stop and inspect before changing production code.

## Task 3: Test Same Order + Same SKU Quantity Merge

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add merge test**

Add this method inside `CoreRegressionTests`:

```python
    def test_same_order_same_sku_merges_quantities_before_write(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-5个.xlsx"
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
```

- [ ] **Step 2: Run the merge test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_same_order_same_sku_merges_quantities_before_write -v
```

Expected: PASS.

## Task 4: Test Same Order + Different SKU Stays Separate

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add different-SKU test**

Add this method inside `CoreRegressionTests`:

```python
    def test_same_order_different_sku_stays_as_two_normal_rows(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-3个.xlsx"
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
```

- [ ] **Step 2: Run the different-SKU test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_same_order_different_sku_stays_as_two_normal_rows -v
```

Expected: PASS.

## Task 5: Test Empty SKU Merge Behavior

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add empty-SKU merge test**

Add this method inside `CoreRegressionTests`:

```python
    def test_same_order_empty_sku_still_merges_when_order_id_exists(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-7个.xlsx"
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
```

- [ ] **Step 2: Run the empty-SKU test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_same_order_empty_sku_still_merges_when_order_id_exists -v
```

Expected: PASS.

## Task 6: Test Exact Duplicate Skip Toggle

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add exact duplicate tests**

Add these methods inside `CoreRegressionTests`:

```python
    def test_exact_duplicate_rows_are_skipped_when_setting_is_enabled(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-2单-2个.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-DUP", "SKU-DUP", 1),
                ("ORDER-DUP", "SKU-DUP", 1),
            ],
        )
        make_zip(self.input_dir / "exact_duplicate_skip.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(skip_exact_duplicates=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 1)
        self.assertGreaterEqual(result["exact_duplicate_count"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 1)

    def test_exact_duplicate_rows_can_be_written_when_setting_is_disabled(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-2单-2个.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-DUP-WRITE", "SKU-DUP", 1),
                ("ORDER-DUP-WRITE", "SKU-DUP", 1),
            ],
        )
        make_zip(self.input_dir / "exact_duplicate_write.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(skip_exact_duplicates=False)

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 1)
        self.assertGreaterEqual(result["exact_duplicate_count"], 0)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["quantity"], 2)
```

Note: this project merges same order + same SKU before exact duplicate detection. The disabled case should still write one merged row with quantity `2`; this protects the intended merge-before-dedupe ordering.

- [ ] **Step 2: Run exact duplicate tests**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest `
  test_core_regression.CoreRegressionTests.test_exact_duplicate_rows_are_skipped_when_setting_is_enabled `
  test_core_regression.CoreRegressionTests.test_exact_duplicate_rows_can_be_written_when_setting_is_disabled -v
```

Expected: PASS.

## Task 7: Test Multi-Subfolder Archive Splitting

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add multi-subfolder test**

Add this method inside `CoreRegressionTests`:

```python
    def test_archive_with_multiple_order_subfolders_processes_each_subfolder(self) -> None:
        first = self.work_dir / "0507-WZY-刀叉-1单-1个.xlsx"
        second = self.work_dir / "0508-WZY-刀叉-1单-2个.xlsx"
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
```

- [ ] **Step 2: Run multi-subfolder test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_archive_with_multiple_order_subfolders_processes_each_subfolder -v
```

Expected: PASS.

## Task 8: Test `修改` Archive Skip Behavior

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add skipped archive tests**

Add these methods inside `CoreRegressionTests`:

```python
    def test_modify_archive_is_copied_to_skip_folder_and_excluded_from_output(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-1个.xlsx"
        make_order_workbook(workbook_path, [("ORDER-SKIP", "SKU-SKIP", 1)])
        make_zip(self.input_dir / "修改-order.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertTrue((self.input_dir / "未处理压缩包" / "修改-order.zip").exists())

    def test_modify_archive_dry_run_does_not_create_skip_folder(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-1单-1个.xlsx"
        make_order_workbook(workbook_path, [("ORDER-SKIP-DRY", "SKU-SKIP", 1)])
        make_zip(self.input_dir / "修改-dry.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool(dry_run=True)

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse((self.input_dir / "未处理压缩包").exists())
```

- [ ] **Step 2: Run skipped archive tests**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest `
  test_core_regression.CoreRegressionTests.test_modify_archive_is_copied_to_skip_folder_and_excluded_from_output `
  test_core_regression.CoreRegressionTests.test_modify_archive_dry_run_does_not_create_skip_folder -v
```

Expected: PASS.

## Task 9: Test Report Artifacts Are Created and User-Visible

**Files:**
- Modify: `test_core_regression.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add report artifact test**

Add this method inside `CoreRegressionTests`:

```python
    def test_reports_and_returned_paths_exist_after_run_with_duplicates(self) -> None:
        workbook_path = self.work_dir / "0507-WZY-刀叉-2单-2个.xlsx"
        make_order_workbook(
            workbook_path,
            [
                ("ORDER-REPORT", "SKU-REPORT", 1),
                ("ORDER-REPORT", "SKU-REPORT", 1),
            ],
        )
        make_zip(self.input_dir / "reports.zip", [(workbook_path, workbook_path.name)])

        result = self.run_tool()

        self.assertTrue(result["success"])
        self.assertTrue(Path(result["process_report_path"]).exists())
        self.assertTrue(Path(result["debug_report_path"]).exists())
        self.assertTrue(Path(result["log_file_path"]).exists())
        duplicate_report = result.get("duplicate_report_path") or result.get("duplicate_report_file_path")
        if duplicate_report:
            self.assertTrue(Path(duplicate_report).exists())
```

- [ ] **Step 2: Run report artifact test**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.CoreRegressionTests.test_reports_and_returned_paths_exist_after_run_with_duplicates -v
```

Expected: PASS.

## Task 10: Run Full Regression Verification

**Files:**
- Test: `test_core_regression.py`, `test_hc_filter.py`

- [ ] **Step 1: Compile source and tests**

Run:

```powershell
python -m py_compile extract_orders.py extract_orders_gui.py test_hc_filter.py test_core_regression.py
```

Expected: no output and exit code `0`.

- [ ] **Step 2: Run all regression tests with project-local dependencies**

Run:

```powershell
$env:PYTHONPATH = (Join-Path (Get-Location) '.codex_test_deps')
python -m unittest test_core_regression.py test_hc_filter.py -v
```

Expected: all tests PASS. Existing `test_hc_filter.py` should still report 4 passing tests.

- [ ] **Step 3: Record verification notes**

If this directory is not a git repository, do not create commits. Report:

```text
Created: test_core_regression.py
Verification:
- py_compile passed
- unittest test_core_regression.py test_hc_filter.py passed with PYTHONPATH=.codex_test_deps
Notes:
- Direct python may not see openpyxl unless dependencies are installed globally or PYTHONPATH points at .codex_test_deps.
```

## Self-Review

Spec coverage:
- Dry Run side-effect boundary: Task 2.
- Same order + same SKU merge: Task 3.
- Same order + different SKU normal rows: Task 4.
- Empty SKU merge: Task 5.
- Exact duplicate toggle and merge-before-dedupe ordering: Task 6.
- Multi-subfolder archive processing: Task 7.
- `修改` skip behavior and Dry Run no-copy behavior: Task 8.
- Report artifact paths and existence: Task 9.
- HC behavior remains covered by existing `test_hc_filter.py` and included in full verification: Task 10.

Placeholder scan:
- No `TBD`, `TODO`, or undefined implementation placeholders remain.

Type consistency:
- Tests use existing `run_extract(...)` return keys already used by `test_hc_filter.py` and `extract_orders_gui.py`.
- Test helper workbook headers use Unicode escapes for Chinese labels to avoid source encoding ambiguity.
