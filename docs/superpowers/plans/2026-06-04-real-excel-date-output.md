# Real Excel Date Output Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Write single-day filename dates to the output workbook as real Excel date values while preserving the `4月17日` cell display.

**Architecture:** Keep the existing row dictionary flow intact. Change filename date parsing so single dates produce `datetime.date`, keep ranges as text, and centralize display formatting for logs and reports. Apply the Excel date number format at the output worksheet date column after appending rows.

**Tech Stack:** Python `datetime.date`, `openpyxl`, existing `unittest` regression tests.

---

### Task 1: Add Failing Regression Tests

**Files:**
- Modify: `test_core_regression.py`

- [x] **Step 1: Add tests for real date cells and range text**

Add tests that run the public `run_extract(...)` API, inspect the generated workbook with `openpyxl`, and assert:

```python
from datetime import date

from extract_orders import OUTPUT_DATE_YEAR, run_extract

...

def test_single_filename_date_is_written_as_real_excel_date(self) -> None:
    workbook_path = self.work_dir / "0417-WZY-knife-1order-2pcs.xlsx"
    make_order_workbook(workbook_path, [("ORDER-DATE", "SKU-DATE", 2)])
    make_zip(self.input_dir / "date.zip", [(workbook_path, workbook_path.name)])

    result = self.run_tool()

    self.assertTrue(result["success"])
    workbook = load_workbook(self.output_path, data_only=True)
    try:
        worksheet = workbook.active
        cell = worksheet.cell(row=2, column=4)
        self.assertEqual(cell.value.date() if hasattr(cell.value, "date") else cell.value, date(OUTPUT_DATE_YEAR, 4, 17))
        self.assertEqual(cell.number_format, 'm"月"d"日"')
    finally:
        workbook.close()
```

- [x] **Step 2: Run the new tests and verify they fail**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_single_filename_date_is_written_as_real_excel_date
```

Expected: FAIL because the output cell currently contains text like `4月17日`, not a date.

### Task 2: Implement Real Date Parsing and Output Formatting

**Files:**
- Modify: `extract_orders.py`

- [x] **Step 1: Add `OUTPUT_DATE_YEAR` and date helpers**

Import `date`, add `OUTPUT_DATE_YEAR = datetime.now().year`, and make single-date helpers return `date(OUTPUT_DATE_YEAR, month, day)`.

- [x] **Step 2: Keep date ranges as text**

Leave the existing range branch returning `5月1日-5月3日` through a text-format helper.

- [x] **Step 3: Apply Excel date format to the output date column**

After headers are ensured and rows appended, set date column cells with `datetime.date` values to:

```python
cell.number_format = 'm"月"d"日"'
```

- [x] **Step 4: Use display formatting for report grouping**

Make daily report grouping and human-readable duplicate/log output use a display formatter so date objects still show as `5月7日` in reports.

### Task 3: Verify Existing Behavior

**Files:**
- Test: `test_core_regression.py`
- Test: `test_hc_filter.py`

- [x] **Step 1: Run targeted date tests**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_single_filename_date_is_written_as_real_excel_date
```

Expected: PASS.

- [x] **Step 2: Run the full test suite**

Run:

```powershell
python -m unittest
```

Expected: all tests pass.
