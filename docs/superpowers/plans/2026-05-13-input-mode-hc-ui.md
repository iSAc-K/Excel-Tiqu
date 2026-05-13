# Input Mode HC UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add UI-selectable input source modes, optional remembered HC filtering, and a multi-file summary mode for both archives and normal folders.

**Architecture:** Keep `extract_orders.py` as the business-logic source of truth and keep `extract_orders_gui.py` as a thin wrapper that passes user choices into `run_extract(...)`. Add source discovery and folder-Excel processing in the core, then reuse the existing row extraction, merge, duplicate, Dry Run, report, and workbook-writing pipeline.

**Tech Stack:** Python 3.14, `unittest`, `openpyxl`, `customtkinter`, PowerShell on Windows.

---

## File Structure

- Modify `extract_orders.py`: add `input_mode`, `enable_hc_filter`, and `excel_group_mode`; add folder Excel discovery and processing; update archive processing to respect HC and multi-file rules; add CLI flags.
- Modify `extract_orders_gui.py`: add saved settings and UI controls for input source, HC filter, and Excel rule; pass values to `run_extract(...)`; allow selecting either a folder or a single archive when archive mode is active.
- Modify `test_hc_filter.py`: update HC tests so default behavior is off, and add enabled-HC assertions.
- Modify `test_core_regression.py`: add tests for folder mode, mixed mode, single-file vs multi-file Excel behavior.
- Modify `README.md`: document the new UI choices and CLI flags.

## Task 1: Core Mode Tests

**Files:**
- Modify: `test_core_regression.py`
- Modify: `test_hc_filter.py`

- [ ] **Step 1: Add input-mode helpers to core regression tests**

In `CoreRegressionTests.run_tool(...)`, add keyword arguments and pass them through to `run_extract(...)`:

```python
    def run_tool(
        self,
        *,
        dry_run: bool = False,
        skip_exact_duplicates: bool = True,
        clear: bool = True,
        input_mode: str = "archives",
        excel_group_mode: str = "single",
        enable_hc_filter: bool = False,
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
            input_mode=input_mode,
            excel_group_mode=excel_group_mode,
            enable_hc_filter=enable_hc_filter,
        )
```

- [ ] **Step 2: Add folder, mixed, and multi-file tests**

Append these tests to `CoreRegressionTests`:

```python
    def test_folder_mode_recursively_extracts_excel_files(self) -> None:
        nested = self.input_dir / "orders" / "day-1"
        nested.mkdir(parents=True)
        workbook_path = nested / "0507-WZY-knife-1order-4pcs.xlsx"
        make_order_workbook(workbook_path, [("ORDER-FOLDER", "SKU-FOLDER", 4)])

        result = self.run_tool(input_mode="folders")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 1)
        rows = read_summary_rows(self.output_path)
        self.assertEqual(rows[0]["order_id"], "ORDER-FOLDER")

    def test_mixed_mode_extracts_archive_and_folder_excel_files(self) -> None:
        folder_excel = self.input_dir / "folder-orders" / "0507-WZY-knife-1order-2pcs.xlsx"
        folder_excel.parent.mkdir(parents=True)
        make_order_workbook(folder_excel, [("ORDER-FOLDER-MIX", "SKU-FOLDER", 2)])
        archive_excel = self.work_dir / "0508-WZY-knife-1order-3pcs.xlsx"
        make_order_workbook(archive_excel, [("ORDER-ARCHIVE-MIX", "SKU-ARCHIVE", 3)])
        make_zip(self.input_dir / "archive_orders.zip", [(archive_excel, archive_excel.name)])

        result = self.run_tool(input_mode="mixed")

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 2)
        rows = sorted(read_summary_rows(self.output_path), key=lambda item: str(item["order_id"]))
        self.assertEqual([row["order_id"] for row in rows], ["ORDER-ARCHIVE-MIX", "ORDER-FOLDER-MIX"])

    def test_single_file_mode_skips_multiple_excel_files_in_one_folder_unit(self) -> None:
        folder = self.input_dir / "same-unit"
        folder.mkdir()
        make_order_workbook(folder / "0507-WZY-knife-1order-1pc.xlsx", [("ORDER-A", "SKU-A", 1)])
        make_order_workbook(folder / "0508-WZY-knife-1order-2pcs.xlsx", [("ORDER-B", "SKU-B", 2)])

        result = self.run_tool(input_mode="folders", excel_group_mode="single")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 0)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertGreaterEqual(result["skipped_archives"], 1)

    def test_multi_file_summary_mode_extracts_multiple_excel_files_in_one_folder_unit(self) -> None:
        folder = self.input_dir / "same-unit"
        folder.mkdir()
        make_order_workbook(folder / "0507-WZY-knife-1order-1pc.xlsx", [("ORDER-A", "SKU-A", 1)])
        make_order_workbook(folder / "0508-WZY-knife-1order-2pcs.xlsx", [("ORDER-B", "SKU-B", 2)])

        result = self.run_tool(input_mode="folders", excel_group_mode="multi")

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 2)
        rows = sorted(read_summary_rows(self.output_path), key=lambda item: str(item["order_id"]))
        self.assertEqual([row["order_id"] for row in rows], ["ORDER-A", "ORDER-B"])

    def test_multi_file_summary_mode_applies_inside_archives(self) -> None:
        first = self.work_dir / "0507-WZY-knife-1order-1pc.xlsx"
        second = self.work_dir / "0508-WZY-knife-1order-2pcs.xlsx"
        make_order_workbook(first, [("ORDER-ZIP-A", "SKU-A", 1)])
        make_order_workbook(second, [("ORDER-ZIP-B", "SKU-B", 2)])
        make_zip(self.input_dir / "multi_excel.zip", [(first, first.name), (second, second.name)])

        result = self.run_tool(excel_group_mode="multi")

        self.assertTrue(result["success"])
        self.assertEqual(result["written_rows"], 2)
        rows = sorted(read_summary_rows(self.output_path), key=lambda item: str(item["order_id"]))
        self.assertEqual([row["order_id"] for row in rows], ["ORDER-ZIP-A", "ORDER-ZIP-B"])
```

- [ ] **Step 3: Update HC tests to prove default off and optional on**

Change `HcFilterTests.run_tool(...)` to accept `enable_hc_filter`:

```python
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
```

Rename `test_hc_excel_is_copied_and_excluded_from_output` to `test_hc_filter_enabled_copies_and_excludes_from_output` and call:

```python
        result = self.run_tool(enable_hc_filter=True)
```

Add this new default-off test:

```python
    def test_hc_filter_default_off_processes_hc_excel_normally(self) -> None:
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
        self.assertGreaterEqual(result["skipped_archives"], 1)
```

Update the Dry Run and copy-failure tests to call `self.run_tool(dry_run=True, enable_hc_filter=True)` and `self.run_tool(enable_hc_filter=True)`.

- [ ] **Step 4: Run focused tests to verify expected failures**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_folder_mode_recursively_extracts_excel_files test_core_regression.CoreRegressionTests.test_mixed_mode_extracts_archive_and_folder_excel_files test_core_regression.CoreRegressionTests.test_multi_file_summary_mode_extracts_multiple_excel_files_in_one_folder_unit test_hc_filter.HcFilterTests.test_hc_filter_default_off_processes_hc_excel_normally -v
```

Expected: failures with `run_extract() got an unexpected keyword argument 'input_mode'` or equivalent missing behavior.

## Task 2: Core Input Discovery and Processing

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`
- Test: `test_hc_filter.py`

- [ ] **Step 1: Add constants and normalization helpers**

Near `ARCHIVE_EXTENSIONS` and `EXCEL_EXTENSIONS`, add:

```python
INPUT_MODE_ARCHIVES = "archives"
INPUT_MODE_FOLDERS = "folders"
INPUT_MODE_MIXED = "mixed"
INPUT_MODES = {INPUT_MODE_ARCHIVES, INPUT_MODE_FOLDERS, INPUT_MODE_MIXED}

EXCEL_GROUP_SINGLE = "single"
EXCEL_GROUP_MULTI = "multi"
EXCEL_GROUP_MODES = {EXCEL_GROUP_SINGLE, EXCEL_GROUP_MULTI}
```

Near `clamp_workers(...)`, add:

```python
def normalize_input_mode(value: str | None) -> str:
    mode = (value or INPUT_MODE_ARCHIVES).strip().lower()
    if mode not in INPUT_MODES:
        raise ValueError(f"input_mode 必须是 archives、folders 或 mixed，当前值：{value}")
    return mode


def normalize_excel_group_mode(value: str | None) -> str:
    mode = (value or EXCEL_GROUP_SINGLE).strip().lower()
    if mode not in EXCEL_GROUP_MODES:
        raise ValueError(f"excel_group_mode 必须是 single 或 multi，当前值：{value}")
    return mode
```

- [ ] **Step 2: Add folder Excel discovery**

After `find_excel_files_in_extracted_dir(...)`, add:

```python
def find_folder_excel_files(input_path: str | Path) -> list[Path]:
    path = Path(input_path).expanduser()
    if path.is_file():
        suffix = path.suffix.lower()
        if suffix in EXCEL_EXTENSIONS and not path.name.startswith("~$"):
            add_log(f"输入为单个 Excel 文件：{path}")
            return [path]
        if suffix == ".xls":
            add_log(f"跳过文件：{path.name}，暂不支持 .xls")
            return []
        return []

    if not path.is_dir():
        return []

    add_log(f"开始扫描文件夹里的 Excel：{path}")
    excel_files: list[Path] = []
    for file_path in sorted(path.rglob("*")):
        if not file_path.is_file():
            continue
        if "未处理压缩包" in file_path.relative_to(path).parts or HC_FOLDER_NAME in file_path.relative_to(path).parts:
            continue
        suffix = file_path.suffix.lower()
        if file_path.name.startswith("~$"):
            continue
        if suffix == ".xls":
            add_log(f"跳过文件：{file_path.name}，暂不支持 .xls")
            continue
        if suffix in EXCEL_EXTENSIONS:
            excel_files.append(file_path)
    add_log(f"找到文件夹 Excel：{len(excel_files)} 个")
    return excel_files
```

- [ ] **Step 3: Let `process_excel_unit(...)` support multi-file mode**

Change the function signature:

```python
def process_excel_unit(
    archive_path: Path,
    extracted_root: Path,
    excel_files: list[Path],
    unit_folder: Path,
    category_keywords: dict[str, list[str]] | None = None,
    skip_dir: Path | None = None,
    dry_run: bool = False,
    excel_group_mode: str = EXCEL_GROUP_SINGLE,
) -> dict[str, Any]:
```

Replace the `if excel_count > 1:` block with:

```python
    if excel_count > 1 and excel_group_mode == EXCEL_GROUP_SINGLE:
        reason = "子文件夹中识别到多个正式 Excel"
        related = "；".join(str(file.relative_to(extracted_root)) for file in excel_files)
        add_exception(f"异常：{archive_path.name} / {subfolder_name or '.'} 中识别到多个正式 Excel，已跳过该子文件夹\n{related}")
        add_structured_error(archive_path, "子文件夹多个正式Excel", reason, related)
        return build_result(False, "跳过", skip=True, reason=reason)
```

Replace the single-file extraction block from `excel_file = excel_files[0]` through `return build_result(True, "成功")` with a loop that accumulates rows:

```python
    selected_excel_name = "；".join(str(file.relative_to(extracted_root)) for file in excel_files)
    for excel_file in excel_files:
        if subfolder_name:
            add_log(f"子文件夹：{subfolder_name}")
        add_log(f"[{archive_path.name}] 找到正式 Excel：{excel_file.name}")
        if excel_file.stem.startswith("修改"):
            reason = "文件名前两个字为“修改”"
            copied_to = copy_skipped_source(archive_path, skip_dir, dry_run=dry_run)
            add_log(f"[{archive_path.name}] 已跳过：{excel_file.name}，原因：{reason}")
            add_structured_error(
                archive_path,
                "修改文件跳过",
                f"{reason}；原文件名：{excel_file.name}；已复制到：{copied_to}",
                excel_file.name,
                status="已跳过",
            )
            if excel_count == 1:
                return build_result(False, "跳过", skip=True, reason=reason, ignored=True, copied_to=copied_to)
            continue

        extracted_rows, workbook_success, current_category, current_date_text, current_header_rows = extract_rows_from_workbook(
            excel_file,
            category_keywords,
        )
        for row in current_header_rows:
            row["压缩包名"] = archive_path.name
            row["子文件夹名"] = subfolder_name
        header_report_rows.extend(current_header_rows)
        for row in extracted_rows:
            row["外层压缩包名"] = archive_path.name
            row["子文件夹名"] = subfolder_name
        if not workbook_success:
            if excel_count == 1:
                return build_result(False, "异常", skip=True, reason="Excel 读取失败或未识别到任何核心表头")
            continue

        current_validation, current_quantity_errors = build_filename_validation(
            archive_path.name,
            excel_file.name,
            extracted_rows,
            current_category,
            current_date_text,
            subfolder_name,
        )
        filename_validation = current_validation if filename_validation is None else filename_validation
        quantity_error_rows.extend(current_quantity_errors)
        if not category:
            category = current_category
        if not date_text:
            date_text = current_date_text
        rows.extend(extracted_rows)

    if rows:
        add_log(f"[{archive_path.name}] 目标工作表：{rows[0]['category']}")
        return build_result(True, "成功")
    return build_result(False, "异常", skip=True, reason="Excel 读取失败或未识别到任何核心表头")
```

- [ ] **Step 4: Pass modes through archive processing**

Change `process_archive(...)` signature:

```python
def process_archive(
    archive_path: Path,
    category_keywords: dict[str, list[str]] | None = None,
    skip_dir: Path | None = None,
    hc_dir: Path | None = None,
    dry_run: bool = False,
    enable_hc_filter: bool = False,
    excel_group_mode: str = EXCEL_GROUP_SINGLE,
) -> dict[str, Any]:
```

Replace HC splitting inside `process_archive(...)`:

```python
            if enable_hc_filter:
                hc_files, normal_excel_files = split_hc_excel_files(excel_files)
            else:
                hc_files, normal_excel_files = [], excel_files
```

Update the `process_excel_unit(...)` call to pass:

```python
                    excel_group_mode=excel_group_mode,
```

- [ ] **Step 5: Add folder Excel task processing**

Add after `process_archive(...)`:

```python
def build_folder_processing_groups(input_root: Path, excel_files: list[Path]) -> list[tuple[Path, list[Path]]]:
    grouped: dict[Path, list[Path]] = {}
    for excel_file in excel_files:
        grouped.setdefault(excel_file.parent, []).append(excel_file)
    return [(folder, sorted(files)) for folder, files in sorted(grouped.items(), key=lambda item: str(item[0]))]


def process_folder_excel_group(
    input_root: Path,
    unit_folder: Path,
    excel_files: list[Path],
    category_keywords: dict[str, list[str]] | None = None,
    skip_dir: Path | None = None,
    hc_dir: Path | None = None,
    dry_run: bool = False,
    enable_hc_filter: bool = False,
    excel_group_mode: str = EXCEL_GROUP_SINGLE,
) -> dict[str, Any]:
    virtual_archive = unit_folder
    folder_process_logs: list[str] = []
    folder_exception_logs: list[str] = []
    folder_error_report_rows: list[dict[str, Any]] = []
    folder_debug_logs: list[str] = []
    set_thread_log_context(virtual_archive, folder_process_logs, folder_exception_logs, folder_error_report_rows, folder_debug_logs)
    hc_report_rows: list[dict[str, Any]] = []
    try:
        if enable_hc_filter:
            hc_files, normal_excel_files = split_hc_excel_files(excel_files)
        else:
            hc_files, normal_excel_files = [], excel_files
        if hc_files and hc_dir is not None:
            for hc_file in hc_files:
                hc_report_rows.append(copy_or_preview_hc_excel(virtual_archive, input_root, hc_file, hc_dir, dry_run))
        result = process_excel_unit(
            virtual_archive,
            input_root,
            normal_excel_files,
            unit_folder,
            category_keywords,
            skip_dir,
            dry_run,
            excel_group_mode=excel_group_mode,
        )
        result["process_logs"] = list(folder_process_logs)
        result["exception_logs"] = list(folder_exception_logs)
        result["debug_logs"] = list(folder_debug_logs)
        result["error_report_rows"] = list(folder_error_report_rows)
        result["hc_report_rows"] = list(hc_report_rows)
        return result
    finally:
        clear_thread_log_context()
```

- [ ] **Step 6: Update `run_extract(...)` signature and task loop**

Change the signature:

```python
def run_extract(
    input_path: str,
    output_path: str,
    *,
    clear: bool = False,
    dry_run: bool = False,
    workers: int = 4,
    detect_duplicate_orders: bool = True,
    skip_exact_duplicates: bool = True,
    category_config_path: str | Path | None = None,
    settings_path: str | Path | None = None,
    report_dir: str | Path | None = None,
    backup_dir: str | Path | None = None,
    log_callback: Callable[[str], None] | None = None,
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    input_mode: str = INPUT_MODE_ARCHIVES,
    enable_hc_filter: bool = False,
    excel_group_mode: str = EXCEL_GROUP_SINGLE,
) -> dict[str, Any]:
```

Near the start of the function, after workers are normalized:

```python
    normalized_input_mode = normalize_input_mode(input_mode)
    normalized_excel_group_mode = normalize_excel_group_mode(excel_group_mode)
```

Replace the current archive-only block starting at `archives = find_archive_files(input_path)` with task discovery:

```python
        archives: list[Path] = []
        folder_groups: list[tuple[Path, list[Path]]] = []
        input_root = Path(input_path).expanduser()
        if normalized_input_mode in {INPUT_MODE_ARCHIVES, INPUT_MODE_MIXED}:
            archives = find_archive_files(input_path)
        if normalized_input_mode in {INPUT_MODE_FOLDERS, INPUT_MODE_MIXED}:
            folder_excel_files = find_folder_excel_files(input_path)
            folder_groups = build_folder_processing_groups(input_root if input_root.is_dir() else input_root.parent, folder_excel_files)

        total_tasks = len(archives) + len(folder_groups)
        stats["total_archives"] = total_tasks
```

Then process archives and folder groups. The implementation can keep archives threaded first and process folder groups in the same executor with task dictionaries:

```python
            with ThreadPoolExecutor(max_workers=normalized_workers) as executor:
                future_map = {}
                task_index = 0
                for archive_path in archives:
                    task_index += 1
                    emit_progress({"current": 0, "total": total_tasks, "archive_name": archive_path.name, "status": "queued", "active_workers": min(normalized_workers, total_tasks), "completed_archives": 0, "failed_archives": 0})
                    future = executor.submit(process_archive, archive_path, category_keywords, skip_dir, hc_dir, dry_run, enable_hc_filter, normalized_excel_group_mode)
                    future_map[future] = (task_index, archive_path.name)
                for unit_folder, unit_files in folder_groups:
                    task_index += 1
                    emit_progress({"current": 0, "total": total_tasks, "archive_name": unit_folder.name, "status": "queued", "active_workers": min(normalized_workers, total_tasks), "completed_archives": 0, "failed_archives": 0})
                    future = executor.submit(process_folder_excel_group, input_root if input_root.is_dir() else input_root.parent, unit_folder, unit_files, category_keywords, skip_dir, hc_dir, dry_run, enable_hc_filter, normalized_excel_group_mode)
                    future_map[future] = (task_index, unit_folder.name)
```

Keep the existing result aggregation logic, changing loops from `range(1, len(archives) + 1)` to `range(1, total_tasks + 1)` and progress totals from `len(archives)` to `total_tasks`.

- [ ] **Step 7: Add CLI flags**

In `main()`, add:

```python
    parser.add_argument("--input-mode", choices=sorted(INPUT_MODES), default=INPUT_MODE_ARCHIVES, help="输入来源：archives 只处理压缩包，folders 只处理文件夹 Excel，mixed 混合模式")
    parser.add_argument("--enable-hc-filter", action="store_true", help="启用 HC 文件过滤；默认不启用")
    parser.add_argument("--excel-group-mode", choices=sorted(EXCEL_GROUP_MODES), default=EXCEL_GROUP_SINGLE, help="Excel 处理规则：single 单文件订单模式，multi 多文件汇总模式")
```

Pass them into `run_extract(...)`:

```python
        input_mode=args.input_mode,
        enable_hc_filter=args.enable_hc_filter,
        excel_group_mode=args.excel_group_mode,
```

- [ ] **Step 8: Run focused core tests**

Run:

```powershell
python -m unittest test_core_regression.py test_hc_filter.py -v
```

Expected: all tests pass.

- [ ] **Step 9: Commit core changes**

Run:

```powershell
git add extract_orders.py test_core_regression.py test_hc_filter.py
git commit -m "feat: add input modes and optional HC filtering"
```

## Task 3: GUI Controls and Settings

**Files:**
- Modify: `extract_orders_gui.py`
- Test: manual GUI smoke check

- [ ] **Step 1: Add settings defaults and UI variables**

Add to `DEFAULT_SETTINGS`:

```python
    "input_mode": "archives",
    "enable_hc_filter": False,
    "excel_group_mode": "single",
```

Add display maps near `PAGE_TITLES`:

```python
INPUT_MODE_LABELS = {
    "只处理压缩包": "archives",
    "只处理文件夹Excel": "folders",
    "混合模式": "mixed",
}
INPUT_MODE_VALUES = {value: label for label, value in INPUT_MODE_LABELS.items()}

EXCEL_GROUP_LABELS = {
    "单文件订单模式": "single",
    "多文件汇总模式": "multi",
}
EXCEL_GROUP_VALUES = {value: label for label, value in EXCEL_GROUP_LABELS.items()}
```

In `ExtractOrdersApp.__init__`, add:

```python
        self.input_mode_var = tk.StringVar(value=INPUT_MODE_VALUES.get(str(self.settings.get("input_mode") or "archives"), "只处理压缩包"))
        self.enable_hc_filter_var = tk.BooleanVar(value=bool(self.settings.get("enable_hc_filter", False)))
        self.excel_group_mode_var = tk.StringVar(value=EXCEL_GROUP_VALUES.get(str(self.settings.get("excel_group_mode") or "single"), "单文件订单模式"))
```

- [ ] **Step 2: Add controls to the start/options page**

In `_build_start_page(...)`, change the first path label from “压缩包所在文件夹” to “输入路径”, and the button text to “选择文件夹”.

In `_build_options(...)`, add these controls after the worker slider row:

```python
        ctk.CTkLabel(frame, text="数据来源", text_color="#172033").grid(row=2, column=0, sticky="w", padx=18, pady=(4, 8))
        self.input_mode_menu = ctk.CTkOptionMenu(frame, values=list(INPUT_MODE_LABELS.keys()), variable=self.input_mode_var, command=lambda _value: self.update_top_status())
        self.input_mode_menu.grid(row=2, column=1, sticky="w", padx=14, pady=(4, 8))

        ctk.CTkLabel(frame, text="Excel 规则", text_color="#172033").grid(row=3, column=0, sticky="w", padx=18, pady=(4, 8))
        self.excel_group_mode_menu = ctk.CTkOptionMenu(frame, values=list(EXCEL_GROUP_LABELS.keys()), variable=self.excel_group_mode_var, command=lambda _value: self.update_top_status())
        self.excel_group_mode_menu.grid(row=3, column=1, sticky="w", padx=14, pady=(4, 8))

        self.enable_hc_filter_check = ctk.CTkCheckBox(frame, text="过滤 HC 文件", variable=self.enable_hc_filter_var, command=self.update_top_status)
        self.enable_hc_filter_check.grid(row=3, column=2, sticky="w", padx=(0, 18), pady=(4, 8))
```

Move the existing `clear_check`, `detect_duplicates_check`, and `skip_exact_duplicates_check` down to row 4.

Extend `self.start_controls`:

```python
        self.start_controls.extend([self.input_mode_menu, self.excel_group_mode_menu, self.enable_hc_filter_check])
```

- [ ] **Step 3: Save and validate settings**

In `save_current_settings(...)`, add:

```python
            "input_mode": INPUT_MODE_LABELS.get(self.input_mode_var.get(), "archives"),
            "enable_hc_filter": self.enable_hc_filter_var.get(),
            "excel_group_mode": EXCEL_GROUP_LABELS.get(self.excel_group_mode_var.get(), "single"),
```

Change `validate_inputs(...)` so the input path can be a folder or a supported archive/Excel file:

```python
        input_item = Path(input_path).expanduser()
        if not input_item.exists():
            messagebox.showwarning("提示", "输入路径不存在")
            return False, "", ""
        if not input_item.is_dir() and input_item.suffix.lower() not in {".zip", ".rar", ".7z", ".xlsx", ".xlsm"}:
            messagebox.showwarning("提示", "输入路径必须是文件夹、压缩包或 Excel 文件")
            return False, "", ""
```

Return `str(input_item)` instead of `str(input_folder)`.

- [ ] **Step 4: Pass choices to `run_extract(...)`**

Add helper methods:

```python
    def selected_input_mode(self) -> str:
        return INPUT_MODE_LABELS.get(self.input_mode_var.get(), "archives")

    def selected_excel_group_mode(self) -> str:
        return EXCEL_GROUP_LABELS.get(self.excel_group_mode_var.get(), "single")
```

In `start_extract(...)`, pass `self.selected_input_mode()`, `self.enable_hc_filter_var.get()`, and `self.selected_excel_group_mode()` into the worker thread args.

Change `run_extract_worker(...)` signature and `run_extract(...)` call:

```python
                input_mode=input_mode,
                enable_hc_filter=enable_hc_filter,
                excel_group_mode=excel_group_mode,
```

- [ ] **Step 5: Update status text**

In `update_top_status(...)`, append source and HC state:

```python
        self.worker_status_var.set(f"并发：{self.workers_var.get()} | 来源：{self.input_mode_var.get()} | {self.excel_group_mode_var.get()}")
```

- [ ] **Step 6: GUI smoke check**

Run:

```powershell
python -m py_compile extract_orders_gui.py
python -c "import extract_orders_gui; app = extract_orders_gui.ExtractOrdersApp(); app.after(800, app.destroy); app.mainloop(); print('gui startup ok')"
```

Expected: `gui startup ok` and no traceback.

- [ ] **Step 7: Commit GUI changes**

Run:

```powershell
git add extract_orders_gui.py app_settings.json
git commit -m "feat: expose input and HC options in GUI"
```

## Task 4: Docs and Final Verification

**Files:**
- Modify: `README.md`
- Verify: `extract_orders.py`, `extract_orders_gui.py`, tests

- [ ] **Step 1: Update README usage text**

Add a short section explaining:

```markdown
### 输入来源和 Excel 规则

界面里可以选择三种数据来源：

- 只处理压缩包：递归处理 `.zip`、`.rar`、`.7z`
- 只处理文件夹里的 Excel：递归处理 `.xlsx`、`.xlsm`
- 混合模式：压缩包和普通 Excel 一起处理

HC 过滤默认关闭。勾选“过滤 HC 文件”后，文件名包含 HC 的 Excel 会从正常提取中排除；正式运行会复制到输入目录下的 `HC` 文件夹，Dry Run 只记录计划，不创建文件夹。

Excel 规则有两种：

- 单文件订单模式：同一处理单元里多个正式 Excel 会跳过并记录异常
- 多文件汇总模式：同一处理单元里多个正式 Excel 会逐个提取并汇总
```

Add CLI examples:

```powershell
python extract_orders.py --input "D:\orders" --output "D:\summary.xlsx" --input-mode folders --excel-group-mode multi
python extract_orders.py --input "D:\orders" --output "D:\summary.xlsx" --input-mode mixed --enable-hc-filter
```

- [ ] **Step 2: Run full source verification**

Run:

```powershell
python -m py_compile extract_orders.py extract_orders_gui.py test_core_regression.py test_hc_filter.py
python -m unittest test_core_regression.py test_hc_filter.py -v
```

Expected: compile succeeds and all tests pass.

- [ ] **Step 3: Run sample CLI checks**

Use existing tests as the authoritative sample check by running:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_folder_mode_recursively_extracts_excel_files test_core_regression.CoreRegressionTests.test_mixed_mode_extracts_archive_and_folder_excel_files test_core_regression.CoreRegressionTests.test_multi_file_summary_mode_applies_inside_archives test_hc_filter.HcFilterTests.test_hc_filter_enabled_copies_and_excludes_from_output -v
```

Expected: all selected tests pass.

- [ ] **Step 4: Commit docs and verification-ready state**

Run:

```powershell
git add README.md docs/superpowers/plans/2026-05-13-input-mode-hc-ui.md
git commit -m "docs: document input mode implementation plan"
```

- [ ] **Step 5: Report final status**

Include:

- Modified files.
- Commands run and pass/fail result.
- Whether GUI startup smoke check passed.
- Current Git commit IDs.
- Any packaging not yet rebuilt, if no EXE rebuild was requested in this phase.
