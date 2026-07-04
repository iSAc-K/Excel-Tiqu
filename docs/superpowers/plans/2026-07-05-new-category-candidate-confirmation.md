# New Category Candidate Confirmation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a safe GUI-confirmed workflow that discovers possible new product categories from unclassified folder names and saves user-approved prefix and category rules.

**Architecture:** Keep category recognition, candidate parsing, config migration, and persistence in `extract_orders.py` so CLI, GUI, and tests share one implementation. Add a small GUI review window in `extract_orders_gui.py` that reads candidates from the last run result, lets the user edit prefix/category values, and calls core save helpers. Keep unconfirmed candidates excluded from formal output.

**Tech Stack:** Python 3, openpyxl, customtkinter/tkinter, unittest.

---

## File Structure

- Modify `extract_orders.py`
  - Add a `CategoryConfigData` dataclass for categories plus supplier/channel prefixes.
  - Preserve old `category_config.json` compatibility while allowing the new structured format.
  - Add candidate parsing helpers and candidate merge/save helpers.
  - Thread candidate metadata through workbook/archive/folder processing results.
  - Add a `新品类候选` sheet to the process report.
- Modify `extract_orders_gui.py`
  - Add an entry point on the reports/config page for reviewing new category candidates.
  - Add `NewCategoryCandidatesWindow` for editing prefix/category values and saving confirmations.
  - Keep the existing category keyword editor working with the category map.
- Modify `test_core_regression.py`
  - Add unit tests for config compatibility, candidate parsing, unclassified behavior, report output, and save helpers.
- Modify `README.md`
  - Document the candidate review workflow and the new structured `category_config.json` format.

## Task 1: Category Config Compatibility

**Files:**
- Modify: `extract_orders.py:80-215`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add failing tests for old and new config formats**

Add these tests near the beginning of `CoreRegressionTests`, after `test_fixture_helpers_create_readable_order_zip`:

```python
    def test_load_old_category_config_returns_empty_prefixes(self) -> None:
        self.category_config_path.write_text(
            json.dumps({"方黑名片架": ["方黑名片架"]}, ensure_ascii=False),
            encoding="utf-8",
        )

        from extract_orders import load_category_config_data

        config_data, path, error = load_category_config_data(self.category_config_path)

        self.assertEqual(error, "")
        self.assertEqual(Path(path), self.category_config_path)
        self.assertEqual(config_data.categories, {"方黑名片架": ["方黑名片架"]})
        self.assertEqual(config_data.prefixes, [])

    def test_load_structured_category_config_reads_prefixes(self) -> None:
        self.category_config_path.write_text(
            json.dumps(
                {
                    "prefixes": ["WZY", "HAL"],
                    "categories": {"纯木名片架": ["纯木名片架"]},
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        from extract_orders import load_category_config_data

        config_data, _, error = load_category_config_data(self.category_config_path)

        self.assertEqual(error, "")
        self.assertEqual(config_data.prefixes, ["WZY", "HAL"])
        self.assertEqual(config_data.categories, {"纯木名片架": ["纯木名片架"]})
```

- [ ] **Step 2: Run the focused tests and verify failure**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_load_old_category_config_returns_empty_prefixes test_core_regression.CoreRegressionTests.test_load_structured_category_config_reads_prefixes
```

Expected: fail with `ImportError` or `AttributeError` because `load_category_config_data` does not exist yet.

- [ ] **Step 3: Add the config data model and validators**

In `extract_orders.py`, add `dataclass` to the imports if it is not already present:

```python
from dataclasses import dataclass
```

Add this near `CATEGORY_KEYWORDS` and the config helpers:

```python
@dataclass
class CategoryConfigData:
    categories: dict[str, list[str]]
    prefixes: list[str]

    def copy(self) -> "CategoryConfigData":
        return CategoryConfigData(
            categories={category: list(keywords) for category, keywords in self.categories.items()},
            prefixes=list(self.prefixes),
        )


def copy_default_category_config_data() -> CategoryConfigData:
    return CategoryConfigData(copy_default_category_keywords(), [])
```

Add prefix validation helpers after `validate_category_config`:

```python
def validate_prefixes(data: Any) -> list[str]:
    if data is None:
        return []
    if not isinstance(data, list):
        raise ValueError("prefixes 必须是列表")
    prefixes: list[str] = []
    seen: set[str] = set()
    for item in data:
        prefix = str(item).strip()
        if prefix and prefix not in seen:
            prefixes.append(prefix)
            seen.add(prefix)
    return prefixes


def validate_category_config_data(data: Any) -> CategoryConfigData:
    if not isinstance(data, dict):
        raise ValueError("配置文件根节点必须是对象")
    if "categories" in data or "prefixes" in data:
        categories = validate_category_config(data.get("categories", {}))
        prefixes = validate_prefixes(data.get("prefixes", []))
        return CategoryConfigData(categories=categories, prefixes=prefixes)
    return CategoryConfigData(categories=validate_category_config(data), prefixes=[])
```

- [ ] **Step 4: Add load/save compatibility helpers**

Replace the current `save_category_config` implementation with this compatibility-preserving version:

```python
def save_category_config(config: dict[str, list[str]], config_path: str | Path | None = None) -> Path:
    path = Path(config_path).expanduser() if config_path else default_category_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    validated = validate_category_config(config)
    path.write_text(json.dumps(validated, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def save_category_config_data(config_data: CategoryConfigData, config_path: str | Path | None = None) -> Path:
    path = Path(config_path).expanduser() if config_path else default_category_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    validated = validate_category_config_data(
        {
            "categories": config_data.categories,
            "prefixes": config_data.prefixes,
        }
    )
    payload = {
        "prefixes": validated.prefixes,
        "categories": validated.categories,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
```

Add this new loader before the existing `load_category_config`:

```python
def load_category_config_data(
    config_path: str | Path | None = None,
    create_if_missing: bool = False,
) -> tuple[CategoryConfigData, str, str]:
    path = Path(config_path).expanduser() if config_path else default_category_config_path()
    if create_if_missing and not path.exists():
        try:
            save_category_config(copy_default_category_keywords(), path)
        except Exception as exc:
            return copy_default_category_config_data(), str(path), f"创建品类配置失败：{exc}"

    if not path.exists():
        return copy_default_category_config_data(), str(path), "品类配置文件不存在，已使用内置默认配置"

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return validate_category_config_data(data), str(path), ""
    except Exception as exc:
        return copy_default_category_config_data(), str(path), f"品类配置文件读取失败：{exc}，已使用内置默认配置"
```

Update `load_category_config` to call the new loader and return only the category map:

```python
def load_category_config(
    config_path: str | Path | None = None,
    create_if_missing: bool = False,
) -> tuple[dict[str, list[str]], str, str]:
    config_data, path, error = load_category_config_data(config_path, create_if_missing)
    return config_data.categories, path, error
```

- [ ] **Step 5: Run the config tests**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_load_old_category_config_returns_empty_prefixes test_core_regression.CoreRegressionTests.test_load_structured_category_config_reads_prefixes
```

Expected: both tests pass.

- [ ] **Step 6: Commit Task 1**

```powershell
git add -- extract_orders.py test_core_regression.py
git commit -m "feat: support structured category config"
```

## Task 2: Candidate Parsing Helpers

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add failing parser tests**

Add these tests after the config tests:

```python
    def test_extract_category_candidate_from_folder_name(self) -> None:
        from extract_orders import build_category_candidate_from_name

        candidate = build_category_candidate_from_name("18~19-0620-0625-方黑名片架-20单-3个", [])

        self.assertIsNotNone(candidate)
        self.assertEqual(candidate["raw_candidate"], "方黑名片架")
        self.assertEqual(candidate["prefix"], "")
        self.assertEqual(candidate["category"], "方黑名片架")

    def test_extract_category_candidate_cleans_separated_prefix(self) -> None:
        from extract_orders import build_category_candidate_from_name

        candidate = build_category_candidate_from_name("17-0625-WZY-纯木名片架-5单5个", ["WZY"])

        self.assertIsNotNone(candidate)
        self.assertEqual(candidate["raw_candidate"], "WZY-纯木名片架")
        self.assertEqual(candidate["prefix"], "WZY")
        self.assertEqual(candidate["category"], "纯木名片架")

    def test_extract_category_candidate_cleans_attached_prefix(self) -> None:
        from extract_orders import build_category_candidate_from_name

        candidate = build_category_candidate_from_name("14-0625-HAL小钢片-4单5个", ["HAL"])

        self.assertIsNotNone(candidate)
        self.assertEqual(candidate["raw_candidate"], "HAL小钢片")
        self.assertEqual(candidate["prefix"], "HAL")
        self.assertEqual(candidate["category"], "小钢片")

    def test_extract_category_candidate_ignores_structure_only_name(self) -> None:
        from extract_orders import build_category_candidate_from_name

        candidate = build_category_candidate_from_name("18~19-0620-0625-20单-3个", ["WZY"])

        self.assertIsNone(candidate)
```

- [ ] **Step 2: Run parser tests and verify failure**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_extract_category_candidate_from_folder_name test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_separated_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_attached_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_ignores_structure_only_name
```

Expected: fail because `build_category_candidate_from_name` does not exist.

- [ ] **Step 3: Implement candidate parsing**

Add these helpers after `first_detected_category_from_names`:

```python
def split_candidate_name_parts(name: str) -> list[str]:
    stem = Path(str(name)).stem
    normalized = re.sub(r"[＿_]+", "-", stem)
    normalized = re.sub(r"\s+", "", normalized)
    return [part.strip(" -—–~") for part in re.split(r"[-—–]+", normalized) if part.strip(" -—–~")]


def is_sequence_part(part: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(?:[~～]\d{1,3})?", part))


def is_date_part(part: str) -> bool:
    if parse_date_from_filename(part):
        return True
    return bool(re.fullmatch(r"\d{1,2}[.月]\d{1,2}(?:日)?", part))


def is_quantity_part(part: str) -> bool:
    return bool(re.fullmatch(r".*\d+\s*(?:单|个|件|pcs?|orders?|qty).*", part, flags=re.IGNORECASE))


def is_candidate_noise_part(part: str) -> bool:
    text = part.strip()
    if not text:
        return True
    return is_sequence_part(text) or is_date_part(text) or is_quantity_part(text)


def clean_candidate_with_prefix(raw_candidate: str, prefixes: list[str]) -> tuple[str, str]:
    candidate = raw_candidate.strip(" -—–")
    for prefix in sorted((item.strip() for item in prefixes if item.strip()), key=len, reverse=True):
        separated = f"{prefix}-"
        if candidate.startswith(separated):
            cleaned = candidate[len(separated):].strip(" -—–")
            return prefix, cleaned
        if candidate.startswith(prefix) and len(candidate) > len(prefix):
            cleaned = candidate[len(prefix):].strip(" -—–")
            return prefix, cleaned
    return "", candidate


def build_category_candidate_from_name(name: str, prefixes: list[str] | None = None) -> dict[str, str] | None:
    parts = split_candidate_name_parts(name)
    kept = [part for part in parts if not is_candidate_noise_part(part)]
    if not kept:
        return None
    raw_candidate = "-".join(kept).strip(" -—–")
    if not raw_candidate or not re.search(r"[^\d\s~～.\-—–]", raw_candidate):
        return None
    prefix, category = clean_candidate_with_prefix(raw_candidate, prefixes or [])
    if not category:
        return None
    return {
        "source_name": str(name),
        "raw_candidate": raw_candidate,
        "prefix": prefix,
        "category": category,
    }
```

- [ ] **Step 4: Run parser tests**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_extract_category_candidate_from_folder_name test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_separated_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_attached_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_ignores_structure_only_name
```

Expected: all pass.

- [ ] **Step 5: Commit Task 2**

```powershell
git add -- extract_orders.py test_core_regression.py
git commit -m "feat: parse new category candidates"
```

## Task 3: Thread Candidate Metadata Through Processing

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add failing processing test**

Replace the existing `test_unclassified_folder_excel_is_copied_and_excluded_from_output` body with assertions that also check candidate metadata:

```python
    def test_unclassified_folder_excel_records_category_candidate(self) -> None:
        folder = self.input_dir / "0607" / "1~2-0605-方白名片架-1单-1个"
        folder.mkdir(parents=True)
        workbook_path = folder / "random-order.xlsx"
        make_order_workbook(workbook_path, [("ORDER-UNCLASSIFIED", "SKU-UNKNOWN", 1)])

        result = self.run_tool(input_mode="folders")

        self.assertTrue(result["success"])
        self.assertEqual(result["total_rows"], 1)
        self.assertEqual(result["written_rows"], 0)
        self.assertFalse(self.output_path.exists())
        self.assertTrue((self.input_dir / UNCLASSIFIED_FOLDER_NAME / workbook_path.name).exists())
        candidates = result["category_candidates"]
        self.assertEqual(len(candidates), 1)
        self.assertEqual(candidates[0]["source_name"], "1~2-0605-方白名片架-1单-1个")
        self.assertEqual(candidates[0]["raw_candidate"], "方白名片架")
        self.assertEqual(candidates[0]["category"], "方白名片架")
        self.assertEqual(candidates[0]["status"], "待确认")
```

- [ ] **Step 2: Run the processing test and verify failure**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_unclassified_folder_excel_records_category_candidate
```

Expected: fail because `category_candidates` is not in the result.

- [ ] **Step 3: Add candidate selection helper**

Add this after `build_category_candidate_from_name`:

```python
def first_category_candidate_from_names(names: list[str], prefixes: list[str] | None = None) -> dict[str, str] | None:
    for name in names:
        candidate = build_category_candidate_from_name(name, prefixes or [])
        if candidate:
            return candidate
    return None
```

- [ ] **Step 4: Extend extraction function return values**

Update `extract_rows_from_workbook` to accept `category_prefixes` and return candidate metadata:

```python
def extract_rows_from_workbook(
    file_path: Path,
    category_keywords: dict[str, list[str]] | None = None,
    recognition_names: list[str] | None = None,
    category_prefixes: list[str] | None = None,
) -> tuple[list[dict[str, Any]], bool, str, str, list[dict[str, Any]], dict[str, str] | None]:
```

Inside the function, immediately after category detection:

```python
    category_candidate = None
    if category == "未分类":
        category_candidate = first_category_candidate_from_names(names, category_prefixes or [])
        if category_candidate:
            add_log(f"发现新品类候选：{category_candidate['category']}；来源：{category_candidate['source_name']}")
```

Update every return in the function to include `category_candidate` as the last item.

- [ ] **Step 5: Pass prefixes through run processing**

In `run_extract`, load config data instead of only categories:

```python
    category_config_data, category_config_file_path, category_config_error = load_category_config_data(category_config_path)
    category_keywords = category_config_data.categories
    category_prefixes = category_config_data.prefixes
```

Add `category_candidates: []` to the initial result/stat aggregation structures where `all_rows` and report lists are initialized:

```python
            category_candidates: list[dict[str, Any]] = []
```

Pass `category_prefixes` into `process_archive` and `process_folder_excel_group`. Update those function signatures to accept `category_prefixes: list[str] | None = None`, and pass it into `process_archive_excel_group`, `extract_rows_from_workbook`, and folder-group processing.

When an extraction call returns a candidate, append metadata:

```python
            if current_category_candidate:
                current_category_candidate = dict(current_category_candidate)
                current_category_candidate.update(
                    {
                        "archive_name": archive_path.name,
                        "excel_file": excel_file.name,
                        "source_path": str(excel_file),
                        "status": "待确认",
                    }
                )
                category_candidates.append(current_category_candidate)
```

Ensure each archive/folder result includes `"category_candidates": category_candidates`, and aggregate them in `run_extract`:

```python
                category_candidates.extend(result.get("category_candidates") or [])
```

Add to the final return dictionary:

```python
        "category_candidates": category_candidates,
```

- [ ] **Step 6: Run the processing test**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_unclassified_folder_excel_records_category_candidate
```

Expected: pass.

- [ ] **Step 7: Commit Task 3**

```powershell
git add -- extract_orders.py test_core_regression.py
git commit -m "feat: record unclassified category candidates"
```

## Task 4: Candidate Report Sheet

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add failing report test**

Add this test after the processing candidate test:

```python
    def test_process_report_contains_new_category_candidate_sheet(self) -> None:
        folder = self.input_dir / "0607" / "1~2-0605-方白名片架-1单-1个"
        folder.mkdir(parents=True)
        workbook_path = folder / "random-order.xlsx"
        make_order_workbook(workbook_path, [("ORDER-UNCLASSIFIED", "SKU-UNKNOWN", 1)])

        result = self.run_tool(input_mode="folders")

        report_path = Path(result["stats"]["process_report_path"])
        workbook = load_workbook(report_path, data_only=True)
        try:
            self.assertIn("新品类候选", workbook.sheetnames)
            rows = list(workbook["新品类候选"].iter_rows(values_only=True))
        finally:
            workbook.close()
        self.assertEqual(rows[0][:6], ("序号", "原始名称", "原始候选", "识别前缀", "候选品类", "状态"))
        self.assertEqual(rows[1][1], "1~2-0605-方白名片架-1单-1个")
        self.assertEqual(rows[1][4], "方白名片架")
```

- [ ] **Step 2: Run report test and verify failure**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_process_report_contains_new_category_candidate_sheet
```

Expected: fail because the report has no `新品类候选` sheet.

- [ ] **Step 3: Add candidate sheet to report**

In `save_process_report`, add this sheet entry before `品类汇总`:

```python
        (
            "新品类候选",
            ["序号", "原始名称", "原始候选", "识别前缀", "候选品类", "状态", "Excel文件名", "来源路径"],
            [
                {
                    "序号": index,
                    "原始名称": row.get("source_name", ""),
                    "原始候选": row.get("raw_candidate", ""),
                    "识别前缀": row.get("prefix", ""),
                    "候选品类": row.get("category", ""),
                    "状态": row.get("status", "待确认"),
                    "Excel文件名": row.get("excel_file", ""),
                    "来源路径": row.get("source_path", ""),
                }
                for index, row in enumerate(report_data.get("category_candidates") or [], 1)
            ],
        ),
```

When building `report_data` in `run_extract`, include:

```python
                "category_candidates": category_candidates,
```

- [ ] **Step 4: Run report test**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_process_report_contains_new_category_candidate_sheet
```

Expected: pass.

- [ ] **Step 5: Commit Task 4**

```powershell
git add -- extract_orders.py test_core_regression.py
git commit -m "feat: report new category candidates"
```

## Task 5: Save Confirmed Candidates

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Add failing save helper tests**

Add these tests after the config tests:

```python
    def test_save_confirmed_category_candidate_adds_prefix_and_category(self) -> None:
        self.category_config_path.write_text(
            json.dumps({"小钢片": ["小钢片"]}, ensure_ascii=False),
            encoding="utf-8",
        )

        from extract_orders import save_confirmed_category_candidate, load_category_config_data

        save_confirmed_category_candidate(
            self.category_config_path,
            prefix="HAL",
            category="方白名片架",
            keyword="方白名片架",
        )
        config_data, _, error = load_category_config_data(self.category_config_path)

        self.assertEqual(error, "")
        self.assertEqual(config_data.prefixes, ["HAL"])
        self.assertEqual(config_data.categories["方白名片架"], ["方白名片架"])
        self.assertIn("小钢片", config_data.categories)

    def test_save_confirmed_category_candidate_merges_without_duplicates(self) -> None:
        self.category_config_path.write_text(
            json.dumps(
                {
                    "prefixes": ["HAL"],
                    "categories": {"方白名片架": ["方白名片架"]},
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

        from extract_orders import save_confirmed_category_candidate, load_category_config_data

        save_confirmed_category_candidate(
            self.category_config_path,
            prefix="HAL",
            category="方白名片架",
            keyword="方白名片架",
        )
        config_data, _, _ = load_category_config_data(self.category_config_path)

        self.assertEqual(config_data.prefixes, ["HAL"])
        self.assertEqual(config_data.categories["方白名片架"], ["方白名片架"])
```

- [ ] **Step 2: Run save helper tests and verify failure**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_adds_prefix_and_category test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_merges_without_duplicates
```

Expected: fail because `save_confirmed_category_candidate` does not exist.

- [ ] **Step 3: Implement save helper**

Add this after config save/load helpers:

```python
def save_confirmed_category_candidate(
    config_path: str | Path | None,
    *,
    prefix: str = "",
    category: str,
    keyword: str = "",
) -> Path:
    category_text = str(category).strip()
    if not category_text:
        raise ValueError("候选品类不能为空")
    keyword_text = str(keyword or category_text).strip()
    prefix_text = str(prefix).strip()

    config_data, _, error = load_category_config_data(config_path, create_if_missing=True)
    if error and "已使用内置默认配置" not in error:
        raise ValueError(error)

    if prefix_text and prefix_text not in config_data.prefixes:
        config_data.prefixes.append(prefix_text)

    keywords = config_data.categories.setdefault(category_text, [])
    if keyword_text and keyword_text not in keywords:
        keywords.append(keyword_text)
    if category_text not in keywords:
        keywords.insert(0, category_text)

    return save_category_config_data(config_data, config_path)
```

- [ ] **Step 4: Run save helper tests**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_adds_prefix_and_category test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_merges_without_duplicates
```

Expected: pass.

- [ ] **Step 5: Commit Task 5**

```powershell
git add -- extract_orders.py test_core_regression.py
git commit -m "feat: save confirmed category candidates"
```

## Task 6: GUI Candidate Review Window

**Files:**
- Modify: `extract_orders_gui.py`
- Test: manual smoke test

- [ ] **Step 1: Store last run result for GUI access**

In `ExtractOrdersApp.__init__`, add:

```python
        self.last_result: dict[str, Any] = {}
```

In the code path that receives the successful `run_extract` result and updates reports, assign:

```python
        self.last_result = result
```

Use `rg -n "run_extract|update_report_buttons|replace_stats|format_stats" extract_orders_gui.py` to locate the existing completion handler and place the assignment next to other result state updates.

- [ ] **Step 2: Add review button to reports page**

In `build_reports_page`, add a button under `self.category_config_button`:

```python
        self.category_candidates_button = ctk.CTkButton(
            config_panel,
            text="新品类候选确认",
            height=38,
            fg_color="#2563eb",
            command=self.open_category_candidates,
        )
        self.category_candidates_button.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 10))
```

Move the existing logs/output/config-path widgets down by one row so they do not overlap.

- [ ] **Step 3: Add the review window class**

Add this class after `CategoryConfigWindow`:

```python
class NewCategoryCandidatesWindow(ctk.CTkToplevel):
    def __init__(self, master: ctk.CTk, config_path: Path, candidates: list[dict[str, Any]]) -> None:
        super().__init__(master)
        self.title("新品类候选确认")
        self.geometry("920x560")
        self.minsize(820, 500)
        self.transient(master)
        self.config_path = config_path
        self.candidates = [dict(candidate) for candidate in candidates]
        self.current_index: int | None = None
        self.status_var = tk.StringVar(value="")
        self.prefix_var = tk.StringVar(value="")
        self.category_var = tk.StringVar(value="")

        self._build_ui()
        self.refresh_candidates()
        self.grab_set()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="新品类候选确认", font=ctk.CTkFont(size=22, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text="保存后会写入 category_config.json，下一次提取自动使用。", text_color="#5b6675").grid(row=1, column=0, sticky="w", pady=(4, 0))

        body = ctk.CTkFrame(self, fg_color="#f6f8fb")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, corner_radius=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(left, text="候选列表", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))
        self.candidate_listbox = tk.Listbox(left, exportselection=False, activestyle="none", relief="flat", highlightthickness=1)
        self.candidate_listbox.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.candidate_listbox.bind("<<ListboxSelect>>", self.on_candidate_select)

        right = ctk.CTkFrame(body, corner_radius=8)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=12)
        right.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(right, text="识别前缀", font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        self.prefix_entry = ctk.CTkEntry(right, textvariable=self.prefix_var)
        self.prefix_entry.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))
        ctk.CTkLabel(right, text="候选品类", font=ctk.CTkFont(size=14, weight="bold")).grid(row=2, column=0, sticky="w", padx=12, pady=(4, 4))
        self.category_entry = ctk.CTkEntry(right, textvariable=self.category_var)
        self.category_entry.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 12))
        ctk.CTkButton(right, text="按前缀重新切分", command=self.apply_prefix_to_category).grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 10))
        ctk.CTkButton(right, text="保存为新品类", command=self.save_current_candidate).grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 10))
        ctk.CTkButton(right, text="忽略此候选", fg_color="#4b5563", command=self.ignore_current_candidate).grid(row=6, column=0, sticky="ew", padx=12, pady=(0, 10))

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(2, 18))
        footer.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(footer, textvariable=self.status_var, anchor="w", text_color="#5b6675").grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(footer, text="关闭", width=88, fg_color="#6b7280", command=self.destroy).grid(row=0, column=1, sticky="e")

    def refresh_candidates(self) -> None:
        selected = self.current_index
        self.candidate_listbox.delete(0, tk.END)
        for candidate in self.candidates:
            label = f"{candidate.get('status', '待确认')} | {candidate.get('category', '')} | {candidate.get('source_name', '')}"
            self.candidate_listbox.insert(tk.END, label)
        if self.candidates:
            index = selected if selected is not None and selected < len(self.candidates) else 0
            self.current_index = index
            self.candidate_listbox.selection_set(index)
            self.candidate_listbox.activate(index)
            self.load_candidate(index)
        else:
            self.current_index = None
            self.prefix_var.set("")
            self.category_var.set("")
            self.status_var.set("没有可确认的新品类候选")

    def load_candidate(self, index: int) -> None:
        candidate = self.candidates[index]
        self.prefix_var.set(str(candidate.get("prefix", "")))
        self.category_var.set(str(candidate.get("category", "")))
        self.status_var.set(f"来源：{candidate.get('source_name', '')}")

    def on_candidate_select(self, _event: tk.Event) -> None:
        selection = self.candidate_listbox.curselection()
        if not selection:
            return
        self.current_index = int(selection[0])
        self.load_candidate(self.current_index)

    def apply_prefix_to_category(self) -> None:
        if self.current_index is None:
            return
        raw_candidate = str(self.candidates[self.current_index].get("raw_candidate", ""))
        prefix = self.prefix_var.get().strip()
        if prefix and raw_candidate.startswith(prefix):
            self.category_var.set(raw_candidate[len(prefix):].strip(" -—–"))

    def save_current_candidate(self) -> None:
        if self.current_index is None:
            return
        category = self.category_var.get().strip()
        if not category:
            messagebox.showwarning("提示", "候选品类不能为空", parent=self)
            return
        candidate = self.candidates[self.current_index]
        prefix = self.prefix_var.get().strip()
        keyword = str(candidate.get("category") or candidate.get("raw_candidate") or category).strip()
        try:
            from extract_orders import save_confirmed_category_candidate

            save_confirmed_category_candidate(self.config_path, prefix=prefix, category=category, keyword=category)
        except Exception as exc:
            messagebox.showerror("保存失败", f"保存新品类失败：{exc}", parent=self)
            return
        candidate["prefix"] = prefix
        candidate["category"] = category
        candidate["status"] = "已保存"
        self.refresh_candidates()
        messagebox.showinfo("保存成功", "新品类和前缀规则已保存，下一次提取会自动使用。", parent=self)

    def ignore_current_candidate(self) -> None:
        if self.current_index is None:
            return
        self.candidates[self.current_index]["status"] = "已忽略"
        self.refresh_candidates()
```

- [ ] **Step 4: Add app opener method**

Add this method near `open_category_config`:

```python
    def open_category_candidates(self) -> None:
        candidates = list((self.last_result or {}).get("category_candidates") or [])
        if not candidates:
            messagebox.showinfo("新品类候选", "当前没有可确认的新品类候选。请先运行一次提取。", parent=self)
            return
        NewCategoryCandidatesWindow(self, category_config_path(), candidates)
```

- [ ] **Step 5: Run syntax check**

Run:

```powershell
python -m py_compile extract_orders_gui.py
```

Expected: no output and exit code 0.

- [ ] **Step 6: Manual GUI smoke test**

Run:

```powershell
python extract_orders_gui.py
```

Expected:

- Reports page shows `新品类候选确认`.
- Before any run, clicking it shows an info dialog.
- After a run with an unclassified folder such as `1~2-0605-方白名片架-1单-1个`, the window lists the candidate.
- Editing prefix/category and saving updates `category_config.json`.

- [ ] **Step 7: Commit Task 6**

```powershell
git add -- extract_orders_gui.py
git commit -m "feat: review category candidates in gui"
```

## Task 7: Documentation and Full Verification

**Files:**
- Modify: `README.md`
- Test: full regression suite

- [ ] **Step 1: Update README**

In the section `## 5. 品类关键词配置`, add this paragraph after the existing rule bullets:

```markdown
新品类候选确认：

- 如果文件夹名里出现未配置的新产品，例如 `方白名片架`，工具会先把它列为“新品类候选”。
- 未确认的候选仍按 `未分类` 处理，不写入正式汇总。
- 在 GUI 的“新品类候选确认”里可以填写或修改供应商/渠道前缀，例如 `WZY`、`HAL`，保存后会写入 `category_config.json`。
- 保存后的品类和前缀会在下一次提取时生效。
```

Also add a config-format note:

````markdown
新版配置可能包含：

```json
{
  "prefixes": ["HAL"],
  "categories": {
    "小钢片": ["小钢片"]
  }
}
```

旧版 `{ "品类": ["关键词"] }` 格式仍可继续读取。
````

- [ ] **Step 2: Run focused regression tests**

Run:

```powershell
python -m unittest test_core_regression.CoreRegressionTests.test_load_old_category_config_returns_empty_prefixes test_core_regression.CoreRegressionTests.test_load_structured_category_config_reads_prefixes test_core_regression.CoreRegressionTests.test_extract_category_candidate_from_folder_name test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_separated_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_cleans_attached_prefix test_core_regression.CoreRegressionTests.test_extract_category_candidate_ignores_structure_only_name test_core_regression.CoreRegressionTests.test_unclassified_folder_excel_records_category_candidate test_core_regression.CoreRegressionTests.test_process_report_contains_new_category_candidate_sheet test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_adds_prefix_and_category test_core_regression.CoreRegressionTests.test_save_confirmed_category_candidate_merges_without_duplicates
```

Expected: all pass.

- [ ] **Step 3: Run full suite**

Run:

```powershell
python -m unittest test_core_regression test_hc_filter
```

Expected: all tests pass.

- [ ] **Step 4: Inspect final diff**

Run:

```powershell
git diff -- README.md extract_orders.py extract_orders_gui.py test_core_regression.py
git status --short
```

Expected:

- Diff only contains candidate confirmation implementation, tests, and docs.
- Existing unrelated dirty files remain identifiable and are not accidentally staged unless they are part of this feature.

- [ ] **Step 5: Commit Task 7**

```powershell
git add -- README.md extract_orders.py extract_orders_gui.py test_core_regression.py
git commit -m "docs: describe category candidate confirmation"
```

## Self-Review

- Spec coverage: covered config compatibility, candidate extraction, prefix cleanup, GUI confirmation, saving, reporting, and tests.
- Scope: one feature touching shared extraction logic plus GUI review. No packaging or release rebuild is included.
- Known execution caution: the current worktree had unrelated uncommitted changes before this plan. During implementation, inspect `git status --short` before each commit and only stage files intended for that task.
