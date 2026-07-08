# Save Candidate To Existing Category Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an explicit GUI path for saving a new-category candidate keyword into an existing category.

**Architecture:** Add a small core helper in `extract_orders.py` that preserves structured config and appends a keyword to an existing category. Update `NewCategoryCandidatesWindow` to load existing category names, show a selector, and call that helper while mutating the live candidate status.

**Tech Stack:** Python, customtkinter/tkinter, JSON category config, pytest/unittest regression tests.

---

### Task 1: Core Config Helper

**Files:**
- Modify: `extract_orders.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Write failing tests for saving to an existing category**

Add tests near the existing `save_confirmed_category_candidate` tests:

```python
def test_save_candidate_keyword_to_existing_category_preserves_category_and_prefix(self) -> None:
    self.category_config_path.write_text(
        json.dumps(
            {
                "prefixes": ["HAL"],
                "categories": {
                    "心形钥匙扣": ["心形钥匙扣"],
                    "小钢片": ["小钢片"],
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    from extract_orders import save_candidate_keyword_to_existing_category, load_category_config_data

    save_candidate_keyword_to_existing_category(
        self.category_config_path,
        prefix="CBZ",
        target_category="心形钥匙扣",
        keyword="CBZ-心形钥匙扣扣",
    )

    config_data, _, error = load_category_config_data(self.category_config_path)
    self.assertEqual(error, "")
    self.assertEqual(config_data.prefixes, ["HAL", "CBZ"])
    self.assertEqual(config_data.categories["心形钥匙扣"], ["心形钥匙扣", "CBZ-心形钥匙扣扣"])
    self.assertNotIn("CBZ-心形钥匙扣扣", config_data.categories)

def test_save_candidate_keyword_to_existing_category_rejects_missing_category(self) -> None:
    self.category_config_path.write_text(
        json.dumps({"心形钥匙扣": ["心形钥匙扣"]}, ensure_ascii=False),
        encoding="utf-8",
    )

    from extract_orders import save_candidate_keyword_to_existing_category

    with self.assertRaises(ValueError):
        save_candidate_keyword_to_existing_category(
            self.category_config_path,
            target_category="不存在",
            keyword="CBZ-心形钥匙扣扣",
        )
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_preserves_category_and_prefix test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_rejects_missing_category
```

Expected: both fail because `save_candidate_keyword_to_existing_category` does not exist.

- [ ] **Step 3: Implement the core helper**

Add this function after `save_confirmed_category_candidate`:

```python
def save_candidate_keyword_to_existing_category(
    config_path: str | Path | None,
    *,
    prefix: str = "",
    target_category: str,
    keyword: str,
) -> Path:
    target_text = str(target_category).strip()
    if not target_text:
        raise ValueError("已有品类不能为空")
    keyword_text = str(keyword).strip()
    if not keyword_text:
        raise ValueError("候选关键词不能为空")
    prefix_text = str(prefix).strip()

    path = Path(config_path).expanduser() if config_path else default_category_config_path()
    existed_before = path.exists()
    config_data, _, error = load_category_config_data(path, create_if_missing=True)
    if error and existed_before:
        raise ValueError(error)
    if error and not existed_before and "已使用内置默认配置" not in error:
        raise ValueError(error)
    if target_text not in config_data.categories:
        raise ValueError(f"已有品类不存在：{target_text}")
    if prefix_text and prefix_text not in config_data.prefixes:
        config_data.prefixes.append(prefix_text)
    keywords = config_data.categories[target_text]
    if keyword_text not in keywords:
        keywords.append(keyword_text)
    return save_category_config_data(config_data, path)
```

- [ ] **Step 4: Run core tests**

Run:

```powershell
python -m pytest test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_preserves_category_and_prefix test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_rejects_missing_category
```

Expected: both pass.

### Task 2: GUI Selector And Save Action

**Files:**
- Modify: `extract_orders_gui.py`
- Test: `test_core_regression.py`

- [ ] **Step 1: Write failing GUI helper test**

Add a test near `test_category_config_window_preserves_prefixes_when_saving_keywords`:

```python
def test_candidate_window_save_to_existing_category_updates_live_candidate(self) -> None:
    import tkinter as tk
    import extract_orders_gui

    root = tk.Tk()
    root.withdraw()
    try:
        self.category_config_path.write_text(
            json.dumps({"心形钥匙扣": ["心形钥匙扣"]}, ensure_ascii=False),
            encoding="utf-8",
        )
        candidates = [{"prefix": "CBZ", "category": "CBZ-心形钥匙扣扣", "status": "待确认"}]
        window = object.__new__(extract_orders_gui.NewCategoryCandidatesWindow)
        window.config_path = self.category_config_path
        window.candidates = candidates
        window.current_index = 0
        window.prefix_var = tk.StringVar(root, value="CBZ")
        window.category_var = tk.StringVar(root, value="心形钥匙扣扣")
        window.existing_category_var = tk.StringVar(root, value="心形钥匙扣")
        window.refresh_candidates = lambda: None
        window.on_candidate_select = lambda _event=None: None
        window.candidate_listbox = type(
            "DummyListbox",
            (),
            {"selection_set": lambda *_: None, "activate": lambda *_: None, "see": lambda *_: None},
        )()

        window.save_to_existing_category()

        from extract_orders import load_category_config_data

        config_data, _, error = load_category_config_data(self.category_config_path)
        self.assertEqual(error, "")
        self.assertEqual(config_data.prefixes, ["CBZ"])
        self.assertEqual(config_data.categories["心形钥匙扣"], ["心形钥匙扣", "心形钥匙扣扣"])
        self.assertEqual(candidates[0]["status"], "已保存到已有品类：心形钥匙扣")
    finally:
        root.destroy()
```

- [ ] **Step 2: Run the GUI helper test to verify it fails**

Run:

```powershell
python -m pytest test_core_regression.py::CoreRegressionTests::test_candidate_window_save_to_existing_category_updates_live_candidate
```

Expected: fail because `existing_category_var` and `save_to_existing_category` are not implemented.

- [ ] **Step 3: Add category loading and selector state**

In `NewCategoryCandidatesWindow.__init__`, add:

```python
self.existing_categories = self.load_existing_categories()
self.existing_category_var = tk.StringVar(value=self.existing_categories[0] if self.existing_categories else "")
```

Add method:

```python
def load_existing_categories(self) -> list[str]:
    try:
        from extract_orders import load_category_config_data

        config_data, _, _ = load_category_config_data(self.config_path)
        return list(config_data.categories)
    except Exception:
        return []
```

- [ ] **Step 4: Add the selector and button**

In `_build_ui`, add a row after the candidate keyword entry:

```python
ctk.CTkLabel(right, text="已有品类").grid(row=4, column=0, sticky="w", padx=(14, 8), pady=(0, 10))
self.existing_category_menu = ctk.CTkOptionMenu(
    right,
    variable=self.existing_category_var,
    values=self.existing_categories or ["无可用品类"],
    height=36,
)
self.existing_category_menu.grid(row=4, column=1, sticky="ew", padx=(0, 14), pady=(0, 10))
if not self.existing_categories:
    self.existing_category_menu.configure(state=tk.DISABLED)
```

Move the actions frame to row 5 and status label to row 6.

Add the button inside the actions frame:

```python
ctk.CTkButton(actions, text="保存到已有品类", width=128, fg_color="#0f766e", command=self.save_to_existing_category).pack(side=tk.LEFT, padx=(8, 0))
```

- [ ] **Step 5: Implement the save action**

Add method:

```python
def save_to_existing_category(self) -> None:
    if self.current_index is None:
        return
    prefix = self.prefix_var.get().strip()
    keyword = self.category_var.get().strip()
    target_category = self.existing_category_var.get().strip()
    if not keyword:
        messagebox.showwarning("新品类候选", "候选关键词不能为空。", parent=self)
        return
    if not target_category or target_category == "无可用品类":
        messagebox.showwarning("新品类候选", "请选择已有品类。", parent=self)
        return
    try:
        from extract_orders import save_candidate_keyword_to_existing_category

        save_candidate_keyword_to_existing_category(
            self.config_path,
            prefix=prefix,
            target_category=target_category,
            keyword=keyword,
        )
    except Exception as exc:
        messagebox.showerror("保存失败", f"保存到已有品类失败：{exc}", parent=self)
        return
    self.update_current_candidate(status=f"已保存到已有品类：{target_category}")
    messagebox.showinfo("保存成功", f"候选关键词已保存到已有品类：{target_category}", parent=self)
```

- [ ] **Step 6: Run targeted tests**

Run:

```powershell
python -m pytest test_core_regression.py::CoreRegressionTests::test_candidate_window_save_to_existing_category_updates_live_candidate test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_preserves_category_and_prefix test_core_regression.py::CoreRegressionTests::test_save_candidate_keyword_to_existing_category_rejects_missing_category
```

Expected: pass.

### Task 3: Verification And Packaging

**Files:**
- Modify: `README.md` if needed to mention save-to-existing behavior.

- [ ] **Step 1: Run full regression and compile checks**

Run:

```powershell
python -m pytest
python -m py_compile extract_orders.py extract_orders_gui.py
```

Expected: all tests pass, compile exits 0.

- [ ] **Step 2: Rebuild v3.0 release assets**

Run:

```powershell
$env:SKIP_DEP_INSTALL='1'
$env:BUILD_RELEASE_ZIP='1'
$env:CODEX_NO_OPEN_EXPLORER='1'
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Expected: `dist\Excel订单数据提取工具_v3.0`, `dist\Excel-Tiqu-v3.0.zip`, and `dist\update.json` are regenerated.

- [ ] **Step 3: Verify update manifest hash**

Run:

```powershell
$manifest = Get-Content "dist\update.json" -Raw | ConvertFrom-Json
$actual = (Get-FileHash -LiteralPath "dist\Excel-Tiqu-v3.0.zip" -Algorithm SHA256).Hash.ToLowerInvariant()
if ($manifest.sha256 -ne $actual) { throw "sha mismatch" }
```

Expected: exits 0.

- [ ] **Step 4: Commit and upload**

Run:

```powershell
git add -- extract_orders.py extract_orders_gui.py test_core_regression.py README.md
git commit -m "feat: save candidates to existing categories"
git push origin master
gh release upload v3.0 "dist\Excel-Tiqu-v3.0.zip" "dist\update.json" --repo iSAc-K/Excel-Tiqu --clobber
```

Expected: commit, push, and release asset upload succeed.
