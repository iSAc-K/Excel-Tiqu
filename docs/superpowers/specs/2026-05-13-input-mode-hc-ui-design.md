# 输入来源、HC 过滤与多文件汇总模式设计

日期：2026-05-13

## 背景

当前工具主要面向压缩包输入：GUI 让用户选择“压缩包所在文件夹”，核心 `run_extract(...)` 会递归查找压缩包，逐个解压并提取其中的 Excel 数据。现有 HC 过滤功能在核心层默认开启：只要 Excel 文件名包含 `HC`，就从正常提取流程中排除，并在正式运行时复制到输入根目录下的 `HC` 文件夹。

本次改动要把 HC 过滤从默认行为改成用户可选项，并新增三种输入来源选择：只处理压缩包、只处理文件夹里的 Excel、混合模式。同时新增“多文件汇总模式”，允许同一处理单元内存在多个正式 Excel，并把这些 Excel 的数据汇总到同一个输出工作簿。

## 用户可见功能

GUI 起始页新增三组选择：

1. 数据来源
   - 只处理压缩包
   - 只处理文件夹里的 Excel
   - 混合模式

2. HC 过滤
   - 过滤 HC 文件
   - 第一次默认关闭
   - 之后从 `app_settings.json` 读取并记住上次选择

3. Excel 处理规则
   - 单文件订单模式
   - 多文件汇总模式

数据来源决定扫描哪些文件：

- 只处理压缩包：递归查找 `.zip`、`.rar`、`.7z`，沿用现有解压处理流程。
- 只处理文件夹里的 Excel：递归查找 `.xlsx`、`.xlsm`，忽略 `~$` 临时文件和 `.xls`。
- 混合模式：同时查找压缩包和普通 Excel 文件，两类数据统一进入同一个汇总流程。

Excel 处理规则决定同一处理单元里多个正式 Excel 如何处理：

- 单文件订单模式：沿用现有规则。同一处理单元中如果有多个正式 Excel，则跳过该单元并记录异常。
- 多文件汇总模式：同一处理单元中允许多个正式 Excel，逐个提取，最后统一进入分类、合并、重复检测和输出写入流程。

多文件汇总模式同时适用于压缩包内部和普通文件夹 Excel。

## 核心接口

`run_extract(...)` 增加三个参数：

```python
input_mode: str = "archives"       # archives / folders / mixed
enable_hc_filter: bool = False
excel_group_mode: str = "single"   # single / multi
```

默认值保持命令行和旧调用方尽量兼容：

- `input_mode="archives"` 保持原来的压缩包扫描入口。
- `enable_hc_filter=False` 符合本次“HC 过滤不是默认项”的要求。
- `excel_group_mode="single"` 保持原来的单文件校验规则。

命令行入口同步增加参数：

- `--input-mode archives|folders|mixed`
- `--enable-hc-filter`
- `--excel-group-mode single|multi`

GUI 调用 `run_extract(...)` 时只传递用户选择，不在 UI 层重新实现扫描、HC 判断、Excel 提取或合并逻辑。

## 数据流

核心层拆出输入发现阶段：

1. 根据 `input_mode` 查找压缩包任务。
2. 根据 `input_mode` 查找普通文件夹 Excel 任务。
3. 压缩包任务继续走 `process_archive(...)`。
4. 文件夹 Excel 任务走新增的文件夹处理函数，复用 `process_excel_unit(...)`。
5. 两类任务产出的 rows、异常、报告明细、HC 明细合并到 `run_extract(...)` 的统一后处理流程。
6. 后处理继续使用现有的同订单同 SKU 合并、重复检测、重复报告、处理报告、Dry Run、备份和写入逻辑。

普通文件夹 Excel 的处理单元按直接父文件夹分组。这样递归扫描时，一个订单子文件夹可以作为一个自然单元处理；如果 Excel 直接放在输入根目录，则输入根目录本身作为一个处理单元。

## HC 过滤规则

HC 过滤由 `enable_hc_filter` 控制：

- 关闭时：文件名包含 `HC` 的 Excel 不再自动排除，按普通 Excel 处理。
- 开启时：只用 `path.name.casefold()` 判断文件名，不看父级文件夹。
- 压缩包中的 HC 文件：保持现有排除和复制规则。
- 普通文件夹中的 HC 文件：同样排除并复制到输入根目录下的 `HC` 文件夹。
- Dry Run 下只记录预计复制路径，不创建 `HC` 文件夹，不复制文件。
- 即使 HC 复制失败，文件也保持排除，不回落到正常提取流程。

## 报告与日志

报告需要区分来源类型，避免把普通文件夹 Excel 误显示成压缩包数据：

- 压缩包任务：继续填写现有压缩包字段。
- 文件夹 Excel 任务：来源类型标记为“文件夹Excel”，压缩包名称可为空或显示为对应父文件夹名，报告中保留 Excel 文件名、所在文件夹、处理状态、提取行数、异常原因。
- HC 明细继续写入 `HC文件明细` sheet，并覆盖压缩包和普通文件夹两类来源。

统计面板继续显示总任务数、完成数、异常数、写入行数、重复数和 HC 数。总任务数在文件夹模式下按处理单元计数，在混合模式下为压缩包任务数加文件夹 Excel 处理单元数。

## 设置保存

`app_settings.json` 新增字段：

```json
{
  "input_mode": "archives",
  "enable_hc_filter": false,
  "excel_group_mode": "single"
}
```

第一次没有这些字段时，使用上述默认值。用户运行或关闭窗口时保存当前选择，下次打开 GUI 时恢复。

## 测试与验收

最低验证项：

1. `python -m py_compile extract_orders.py extract_orders_gui.py`
2. 现有 HC 回归测试更新为默认关闭和开启两组断言。
3. 压缩包模式：旧样例仍按原逻辑处理，除非用户开启 HC 过滤。
4. 文件夹模式：递归扫描普通 Excel，能提取数据并写入汇总表。
5. 混合模式：同一次运行同时处理压缩包和普通文件夹 Excel。
6. 单文件订单模式：同一处理单元多个正式 Excel 仍跳过并记录异常。
7. 多文件汇总模式：同一处理单元多个正式 Excel 能逐个提取并汇总。
8. Dry Run：不写输出工作簿，不创建 HC 文件夹，不复制 HC 文件。
9. GUI：能保存并恢复数据来源、HC 过滤、多文件汇总模式选择。

打包后仍按现有 Windows onedir 交付方式验证，确保 `README.md`、`category_config.json`、`app_settings.json` 和 EXE 在 `dist` 目录中保持同步。
