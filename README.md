# Excel 订单数据提取工具 v1.4

本工具用于批量处理压缩包或文件夹里的 Excel 订单文件，把订单数据按品类汇总到一个 Excel 文件中。普通用户建议使用图形界面，熟悉 PowerShell 的用户也可以使用命令行。

## 1. v1.4 新增功能

- 品类关键词可在 GUI 中查看、添加、修改、删除，并保存到 `category_config.json`。
- CLI 和 GUI 都使用同一份品类配置。
- 从正式 Excel 文件名识别预计单量和预计数量，例如 `0507-WZY-刀叉-13单18个.xlsx`。
- 每次运行都生成处理报告 Excel：`logs/处理报告_YYYYMMDD_HHMMSS.xlsx`。
- 写入旧汇总 Excel 前自动备份到 `backups/`，Dry Run 不备份、不写入。
- GUI 会记住上次输入路径、输出路径、数据来源、HC 过滤、Excel 规则、workers、clear、dry-run、重复检测等设置。
- 可选择只处理压缩包、只处理文件夹里的 Excel，或使用混合模式同时处理两类来源。
- HC 过滤默认关闭；勾选后，文件名包含 `HC`、`hc`、`Hc`、`hC` 的 Excel 会作为 HC 文件排除，不参与订单提取和汇总。
- 日期和品类优先从订单文件夹名识别，再看外层压缩包名/日期目录名，最后看 Excel 文件名。例如订单文件夹名包含 `0605` 和 `军牌钥匙扣` 时，会写入日期“6月5日”和“军牌钥匙扣”工作表。

## 2. 文件清单

```text
extract_orders.py        命令行入口和核心处理逻辑
extract_orders_gui.py    图形界面入口
build_exe.ps1            文件夹版 EXE 打包脚本
requirements.txt         Python 依赖
README.md                使用说明
category_config.json     品类关键词配置
app_settings.json        GUI 上次路径和设置
logs/                    处理日志和处理报告
backups/                 旧汇总 Excel 自动备份
```

## 3. GUI 使用方法

运行：

```powershell
python .\extract_orders_gui.py
```

操作步骤：

1. 选择输入路径。
2. 选择汇总 Excel 保存位置。
3. 展开处理选项，选择数据来源、Excel 规则、是否过滤 HC 文件。
4. 设置是否清空旧数据、是否 Dry Run、workers 数量、重复检测选项。
5. 如需调整品类，点击“品类关键词配置”。
6. 点击“开始提取”。
7. 完成后在日志中查看输出汇总 Excel、重复报告、处理报告、备份路径。

GUI 关闭时会自动保存上次路径和设置，下次打开会自动恢复。

## 4. CLI 使用方法

基本运行：

```powershell
python .\extract_orders.py --input "C:\订单压缩包" --output "C:\汇总.xlsx" --workers 4
```

清空旧汇总重新生成：

```powershell
python .\extract_orders.py --input "C:\订单压缩包" --output "C:\汇总.xlsx" --workers 4 --clear
```

Dry Run 预览：

```powershell
python .\extract_orders.py --input "C:\订单压缩包" --output "C:\汇总.xlsx" --workers 4 --dry-run
```

可选参数：

```text
--category-config   指定品类配置文件，默认 category_config.json
--report-dir        指定 logs 目录
--backup-dir        指定 backups 目录
--workers           同时处理压缩包数量，范围 1-8
--clear             先备份旧汇总，再清空重建
--dry-run           只扫描、校验、生成日志和处理报告，不写汇总、不备份
--input-mode        输入来源：archives 只处理压缩包，folders 只处理文件夹 Excel，mixed 混合模式
--enable-hc-filter  启用 HC 文件过滤，默认不启用
--excel-group-mode  Excel 规则：single 单文件订单模式，multi 多文件汇总模式
```

示例：

```powershell
python .\extract_orders.py --input "C:\订单文件夹" --output "C:\汇总.xlsx" --input-mode folders --excel-group-mode multi
python .\extract_orders.py --input "C:\订单文件夹" --output "C:\汇总.xlsx" --input-mode mixed --enable-hc-filter
```

## 5. 品类关键词配置

默认配置文件是：

```text
category_config.json
```

GUI 点击“品类关键词配置”后，可以：

- 新增品类
- 修改品类名
- 删除品类
- 新增、修改、删除关键词
- 保存配置
- 恢复默认配置

删除品类和恢复默认配置前会弹窗确认。保存成功后会提示。

识别规则：

- 文件名命中关键词后识别品类。
- 如果命中多个关键词，优先选择关键词长度最长的。
- 如果关键词长度一样，按 `category_config.json` 里的顺序选择靠前的品类。
- 未命中时归入 `未分类`。

## 6. 文件名单量 / 数量校验

工具会从正式 Excel 文件名中识别：

```text
13单18个
13单-18个
13单_18个
13单 18个
13单量18数量
13单18件
5单7套
5单7pcs
```

注意：只有带单位的数字才会识别，`0507-WZY-刀叉-13单18个.xlsx` 里的 `0507` 会作为日期，不会误判为单量或数量。

数量单位支持：`个`、`件`、`套`、`只`、`条`、`张`、`份`、`包`、`箱`、`袋`、`盒`、`对`、`双`、`支`、`根`、`瓶`、`罐`、`卷`、`片`、`台`、`把`、`枚`、`块`、`组`、`本`、`部`、`副`、`串`、`数量`、`pc`、`pcs`、`piece`、`pieces`。

为了避免误判，`月`、`日`、`年` 不作为数量单位；`单` 只作为单量单位，不作为数量单位。

日期支持从正式 Excel 文件名中识别这些写法：

```text
0507
0501-0503
4.17
45-4.17-CBZ
33~35-0418-20
24.4.21-CBZ
2026-05-07
```

识别后会写成 `4月17日`、`5月7日` 这类中文日期。

实际单量：

- 按“亚马逊订单号”非空的唯一订单号数量计算。
- 同一个订单号多行只算 1 单。

实际数量：

- 对“数量”列里可转成数字的值求和。
- 空值和无法转成数字的值不参与求和。
- 无法转成数字的行会写入处理报告的“异常明细”。

校验不通过不会阻止写入汇总 Excel，但会写入 GUI 日志、CLI 日志和处理报告。

## 7. 处理报告 Excel

每次运行都会生成：

```text
logs/处理报告_YYYYMMDD_HHMMSS.xlsx
```

包括这些 Sheet：

- `运行概览`
- `压缩包明细`
- `文件名校验`
- `异常明细`
- `重复明细`
- `品类汇总`

Dry Run 也会生成处理报告。Dry Run 的运行概览会明确显示：

```text
dry-run 模式：是
是否写入汇总 Excel：否
是否备份旧汇总 Excel：否
```

## 8. 自动备份

非 Dry Run 且本次确实有数据准备写入时：

- 如果输出汇总 Excel 已存在，会先备份到 `backups/`。
- `--clear` 模式也会先备份，再清空重建。
- 如果输出文件不存在，日志会显示“无需备份”。
- 如果本次所有压缩包都异常或没有可写入数据，不修改汇总 Excel，也不备份。

备份文件示例：

```text
backups/汇总_backup_20260509_153012.xlsx
```

## 9. 重复检测和异常

v1.1 功能继续保留：

- 多个压缩包批量处理。
- 单文件订单模式下，多个正式 Excel 时跳过。
- 多文件汇总模式下，同一处理单元里的多个正式 Excel 会逐个提取并汇总。
- `~$xxx.xlsx` 临时 Excel 自动忽略。
- 重复订单号只提示。
- 完全重复行按设置跳过。
- 重复报告自动生成。
- Sheet 标签颜色保留。
- GUI 可导出异常列表 Excel。

## 10. HC 文件过滤

HC 过滤默认关闭。关闭时，文件名包含 `HC`（不区分大小写）的 Excel 会按普通 Excel 参与处理。

在 GUI 勾选“过滤 HC 文件”，或命令行使用 `--enable-hc-filter` 后，工具会把这类 Excel 视为 HC 文件：

- 不参与订单提取。
- 不参与数量统计。
- 不参与重复检测。
- 不写入汇总 Excel。
- 正式运行时会复制到输入目录下的 `HC` 文件夹，原始压缩包或原始文件夹不会被改写。
- Dry Run 只在日志和处理报告中显示预计复制路径，不创建 `HC` 文件夹、不复制文件。

判断规则只看 Excel 文件名本身，不看父级文件夹名。处理报告会生成 `HC文件明细` Sheet，记录来源文件、目标路径和复制状态。

## 11. v1.4 后续修复

本版继续增强了几类常见源文件情况：

- 支持读取合并单元格。如果“亚马逊订单号”合并了 2 行、10 行或更多行，合并区域内每一行都会读取左上角订单号。
- 表头识别和数据读取都会按合并单元格取值。
- 同一订单号下相同 SKU 会在写入汇总前自动合并，数量自动相加；同一订单号下不同 SKU 会保留多行。
- 重复检测按“亚马逊订单号 + SKU”判断，同一订单号下不同 SKU 属于正常多商品订单，不再计入重复异常。
- 一个压缩包内如果有多个订单子文件夹，会按子文件夹拆分处理；只有某个子文件夹内多个正式 Excel 或没有正式 Excel 时，才跳过该子文件夹。
- 数量支持更多文本格式，例如 `数量：1`、`1个`、`共2件`、`Qty 3`、`QTY: 5`、`quantity 6`、`x4`、`X10`。
- 数量无法转换时保留原值，并在处理报告的“异常明细”里记录原始数量、文件名和行号。
- 压缩包或正式 Excel 文件名以“修改”或“售后”开头时才会跳过；文件名中间出现“修改”“售后”不会跳过。
- “补发”不再作为跳过条件，补发订单会按普通订单正常提取。
- 日期和品类识别优先使用订单文件夹名，再使用外层压缩包名或日期目录名，最后使用 Excel 文件名。
- 最终仍识别为“未分类”的 Excel 不写入正式汇总，会复制到输入目录下的 `未分类Excel` 文件夹；Dry Run 只预览，不实际复制。

## 12. 打包 EXE

运行：

```powershell
.\build_exe.ps1
```

打包完成后会生成文件夹版 EXE：

```text
dist/Excel订单数据提取工具_v1.4/
```

该目录会包含：

```text
Excel订单数据提取工具_v1.4.exe
README.md
category_config.json
app_settings.json
```

用户可编辑的配置文件放在 EXE 同级目录，不会放到 PyInstaller 临时目录。

## 13. 如何验证 v1.4 是否正常

验证 CLI：

```powershell
python .\extract_orders.py --input "C:\订单压缩包" --output "C:\汇总.xlsx" --workers 4
```

验证 GUI：

```powershell
python .\extract_orders_gui.py
```

查看处理报告：

- 运行后打开 `logs/处理报告_YYYYMMDD_HHMMSS.xlsx`。
- 检查 `运行概览`、`压缩包明细`、`文件名校验`。

查看备份文件：

- 先准备一个已经存在的汇总 Excel。
- 再正常运行一次。
- 检查 `backups/` 是否出现 `原文件名_backup_时间.xlsx`。

测试 Dry Run：

```powershell
python .\extract_orders.py --input "C:\订单压缩包" --output "C:\汇总.xlsx" --workers 4 --dry-run
```

确认：

- 不写入汇总 Excel。
- 不生成备份。
- 会生成日志和处理报告。

测试品类配置：

1. 打开 GUI。
2. 点击“品类关键词配置”。
3. 新增一个品类和关键词并保存。
4. 准备一个文件名包含该关键词的 Excel 压缩包。
5. 再次运行，检查汇总 Excel 和处理报告中是否识别为新品类。

测试 `13单18个` 校验：

1. 准备正式 Excel 文件名：`0507-WZY-刀叉-13单18个.xlsx`。
2. 让文件里实际唯一订单数为 13，数量合计为 18。
3. 运行工具。
4. 打开处理报告的 `文件名校验` Sheet，确认单量和数量均为“匹配”。
5. 修改数据让实际单量或数量不一致，再运行，确认报告显示“不匹配”。

## 13. 常见问题

如果 `.rar` 或 `.7z` 处理失败：

- 请安装 WinRAR。
- 工具会优先调用电脑上的 WinRAR。

如果提示找不到 `openpyxl`：

```powershell
python -m pip install -r requirements.txt
```

如果品类配置损坏：

- GUI 会提示读取失败。
- 工具会自动回退到内置默认品类。
- 可以在 GUI 里点击“恢复默认配置”后保存。

如果没有生成备份：

- 检查是否为 Dry Run。
- 检查输出汇总 Excel 是否原本存在。
- 检查本次是否确实有可写入数据。
