# Excel Tiqu 自动更新设计

## 1. 目标

为 Excel 订单数据提取工具增加与 Windows 文件整理助手一致的联网自动更新能力。

目标行为：

- 使用新 GitHub 仓库 `Excel-Tiqu` 发布版本。
- GUI 启动后异步检查 GitHub Releases 是否有新版本。
- 用户确认后下载发布 ZIP，校验 SHA-256，通过独立更新器安装。
- 安装前备份被替换文件，失败时自动回滚。
- 更新失败或无网络时不影响订单提取功能。

GitHub 更新清单地址：

```text
https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json
```

仓库显示名可使用 `Excel Tiqu`，实际仓库 slug 使用 URL 友好的 `Excel-Tiqu`。

## 2. 当前项目边界

当前项目是免安装 onedir Windows 工具：

```text
extract_orders.py        核心处理逻辑和 CLI
extract_orders_gui.py    CustomTkinter GUI 入口
build_exe.ps1            PyInstaller 打包脚本
category_config.json     用户可编辑品类配置
app_settings.json        GUI 上次路径和选项
README.md                使用说明
```

`build_exe.ps1` 当前生成 `dist/Excel订单数据提取工具_v2.1/`，并复制 `README.md`、`category_config.json`、`app_settings.json` 到发布目录。自动更新设计继续沿用 onedir 发布方式，不改成安装程序，不引入后台常驻服务。

## 3. 发布目录

更新后发布目录应包含：

```text
Excel订单数据提取工具_vX.Y/
|-- Excel订单数据提取工具_vX.Y.exe
|-- updater.exe
|-- VERSION.txt
|-- README.md
|-- category_config.json
|-- app_settings.json
|-- _internal/
```

`VERSION.txt` 是唯一版本来源。主程序、打包脚本、更新检查都从这里读取版本号。第一行格式使用 `v2.1` 或 `2.1` 均可，读取时去掉可选的 `v` 前缀。

## 4. 更新清单

GitHub Release 附带 `update.json`：

```json
{
  "version": "2.2",
  "download_url": "https://github.com/iSAc-K/Excel-Tiqu/releases/download/v2.2/Excel-Tiqu-v2.2.zip",
  "sha256": "<64 hex chars>",
  "notes": [
    "新增自动更新。"
  ]
}
```

清单规则：

- `version` 必须是语义化数字版本。
- `download_url` 必须是 HTTPS ZIP。
- `sha256` 必须是 64 位十六进制字符串。
- `notes` 可以是字符串或字符串列表，GUI 显示为更新说明。
- 版本比较使用语义化比较，不使用字符串排序。

## 5. 模块设计

### 5.1 `update_manager.py`

负责联网检查和下载，不操作安装目录。

职责：

- 定义 `UPDATE_MANIFEST_URL`。
- 读取并解析 `update.json`。
- 比较当前版本和线上版本。
- 下载 ZIP 到临时目录。
- 计算并校验 SHA-256。
- 汇报下载、校验进度。
- 支持下载和校验阶段取消。

主要接口沿用参考项目：

- `UpdateInfo`
- `DownloadProgress`
- `UpdateCancelled`
- `is_newer_version()`
- `parse_update_manifest()`
- `fetch_update_info()`
- `fetch_update_info_with_retry()`
- `verify_sha256()`
- `download_update()`

联网失败、清单格式错误、SHA 错误都作为普通错误返回给 GUI，不影响主程序启动和订单提取。

### 5.2 `updater.py`

负责真正替换文件。主程序不能覆盖正在运行的 EXE，所以安装由独立 `updater.exe` 完成。

职责：

- 等待主程序进程退出。
- 解压下载好的 ZIP 到临时目录。
- 拒绝绝对路径和 `..` 路径，防止 zip-slip。
- 复制新版本文件到安装目录。
- 替换前备份旧文件。
- 任一步骤失败时按已替换文件反向回滚。
- 成功后启动新主程序。
- 用小窗口显示等待、备份、安装、回滚、完成状态。

安装阶段不可取消，避免程序目录处于半替换状态。

### 5.3 `extract_orders_gui.py`

GUI 增加更新入口和状态机，仍保持订单提取主流程不变。

新增行为：

- 启动后延迟异步检查更新。
- 在侧边栏或“报告与配置”区域增加“软件更新”入口。
- 发现新版本后显示当前版本、最新版本、更新说明。
- 用户点击“立即更新”后开始下载。
- 下载和校验阶段显示进度、已下载大小、平均速度、预计剩余时间。
- 下载和校验阶段允许用户停止更新。
- 准备安装后启动 `updater.exe`，主程序退出。
- 正在处理订单时禁止开始更新。
- 更新下载中禁止开始新的订单处理。

主窗口和更新窗口通过队列或 `after()` 轮询通信，不从后台线程直接修改 Tk 控件。

## 6. 用户文件保护

更新器不得覆盖或删除：

```text
category_config.json
app_settings.json
logs/
backups/
```

说明：

- `category_config.json` 是用户可编辑品类配置。
- `app_settings.json` 保存 GUI 路径和选项。
- `logs/` 和 `backups/` 是运行结果和历史备份。
- 用户输出的 Excel 汇总文件通常在用户选择的输出目录，不属于程序目录替换范围。

更新器采用“复制 ZIP 中允许更新的文件”策略，不清空安装目录。

## 7. 打包与发布

`build_exe.ps1` 需要扩展：

1. 安装 PyInstaller 依赖。
2. 构建主程序 EXE。
3. 构建 `updater.exe`。
4. 复制 `README.md`、`category_config.json`、`app_settings.json`、`VERSION.txt` 到发布目录。
5. 确认发布目录包含主 EXE 和 `updater.exe`。

发布流程：

1. 更新 `VERSION.txt`。
2. 运行打包脚本。
3. 将发布目录压缩为 `Excel-Tiqu-vX.Y.zip`。
4. 计算 ZIP 的 SHA-256。
5. 生成 `update.json`。
6. 在 GitHub `iSAc-K/Excel-Tiqu` 创建 Release，上传 ZIP 和 `update.json`。
7. 验证公开地址 `releases/latest/download/update.json` 可读取，且版本和 SHA 与本地 ZIP 一致。

## 8. 错误处理

- 无网络或 GitHub 不可访问：GUI 保持可用，可显示“检查更新失败”。
- 清单格式错误：拒绝更新并显示原因。
- 下载中断：删除临时 ZIP，允许重新开始。
- 用户停止：删除临时 ZIP，不修改程序文件。
- SHA-256 不匹配：删除 ZIP，拒绝安装。
- 找不到 `updater.exe`：提示更新器缺失，不退出主程序。
- 安装失败：更新器回滚旧文件。
- 回滚失败：更新器保留备份目录路径供人工恢复。

## 9. 测试与验收

新增单元测试：

- 版本比较：`2.10` 大于 `2.9.9`，`2.1.0` 等于 `2.1`。
- manifest 校验：拒绝非 HTTPS、非 ZIP、错误 SHA。
- 下载进度：有 `Content-Length` 时显示百分比，无总大小时使用不确定进度。
- 下载取消：删除临时目录。
- 校验取消：删除完整但未安装的 ZIP。
- SHA 错误：拒绝安装并清理临时文件。
- 安装保护：`category_config.json`、`app_settings.json`、`logs/`、`backups/` 保持不变。
- zip-slip：包含 `../outside.txt` 的 ZIP 被拒绝。
- 安装失败回滚：旧文件恢复。

手工或烟雾验证：

- GUI 正常启动。
- “软件更新”入口能打开。
- 订单提取运行中点击更新会被阻止。
- 更新下载中无法开始订单提取。
- 打包目录包含主 EXE、`updater.exe`、`VERSION.txt`、配置文件。
- 使用测试 Release 执行一次真实更新，确认用户配置和日志保留。

## 10. 不在本次范围

- 不做静默自动安装。
- 不做后台常驻更新服务。
- 不做断点续传。
- 不做增量补丁。
- 不做测试版、灰度版或多更新频道。
- 不改订单提取核心规则。
- 不把用户配置迁移到 `%APPDATA%`。
