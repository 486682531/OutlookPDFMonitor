# Outlook PDF 自动监控器

> 当前版本：**v1.0.10**（构建日期 2026-08-20）

自动监控你的 **Outlook / Hotmail 个人邮箱**，把新邮件里的 PDF 附件下载到本地。

## 当前方案：Microsoft Graph API + OAuth2

- 通过 **Microsoft Graph REST API**（`/me/mailFolders/inbox/messages`）读取收件箱，
  不再使用 IMAP / 应用专用密码（微软已于 2026 年禁用 Outlook.com 个人账户的 IMAP 基本认证）。
- 认证使用 **OAuth2 访问令牌（Bearer）**，scope = `Mail.Read` + `offline_access` + `User.Read`。
- 令牌由 MSAL **公共客户端流**管理并本地缓存（`config/token_cache.bin`），自动刷新，
  **绝不保存邮箱密码、绝不保存明文令牌**；日志中禁止出现令牌 / 密码。
- 已内置一个 Azure 桌面应用（公共客户端）的**客户端 ID**，一般情况无需自行注册 Azure 应用；
  如需替换，可在【设置】页「OAuth 客户端 ID」填写并重启程序。

### 方案演进（简述）
最初用本机 Outlook COM/MAPI → 改用 Graph → 改用 IMAP + 应用专用密码 → IMAP + XOAUTH2，
最终因微软禁用 IMAP 基本认证而定稿为 **Graph API + OAuth2**（最简洁、最贴近微软官方推荐路径）。

## 特性

- 多邮箱支持：每个邮箱独立监控、独立去重、单邮箱失败不影响其他邮箱
- 最小权限：**只读**访问收件箱，绝不发送 / 删除 / 移动邮件
- 认证安全：OAuth2 令牌本地缓存自动刷新；日志中禁止出现令牌 / 密码
- 按邮箱分文件夹保存、重名自动加序号、失败自动重试（默认 3 次）
- 任务状态管理（正常 / 失败 / 需授权 / Token 过期 / 已禁用）、系统托盘常驻、可选开机自启、详细日志
- `PDF_DOWNLOADED` 事件接口（为第二阶段【报关单智能分析工具】预留）

## 环境要求

- Windows 10 / 11
- **不需要**安装 Microsoft Outlook 桌面版
- 一次性的浏览器登录授权（标准微软账户登录，需能访问 login.microsoftonline.com）

## 第一次使用：OAuth 授权（约 1 分钟）

> 不再使用账户密码或「应用专用密码」——改为在浏览器中完成 OAuth2 授权。

1. 打开软件 → 【邮箱】页点「**添加邮箱**」，输入邮箱地址（如 `xxx@outlook.com`）
2. 程序自动弹出**微软登录页**（系统默认浏览器），用该邮箱登录并点击「接受」
   （此处输入的密码只在微软页面，本程序绝不经手）
3. 授权成功后状态变为「已授权未测试」；回列表**勾选「启用」**
4. 点「**测试**」验证连接；连接正常后状态显示「正常」
5. 之后即可在后台自动下载新邮件中的 PDF

> 也可在【设置】页的「Microsoft Graph OAuth」分组里点「授权 / 重新授权 / 注销授权」，
> 并实时查看授权状态（未授权 / 已授权 / Token 过期）。

## 运行（绿色便携，支持任意安装目录）

**双击 `启动.bat` 即可**，无需安装 Python、无需安装依赖、不依赖任何开发机路径。

- 软件包已内置完整的运行环境：`runtime\`（Python 解释器）与 `venv\`（依赖包）。
- 可放到任意目录（如 `D:\软件\Outlook PDF自动监控器`），从该目录直接启动即可，
  **不要**把 `runtime`、`venv`、`启动.bat`、`main.py` 分开移动。
- 启动脚本使用相对路径（`%~dp0`），不写死任何用户名或绝对路径。

若弹出提示 **「未检测到运行环境（runtime / venv 缺失）」**：

> 说明软件包不完整（解压时被拦截或只复制了部分文件）。
> 请重新解压**完整的**软件包，或联系管理员获取安装包，不要直接运行 Python 启动器。

排错：

- 看不到界面：确认 `runtime\pythonw.exe` 与 `venv\Scripts\pythonw.exe` 均存在；
  用 `启动调试.bat` 以控制台方式启动，可看到具体报错。
- 需要源码运行调试：`venv\Scripts\python.exe main.py`（同样走包内环境）。

## 配置与使用

1. 打开软件 → 【设置】页可选：监控间隔（30 / 60 / 120 / 300 秒）、失败重试次数、开机自启
2. 【邮箱】页点「**添加邮箱**」→ 输入邮箱地址 → 自动弹出微软登录页完成授权
   （令牌本地缓存于 `config/token_cache.bin`，绝不存明文）
3. 勾选「启用」→ 点「**测试**」验证；连接正常后状态显示「正常」
4. 若令牌过期，状态变「Token 过期」，点「授权」或「重新授权」重新登录即可，无需重启
5. 监控默认每 60 秒检查一次
6. 点「**注销授权**」可清除本地所有 OAuth 令牌（下次使用需重新授权）

## 去重机制

唯一键：**邮箱 + message_id(Graph 邮件 id) + attachment_id(Graph 附件 id)**。
已成功下载过的附件绝不重复下载。
数据库：`data/outlookauto.db`。旧版（`entry_id + attachment_index`）结构会在首次
启动时**自动迁移**为新键，无需手动处理。

## 安全

- **绝不保存明文密码 / 令牌**：仅在本地缓存 MSAL 刷新令牌（`config/token_cache.bin`），
  由 MSAL 负责加密与自动续期
- 仅**只读**访问邮箱（Graph 仅请求 `Mail.Read`），不发送 / 删除 / 移动任何邮件
- 日志中**禁止出现密码 / 令牌**；任何异常信息都不回显敏感凭据

## 自动更新（OTA）

程序支持通过 **GitHub Releases** 一键在线更新。

- 启动时（约 8 秒后）自动检查一次，之后每 6 小时检查一次；发现新版本会在**系统托盘**弹出提示「更新并重启」。
- 更新方式：在【设置】页「更新」分组填写 GitHub 更新源 `owner/repo`（例如 `486682531/OutlookPDFMonitor`），点「保存」即生效；也可点「检查更新」手动触发。
- 更新包为新版整个项目源码压缩包（`source.zip` 形态），由独立的 `update_apply.py` 在后台解压覆盖（**保留** `venv / runtime / config / data / logs / downloads` 等你的运行环境与数据），并校验 sha256 后无窗口重启。
- 校验失败会自动中止，**不改动任何文件**，可放心更新。

> 当前版本已内置默认更新源 `486682531/OutlookPDFMonitor`，无需手动设置即可接收后续更新。

## 目录结构

```
main.py                程序入口
graph_source.py        Graph 邮件源（OAuth2 Bearer 令牌，REST 读取收件箱）
oauth_manager.py       OAuth2 令牌管理（MSAL 公共客户端，令牌本地缓存）
database.py            SQLite 持久化（邮箱 / 去重 / 日志）
pdf_downloader.py      PDF 下载、去重、重名、重试
monitor_engine.py      多邮箱轮询引擎（线程池 + 错误隔离）
event_bus.py           PDF_DOWNLOADED 事件
config.py              配置（oauth_client_id / 监控参数 / update_repo）
updater.py             OTA 更新检查器（GitHub Releases latest）
update_apply.py        独立更新脚本（后台解压覆盖 + 校验 sha256 + 重启）
make_release.py        打包脚本（产出 source.zip 与 source.sha256）
ui/                    首页 / 邮箱 / 下载设置 / 日志 / 设置
requirements.txt       PySide6 / msal / pyinstaller
test_core.py           核心逻辑测试（MockGraphSource，覆盖授权/读取/解析/去重/重试等）
test_gui_smoke.py      GUI 冒烟测试
```

## 测试

```
venv\Scripts\python.exe test_core.py
venv\Scripts\python.exe test_gui_smoke.py
```

`test_core.py` 使用 `MockGraphSource`（不依赖网络 / OAuth / PySide6），
覆盖场景：OAuth 授权状态、Graph 消息读取模拟、PDF 附件解析、重复邮件、下载失败重试、
配置 / 数据库 / 去重 / 重名 / 分类 / 多邮箱 / 时间窗 / 网络异常 / 令牌失效 / 程序重启恢复。

## 在此开发机无法验证、需你在自己电脑验收的部分

- 真实 Graph 调用（需你的邮箱完成 OAuth2 授权，网络可达 graph.microsoft.com）
- 实际邮件读取与附件下载

验收步骤：添加邮箱 → 浏览器授权 → 勾选启用 → 点「测试」→ 观察【日志】页与
`downloads/` 下是否出现按邮箱分类的 PDF 文件。
