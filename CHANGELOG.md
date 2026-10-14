# Changelog

> 说明：`0.1.7` 到 `0.1.46` 为根据 git 历史补录（仓库没有 tag，按 `发布: 0.1.x 版本` 提交划分版本边界，日期取发布提交日期）。每条只写 commit 内容能支撑的要点，粒度较粗。

## 0.2.0 - 2026-10-14

Breaking changes：

- **移除 EWS（exchangelib）后端**，项目现在是 Graph-only。`provider: "ews"` 配置、EWS 相关环境变量（`OUTLOOK_EMAIL`、`OUTLOOK_AUTH_MODE`、`OUTLOOK_SERVER`、`OUTLOOK_ACCESS_TYPE`、`OUTLOOK_CLIENT_*`、`OUTLOOK_USERNAME` / `OUTLOOK_PASSWORD`）和 `client.py` 全部删除。迁移指引：仍在使用 EWS（Exchange on-prem / 企业租户走 exchangelib）的用户请停留在 `0.1.48`，不要升级。
- **HTTP 模式改为单用户自托管自动 provisioning**：server 启动时自动创建单个 credential，account 取 `outlook-mcp.json` 的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`；未设置时用 `secrets.token_urlsafe(24)` 生成随机 token 并在启动日志打印 setup URL（`http://{OUTLOOK_MCP_HTTP_HOST}:{OUTLOOK_MCP_HTTP_PORT}/graph/{email}/{token}`）。不再需要手工执行 `credential_store register`；register/rotate CLI 降级为运维恢复工具。非 base config `email` 的 credential 按未注册处理（404 页面 / 403 `auth_required`）。
- Bearer token 移除最小长度限制（`0.1.47` 曾默认要求 `16 characters`）；字符集仍为 `A-Za-z0-9._~-`，公网部署仍应使用高熵 token。

其他：

- `OUTLOOK_MCP_DB` 语义明确为本地状态库（msal cache、日历草稿、操作日志、上传元数据），不再是多账号凭据库。
- `start_outlook_mcp_http.sh` / `.bat` 删除显式 register 步骤，只设环境变量并启动。

## 0.1.48 - 2026-09-09

修复与健壮性（来自一次全项目体检，commit a715ec6）：

- 修复 HTTP API `contact/list`、`contact/search`、`task/list` 在 mode/tools 过滤下误报 `tool_disabled`（工具名改从函数 `__name__` 推导）。
- HTTP API 未知参数由 502 `exchange_error` 改为 400 `validation_error`。
- Graph `task_update` 支持 `complete=False` 重开任务；Graph 附件补 3 MB/文件 + 3.5 MB 总量预算校验，与 EWS 行为对齐。
- 修复 SQLite 凭据库 0600 首次启动未生效，并覆盖 `-wal`/`-shm` sidecar；`redact_secrets` 覆盖 `outlook-mcp.json` 加载的凭据；pending flow/binding/quarantine/uploads 文件权限统一 0700/0600；`redact_text` 对 JSON 敏感字段值结构化打码。
- EWS 配置校验 email 格式；联系人 phone label 两端共用白名单；4 处 `timeout=30` 收敛为 `GRAPH_HTTP_TIMEOUT_SECONDS` 常量；413/400 判断改用 `RequestBodyTooLargeError` 专用异常。
- Dockerfile 基础镜像 ARG 化（`--build-arg BASE_IMAGE=...` 可覆盖，默认仍 daocloud）；`PIP_INDEX_URL` 不再残留到容器运行时；README 去掉写死的 `D:\PythonWorkspace` 路径。

测试：

- 新增 15 条回归用例，含 31 个 MCP 工具在 server 注册、HTTP dispatch 表、mode presets 三处的一致性测试（522 passed）。

## 0.1.47 - 2026-09-09

本次包含 0.1.46 之后 6 个未记录的 commit（d5c31ac 到 f6e5aef）与一批新的安全/健壮性修复。

行为变化（用户可见）：

- Bearer token 默认最小长度从 `4` 提升到 `16 characters`（`OUTLOOK_MCP_SECRET_MIN_LENGTH` 可调）；文档全部对齐。
- 上传增加 per-credential 总量配额：默认 `30 MiB` / `100` 个文件（`OUTLOOK_MCP_UPLOAD_QUOTA_BYTES` / `OUTLOOK_MCP_UPLOAD_QUOTA_FILES`）。
- 重复认证失败按来源 IP 限速（burst 10 次，之后每秒 1 次），超限返回 HTTP `429 rate_limited`。
- 日历草稿提交成功但本地清理冲突时，返回 `submitted=true` + `draft_cleanup="pending"`，不再误报 409 导致重复提交。

修复与健壮性：

- MSAL token 刷新与 device flow 轮询增加 30 秒网络超时（`TimeoutSession`），不再永久阻塞；device flow 轮询移出全局 token 锁，并发请求同账号不再互相阻塞。
- `auth_setup` token cache 改原子写；`verify_graph_token` 自建 Session 正确关闭；`CredentialStore.close()` 停止 flush 线程并注销 atexit，CLI register/rotate 与 `create_app` 重建时不再累积线程。
- 删除必然 401 的 4 段式 secret uploads 死路由；endpoint account 增加 email 格式校验；uvicorn access log 关闭，避免 URL 中的 secret 进入 Docker logs。
- FastMCP 实例 host 从 `0.0.0.0` 收紧为 `127.0.0.1`（HTTP 暴露仍由 http_server 包装层负责 bearer 认证）。
- `start_outlook_mcp_http.bat` 的 `OUTLOOK_MCP_SKIP_PROMPT` 语义与 sh 对齐（仅值为 `1` 时跳过）。
- `.gitignore` 补 `*.sqlite3*`；`.env.example` 补全 HTTP/上传/启动器环境变量。

此前 6 个 commit 补录：

- `.gitattributes` 固定行尾（源码 LF、bat/ps1 CRLF）。
- 后端安全：token cache/SQLite 权限 0o600、cache 原子写、损坏 cache 隔离备份、删除 legacy `graph_oauth_state` 表、ICS 按 RFC 5545 折行。
- 工具层：`mail_forward` 修复 html 参数被忽略、写工具 confirm 校验顺序统一、删除/标记操作返回 alias、`task_update` 支持取消完成、base64 附件大小校验、上传保留中文文件名。
- HTTP 层：API 错误按类别映射状态码（400/401/404/502），`_dispatch_api` 异常统一走 `run_safely` + `redact_secrets`。

测试：

- 新增全部 16 个写工具的 confirm 负向回归测试（缺省与 `confirm=false` 都必须返回 `validation_error`）。
- 新增认证限速 429、非 email account 404、死路由移除、MSAL http_client 超时、device flow 并发共享单次轮询、凭证库 close 幂等、草稿 pending cleanup 等用例。

## 0.1.46 - 2026-07-16

- `mail_forward` 支持 `folder` 与 `html` 参数；`html=true` 时转发 comment 自动转成纯文本。

## 0.1.45 - 2026-07-16

- Mail WebUI 读取 API 增加按 HTTP credential 隔离的进程内缓存：文件夹和附件信息 `5 minutes`，列表和搜索 `30 seconds`，正文 `10 minutes`；写操作立即失效对应缓存。
- Graph 邮件正文读取修正：优先使用 body content，回退 `bodyPreview`。

## 0.1.44 - 2026-07-16

- Mail WebUI 样式重构（`mail.css` 大幅调整）。

## 0.1.43 - 2026-07-16

- Mail WebUI 移动端布局：文件夹抽屉、返回按钮、写信弹层；TinyMCE 工具栏自动换行。

## 0.1.42 - 2026-07-15

- Mail WebUI 写信、回复、转发支持 `to` / `cc` / `bcc` 和 HTML 正文。
- 新增 Mail UI 状态测试。

## 0.1.41 - 2026-07-15

- 新增日历 UI 状态场景测试（基于 Node.js 校验 `calendar.js` 行为，环境缺少 Node.js 时跳过）。

## 0.1.40 - 2026-07-15

- Calendar WebUI 支持服务端草稿：草稿存 SQLite、按 credential 隔离、可跨设备恢复，版本号检测多设备写入冲突，支持按 credential 配置保留期。

## 0.1.39 - 2026-07-15

- 新增基于 Node.js 的日历时区/UI 测试（`tests/test_calendar_timezone.py`）。

## 0.1.38 - 2026-07-14

- “接下来”日程列表改为 grid 布局。

## 0.1.37 - 2026-07-14

- 修复 FullCalendar 表格间距样式。

## 0.1.36 - 2026-07-14

- 日历日程自适应布局：关闭 slot 内日程重叠，启用 `dayMaxEvents` 折叠。

## 0.1.35 - 2026-07-14

- 日期时间输入改用原生 picker；结束时间非法时自动按半小时取整纠正。

## 0.1.34 - 2026-07-14

- 日历增删改结果同步投影到所有已缓存时间范围，避免切换月份后数据不一致。

## 0.1.33 - 2026-07-14

- 取消 API 和 upload 的请求 `Origin` 校验，访问控制只由 Bearer token 和 credential 状态决定。

## 0.1.32 - 2026-07-14

- 新增 `OUTLOOK_MCP_SECRET_MIN_LENGTH` 环境变量，配置 HTTP Bearer token 最小长度（默认 `4 characters`）。

## 0.1.31 - 2026-07-13

- 构建统一使用清华 Python 包源，uv 默认 index 收拢到 `pyproject.toml` 的 `[[tool.uv.index]]`。

## 0.1.30 - 2026-07-13

- 大范围重构与稳定性修正：`tools/`、`uploads`、`http_server`、`credential_store`、`graph_auth` 等模块全量梳理，测试同步更新（38 个文件，约 4570 行新增）。
- README 补充 HTTP credential 显式注册与授权 URL 打印流程。

## 0.1.29 - 2026-07-13

- 周视图时间槽背景样式修正。

## 0.1.28 - 2026-07-10

- 日历键盘导航与 aria 无障碍标注：月视图、当天列表、描述编辑器均可键盘操作。

## 0.1.27 - 2026-07-10

- 周视图当前时间指示线样式与定时同步。
- 同步 `uv.lock` 锁文件版本。

## 0.1.26 - 2026-07-10

- 仅同步版本号与 Docker 镜像 tag，无功能变更。

## 0.1.25 - 2026-07-10

- 新建日程默认时间改为当前小时整点、时长 `30 minutes`。

## 0.1.24 - 2026-07-10

- 日历新增当天日程列表面板（day agenda），支持点击日期查看与新建。

## 0.1.23 - 2026-07-10

- `auth_setup` 授权后校验 Graph token 属于配置邮箱，不匹配时报错；支持 silent token 获取。
- 新增 `graph_auth` 模块做 MSAL 账号身份匹配。

## 0.1.22 - 2026-07-10

- WebUI 页面增加 SVG favicon 和 `theme-color`。

## 0.1.21 - 2026-07-09

- 授权状态校验容忍 MSAL cache 缺少 username 的账号记录。
- Calendar WebUI 的 API 请求增加 `60 seconds` 超时（AbortController）。

## 0.1.20 - 2026-07-09

- Graph 会议邀请返回 `ics_uid`，`id` 为 `null`；`calendar_update` 部分更新时保留现有 attendees。
- MSAL 账号查找按配置的 username 过滤。
- 全天日程时间行布局调整。

## 0.1.19 - 2026-07-09

- 编辑日程时检测远端变更并提示“覆盖为远端”；日程详情增加本地缓存。

## 0.1.18 - 2026-07-09

- 日程详情增加“正在读取日程详情...”加载态；面板头部操作按钮调整。

## 0.1.17 - 2026-07-09

- 日历正文支持完整 HTML document：提取 `<body>` 内容，剥离 `head` / `script` / `style` / 注释。

## 0.1.16 - 2026-07-09

- 日程描述改回纯文本输入，移除富文本预览；没有变更时保存给出提示。

## 0.1.15 - 2026-07-09

- 引入本地 vendored TinyMCE 富文本编辑器；`/static/` 静态资源改由 `StaticFiles` 托管。

## 0.1.14 - 2026-07-09

- 编辑面板标题简化为“新建日程 / 编辑日程”。

## 0.1.13 - 2026-07-09

- 纯文本转换支持已转义的 HTML（如 `&lt;p&gt;` 自动 unescape）。
- 日历表单提示参会人数量及“Graph 会自动发送邀请”。

## 0.1.12 - 2026-07-09

- 日历编辑面板双栏布局调整。

## 0.1.11 - 2026-07-09

- 日历支持拖拽选区直接新建日程草稿；修正日程时间显示。

## 0.1.10 - 2026-07-09

- 日历改为 cache-first 加载，移除后台自动刷新，提示“点击刷新获取最新日程”。

## 0.1.9 - 2026-07-09

- Dockerfile 简化启动：通过 `VIRTUAL_ENV` / `PATH` 直接运行 `outlook-mcp-http`，不再 `uv run`；compose 移除 `version` 字段。

## 0.1.8 - 2026-07-09

- Graph 后端增加时区别名表（`GRAPH_TIMEZONE_ALIASES`）。
- 日历表单支持参会人（attendees）输入。

## 0.1.7 - 2026-07-09

- Calendar WebUI 中文化：加载本地 FullCalendar 中文 locale，页面文案中文化。

## 0.1.6 - 2026-07-09

- 统一项目版本号到 `0.1.6`，同步 Docker demo / 安装文档里的镜像 tag。
- 改进 HTTP WebUI 日历编辑器，使用本地静态资源并收敛为 Graph-only 可视化编辑路径。

## 0.1.1 - 2026-06-16

- 小版本发布，更新 Docker demo 镜像 tag。
- Docker 镜像 tag：`outlook-mcp:0.1.1`。

## 0.1.0 - 2026-06-16

- `calendar_create` 默认关闭 Outlook 提醒，可用 `reminder_minutes_before_start` 显式开启。
- Docker 镜像 tag：`outlook-mcp:0.1.0`。

## 0.0.6 - 2026-06-15

- 优化 endpoint 页面安装区，改为带复制按钮的 copy panel。
- 使用本地 vendored Pico CSS 和项目样式文件，不依赖 CDN。
- Docker 镜像 tag：`outlook-mcp:0.0.6`。

## 0.0.5 - 2026-06-15

- 修复 HTTPS 页面中复制用 MCP JSON 仍显示 `http://...` 的问题。
- 页面安装提示区改为浏览器端按当前 `window.location.origin` 重写 endpoint URL。
- Docker 镜像 tag：`outlook-mcp:0.0.5`。
