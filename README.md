# Outlook MCP

这是一个 Python MCP server，用于让 AI Agent 通过高层工具操作 Outlook 邮箱、日历、联系人和任务。

Agent 只调用 MCP 暴露的业务操作，例如 `mail_send`、`calendar_create`。底层统一走 Microsoft Graph；普通 event 还是 MIME/ICS 会议邀请，由本项目脚本根据参数自动处理，不需要 Agent 知道底层细节。

## 快速开始

如果只是想先跑起来，推荐走 Docker Compose：

> **前提**：`docker-compose.demo.yml` 面向已有 Traefik 反向代理的部署，声明了 external Docker network `traefik`（下面命令块的第一步会创建它），且默认使用本地构建的镜像 `outlook-mcp:0.2.0`（构建命令见 [直接用 Dockerfile 构建镜像](#直接用-dockerfile-构建镜像)）。没有 Traefik 环境时，新人主路径是本地构建镜像后用 `docker run` 端口映射直接启动，见 [直接用 Dockerfile 构建镜像](#直接用-dockerfile-构建镜像)；或者去掉 compose 文件中的 `networks` / `labels` 配置并加上 `ports: ["8000:8000"]`。镜像 tag 不写死，以 `docker-compose.demo.yml` 中的 `image:` 为准。

```bash
docker network create traefik
mkdir -p /volume1/Docker/outlookmcp/data
chown 1000:1000 /volume1/Docker/outlookmcp/data
cp outlook-mcp.example.json /volume1/Docker/outlookmcp/outlook-mcp.json
# 启动前先编辑 /volume1/Docker/outlookmcp/outlook-mcp.json，把 email 改成自己的 Outlook.com 邮箱
export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_CSRF_KEY="$(openssl rand -hex 32)"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"
docker compose -f docker-compose.demo.yml up -d
```

> **权限提示**：容器以 UID 1000 运行，`/data` 挂载目录及其中的 SQLite 文件必须对 UID 1000 可读写。如果启动时报 `attempt to write a readonly database`，说明已有文件权限不对，执行：
>
> ```bash
> chown -R 1000:1000 /volume1/Docker/outlookmcp/data
> chmod 666 /volume1/Docker/outlookmcp/data/*.sqlite3*
> ```
>
> 如果宿主用户 UID 不是 1000，也可以在 `docker-compose.demo.yml` 中通过 `user:` 指定匹配的 UID:GID。

`tenant` 默认 `consumers` 适合个人 Microsoft 账户。上面的命令块已在 `up -d` 前提示编辑配置，不要带着默认占位邮箱启动容器。

HTTP server 启动时会自动 provisioning 单个 credential：account 取 `outlook-mcp.json` 里的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`；未设置 token 时会自动生成随机 token 并打印到启动日志。浏览器 GET 本身不会创建 credential。容器启动后，用下面的命令打印授权 URL：

```bash
printf 'https://outlook-mcp.example.com/graph/%s/%s\n' \
  "$OUTLOOK_MCP_ACCOUNT" "$OUTLOOK_MCP_BEARER_TOKEN"
```

首次打开该 URL 会显示 Microsoft device code 授权信息。授权完成后，MCP 客户端使用 `https://outlook-mcp.example.com/graph/you@outlook.com`，并配置同一个 `Authorization: Bearer ...` token。token 只允许字母、数字、`.`、`_`、`~`、`-`，不限制长度；公网部署仍应使用高熵 token。操作日志 URL 是授权 URL 后追加 `/log`。

需要换 token 时使用运维恢复命令 `uv run python -m outlook_mcp.credential_store rotate`（Docker 下加 `--db /data/outlook-mcp.sqlite3`），它会原地更新 credential，保留该 credential 的 OAuth、日志和上传隔离状态；换 token 后注意旧 token 不要再出现在客户端配置里。

本地开发可以直接用：

```powershell
uv sync
uv run pytest
uv run outlook-mcp
```

## 文档导航

- [功能](#功能)
- [实现结构](#实现结构)
- [可用 MCP 工具](#可用-mcp-工具)
- [本地安装](#本地安装)
- [MCP 客户端配置](#mcp-客户端配置)
- [HTTP MCP 模式](#http-mcp-模式)
- [Docker demo](#docker-demo)
- [Graph 个人 Outlook.com 配置](#graph-个人-outlookcom-配置)
- [故障排查](#故障排查)
- [AI Agent 使用约定](#ai-agent-使用约定)

## 功能

- 邮件：列表、搜索、读取、发送、回复、转发、移动、删除、标记已读/未读、附件列表、文件夹列表
- Mail WebUI：文件夹、分页列表、主题搜索、HTML 详情、写信、回复、转发、移动、删除和已读状态
- 日历：列表、读取、搜索、创建、更新、删除
- 日历 WebUI：跨设备草稿、草稿列表、月/周/日 ghost 占位和键盘快捷键
- 联系人：列表、搜索、读取、创建、更新、删除
- 任务：列表、读取、搜索、创建、更新、删除
- 账号信息：读取当前账号邮箱和显示名
- 账号绑定：Microsoft Graph（个人 Outlook.com 账号）
- MCP 传输：`stdio`、Streamable HTTP

所有写操作都需要 `confirm=true`，避免 Agent 误发邮件或误改数据。

## 实现结构

这部分主要给维护者阅读；如果你只是安装或接入 MCP 客户端，可以先看 [快速开始](#快速开始)、[Docker demo](#docker-demo) 和 [Graph 个人 Outlook.com 配置](#graph-个人-outlookcom-配置)。

```text
AI Agent
  |
  | 调用 MCP 高层工具
  v
Outlook MCP
  |
  +-- mail_send
  |     |
  |     +-- Microsoft Graph /me/sendMail
  |         通过 Graph 发普通邮件
  |
  +-- calendar_create
        |
        +-- 没有邀请人
        |     |
        |     +-- Microsoft Graph /me/events
        |         创建普通日程
        |
        +-- 有邀请人
              |
              +-- Microsoft Graph /me/sendMail
                      Content = MIME 邮件 + ICS 日历邀请
                      通过邮件形式发送会议邀请
```

简化理解：

```text
Graph 账号（个人 Outlook.com）
  |
  +-- 发邮件：Graph sendMail
  +-- 普通日程：Graph /me/events
  +-- 会议邀请：Graph sendMail + MIME/ICS
```

## 可用 MCP 工具

邮件：

- `mail_list`
- `mail_search`
- `mail_get`
- `mail_send`
- `mail_reply`
- `mail_forward`
- `mail_move`
- `mail_delete`
- `mail_mark_read`
- `mail_mark_unread`
- `mail_attachments`
- `mail_folders`

日历：

- `calendar_list`
- `calendar_get`
- `calendar_search`
- `calendar_create`
- `calendar_update`
- `calendar_delete`

`calendar_list`、`calendar_get`、`calendar_search`、`calendar_create`、`calendar_update` 支持可选 `timezone` 参数，值为 IANA 时区名，例如 `Asia/Shanghai`。定时日程的 `start` / `end` 不带 `Z` 或 `+08:00` 这类 offset 时，会按 `timezone` 指定的本地墙上时间解释；未传 `timezone` 时默认使用 `Asia/Shanghai`。定时日程返回的 `start` / `end` 是 UTC ISO 8601，`start_local` / `end_local` 是本地显示时间。全天日程则返回 `YYYY-MM-DD` 日期，`end` 是 exclusive end date，不做时区偏移。

`calendar_create` 默认不设置 Outlook 提醒；如需提醒，传 `reminder_minutes_before_start=10` 表示开始前 10 分钟提醒。

Graph 普通日程已创建、但后续 attachment 写入失败时，会返回 `ok=true`、`status="partial_success"`、`event_id` 和 `attachment_errors`。此时不要重试整个 `calendar_create`，否则会创建重复日程。Graph 写操作在 network timeout 或 HTTP `5xx` 后无法确认结果时返回 `error="outcome_unknown"`，也不会自动重试；应先通过 list/search/get 核对远端状态。

联系人：

- `contacts_list`
- `contacts_search`
- `contact_get`
- `contact_create`
- `contact_update`
- `contact_delete`

任务：

- `tasks_list`
- `task_get`
- `task_search`
- `task_create`
- `task_update`
- `task_delete`

账号：

- `account_info`

## 本地安装

安装依赖：

```powershell
uv sync
```

运行测试：

```powershell
uv run pytest
```

启动 MCP server：

```powershell
uv run outlook-mcp
```

启动 HTTP MCP server：

```powershell
uv run outlook-mcp-http
```

调试时也可以用根目录启动器。启动前必须显式设置 `OUTLOOK_MCP_ACCOUNT` 和 `OUTLOOK_MCP_BEARER_TOKEN`。启动器会打印完整 MCP JSON，等你复制后按回车，再在当前窗口启动 HTTP 服务；server 启动时会用 `OUTLOOK_MCP_BEARER_TOKEN` 自动 provisioning 对应的 credential：

```powershell
$env:OUTLOOK_MCP_ACCOUNT = "you@outlook.com"
$env:OUTLOOK_MCP_BEARER_TOKEN = [guid]::NewGuid().ToString("N")
.\start_outlook_mcp_http.bat
.\start_outlook_mcp_http.bat 8000 127.0.0.1
```

Linux 下使用同名 shell 启动器：

```bash
chmod +x ./start_outlook_mcp_http.sh
./start_outlook_mcp_http.sh
./start_outlook_mcp_http.sh 8000 127.0.0.1
```

启动器默认读取根目录的 `outlook-mcp.json`，默认 SQLite 数据库是 `~/.outlook-mcp/outlook-mcp.sqlite3`。常用环境变量可以直接覆盖：

```bash
export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"
export OUTLOOK_MCP_MODE="calendar"
./start_outlook_mcp_http.sh 8000 0.0.0.0
```

Python 依赖源已统一配置在 `pyproject.toml` 的 `[[tool.uv.index]]`，裸 `uv lock`、`uv sync` 和 `uv run` 都使用清华源，无需额外环境变量。

## MCP 客户端配置

本项目已提供 `.mcp.json`。如果 MCP 客户端支持项目级配置，直接打开本项目即可加载。注意 `.mcp.json` 里的 `cwd` 字段和相对路径需要客户端支持项目级配置解析；不支持 `cwd` 的客户端请把配置复制到客户端自己的 MCP 配置文件，并把 `cwd` 与 `OUTLOOK_CONFIG_FILE` 改成实际 checkout 的绝对路径。

本机使用优先选 `stdio`，也就是下面的 `outlook` server；它不经过 HTTP/TLS，最少出网络和证书问题。只有需要局域网、Docker、Traefik 或远程访问时，才使用 HTTP MCP endpoint。

当前配置示例：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["run", "outlook-mcp"],
      "cwd": "/absolute/path/to/OutlookMCP",
      "env": {
        "OUTLOOK_CONFIG_FILE": "/absolute/path/to/OutlookMCP/outlook-mcp.json",
        "PYTHONUTF8": "1",
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

Windows 路径需要使用转义后的反斜杠，例如 `C:\\path\\to\\OutlookMCP`（以你的实际 checkout 路径为准）。

如果你的 MCP 客户端不读取项目级 `.mcp.json`，把上面的 `outlook` server 配置复制到客户端自己的 MCP 配置文件中。

HTTP MCP 的客户端配置由 WebUI 自动生成。启动 `uv run outlook-mcp-http` 或 Docker 服务时会自动 provisioning credential（account 取配置文件的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`，未设置时自动生成并打印到启动日志），然后用浏览器打开 `/{method}/{account}/{token}` 完成授权。未知 token 的 GET 只返回 `credential_not_registered`，不会写 SQLite。copy panel 给 MCP 客户端的仍是 `/{method}/{account}` + `Authorization: Bearer ...`。

## HTTP MCP 模式

HTTP 模式用于内部服务或 Traefik 反代 HTTPS。启动：

```powershell
$env:OUTLOOK_CONFIG_FILE = "C:\path\to\OutlookMCP\outlook-mcp.json"  # 以实际 checkout 路径为准
$env:OUTLOOK_MCP_HTTP_HOST = "127.0.0.1"
$env:OUTLOOK_MCP_HTTP_PORT = "8000"
uv run outlook-mcp-http
```

建议显式设置固定的 `OUTLOOK_MCP_CSRF_KEY`（例如 `openssl rand -hex 32`）。不设置时每次重启都会随机生成，重启后 WebUI 已打开页面里的授权/取消授权表单会因 CSRF token 失效而全部校验失败。

本机调试也可以直接运行：

```powershell
.\start_outlook_mcp_http.bat 8000 127.0.0.1
```

MCP endpoint 形式固定为：

```text
/{method}/{account}
Authorization: Bearer {token}
```

示例：

```json
{
  "mcpServers": {
    "outlook-http": {
      "type": "streamable-http",
      "url": "http://127.0.0.1:8000/graph/you@outlook.com",
      "headers": {
        "Authorization": "Bearer replace-with-random-token-32-chars",
        "X-Outlook-MCP-Mode": "calendar"
      }
    }
  }
}
```

如果前面有 Traefik、nginx、Caddy 这类反向代理终止 HTTPS，后端仍然监听 HTTP。浏览器页面里的 MCP JSON 会在渲染时使用当前 `window.location.origin`，所以通过 HTTPS 打开的页面会显示 `https://...` 复制 URL；页面样式使用本地 vendored Pico CSS，不依赖 CDN；`OUTLOOK_MCP_TRUST_PROXY_HEADERS=1` 仍用于后端需要信任反代 headers 的场景。

反代必须把完整证书链配置好。MCP 客户端报 `streamableHttp connect failed: fetch failed`、`sse connect failed` 或 `unable to verify the first certificate` 时，先检查外部 HTTPS 证书链；不要通过暴露新的 MCP 底层工具绕过这个问题。

首次使用时：

- HTTP server 启动时会自动 provisioning 单个 credential：account 取 `outlook-mcp.json` 的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`。token 只允许字母、数字、`.`、`_`、`~`、`-`，不限制长度；公网部署仍应使用高熵 token。未设置 `OUTLOOK_MCP_BEARER_TOKEN` 时会生成随机 token，并在启动日志打印形如 `http://{host}:{port}/graph/{email}/{token}` 的 setup URL。
- MCP 客户端调用 `/{method}/{account}`，必须带 `Authorization: Bearer {token}`；token 不匹配时返回 `auth_required`。
- 用户手工打开 `/{method}/{account}/{token}`，页面才会展示 Microsoft device code 授权信息。
- 用户完成 Microsoft 登录后刷新页面；授权成功后，MCP 客户端继续使用无 token URL + Bearer header。

`credential_store` 的 register/rotate CLI 只是运维恢复工具，正常启动不再需要手工 register。轮换 token 使用 `uv run python -m outlook_mcp.credential_store rotate graph you@outlook.com old-token-1234 new-token-5678`。旧 token 立即失效，credential ID 与已绑定 OAuth 保持不变。

WebUI 默认选中 `Full`，不会写 `X-Outlook-MCP-Mode` 或 `X-Outlook-MCP-Tools`，因此暴露全部高层工具。切到 `Calendar`、`Mail` 等 mode 时，页面会自动写入对应的 `X-Outlook-MCP-Mode`；只有 `Custom` 会展示工具输入框并写入 `X-Outlook-MCP-Tools`。从 `Full` 切到 `Custom` 会默认填入全部工具，从单一 mode 切到 `Custom` 会默认填入该 mode 的工具，便于删减或微调。

操作日志页面：

```text
/{method}/{account}/{token}/log
```

日志页展示同一 endpoint 的 HTTP MCP 调用摘要，包括工具名、状态、错误类型和耗时。页面支持 `range=all|24h|7d|30d|custom`，自定义区间使用 `from=YYYY-MM-DD&to=YYYY-MM-DD`。日志不会保存 raw secret、token、完整请求参数或完整响应体。

WebUI JSON API 只接受无 secret 路径的 JSON `POST`：

```text
POST /{method}/{account}/api/calendar/list
Authorization: Bearer {token}
Content-Type: application/json
```

完成 Graph 授权后，可以打开 Mail WebUI 和 Calendar WebUI：

```text
/{method}/{account}/{token}/mail
/{method}/{account}/{token}/calendar
```

Mail WebUI 通过无 secret 的 `/api/mail/...` endpoint 加 Bearer header 调用现有高层 Mail 工具。页面支持顶层文件夹、每页 `20 items` 的邮件列表、当前文件夹内主题搜索、邮件详情、写信、HTML 回复/转发、移动、移入已删除邮件和显式标记已读/未读。HTML 邮件会先清理危险元素和属性，再放入 sandbox iframe；远程图片默认阻断，用户可以显式加载，`cid:` 内嵌图片首版显示占位。

Mail WebUI 的读取 API 使用按 HTTP `credential_id` 隔离的进程内短 TTL 缓存：文件夹和附件信息缓存 `5 minutes`，列表和搜索缓存 `30 seconds`，正文缓存 `10 minutes`。缓存总量最多 `32 MiB`，单条最多 `512 KiB`；发送、回复、转发、移动、删除、标记已读/未读以及清除授权都会立即失效该 credential 的邮件缓存。缓存不会跨进程或容器重启保留，也不会缓存附件二进制内容。

Mail WebUI 首版仅展示附件名称、类型和大小，不提供附件下载、普通文件上传或服务端邮件草稿。现有 HTTP 图片上传仍只用于 MCP/API 的 PNG、JPEG、GIF、WebP attachment 输入，不会出现在 Mail WebUI 编辑器里。

日历 WebUI 支持服务端草稿。草稿保存在 `OUTLOOK_MCP_DB` 指向的 SQLite 文件中，
不会提前创建 Graph event，也不会发送会议邀请。草稿按 HTTP `credential_id` 隔离：
使用同一个 endpoint credential 的设备可以恢复草稿；`credential_store rotate` 不改变
`credential_id`，重新 `register` 则会创建新的隔离范围。

键盘操作：

- `Ctrl+S` / `Cmd+S`：把当前新建或编辑表单暂存为跨设备草稿。
- `Ctrl+Enter` / `Cmd+Enter`：通过现有 `calendar_create` / `calendar_update` 流程最终保存。
- 日程、当天日程列表或草稿列表中的条目获得键盘焦点后，按 `E` 进入编辑。

“草稿”Tab 会列出当前 credential 的全部草稿。具有完整合法时间范围的草稿还会在
月、周、日和列表视图中显示特殊颜色的 `[草稿]` ghost 占位；编辑草稿不会隐藏或删除
真实 Outlook 日程。每个真实日程最多有一个编辑草稿，多设备同时写入使用版本号检测冲突。
草稿默认永久保存，也可以在草稿 Tab 中为当前 credential 设置 `1 day` 到 `3650 days`
的保留期限。提交返回 `outcome_unknown` 时草稿会保留并锁定；在 Outlook 中确认操作未生效后，
可以通过“已核对未提交”显式解锁，系统不会自动重试。

`/{method}/{account}/{token}/api/...` 已禁用。API 和 upload 不校验请求 `Origin`，但仍必须使用有效 Bearer token；同步 Graph 调用在线程池执行，不会阻塞 ASGI event loop。取消服务端 Origin 校验不等于启用了浏览器跨站 CORS/preflight。

HTTP 图片上传 endpoint：

```text
POST /{method}/{account}/uploads
Authorization: Bearer {token}
multipart/form-data field: file
```

上传前必须满足两个条件：`method + account + token` 已经注册为 HTTP credential，且该 credential 曾经完成过 Microsoft OAuth 授权。只有 pending device flow 或仅打开过 endpoint 页面不算完成授权，接口会返回 `403 auth_required`。

上传成功后返回 `relative_path`、`width`、`height`，例如 `uploads/12/20260522/6f1a2b3c4d5e6789a0b1c2d3e4f56789/photo.png`。`mail_send.attachments` 和 `calendar_create.attachments` 可以直接引用这个相对路径：

```json
{"relative_path": "uploads/12/20260522/6f1a2b3c4d5e6789a0b1c2d3e4f56789/photo.png", "disposition": "attachment"}
```

HTML 正文内嵌图片使用 `inline + content_id`：

```json
{"relative_path": "uploads/12/20260522/6f1a2b3c4d5e6789a0b1c2d3e4f56789/photo.png", "disposition": "inline", "content_id": "photo1"}
```

正文里对应写 `<img src="cid:photo1">`，并设置 `html=true`。如果这个 `<img>` 没有 `width`、`height` 或 `style` 尺寸约束，服务端会按图片实际像素自动补默认显示尺寸；大图最大显示宽度为 `600px`，小图不会被放大。邮件和日程都支持普通附件与 inline 图片；带邀请人的 Graph 日程会通过 MIME/ICS 会议邀请发送附件，没有邀请人的普通 Graph 日程会创建 event 后再写入 event attachments。上传目录默认是 `~/.outlook-mcp/uploads`，可用 `OUTLOOK_MCP_UPLOAD_DIR` 覆盖；默认只允许 PNG、JPEG、GIF、WebP，单文件上限 `3 MiB`（`OUTLOOK_MCP_UPLOAD_MAX_BYTES`），TTL `24 hours`（`OUTLOOK_MCP_UPLOAD_TTL_SECONDS`）。每个 credential 还有总量配额：默认 `30 MiB`（`OUTLOOK_MCP_UPLOAD_QUOTA_BYTES`）和 `100` 个文件（`OUTLOOK_MCP_UPLOAD_QUOTA_FILES`），超限后新上传会被拒绝。

HTTP 模式默认 SQLite 文件：

```text
~/.outlook-mcp/outlook-mcp.sqlite3
```

可以用 `OUTLOOK_MCP_DB` 改位置。这个 SQLite 是本地状态库：HTTP credential、Graph 的 MSAL token cache、pending device flow、日历草稿、操作日志和上传元数据都存在这里；`outlook-mcp.json` 里的 `graph.client_id`、`tenant`、`scopes` 仍然作为 Graph 基础配置使用。

## Docker demo

本项目的 Docker 镜像默认启动 HTTP MCP server。`Dockerfile` 使用 `docker.m.daocloud.io/library/python:3.12-slim` 作为基础镜像；安装 `uv` 的 bootstrap pip 源和 `pyproject.toml` 中的 uv 默认源均使用清华源，适合国内网络环境。

使用 Docker 前先准备配置文件：

```bash
mkdir -p /volume1/Docker/outlookmcp/data
cp outlook-mcp.example.json /volume1/Docker/outlookmcp/outlook-mcp.json
export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_CSRF_KEY="$(openssl rand -hex 32)"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"
```

然后编辑 `/volume1/Docker/outlookmcp/outlook-mcp.json`，把 `email` 改成自己的 Outlook.com 邮箱。

### 从 clone 到 Docker Compose 启动

`docker-compose.demo.yml` 面向 Traefik 反向代理部署，要求宿主机已有外部 Docker network `traefik`，镜像 tag 为本地构建的 `outlook-mcp:0.2.0`。

```bash
git clone https://github.com/ljq29/MCP_outlook.git
cd MCP_outlook

mkdir -p /volume1/Docker/outlookmcp/data
chown 1000:1000 /volume1/Docker/outlookmcp/data
cp outlook-mcp.example.json /volume1/Docker/outlookmcp/outlook-mcp.json
nano /volume1/Docker/outlookmcp/outlook-mcp.json

export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_CSRF_KEY="$(openssl rand -hex 32)"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"

docker compose -f docker-compose.demo.yml up -d
docker compose -f docker-compose.demo.yml ps
docker compose -f docker-compose.demo.yml logs -f
```

`OUTLOOK_MCP_CSRF_KEY` 必须是固定 secret。否则容器重启后浏览器里的 CSRF token 会失效；多 worker 部署时还会直接校验失败。

`docker-compose.demo.yml` 默认使用本地镜像 `outlook-mcp:0.2.0`。如果镜像不存在，先运行 `docker build -t outlook-mcp:0.2.0 .`，再执行 `docker compose -f docker-compose.demo.yml up -d`。

已有部署升级到当前镜像：

```bash
docker build -t outlook-mcp:0.2.0 .
docker compose -f docker-compose.demo.yml up -d --force-recreate
docker compose -f docker-compose.demo.yml logs -f outlookmcp
```

如果升级是为了处理 HTTPS 页面“不安全”提示，升级后检查页面源码不应再出现 `action="http://...`：

```bash
curl -sk "https://outlook-mcp.example.com/graph/$OUTLOOK_MCP_ACCOUNT/$OUTLOOK_MCP_BEARER_TOKEN" \
  | grep 'action="http://'
```

这个命令没有输出就是预期结果。

启动后打印浏览器授权 URL：

```bash
printf 'https://outlook-mcp.example.com/graph/%s/%s\n' \
  "$OUTLOOK_MCP_ACCOUNT" "$OUTLOOK_MCP_BEARER_TOKEN"
```

### 从 GHCR 拉取自动构建的镜像

仓库内置 GitHub Actions workflow（`.github/workflows/docker-image.yml`）：push 到 `main` 时构建并推送 `main` / `sha-*` tag，push `v*` 版本 tag（例如 `v0.2.0`）时额外构建 `0.2.0`、`0.1`、`latest` tag。镜像发布在 GHCR：

```bash
docker pull ghcr.io/ljq29/mcp_outlook:main
```

首次发布后，GHCR package 默认是 private；想公开拉取需在 GitHub → Packages → `mcp_outlook` → Package settings 中把可见性改为 Public。

### 直接用 Dockerfile 构建镜像

先构建镜像：

```bash
docker build -t outlook-mcp:0.2.0 .
```

然后直接启动容器：

```bash
docker run -d \
  --name outlook-mcp \
  -p 8000:8000 \
  -e OUTLOOK_GRAPH_CLIENT_ID="$OUTLOOK_GRAPH_CLIENT_ID" \
  -e OUTLOOK_MCP_CSRF_KEY="$OUTLOOK_MCP_CSRF_KEY" \
  -e OUTLOOK_CONFIG_FILE=/config/outlook-mcp.json \
  -e OUTLOOK_MCP_DB=/data/outlook-mcp.sqlite3 \
  -e OUTLOOK_MCP_BEARER_TOKEN="$OUTLOOK_MCP_BEARER_TOKEN" \
  -v "$(pwd)/outlook-mcp.json:/config/outlook-mcp.json:ro" \
  -v outlookmcp-data:/data \
  outlook-mcp:0.2.0
```

启动后访问对应的 `/{method}/{account}/{token}` 授权 URL。未知 token 不会因浏览器 GET 自动写入数据库：

```bash
printf 'http://127.0.0.1:8000/graph/%s/%s\n' \
  "$OUTLOOK_MCP_ACCOUNT" "$OUTLOOK_MCP_BEARER_TOKEN"
```

MCP 客户端不要使用上面的浏览器授权 URL；客户端配置应使用 `http://127.0.0.1:8000/graph/you@outlook.com`，并在 `headers.Authorization` 中写同一个 Bearer token。

### 构建并推送到自己的 Registry

```bash
# 构建镜像
docker build -t <your-registry>/outlook-mcp:0.2.0 .

# 推送到自己的 registry，例如 ghcr.io、Docker Hub 或自建 registry
docker push <your-registry>/outlook-mcp:0.2.0
```

发布新版本时，更新版本号并重复上述步骤。当前最新版本：`0.2.0`。

### 用 Docker Compose 直接创建容器

Docker demo 默认启动 HTTP MCP server：

```bash
docker compose -f docker-compose.demo.yml up -d
```

如果只想先创建容器而不立刻启动，可以用：

```bash
docker compose -f docker-compose.demo.yml create
docker compose -f docker-compose.demo.yml start
```

Compose 会把配置文件和数据库持久化：

- `/volume1/Docker/outlookmcp/outlook-mcp.json` 挂载到容器内 `/config/outlook-mcp.json:rw`
- `/volume1/Docker/outlookmcp/data` 挂载到 `/data`（目录及文件必须对 UID 1000 可读写）
- SQLite 文件固定为 `/data/outlook-mcp.sqlite3`
- 服务加入外部 Docker network `traefik`，默认域名是 `outlook-mcp.example.com`
- 页面提示的 MCP URL 会在浏览器端按当前 `window.location.origin` 重写；`OUTLOOK_MCP_TRUST_PROXY_HEADERS=1` 会信任 Traefik 注入的 `Forwarded` / `X-Forwarded-*`

Docker 模式下 server 启动时会自动 provisioning credential：account 取 `/config/outlook-mcp.json` 的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`，未设置时生成随机 token 并打印到容器日志。随后打开 `/graph/{account}/{token}` 获取 Microsoft device code；MCP 客户端后续使用 `/graph/{account}` + `Authorization: Bearer {token}`。credential 存储在持久化的 `/data/outlook-mcp.sqlite3` 中，升级或重启不会清除。

## Graph 个人 Outlook.com 配置

普通 `outlook.com` 个人账号推荐使用 Graph。

复制示例配置：

```powershell
Copy-Item outlook-mcp.example.json outlook-mcp.json
```

编辑 `outlook-mcp.json`：

```json
{
  "provider": "graph",
  "email": "yourname@outlook.com",
  "graph": {
    "client_id": "9966331f-d9e2-42b2-85bf-5a54280ee60e",
    "tenant": "consumers",
    "token_cache_path": "~/.outlook-mcp/msal-token-cache.json",
    "scopes": [
      "User.Read",
      "Mail.ReadWrite",
      "Mail.Send",
      "Calendars.ReadWrite",
      "Contacts.ReadWrite",
      "Tasks.ReadWrite"
    ]
  }
}
```

Azure / Microsoft Entra App Registration 要求：

- 应用类型使用 public client
- Authentication 中启用 `Allow public client flows`
- 个人 Microsoft 账户使用 `tenant: "consumers"`
- 不需要 `client_secret`

`OUTLOOK_GRAPH_CLIENT_ID` 的优先级高于 `outlook-mcp.json` 里的 `graph.client_id`。也就是说，容器或系统环境变量里设置了 `OUTLOOK_GRAPH_CLIENT_ID` 时会直接使用它；没有设置时才回退到配置文件。

首次授权可以直接运行：

```powershell
authorize_outlook.bat
```

Linux 下运行等价的 Python 入口（`authorize_outlook.bat` 只是它的包装）：

```bash
mkdir -p logs
uv run python -m outlook_mcp.auth_setup --debug \
  --config-file ./outlook-mcp.json \
  --log-file "./logs/outlook-auth-$(date +%Y%m%d-%H%M%S).log"
```

授权日志会写入 `logs/`，token 会写入 `token_cache_path`。`outlook-mcp.json` 和 token cache 都不要提交到 Git。

## 故障排查

`docker build` 拉取基础镜像失败：当前 `Dockerfile` 默认使用 `docker.m.daocloud.io/library/python:3.12-slim`。如果仍失败，先确认 Docker Desktop 能访问该镜像源，再执行 `docker pull docker.m.daocloud.io/library/python:3.12-slim`；需要用官方源覆盖时执行 `docker build --build-arg BASE_IMAGE=python:3.12-slim .`。

`uv sync` 或 Docker 构建下载 Python 包很慢：先检查 `pyproject.toml` 的 `[[tool.uv.index]]` 与 `uv.lock` 是否仍指向清华源。Dockerfile 的 `PIP_INDEX_URL` 只负责安装 `uv` 自身，项目依赖源由 `pyproject.toml` 统一管理。

启动时报 `graph.client_id is required`：设置 `OUTLOOK_GRAPH_CLIENT_ID`，或者在 `outlook-mcp.json` 的 `graph.client_id` 中填写 Microsoft Entra / Azure App Registration 的 Application client ID。

HTTP endpoint 返回 `auth_required`：这是首次授权的正常状态。用浏览器打开同一个 endpoint，按页面提示完成 Microsoft device code 登录，然后让 MCP 客户端重试。

MCP 客户端报 `streamableHttp connect failed: fetch failed; sse connect failed: ... unable to verify the first certificate`：这是客户端校验 HTTPS 证书链失败。公网/反代部署应修服务端 `fullchain.pem` 或补齐 intermediate cert；公司代理或自签 CA 场景，把 root CA 导出为 PEM，并在启动 MCP 客户端的环境设置 `NODE_EXTRA_CA_CERTS=/etc/ssl/certs/company-root-ca.pem`。本机调试建议改用 `stdio` 配置，不要长期使用 `NODE_TLS_REJECT_UNAUTHORIZED=0`。

浏览器提示 HTTPS 页面“某些部分不安全”：检查页面源码里是否还有 `action="http://...`。授权按钮应使用同源 path，例如 `action="/graph/you@outlook.com/replace-with-random-token-32-chars/auth/reauthorize"`；如果仍是 `http://...`，说明容器镜像或代码还没更新。

日志页面没有数据：日志只记录同一个 `method + account + token` credential 的 HTTP MCP 调用摘要。先让 MCP 客户端调用一次 `/{method}/{account}` 且带 `Authorization: Bearer {token}`，再打开 `/{method}/{account}/{token}/log`。

端口 `8000` 被占用：本地启动时设置 `OUTLOOK_MCP_HTTP_PORT`，或者 Docker Compose 里修改 `ports`，例如 `"18000:8000"`。

## AI Agent 使用约定

- Agent 只需要调用 MCP 工具，不要绕过 MCP 直接调用 Graph API。
- 写操作必须传 `confirm=true`。
- `calendar_create` 创建普通日程或会议都用同一个入口。
- 日历时间推荐传本地墙上时间加 `timezone`，例如 `start="2026-03-16T20:00:00"`、`timezone="Asia/Shanghai"` 表示北京时间晚上 8 点。
- 如果传入 `required_attendees` 或 `optional_attendees`，后端会自动发送会议邀请。
- Graph 个人账号下，有邀请人的会议会自动走 MIME/ICS，避免 Outlook.com 把外部地址归一化成当前账号。
- Graph 带邀请人的 `calendar_create` 不支持 `send_invites=false`；Graph `calendar_update` 也不支持 `send_invites=false`，这两种情况会返回 `validation_error`。
- Graph 带邀请人的 `calendar_create` 返回 `sent=true` 和 `ics_uid`，`id` 为 `null`；后续更新或删除要先用 `calendar_list` / `calendar_search` 获取 Graph event id。
- 如果返回 `auth_required`，把 `verification_uri` 和 `user_code` 展示给用户，等用户完成登录后重试原工具。

## 常用命令

```powershell
uv sync
uv run pytest
uv run outlook-mcp
uv run outlook-mcp-http
.\start_outlook_mcp_http.bat
./start_outlook_mcp_http.sh
docker compose -f docker-compose.demo.yml up -d
authorize_outlook.bat
```

## 重要文件

- `src/outlook_mcp/server.py`：MCP 工具入口
- `src/outlook_mcp/http_server.py`：HTTP MCP 入口
- `src/outlook_mcp/credential_store.py`：SQLite endpoint 和 Graph token 存储
- `src/outlook_mcp/graph_client.py`：Graph 适配和 MIME/ICS 会议邀请逻辑
- `src/outlook_mcp/tools/`：高层工具实现
- `src/outlook_mcp/static/mail.html`、`mail.js`、`mail.css`：Mail WebUI
- `.mcp.json`：项目级 MCP server 配置
- `Dockerfile`：HTTP MCP 容器镜像
- `docker-compose.demo.yml`：Traefik Docker demo，使用本地构建的 `outlook-mcp:0.2.0` 镜像并挂载配置和数据目录
- `start_outlook_mcp_http.sh`：Linux HTTP MCP 调试启动器
- `MCP_INSTALL.md`：给其他 AI Agent 使用的安装说明
- `outlook-mcp.example.json`：配置示例
