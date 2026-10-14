# Outlook MCP 安装与使用说明

这份说明给其他 AI agent 或开发者使用，用于把本项目作为 MCP server 挂载到支持 MCP 的客户端中。

## 1. 项目位置

项目路径：本项目 checkout 根目录（下文命令均假设在该目录下执行；历史上文档写死过 `D:\PythonWorkspace\OutlookMCP`，请以你的实际 checkout 路径为准）。

MCP server 启动命令：

```powershell
uv run outlook-mcp
```

HTTP MCP 启动命令：

```powershell
uv run outlook-mcp-http
```

HTTP MCP 调试启动器：

```powershell
.\start_outlook_mcp_http.bat
.\start_outlook_mcp_http.bat 8000 127.0.0.1
```

Linux HTTP MCP 调试启动器：

```bash
chmod +x ./start_outlook_mcp_http.sh
./start_outlook_mcp_http.sh
./start_outlook_mcp_http.sh 8000 127.0.0.1
```

仓库已提交的 `.mcp.json` 内容如下（`cwd: "."` 与相对 `OUTLOOK_CONFIG_FILE` 依赖客户端从项目根解析；不支持相对路径的客户端请展开为绝对路径）：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uv",
      "args": ["run", "outlook-mcp"],
      "cwd": ".",
      "env": {
        "OUTLOOK_CONFIG_FILE": "outlook-mcp.json",
        "PYTHONUTF8": "1",
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

如果 MCP 客户端支持项目级 `.mcp.json`，直接打开本项目即可读取。若客户端不自动读取项目配置，请把上面的 `outlook` server 配置复制到该客户端自己的 MCP 配置文件中。

本机接入优先使用上面的 `stdio` 配置。HTTP MCP 适合 Docker、Traefik、局域网或远程访问；它需要先启动 `uv run outlook-mcp-http` 或 Docker 服务，再从 endpoint 页面 copy panel 复制 `streamable-http` JSON。

## 2. 安装依赖

在项目目录执行：

```powershell
uv sync
```

项目已在 `pyproject.toml` 的 `[[tool.uv.index]]` 中统一使用清华源，直接执行 `uv lock`、`uv sync` 或 `uv run` 即可；不需要再手工设置 `UV_DEFAULT_INDEX`。

如果你要先写入个人账户授权，先执行：

```powershell
authorize_outlook.bat
```

可选验证：

```powershell
uv run pytest
```

## 3. 配置个人 Outlook.com 账户

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

`client_id` 来自 Microsoft Entra / Azure App Registration。该应用应按 public client / device code flow 使用，不需要 `client_secret`。在 Azure Portal 中打开该 App Registration，进入 **Authentication**，启用 **Allow public client flows** 并保存。个人 Microsoft 账户使用 `tenant: "consumers"`。

`OUTLOOK_GRAPH_CLIENT_ID` 的优先级高于 `outlook-mcp.json` 里的 `graph.client_id`。系统环境变量存在时直接使用它；没有设置时才回退到配置文件。这适合 Docker、systemd 或多环境部署共用同一个 App Registration。

## 4. 首次登录流程

首次调用任意 Outlook MCP 工具时，如果本地没有 token cache，工具会返回：

```json
{
  "ok": false,
  "error": "auth_required",
  "data": {
    "verification_uri": "https://microsoft.com/devicelogin",
    "user_code": "ABCD-EFGH"
  }
}
```

AI agent 应把 `verification_uri` 和 `user_code` 展示给用户，让用户在浏览器中完成 Microsoft 登录。登录完成后，重新调用同一个 MCP 工具即可。pending 授权状态会临时保存到 `token_cache_path` 旁边的 `.pending.json` 文件，所以即使 MCP 进程重启，也能继续完成换 token。成功后 token 会缓存到 `token_cache_path`，之后通常无需再次登录。

如果你不想等 MCP 工具返回 `auth_required`，可以直接先运行根目录下的 `authorize_outlook.bat`。它会把授权日志写到 `logs\outlook-auth-*.log`，并在完成后把 token 写入本地缓存。

不要向用户索要 Outlook.com 密码；个人账户登录必须走 Microsoft Graph delegated OAuth。

## 4.1 HTTP MCP 模式

HTTP 模式启动：

```powershell
$env:OUTLOOK_CONFIG_FILE = ".\outlook-mcp.json"
$env:OUTLOOK_MCP_HTTP_HOST = "127.0.0.1"
$env:OUTLOOK_MCP_HTTP_PORT = "8000"
uv run outlook-mcp-http
```

建议显式设置固定的 `OUTLOOK_MCP_CSRF_KEY`（例如 `openssl rand -hex 32`）。不设置时每次重启都会随机生成，重启后 WebUI 已打开页面里的授权/取消授权表单会因 CSRF token 失效而全部校验失败。

如果只是本机调试，直接运行：

```powershell
$env:OUTLOOK_MCP_ACCOUNT = "you@outlook.com"
$env:OUTLOOK_MCP_BEARER_TOKEN = [guid]::NewGuid().ToString("N")
.\start_outlook_mcp_http.bat 8000 127.0.0.1
```

Linux 下使用：

```bash
chmod +x ./start_outlook_mcp_http.sh
export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"
export OUTLOOK_MCP_MODE="calendar"
./start_outlook_mcp_http.sh 8000 0.0.0.0
```

启动器要求显式提供 account 与 Bearer token；server 启动时会用它们自动 provisioning 该 credential。token 只允许字母、数字、`.`、`_`、`~`、`-`，不限制长度；公网部署仍应使用高熵 token。直接运行 `uv run outlook-mcp-http` 而未设置 `OUTLOOK_MCP_BEARER_TOKEN` 时，server 会自动生成随机 token 并在启动日志打印 setup URL。

Traefik 可以在前面终止 HTTPS 并转发到本地 HTTP 服务。MCP endpoint 直接使用：

```text
/{method}/{account}
Authorization: Bearer {token}
```

反代 HTTPS 时，浏览器页面里的 MCP JSON 会在渲染时使用当前 `window.location.origin`，所以通过 HTTPS 打开的页面会显示 `https://...` 复制 URL。页面样式使用本地 vendored Pico CSS，不依赖 CDN；`OUTLOOK_MCP_TRUST_PROXY_HEADERS=1` 仍用于后端需要信任反代 headers 的场景。

如果 MCP 客户端报 `streamableHttp connect failed: fetch failed`、`sse connect failed` 或 `unable to verify the first certificate`，问题在外部 HTTPS 证书链。公网/反代部署应修 `fullchain.pem` 或补齐 intermediate cert；公司代理或自签 CA 场景，把 root CA 导出为 PEM，并在启动 MCP 客户端的环境设置 `NODE_EXTRA_CA_CERTS=/etc/ssl/certs/company-root-ca.pem`。本机调试优先改用 `stdio`。

浏览器提示 HTTPS 页面“某些部分不安全”时，检查页面源码里是否还有 `action="http://...`。授权按钮应使用同源 path，例如 `action="/graph/you@outlook.com/replace-with-random-token-32-chars/auth/reauthorize"`；如果仍是 `http://...`，说明容器镜像或代码还没更新。

示例：

```json
{
  "mcpServers": {
    "outlook-http": {
      "type": "streamable-http",
      "url": "https://mcp.example.com/graph/you@outlook.com",
      "headers": {
        "Authorization": "Bearer replace-with-random-token-32-chars",
        "X-Outlook-MCP-Mode": "calendar"
      }
    }
  }
}
```

HTTP server 启动时会自动 provisioning 单个 credential：account 取 `outlook-mcp.json` 的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`；未设置时会自动生成随机 token 并在启动日志打印 setup URL。浏览器打开 `https://mcp.example.com/graph/you@outlook.com/replace-with-random-token-32-chars` 后才会展示 device code；未知 token 的 GET 返回 `credential_not_registered`，不会写 SQLite。完成登录后刷新页面，再让 MCP 客户端重试无 token URL。

`credential_store` 的 register/rotate CLI 只是运维恢复工具，正常启动不再需要手工 register。轮换 token 使用 `uv run python -m outlook_mcp.credential_store rotate graph you@outlook.com old-token-1234 new-token-5678`；旧 token 失效，但 credential ID 与 OAuth 状态保留。

WebUI 默认选中 `Full`，不会写 `X-Outlook-MCP-Mode` 或 `X-Outlook-MCP-Tools`，因此暴露全部高层工具。页面上可以直接点 `Full`、`Mail`、`Calendar`、`Contacts`、`Tasks`、`Custom`；选中项会高亮，MCP JSON 会自动切换。只有 `Custom` 会展示工具输入框并写入 `X-Outlook-MCP-Tools`。从 `Full` 切到 `Custom` 会默认填入全部工具，从单一 mode 切到 `Custom` 会默认填入该 mode 的工具，便于删减或微调。

同一个 credential 的操作日志页面：

```text
/{method}/{account}/{token}/log
```

日志页展示 HTTP MCP 调用摘要，支持 `range=all|24h|7d|30d|custom`，自定义区间使用 `from=YYYY-MM-DD&to=YYYY-MM-DD`。日志不会保存 raw secret、token、完整请求参数或完整响应体。

WebUI JSON API 只接受无 secret 路径的 JSON `POST`：

```text
POST /{method}/{account}/api/calendar/list
Authorization: Bearer {token}
Content-Type: application/json
```

Graph credential 完成授权后，可直接打开 Mail WebUI 和 Calendar WebUI：

```text
/{method}/{account}/{token}/mail
/{method}/{account}/{token}/calendar
```

页面提供顶层文件夹、每页 `20 items` 的列表、当前文件夹主题搜索、HTML 邮件详情、写信、HTML 回复/转发、移动、移入已删除邮件和显式已读/未读操作。所有写请求仍由页面携带 `confirm=true` 调用现有 `mail_*` API。

HTML 正文经过清理后在 sandbox iframe 中显示；远程图片默认阻断，`cid:` 内嵌图片首版显示占位。首版附件只展示名称、类型和大小，不支持下载、普通文件上传或邮件草稿。

`/{method}/{account}/{token}/api/...` 已禁用。API 和 upload 不校验请求 `Origin`，但仍必须携带有效 Bearer token。取消服务端 Origin 校验不等于启用了浏览器跨站 CORS/preflight。

HTTP 图片上传 endpoint：

```text
POST /{method}/{account}/uploads
Authorization: Bearer {token}
multipart/form-data field: file
```

上传前必须满足两个条件：`method + account + token` 已经注册为 HTTP credential，且该 credential 曾经完成过 Microsoft OAuth 授权。只有 pending device flow 或仅打开过 endpoint 页面不算完成授权，接口会返回 `403 auth_required`。

返回的 `relative_path`、`width`、`height` 可用于 `mail_send.attachments` 和 `calendar_create.attachments`：

```json
{"relative_path": "uploads/12/20260522/6f1a2b3c4d5e6789a0b1c2d3e4f56789/photo.png", "disposition": "attachment"}
```

正文内嵌图片使用：

```json
{"relative_path": "uploads/12/20260522/6f1a2b3c4d5e6789a0b1c2d3e4f56789/photo.png", "disposition": "inline", "content_id": "photo1"}
```

HTML 正文写 `<img src="cid:photo1">`，并设置 `html=true`。如果这个 `<img>` 没有 `width`、`height` 或 `style` 尺寸约束，服务端会按图片实际像素自动补默认显示尺寸；大图最大显示宽度为 `600px`，小图不会被放大。邮件和日程都支持普通附件与 inline 图片；带邀请人的 Graph 日程会通过 MIME/ICS 会议邀请发送附件，没有邀请人的普通 Graph 日程会创建 event 后再写入 event attachments。上传目录默认 `~/.outlook-mcp/uploads`，可用 `OUTLOOK_MCP_UPLOAD_DIR` 覆盖；默认只允许 PNG、JPEG、GIF、WebP，单文件上限 `3 MiB`（`OUTLOOK_MCP_UPLOAD_MAX_BYTES`），TTL `24 hours`（`OUTLOOK_MCP_UPLOAD_TTL_SECONDS`）。每个 credential 还有总量配额：默认 `30 MiB`（`OUTLOOK_MCP_UPLOAD_QUOTA_BYTES`）和 `100` 个文件（`OUTLOOK_MCP_UPLOAD_QUOTA_FILES`）。

HTTP 模式默认 SQLite 文件是 `~/.outlook-mcp/outlook-mcp.sqlite3`，可用 `OUTLOOK_MCP_DB` 覆盖。它是本地状态库：HTTP credential、msal cache、日历草稿、操作日志和上传元数据都存在这里。当前只支持 `graph`。

## 4.2 Docker demo

Docker demo 使用 HTTP MCP server 作为默认入口：

```bash
mkdir -p /volume1/Docker/outlookmcp/data
cp outlook-mcp.example.json /volume1/Docker/outlookmcp/outlook-mcp.json
export OUTLOOK_GRAPH_CLIENT_ID="9966331f-d9e2-42b2-85bf-5a54280ee60e"
export OUTLOOK_MCP_CSRF_KEY="$(openssl rand -hex 32)"
export OUTLOOK_MCP_ACCOUNT="you@outlook.com"
export OUTLOOK_MCP_BEARER_TOKEN="$(openssl rand -hex 32)"
docker compose -f docker-compose.demo.yml up -d
```

已有 Docker 部署升级到当前镜像（本地重新构建后重建容器）：

```bash
docker build -t outlook-mcp:0.1.48 .
docker compose -f docker-compose.demo.yml up -d --force-recreate
docker compose -f docker-compose.demo.yml logs -f outlookmcp
```

如果升级是为了处理 HTTPS 页面“不安全”提示，升级后检查页面源码不应再出现 `action="http://...`：

```bash
curl -sk "https://outlook-mcp.example.com/graph/$OUTLOOK_MCP_ACCOUNT/$OUTLOOK_MCP_BEARER_TOKEN" \
  | grep 'action="http://'
```

这个命令没有输出就是预期结果。

持久化约定：

- 本地构建镜像 `outlook-mcp:0.1.48`（`docker build -t outlook-mcp:0.1.48 .`）
- `/volume1/Docker/outlookmcp/outlook-mcp.json` 挂载到容器内 `/config/outlook-mcp.json:rw`
- `/volume1/Docker/outlookmcp/data` 挂载到 `/data`
- `OUTLOOK_MCP_DB=/data/outlook-mcp.sqlite3`
- 服务加入外部 Docker network `traefik`，默认域名是 `outlook-mcp.example.com`
- 页面提示的 MCP URL 会在浏览器端按当前 `window.location.origin` 重写；`OUTLOOK_MCP_TRUST_PROXY_HEADERS=1` 会信任 Traefik 注入的 `Forwarded` / `X-Forwarded-*`
- `OUTLOOK_GRAPH_CLIENT_ID=${OUTLOOK_GRAPH_CLIENT_ID}` 从宿主机环境变量注入，未设置时读取配置文件里的 `graph.client_id`
- `OUTLOOK_MCP_CSRF_KEY=${OUTLOOK_MCP_CSRF_KEY}` 必须是固定 secret；建议用 `openssl rand -hex 32` 生成

启动后打印授权 URL：

```bash
printf 'https://outlook-mcp.example.com/graph/%s/%s\n' \
  "$OUTLOOK_MCP_ACCOUNT" "$OUTLOOK_MCP_BEARER_TOKEN"
```

Docker demo 的 credential 由 server 启动时自动 provisioning：account 取 `/config/outlook-mcp.json` 的 `email`，token 取 `OUTLOOK_MCP_BEARER_TOKEN`，未设置时生成随机 token 并打印到容器日志。随后打开 `/graph/{account}/{token}` 获取 Microsoft device code；浏览器 GET 本身不会创建 credential。MCP 客户端后续使用 `/graph/{account}` + `Authorization: Bearer {token}`。credential 存在持久化的 `/data/outlook-mcp.sqlite3` 中。

## 5. 可用工具

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

`calendar_list`、`calendar_get`、`calendar_search`、`calendar_create`、`calendar_update` 支持可选 `timezone` 参数。定时日程不带 offset 时按该时区的本地墙上时间解释；返回的 `start` / `end` 是 UTC ISO 8601，`start_local` / `end_local` 是本地显示时间。全天日程返回 `YYYY-MM-DD`，其中 `end` 是 exclusive end date，不做时区偏移。

`calendar_create` 默认不设置 Outlook 提醒；如需提醒，传 `reminder_minutes_before_start=10` 表示开始前 10 分钟提醒。

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

所有写操作必须传 `confirm=true`，包括发送/回复/转发邮件、移动/删除邮件、标记已读/未读、创建/更新/删除日历、联系人和任务。

Graph 带邀请人的 `calendar_create` 不支持 `send_invites=false`；Graph `calendar_update` 也不支持 `send_invites=false`，这两种情况会返回 `validation_error`。

Graph 带邀请人的 `calendar_create` 返回 `sent=true` 和 `ics_uid`，`id` 为 `null`；后续更新或删除要先用 `calendar_list` / `calendar_search` 获取 Graph event id。

Graph 普通 event 创建成功、但 attachment 写入失败时返回 `ok=true` 与 `status="partial_success"`，并附带 `event_id`、`attachment_errors`；不要重试整个 create。Graph 写操作遇到 timeout 或 HTTP `5xx`、无法确认远端结果时返回 `error="outcome_unknown"`，后端不会自动重试，应先查询远端状态。

## 6. AI agent 使用约定

当工具返回 `auth_required`：

1. 把登录地址和验证码原样展示给用户。
2. 等用户确认已完成网页登录。
3. 重试原 MCP 工具调用。

当工具返回 `validation_error`：

1. 检查参数是否缺失。
2. 对写操作确认是否传入 `confirm=true`。

当工具返回 `exchange_error` 或 `outlook_error`：

1. 不要输出 token、client_id 之外的敏感信息。
2. 提醒用户检查 `outlook-mcp.json`、Graph 权限和 Microsoft 账户授权状态。
