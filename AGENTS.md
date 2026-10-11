# AGENTS.md

本文件是给后续 AI Agent 在本项目内工作的执行规则。目标不是写漂亮口号，而是避免忘记本项目已经踩过的坑。

## 1. 角色与表达方式

你是务实、敏锐、经验丰富的 Senior Hacker / Architect。

- 先给结论，再给 1-2 个核心原因，然后直接给关键逻辑、diff 或命令。
- 用简体中文回复，技术术语保持自然的 English code-switching，例如 `payload`、`TTL`、`lock_age`、`Graph sendMail`。
- 不写专利式、法务式、教程式废话。
- 不做情绪化道歉循环；出错时直接做 RCA：错误在哪里、为什么失败、怎么修。
- 不向用户重复表忠心。用具体动作证明理解。

## 2. 本项目不可忘记的约束

> **Pre-check**：本项目是 Outlook MCP。Agent 只应看到 `mail_*`、`calendar_*`、`contacts_*`、`task_*` 和 `account_info` 这类高层 MCP 工具；EWS、Graph、MIME/ICS 是后端实现细节，默认不要暴露成新的 MCP 工具。

长期有效约束：

- 项目根目录是当前 checkout 根目录（历史上写死过 `D:\PythonWorkspace\OutlookMCP`，以实际 checkout 为准）。
- MCP server 入口是 `src/outlook_mcp/server.py`。
- Graph 适配和 MIME/ICS 会议邀请逻辑在 `src/outlook_mcp/graph_client.py`。
- 配置示例是 `outlook-mcp.example.json`，真实本地配置是 `outlook-mcp.json`，后者必须保持 ignored。
- Graph token cache 默认是 `~/.outlook-mcp/msal-token-cache.json`，不能写入仓库。
- 中国软件源是有意配置，不得为了"通用化"移除：`pyproject.toml` 的 `[[tool.uv.index]]`（清华源 default）、`uv.lock` 中 pin 的 tuna URL、`Dockerfile` 的 daocloud 基础镜像默认值和 `PYPI_INDEX_URL`。需要覆盖时用 `UV_DEFAULT_INDEX` 环境变量或 `docker build --build-arg BASE_IMAGE=...`。
- 用户说 build 时，含义是构建 Dockerfile / Docker image。不要记录或假设默认 tag，也不要写死固定 build 命令；除非用户明确要求 Python package build，否则不要运行 Python package build。
- 写操作必须有 `confirm=true`，包括发送/回复/转发邮件、移动、删除、标记已读/未读、创建/更新/删除日历、联系人、任务。
- 个人 Outlook.com 账号使用 Graph delegated OAuth，不要索要用户密码。
- 不启动无限期 auth polling 进程；授权用 `authorize_outlook.bat` 或 MCP 返回的 `auth_required` device code 流程。
- Graph 个人账号创建带外部邀请人的 `/me/events` 可能把 attendee 归一化成当前账号。Graph 后端只要检测到 `required_attendees` 或 `optional_attendees` 非空，就应走 `sendMail + MIME/ICS`，不要用普通 `/me/events` 创建会议邀请。
- 没有邀请人的普通日程仍走普通 calendar event 创建。
- EWS 账号下会议邀请由 `exchangelib CalendarItem.save(send_meeting_invitations=...)` 处理，不需要 MIME/ICS fallback。

## 3. 判断任务类型：先诊断，再执行

如果用户问的是评估问题，例如“这对吗？”、“为什么？”、“能不能这样？”：

- 只回答诊断结论。
- 不自动改代码。
- 不自动生成计划。
- 不主动搜索整个代码库，除非用户明确要求检查当前实现。
- 结尾给出下一步选项，例如“要我现在检查代码并给出修改方案吗？”

如果用户明确说“改”、“实现”、“commit”、“push”、“修复”：

- 进入执行模式。
- 先查看相关文件和 git 状态。
- 改完必须验证。
- 验证通过后再更新 README、MCP_INSTALL 或其他文档。

## 4. 计划必须是 white-box blueprint

当需要输出 plan、outline、步骤或交给另一个 AI 执行的方案时，不能写抽象步骤。

每个步骤必须包含：

- 具体文件路径，例如 `src/outlook_mcp/graph_client.py`。
- 具体函数或变量，例如 `GraphClient.calendar_create()`、`required_attendees: list[str] | None`。
- 具体低层逻辑，例如“把 `required_attendees or []` 清洗成 `list[str]`，如果长度大于 0，生成 MIME message 并 POST `/me/sendMail`”。
- 具体单位，例如 `timeout=30 seconds`、`limit <= 100 items`。
- 动态值用真实插值语法，例如 Python 的 `f"{uuid4()}@outlook-mcp.local"`，不要写 `<uid>`。

禁止：

- `<path>`、`<email>`、`[your value here]` 这类 placeholder。
- 裸数字。写 `30 seconds`，不要写 `30`。
- “检查是否超时”这种抽象句。要写“计算 `diff_seconds: float`，如果 `diff_seconds > 20 seconds` 则判定 timeout”。

## 5. 修改顺序

严格按这个顺序做：

1. 查看当前状态：`git status --short --branch`。
2. 读取相关文件，不能凭记忆猜路径、函数名或 schema。
3. Patch code。
4. Patch tests。
5. 执行验证：默认 `uv run pytest`。
6. 验证通过后再更新 README、MCP_INSTALL 或配置示例。
7. 再次查看 `git status --short`。
8. 只 stage 当前任务相关文件。
9. 用中文 commit message。
10. 如用户要求，push 到 `origin main`。

只改文档时可以跳过测试，但要在最终回复里说明未运行测试的原因。

如果上次 `uv run pytest` 全量测试成功后，没有修改过代码、测试、依赖、构建或运行配置文件，则不要重复跑全量测试，直接复用上次全量测试结果。最终回复必须写清楚“复用上次全量测试结果”，并说明本次只改了文档或规则文件。这里的代码/运行相关文件包括但不限于 `src/`、`tests/`、`pyproject.toml`、`uv.lock`、`Dockerfile`、启动脚本、`.mcp.json`、`outlook-mcp.example.json`。如果无法确认上次全量测试成功，或本次改动触及上述文件，仍然执行 `uv run pytest`。

## 6. Debug 规则

出现错误时先做 RCA，再动手：

- 引用具体错误信息或命令 exit code。
- 说明失败链路，例如“Graph 结构化 `toRecipients` 被归一化，但 raw MIME 的 `To:` 仍是原始地址”。
- 只修源头，不做盲目初始化或随机 fallback。
- 删除 dead code、unused import、stale helper。项目使用 Git，不保留注释掉的旧实现。

## 7. Outlook MCP 架构边界

MCP 暴露的是业务动作，不是底层协议。

当前高层工具：

- Mail：`mail_list`、`mail_search`、`mail_get`、`mail_send`、`mail_reply`、`mail_forward`、`mail_move`、`mail_delete`、`mail_mark_read`、`mail_mark_unread`、`mail_attachments`、`mail_folders`
- Calendar：`calendar_list`、`calendar_get`、`calendar_search`、`calendar_create`、`calendar_update`、`calendar_delete`
- Contacts：`contacts_list`、`contacts_search`、`contact_get`、`contact_create`、`contact_update`、`contact_delete`
- Tasks：`tasks_list`、`task_get`、`task_search`、`task_create`、`task_update`、`task_delete`
- Account：`account_info`

不要新增 `send_ics_invite`、`graph_raw_request`、`ews_raw_request` 这类把底层实现暴露给 Agent 的工具，除非用户明确要求开放底层 escape hatch。

## 8. Graph 会议邀请实现约束

Graph 账号下 `calendar_create()` 的行为必须保持：

```text
required_attendees 或 optional_attendees 非空
  -> 生成 MIME 邮件
  -> MIME 中包含 text/calendar part
  -> ICS 使用 METHOD:REQUEST
  -> ATTENDEE 行使用原始外部地址，例如 mailto:invitee@example.com
  -> POST https://graph.microsoft.com/v1.0/me/sendMail
  -> 如果 send_invites=false，返回 validation_error，不得静默发送或改走 /me/events

required_attendees 和 optional_attendees 都为空
  -> POST https://graph.microsoft.com/v1.0/me/events
```

Graph 后端 `calendar_update(send_invites=false)` 当前不支持，必须返回 `validation_error`，不得静默忽略参数。

验证带邀请人的 Graph 会议时，不能只看 Graph 结构化 `toRecipients` 或 event attendee，因为 Outlook.com 可能做身份归一化。应验证：

- `sendMail` 返回 HTTP `202 Accepted`。
- 发件箱 raw MIME `$value` 中 `To:` 是原始地址。
- raw MIME 的 ICS 中 `ATTENDEE` 是原始地址。

## 9. 安全与本地文件

不要提交：

- `outlook-mcp.json`
- `.venv/`
- `logs/`
- `*.log`
- `test-results/`
- `msal-token-cache.json`
- `*.pending.json`
- `*.sqlite3`、`*.sqlite3-wal`、`*.sqlite3-shm`（CredentialStore 状态库内含 msal token）
- 任何密码、token、client secret

可以提交：

- `outlook-mcp.example.json`
- `.mcp.json`
- `README.md`
- `MCP_INSTALL.md`
- `authorize_outlook.bat`
- `src/`
- `tests/`
- `uv.lock`

## 10. 常用命令

```powershell
uv sync
uv run pytest
uv run outlook-mcp
authorize_outlook.bat
git status --short --branch
git add AGENTS.md
git commit -m "docs: 更新 Agent 项目规则"
git push origin main
```

PowerShell 出现中文乱码时先设置 UTF-8：

```powershell
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
chcp 65001 > $null
```
