# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
uv run pytest                    # Run all tests (~430 tests)
uv run pytest tests/test_file.py # Run a single test file
uv run pytest -k "test_name"     # Run tests matching keyword
uv run ruff check src/           # Lint
uv run outlook-mcp               # Start stdio MCP server
uv run outlook-mcp-http          # Start HTTP MCP server (port 8000)
```

When the user says "build", interpret it as building the Docker image from `Dockerfile`. Do not record or assume a default tag or fixed build command — the user must specify the tag each time. Do not run Python package builds unless the user explicitly asks for one.

The repository has a staged `uv.lock` — do not regenerate it unless the user asks.

## Architecture

This is an MCP (Model Context Protocol) server that lets AI agents operate Outlook mail/calendar/contacts/tasks through high-level tools. The agent calls MCP tools (`mail_send`, `calendar_create`, etc.) — it never knows or cares whether the backend is EWS or Microsoft Graph.

### Two entry points

- **`server.py`** — stdio MCP mode, using `FastMCP`. Used for local MCP clients (Claude Desktop, VS Code). Direct EWS or Graph via config file.
- **`http_server.py`** — Streamable HTTP MCP mode (Starlette + uvicorn). Used for Docker/Traefik deployments. **Graph-only** in HTTP mode (`ews` returns `method_unsupported`).

### Two backends

Every tool (`mail_send`, `calendar_create`, etc.) checks `hasattr(client, "graph_method")` to dispatch:
- **EWS** (`client.py`): Uses `exchangelib` (Message, CalendarItem, Task, Contact)
- **Graph** (`graph_client.py`): Uses `requests.Session` + Microsoft Graph REST API

The Graph path handles calendar invites differently: with attendees → constructs MIME message with ICS attachment and sends via `/me/sendMail` (because Outlook.com personal accounts would normalize external addresses). Without attendees → straightforward REST `/me/events`.

### HTTP auth flow (important)

HTTP endpoints follow `/{method}/{account}/{secret}`. The endpoint must be registered explicitly first (`uv run python -m outlook_mcp.credential_store register ...`); a browser GET with an unknown token returns `credential_not_registered` and never writes to SQLite. Opening a registered URL shows a Microsoft device-code OAuth page. The MCP client uses the same URL; credentials live in `credential_store.py` (SQLite). MSAL token cache and pending device flows are stored in SQLite, not files.

### Key patterns

- **`run_safely`** in `tools/common.py`: Wraps every tool action to catch exceptions and return structured `{"ok": False, "error": ...}` responses. All tool implementations use this.
- **`credential_store.py`**: `_connect()` is a `@contextmanager` that creates a **new** SQLite connection per call, runs PRAGMAs, commits on success, rolls back on exception, and closes in `finally`. All DB access goes through it. No shared/global connection — each call gets its own connection so concurrent requests are safe.
- **`server.py` globals**: `_http_clients` (OrderedDict LRU cache of GraphClient instances), `_http_store`, `_http_base_config` are module-level and protected by `_http_lock` (threading.Lock). `_client` is protected by `_client_lock`.
- **All write tools require `confirm=True`** to prevent accidental mutations by LLMs.
- **Alias registry** (`tools/alias_registry.py`): Long EWS/Graph item IDs are mapped to short deterministic aliases (`cal:a3f29b`) so LLMs don't truncate them across tool calls.

### SQLite schema

Five tables in `outlook-mcp.sqlite3`:
- `http_credentials` — endpoint registration (method + account + secret_hash)
- `graph_authorizations` — MSAL token cache and pending device flow per credential (one-time automatic migration drops the legacy `graph_oauth_state` table)
- `http_operation_logs` — buffered operation audit log (5s flush interval, 30-day retention, batched DELETE pruning)
- `calendar_drafts` — credential-scoped WebUI create/update drafts with optimistic versions, submission state, projections, and optional expiry
- `calendar_draft_settings` — per-credential draft retention (`NULL` means permanent)

Calendar draft CRUD uses the same Bearer credential as the WebUI but does not require a live Graph OAuth session.
Only draft submission obtains a `GraphClient`. Drafts remain an HTTP WebUI concern and are not exposed as MCP tools.

### Uploads

`uploads.py` handles multipart file uploads at `/{method}/{account}/uploads` (Bearer auth; the secret-scoped path variant was removed). Files stored under `OUTLOOK_MCP_UPLOAD_DIR` with per-credential subdirectories. 3 MiB limit, PNG/JPEG/GIF/WebP only, 24h TTL, per-credential quota 30 MiB / 100 files (`OUTLOOK_MCP_UPLOAD_*` envs), image dimensions auto-detected.

### Environment variables

| Variable | Default | Notes |
|---|---|---|
| `OUTLOOK_CONFIG_FILE` | — | Path to config JSON (required at runtime) |
| `OUTLOOK_MCP_DB` | `~/.outlook-mcp/outlook-mcp.sqlite3` | SQLite database path; Docker sets `/data/outlook-mcp.sqlite3` |
| `OUTLOOK_MCP_HTTP_HOST` | `127.0.0.1` | Bind host for HTTP server; Docker sets `0.0.0.0` |
| `OUTLOOK_MCP_HTTP_PORT` | `8000` | Bind port for HTTP server |
| `OUTLOOK_MCP_TRUST_PROXY_HEADERS` | — | Trust X-Forwarded-* from proxies allowed by `FORWARDED_ALLOW_IPS` (default `127.0.0.1`) |
| `OUTLOOK_GRAPH_CLIENT_ID` | — | Overrides `graph.client_id` from the config file |
| `OUTLOOK_MCP_SECRET_MIN_LENGTH` | `16` | Minimum Bearer token length for register/rotate |
| `OUTLOOK_MCP_CSRF_KEY` | random per boot | Fixed HMAC key for WebUI CSRF tokens |
| `OUTLOOK_MCP_UPLOAD_DIR` | `~/.outlook-mcp/uploads` | Upload temp directory |
| `OUTLOOK_MCP_UPLOAD_MAX_BYTES` | `3145728` | Single upload limit (3 MiB) |
| `OUTLOOK_MCP_UPLOAD_TTL_SECONDS` | `86400` | Upload TTL (24 hours) |
| `OUTLOOK_MCP_UPLOAD_QUOTA_BYTES` | `31457280` | Per-credential upload quota (30 MiB) |
| `OUTLOOK_MCP_UPLOAD_QUOTA_FILES` | `100` | Per-credential upload file count quota |
| `OUTLOOK_MCP_ACCOUNT` / `OUTLOOK_MCP_BEARER_TOKEN` | — | Required by the debug launcher scripts (`start_outlook_mcp_http.sh` / `.bat`) |

### Deployment

- **`Dockerfile`**: Multi-stage-ish with uv, runs `outlook-mcp-http` as UID 1000. Health check hits `/health`. Builds from current repo source — `pyproject.toml` version is authoritative.
- **`docker-compose.demo.yml`**: Demo stack with Traefik labels for `outlook-mcp.example.com`, TLS via Let's Encrypt, bind-mounts config and data dirs.
