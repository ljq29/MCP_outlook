"""HTTP transport entry point for Outlook MCP."""

from __future__ import annotations

import asyncio
import hmac
import html
import inspect
import json
import logging
import os
import secrets
import shlex
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlsplit

import uvicorn
from starlette.datastructures import UploadFile
from starlette.formparsers import MultiPartException
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, PlainTextResponse
from starlette.staticfiles import StaticFiles

from . import __version__, server
from .calendar_drafts import dispatch_calendar_draft_api, is_calendar_draft_api
from .config import EMAIL_RE, RuntimeConfig, load_config
from .credential_store import DEFAULT_DB_PATH, CredentialRecord, CredentialStore, hash_secret
from .graph_client import GraphClient
from .launcher_json import validate_secret
from .mail_cache import MailResponseCache
from .tools.common import (
    DEFAULT_CALENDAR_TIMEZONE,
    RequestBodyTooLargeError,
    ToolValidationError,
    run_safely,
)
from .uploads import (
    cleanup_expired_uploads,
    reset_upload_credential_id,
    save_uploaded_file,
    set_upload_credential_id,
    upload_max_bytes,
)


HTTP_METHODS = {"GET", "POST", "DELETE"}
ENDPOINT_METHODS = {"graph"}
DEFAULT_EXAMPLE_TOKEN = "example-token-not-registered"
DEFAULT_EXAMPLE_PATH = "/graph/you@outlook.com"
DEFAULT_EXAMPLE_AUTH_PATH = f"{DEFAULT_EXAMPLE_PATH}/{DEFAULT_EXAMPLE_TOKEN}"
EXPECTED_ENDPOINT_PATH = "/{method}/{account}"
MCP_SERVER_NAME = "outlook-http"
LOG_RANGES = {"all", "24h", "7d", "30d", "custom"}
AUTH_ACTIONS = {"auth/cancel", "auth/reauthorize"}
_PROCESS_STARTED_AT = time.monotonic()
_MAIL_API_CACHE_TTLS = {
    "mail/folders": 300.0,
    "mail/list": 30.0,
    "mail/search": 30.0,
    "mail/get": 600.0,
    "mail/attachments": 300.0,
}
_MAIL_API_WRITE_PATHS = frozenset(
    {
        "mail/send",
        "mail/reply",
        "mail/forward",
        "mail/move",
        "mail/delete",
        "mail/mark_read",
        "mail/mark_unread",
    }
)
_MAIL_RESPONSE_CACHE = MailResponseCache(max_entries=256)
TOOL_MODE_PRESETS: tuple[dict[str, str | None], ...] = (
    {"mode": "full", "label": "Full", "tools": None},
    {
        "mode": "mail",
        "label": "Mail",
        "tools": (
            "mail_list,mail_search,mail_get,mail_send,mail_reply,mail_forward,"
            "mail_move,mail_delete,mail_mark_read,mail_mark_unread,mail_attachments,mail_folders,account_info"
        ),
    },
    {
        "mode": "calendar",
        "label": "Calendar",
        "tools": "calendar_list,calendar_get,calendar_search,calendar_create,calendar_update,calendar_delete,account_info",
    },
    {
        "mode": "contacts",
        "label": "Contacts",
        "tools": "contacts_list,contacts_search,contact_get,contact_create,contact_update,contact_delete,account_info",
    },
    {
        "mode": "tasks",
        "label": "Tasks",
        "tools": "tasks_list,task_get,task_search,task_create,task_update,task_delete,account_info",
    },
    {"mode": "custom", "label": "Custom", "tools": ""},
)
TOOL_MODE_BY_NAME = {str(item["mode"]): item for item in TOOL_MODE_PRESETS}

_CSRF_KEY = os.environ.get("OUTLOOK_MCP_CSRF_KEY")
if not _CSRF_KEY:
    _CSRF_KEY = secrets.token_hex(32)
    logging.warning(
        "\nOUTLOOK_MCP_CSRF_KEY is not set — a random key was generated.\n"
        "  Single-worker: CSRF works but tokens expire on restart.\n"
        "  Multi-worker:  Each worker has a DIFFERENT key — CSRF verification WILL FAIL.\n"
        "  Set OUTLOOK_MCP_CSRF_KEY to a fixed secret (e.g. openssl rand -hex 32) for production."
    )

_CSRF_TTL_SECONDS = 900


def _memory_rss_bytes() -> int | None:
    try:
        import resource
    except ImportError:
        return None
    try:
        max_rss = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    except (OSError, ValueError):
        return None
    if sys.platform == "darwin":
        return max_rss
    return max_rss * 1024


def _load_average() -> list[float] | None:
    getloadavg = getattr(os, "getloadavg", None)
    if getloadavg is None:
        return None
    try:
        return [round(value, 3) for value in getloadavg()]
    except OSError:
        return None


def _health_payload() -> dict[str, Any]:
    return {
        "status": "ok",
        "service": MCP_SERVER_NAME,
        "uptime_seconds": round(time.monotonic() - _PROCESS_STARTED_AT, 3),
        "resources": {
            "process_cpu_seconds": round(time.process_time(), 3),
            "memory_rss_bytes": _memory_rss_bytes(),
            "load_average": _load_average(),
            "cpu_count": os.cpu_count(),
        },
    }


def _deep_health_payload() -> dict[str, Any]:
    return {
        "status": "ok",
        "service": MCP_SERVER_NAME,
        "checks": {
            "process": _health_payload(),
        },
    }


def _read_static_text(filename: str) -> str:
    try:
        return (Path(__file__).with_name("static") / filename).read_text(encoding="utf-8")
    except OSError:
        logging.getLogger("outlook_mcp").warning("failed to load static asset %s", filename, exc_info=True)
        return ""


PICO_CSS = _read_static_text("pico.min.css")
HTML_PAGE_CSS = _read_static_text("outlook-webui.css")
CALENDAR_CSS = _read_static_text("calendar.css")
MAIL_CSS = _read_static_text("mail.css")
FULLCALENDAR_JS = "fullcalendar-index.global.min.js"
FULLCALENDAR_LOCALE_JS = "fullcalendar-zh-cn.global.min.js"
TINYMCE_JS = "tinymce/tinymce.min.js"
STATIC_ROOT = Path(__file__).with_name("static")

CALENDAR_CAPABILITIES = {
    "provider": "graph",
    "supports_send_invites_toggle_on_create_without_attendees": True,
    "supports_send_invites_false_with_attendees_create": False,
    "supports_send_invites_false_on_update": False,
    "supports_delete_without_cancellation": False,
    "disable_drag_resize_for_all_day_meetings": True,
}

def _static_path(scope: dict[str, Any], filename: str) -> str:
    return f"{_request_prefix(scope)}/static/{filename}".replace("//", "/")


def _versioned_static_path(scope: dict[str, Any], filename: str) -> str:
    return f"{_static_path(scope, filename)}?v={__version__}"


def _request_prefix(scope: dict[str, Any]) -> str:
    headers = _headers(scope)
    trust_proxy = _trust_proxy_headers()
    forwarded_prefix = _first_header_value(headers.get("x-forwarded-prefix")) if trust_proxy else ""
    return _normalized_prefix(forwarded_prefix or str(scope.get("root_path", "")))


def _json_script_value(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")

# Per-credential upload rate limiting (token bucket)
_UPLOAD_RATE = 10.0 / 60.0    # 10 uploads per minute per credential
_UPLOAD_BURST = 5.0           # burst capacity
_UPLOAD_BUCKETS: dict[int, tuple[float, float]] = {}
_UPLOAD_BUCKETS_LOCK = threading.Lock()


def _check_upload_rate(credential_id: int) -> bool:
    """Return True if the upload is allowed under the rate limit for *credential_id*."""
    now = time.monotonic()
    with _UPLOAD_BUCKETS_LOCK:
        tokens, last = _UPLOAD_BUCKETS.get(credential_id, (_UPLOAD_BURST, now))
        elapsed = now - last
        tokens = min(_UPLOAD_BURST, tokens + elapsed * _UPLOAD_RATE)
        if tokens >= 1.0:
            _UPLOAD_BUCKETS[credential_id] = (tokens - 1.0, now)
            # Lazy cleanup: evict oldest if too many buckets
            if len(_UPLOAD_BUCKETS) > 1000:
                oldest = min(
                    _UPLOAD_BUCKETS,
                    key=lambda bucket_id: _UPLOAD_BUCKETS[bucket_id][1],
                )
                del _UPLOAD_BUCKETS[oldest]
            return True
        _UPLOAD_BUCKETS[credential_id] = (tokens, now)
        return False


# Per-source-IP authentication failure rate limiting (token bucket)
_AUTH_FAILURE_RATE = 1.0        # 1 failed auth per second sustained per source IP
_AUTH_FAILURE_BURST = 10.0      # burst capacity
_AUTH_FAILURE_BUCKETS: dict[str, tuple[float, float]] = {}
_AUTH_FAILURE_BUCKETS_LOCK = threading.Lock()


def _client_source(scope: dict[str, Any]) -> str:
    client = scope.get("client")
    if isinstance(client, (list, tuple)) and client:
        return str(client[0])
    return "unknown"


def _auth_rate_limited_response(scope: dict[str, Any]) -> JSONResponse:
    return JSONResponse(
        {"ok": False, "error": "rate_limited", "message": "Too many failed authentication attempts. Retry later."},
        status_code=429,
    )


def _check_auth_rate(source: str) -> bool:
    """Return True if another failed authentication from *source* is allowed."""
    now = time.monotonic()
    with _AUTH_FAILURE_BUCKETS_LOCK:
        tokens, last = _AUTH_FAILURE_BUCKETS.get(source, (_AUTH_FAILURE_BURST, now))
        elapsed = now - last
        tokens = min(_AUTH_FAILURE_BURST, tokens + elapsed * _AUTH_FAILURE_RATE)
        if tokens >= 1.0:
            _AUTH_FAILURE_BUCKETS[source] = (tokens - 1.0, now)
            if len(_AUTH_FAILURE_BUCKETS) > 1000:
                oldest = min(
                    _AUTH_FAILURE_BUCKETS,
                    key=lambda bucket_key: _AUTH_FAILURE_BUCKETS[bucket_key][1],
                )
                del _AUTH_FAILURE_BUCKETS[oldest]
            return True
        _AUTH_FAILURE_BUCKETS[source] = (tokens, now)
        return False


def _make_csrf_token(action: str, path: str) -> str:
    expires = int(time.time()) + _CSRF_TTL_SECONDS
    msg = f"action={action}&path={path}&expires={expires}"
    mac = hmac.new(_CSRF_KEY.encode("utf-8"), msg.encode("utf-8"), "sha256").hexdigest()
    return f"{expires}:{mac}"


def _verify_csrf_token(token: str, action: str, path: str) -> bool:
    try:
        expires_str, mac = token.split(":", 1)
        expires = int(expires_str)
    except (ValueError, TypeError):
        return False
    if time.time() > expires:
        return False
    msg = f"action={action}&path={path}&expires={expires}"
    expected = hmac.new(_CSRF_KEY.encode("utf-8"), msg.encode("utf-8"), "sha256").hexdigest()
    return hmac.compare_digest(expected, mac)

MCP_INSTALL_SCRIPT = """
<script>
(() => {
  function mcpConfig(serverName, endpointUrl, bearerToken, mode, tools) {
    const headers = {
      Authorization: `Bearer ${bearerToken}`,
    };
    const normalizedMode = (mode || "full").trim().toLowerCase();
    const customTools = (tools || "").trim();
    if (normalizedMode === "custom") {
      if (customTools) {
        headers["X-Outlook-MCP-Tools"] = customTools;
      } else {
        headers["X-Outlook-MCP-Mode"] = "custom";
      }
    } else if (normalizedMode && normalizedMode !== "full") {
      headers["X-Outlook-MCP-Mode"] = normalizedMode;
    }
    return {
      mcpServers: {
        [serverName]: {
          type: "streamable-http",
          url: endpointUrl,
          headers,
        },
      },
    };
  }

  function agentPrompt(mcpJson) {
    return [
      "把下面的 MCP server 配置 merge 到当前项目的 .mcp.json 或你的 MCP 客户端配置里。",
      "保留已有 mcpServers，不要覆盖其他 server。",
      "",
      mcpJson,
    ].join("\\n");
  }

  function selectedMode(section) {
    const selected = section.querySelector("[data-mcp-mode][aria-pressed='true']");
    return selected ? selected.getAttribute("data-mode") || "full" : "full";
  }

  function selectedTools(section) {
    if (selectedMode(section) === "custom") {
      const input = section.querySelector("[data-mcp-custom-tools]");
      return input ? input.value.trim() : "";
    }
    return "";
  }

  function maybeSeedCustomTools(section, previousSelected, selectedButton) {
    if (selectedButton.getAttribute("data-mode") !== "custom") {
      return;
    }
    if (previousSelected === selectedButton) {
      return;
    }
    const customInput = section.querySelector("[data-mcp-custom-tools]");
    if (!customInput || customInput.getAttribute("data-custom-tools-dirty") === "1") {
      return;
    }
    const seedSource = previousSelected || selectedButton;
    const seedTools = seedSource.getAttribute("data-custom-tools") || "";
    customInput.value = seedTools;
  }

  function endpointUrlFor(section) {
    const endpointPath = section.getAttribute("data-endpoint-path") || "";
    const fallbackUrl = section.getAttribute("data-endpoint-url") || "";
    if (!endpointPath || !fallbackUrl) {
      return "";
    }
    const origin = window.location.origin;
    return origin ? origin + endpointPath : fallbackUrl;
  }

  function updateInstallSection(section) {
    const endpointUrl = endpointUrlFor(section);
    if (!endpointUrl) {
      return;
    }
    const serverName = section.getAttribute("data-server-name") || "outlook-http";
    const bearerToken = section.getAttribute("data-bearer-token") || "";
    const mode = selectedMode(section);
    const tools = selectedTools(section);
    const mcpJson = JSON.stringify(mcpConfig(serverName, endpointUrl, bearerToken, mode, tools), null, 2);

    section.querySelectorAll("[data-mcp-endpoint-url]").forEach((element) => {
      element.textContent = endpointUrl;
    });
    section.querySelectorAll("[data-copy-name='mcp-json']").forEach((element) => {
      element.textContent = mcpJson;
    });
    section.querySelectorAll("[data-copy-name='agent-prompt']").forEach((element) => {
      element.textContent = agentPrompt(mcpJson);
    });

    const customInput = section.querySelector("[data-mcp-custom-tools]");
    const customField = section.querySelector("[data-mcp-custom-field]");
    const customSelected = Boolean(section.querySelector("[data-mcp-mode][data-mode='custom'][aria-pressed='true']"));
    if (customInput) {
      customInput.disabled = !customSelected;
    }
    if (customField) {
      customField.hidden = !customSelected;
    }
  }

  function rewriteMcpInstallUrls() {
    document.querySelectorAll("[data-mcp-install]").forEach((section) => {
      updateInstallSection(section);
    });
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise((resolve, reject) => {
      const textarea = document.createElement("textarea");
      textarea.value = text;
      textarea.setAttribute("readonly", "");
      textarea.style.position = "fixed";
      textarea.style.left = "-9999px";
      document.body.appendChild(textarea);
      textarea.select();
      try {
        if (document.execCommand("copy")) {
          resolve();
        } else {
          reject(new Error("copy command failed"));
        }
      } catch (error) {
        reject(error);
      } finally {
        document.body.removeChild(textarea);
      }
    });
  }

  function setCopyButtonState(button, label, state) {
    const originalLabel = button.getAttribute("data-copy-label") || button.textContent;
    button.setAttribute("data-copy-label", originalLabel);
    button.setAttribute("data-copy-state", state);
    button.textContent = label;
    window.setTimeout(() => {
      button.removeAttribute("data-copy-state");
      button.textContent = originalLabel;
    }, 1400);
  }

  function bindCopyButtons() {
    document.querySelectorAll("[data-copy-source]").forEach((button) => {
      if (button.getAttribute("data-copy-bound") === "1") {
        return;
      }
      button.setAttribute("data-copy-bound", "1");
      button.addEventListener("click", async () => {
        const section = button.closest("[data-mcp-install]");
        const source = button.getAttribute("data-copy-source") || "";
        const target = section ? section.querySelector(`[data-copy-name="${source}"]`) : null;
        if (!target) {
          setCopyButtonState(button, "失败", "error");
          return;
        }
        try {
          await copyText(target.textContent.trim());
          setCopyButtonState(button, "已复制", "done");
        } catch (error) {
          setCopyButtonState(button, "失败", "error");
        }
      });
    });
  }

  function bindModeControls() {
    document.querySelectorAll("[data-mcp-install]").forEach((section) => {
      section.querySelectorAll("[data-mcp-mode]").forEach((button) => {
        if (button.getAttribute("data-mode-bound") === "1") {
          return;
        }
        button.setAttribute("data-mode-bound", "1");
        button.addEventListener("click", () => {
          const previousSelected = section.querySelector("[data-mcp-mode][aria-pressed='true']");
          maybeSeedCustomTools(section, previousSelected, button);
          section.querySelectorAll("[data-mcp-mode]").forEach((item) => {
            item.setAttribute("aria-pressed", item === button ? "true" : "false");
          });
          updateInstallSection(section);
        });
      });
      const customInput = section.querySelector("[data-mcp-custom-tools]");
      if (customInput && customInput.getAttribute("data-mode-bound") !== "1") {
        customInput.setAttribute("data-mode-bound", "1");
        customInput.addEventListener("input", () => {
          customInput.setAttribute("data-custom-tools-dirty", "1");
          section.querySelectorAll("[data-mcp-mode]").forEach((item) => {
            item.setAttribute("aria-pressed", item.getAttribute("data-mode") === "custom" ? "true" : "false");
          });
          updateInstallSection(section);
        });
      }
    });
  }

  function boot() {
    bindModeControls();
    rewriteMcpInstallUrls();
    bindCopyButtons();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
</script>
"""


@dataclass(frozen=True)
class EndpointRoute:
    method: str
    account: str
    secret: str | None = None
    action: str | None = None


def _lookup_single_account_endpoint(
    store: CredentialStore,
    base_config: RuntimeConfig | None,
    method: str,
    account: str,
    secret: str,
) -> CredentialRecord | None:
    """Single-user guard: only the endpoint matching the configured Graph account is served."""
    record = store.lookup_endpoint(method, account, secret)
    if record is None or base_config is None:
        return record
    if (record.method, record.account) != ("graph", base_config.email):
        return None
    return record


class EndpointScopedMCPApp:
    def __init__(self, store: CredentialStore, base_config: RuntimeConfig | None = None):
        self.store = store
        self.base_config = base_config
        self.mcp_app = server.mcp.streamable_http_app()
        self.static_app = StaticFiles(directory=STATIC_ROOT)

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        if scope["type"] != "http":
            await self.mcp_app(scope, receive, send)
            return

        path = _request_path(scope)
        if path.startswith("/static/") and scope["method"] in {"GET", "HEAD"}:
            static_scope = dict(scope)
            static_scope["path"] = path[len("/static"):] or "/"
            static_scope["raw_path"] = static_scope["path"].encode("utf-8")
            await self.static_app(static_scope, receive, send)
            return

        if path == "/health":
            await JSONResponse(_health_payload())(scope, receive, send)
            return

        if path == "/health/deep":
            await JSONResponse(_deep_health_payload())(scope, receive, send)
            return

        try:
            route = _parse_endpoint(path)
        except ValueError as exc:
            await _invalid_secret_response(scope, str(exc))(scope, receive, send)
            return
        if route is None:
            await _invalid_endpoint_response(scope)(scope, receive, send)
            return

        method, account = route.method, route.account
        auth_url = _auth_url(scope, method, account, route.secret or DEFAULT_EXAMPLE_TOKEN)
        if route.action == "log":
            if scope["method"] != "GET" or _wants_mcp_stream(scope) or route.secret is None:
                await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
                return
            secret = route.secret
            auth_url = _auth_url(scope, method, account, secret)
            record = await asyncio.to_thread(
                _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
            )
            if record is None and not _check_auth_rate(_client_source(scope)):
                await _auth_rate_limited_response(scope)(scope, receive, send)
                return
            if record is None:
                await _unregistered_credential_page(method, account, secret)(scope, receive, send)
                return
            response = await asyncio.to_thread(_log_page, self.store, scope, method, account, record)
            await response(scope, receive, send)
            return

        if route.action == "calendar":
            if scope["method"] != "GET" or _wants_mcp_stream(scope) or route.secret is None:
                await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
                return
            secret = route.secret
            if method != "graph":
                await _calendar_unsupported_page(method, account, auth_url)(scope, receive, send)
                return
            record = await asyncio.to_thread(
                _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
            )
            if record is None and not _check_auth_rate(_client_source(scope)):
                await _auth_rate_limited_response(scope)(scope, receive, send)
                return
            if record is None:
                await _unregistered_credential_page(method, account, secret)(scope, receive, send)
                return
            await _calendar_page(scope, method, account, secret)(scope, receive, send)
            return

        if route.action == "mail":
            if scope["method"] != "GET" or _wants_mcp_stream(scope) or route.secret is None:
                await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
                return
            secret = route.secret
            if method != "graph":
                await _mail_unsupported_page(method, account, auth_url)(scope, receive, send)
                return
            record = await asyncio.to_thread(
                _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
            )
            if record is None and not _check_auth_rate(_client_source(scope)):
                await _auth_rate_limited_response(scope)(scope, receive, send)
                return
            if record is None:
                await _unregistered_credential_page(method, account, secret)(scope, receive, send)
                return
            await _mail_page(scope, method, account, secret)(scope, receive, send)
            return

        if route.action in AUTH_ACTIONS:
            if scope["method"] != "POST" or _wants_mcp_stream(scope) or route.secret is None:
                await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
                return
            secret = route.secret
            auth_url = _auth_url(scope, method, account, secret)
            body = await _read_request_body(receive, max_bytes=_MAX_AUTH_BODY_BYTES)
            try:
                form_body = body.decode("utf-8")
            except UnicodeDecodeError:
                form_body = body.decode("utf-8", errors="replace")
            form_data = parse_qs(form_body)
            csrf_token = form_data.get("csrf_token", [""])[0]
            csrf_action = "reauthorize" if route.action == "auth/reauthorize" else "cancel"
            if not _verify_csrf_token(csrf_token, csrf_action, auth_url):
                await _auth_action_csrf_error_page(method, account, auth_url)(scope, receive, send)
                return
            record = await asyncio.to_thread(
                _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
            )
            if record is None and not _check_auth_rate(_client_source(scope)):
                await _auth_rate_limited_response(scope)(scope, receive, send)
                return
            if record is None:
                await _unregistered_credential_page(method, account, secret)(scope, receive, send)
                return
            await asyncio.to_thread(server.clear_http_auth_state, record.id, account)
            _MAIL_RESPONSE_CACHE.invalidate(record.id)
            if route.action == "auth/reauthorize":
                try:
                    mode, tools = _install_mode_and_tools(scope)
                except ToolValidationError as exc:
                    await _html_page("Outlook MCP Endpoint", "validation_error", f"<p>{html.escape(str(exc))}</p>", status_code=400)(scope, receive, send)
                    return
                response = await asyncio.to_thread(
                    _auth_page,
                    method,
                    account,
                    record,
                    auth_url,
                    mode=mode,
                    tools=tools,
                )
                await response(scope, receive, send)
            else:
                await _auth_cleared_page(method, account, auth_url)(scope, receive, send)
            return

        if route.action == "uploads":
            response = await _upload_response(self.store, self.base_config, scope, receive, method, account)
            await response(scope, receive, send)
            return

        if isinstance(route.action, str) and route.action.startswith("api/"):
            if route.secret is not None:
                await JSONResponse(
                    {
                        "ok": False,
                        "error": "invalid_endpoint",
                        "message": (
                            "Secret-scoped API paths are disabled. Use /{method}/{account}/api/... "
                            "with an Authorization: Bearer header."
                        ),
                    },
                    status_code=404,
                )(scope, receive, send)
                return
            await _api_route(route.action, scope, receive, send, method, account, self.store, self.base_config)
            return

        if scope["method"] == "GET" and not _wants_mcp_stream(scope):
            if route.secret is None:
                await _invalid_endpoint_response(scope)(scope, receive, send)
                return
            secret = route.secret
            auth_url = _auth_url(scope, method, account, secret)
            record = await asyncio.to_thread(
                _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
            )
            if record is None and not _check_auth_rate(_client_source(scope)):
                await _auth_rate_limited_response(scope)(scope, receive, send)
                return
            if record is None:
                await _unregistered_credential_page(method, account, secret)(scope, receive, send)
                return
            try:
                mode, tools = _install_mode_and_tools(scope)
            except ToolValidationError as exc:
                await _html_page("Outlook MCP Endpoint", "validation_error", f"<p>{html.escape(str(exc))}</p>", status_code=400)(scope, receive, send)
                return
            response = await asyncio.to_thread(
                _auth_page,
                method,
                account,
                record,
                auth_url,
                mode=mode,
                tools=tools,
            )
            await response(scope, receive, send)
            return

        if route.secret is not None:
            await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
            return

        if scope["method"] not in HTTP_METHODS:
            await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
            return

        try:
            secret = _bearer_token(scope)
            enabled_tools = _tools_allowlist(scope)
        except ToolValidationError as exc:
            await JSONResponse({"ok": False, "error": "validation_error", "message": str(exc)}, status_code=400)(scope, receive, send)
            return
        except ValueError as exc:
            await JSONResponse({"ok": False, "error": "invalid_secret", "message": str(exc)}, status_code=400)(scope, receive, send)
            return
        if secret is None:
            await _missing_bearer_response(scope)(scope, receive, send)
            return
        auth_url = _auth_url(scope, method, account, secret)
        record = await asyncio.to_thread(
            _lookup_single_account_endpoint, self.store, self.base_config, method, account, secret
        )
        if record is None and not _check_auth_rate(_client_source(scope)):
            await _auth_rate_limited_response(scope)(scope, receive, send)
            return
        context = server.HttpRequestContext(
            method=method,
            account=account,
            secret_hash=hash_secret(secret),
            credential_id=record.id if record else None,
            auth_url=auth_url,
            enabled_tools=enabled_tools,
        )
        token = server.set_http_request_context(context)
        upload_token = set_upload_credential_id(record.id if record else None)
        started_at = time.perf_counter()
        capture = _RequestResponseCapture(receive, send)
        error: str | None = None
        try:
            await self.mcp_app(_mcp_scope(scope), capture.receive, capture.send)
        except Exception as exc:
            error = exc.__class__.__name__
            raise
        finally:
            reset_upload_credential_id(upload_token)
            server.reset_http_request_context(token)
            if record is not None:
                try:
                    duration_ms = (time.perf_counter() - started_at) * 1000
                    self.store.append_operation_log(
                        record.id,
                        operation=_operation_from_request_body(capture.request_body),
                        http_method=scope["method"],
                        status_code=capture.status_code,
                        ok=_operation_ok(capture.status_code, capture.response_body, error),
                        error=error or _operation_error(capture.status_code, capture.response_body),
                        duration_ms=duration_ms,
                    )
                except Exception:
                    logging.getLogger("outlook_mcp").warning("failed to append operation log", exc_info=True)


_owned_store: CredentialStore | None = None


def _provision_single_user_endpoint(store: CredentialStore, base_config: RuntimeConfig) -> None:
    """Provision the single-user HTTP endpoint from OUTLOOK_MCP_BEARER_TOKEN or a generated token."""
    env_token = os.environ.get("OUTLOOK_MCP_BEARER_TOKEN", "").strip()
    if env_token:
        # An invalid env token is a configuration error; fail fast instead of provisioning garbage.
        validate_secret(env_token)
        token = env_token
        generated = False
    else:
        token = secrets.token_urlsafe(24)
        generated = True
    try:
        store.get_or_create_endpoint("graph", base_config.email, token)
    except Exception:
        logging.getLogger("outlook_mcp").warning(
            "failed to provision HTTP endpoint credential for %s", base_config.email, exc_info=True
        )
        return
    if generated:
        host = os.environ.get("OUTLOOK_MCP_HTTP_HOST", "127.0.0.1")
        port = os.environ.get("OUTLOOK_MCP_HTTP_PORT", "8000")
        logging.getLogger("outlook_mcp").info(
            "no OUTLOOK_MCP_BEARER_TOKEN set; generated a random token. Setup URL: http://%s:%s/graph/%s/%s",
            host,
            port,
            base_config.email,
            token,
        )


def create_app(
    db_path: str | Path | None = None,
    *,
    store: CredentialStore | None = None,
    base_config: RuntimeConfig | None = None,
) -> EndpointScopedMCPApp:
    global _owned_store
    credential_store = store if store is not None else CredentialStore(db_path or os.environ.get("OUTLOOK_MCP_DB") or DEFAULT_DB_PATH)
    previous_owned = _owned_store
    _owned_store = credential_store if store is None else None
    if previous_owned is not None and previous_owned is not credential_store:
        previous_owned.close()
    resolved_config = base_config if base_config is not None else load_config()
    _provision_single_user_endpoint(credential_store, resolved_config)
    server.configure_http_runtime(credential_store, base_config=resolved_config)
    _MAIL_RESPONSE_CACHE.clear()
    _AUTH_FAILURE_BUCKETS.clear()
    return EndpointScopedMCPApp(credential_store, resolved_config)


def main() -> None:
    host = os.environ.get("OUTLOOK_MCP_HTTP_HOST", "127.0.0.1")
    port = int(os.environ.get("OUTLOOK_MCP_HTTP_PORT", "8000"))
    db_path = os.environ.get("OUTLOOK_MCP_DB") or None
    app = create_app(db_path)
    uvicorn.run(
        app,
        host=host,
        port=port,
        proxy_headers=_trust_proxy_headers(),
        forwarded_allow_ips=os.environ.get("FORWARDED_ALLOW_IPS", "127.0.0.1"),
        access_log=False,
    )


def _html_page(title: str, heading: str, body: str, *, status_code: int = 200) -> HTMLResponse:
    content = f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <meta name="referrer" content="no-referrer">
      <title>{html.escape(title)}</title>
      <style>{PICO_CSS}</style>
      <style>{HTML_PAGE_CSS}</style>
    </head>
    <body>
      <main>
        <h1>{html.escape(heading)}</h1>
        {body}
      </main>
      {MCP_INSTALL_SCRIPT}
    </body>
    </html>
    """
    return HTMLResponse(content, status_code=status_code)


def _auth_page(
    method: str,
    account: str,
    record: CredentialRecord,
    auth_url: str,
    *,
    mode: str = "full",
    tools: str | None = None,
) -> HTMLResponse:
    status = server.http_auth_status(method, account, record.id, auth_url)
    escaped_account = html.escape(account)
    escaped_method = html.escape(method)
    install_help = _mcp_install_help_html(_mcp_url_from_auth_url(auth_url), _secret_from_auth_url(auth_url), mode=mode, tools=tools)
    auth_actions = _auth_action_forms(auth_url)
    if status["ok"]:
        mail_url = html.escape(_same_origin_path(auth_url, "mail"), quote=True)
        calendar_url = html.escape(_same_origin_path(auth_url, "calendar"), quote=True)
        body = f"""
        <p><strong>method</strong>: {escaped_method}</p>
        <p><strong>account</strong>: {escaped_account}</p>
        <p>This endpoint is ready for MCP calls.</p>
        <p><a href="{mail_url}">打开 Mail WebUI</a> · <a href="{calendar_url}">打开 Calendar WebUI</a></p>
        {auth_actions}
        {install_help}
        """
        return _html_page("Outlook MCP Authorized", "Outlook MCP Authorized", body)

    if status["error"] == "auth_required":
        data = status.get("data", {})
        verification_uri = html.escape(str(data.get("verification_uri") or ""))
        verification_uri_complete = html.escape(str(data.get("verification_uri_complete") or ""))
        user_code = html.escape(str(data.get("user_code") or ""))
        link = verification_uri_complete or verification_uri
        link_html = f'<p><a href="{link}" target="_blank" rel="noreferrer">Open Microsoft login</a></p>' if link else ""
        code_html = f"<p><strong>User code</strong>: <code>{user_code}</code></p>" if user_code else ""
        body = f"""
        <p><strong>method</strong>: {escaped_method}</p>
        <p><strong>account</strong>: {escaped_account}</p>
        {link_html}
        {code_html}
        <p>After finishing Microsoft login, refresh this page.</p>
        {auth_actions}
        {install_help}
        """
        return _html_page("Authorize Outlook MCP", "Authorize Outlook MCP", body)

    body = f"""
    <p><strong>method</strong>: {escaped_method}</p>
    <p><strong>account</strong>: {escaped_account}</p>
    <p><strong>{html.escape(status["error"])}</strong>: {html.escape(status.get("message", ""))}</p>
    {install_help}
    """
    return _html_page("Outlook MCP Endpoint", "Outlook MCP Endpoint", body, status_code=400)


def _auth_cleared_page(method: str, account: str, auth_url: str) -> HTMLResponse:
    escaped_account = html.escape(account)
    escaped_method = html.escape(method)
    body = f"""
    <p><strong>method</strong>: {escaped_method}</p>
    <p><strong>account</strong>: {escaped_account}</p>
    <p>Authorization for this endpoint has been cleared.</p>
    {_auth_action_forms(auth_url, include_cancel=False)}
    {_mcp_install_help_html(_mcp_url_from_auth_url(auth_url), _secret_from_auth_url(auth_url))}
    """
    return _html_page("Outlook MCP Authorization Cleared", "Outlook MCP Authorization Cleared", body)


def _auth_action_forms(auth_url: str, *, include_cancel: bool = True) -> str:
    cancel_url = html.escape(_same_origin_path(auth_url, "auth/cancel"))
    reauthorize_url = html.escape(_same_origin_path(auth_url, "auth/reauthorize"))
    cancel_form = (
        f"""
        <form method="post" action="{cancel_url}">
          <input type="hidden" name="csrf_token" value="{html.escape(_make_csrf_token("cancel", auth_url))}">
          <button type="submit">取消授权</button>
        </form>
        """
        if include_cancel
        else ""
    )
    return f"""
    <h2>Authorization actions</h2>
    {cancel_form}
    <form method="post" action="{reauthorize_url}">
      <input type="hidden" name="csrf_token" value="{html.escape(_make_csrf_token("reauthorize", auth_url))}">
      <button type="submit">重新授权</button>
    </form>
    """


def _same_origin_path(endpoint_url: str, suffix: str = "") -> str:
    path = urlsplit(endpoint_url).path or "/"
    if not suffix:
        return path
    return f"{path.rstrip('/')}/{suffix.lstrip('/')}"


def _secret_from_auth_url(auth_url: str) -> str:
    path = urlsplit(auth_url).path.strip("/")
    secret = path.rsplit("/", 1)[-1] if path else ""
    validate_secret(secret)
    return secret


def _mcp_url_from_auth_url(auth_url: str) -> str:
    split = urlsplit(auth_url)
    parts = split.path.strip("/").split("/")
    if len(parts) < 3:
        return auth_url
    mcp_path = "/" + "/".join(parts[:-1])
    return split._replace(path=mcp_path, query="", fragment="").geturl()


def _auth_action_csrf_error_page(method: str, account: str, auth_url: str) -> HTMLResponse:
    body = f"""
    <p><strong>method</strong>: {html.escape(method)}</p>
    <p><strong>account</strong>: {html.escape(account)}</p>
    <p>Invalid or expired CSRF token. Please go back and re-submit.</p>
    <p><a href="{html.escape(_same_origin_path(auth_url))}">Back to authorization page</a></p>
    """
    return _html_page("Outlook MCP CSRF Error", "csrf_error", body, status_code=403)


async def _api_route(
    action: str,
    scope: dict[str, Any],
    receive: Any,
    send: Any,
    method: str,
    account: str,
    store: CredentialStore,
    base_config: RuntimeConfig | None = None,
) -> None:
    api_path = action[4:]  # strip "api/"
    if scope["method"] != "POST":
        await PlainTextResponse("Method Not Allowed", status_code=405)(scope, receive, send)
        return
    content_type = _headers(scope).get("content-type", "").split(";", 1)[0].strip().lower()
    if content_type != "application/json":
        await JSONResponse(
            {
                "ok": False,
                "error": "unsupported_media_type",
                "message": "API requests require Content-Type: application/json.",
            },
            status_code=415,
        )(scope, receive, send)
        return
    try:
        body = await _read_request_body(receive)
    except ToolValidationError as exc:
        await JSONResponse({"ok": False, "error": "validation_error", "message": str(exc)}, status_code=400)(scope, receive, send)
        return
    params: dict[str, Any] = {}
    if body:
        try:
            parsed = json.loads(body)
        except json.JSONDecodeError:
            await JSONResponse({"ok": False, "error": "invalid_request", "message": "Invalid JSON"}, status_code=400)(scope, receive, send)
            return
        if not isinstance(parsed, dict):
            await JSONResponse(
                {"ok": False, "error": "invalid_request", "message": "JSON request body must be an object."},
                status_code=400,
            )(scope, receive, send)
            return
        params = parsed

    try:
        enabled_tools = _tools_allowlist(scope)
        secret = _bearer_token(scope)
    except ToolValidationError as exc:
        await JSONResponse({"ok": False, "error": "validation_error", "message": str(exc)}, status_code=400)(scope, receive, send)
        return
    except ValueError as exc:
        await JSONResponse({"ok": False, "error": "invalid_secret", "message": str(exc)}, status_code=400)(scope, receive, send)
        return
    if secret is None:
        await _missing_bearer_response(scope)(scope, receive, send)
        return
    auth_url = _auth_url(scope, method, account, secret)
    record = await asyncio.to_thread(_lookup_single_account_endpoint, store, base_config, method, account, secret)
    if record is None and not _check_auth_rate(_client_source(scope)):
        await _auth_rate_limited_response(scope)(scope, receive, send)
        return
    context = server.HttpRequestContext(
        method=method, account=account, secret_hash=hash_secret(secret),
        credential_id=record.id if record else None, auth_url=auth_url, enabled_tools=enabled_tools,
    )
    draft_api = is_calendar_draft_api(api_path)
    if draft_api and record is None:
        await JSONResponse(
            {"ok": False, "error": "auth_required", "message": "Credential is not registered."},
            status_code=403,
        )(scope, receive, send)
        return

    client: GraphClient | None = None
    if not draft_api or api_path == "calendar/drafts/submit":
        resolved_client = await asyncio.to_thread(server._get_http_client, context)
        if not isinstance(resolved_client, GraphClient):
            message = getattr(resolved_client, "message", "Not authorized")
            await JSONResponse(
                {"ok": False, "error": "auth_required", "message": str(message)},
                status_code=403,
            )(scope, receive, send)
            return
        client = resolved_client

    cache_namespace = record.id if record is not None else None
    cache_ttl = _MAIL_API_CACHE_TTLS.get(api_path)
    cache_params = {
        "params": params,
        "enabled_tools": sorted(enabled_tools) if enabled_tools is not None else None,
    }
    result = (
        _MAIL_RESPONSE_CACHE.get(cache_namespace, api_path, cache_params)
        if cache_namespace is not None and cache_ttl is not None
        else None
    )
    if result is None:
        upload_token = set_upload_credential_id(cache_namespace)
        try:
            if draft_api:
                result = await asyncio.to_thread(
                    dispatch_calendar_draft_api,
                    api_path,
                    store,
                    record.id,
                    params,
                    client=client,
                    enabled_tools=enabled_tools,
                )
            else:
                result = await asyncio.to_thread(
                    _dispatch_api,
                    api_path,
                    client,
                    params,
                    enabled_tools=enabled_tools,
                )
        finally:
            reset_upload_credential_id(upload_token)
        if cache_namespace is not None and result.get("ok", True):
            if cache_ttl is not None:
                _MAIL_RESPONSE_CACHE.put(
                    cache_namespace,
                    api_path,
                    cache_params,
                    result,
                    ttl_seconds=cache_ttl,
                )
            elif api_path in _MAIL_API_WRITE_PATHS:
                _MAIL_RESPONSE_CACHE.invalidate(cache_namespace)
    status_code = _result_status_code(result)
    await JSONResponse(result, status_code=status_code)(scope, receive, send)


_API_ERROR_STATUS_CODES = {
    "validation_error": 400,
    "tool_disabled": 400,
    "auth_required": 401,
    "not_found": 404,
}


def _result_status_code(result: dict[str, Any]) -> int:
    explicit = result.pop("_status_code", None)
    if explicit is not None:
        return int(explicit)
    if result.get("ok", True):
        return 200
    return _API_ERROR_STATUS_CODES.get(str(result.get("error") or ""), 502)


def _dispatch_api(api_path: str, client: Any, params: dict[str, Any], *, enabled_tools: frozenset[str] | None = None) -> dict[str, Any]:
    try:
        return _do_dispatch(api_path, client, params, enabled_tools=enabled_tools)
    except Exception as exc:
        def failed() -> Any:
            raise exc
        return run_safely(failed)


def _do_dispatch(api_path: str, client: Any, params: dict[str, Any], *, enabled_tools: frozenset[str] | None = None) -> dict[str, Any]:
    from .tools import get_api_dispatch

    tool_func = get_api_dispatch().get(api_path)
    if tool_func is None:
        return {"ok": False, "error": "not_found", "message": f"Unknown API path: {api_path}"}
    tool_name = getattr(tool_func, "__name__", None) or api_path.replace("/", "_")
    if enabled_tools is not None and tool_name not in enabled_tools:
        return {
            "ok": False,
            "error": "tool_disabled",
            "message": f"MCP tool is disabled by current MCP tool selection: {tool_name}",
        }
    try:
        inspect.signature(tool_func).bind(client, **params)
    except TypeError as exc:
        return {"ok": False, "error": "validation_error", "message": f"invalid parameters: {exc}"}
    return tool_func(client, **params)


_MAX_AUTH_BODY_BYTES = 65536
_MAX_API_BODY_BYTES = 1_048_576
_MAX_UPLOAD_MULTIPART_OVERHEAD_BYTES = 256 * 1024


async def _read_request_body(receive: Any, max_bytes: int = _MAX_API_BODY_BYTES) -> bytes:
    chunks: list[bytes] = []
    size = 0
    more_body = True
    while more_body:
        msg = await receive()
        if msg.get("type") == "http.request":
            chunk = msg.get("body", b"")
            size += len(chunk)
            if size > max_bytes:
                raise RequestBodyTooLargeError("request body too large")
            chunks.append(chunk)
            more_body = msg.get("more_body", False)
        elif msg.get("type") == "http.disconnect":
            raise ToolValidationError("request disconnected")
    return b"".join(chunks)


def _body_receive(body: bytes) -> Any:
    sent = False

    async def receive() -> dict[str, Any]:
        nonlocal sent
        if sent:
            return {"type": "http.disconnect"}
        sent = True
        return {"type": "http.request", "body": body, "more_body": False}

    return receive


def _make_endpoint_route(method: str, account: str, secret: str | None = None, action: str | None = None) -> EndpointRoute | None:
    if not method or not account:
        return None
    method = method.strip().lower()
    if method not in ENDPOINT_METHODS:
        return None
    account = account.strip()
    if not EMAIL_RE.match(account):
        return None
    if secret is not None:
        validate_secret(secret)
    return EndpointRoute(method, account, secret, action)


def _parse_endpoint(path: str) -> EndpointRoute | None:
    parts = [unquote(part) for part in path.strip("/").split("/")]
    if len(parts) >= 3 and parts[2] == "api":
        method, account, _, *api_parts = parts
        return _make_endpoint_route(method, account, None, f"api/{'/'.join(api_parts)}")
    if len(parts) == 3 and parts[2] == "uploads":
        method, account, _ = parts
        return _make_endpoint_route(method, account, None, "uploads")
    if len(parts) >= 4 and parts[3] == "api":
        method, account, secret, _, *api_parts = parts
        return _make_endpoint_route(method, account, secret, f"api/{'/'.join(api_parts)}")
    if len(parts) == 4 and parts[3] == "log":
        method, account, secret, _ = parts
        return _make_endpoint_route(method, account, secret, "log")
    if len(parts) == 4 and parts[3] == "calendar":
        method, account, secret, _ = parts
        return _make_endpoint_route(method, account, secret, "calendar")
    if len(parts) == 4 and parts[3] == "mail":
        method, account, secret, _ = parts
        return _make_endpoint_route(method, account, secret, "mail")
    if len(parts) == 5 and parts[3] == "auth" and parts[4] in {"cancel", "reauthorize"}:
        method, account, secret, auth, action = parts
        return _make_endpoint_route(method, account, secret, f"{auth}/{action}")
    if len(parts) == 2 and not any(part == "" for part in parts):
        method, account = parts
        return _make_endpoint_route(method, account)
    if len(parts) != 3 or any(part == "" for part in parts):
        return None
    method, account, secret = parts
    return _make_endpoint_route(method, account, secret)


def _invalid_secret_response(scope: dict[str, Any], message: str) -> HTMLResponse | JSONResponse:
    if scope["method"] == "GET" and not _wants_mcp_stream(scope):
        body = f"""
        <p>{html.escape(message)}</p>
        {_endpoint_help_html(_external_base_url(scope))}
        """
        return _html_page("Invalid Outlook MCP Secret", "invalid_secret", body, status_code=400)
    return JSONResponse({"ok": False, "error": "invalid_secret", "message": message}, status_code=400)


def _invalid_endpoint_response(scope: dict[str, Any]) -> HTMLResponse | JSONResponse:
    base_url = _external_base_url(scope)
    help_payload = _endpoint_help_payload(base_url)
    if scope["method"] == "GET" and not _wants_mcp_stream(scope):
        body = f"""
        {_endpoint_help_html(base_url)}
        """
        return _html_page("Outlook MCP Endpoint Help", "invalid_endpoint", body, status_code=404)
    return JSONResponse({"ok": False, "error": "invalid_endpoint", **help_payload}, status_code=404)


def _unregistered_credential_page(method: str, account: str, secret: str) -> HTMLResponse:
    linux_command = " ".join(
        shlex.quote(value)
        for value in (
            "uv",
            "run",
            "python3",
            "-m",
            "outlook_mcp.credential_store",
            "register",
            method,
            account,
            secret,
        )
    )
    windows_args = (method, account, secret)
    windows_command = (
        "uv run python -m outlook_mcp.credential_store register "
        + " ".join("'" + value.replace("'", "''") + "'" for value in windows_args)
    )
    body = f"""
    <p>This HTTP credential is not registered. Browser GET requests never create credentials.</p>
    <p><strong>Linux</strong></p>
    <pre><code>{html.escape(linux_command)}</code></pre>
    <p><strong>Windows PowerShell</strong></p>
    <pre><code>{html.escape(windows_command)}</code></pre>
    <p>Run one command on the server, then reload this page to start authorization.</p>
    """
    return _html_page(
        "Outlook MCP Credential Not Registered",
        "credential_not_registered",
        body,
        status_code=404,
    )


def _missing_bearer_response(scope: dict[str, Any]) -> JSONResponse:
    return JSONResponse(
        {
            "ok": False,
            "error": "auth_required",
            "message": "HTTP MCP requests must include Authorization: Bearer <token>.",
        },
        status_code=401,
    )


def _bearer_token(scope: dict[str, Any]) -> str | None:
    header = _headers(scope).get("authorization", "").strip()
    scheme, _, value = header.partition(" ")
    if scheme.lower() != "bearer" or not value.strip():
        return None
    token = value.strip()
    validate_secret(token)
    return token


def _tools_allowlist(scope: dict[str, Any]) -> frozenset[str] | None:
    headers = _headers(scope)
    raw_names = _split_tools(headers.get("x-outlook-mcp-tools", ""))
    if not raw_names:
        mode = headers.get("x-outlook-mcp-mode", "").strip().lower()
        if not mode:
            return None
        return _tools_for_mode(mode, source="X-Outlook-MCP-Mode")
    _validate_tool_names(raw_names, source="X-Outlook-MCP-Tools")
    return frozenset(raw_names)


def _tools_for_mode(mode: str, *, source: str) -> frozenset[str] | None:
    normalized_mode = mode.strip().lower() or "full"
    preset = TOOL_MODE_BY_NAME.get(normalized_mode)
    if preset is None:
        raise ToolValidationError(f"Unknown MCP tools mode: {normalized_mode}")
    tools = str(preset["tools"] or "")
    names = _split_tools(tools)
    if not names:
        return frozenset() if normalized_mode == "custom" else None
    _validate_tool_names(names, source=source)
    return frozenset(names)


def _custom_seed_tools_for_mode(mode: str) -> str:
    normalized_mode = mode.strip().lower() or "full"
    preset = TOOL_MODE_BY_NAME.get(normalized_mode)
    if preset is None:
        raise ToolValidationError(f"Unknown MCP tools mode: {normalized_mode}")
    if normalized_mode == "custom":
        return ""
    if normalized_mode != "full":
        return str(preset["tools"] or "")

    names: list[str] = []
    seen: set[str] = set()
    has_account_info = False
    for item in TOOL_MODE_PRESETS:
        item_mode = str(item["mode"])
        if item_mode in {"full", "custom"}:
            continue
        for name in _split_tools(str(item["tools"] or "")):
            if name == "account_info":
                has_account_info = True
                continue
            if name not in seen:
                names.append(name)
                seen.add(name)
    if has_account_info:
        names.append("account_info")
    return ",".join(names)


def _install_mode_and_tools(scope: dict[str, Any]) -> tuple[str, str | None]:
    query = _query_params(scope)
    tools_value = _query_value(query, "tools", "")
    names = _split_tools(tools_value)
    if not names:
        mode = _query_value(query, "mode", "full").strip().lower() or "full"
        _tools_for_mode(mode, source=f"mode {mode}")
        return mode, None
    _validate_tool_names(names, source="tools")
    return "custom", ",".join(names)


def _split_tools(value: str) -> list[str]:
    return [part.strip() for part in value.split(",") if part.strip()]


def _validate_tool_names(names: list[str], *, source: str) -> None:
    available = server.available_tool_names()
    unknown = sorted(set(names) - available)
    if unknown:
        raise ToolValidationError(f"Unknown MCP tools in {source}: {', '.join(unknown)}")


async def _upload_response(
    store: CredentialStore,
    base_config: RuntimeConfig | None,
    scope: dict[str, Any],
    receive: Any,
    method: str,
    account: str,
) -> JSONResponse | PlainTextResponse:
    if scope["method"] != "POST":
        return PlainTextResponse("Method Not Allowed", status_code=405)
    content_type = _headers(scope).get("content-type", "").split(";", 1)[0].strip().lower()
    if content_type != "multipart/form-data":
        return JSONResponse(
            {
                "ok": False,
                "error": "unsupported_media_type",
                "message": "Uploads require Content-Type: multipart/form-data.",
            },
            status_code=415,
        )
    try:
        secret = _bearer_token(scope)
    except ValueError as exc:
        return JSONResponse({"ok": False, "error": "invalid_secret", "message": str(exc)}, status_code=400)
    if secret is None:
        return _missing_bearer_response(scope)
    record = await asyncio.to_thread(_lookup_single_account_endpoint, store, base_config, method, account, secret)
    if record is None:
        return JSONResponse(
            {
                "ok": False,
                "error": "auth_required",
                "message": "Open the endpoint auth page before uploading files.",
            },
            status_code=403,
        )
    cache_json = await asyncio.to_thread(store.load_msal_cache, record.id)
    if not cache_json or not _has_msal_tokens(cache_json):
        return JSONResponse(
            {
                "ok": False,
                "error": "auth_required",
                "message": "Complete Microsoft OAuth authorization for this endpoint before uploading files.",
            },
            status_code=403,
        )

    if not _check_upload_rate(record.id):
        return JSONResponse(
            {"ok": False, "error": "rate_limited", "message": "Too many uploads. Please wait before uploading again."},
            status_code=429,
        )

    max_request_bytes = upload_max_bytes() + _MAX_UPLOAD_MULTIPART_OVERHEAD_BYTES
    content_length = _headers(scope).get("content-length", "").strip()
    if content_length:
        try:
            declared_size = int(content_length)
        except ValueError:
            return JSONResponse(
                {"ok": False, "error": "invalid_request", "message": "Content-Length must be an integer."},
                status_code=400,
            )
        if declared_size < 0 or declared_size > max_request_bytes:
            return JSONResponse(
                {"ok": False, "error": "request_too_large", "message": "Upload request body is too large."},
                status_code=413,
            )

    file: UploadFile | None = None
    upload_token = set_upload_credential_id(record.id)
    try:
        body = await _read_request_body(receive, max_bytes=max_request_bytes)
        request = Request(scope, _body_receive(body))
        form = await request.form(max_files=1, max_fields=8, max_part_size=upload_max_bytes())
        uploaded_files = [
            (name, value)
            for name, value in form.multi_items()
            if isinstance(value, UploadFile)
        ]
        if len(uploaded_files) != 1 or uploaded_files[0][0] != "file":
            raise ToolValidationError("multipart request must contain exactly one file field")
        file = uploaded_files[0][1]
        if not isinstance(file, UploadFile):
            raise ToolValidationError("multipart field file is required")
        content = await file.read(upload_max_bytes() + 1)
        await asyncio.to_thread(cleanup_expired_uploads)
        uploaded = await asyncio.to_thread(
            save_uploaded_file,
            credential_id=record.id,
            filename=file.filename or "",
            mime_type=file.content_type or "",
            content=content,
        )
    except RequestBodyTooLargeError as exc:
        return JSONResponse({"ok": False, "error": "request_too_large", "message": str(exc)}, status_code=413)
    except (ToolValidationError, MultiPartException) as exc:
        return JSONResponse({"ok": False, "error": "validation_error", "message": str(exc)}, status_code=400)
    except Exception:
        logging.getLogger("outlook_mcp").exception("upload handler failed")
        return JSONResponse(
            {"ok": False, "error": "internal_error", "message": "Upload processing failed"},
            status_code=500,
        )
    finally:
        reset_upload_credential_id(upload_token)
        if isinstance(file, UploadFile):
            await file.close()

    return JSONResponse(
        {
            "ok": True,
            "relative_path": uploaded.relative_path,
            "mime_type": uploaded.mime_type,
            "size": uploaded.size,
            "sha256": uploaded.sha256,
            "width": uploaded.width,
            "height": uploaded.height,
        }
    )


def _calendar_bootstrap(scope: dict[str, Any], method: str, account: str, secret: str) -> dict[str, Any]:
    auth_url = _auth_url(scope, method, account, secret)
    endpoint_url = _mcp_endpoint_url(scope, method, account)
    return {
        "provider": method,
        "method": method,
        "account": account,
        "apiBase": _same_origin_path(endpoint_url, "api"),
        "authPath": _same_origin_path(auth_url),
        "calendarTimezone": DEFAULT_CALENDAR_TIMEZONE,
        "capabilities": CALENDAR_CAPABILITIES,
    }


def _calendar_page(scope: dict[str, Any], method: str, account: str, secret: str) -> HTMLResponse:
    html_content = _read_static_text("calendar.html")
    html_content = html_content.replace("__ACCOUNT__", html.escape(account))
    html_content = html_content.replace("__CALENDAR_CSS__", CALENDAR_CSS)
    html_content = html_content.replace("__PICO_CSS__", PICO_CSS)
    html_content = html_content.replace("__APP_BOOTSTRAP__", _json_script_value(_calendar_bootstrap(scope, method, account, secret)))
    html_content = html_content.replace("__FULLCALENDAR_JS_URL__", html.escape(_static_path(scope, FULLCALENDAR_JS), quote=True))
    html_content = html_content.replace("__FULLCALENDAR_LOCALE_JS_URL__", html.escape(_static_path(scope, FULLCALENDAR_LOCALE_JS), quote=True))
    html_content = html_content.replace("__TINYMCE_JS_URL__", html.escape(_static_path(scope, TINYMCE_JS), quote=True))
    html_content = html_content.replace("__CALENDAR_JS_URL__", html.escape(_static_path(scope, "calendar.js"), quote=True))
    html_content = html_content.replace("__APP_ICON_URL__", html.escape(_static_path(scope, "outlook-mcp-icon.svg"), quote=True))
    return HTMLResponse(html_content)


def _mail_bootstrap(scope: dict[str, Any], method: str, account: str, secret: str) -> dict[str, Any]:
    auth_url = _auth_url(scope, method, account, secret)
    endpoint_url = _mcp_endpoint_url(scope, method, account)
    return {
        "provider": method,
        "method": method,
        "account": account,
        "apiBase": _same_origin_path(endpoint_url, "api"),
        "authPath": _same_origin_path(auth_url),
        "pageSize": 20,
        "bodyMaxChars": 200_000,
    }


def _mail_page(scope: dict[str, Any], method: str, account: str, secret: str) -> HTMLResponse:
    html_content = _read_static_text("mail.html")
    html_content = html_content.replace("__ACCOUNT__", html.escape(account))
    html_content = html_content.replace("__MAIL_CSS__", MAIL_CSS)
    html_content = html_content.replace("__PICO_CSS__", PICO_CSS)
    html_content = html_content.replace("__APP_BOOTSTRAP__", _json_script_value(_mail_bootstrap(scope, method, account, secret)))
    html_content = html_content.replace("__TINYMCE_JS_URL__", html.escape(_static_path(scope, TINYMCE_JS), quote=True))
    html_content = html_content.replace("__MAIL_JS_URL__", html.escape(_versioned_static_path(scope, "mail.js"), quote=True))
    html_content = html_content.replace("__APP_ICON_URL__", html.escape(_static_path(scope, "outlook-mcp-icon.svg"), quote=True))
    return HTMLResponse(html_content)


def _calendar_unsupported_page(method: str, account: str, auth_url: str) -> HTMLResponse:
    body = f"""
    <p><strong>method</strong>: {html.escape(method)}</p>
    <p><strong>account</strong>: {html.escape(account)}</p>
    <p>The visual calendar editor is currently available only for Graph HTTP endpoints.</p>
    <p><a href="{html.escape(_same_origin_path(auth_url))}">Back to endpoint page</a></p>
    """
    return _html_page("Outlook Calendar Unsupported", "Outlook Calendar Unsupported", body, status_code=400)


def _mail_unsupported_page(method: str, account: str, auth_url: str) -> HTMLResponse:
    body = f"""
    <p><strong>method</strong>: {html.escape(method)}</p>
    <p><strong>account</strong>: {html.escape(account)}</p>
    <p>The visual mail client is currently available only for Graph HTTP endpoints.</p>
    <p><a href="{html.escape(_same_origin_path(auth_url))}">Back to endpoint page</a></p>
    """
    return _html_page("Outlook Mail Unsupported", "Outlook Mail Unsupported", body, status_code=400)


def _log_page(
    store: CredentialStore,
    scope: dict[str, Any],
    method: str,
    account: str,
    record: CredentialRecord,
) -> HTMLResponse:
    selected_range, start_at, end_at, from_value, to_value, warning = _log_filter(scope)
    store.flush_logs()
    logs = store.list_operation_logs(record.id, start_at=start_at, end_at=end_at, limit=500)
    body = f"""
    <p><strong>method</strong>: {html.escape(method)}</p>
    <p><strong>account</strong>: {html.escape(account)}</p>
    {_log_filter_form(selected_range, from_value, to_value)}
    {f'<p><strong>filter_error</strong>: {html.escape(warning)}</p>' if warning else ''}
    {_log_table(logs)}
    """
    return _html_page("Outlook MCP Operation Logs", "Outlook MCP Operation Logs", body)


def _log_filter(scope: dict[str, Any]) -> tuple[str, float | None, float | None, str, str, str | None]:
    query = _query_params(scope)
    selected = _query_value(query, "range", "all")
    if selected not in LOG_RANGES:
        selected = "all"
    from_value = _query_value(query, "from", "")
    to_value = _query_value(query, "to", "")
    now = time.time()
    if selected == "24h":
        return selected, now - 24 * 60 * 60, None, from_value, to_value, None
    if selected == "7d":
        return selected, now - 7 * 24 * 60 * 60, None, from_value, to_value, None
    if selected == "30d":
        return selected, now - 30 * 24 * 60 * 60, None, from_value, to_value, None
    if selected == "custom":
        try:
            start_at = _date_start_timestamp(from_value) if from_value else None
            end_at = _date_start_timestamp(to_value, add_days=1) if to_value else None
        except ValueError:
            return "all", None, None, from_value, to_value, "custom dates must use YYYY-MM-DD"
        return selected, start_at, end_at, from_value, to_value, None
    return "all", None, None, from_value, to_value, None


def _log_filter_form(selected_range: str, from_value: str, to_value: str) -> str:
    options = "".join(
        f'<option value="{value}"{" selected" if value == selected_range else ""}>{label}</option>'
        for value, label in (
            ("all", "All"),
            ("24h", "Last 24 hours"),
            ("7d", "Last 7 days"),
            ("30d", "Last 30 days"),
            ("custom", "Custom"),
        )
    )
    return f"""
    <form method="get">
      <label>Range
        <select name="range">{options}</select>
      </label>
      <label>From
        <input type="date" name="from" value="{html.escape(from_value)}">
      </label>
      <label>To
        <input type="date" name="to" value="{html.escape(to_value)}">
      </label>
      <button type="submit">Apply</button>
    </form>
    """


def _log_table(logs: list[dict[str, Any]]) -> str:
    if not logs:
        return "<p>No operation logs found.</p>"
    rows = []
    for item in logs:
        rows.append(
            "<tr>"
            f"<td>{html.escape(_utc_text(item['created_at']))}</td>"
            f"<td>{html.escape(item['http_method'])}</td>"
            f"<td>{html.escape(item['operation'])}</td>"
            f"<td>{'' if item['status_code'] is None else int(item['status_code'])}</td>"
            f"<td>{'ok' if item['ok'] else 'failed'}</td>"
            f"<td>{html.escape(str(item['error'] or ''))}</td>"
            f"<td>{item['duration_ms']:.1f}</td>"
            "</tr>"
        )
    return (
        "<table>"
        "<thead><tr><th>Time (UTC)</th><th>HTTP</th><th>Operation</th><th>Status</th>"
        "<th>Result</th><th>Error</th><th>Duration ms</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody>"
        "</table>"
    )


def _query_params(scope: dict[str, Any]) -> dict[str, list[str]]:
    raw = scope.get("query_string", b"")
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")
    return parse_qs(str(raw), keep_blank_values=True)


def _query_value(query: dict[str, list[str]], key: str, default: str) -> str:
    values = query.get(key)
    return values[0] if values else default


def _date_start_timestamp(value: str, *, add_days: int = 0) -> float:
    parsed = datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    if add_days:
        parsed += timedelta(days=add_days)
    return parsed.timestamp()


def _utc_text(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp, tz=timezone.utc).isoformat()


def _wants_mcp_stream(scope: dict[str, Any]) -> bool:
    headers = _headers(scope)
    accept = headers.get("accept", "")
    if "text/event-stream" in accept:
        return True
    if "text/html" in accept:
        return False
    if "application/json" in accept:
        return not headers.get("sec-fetch-dest")
    return False


def _mcp_scope(scope: dict[str, Any]) -> dict[str, Any]:
    rewritten = dict(scope)
    rewritten["path"] = "/"
    rewritten["raw_path"] = b"/"
    return rewritten


class _RequestResponseCapture:
    _MAX_BODY_BYTES = 65536

    def __init__(self, receive: Any, send: Any):
        self._receive = receive
        self._send = send
        self.request_body = bytearray()
        self.response_body = bytearray()
        self.status_code: int | None = None

    async def receive(self) -> dict[str, Any]:
        message = await self._receive()
        if message.get("type") == "http.request" and message.get("body"):
            remaining = self._MAX_BODY_BYTES - len(self.request_body)
            if remaining > 0:
                self.request_body.extend(message["body"][:remaining])
        return message

    async def send(self, message: dict[str, Any]) -> None:
        if message.get("type") == "http.response.start":
            self.status_code = int(message.get("status") or 0)
        elif message.get("type") == "http.response.body" and message.get("body"):
            remaining = self._MAX_BODY_BYTES - len(self.response_body)
            if remaining > 0:
                self.response_body.extend(message["body"][:remaining])
        await self._send(message)


def _operation_from_request_body(body: bytes | bytearray) -> str:
    try:
        payload = json.loads(bytes(body).decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return "mcp_request"
    if isinstance(payload, list) and payload:
        payload = payload[0]
    if not isinstance(payload, dict):
        return "mcp_request"
    method = str(payload.get("method") or "")
    params = payload.get("params")
    if method == "tools/call" and isinstance(params, dict) and params.get("name"):
        return str(params["name"])
    return method or "mcp_request"


def _operation_ok(status_code: int | None, body: bytes | bytearray, error: str | None) -> bool:
    if error:
        return False
    if status_code is not None and status_code >= 400:
        return False
    payload = _json_body(body)
    if isinstance(payload, dict):
        if payload.get("ok") is False or "error" in payload:
            return False
        result = payload.get("result")
        if isinstance(result, dict):
            result_payload = _jsonrpc_text_payload(result)
            if isinstance(result_payload, dict) and result_payload.get("ok") is False:
                return False
    return True


def _operation_error(status_code: int | None, body: bytes | bytearray) -> str | None:
    payload = _json_body(body)
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, str):
            return error
        if isinstance(error, dict):
            return str(error.get("code") or error.get("message") or "jsonrpc_error")
        result = payload.get("result")
        if isinstance(result, dict):
            result_payload = _jsonrpc_text_payload(result)
            if isinstance(result_payload, dict) and result_payload.get("ok") is False:
                return str(result_payload.get("error") or "operation_failed")
    if status_code is not None and status_code >= 400:
        return f"http_{status_code}"
    return None


def _json_body(body: bytes | bytearray) -> Any:
    if not body:
        return None
    try:
        return json.loads(bytes(body).decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return None


def _jsonrpc_text_payload(result: dict[str, Any]) -> Any:
    content = result.get("content")
    if not isinstance(content, list):
        return None
    for item in content:
        if not isinstance(item, dict) or item.get("type") != "text":
            continue
        text = item.get("text")
        if not isinstance(text, str):
            continue
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            return None
    return None


def _has_msal_tokens(cache_json: str) -> bool:
    try:
        data = json.loads(cache_json)
    except (json.JSONDecodeError, TypeError):
        return False
    return "AccessToken" in data or "RefreshToken" in data


def _auth_url(scope: dict[str, Any], method: str, account: str, secret: str) -> str:
    return f"{_external_base_url(scope)}/{method}/{account}/{secret}"


def _mcp_endpoint_url(scope: dict[str, Any], method: str, account: str) -> str:
    return f"{_external_base_url(scope)}/{method}/{account}"


def _external_base_url(scope: dict[str, Any]) -> str:
    headers = _headers(scope)
    trust_proxy = _trust_proxy_headers()
    forwarded = _forwarded_values(headers.get("forwarded", "")) if trust_proxy else {}
    scheme = scope.get("scheme", "http")
    host = headers.get("host")
    if trust_proxy:
        scheme = forwarded.get("proto") or _first_header_value(headers.get("x-forwarded-proto")) or scheme
        host = forwarded.get("host") or _first_header_value(headers.get("x-forwarded-host")) or host
    if not host:
        server_info = scope.get("server") or ("127.0.0.1", 8000)
        host = f"{server_info[0]}:{server_info[1]}"
    port = _first_header_value(headers.get("x-forwarded-port")) if trust_proxy else ""
    if port and not _host_has_port(host) and not _is_default_port(scheme, port):
        host = f"{host}:{port}"
    forwarded_prefix = _first_header_value(headers.get("x-forwarded-prefix")) if trust_proxy else ""
    prefix = _normalized_prefix(forwarded_prefix or scope.get("root_path", ""))
    return f"{scheme}://{host}{prefix}"


def _trust_proxy_headers() -> bool:
    explicit = os.environ.get("OUTLOOK_MCP_TRUST_PROXY_HEADERS")
    if explicit is not None:
        return explicit.strip().lower() in {"1", "true", "yes", "on"}
    return bool(os.environ.get("FORWARDED_ALLOW_IPS", "").strip())


def _endpoint_help_payload(base_url: str) -> dict[str, Any]:
    example_url = f"{base_url}{DEFAULT_EXAMPLE_PATH}"
    return {
        "expected_path": EXPECTED_ENDPOINT_PATH,
        "example_path": DEFAULT_EXAMPLE_PATH,
        "example_auth_path": DEFAULT_EXAMPLE_AUTH_PATH,
        "example_url": example_url,
        "mcp_json": _mcp_config(example_url, DEFAULT_EXAMPLE_TOKEN),
        "agent_prompt": _agent_prompt(example_url, DEFAULT_EXAMPLE_TOKEN),
    }


def _endpoint_help_html(base_url: str) -> str:
    payload = _endpoint_help_payload(base_url)
    expected_path = html.escape(payload["expected_path"])
    example_path = html.escape(payload["example_path"])
    example_auth_path = html.escape(payload["example_auth_path"])
    example_url = html.escape(payload["example_url"])
    return f"""
    <h2>Correct URL format</h2>
    <p>Use <code>{expected_path}</code>.</p>
    <p>Example path: <code>{example_path}</code></p>
    <p>Browser auth page: <code>{example_auth_path}</code></p>
    <p>Example URL: <code>{example_url}</code></p>
    {_mcp_install_help_html(payload["example_url"], DEFAULT_EXAMPLE_TOKEN)}
    """


def _copy_panel_html(
    title: str,
    copy_name: str,
    content: str,
    *,
    block: bool = True,
    endpoint_url: bool = False,
    rewrite_url: bool = False,
) -> str:
    data_attrs = ""
    if endpoint_url:
        data_attrs += " data-mcp-endpoint-url"
    if rewrite_url:
        data_attrs += " data-mcp-rewrite-url"
    escaped_title = html.escape(title)
    escaped_copy_name = html.escape(copy_name, quote=True)
    copy_label = html.escape(f"复制 {title}", quote=True)
    if block:
        body = (
            f'<pre class="copy-panel__body">'
            f'<code data-copy-name="{escaped_copy_name}"{data_attrs}>{content}</code>'
            f"</pre>"
        )
    else:
        body = (
            f'<div class="copy-panel__inline">'
            f'<code data-copy-name="{escaped_copy_name}"{data_attrs}>{content}</code>'
            f"</div>"
        )
    return f"""
    <div class="copy-panel">
      <div class="copy-panel__header">
        <p class="copy-panel__title">{escaped_title}</p>
        <button type="button" class="secondary copy-button" data-copy-source="{escaped_copy_name}" aria-label="{copy_label}">复制</button>
      </div>
      {body}
    </div>
    """


def _mcp_install_help_html(endpoint_url: str, bearer_token: str, *, mode: str = "full", tools: str | None = None) -> str:
    mcp_json = html.escape(json.dumps(_mcp_config(endpoint_url, bearer_token, mode=mode, tools=tools), ensure_ascii=False, indent=2))
    agent_prompt = html.escape(_agent_prompt(endpoint_url, bearer_token, mode=mode, tools=tools))
    endpoint = html.escape(endpoint_url)
    endpoint_path = html.escape(_same_origin_path(endpoint_url), quote=True)
    endpoint_attr = html.escape(endpoint_url, quote=True)
    bearer_attr = html.escape(bearer_token, quote=True)
    server_attr = html.escape(MCP_SERVER_NAME, quote=True)
    return f"""
    <section class="mcp-install" data-mcp-install data-endpoint-url="{endpoint_attr}" data-endpoint-path="{endpoint_path}" data-bearer-token="{bearer_attr}" data-server-name="{server_attr}">
    <h2>Add this endpoint to MCP</h2>
    {_default_endpoint_warning_html(endpoint_url)}
    {_mode_selector_html(mode, tools)}
    <p>Manual values:</p>
    <ul>
      <li><strong>server name</strong>: <code>{MCP_SERVER_NAME}</code></li>
      <li><strong>type</strong>: <code>streamable-http</code></li>
      <li><strong>auth</strong>: <code>Authorization: Bearer ...</code></li>
    </ul>
    <div class="copy-grid">
      {_copy_panel_html("Endpoint URL", "endpoint-url", endpoint, block=False, endpoint_url=True, rewrite_url=True)}
      {_copy_panel_html("MCP JSON", "mcp-json", mcp_json, rewrite_url=True)}
      {_copy_panel_html("Agent prompt", "agent-prompt", agent_prompt, rewrite_url=True)}
    </div>
    </section>
    """


def _mode_selector_html(selected_mode: str, tools: str | None) -> str:
    buttons = []
    for preset in TOOL_MODE_PRESETS:
        mode = str(preset["mode"])
        label = str(preset["label"])
        pressed = "true" if mode == selected_mode else "false"
        custom_seed_tools = _custom_seed_tools_for_mode(mode)
        buttons.append(
            f'<button type="button" class="secondary mode-button" data-mcp-mode data-mode="{html.escape(mode, quote=True)}" '
            f'aria-pressed="{pressed}" data-custom-tools="{html.escape(custom_seed_tools, quote=True)}">{html.escape(label)}</button>'
        )
    custom_tools = html.escape(tools or "", quote=True) if selected_mode == "custom" else ""
    disabled = "" if selected_mode == "custom" else " disabled"
    hidden = "" if selected_mode == "custom" else " hidden"
    dirty = ' data-custom-tools-dirty="1"' if selected_mode == "custom" and tools else ""
    return f"""
    <div class="mode-selector" data-mcp-mode-selector>
      <p class="mode-selector__title">Choose tool mode</p>
      <div class="mode-selector__buttons" role="group" aria-label="MCP tool mode">
        {''.join(buttons)}
      </div>
      <label class="mode-selector__custom" data-mcp-custom-field{hidden}>Custom tools
        <input type="text" data-mcp-custom-tools value="{custom_tools}" placeholder="calendar_list,account_info"{dirty}{disabled}>
      </label>
    </div>
    """


def _default_endpoint_warning_html(endpoint_url: str) -> str:
    if not endpoint_url.endswith(DEFAULT_EXAMPLE_PATH):
        return ""
    return """
    <p><strong>Warning</strong>: the default <code>/graph/you@outlook.com</code>
    endpoint and default bearer token are for local debugging only. Do not use this account or token on a public HTTP endpoint.</p>
    """


def _mcp_config(endpoint_url: str, bearer_token: str, *, mode: str = "full", tools: str | None = None) -> dict[str, Any]:
    headers = {"Authorization": f"Bearer {bearer_token}"}
    normalized_mode = mode.strip().lower() or "full"
    if normalized_mode == "custom":
        if tools:
            headers["X-Outlook-MCP-Tools"] = tools
        else:
            headers["X-Outlook-MCP-Mode"] = "custom"
    elif normalized_mode != "full":
        headers["X-Outlook-MCP-Mode"] = normalized_mode
    return {
        "mcpServers": {
            MCP_SERVER_NAME: {
                "type": "streamable-http",
                "url": endpoint_url,
                "headers": headers,
            }
        }
    }


def _agent_prompt(endpoint_url: str, bearer_token: str, *, mode: str = "full", tools: str | None = None) -> str:
    mcp_json = json.dumps(_mcp_config(endpoint_url, bearer_token, mode=mode, tools=tools), ensure_ascii=False, indent=2)
    return (
        "把下面的 MCP server 配置 merge 到当前项目的 .mcp.json 或你的 MCP 客户端配置里。"
        "保留已有 mcpServers，不要覆盖其他 server。\n\n"
        f"{mcp_json}"
    )


def _forwarded_values(header_value: str) -> dict[str, str]:
    first_entry = _first_header_value(header_value)
    values: dict[str, str] = {}
    for raw_pair in first_entry.split(";"):
        if "=" not in raw_pair:
            continue
        key, value = raw_pair.split("=", 1)
        values[key.strip().lower()] = value.strip().strip('"')
    return values


def _first_header_value(value: str | None) -> str:
    if not value:
        return ""
    return value.split(",", 1)[0].strip()


def _host_has_port(host: str) -> bool:
    if host.startswith("["):
        return "]:" in host
    return ":" in host


def _is_default_port(scheme: str, port: str) -> bool:
    return (scheme == "https" and port == "443") or (scheme == "http" and port == "80")


def _normalized_prefix(prefix: str) -> str:
    value = prefix.strip()
    if not value or value == "/":
        return ""
    return "/" + value.strip("/")


def _request_path(scope: dict[str, Any]) -> str:
    raw_path = scope.get("raw_path")
    if isinstance(raw_path, bytes):
        return raw_path.decode("latin1")
    return str(scope.get("path", ""))


def _headers(scope: dict[str, Any]) -> dict[str, str]:
    return {
        key.decode("latin1").lower(): value.decode("latin1")
        for key, value in scope.get("headers", [])
    }


if __name__ == "__main__":
    main()
