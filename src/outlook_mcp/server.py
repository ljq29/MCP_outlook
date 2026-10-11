"""MCP server entry point."""

from __future__ import annotations

import atexit
import logging
import time
import threading
from collections import OrderedDict

from contextvars import ContextVar, Token
from dataclasses import dataclass, replace
from typing import Any

from mcp.server.fastmcp import FastMCP

from .client import create_client
from .config import GraphConfig, RuntimeConfig, load_config
from .credential_store import CredentialStore, SQLiteGraphTokenStorage
from .graph_auth import graph_config_fingerprint
from .graph_client import GraphClient, GraphTokenProvider
from .tools import mail as mail_tools
from .tools import calendar as cal_tools
from .tools import contacts as contact_tools
from .tools import tasks as task_tools
from .tools import account as account_tools
from .tools.common import AuthRequiredError, MethodUnsupportedError


_logger = logging.getLogger(__name__)


class OutlookFastMCP(FastMCP):
    async def list_tools(self) -> list[Any]:
        tools = await super().list_tools()
        enabled_tools = _current_enabled_tools()
        if enabled_tools is None:
            return tools
        return [tool for tool in tools if tool.name in enabled_tools]

    async def call_tool(self, name: str, arguments: dict[str, Any]) -> Any:
        enabled_tools = _current_enabled_tools()
        if enabled_tools is not None and name not in enabled_tools:
            return {
                "ok": False,
                "error": "tool_disabled",
                "message": f"MCP tool is disabled by current MCP tool selection: {name}",
            }
        return await super().call_tool(name, arguments)


mcp = OutlookFastMCP("Outlook MCP", host="127.0.0.1", streamable_http_path="/", json_response=True, stateless_http=True)
_client: Any | None = None
_client_lock = threading.Lock()
_http_store: CredentialStore | None = None
_http_base_config: RuntimeConfig | None = None
_HTTP_CLIENTS_MAX = 32
_HTTP_CLIENTS_MAX_IDLE_SECONDS = 1800  # 30 minutes
_http_clients: OrderedDict[int, Any] = OrderedDict()
_http_lock = threading.Lock()


@dataclass(frozen=True)
class HttpRequestContext:
    method: str
    account: str
    secret_hash: str
    credential_id: int | None
    auth_url: str
    enabled_tools: frozenset[str] | None = None


_http_context: ContextVar[HttpRequestContext | None] = ContextVar("outlook_mcp_http_context", default=None)


def _safe_close(client: Any) -> None:
    try:
        if hasattr(client, "close"):
            client.close()
    except Exception:
        pass


def _close_all_clients_unsafe() -> None:
    """Drain ``_http_clients``, closing every cached client. Caller must hold ``_http_lock``."""
    while _http_clients:
        _, (client, __) = _http_clients.popitem(last=False)
        _safe_close(client)


def configure_http_runtime(store: CredentialStore, base_config: RuntimeConfig | None = None) -> None:
    global _http_store, _http_base_config
    with _http_lock:
        _http_store = store
        _http_base_config = base_config if base_config is not None else load_config()
        _close_all_clients_unsafe()


def set_http_request_context(context: HttpRequestContext) -> Token[HttpRequestContext | None]:
    return _http_context.set(context)


def reset_http_request_context(token: Token[HttpRequestContext | None]) -> None:
    _http_context.reset(token)


def _current_enabled_tools() -> frozenset[str] | None:
    context = _http_context.get()
    return None if context is None else context.enabled_tools


def available_tool_names() -> set[str]:
    tool_manager = getattr(mcp, "_tool_manager", None)
    tools = getattr(tool_manager, "_tools", None)
    if not isinstance(tools, dict):
        _logger.warning("FastMCP private tool registry is unavailable; returning no tool names")
        return set()
    return set(tools.keys())


def http_auth_status(method: str, account: str, credential_id: int | None, auth_url: str) -> dict[str, Any]:
    method = method.strip().lower()
    if method != "graph":
        return {"ok": False, "error": "method_unsupported", "message": f"HTTP MCP method is not implemented yet: {method}"}
    if credential_id is None:
        return {"ok": False, "error": "auth_required", "message": "Open the auth_url in a browser, finish authorization, then retry the MCP call."}

    context = HttpRequestContext(
        method=method,
        account=account,
        secret_hash="",
        credential_id=credential_id,
        auth_url=auth_url,
    )
    client = _get_http_client(context)
    if isinstance(client, _UnsupportedMethodClient):
        return {"ok": False, "error": "method_unsupported", "message": client.message}
    if isinstance(client, _AuthRequiredClient):
        return {"ok": False, "error": "auth_required", "message": client.message, "data": client.data}
    try:
        client.token_provider.get_access_token()
    except AuthRequiredError as exc:
        data = {"auth_url": auth_url, "method": context.method, "account": account, **exc.data}
        return {"ok": False, "error": "auth_required", "message": str(exc), "data": data}
    except Exception as exc:
        return {"ok": False, "error": "outlook_error", "message": str(exc) or exc.__class__.__name__}
    return {"ok": True, "message": "Graph authorization is ready.", "data": {"method": context.method, "account": account}}


def _cleanup_http_clients() -> None:
    with _http_lock:
        _close_all_clients_unsafe()


atexit.register(_cleanup_http_clients)


def clear_http_auth_state(credential_id: int, account: str) -> None:
    with _http_lock:
        store = _http_store
        authorization_id = store.get_graph_authorization_id(credential_id) if store is not None else None
        entry = _http_clients.pop(authorization_id, None) if authorization_id is not None else None
        if entry is not None:
            client, _ = entry
            _safe_close(client)
    if store is not None:
        if authorization_id is None:
            authorization_id = store.resolve_graph_authorization(credential_id, account, "")
        store.clear_graph_authorization(authorization_id)


def get_client() -> Any:
    global _client
    http_context = _http_context.get()
    if http_context is not None:
        return _get_http_client(http_context)
    if _client is not None:
        return _client
    with _client_lock:
        if _client is None:
            _client = create_client(load_config())
            atexit.register(_safe_close, _client)
        return _client


def _get_http_client(context: HttpRequestContext) -> Any:
    if context.credential_id is None:
        return _AuthRequiredClient(context)
    if context.method != "graph":
        return _UnsupportedMethodClient(f"HTTP MCP method is not implemented yet: {context.method}")
    with _http_lock:
        store = _http_store
        base_config = _http_base_config
    if store is None:
        return _UnsupportedMethodClient("HTTP credential store is not configured")
    config = _http_graph_config(context.account, base_config)
    if config is None:
        return _UnsupportedMethodClient("HTTP graph mode requires a provider=graph config with graph.client_id")
    authorization_id = store.resolve_graph_authorization(
        context.credential_id,
        context.account,
        graph_config_fingerprint(config),
    )
    key = authorization_id
    with _http_lock:
        entry = _http_clients.get(key)
        if entry is not None:
            client, last_used = entry
            if time.monotonic() - last_used <= _HTTP_CLIENTS_MAX_IDLE_SECONDS:
                _http_clients[key] = (client, time.monotonic())
                _http_clients.move_to_end(key)
                return client
            # Stale entry — evict and recreate below
            _safe_close(client)
            del _http_clients[key]
    storage = SQLiteGraphTokenStorage(store, authorization_id)
    token_provider = GraphTokenProvider(config, storage=storage)
    client = GraphClient(config, token_provider=token_provider)
    with _http_lock:
        existing = _http_clients.get(key)
        if existing is not None:
            _safe_close(client)
            return existing[0]
        while len(_http_clients) >= _HTTP_CLIENTS_MAX:
            attempts = len(_http_clients)
            evicted_key: Any = None
            evicted_client: Any = None
            while attempts > 0:
                evicted_key, entry = _http_clients.popitem(last=False)
                evicted_client, last_used = entry
                if time.monotonic() - last_used < 10.0 and attempts > 1:
                    _http_clients[evicted_key] = entry
                    _http_clients.move_to_end(evicted_key)
                    attempts -= 1
                    continue
                break
            _safe_close(evicted_client)
        _http_clients[key] = (client, time.monotonic())
    return client


def _http_graph_config(account: str, base_config: RuntimeConfig | None = None) -> GraphConfig | None:
    cfg = base_config if base_config is not None else _http_base_config
    if not isinstance(cfg, GraphConfig):
        return None
    return replace(cfg, email=account)


class _AuthRequiredClient:
    def __init__(self, context: HttpRequestContext):
        self.context = context
        self.message = "Open the auth_url in a browser, finish authorization, then retry the MCP call."
        self.data = {
            "auth_url": context.auth_url,
            "method": context.method,
            "account": context.account,
            "credential_exists": context.credential_id is not None,
        }

    def __getattr__(self, name: str) -> Any:
        def action(*args: Any, **kwargs: Any) -> Any:
            raise AuthRequiredError(self.message, self.data)

        return action


class _UnsupportedMethodClient:
    def __init__(self, message: str):
        self.message = message

    def __getattr__(self, name: str) -> Any:
        def action(*args: Any, **kwargs: Any) -> Any:
            raise MethodUnsupportedError(self.message)

        return action


@mcp.tool()
def mail_list(folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List messages from an Outlook folder."""
    return mail_tools.mail_list(get_client(), folder=folder, limit=limit, offset=offset)


@mcp.tool()
def mail_search(query: str, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search messages by subject within a folder."""
    return mail_tools.mail_search(get_client(), query=query, folder=folder, limit=limit, offset=offset)


@mcp.tool()
def mail_get(item_id: str, folder: str | None = None, include_body: bool = True, body_max_chars: int = 4000) -> dict[str, Any]:
    """Read a single mail item."""
    return mail_tools.mail_get(get_client(),
        item_id=item_id, folder=folder, include_body=include_body, body_max_chars=body_max_chars)


@mcp.tool()
def mail_send(
    to: list[str], subject: str, body: str,
    cc: list[str] | None = None, bcc: list[str] | None = None,
    html: bool = False, attachments: list[dict[str, Any]] | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Send a new message and save it to Sent Items."""
    return mail_tools.mail_send(get_client(),
        to=to, subject=subject, body=body, cc=cc, bcc=bcc,
        html=html, attachments=attachments, confirm=confirm)


@mcp.tool()
def mail_reply(
    item_id: str, body: str, subject: str | None = None,
    folder: str | None = None, to: list[str] | None = None,
    cc: list[str] | None = None, bcc: list[str] | None = None,
    html: bool = False,
    confirm: bool = False,
) -> dict[str, Any]:
    """Reply to an existing message."""
    return mail_tools.mail_reply(get_client(),
        item_id=item_id, body=body, subject=subject, folder=folder,
        to=to, cc=cc, bcc=bcc, html=html, confirm=confirm)


@mcp.tool()
def mail_move(item_id: str, target_folder: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Move a mail item to another folder."""
    return mail_tools.mail_move(get_client(),item_id=item_id, target_folder=target_folder, folder=folder, confirm=confirm)


@mcp.tool()
def mail_delete(item_id: str, folder: str | None = None, mode: str = "move_to_trash", confirm: bool = False) -> dict[str, Any]:
    """Delete a mail item using move_to_trash, soft_delete or hard_delete."""
    return mail_tools.mail_delete(get_client(),item_id=item_id, folder=folder, mode=mode, confirm=confirm)


@mcp.tool()
def calendar_list(start: str | None = None, end: str | None = None,
                  limit: int = 20, offset: int = 0, timezone: str | None = None) -> dict[str, Any]:
    """List calendar items in a time range."""
    return cal_tools.calendar_list(get_client(),start=start, end=end, limit=limit, offset=offset, timezone=timezone)


@mcp.tool()
def calendar_create(
    subject: str, start: str, end: str, body: str = "",
    location: str | None = None, required_attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None, is_all_day: bool = False,
    send_invites: bool = True, html: bool = False,
    attachments: list[dict[str, Any]] | None = None, confirm: bool = False,
    timezone: str | None = None, reminder_minutes_before_start: int | None = None,
) -> dict[str, Any]:
    """Create a calendar item or meeting."""
    return cal_tools.calendar_create(get_client(),
        subject=subject, start=start, end=end, body=body, location=location,
        required_attendees=required_attendees, optional_attendees=optional_attendees,
        is_all_day=is_all_day, send_invites=send_invites, html=html,
        attachments=attachments, confirm=confirm, timezone=timezone,
        reminder_minutes_before_start=reminder_minutes_before_start)


@mcp.tool()
def calendar_update(item_id: str, updates: dict[str, Any],
                    send_invites: bool = True, confirm: bool = False,
                    timezone: str | None = None,
                    html: bool = False) -> dict[str, Any]:
    """Update a calendar item."""
    return cal_tools.calendar_update(get_client(),
        item_id=item_id, updates=updates, send_invites=send_invites, confirm=confirm, timezone=timezone,
        html=html)


@mcp.tool()
def calendar_delete(item_id: str, send_cancellations: bool = True, confirm: bool = False) -> dict[str, Any]:
    """Delete a calendar item."""
    return cal_tools.calendar_delete(get_client(),
        item_id=item_id, send_cancellations=send_cancellations, confirm=confirm)


@mcp.tool()
def contacts_list(limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List contacts."""
    return contact_tools.contacts_list(get_client(),limit=limit, offset=offset)


@mcp.tool()
def contacts_search(query: str, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search contacts by display name."""
    return contact_tools.contacts_search(get_client(),query=query, limit=limit, offset=offset)


@mcp.tool()
def contact_create(
    display_name: str, given_name: str | None = None,
    surname: str | None = None, company_name: str | None = None,
    job_title: str | None = None, department: str | None = None,
    email_addresses: list[str] | None = None,
    phone_numbers: list[dict[str, str]] | None = None,
    notes: str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Create a contact."""
    return contact_tools.contact_create(get_client(),
        display_name=display_name, given_name=given_name, surname=surname,
        company_name=company_name, job_title=job_title, department=department,
        email_addresses=email_addresses, phone_numbers=phone_numbers,
        notes=notes, confirm=confirm)


@mcp.tool()
def contact_update(item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a contact."""
    return contact_tools.contact_update(get_client(),item_id=item_id, updates=updates, confirm=confirm)


@mcp.tool()
def contact_delete(item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a contact."""
    return contact_tools.contact_delete(get_client(),item_id=item_id, confirm=confirm)


@mcp.tool()
def tasks_list(status: str | None = None, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List tasks."""
    return task_tools.tasks_list(get_client(),status=status, limit=limit, offset=offset)


@mcp.tool()
def task_create(
    subject: str, body: str | None = None,
    start_date: str | None = None, due_date: str | None = None,
    status: str = "NotStarted",
    percent_complete: int | float | str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Create a task."""
    return task_tools.task_create(get_client(),
        subject=subject, body=body, start_date=start_date, due_date=due_date,
        status=status, percent_complete=percent_complete, confirm=confirm)


@mcp.tool()
def task_update(item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a task."""
    return task_tools.task_update(get_client(),item_id=item_id, updates=updates, confirm=confirm)


@mcp.tool()
def task_delete(item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a task."""
    return task_tools.task_delete(get_client(),item_id=item_id, confirm=confirm)


# -- new MCP tools (via shared service layer) --

@mcp.tool()
def mail_forward(item_id: str, to: list[str], body: str = "", folder: str | None = None,
                 html: bool = False, confirm: bool = False) -> dict[str, Any]:
    """Forward a message to recipients."""
    return mail_tools.mail_forward(get_client(),
        item_id=item_id, to=to, body=body, folder=folder, html=html, confirm=confirm,
    )


@mcp.tool()
def mail_mark_read(item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as read."""
    return mail_tools.mail_mark_read(get_client(),item_id=item_id, folder=folder, confirm=confirm)


@mcp.tool()
def mail_mark_unread(item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as unread."""
    return mail_tools.mail_mark_unread(get_client(),item_id=item_id, folder=folder, confirm=confirm)


@mcp.tool()
def mail_attachments(item_id: str, folder: str | None = None) -> dict[str, Any]:
    """List attachments on a message."""
    return mail_tools.mail_attachments(get_client(),item_id=item_id, folder=folder)


@mcp.tool()
def mail_folders() -> dict[str, Any]:
    """List available mail folders."""
    return mail_tools.mail_folders(get_client())


@mcp.tool()
def calendar_get(item_id: str, timezone: str | None = None) -> dict[str, Any]:
    """Get a single calendar item by ID."""
    return cal_tools.calendar_get(get_client(),item_id=item_id, timezone=timezone)


@mcp.tool()
def calendar_search(query: str, start: str | None = None, end: str | None = None,
                    limit: int = 20, offset: int = 0,
                    timezone: str | None = None) -> dict[str, Any]:
    """Search calendar items by subject."""
    return cal_tools.calendar_search(get_client(),
        query=query, start=start, end=end, limit=limit, offset=offset, timezone=timezone,
    )


@mcp.tool()
def contact_get(item_id: str) -> dict[str, Any]:
    """Get a single contact by ID."""
    return contact_tools.contact_get(get_client(),item_id=item_id)


@mcp.tool()
def task_get(item_id: str) -> dict[str, Any]:
    """Get a single task by ID."""
    return task_tools.task_get(get_client(),item_id=item_id)


@mcp.tool()
def task_search(query: str, status: str | None = None,
                limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search tasks by subject."""
    return task_tools.task_search(get_client(),
        query=query, status=status, limit=limit, offset=offset,
    )


@mcp.tool()
def account_info() -> dict[str, Any]:
    """Return the current account email and display name."""
    return account_tools.account_info(get_client())


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
