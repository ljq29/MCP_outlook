from __future__ import annotations

from typing import Any

from outlook_mcp import server
from outlook_mcp.credential_store import CredentialStore
from outlook_mcp.tools.common import AuthRequiredError

from conftest import FakeGraphClient

def test_server_tool_wrappers_delegate_to_tool_modules(monkeypatch) -> None:
    client = FakeGraphClient()
    monkeypatch.setattr(server, "get_client", lambda: client)

    assert server.mail_list()["ok"] is True
    assert server.mail_search("report")["ok"] is True
    assert server.mail_get("mail-1")["ok"] is True
    assert server.mail_send(["a@example.com"], "Hi", "Body", confirm=True)["ok"] is True
    assert server.mail_reply("mail-1", "<p>Thanks</p>", html=True, confirm=True)["ok"] is True
    assert server.mail_move("mail-1", "trash", confirm=True)["ok"] is True
    assert server.mail_delete("mail-1", confirm=True)["ok"] is True

    start = "2026-01-02T00:00:00+00:00"
    end = "2026-01-03T00:00:00+00:00"
    assert server.calendar_list(start=start, end=end)["ok"] is True
    assert server.calendar_create("Meet", start, end, confirm=True)["ok"] is True
    assert server.calendar_create("Meet", start, end, confirm=True, reminder_minutes_before_start=10)["ok"] is True
    assert client.last_call("calendar_create")["reminder_minutes_before_start"] == 10
    assert server.calendar_update("cal-1", {"location": "Room 3"}, confirm=True)["ok"] is True
    assert server.calendar_delete("cal-1", confirm=True)["ok"] is True

    assert server.contacts_list()["ok"] is True
    assert server.contacts_search("Ada")["ok"] is True
    assert server.contact_create("Grace Hopper", confirm=True)["ok"] is True
    assert server.contact_update("contact-1", {"job_title": "Engineer"}, confirm=True)["ok"] is True
    assert server.contact_delete("contact-1", confirm=True)["ok"] is True

    assert server.tasks_list()["ok"] is True
    assert server.task_create("Write notes", confirm=True)["ok"] is True
    assert server.task_update("task-1", {"status": "InProgress"}, confirm=True)["ok"] is True
    assert server.task_delete("task-1", confirm=True)["ok"] is True


def test_new_mcp_tools(monkeypatch) -> None:
    """New tools added after architecture refactoring."""
    client = FakeGraphClient()
    monkeypatch.setattr(server, "get_client", lambda: client)

    # mail
    assert server.mail_forward("mail-1", ["a@example.com"], confirm=True)["ok"] is True
    blocked = server.mail_mark_read("mail-1")
    assert blocked["ok"] is False
    assert blocked["error"] == "validation_error"
    result = server.mail_mark_read("mail-1", confirm=True)
    assert result["ok"] is True
    assert result["data"]["marked"] == "read"
    blocked = server.mail_mark_unread("mail-1")
    assert blocked["ok"] is False
    assert blocked["error"] == "validation_error"
    result = server.mail_mark_unread("mail-1", confirm=True)
    assert result["data"]["marked"] == "unread"
    assert server.mail_attachments("mail-1")["ok"] is True
    folders = server.mail_folders()
    assert folders["ok"] is True
    assert "has_more" in folders["data"]

    # calendar
    assert server.calendar_get("cal-1")["ok"] is True
    cal_search = server.calendar_search("Planning")
    assert cal_search["ok"] is True
    assert "has_more" in cal_search["data"]

    # contact
    assert server.contact_get("contact-1")["ok"] is True

    # task
    assert server.task_get("task-1")["ok"] is True
    task_search = server.task_search("Follow")
    assert task_search["ok"] is True
    assert "has_more" in task_search["data"]

    # account
    info = server.account_info()
    assert info["ok"] is True
    assert "display_name" in info["data"]

    # pagination on existing tools
    mail = server.mail_list(limit=5)
    assert "has_more" in mail["data"]
    assert "total" not in mail["data"]


def test_lru_eviction_does_not_hang_on_all_recent_clients(monkeypatch, tmp_path) -> None:
    """Drive the real _get_http_client eviction path with a full all-recent cache."""
    import time

    from outlook_mcp.config import GraphConfig
    from outlook_mcp.server import _HTTP_CLIENTS_MAX

    class _FakeTokenProvider:
        def __init__(self, config: Any, storage: Any = None):
            self.config = config

    class _FakeGraphClient:
        def __init__(self, config: Any, token_provider: Any = None):
            self.closed = False

        def close(self) -> None:
            self.closed = True

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    base_cfg = GraphConfig(
        email="person@outlook.com",
        client_id="test-client-id",
        tenant="consumers",
        token_cache_path=str(tmp_path / "cache.json"),
    )
    server.configure_http_runtime(store, base_config=base_cfg)
    monkeypatch.setattr(server, "GraphTokenProvider", _FakeTokenProvider)
    monkeypatch.setattr(server, "GraphClient", _FakeGraphClient)

    prefilled: list[_FakeGraphClient] = []
    try:
        with server._http_lock:
            for i in range(_HTTP_CLIENTS_MAX):
                cached = _FakeGraphClient(base_cfg)
                prefilled.append(cached)
                server._http_clients[f"existing-{i}"] = (cached, time.monotonic())

        context = server.HttpRequestContext(
            method="graph",
            account="person@outlook.com",
            secret_hash="",
            credential_id=record.id,
            auth_url="http://localhost/graph/person@outlook.com/123456",
        )
        client = server._get_http_client(context)

        assert isinstance(client, _FakeGraphClient)
        assert client not in prefilled
        assert not client.closed
        assert sum(1 for cached in prefilled if cached.closed) == 1
        assert len(server._http_clients) == _HTTP_CLIENTS_MAX
    finally:
        with server._http_lock:
            server._http_store = None
            server._http_base_config = None
            server._http_clients.clear()


def test_lru_eviction_stale_entry_by_ttl(monkeypatch, tmp_path) -> None:
    """A stale cached client (older than the idle TTL) is closed and recreated on cache hit."""
    import time

    from outlook_mcp.config import GraphConfig
    from outlook_mcp.graph_auth import graph_config_fingerprint
    from outlook_mcp.server import _HTTP_CLIENTS_MAX_IDLE_SECONDS

    class _FakeTokenProvider:
        def __init__(self, config: Any, storage: Any = None):
            self.config = config

    class _FakeGraphClient:
        def __init__(self, config: Any, token_provider: Any = None):
            self.closed = False

        def close(self) -> None:
            self.closed = True

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    base_cfg = GraphConfig(
        email="person@outlook.com",
        client_id="test-client-id",
        tenant="consumers",
        token_cache_path=str(tmp_path / "cache.json"),
    )
    server.configure_http_runtime(store, base_config=base_cfg)
    monkeypatch.setattr(server, "GraphTokenProvider", _FakeTokenProvider)
    monkeypatch.setattr(server, "GraphClient", _FakeGraphClient)

    try:
        config = server._http_graph_config(record.account, base_cfg)
        authorization_id = store.resolve_graph_authorization(
            record.id, record.account, graph_config_fingerprint(config)
        )
        stale = _FakeGraphClient(base_cfg)
        stale_time = time.monotonic() - _HTTP_CLIENTS_MAX_IDLE_SECONDS - 1
        with server._http_lock:
            server._http_clients[authorization_id] = (stale, stale_time)

        context = server.HttpRequestContext(
            method="graph",
            account="person@outlook.com",
            secret_hash="",
            credential_id=record.id,
            auth_url="http://localhost/graph/person@outlook.com/123456",
        )
        client = server._get_http_client(context)

        assert stale.closed
        assert isinstance(client, _FakeGraphClient)
        assert client is not stale
        assert server._http_clients[authorization_id][0] is client
    finally:
        with server._http_lock:
            server._http_store = None
            server._http_base_config = None
            server._http_clients.clear()


def test_lru_eviction_fresh_entry_not_evicted_by_ttl() -> None:
    """Fresh entries (within TTL) must remain in cache."""
    import time
    from outlook_mcp.server import _http_clients, _HTTP_CLIENTS_MAX_IDLE_SECONDS, _http_lock

    class _FakeHttpClient:
        closed = False

        def close(self):
            self.closed = True

    with _http_lock:
        _http_clients.clear()
        try:
            client = _FakeHttpClient()
            _http_clients[100] = (client, time.monotonic())

            entry = _http_clients.get(100)
            _c, last_used = entry
            assert time.monotonic() - last_used <= _HTTP_CLIENTS_MAX_IDLE_SECONDS
            assert not client.closed
        finally:
            _http_clients.clear()


def test_paginate_does_not_mutate_error_responses() -> None:
    from outlook_mcp.tools.common import run_safely

    def action():
        raise ValueError("test error")

    result = run_safely(action)
    assert result["ok"] is False
    assert result["error"] == "exchange_error"
    assert "has_more" not in result


# --- http_auth_status authorization boundary ---

def test_http_auth_status_delegates_to_authorization_provider(monkeypatch, tmp_path) -> None:
    from outlook_mcp.server import http_auth_status, configure_http_runtime
    from outlook_mcp.config import GraphConfig

    class ReadyClient:
        class TokenProvider:
            def get_access_token(self) -> str:
                return "access-token"

        token_provider = TokenProvider()

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    base_cfg = GraphConfig(
        email="person@outlook.com",
        client_id="test-client-id",
        tenant="consumers",
        token_cache_path=str(tmp_path / "cache.json"),
    )
    configure_http_runtime(store, base_config=base_cfg)
    monkeypatch.setattr(server, "_get_http_client", lambda context: ReadyClient())

    try:
        result = http_auth_status("graph", "person@outlook.com", record.id, "http://localhost/graph/person@outlook.com/123456")
        assert result["ok"] is True
        assert "ready" in result["message"]
    finally:
        with server._http_lock:
            server._http_store = None
            server._http_base_config = None
            server._http_clients.clear()


def test_http_auth_status_does_not_use_msal_username_heuristic(monkeypatch, tmp_path) -> None:
    from outlook_mcp.server import http_auth_status, configure_http_runtime
    from outlook_mcp.config import GraphConfig

    class ReadyClient:
        class TokenProvider:
            def get_access_token(self) -> str:
                return "access-token"

        token_provider = TokenProvider()

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    store.save_msal_cache(
        record.id,
        '{"AccessToken": {"secret": "tok"}, "Account": {"home": {"username": "login@live.com"}}}',
    )
    base_cfg = GraphConfig(
        email="person@outlook.com",
        client_id="test-client-id",
        tenant="consumers",
        token_cache_path=str(tmp_path / "cache.json"),
    )
    configure_http_runtime(store, base_config=base_cfg)
    monkeypatch.setattr(server, "_get_http_client", lambda context: ReadyClient())

    try:
        result = http_auth_status("graph", "person@outlook.com", record.id, "http://localhost/graph/person@outlook.com/123456")
        assert result["ok"] is True
        assert "ready" in result["message"]
    finally:
        with server._http_lock:
            server._http_store = None
            server._http_base_config = None
            server._http_clients.clear()


def test_http_auth_status_does_not_accept_mismatched_msal_account(monkeypatch, tmp_path) -> None:
    from outlook_mcp.server import http_auth_status, configure_http_runtime
    from outlook_mcp.config import GraphConfig

    class FakeClient:
        class TokenProvider:
            def get_access_token(self) -> str:
                raise AuthRequiredError("Login required", {"user_code": "ABCD-EFGH"})

        token_provider = TokenProvider()

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    store.save_msal_cache(
        record.id,
        '{"AccessToken": {"secret": "tok"}, "Account": {"home": {"username": "other@outlook.com"}}}',
    )
    base_cfg = GraphConfig(
        email="person@outlook.com",
        client_id="test-client-id",
        tenant="consumers",
        token_cache_path=str(tmp_path / "cache.json"),
    )
    configure_http_runtime(store, base_config=base_cfg)
    monkeypatch.setattr(server, "_get_http_client", lambda context: FakeClient())

    try:
        result = http_auth_status("graph", "person@outlook.com", record.id, "http://localhost/graph/person@outlook.com/123456")
        assert result["ok"] is False
        assert result["error"] == "auth_required"
        assert result["data"]["user_code"] == "ABCD-EFGH"
    finally:
        with server._http_lock:
            server._http_store = None
            server._http_base_config = None
            server._http_clients.clear()


def test_http_auth_status_no_credential_id() -> None:
    from outlook_mcp.server import http_auth_status

    result = http_auth_status("graph", "e@outlook.com", None, "http://localhost/graph/e@outlook.com/123456")
    assert result["ok"] is False
    assert result["error"] == "auth_required"


def test_http_context_filters_fastmcp_tools() -> None:
    import asyncio

    context = server.HttpRequestContext(
        method="graph",
        account="person@outlook.com",
        secret_hash="hash",
        credential_id=1,
        auth_url="http://localhost/graph/person@outlook.com/123456",
        enabled_tools=frozenset({"calendar_list", "account_info"}),
    )
    token = server.set_http_request_context(context)
    try:
        tools = asyncio.run(server.mcp.list_tools())
    finally:
        server.reset_http_request_context(token)

    assert {tool.name for tool in tools} == {"calendar_list", "account_info"}


def test_http_context_blocks_disabled_fastmcp_tool() -> None:
    import asyncio

    context = server.HttpRequestContext(
        method="graph",
        account="person@outlook.com",
        secret_hash="hash",
        credential_id=1,
        auth_url="http://localhost/graph/person@outlook.com/123456",
        enabled_tools=frozenset({"calendar_list"}),
    )
    token = server.set_http_request_context(context)
    try:
        result = asyncio.run(server.mcp.call_tool("mail_list", {"limit": 1}))
    finally:
        server.reset_http_request_context(token)

    assert result["ok"] is False
    assert result["error"] == "tool_disabled"
    assert "mail_list" in result["message"]


# --- atexit cleanup ---

def test_atexit_cleanup_closes_all_cached_clients() -> None:
    from outlook_mcp.server import _cleanup_http_clients, _http_clients, _http_lock
    import time

    class _CloseTrackingClient:
        closed = False

        def close(self):
            self.closed = True

    with _http_lock:
        _http_clients.clear()
        c1 = _CloseTrackingClient()
        c2 = _CloseTrackingClient()
        _http_clients[1] = (c1, time.monotonic())
        _http_clients[2] = (c2, time.monotonic())

    try:
        _cleanup_http_clients()

        assert c1.closed
        assert c2.closed
        assert len(_http_clients) == 0
    finally:
        with _http_lock:
            _http_clients.clear()


def test_registered_tools_match_api_dispatch_and_mode_presets() -> None:
    from outlook_mcp.http_server import TOOL_MODE_PRESETS
    from outlook_mcp.tools import get_api_dispatch

    registered = server.available_tool_names()
    assert registered, "available_tool_names() returned an empty set — FastMCP private registry changed?"
    dispatch_names = {getattr(func, "__name__", "") for func in get_api_dispatch().values()}
    assert registered == dispatch_names
    for preset in TOOL_MODE_PRESETS:
        tools = preset["tools"]
        if not tools:
            continue
        for name in str(tools).split(","):
            assert name in registered, f"mode {preset['mode']} references unknown tool {name}"
