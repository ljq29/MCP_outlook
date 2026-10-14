from __future__ import annotations

import hashlib
import json
import os
import sys
import threading
from datetime import datetime, timezone
from typing import Any

import pytest
from starlette.testclient import TestClient

from outlook_mcp import server
from outlook_mcp.client import OutlookClientError
from outlook_mcp.config import GraphConfig
from outlook_mcp.credential_store import CredentialStore
from outlook_mcp.graph_client import GraphClient
from outlook_mcp.http_server import (
    _MAIL_RESPONSE_CACHE,
    _MAX_API_BODY_BYTES,
    _RequestResponseCapture,
    _UPLOAD_BUCKETS,
    _check_upload_rate,
    _date_start_timestamp,
    _dispatch_api,
    _first_header_value,
    _forwarded_values,
    _host_has_port,
    _is_default_port,
    _load_average,
    _log_filter,
    _make_csrf_token,
    _memory_rss_bytes,
    _normalized_prefix,
    _operation_error,
    _operation_from_request_body,
    _operation_ok,
    _read_request_body,
    _result_status_code,
    _verify_csrf_token,
    _wants_mcp_stream,
    create_app,
)
from conftest import PNG_BYTES
from outlook_mcp.tools import mail as mail_tools
from outlook_mcp.tools.common import AuthRequiredError, ToolValidationError
from outlook_mcp.uploads import current_upload_credential_id


ALL_CUSTOM_TOOLS = (
    "mail_list,mail_search,mail_get,mail_send,mail_reply,mail_forward,"
    "mail_move,mail_delete,mail_mark_read,mail_mark_unread,mail_attachments,mail_folders,"
    "calendar_list,calendar_get,calendar_search,calendar_create,calendar_update,calendar_delete,"
    "contacts_list,contacts_search,contact_get,contact_create,contact_update,contact_delete,"
    "tasks_list,task_get,task_search,task_create,task_update,task_delete,account_info"
)
CALENDAR_CUSTOM_TOOLS = "calendar_list,calendar_get,calendar_search,calendar_create,calendar_update,calendar_delete,account_info"


def _normalize_line_endings(text: str) -> str:
    return text.replace("\r\n", "\n")


def test_health_is_liveness_only_and_does_not_touch_graph(monkeypatch, tmp_path) -> None:
    def fail_if_called(*args: Any, **kwargs: Any) -> None:
        raise AssertionError("health must not touch Graph/auth")

    monkeypatch.setattr("outlook_mcp.server._get_http_client", fail_if_called)
    monkeypatch.setattr("outlook_mcp.server.http_auth_status", fail_if_called)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/health")

    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["service"] == "outlook-http"
    assert isinstance(payload["uptime_seconds"], int | float)
    assert payload["uptime_seconds"] >= 0
    assert isinstance(payload["resources"]["process_cpu_seconds"], int | float)
    assert "memory_rss_bytes" in payload["resources"]
    assert "load_average" in payload["resources"]
    assert "cpu_count" in payload["resources"]


def test_health_deep_is_separate_from_liveness(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/health/deep")

    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["service"] == "outlook-http"
    assert payload["checks"]["process"]["status"] == "ok"
    assert payload["checks"]["process"]["service"] == "outlook-http"


def test_health_degrades_when_resource_metrics_unavailable(monkeypatch, tmp_path) -> None:
    monkeypatch.setitem(sys.modules, "resource", None)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/health")

    assert response.status_code == 200
    assert response.json()["resources"]["memory_rss_bytes"] is None


def test_health_degrades_when_load_average_unavailable(monkeypatch, tmp_path) -> None:
    monkeypatch.delattr(os, "getloadavg")
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/health")

    assert response.status_code == 200
    assert response.json()["resources"]["load_average"] is None


def test_resource_helpers_degrade_on_os_errors(monkeypatch) -> None:
    import resource

    def fail_getrusage(_who: int) -> None:
        raise OSError("rusage unavailable")

    def fail_getloadavg() -> None:
        raise OSError("load average unavailable")

    monkeypatch.setattr(resource, "getrusage", fail_getrusage)
    monkeypatch.setattr(os, "getloadavg", fail_getloadavg)

    assert _memory_rss_bytes() is None
    assert _load_average() is None


def test_api_graph_timeout_is_business_error_not_health_error(monkeypatch, tmp_path) -> None:
    def graph_timeout(*args: Any, **kwargs: Any) -> None:
        raise OutlookClientError("Network error connecting to Microsoft Graph: timed out")

    monkeypatch.setattr("outlook_mcp.http_server._do_dispatch", graph_timeout)
    result = _dispatch_api("mail/list", object(), {"limit": 20, "offset": 0})

    assert result["ok"] is False
    assert result["error"] == "outlook_error"
    assert "timed out" in result["message"]

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/health")

    assert response.status_code == 200
    assert response.json()["status"] == "ok"


@pytest.mark.parametrize(
    ("error", "expected_status"),
    [
        ("validation_error", 400),
        ("tool_disabled", 400),
        ("auth_required", 401),
        ("not_found", 404),
        ("outlook_error", 502),
        ("exchange_error", 502),
        ("method_unsupported", 502),
        ("outcome_unknown", 502),
        ("unexpected_internal", 502),
    ],
)
def test_result_status_code_maps_error_categories(error: str, expected_status: int) -> None:
    assert _result_status_code({"ok": False, "error": error, "message": "boom"}) == expected_status


def test_result_status_code_prefers_explicit_status_and_ok() -> None:
    assert _result_status_code({"ok": False, "error": "draft_conflict", "_status_code": 409}) == 409
    assert _result_status_code({"ok": True, "data": {}}) == 200


@pytest.mark.parametrize(
    ("error", "expected_status"),
    [
        ("validation_error", 400),
        ("auth_required", 401),
        ("not_found", 404),
        ("outlook_error", 502),
        ("exchange_error", 502),
    ],
)
def test_http_api_error_category_maps_to_status(monkeypatch, tmp_path, error: str, expected_status: int) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    monkeypatch.setattr(
        "outlook_mcp.server._get_http_client",
        lambda context: GraphClient.__new__(GraphClient),
    )
    monkeypatch.setattr(
        "outlook_mcp.http_server._dispatch_api",
        lambda *args, **kwargs: {"ok": False, "error": error, "message": "boom"},
    )

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/api/mail/list",
            json={"limit": 20, "offset": 0},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == expected_status
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == error


def test_dispatch_api_redacts_secrets_in_exception_messages(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_CLIENT_SECRET", "s3cr3t-token-value")

    def leak(*args: Any, **kwargs: Any) -> None:
        raise ToolValidationError("bad header s3cr3t-token-value rejected")

    monkeypatch.setattr("outlook_mcp.http_server._do_dispatch", leak)
    result = _dispatch_api("mail/list", object(), {})

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "s3cr3t-token-value" not in result["message"]
    assert "[redacted]" in result["message"]


def test_dispatch_api_auth_required_carries_data(monkeypatch) -> None:
    def need_auth(*args: Any, **kwargs: Any) -> None:
        raise AuthRequiredError("login required", {"auth_url": "http://example.test/auth"})

    monkeypatch.setattr("outlook_mcp.http_server._do_dispatch", need_auth)
    result = _dispatch_api("mail/list", object(), {})

    assert result == {
        "ok": False,
        "error": "auth_required",
        "message": "login required",
        "data": {"auth_url": "http://example.test/auth"},
    }


def test_dispatch_api_unexpected_exception_maps_to_exchange_error(monkeypatch) -> None:
    def boom(*args: Any, **kwargs: Any) -> None:
        raise RuntimeError("kaboom")

    monkeypatch.setattr("outlook_mcp.http_server._do_dispatch", boom)
    result = _dispatch_api("mail/list", object(), {})

    assert result["ok"] is False
    assert result["error"] == "exchange_error"
    assert "kaboom" in result["message"]


def test_http_mcp_missing_endpoint_returns_auth_required(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "Origin": "http://testserver",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "auth_required"
    assert payload["data"]["auth_url"].endswith("/graph/person@outlook.com/test-secret-at-least-24-chars")
    assert store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars") is None


def test_http_mcp_bearer_endpoint_uses_authorization_header(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_context_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/call", "params": {"name": "mail_list"}},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["data"]["method"] == "graph"
    assert payload["data"]["account"] == "person@outlook.com"
    assert payload["data"]["credential_id"] == record.id
    assert payload["data"]["auth_url"].endswith("/graph/person@outlook.com/test-secret-at-least-24-chars")


def test_http_mcp_missing_bearer_header_returns_auth_required(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com", json={})

    assert response.status_code == 401
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "auth_required"
    assert "Authorization: Bearer" in payload["message"]


def test_http_mcp_legacy_secret_url_no_longer_dispatches_to_mcp(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars", json={})

    assert response.status_code == 405
    assert response.text == "Method Not Allowed"


def test_http_api_secret_scoped_path_is_rejected(monkeypatch, tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    monkeypatch.setattr("outlook_mcp.http_server._dispatch_api", lambda *args, **kwargs: {"ok": True, "data": {"items": []}})

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/api/calendar/list", json={"limit": 20, "offset": 0})

    assert response.status_code == 404
    assert response.json()["error"] == "invalid_endpoint"


def test_http_api_accepts_arbitrary_origin_and_uses_bearer_context(monkeypatch, tmp_path) -> None:
    captured: dict[str, Any] = {}
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
    )
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    def fake_get_http_client(context) -> GraphClient:
        captured["credential_id"] = context.credential_id
        return GraphClient.__new__(GraphClient)

    def fake_dispatch(*args: Any, **kwargs: Any) -> dict[str, Any]:
        captured["upload_credential_id"] = current_upload_credential_id()
        captured["dispatch_thread"] = threading.get_ident()
        return {"ok": True, "data": {"items": []}}

    monkeypatch.setattr("outlook_mcp.server._get_http_client", fake_get_http_client)
    monkeypatch.setattr("outlook_mcp.http_server._dispatch_api", fake_dispatch)

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/api/calendar/list",
            json={"limit": 20, "offset": 0},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "Origin": "https://attacker.example",
            },
        )

    assert response.status_code == 200
    assert response.json()["ok"] is True
    assert captured["credential_id"] == record.id
    assert captured["upload_credential_id"] == record.id
    assert isinstance(captured["dispatch_thread"], int)


def test_http_mail_read_cache_hits_and_write_invalidates(monkeypatch, tmp_path) -> None:
    calls: list[str] = []
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
    )
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app
    _MAIL_RESPONSE_CACHE.clear()

    monkeypatch.setattr(
        "outlook_mcp.server._get_http_client",
        lambda context: GraphClient.__new__(GraphClient),
    )

    def fake_dispatch(api_path: str, *args: Any, **kwargs: Any) -> dict[str, Any]:
        calls.append(api_path)
        if api_path == "mail/list":
            generation = calls.count("mail/list")
            return {"ok": True, "data": {"items": [{"id": f"mail-{generation}"}]}}
        return {"ok": True, "data": {"marked": "read"}}

    monkeypatch.setattr("outlook_mcp.http_server._dispatch_api", fake_dispatch)
    headers = {"Authorization": "Bearer test-secret-at-least-24-chars"}

    try:
        with TestClient(app) as client:
            first = client.post(
                "/graph/person@outlook.com/api/mail/list",
                json={"folder": "inbox", "limit": 20, "offset": 0},
                headers=headers,
            )
            second = client.post(
                "/graph/person@outlook.com/api/mail/list",
                json={"folder": "inbox", "limit": 20, "offset": 0},
                headers=headers,
            )
            write = client.post(
                "/graph/person@outlook.com/api/mail/mark_read",
                json={"item_id": "mail-1", "folder": "inbox", "confirm": True},
                headers=headers,
            )
            third = client.post(
                "/graph/person@outlook.com/api/mail/list",
                json={"folder": "inbox", "limit": 20, "offset": 0},
                headers=headers,
            )
    finally:
        _MAIL_RESPONSE_CACHE.clear()

    assert first.json()["data"]["items"] == [{"id": "mail-1"}]
    assert second.json()["data"]["items"] == [{"id": "mail-1"}]
    assert write.status_code == 200
    assert third.json()["data"]["items"] == [{"id": "mail-2"}]
    assert calls == ["mail/list", "mail/mark_read", "mail/list"]


def test_calendar_draft_api_works_without_graph_authorization(monkeypatch, tmp_path) -> None:
    def fail_if_graph_is_requested(*args: Any, **kwargs: Any) -> None:
        raise AssertionError("draft CRUD must not require Graph authorization")

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
    )
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app
    monkeypatch.setattr("outlook_mcp.server._get_http_client", fail_if_graph_is_requested)
    headers = {"Authorization": "Bearer test-secret-at-least-24-chars"}

    with TestClient(app) as client:
        saved = client.post(
            "/graph/person@outlook.com/api/calendar/drafts/save",
            json={
                "confirm": True,
                "draft_type": "create",
                "payload": {
                    "subject": "跨设备草稿",
                    "start": "",
                    "end": "",
                    "timezone": "Asia/Shanghai",
                },
            },
            headers=headers,
        )
        listed = client.post(
            "/graph/person@outlook.com/api/calendar/drafts/list",
            json={"limit": 100, "offset": 0},
            headers=headers,
        )

    assert saved.status_code == 200
    assert saved.json()["data"]["draft"]["subject"] == "跨设备草稿"
    assert listed.status_code == 200
    assert listed.json()["data"]["total"] == 1


def test_calendar_draft_api_returns_http_conflict_for_stale_version(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
    )
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app
    headers = {"Authorization": "Bearer test-secret-at-least-24-chars"}
    payload = {
        "confirm": True,
        "draft_type": "create",
        "payload": {
            "subject": "第一版",
            "start": "2026-07-16T09:00",
            "end": "2026-07-16T10:00",
            "timezone": "Asia/Shanghai",
        },
    }

    with TestClient(app) as client:
        created = client.post(
            "/graph/person@outlook.com/api/calendar/drafts/save",
            json=payload,
            headers=headers,
        ).json()["data"]["draft"]
        update_payload = {
            **payload,
            "draft_id": created["id"],
            "version": created["version"],
        }
        first = client.post(
            "/graph/person@outlook.com/api/calendar/drafts/save",
            json=update_payload,
            headers=headers,
        )
        stale = client.post(
            "/graph/person@outlook.com/api/calendar/drafts/save",
            json=update_payload,
            headers=headers,
        )

    assert first.status_code == 200
    assert stale.status_code == 409
    assert stale.json()["error"] == "draft_conflict"
    assert stale.json()["data"]["draft"]["version"] == 2


def test_http_api_rejects_wrong_method_and_media_type(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        wrong_method = client.get("/graph/person@outlook.com/api/calendar/list")
        wrong_type = client.post(
            "/graph/person@outlook.com/api/calendar/list",
            content=b"{}",
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "Content-Type": "text/plain",
            },
        )
    assert wrong_method.status_code == 405
    assert wrong_type.status_code == 415
    assert wrong_type.json()["error"] == "unsupported_media_type"


def test_http_upload_image_success(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/uploads",
            files={"file": ("../photo.png", PNG_BYTES, "image/png")},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "Origin": "https://attacker.example",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["relative_path"].startswith(f"uploads/{record.id}/")
    assert payload["relative_path"].endswith("/photo.png")
    assert str(tmp_path) not in payload["relative_path"]
    assert "test-secret-at-least-24-chars" not in payload["relative_path"]
    assert payload["mime_type"] == "image/png"
    assert payload["size"] == len(PNG_BYTES)
    assert payload["sha256"] == hashlib.sha256(PNG_BYTES).hexdigest()
    assert payload["width"] == 320
    assert payload["height"] == 180


def test_http_upload_rejects_endpoint_without_completed_oauth(monkeypatch, tmp_path) -> None:
    upload_dir = tmp_path / "uploads"
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(upload_dir))
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_pending_flow(record.id, {"user_code": "ABCD-EFGH", "expires_in": 600})
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/uploads",
            files={"file": ("photo.png", PNG_BYTES, "image/png")},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == 403
    payload = response.json()
    assert payload["error"] == "auth_required"
    assert "Microsoft OAuth authorization" in payload["message"]
    assert not upload_dir.exists()


def test_http_upload_rejects_unregistered_endpoint_and_bad_file(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        missing = client.post(
            "/graph/person@outlook.com/uploads",
            files={"file": ("photo.png", PNG_BYTES, "image/png")},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )
        record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
        store.save_msal_cache(record.id, '{"AccessToken": {}}')
        bad_file = client.post(
            "/graph/person@outlook.com/uploads",
            files={"file": ("note.txt", b"hello", "text/plain")},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert missing.status_code == 403
    assert missing.json()["error"] == "auth_required"
    assert bad_file.status_code == 400
    assert bad_file.json()["error"] == "validation_error"


def test_http_upload_rejects_aggregate_body_before_multipart_parsing(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_MAX_BYTES", "8")
    _UPLOAD_BUCKETS.clear()
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/uploads",
            content=b"x" * (8 + 256 * 1024 + 1),
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "Content-Type": "multipart/form-data; boundary=test",
            },
        )

    assert response.status_code == 413
    assert response.json()["error"] == "request_too_large"
    assert not (tmp_path / "uploads").exists()


def test_http_upload_rejects_more_than_one_file(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    _UPLOAD_BUCKETS.clear()
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/uploads",
            files=[
                ("file", ("one.png", PNG_BYTES, "image/png")),
                ("file", ("two.png", PNG_BYTES, "image/png")),
            ],
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == 400
    assert response.json()["error"] == "validation_error"
    assert not (tmp_path / "uploads").exists()


@pytest.mark.parametrize("suffix", ["", "/log", "/calendar", "/mail"])
def test_http_browser_pages_require_preregistered_graph_endpoint(tmp_path, suffix: str) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            f"/graph/person@outlook.com/test-secret-at-least-24-chars{suffix}"
        )

    assert response.status_code == 404
    assert "credential_not_registered" in response.text
    assert "outlook_mcp.credential_store register" in response.text
    assert store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars") is None


def test_http_auth_page_starts_graph_authorization(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars")

    assert response.status_code == 200
    assert "<style>" in response.text
    assert "Pico CSS" in response.text
    assert "cdn.jsdelivr" not in response.text
    assert "Authorize Outlook MCP" in response.text
    assert "ABCD-EFGH" in response.text
    assert "Add this endpoint to MCP" in response.text
    assert "Choose tool mode" in response.text
    assert 'data-mcp-mode-selector' in response.text
    assert 'data-mode="full"' in response.text
    assert 'data-mode="calendar"' in response.text
    assert 'data-mode="custom"' in response.text
    assert 'data-mode="full" aria-pressed="true"' in response.text
    assert f'data-custom-tools="{ALL_CUSTOM_TOOLS}"' in response.text
    assert "customInput.value = seedTools" in response.text
    assert "previousSelected === selectedButton" in response.text
    assert "streamable-http" in response.text
    assert "outlook-http" in response.text
    assert "http://testserver/graph/person@outlook.com" in response.text
    assert "&quot;Authorization&quot;: &quot;Bearer test-secret-at-least-24-chars&quot;" in response.text
    assert "window.location.origin" in response.text
    assert "data-mcp-install" in response.text
    assert 'data-endpoint-path="/graph/person@outlook.com"' in response.text
    assert "data-mcp-rewrite-url" in response.text
    assert 'class="secondary copy-button"' in response.text
    assert 'data-copy-source="mcp-json"' in response.text
    assert "navigator.clipboard" in response.text
    assert 'document.execCommand("copy")' in response.text
    assert "把下面的 MCP server 配置 merge" in response.text
    record = store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    assert record is not None
    assert store.load_pending_flow(record.id)["user_code"] == "ABCD-EFGH"


def test_http_auth_page_calendar_mode_updates_mcp_json(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars?mode=calendar")

    assert response.status_code == 200
    assert 'data-mode="calendar" aria-pressed="true"' in response.text
    assert f'data-custom-tools="{CALENDAR_CUSTOM_TOOLS}"' in response.text
    assert "&quot;X-Outlook-MCP-Mode&quot;: &quot;calendar&quot;" in response.text
    assert f"&quot;X-Outlook-MCP-Tools&quot;: &quot;{CALENDAR_CUSTOM_TOOLS}&quot;" not in response.text


def test_http_auth_page_tools_query_selects_custom_mode(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars?mode=calendar&tools=calendar_list,account_info")

    assert response.status_code == 200
    assert 'data-mode="custom" aria-pressed="true"' in response.text
    assert 'value="calendar_list,account_info"' in response.text
    assert 'data-custom-tools-dirty="1"' in response.text
    assert "&quot;X-Outlook-MCP-Tools&quot;: &quot;calendar_list,account_info&quot;" in response.text


def test_http_auth_page_rejects_unknown_mode(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars?mode=bad")

    assert response.status_code == 400
    assert "validation_error" in response.text
    assert "Unknown MCP tools mode: bad" in response.text


def test_http_invalid_endpoint_get_returns_help_page(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/bad/path")

    assert response.status_code == 404
    assert "<style>" in response.text
    assert "invalid_endpoint" in response.text
    assert "/{method}/{account}" in response.text
    assert "/graph/you@outlook.com" in response.text
    assert "http://testserver/graph/you@outlook.com" in response.text
    assert "window.location.origin" in response.text
    assert 'data-endpoint-path="/graph/you@outlook.com"' in response.text
    assert "streamable-http" in response.text
    assert "outlook-http" in response.text


def test_http_invalid_endpoint_json_returns_help_payload(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/bad/path", headers={"accept": "application/json"})

    assert response.status_code == 404
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "invalid_endpoint"
    assert payload["expected_path"] == "/{method}/{account}"
    assert payload["example_url"] == "http://testserver/graph/you@outlook.com"
    assert payload["mcp_json"]["mcpServers"]["outlook-http"]["type"] == "streamable-http"
    assert payload["mcp_json"]["mcpServers"]["outlook-http"]["url"] == payload["example_url"]
    assert payload["mcp_json"]["mcpServers"]["outlook-http"]["headers"]["Authorization"] == "Bearer example-token-not-registered"
    assert payload["example_url"] in payload["agent_prompt"]


def test_http_forwarded_headers_build_external_url(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "1")
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars",
            headers={
                "x-forwarded-proto": "https",
                "x-forwarded-host": "mcp.example.com",
                "x-forwarded-port": "443",
                "x-forwarded-prefix": "/outlook",
            },
        )

    assert response.status_code == 200
    assert "https://mcp.example.com/outlook/graph/person@outlook.com" in response.text
    assert 'action="/outlook/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel"' in response.text
    assert 'action="/outlook/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize"' in response.text


def test_http_forwarded_port_is_added_when_non_default(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "1")
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars",
            headers={
                "x-forwarded-proto": "https",
                "x-forwarded-host": "mcp.example.com",
                "x-forwarded-port": "8443",
            },
        )

    assert response.status_code == 200
    assert "https://mcp.example.com:8443/graph/person@outlook.com" in response.text


def test_calendar_page_uses_local_assets_and_bearer_api(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "1")
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars/calendar",
            headers={
                "x-forwarded-proto": "https",
                "x-forwarded-host": "mcp.example.com",
                "x-forwarded-port": "443",
                "x-forwarded-prefix": "/outlook",
            },
        )

    assert response.status_code == 200
    assert "cdn.jsdelivr" not in response.text
    assert 'window.__OUTLOOK_CALENDAR_BOOTSTRAP__ = {"provider": "graph"' in response.text
    assert '"calendarTimezone": "Asia/Shanghai"' in response.text
    assert '"/outlook/graph/person@outlook.com/api"' in response.text
    assert '"/outlook/graph/person@outlook.com/test-secret-at-least-24-chars/api"' not in response.text
    assert 'src="/outlook/static/fullcalendar-index.global.min.js"' in response.text
    assert 'src="/outlook/static/fullcalendar-zh-cn.global.min.js"' in response.text
    assert 'src="/outlook/static/tinymce/tinymce.min.js"' in response.text
    assert 'src="/outlook/static/calendar.js"' in response.text
    assert 'rel="icon" type="image/svg+xml" href="/outlook/static/outlook-mcp-icon.svg"' in response.text
    assert '<meta name="theme-color" content="#1459c7">' in response.text
    assert "Graph 日历" in response.text
    assert "新建日程" in response.text
    assert 'id="panel-mode-label"' not in response.text
    assert "刷新" in response.text
    assert "等待加载最新日程" in response.text
    assert "会议与提醒" in response.text
    assert 'name="reminder_minutes_before_start" min="0" max="10080" step="1"' in response.text
    assert 'name="reminder_minutes_before_start" value="15"' not in response.text
    assert 'id="panel-loading"' in response.text
    assert "正在读取日程详情..." in response.text
    assert 'class="cal-panel__header-actions"' in response.text
    assert 'id="remote-update-banner"' in response.text
    assert 'id="btn-apply-remote-detail"' in response.text
    assert "覆盖为远端" in response.text
    assert 'id="btn-save"' in response.text
    assert 'id="btn-save-draft"' in response.text
    assert 'id="btn-unlock-draft"' in response.text
    assert 'title="Ctrl+Enter"' in response.text
    assert 'title="Ctrl+S"' in response.text
    assert 'id="tab-calendar"' in response.text
    assert 'id="tab-drafts"' in response.text
    assert 'id="drafts-view"' in response.text
    assert 'id="drafts-list" role="list"' in response.text
    assert 'id="draft-retention-mode"' in response.text
    assert 'id="draft-retention-days"' in response.text
    assert 'class="datetime-row"' in response.text
    assert 'class="datetime-label">全天' in response.text
    assert 'id="event-start"' in response.text
    assert 'id="event-end"' in response.text
    assert "接下来" in response.text
    assert 'id="panel-day-agenda"' in response.text
    assert 'id="day-agenda-title"' in response.text
    assert 'id="calendar" aria-label="Outlook 日历"' in response.text
    assert 'id="editor-panel" aria-label="日历详情面板"' in response.text
    assert 'id="day-agenda-list" role="list" aria-labelledby="day-agenda-title"' in response.text
    assert 'id="btn-day-new-event"' in response.text
    assert 'id="btn-close-day-agenda"' in response.text
    assert 'id="description-textarea"' in response.text
    assert 'id="description-preview"' not in response.text
    assert 'id="btn-toggle-description-mode"' not in response.text
    assert 'class="cal-form__footer"' not in response.text
    assert 'Authorization": "Bearer' not in response.text


def test_mail_page_uses_local_assets_and_bearer_api(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "1")
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars/mail",
            headers={
                "x-forwarded-proto": "https",
                "x-forwarded-host": "mcp.example.com",
                "x-forwarded-port": "443",
                "x-forwarded-prefix": "/outlook",
            },
        )

    assert response.status_code == 200
    assert 'window.__OUTLOOK_MAIL_BOOTSTRAP__ = {"provider": "graph"' in response.text
    assert '"pageSize": 20' in response.text
    assert '"bodyMaxChars": 200000' in response.text
    assert '"/outlook/graph/person@outlook.com/api"' in response.text
    assert '"/outlook/graph/person@outlook.com/test-secret-at-least-24-chars/api"' not in response.text
    assert 'src="/outlook/static/tinymce/tinymce.min.js"' in response.text
    assert 'src="/outlook/static/mail.js?v=' in response.text
    assert 'href="/outlook/static/outlook-mcp-icon.svg"' in response.text
    assert "Outlook Mail" in response.text
    assert "按主题搜索" in response.text
    assert 'id="mail-workspace" data-mobile-view="list"' in response.text
    assert 'id="folder-pane" data-open="false"' in response.text
    assert 'id="btn-folders-open" aria-expanded="false"' in response.text
    assert 'id="btn-detail-back"' in response.text
    assert 'id="mail-body-frame"' in response.text
    assert 'sandbox="allow-popups"' in response.text
    assert 'referrerpolicy="no-referrer"' in response.text
    assert 'id="compose-panel" hidden role="dialog" aria-modal="true"' in response.text
    assert 'Authorization": "Bearer' not in response.text


def test_mail_page_and_api_bind_bearer_to_account(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        page_response = client.get(
            "/graph/other@outlook.com/test-secret-at-least-24-chars/mail"
        )
        api_response = client.post(
            "/graph/other@outlook.com/api/mail/list",
            json={"folder": "inbox", "limit": 20, "offset": 0},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert page_response.status_code == 404
    assert "credential_not_registered" in page_response.text
    assert api_response.status_code == 403
    assert api_response.json()["error"] == "auth_required"


def test_mail_page_serves_local_assets(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        script_response = client.get("/static/mail.js")
        style_response = client.get("/static/mail.css")

    assert script_response.status_code == 200
    assert "javascript" in script_response.headers["content-type"]
    assert "sanitizeMailHtml" in script_response.text
    assert "Content-Security-Policy" in script_response.text
    assert 'node.removeAttribute("src")' in script_response.text
    assert 'license_key: "gpl"' in script_response.text
    assert 'toolbar_mode: "wrap"' in script_response.text
    assert style_response.status_code == 200
    assert "text/css" in style_response.headers["content-type"]
    assert ".mail-workspace" in style_response.text
    assert ".mail-rail" in style_response.text
    assert "mobileView" in script_response.text
    assert "prefers-reduced-motion" in style_response.text
    assert ".compose-panel" in style_response.text


def test_calendar_page_serves_local_locale_asset(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/fullcalendar-zh-cn.global.min.js")

    assert response.status_code == 200
    assert "FullCalendar.globalLocales.push" in response.text
    assert 'code:"zh-cn"' in response.text


def test_calendar_page_serves_icon_asset(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/outlook-mcp-icon.svg")

    assert response.status_code == 200
    assert response.headers["content-type"].startswith("image/svg+xml")
    assert "<svg" in response.text
    assert "Outlook MCP" in response.text


def test_calendar_page_serves_nested_tinymce_asset(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/tinymce/themes/silver/theme.min.js")

    assert response.status_code == 200
    assert "silver.uistate" in response.text


def test_calendar_page_refreshes_expired_cache_and_preserves_all_day_dates(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/calendar.js")

    assert response.status_code == 200
    assert "点击刷新获取最新日程" in response.text
    assert "缓存较旧，正在后台刷新最新日程" in response.text
    assert "isExpired: Date.now() - parsed.fetched_at > CACHE_STALE_AFTER_MS" in response.text
    assert "refreshCalendar({ forceNetwork: true });" in response.text
    assert '"Authorization": "Bearer " + apiBearerToken()' in response.text
    assert "start: allDay ? calendarInputValue(startValue, true) : startValue" in response.text
    assert "end: allDay ? calendarInputValue(endValue, true) : endValue" in response.text
    assert "buildCreateDraftFromCalendarRange" in response.text
    assert "projectDateTimeValue(new Date())" in response.text
    assert "suggestedEndValue(startValue)" in response.text
    assert "reminder: \"\"" in response.text
    assert "reminder: 15" not in response.text
    assert "API_TIMEOUT_MS = 60 * 1000" in response.text
    assert "new AbortController()" in response.text
    assert "请求超过 60 秒未返回" in response.text
    assert "event.startStr" in response.text
    assert "event.endStr" in response.text
    assert "getUTCHours()" in response.text
    assert "tinymce.init({" in response.text
    assert 'license_key: "gpl"' in response.text
    assert 'apiCall("calendar/get"' in response.text
    assert "normalizeEditorHtml" in response.text
    assert "extractHtmlBodyFragment" in response.text
    assert "looksLikeHtmlDocument" in response.text
    assert "showDetailLoading" in response.text
    assert "detailLoadSeq" in response.text
    assert "DETAIL_CACHE_TTL_MS" in response.text
    assert "cachedDetailItem" in response.text
    assert "buildEditDraftFromEvent" in response.text
    assert "showRemoteUpdatedNotice" in response.text
    assert "applyRemoteDetailDraft" in response.text
    assert "远端详情已更新，可点击覆盖为远端。" in response.text
    assert "syncSelectedEventClass" in response.text
    assert "syncSelectedDateClass" in response.text
    assert "selectCalendarDate" in response.text
    assert "renderDayAgenda" in response.text
    assert "suppressNextSelect" in response.text
    assert "dateClick" in response.text
    assert "datesSet" in response.text
    assert "syncWeeklyNowIndicator" in response.text
    assert "scheduleWeeklyNowIndicatorSync" in response.text
    assert "nowIndicatorSyncFrame" in response.text
    assert "focusMode" in response.text
    assert "keyboardFocusDate" in response.text
    assert "keyboardFocusEventId" in response.text
    assert "focusDateAfterRender" in response.text
    assert "handleMonthGridKeydown" in response.text
    assert "handleDayAgendaKeydown" in response.text
    assert "handleDescriptionEditorKeydown" in response.text
    assert "requestSubmit" in response.text
    assert "aria-selected" in response.text
    assert "tabIndex = isActive ? 0 : -1" in response.text
    assert "eventResizableFromStart: true" in response.text
    assert "height: 190" in response.text
    assert "currentDescriptionHtml()" in response.text
    assert 'apiCall("calendar/drafts/save", params)' in response.text
    assert 'apiCall("calendar/drafts/submit"' in response.text
    assert 'apiCall("calendar/drafts/list"' in response.text
    assert "function isDraftShortcut(event)" in response.text
    assert "function isSubmitShortcut(event)" in response.text
    assert "function handleEditShortcut(event)" in response.text
    assert 'String(event.key || "").toLowerCase() !== "e"' in response.text
    assert 'id: "drafts", events: loadDraftEventInputs' in response.text
    assert 'title: "[草稿] " + (record.subject || "(无标题)")' in response.text
    assert 'classNames: ["cal-event-draft"]' in response.text
    assert 'stageCalendarMutation(info, "移动结果已暂存为编辑草稿。")' in response.text
    assert 'stageCalendarMutation(info, "时间调整已暂存为编辑草稿。")' in response.text
    assert 'apiCall("calendar/update"' not in response.text
    assert 'refs.panelTitle.textContent = state.currentDraftId' in response.text
    assert "useOriginalHtml" not in response.text
    assert "renderDescriptionPreview" not in response.text
    assert "sanitizePreviewNode" not in response.text
    assert "btn-toggle-description-mode" not in response.text
    assert "panelMode" not in response.text
    assert "scheduleBackgroundRefresh" not in response.text


def test_calendar_page_opens_native_picker_and_adjusts_invalid_end(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/calendar.js")

    assert response.status_code == 200
    assert "var HALF_HOUR_MINUTES = 30;" in response.text
    assert "function roundUpToHalfHour(value)" in response.text
    assert "function suggestedEndValue(startValue)" in response.text
    assert "function syncEndAfterStartChange()" in response.text
    assert 'refs.form.start.addEventListener("change", syncEndAfterStartChange);' in response.text
    assert 'refs.form.start.addEventListener("click", showNativePicker);' in response.text
    assert 'refs.form.end.addEventListener("click", showNativePicker);' in response.text
    assert 'typeof input.showPicker !== "function"' in response.text


def test_calendar_page_projects_mutations_into_all_cached_ranges(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/calendar.js")

    assert response.status_code == 200
    assert "var CACHE_VERSION = 3;" in response.text
    assert "range_start: rangeStart" in response.text
    assert "range_end: rangeEnd" in response.text
    assert "saveCachedRange(rangeStart, rangeEnd, eventInputs, !truncated)" in response.text
    assert "projectCalendarMutationIntoCachedRanges" in response.text
    assert "eventInputOverlapsRange" in response.text
    assert 'syncCalendarMutation("delete", data, deletedItemId)' in response.text
    assert 'syncCalendarMutation("upsert", calendarData, previousItemId)' in response.text
    assert "refreshCalendar({ forceNetwork: !cacheUpdated });" in response.text


def test_calendar_page_uses_adaptive_event_layout(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        script_response = client.get("/static/calendar.js")
        style_response = client.get("/static/calendar.css")

    assert script_response.status_code == 200
    assert "slotEventOverlap: false" in script_response.text
    assert "dayMaxEvents: true" in script_response.text
    assert "dayMaxEvents: 3" not in script_response.text
    assert style_response.status_code == 200
    assert ".fc .fc-daygrid-day .fc-daygrid-event .fc-event-title" in style_response.text
    assert ".fc .fc-timegrid-event .fc-event-title" in style_response.text
    assert ".fc .cal-event-draft" in style_response.text
    assert "border: 1px dashed var(--cal-draft)" in style_response.text
    assert ".cal-draft-row" in style_response.text
    assert "text-overflow: ellipsis;" in style_response.text
    assert "white-space: nowrap;" in style_response.text


def test_calendar_page_resets_fullcalendar_table_spacing(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/calendar.css")

    assert response.status_code == 200
    assert _normalize_line_endings(""".fc table {
  margin-bottom: 0;
}""") in _normalize_line_endings(response.text)


def test_calendar_page_serves_expired_cache_style(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/static/calendar.css")

    assert response.status_code == 200
    assert '.cal-cache-status[data-tone="expired"]' in response.text
    assert "font-weight: 600;" in response.text
    assert ".fc .fc-daygrid-event" in response.text
    assert ".fc .fc-daygrid-day.cal-day-selected" in response.text
    assert "--cal-selected-day-bg" in response.text
    assert ".cal-mobile-backdrop[hidden]" in response.text
    assert ".fc .fc-daygrid-dot-event" in response.text
    assert "font-size: 10px;" in response.text
    assert _normalize_line_endings(""".fc .fc-timegrid-slot {
  background-color: transparent;
}""") in _normalize_line_endings(response.text)
    assert "grid-template-columns: minmax(0, 1fr) clamp(440px, 34vw, 520px);" in response.text
    assert _normalize_line_endings(""".cal-upcoming-list {
  display: grid;
  flex: 1 1 auto;
  align-content: start;""") in _normalize_line_endings(response.text)
    assert ".cal-form__body > *" in response.text
    assert "padding: 8px 16px 10px;" in response.text
    assert "height: auto;" in response.text
    assert "min-height: 36px;" in response.text
    assert '.cal-form label > input[type="datetime-local"]' in response.text
    assert "grid-template-columns: minmax(76px, 0.45fr) minmax(0, 1fr) minmax(0, 1fr);" in response.text
    assert ".datetime-label" in response.text
    assert ".cal-form .cal-toggle" in response.text
    assert ".cal-panel__header-actions" in response.text
    assert ".cal-panel__loading-body" in response.text
    assert ".cal-sync-banner" in response.text
    assert '.cal-sync-banner[data-tone="updated"]' in response.text
    assert ".fc .cal-event-selected" in response.text
    assert ".fc .fc-timegrid-event.cal-event-selected .fc-event-resizer-start" in response.text
    assert "--cal-now-indicator-color" in response.text
    assert ".fc .fc-timegrid-now-indicator-container.cal-timegrid-now-indicator-expanded" in response.text
    assert "overflow: visible;" in response.text
    assert ".fc .fc-daygrid-day.cal-day-keyboard-active" in response.text
    assert ".fc .fc-daygrid-day:focus-visible" in response.text
    assert ".cal-day-agenda-item:focus-visible" in response.text
    assert "overflow-wrap: anywhere;" in response.text
    assert ".cal-description-field h4" in response.text
    assert ".cal-description-field .tox-tinymce" in response.text
    assert ".cal-richtext-preview" not in response.text
    assert "max-height: 132px;" in response.text
    assert ".cal-panel .cal-btn {" in response.text
    assert ".cal-form__footer" not in response.text


def test_calendar_page_rejects_non_graph_method(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/ews/person@example.com/test-secret-at-least-24-chars/calendar")

    assert response.status_code == 400
    assert "visual calendar editor is currently available only for Graph HTTP endpoints" in response.text


def test_mail_page_rejects_non_graph_method(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/ews/person@example.com/test-secret-at-least-24-chars/mail")

    assert response.status_code == 400
    assert "visual mail client is currently available only for Graph HTTP endpoints" in response.text


def test_http_forwarded_header_takes_precedence(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "1")
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars",
            headers={
                "forwarded": 'for=192.0.2.60;proto=https;host="forwarded.example.com"',
                "x-forwarded-proto": "http",
                "x-forwarded-host": "ignored.example.com",
            },
        )

    assert response.status_code == 200
    assert "https://forwarded.example.com/graph/person@outlook.com" in response.text


def test_http_forwarded_headers_enabled_by_forwarded_allow_ips(monkeypatch, tmp_path) -> None:
    monkeypatch.delenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", raising=False)
    monkeypatch.setenv("FORWARDED_ALLOW_IPS", "10.0.0.0/8")
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars",
            headers={
                "x-forwarded-proto": "https",
                "x-forwarded-host": "mcp.example.com",
            },
        )

    assert response.status_code == 200
    assert "https://mcp.example.com/graph/person@outlook.com" in response.text


def test_http_forwarded_headers_explicitly_disabled_even_with_forwarded_allow_ips(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_TRUST_PROXY_HEADERS", "0")
    monkeypatch.setenv("FORWARDED_ALLOW_IPS", "10.0.0.0/8")
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/person@outlook.com/test-secret-at-least-24-chars",
            headers={
                "forwarded": 'for=192.0.2.60;proto=https;host="spoofed.example.com"',
                "x-forwarded-proto": "https",
                "x-forwarded-host": "spoofed.example.com",
            },
        )

    assert response.status_code == 200
    assert "http://testserver/graph/person@outlook.com" in response.text
    assert "spoofed.example.com" not in response.text


def test_http_method_placeholder_returns_method_unsupported(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("ews", "person@example.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        client.get("/ews/person@example.com/test-secret-at-least-24-chars")
        response = client.post(
            "/ews/person@example.com",
            json={},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "method_unsupported"


def test_http_authorized_page_shows_auth_action_buttons(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars")

    assert response.status_code == 200
    assert "Outlook MCP Authorized" in response.text
    assert "取消授权" in response.text
    assert "重新授权" in response.text
    assert 'action="/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel"' in response.text
    assert 'action="/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize"' in response.text
    assert 'action="http://testserver' not in response.text


def test_http_cancel_auth_clears_only_current_endpoint(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    other = store.get_or_create_endpoint("graph", "other@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"cache": true}')
    store.save_pending_flow(record.id, {"user_code": "ABCD", "expires_at": 9999999999})
    store.save_msal_cache(other.id, '{"other": true}')
    store.save_pending_flow(other.id, {"user_code": "WXYZ", "expires_at": 9999999999})
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        auth_url = "http://testserver/graph/person@outlook.com/test-secret-at-least-24-chars"
        csrf_token = _make_csrf_token("cancel", auth_url)
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel", data={"csrf_token": csrf_token})

    assert response.status_code == 200
    assert "Authorization for this endpoint has been cleared." in response.text
    assert "重新授权" in response.text
    assert "取消授权" not in response.text
    assert 'action="/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize"' in response.text
    assert 'action="http://testserver' not in response.text
    assert "/auth/cancel/auth/reauthorize" not in response.text
    assert store.load_msal_cache(record.id) is None
    assert store.load_pending_flow(record.id) is None
    assert store.load_msal_cache(other.id) == '{"other": true}'
    assert store.load_pending_flow(other.id)["user_code"] == "WXYZ"


def test_http_reauthorize_clears_cache_and_starts_device_flow(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"cache": true}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        auth_url = "http://testserver/graph/person@outlook.com/test-secret-at-least-24-chars"
        csrf_token = _make_csrf_token("reauthorize", auth_url)
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize", data={"csrf_token": csrf_token})

    assert response.status_code == 200
    assert "Authorize Outlook MCP" in response.text
    assert "ABCD-EFGH" in response.text
    assert "http://testserver/graph/person@outlook.com" in response.text
    assert "/auth/reauthorize/auth/" not in response.text
    assert store.load_msal_cache(record.id) is None
    assert store.load_pending_flow(record.id)["user_code"] == "ABCD-EFGH"


def test_http_auth_actions_clear_cached_graph_client(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    created_clients: list[GraphClient] = []

    def client_factory(*args: Any, **kwargs: Any) -> GraphClient:
        client = GraphClient(*args, **kwargs)
        created_clients.append(client)
        return client

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    monkeypatch.setattr("outlook_mcp.server.GraphClient", client_factory)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        assert client.get("/graph/person@outlook.com/test-secret-at-least-24-chars").status_code == 200
        auth_url = "http://testserver/graph/person@outlook.com/test-secret-at-least-24-chars"
        csrf_token = _make_csrf_token("cancel", auth_url)
        assert client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel", data={"csrf_token": csrf_token}).status_code == 200
        assert client.get("/graph/person@outlook.com/test-secret-at-least-24-chars").status_code == 200

    assert len(created_clients) == 2


def test_http_auth_action_get_is_method_not_allowed(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        cancel = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel")
        reauthorize = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize")

    assert cancel.status_code == 405
    assert reauthorize.status_code == 405


def test_http_get_invalid_secret_returns_error_page(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/a%2Fb")

    assert response.status_code == 400
    assert "<style>" in response.text
    assert "invalid_secret" in response.text
    assert "secret may only contain" in response.text
    assert store.lookup_endpoint("graph", "person@outlook.com", "a/b") is None


def test_http_post_invalid_secret_returns_json_error(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/a%2Fb", json={})

    assert response.status_code == 400
    assert response.json()["error"] == "invalid_secret"
    assert "secret may only contain" in response.json()["message"]


def test_http_log_page_lists_operation_logs_with_css_and_custom_range(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.append_operation_log(
        record.id,
        operation="mail_list",
        http_method="POST",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=12.5,
        created_at=datetime(2026, 1, 1, 12, tzinfo=timezone.utc).timestamp(),
    )
    store.append_operation_log(
        record.id,
        operation="calendar_create",
        http_method="POST",
        status_code=400,
        ok=False,
        error="validation_error",
        duration_ms=20,
        created_at=datetime(2025, 1, 1, 12, tzinfo=timezone.utc).timestamp(),
    )
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars/log?range=custom&from=2026-01-01&to=2026-01-02")

    assert response.status_code == 200
    assert "<style>" in response.text
    assert "Outlook MCP Operation Logs" in response.text
    assert "mail_list" in response.text
    assert "calendar_create" not in response.text
    assert "custom" in response.text


def test_http_mcp_request_is_logged_without_secret_or_body(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        client.get("/graph/person@outlook.com/test-secret-at-least-24-chars")
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/call", "params": {"name": "mail_list", "arguments": {"secret": "do-not-log"}}},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )
        log_page = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars/log")

    assert response.status_code == 200
    logs = store.list_operation_logs(store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars").id)
    assert logs[0]["operation"] == "mail_list"
    assert logs[0]["ok"] is False
    assert logs[0]["error"] == "auth_required"
    assert "mail_list" in log_page.text
    assert "test-secret-at-least-24-chars" not in log_page.text
    assert "do-not-log" not in log_page.text


def test_http_mcp_tools_header_is_stored_in_request_context(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_context_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Tools": "calendar_list, account_info",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["data"]["enabled_tools"] == ["account_info", "calendar_list"]


def test_http_mcp_mode_header_is_stored_in_request_context(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_context_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Mode": "calendar",
            },
        )

    assert response.status_code == 200
    payload = response.json()
    assert payload["ok"] is True
    assert payload["data"]["enabled_tools"] == [
        "account_info",
        "calendar_create",
        "calendar_delete",
        "calendar_get",
        "calendar_list",
        "calendar_search",
        "calendar_update",
    ]


def test_http_mcp_empty_custom_mode_disables_all_tools(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
    )
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_context_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Mode": "custom",
            },
        )

    assert response.status_code == 200
    assert response.json()["data"]["enabled_tools"] == []


def test_http_mcp_tools_header_rejects_unknown_tool(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Tools": "calendar_list,unknown_tool",
            },
        )

    assert response.status_code == 400
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "validation_error"
    assert "unknown_tool" in payload["message"]


def test_http_mcp_mode_header_rejects_unknown_mode(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Mode": "bad",
            },
        )

    assert response.status_code == 400
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "validation_error"
    assert "Unknown MCP tools mode: bad" in payload["message"]


def test_http_mcp_scope_is_rewritten_to_root_path(tmp_path) -> None:
    captured: dict[str, Any] = {}

    async def recording_mcp_app(scope: dict[str, Any], receive: Any, send: Any) -> None:
        if scope["type"] == "lifespan":
            await _lifespan(receive, send)
            return
        captured["path"] = scope["path"]
        captured["raw_path"] = scope["raw_path"]
        captured["method"] = scope["method"]
        captured["headers"] = dict(scope["headers"])
        while True:
            message = await receive()
            if message["type"] == "http.request" and not message.get("more_body"):
                break
        await send(
            {
                "type": "http.response.start",
                "status": 200,
                "headers": [(b"content-type", b"application/json")],
            }
        )
        await send({"type": "http.response.body", "body": b"{}"})

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = recording_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com",
            json={"method": "tools/list"},
            headers={"Authorization": "Bearer test-secret-at-least-24-chars"},
        )

    assert response.status_code == 200
    assert captured["path"] == "/"
    assert captured["raw_path"] == b"/"
    assert captured["method"] == "POST"
    assert captured["headers"][b"authorization"] == b"Bearer test-secret-at-least-24-chars"
    assert captured["headers"][b"host"] == b"testserver"


def test_http_api_respects_tools_header(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/api/mail/list",
            json={"limit": 20, "offset": 0},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Tools": "calendar_list",
            },
        )

    assert response.status_code == 400
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "tool_disabled"
    assert "mail_list" in payload["message"]


def test_http_api_respects_mode_header(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.save_msal_cache(record.id, '{"AccessToken": {}}')
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/api/mail/list",
            json={"limit": 20, "offset": 0},
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "X-Outlook-MCP-Mode": "calendar",
            },
        )

    assert response.status_code == 400
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "tool_disabled"
    assert "mail_list" in payload["message"]


async def _fake_mcp_app(scope: dict[str, Any], receive: Any, send: Any) -> None:
    if scope["type"] == "lifespan":
        await _lifespan(receive, send)
        return
    while True:
        message = await receive()
        if message["type"] == "http.request" and not message.get("more_body"):
            break
    payload = mail_tools.mail_list(server.get_client())
    body = json.dumps(payload).encode("utf-8")
    await send(
        {
            "type": "http.response.start",
            "status": 200,
            "headers": [(b"content-type", b"application/json")],
        }
    )
    await send({"type": "http.response.body", "body": body})


async def _fake_context_mcp_app(scope: dict[str, Any], receive: Any, send: Any) -> None:
    if scope["type"] == "lifespan":
        await _lifespan(receive, send)
        return
    while True:
        message = await receive()
        if message["type"] == "http.request" and not message.get("more_body"):
            break
    context = server._http_context.get()
    payload = {
        "ok": context is not None,
        "data": {
            "method": context.method if context else None,
            "account": context.account if context else None,
            "credential_id": context.credential_id if context else None,
            "auth_url": context.auth_url if context else None,
            "enabled_tools": sorted(context.enabled_tools) if context and context.enabled_tools is not None else None,
        },
    }
    body = json.dumps(payload).encode("utf-8")
    await send(
        {
            "type": "http.response.start",
            "status": 200,
            "headers": [(b"content-type", b"application/json")],
        }
    )
    await send({"type": "http.response.body", "body": body})


async def _lifespan(receive: Any, send: Any) -> None:
    while True:
        message = await receive()
        if message["type"] == "lifespan.startup":
            await send({"type": "lifespan.startup.complete"})
        elif message["type"] == "lifespan.shutdown":
            await send({"type": "lifespan.shutdown.complete"})
            return


def _patch_pending_device_flow(monkeypatch) -> None:
    class FakeCache:
        has_state_changed = False

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "message": "Use ABCD-EFGH",
            }

        def acquire_token_by_device_flow(self, flow: dict[str, Any], exit_condition: Any = None) -> dict[str, Any]:
            return {"error": "authorization_pending"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)


def _graph_config(tmp_path) -> GraphConfig:
    return GraphConfig(
        email="configured@outlook.com",
        client_id="client-id",
        token_cache_path=str(tmp_path / "ignored.json"),
    )


def test_request_response_capture_truncates_oversized_body() -> None:
    import asyncio

    big_body = b"x" * (_RequestResponseCapture._MAX_BODY_BYTES + 1024)

    async def _fake_receive():
        return {"type": "http.request", "body": big_body, "more_body": False}

    async def _fake_send(_msg):
        pass

    capture = _RequestResponseCapture(_fake_receive, _fake_send)
    msg = asyncio.run(capture.receive())
    assert msg["body"] == big_body
    assert len(capture.request_body) == _RequestResponseCapture._MAX_BODY_BYTES


def test_request_response_capture_does_not_truncate_small_body() -> None:
    import asyncio

    small_body = b"hello"

    async def _fake_receive():
        return {"type": "http.request", "body": small_body, "more_body": False}

    async def _fake_send(_msg):
        pass

    capture = _RequestResponseCapture(_fake_receive, _fake_send)
    msg = asyncio.run(capture.receive())
    assert msg["body"] == small_body
    assert bytes(capture.request_body) == small_body


def test_http_auth_cancel_rejects_missing_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel")

    assert response.status_code == 403
    assert "csrf_error" in response.text
    assert 'href="/graph/person@outlook.com/test-secret-at-least-24-chars"' in response.text
    assert 'href="http://testserver' not in response.text


def test_http_auth_cancel_rejects_invalid_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel", data={"csrf_token": "invalid:token"})

    assert response.status_code == 403
    assert "csrf_error" in response.text


def test_http_auth_cancel_rejects_wrong_action_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        auth_url = "http://testserver/graph/person@outlook.com/test-secret-at-least-24-chars"
        reauth_token = _make_csrf_token("reauthorize", auth_url)
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/cancel", data={"csrf_token": reauth_token})

    assert response.status_code == 403
    assert "csrf_error" in response.text


def test_http_auth_forms_include_csrf_tokens(monkeypatch, tmp_path) -> None:
    _patch_pending_device_flow(monkeypatch)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get("/graph/person@outlook.com/test-secret-at-least-24-chars")

    assert response.status_code == 200
    assert 'name="csrf_token"' in response.text
    assert "csrf_error" not in response.text


def test_read_request_body_rejects_disconnect() -> None:
    import asyncio

    async def _fake_receive():
        return {"type": "http.disconnect"}

    with pytest.raises(Exception, match="request disconnected"):
        asyncio.run(_read_request_body(_fake_receive))


# --- _RequestResponseCapture additional tests ---

def test_request_response_capture_truncates_oversized_response_body() -> None:
    import asyncio

    async def _fake_receive():
        return {"type": "http.request", "body": b"{}", "more_body": False}

    sent_messages: list = []

    async def _fake_send(msg):
        sent_messages.append(msg)

    capture = _RequestResponseCapture(_fake_receive, _fake_send)
    asyncio.run(capture.receive())
    asyncio.run(capture.send({"type": "http.response.start", "status": 200}))
    big_body = b"y" * 70000
    asyncio.run(capture.send({"type": "http.response.body", "body": big_body}))

    assert len(capture.response_body) == _RequestResponseCapture._MAX_BODY_BYTES
    assert capture.status_code == 200
    assert sent_messages[1]["body"] == big_body


def test_request_response_capture_multi_chunk_truncation() -> None:
    import asyncio

    chunk_a = b"x" * 61440
    chunk_b = b"z" * 10240
    chunks = [chunk_a, chunk_b]
    call_count = 0

    async def _fake_receive():
        nonlocal call_count
        idx = call_count
        call_count += 1
        more = idx < len(chunks) - 1
        return {"type": "http.request", "body": chunks[idx], "more_body": more}

    async def _fake_send(_msg):
        pass

    capture = _RequestResponseCapture(_fake_receive, _fake_send)
    asyncio.run(capture.receive())
    asyncio.run(capture.receive())

    assert len(capture.request_body) == _RequestResponseCapture._MAX_BODY_BYTES


def test_request_response_capture_full_cycle() -> None:
    import asyncio

    async def _fake_receive():
        return {"type": "http.request", "body": b'{"method":"tools/call"}', "more_body": False}

    async def _fake_send(_msg):
        pass

    capture = _RequestResponseCapture(_fake_receive, _fake_send)
    asyncio.run(capture.receive())
    asyncio.run(capture.send({"type": "http.response.start", "status": 201}))
    asyncio.run(capture.send({"type": "http.response.body", "body": b'{"result":"ok"}'}))

    assert b'tools/call' in bytes(capture.request_body)
    assert b'"result":"ok"' in bytes(capture.response_body)
    assert capture.status_code == 201


# --- CSRF reauthorize error paths ---

def test_http_reauthorize_rejects_missing_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize")

    assert response.status_code == 403
    assert "csrf_error" in response.text


def test_http_reauthorize_rejects_invalid_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize", data={"csrf_token": "invalid:token"})

    assert response.status_code == 403
    assert "csrf_error" in response.text


def test_http_reauthorize_rejects_wrong_action_csrf_token(monkeypatch, tmp_path) -> None:
    class ReadyTokenProvider:
        def get_access_token(self) -> str:
            return "access-token"

    monkeypatch.setattr("outlook_mcp.server.GraphTokenProvider", lambda *args, **kwargs: ReadyTokenProvider())
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        auth_url = "http://testserver/graph/person@outlook.com/test-secret-at-least-24-chars"
        cancel_token = _make_csrf_token("cancel", auth_url)
        response = client.post("/graph/person@outlook.com/test-secret-at-least-24-chars/auth/reauthorize", data={"csrf_token": cancel_token})

    assert response.status_code == 403
    assert "csrf_error" in response.text


# --- _verify_csrf_token unit tests ---

def test_csrf_token_expired_rejection() -> None:
    assert _verify_csrf_token("0:anymac", "cancel", "http://testserver/path") is False


def test_verify_csrf_token_edge_cases() -> None:
    path = "http://testserver/graph/p@o.com/secret"
    assert _verify_csrf_token("noccolon", "cancel", path) is False
    assert _verify_csrf_token(":macvalue", "cancel", path) is False
    assert _verify_csrf_token("-1:macvalue", "cancel", path) is False
    assert _verify_csrf_token("12dot5:macvalue", "cancel", path) is False
    assert _verify_csrf_token("", "cancel", path) is False


# --- _operation_* helper tests ---

def test_operation_from_request_body() -> None:
    payload = json.dumps({"method": "tools/call", "params": {"name": "mail_list"}})
    assert _operation_from_request_body(payload.encode()) == "mail_list"

    payload2 = json.dumps({"method": "tools/list"})
    assert _operation_from_request_body(payload2.encode()) == "tools/list"

    list_payload = json.dumps([{"method": "tools/call", "params": {"name": "first_op"}}, {}])
    assert _operation_from_request_body(list_payload.encode()) == "first_op"

    assert _operation_from_request_body(b"not json") == "mcp_request"

    assert _operation_from_request_body(json.dumps({}).encode()) == "mcp_request"


def test_operation_ok() -> None:
    assert _operation_ok(200, b'{"ok":true}', None) is True
    assert _operation_ok(200, b'{"ok":true}', "SomeError") is False
    assert _operation_ok(400, b'{"ok":true}', None) is False
    assert _operation_ok(200, b'{"ok":false}', None) is False
    assert _operation_ok(200, b'{"error":"bad"}', None) is False
    nested_fail = json.dumps({"ok": False})
    nested_ok = json.dumps({"ok": True})
    wrapped_fail = json.dumps({"result": {"content": [{"type": "text", "text": nested_fail}]}})
    wrapped_ok = json.dumps({"result": {"content": [{"type": "text", "text": nested_ok}]}})
    assert _operation_ok(200, wrapped_fail.encode(), None) is False
    assert _operation_ok(200, wrapped_ok.encode(), None) is True


def test_operation_error() -> None:
    assert _operation_error(200, b'{"ok":true}') is None
    assert _operation_error(200, b'{"error":"auth_required"}') == "auth_required"
    assert _operation_error(200, b'{"error":{"code":123}}') == "123"
    assert _operation_error(200, b'{"error":{"message":"bad"}}') == "bad"
    assert _operation_error(200, b'{"error":{}}') == "jsonrpc_error"
    nested_fail = json.dumps({"ok": False, "error": "inner"})
    wrapped_fail = json.dumps({"result": {"content": [{"type": "text", "text": nested_fail}]}})
    assert _operation_error(200, wrapped_fail.encode()) == "inner"
    assert _operation_error(500, b'not json') == "http_500"
    assert _operation_error(200, b'not json') is None


# --- _log_filter tests ---

def _fake_scope(query_string: str) -> dict[str, Any]:
    return {"query_string": query_string.encode()}


def test_log_filter_all_range() -> None:
    selected, start_at, end_at, _from, _to, warning = _log_filter(_fake_scope("range=all"))
    assert selected == "all"
    assert start_at is None
    assert end_at is None


def test_log_filter_24h_range() -> None:
    selected, start_at, end_at, _from, _to, warning = _log_filter(_fake_scope("range=24h"))
    assert selected == "24h"
    assert start_at is not None
    assert end_at is None


def test_log_filter_custom_range() -> None:
    selected, start_at, end_at, _from, _to, warning = _log_filter(_fake_scope("range=custom&from=2026-01-01&to=2026-01-31"))
    assert selected == "custom"
    assert start_at is not None
    assert end_at is not None
    assert warning is None


def test_log_filter_custom_invalid_date() -> None:
    selected, start_at, end_at, _from, _to, warning = _log_filter(_fake_scope("range=custom&from=bad"))
    assert selected == "all"
    assert warning is not None


def test_log_filter_invalid_range_defaults_to_all() -> None:
    selected, start_at, end_at, _from, _to, warning = _log_filter(_fake_scope("range=bogus"))
    assert selected == "all"


def test_date_start_timestamp_parses_iso_date() -> None:
    ts = _date_start_timestamp("2026-06-15")
    assert ts > 0


def test_date_start_timestamp_with_add_days() -> None:
    ts1 = _date_start_timestamp("2026-06-15", add_days=0)
    ts2 = _date_start_timestamp("2026-06-15", add_days=1)
    assert ts2 - ts1 == 86400.0


def test_wants_mcp_stream_detects_event_stream() -> None:
    assert _wants_mcp_stream({"headers": [(b"accept", b"text/event-stream")]}) is True


def test_wants_mcp_stream_detects_application_json() -> None:
    assert _wants_mcp_stream({"headers": [(b"accept", b"application/json")]}) is True


def test_wants_mcp_stream_rejects_plain_html() -> None:
    assert _wants_mcp_stream({"headers": [(b"accept", b"text/html")]}) is False


def test_wants_mcp_stream_handles_missing_accept_header() -> None:
    assert _wants_mcp_stream({"headers": []}) is False


def test_forwarded_values_parses_host_and_proto() -> None:
    result = _forwarded_values("for=192.168.1.1;proto=https;host=example.com")
    assert result == {"for": "192.168.1.1", "proto": "https", "host": "example.com"}


def test_forwarded_values_picks_first_entry() -> None:
    result = _forwarded_values("host=first.com, host=second.com")
    assert result == {"host": "first.com"}


def test_forwarded_values_empty_header() -> None:
    assert _forwarded_values("") == {}


def test_first_header_value_returns_first_comma_separated() -> None:
    assert _first_header_value("a, b, c") == "a"
    assert _first_header_value("only") == "only"
    assert _first_header_value("") == ""
    assert _first_header_value(None) == ""


def test_host_has_port_detection() -> None:
    assert _host_has_port("example.com:8080") is True
    assert _host_has_port("example.com") is False
    assert _host_has_port("[::1]:443") is True
    assert _host_has_port("[::1]") is False


def test_is_default_port() -> None:
    assert _is_default_port("https", "443") is True
    assert _is_default_port("http", "80") is True
    assert _is_default_port("https", "8443") is False
    assert _is_default_port("http", "8080") is False


def test_normalized_prefix() -> None:
    assert _normalized_prefix("/api") == "/api"
    assert _normalized_prefix("/api/") == "/api"
    assert _normalized_prefix("/") == ""
    assert _normalized_prefix("") == ""
    assert _normalized_prefix("  ") == ""


def test_api_dispatch_auto_generated_routes_exist() -> None:
    from outlook_mcp.tools import get_api_dispatch
    from outlook_mcp.http_server import _do_dispatch

    api_dispatch = get_api_dispatch()
    # Every expected service area should have dispatch entries
    assert api_dispatch.get("mail/list") is not None
    assert api_dispatch.get("mail/get") is not None
    assert api_dispatch.get("mail/forward") is not None
    assert api_dispatch.get("calendar/list") is not None
    assert api_dispatch.get("calendar/search") is not None
    assert api_dispatch.get("contact/get") is not None
    assert api_dispatch.get("task/create") is not None
    assert api_dispatch.get("account/info") is not None

    # Unknown paths return not_found
    result = _do_dispatch("mail/unknown_op", None, {})
    assert result == {"ok": False, "error": "not_found", "message": "Unknown API path: mail/unknown_op"}


def test_api_dispatch_empty_params_still_routes() -> None:
    from outlook_mcp.tools import get_api_dispatch

    api_dispatch = get_api_dispatch()
    assert len(api_dispatch) == 31
    # check that every entry maps to a callable tool function
    for path, tool_func in api_dispatch.items():
        assert callable(tool_func)
        assert "/" in path


def test_upload_rate_rejection_advances_refill_timestamp(monkeypatch) -> None:
    _UPLOAD_BUCKETS.clear()
    _UPLOAD_BUCKETS[42] = (0.0, 0.0)
    times = iter((3.0, 4.0))
    monkeypatch.setattr("outlook_mcp.http_server.time.monotonic", lambda: next(times))

    assert _check_upload_rate(42) is False
    assert _check_upload_rate(42) is False
    assert _UPLOAD_BUCKETS[42] == pytest.approx((2.0 / 3.0, 4.0))
    _UPLOAD_BUCKETS.clear()


def test_upload_rate_cleanup_evicts_real_oldest_bucket(monkeypatch) -> None:
    _UPLOAD_BUCKETS.clear()
    _UPLOAD_BUCKETS.update({bucket_id: (1.0, 800.0) for bucket_id in range(1, 1001)})
    _UPLOAD_BUCKETS[1] = (1.0, 900.0)
    _UPLOAD_BUCKETS[999] = (1.0, 100.0)
    monkeypatch.setattr("outlook_mcp.http_server.time.monotonic", lambda: 1000.0)

    assert _check_upload_rate(2000) is True
    assert 999 not in _UPLOAD_BUCKETS
    assert 1 in _UPLOAD_BUCKETS
    _UPLOAD_BUCKETS.clear()


def test_http_api_oversized_body_returns_validation_error(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app, raise_server_exceptions=False) as client:
        response = client.post(
            "/graph/person@outlook.com/api/mail/list",
            content=b"x" * (_MAX_API_BODY_BYTES + 1),
            headers={
                "Authorization": "Bearer test-secret-at-least-24-chars",
                "content-type": "application/json",
            },
        )

    assert response.status_code == 400
    payload = response.json()
    assert payload["ok"] is False
    assert payload["error"] == "validation_error"
    assert payload["message"] == "request body too large"


def test_http_api_rate_limits_repeated_auth_failures(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    statuses: list[int] = []
    with TestClient(app) as client:
        for _ in range(12):
            response = client.post(
                "/graph/person@outlook.com/api/mail/list",
                headers={"Authorization": "Bearer wrong-secret-000000"},
                json={},
            )
            statuses.append(response.status_code)

    assert statuses[-1] == 429
    assert response.json()["error"] == "rate_limited"
    assert 429 not in statuses[:10]


def test_http_endpoint_rejects_non_email_account(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.get(
            "/graph/not-an-email/some-secret-123456",
            headers={"accept": "application/json"},
        )

    assert response.status_code == 404
    assert response.json()["error"] == "invalid_endpoint"


def test_http_secret_scoped_uploads_path_is_not_routed(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    app = create_app(store=store, base_config=_graph_config(tmp_path))
    app.mcp_app = _fake_mcp_app

    with TestClient(app) as client:
        response = client.post(
            "/graph/person@outlook.com/some-secret-123456/uploads",
            headers={"accept": "application/json"},
        )

    assert response.status_code == 404
    assert response.json()["error"] == "invalid_endpoint"


@pytest.mark.parametrize(
    ("api_path", "tool_name", "params"),
    [
        ("contact/list", "contacts_list", {"limit": 20}),
        ("contact/search", "contacts_search", {"query": "ada"}),
        ("task/list", "tasks_list", {"limit": 20}),
    ],
)
def test_http_api_contact_task_paths_respect_enabled_tools(api_path, tool_name, params) -> None:
    result = _dispatch_api(api_path, object(), params, enabled_tools=frozenset({tool_name}))
    assert result.get("error") != "tool_disabled"

    disabled = _dispatch_api(api_path, object(), params, enabled_tools=frozenset({"mail_list"}))
    assert disabled["ok"] is False
    assert disabled["error"] == "tool_disabled"
    assert tool_name in disabled["message"]


def test_http_api_unknown_parameter_returns_validation_error() -> None:
    result = _dispatch_api("mail/list", object(), {"limit": 20, "bogus": 1})
    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "bogus" in result["message"]
    assert _result_status_code(dict(result)) == 400
