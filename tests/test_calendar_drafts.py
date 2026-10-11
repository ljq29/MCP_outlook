from __future__ import annotations

from typing import Any

from outlook_mcp.calendar_drafts import dispatch_calendar_draft_api
from outlook_mcp.credential_store import CalendarDraftConflictError, CredentialStore
from outlook_mcp.graph_client import GraphClient


def _payload(**updates: Any) -> dict[str, Any]:
    payload = {
        "subject": "项目评审",
        "location": "会议室 A",
        "body": "<p>讨论草稿。</p>",
        "start": "2026-07-16T09:00",
        "end": "2026-07-16T10:00",
        "is_all_day": False,
        "timezone": "Asia/Shanghai",
        "send_invites": True,
        "required_attendees": ["required@example.com"],
        "optional_attendees": [],
        "reminder_minutes_before_start": 10,
    }
    payload.update(updates)
    return payload


def _store_and_credential(tmp_path):
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    return store, record


def _save(
    store: CredentialStore,
    credential_id: int,
    *,
    draft_type: str = "create",
    source_event_id: str | None = None,
    base_snapshot: dict[str, Any] | None = None,
    payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    result = dispatch_calendar_draft_api(
        "calendar/drafts/save",
        store,
        credential_id,
        {
            "confirm": True,
            "draft_type": draft_type,
            "source_event_id": source_event_id,
            "base_snapshot": base_snapshot,
            "payload": payload or _payload(),
        },
    )
    assert result["ok"] is True
    return result["data"]["draft"]


def test_draft_crud_allows_incomplete_payload_and_lists_range(tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id, payload=_payload(subject="", start="", end=""))
    assert draft["subject"] == ""
    assert draft["start_at"] is None

    all_result = dispatch_calendar_draft_api(
        "calendar/drafts/list", store, record.id, {"limit": 100, "offset": 0}
    )
    assert all_result["data"]["total"] == 1
    assert "payload" not in all_result["data"]["items"][0]
    detail_result = dispatch_calendar_draft_api(
        "calendar/drafts/get", store, record.id, {"draft_id": draft["id"]}
    )
    assert detail_result["data"]["draft"]["payload"]["subject"] == ""

    range_result = dispatch_calendar_draft_api(
        "calendar/drafts/list",
        store,
        record.id,
        {"start": "2026-07-01T00:00:00+08:00", "end": "2026-08-01T00:00:00+08:00"},
    )
    assert range_result["data"]["items"] == []

    deleted = dispatch_calendar_draft_api(
        "calendar/drafts/delete",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
    )
    assert deleted == {"ok": True, "data": {"deleted": True}}


def test_create_draft_submission_deletes_draft_on_success(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id)
    captured: dict[str, Any] = {}

    def fake_create(client: Any, **kwargs: Any) -> dict[str, Any]:
        captured.update(kwargs)
        return {"ok": True, "data": {"id": "event-1", "subject": kwargs["subject"]}}

    monkeypatch.setattr("outlook_mcp.calendar_drafts.calendar_create", fake_create)
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )

    assert result["ok"] is True
    assert result["data"]["calendar"]["id"] == "event-1"
    assert captured["confirm"] is True
    assert captured["start"] == "2026-07-16T09:00"
    assert store.get_calendar_draft(record.id, draft["id"]) is None


def test_outcome_unknown_is_kept_and_cannot_be_resubmitted(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id)
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_create",
        lambda *args, **kwargs: {
            "ok": False,
            "error": "outcome_unknown",
            "message": "Graph request outcome is unknown",
        },
    )

    first = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    kept = first["data"]["draft"]
    assert first["error"] == "outcome_unknown"
    assert kept["status"] == "outcome_unknown"

    second = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": kept["id"], "version": kept["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    assert second["error"] == "draft_state_conflict"

    resolved = dispatch_calendar_draft_api(
        "calendar/drafts/resolve-outcome",
        store,
        record.id,
        {"confirm": True, "draft_id": kept["id"], "version": kept["version"]},
    )
    assert resolved["data"]["draft"]["status"] == "active"
    assert resolved["data"]["draft"]["last_error_code"] is None


def test_update_draft_detects_remote_change_before_writing(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    base = {
        "id": "event-1",
        "changekey": "change-1",
        "subject": "原日程",
        "start": "2026-07-16T09:00:00+08:00",
        "end": "2026-07-16T10:00:00+08:00",
    }
    draft = _save(
        store,
        record.id,
        draft_type="update",
        source_event_id="event-1",
        base_snapshot=base,
    )
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_get",
        lambda *args, **kwargs: {
            "ok": True,
            "data": {**base, "changekey": "change-2", "subject": "远端已改"},
        },
    )

    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    assert result["error"] == "remote_conflict"
    assert result["_status_code"] == 409
    assert result["data"]["draft"]["status"] == "active"
    assert result["data"]["remote"]["changekey"] == "change-2"


def test_update_draft_submits_full_desired_state(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    base = {
        "id": "event-1",
        "changekey": "change-1",
        "subject": "原日程",
        "start": "2026-07-16T09:00:00+08:00",
        "end": "2026-07-16T10:00:00+08:00",
    }
    draft = _save(
        store,
        record.id,
        draft_type="update",
        source_event_id="event-1",
        base_snapshot=base,
        payload=_payload(subject="修改后的日程"),
    )
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_get",
        lambda *args, **kwargs: {"ok": True, "data": base},
    )
    captured: dict[str, Any] = {}

    def fake_update(client: Any, **kwargs: Any) -> dict[str, Any]:
        captured.update(kwargs)
        return {"ok": True, "data": {"id": "event-1", "subject": "修改后的日程"}}

    monkeypatch.setattr("outlook_mcp.calendar_drafts.calendar_update", fake_update)
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )

    assert result["ok"] is True
    assert captured["item_id"] == "event-1"
    assert captured["updates"]["subject"] == "修改后的日程"
    assert captured["send_invites"] is True


def test_retention_setting_requires_confirm_and_is_per_credential(tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    missing_confirm = dispatch_calendar_draft_api(
        "calendar/drafts/settings/update", store, record.id, {"retention_days": 30}
    )
    assert missing_confirm["error"] == "validation_error"

    updated = dispatch_calendar_draft_api(
        "calendar/drafts/settings/update",
        store,
        record.id,
        {"confirm": True, "retention_days": 30},
    )
    assert updated["data"]["retention_days"] == 30


def test_retention_preview_reports_deletions_before_update(monkeypatch, tmp_path) -> None:
    now = 1_800_000_000.0
    monkeypatch.setattr("outlook_mcp.credential_store.time.time", lambda: now)
    store, record = _store_and_credential(tmp_path)
    _save(store, record.id)
    now += 31 * 86400

    preview = dispatch_calendar_draft_api(
        "calendar/drafts/settings/preview",
        store,
        record.id,
        {"retention_days": 30},
    )
    assert preview["data"]["expired_count"] == 1
    assert dispatch_calendar_draft_api(
        "calendar/drafts/list", store, record.id, {}
    )["data"]["total"] == 1

    updated = dispatch_calendar_draft_api(
        "calendar/drafts/settings/update",
        store,
        record.id,
        {"confirm": True, "retention_days": 30},
    )
    assert updated["data"]["expired_count"] == 1


def test_all_day_draft_uses_exclusive_end_only_for_projection_and_submit(
    monkeypatch,
    tmp_path,
) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(
        store,
        record.id,
        payload=_payload(start="2026-07-16", end="2026-07-16", is_all_day=True),
    )
    assert draft["payload"]["end"] == "2026-07-16"
    assert draft["start_date"] == "2026-07-16"
    assert draft["end_date"] == "2026-07-17"
    captured: dict[str, Any] = {}

    def fake_create(client: Any, **kwargs: Any) -> dict[str, Any]:
        captured.update(kwargs)
        return {"ok": True, "data": {"id": "all-day-1"}}

    monkeypatch.setattr("outlook_mcp.calendar_drafts.calendar_create", fake_create)
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    assert result["ok"] is True
    assert captured["start"] == "2026-07-16"
    assert captured["end"] == "2026-07-17"


def test_invalid_final_submission_keeps_draft_active(tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id, payload=_payload(subject=""))
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    assert result["error"] == "validation_error"
    kept = store.get_calendar_draft(record.id, draft["id"])
    assert kept.status == "active"
    assert kept.version == draft["version"]


def test_missing_update_source_can_be_converted_to_create(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    base = {"id": "missing-1", "changekey": "change-1", "subject": "原日程"}
    draft = _save(
        store,
        record.id,
        draft_type="update",
        source_event_id="missing-1",
        base_snapshot=base,
    )
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_get",
        lambda *args, **kwargs: {
            "ok": False,
            "error": "outlook_error",
            "message": "Graph API error 404: ErrorItemNotFound",
        },
    )
    missing = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )
    assert missing["error"] == "source_missing"
    active = missing["data"]["draft"]

    converted = dispatch_calendar_draft_api(
        "calendar/drafts/convert-to-create",
        store,
        record.id,
        {"confirm": True, "draft_id": active["id"], "version": active["version"]},
    )
    assert converted["data"]["draft"]["draft_type"] == "create"
    assert converted["data"]["draft"]["source_event_id"] is None
    assert converted["data"]["draft"]["base_snapshot"] is None


# --- send_invites passthrough, confirm message style, submit auth order ---


def test_update_draft_submission_honors_stored_send_invites_false(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    base = {
        "id": "event-1",
        "changekey": "change-1",
        "subject": "原日程",
        "start": "2026-07-16T09:00:00+08:00",
        "end": "2026-07-16T10:00:00+08:00",
    }
    draft = _save(
        store,
        record.id,
        draft_type="update",
        source_event_id="event-1",
        base_snapshot=base,
        payload=_payload(send_invites=False),
    )
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_get",
        lambda *args, **kwargs: {"ok": True, "data": base},
    )
    captured: dict[str, Any] = {}

    def fake_update(client: Any, **kwargs: Any) -> dict[str, Any]:
        captured.update(kwargs)
        return {"ok": True, "data": {"id": "event-1"}}

    monkeypatch.setattr("outlook_mcp.calendar_drafts.calendar_update", fake_update)
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )

    assert result["ok"] is True
    assert captured["send_invites"] is False
    assert "send_invites" not in captured["updates"]


def test_create_draft_submission_honors_stored_send_invites_false(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id, payload=_payload(send_invites=False))
    captured: dict[str, Any] = {}

    def fake_create(client: Any, **kwargs: Any) -> dict[str, Any]:
        captured.update(kwargs)
        return {"ok": True, "data": {"id": "event-2"}}

    monkeypatch.setattr("outlook_mcp.calendar_drafts.calendar_create", fake_create)
    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )

    assert result["ok"] is True
    assert captured["send_invites"] is False


def test_confirm_error_message_matches_tools_style(tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)

    result = dispatch_calendar_draft_api(
        "calendar/drafts/save", store, record.id, {"payload": _payload()}
    )

    assert result["error"] == "validation_error"
    assert result["message"] == "calendar/drafts/save requires confirm=true"


def test_submit_without_client_returns_auth_required_before_confirm(tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)

    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"draft_id": "draft-1", "version": 1},
        client=None,
    )

    assert result["error"] == "auth_required"
    assert result["_status_code"] == 403


def test_normalize_payload_uses_shared_default_timezone() -> None:
    from outlook_mcp.calendar_drafts import DEFAULT_CALENDAR_TIMEZONE, _normalize_payload

    payload = _normalize_payload({"subject": "x"})

    assert payload["timezone"] == DEFAULT_CALENDAR_TIMEZONE == "Asia/Shanghai"


def test_submit_reports_pending_cleanup_when_local_delete_conflicts(monkeypatch, tmp_path) -> None:
    store, record = _store_and_credential(tmp_path)
    draft = _save(store, record.id)
    monkeypatch.setattr(
        "outlook_mcp.calendar_drafts.calendar_create",
        lambda client, **kwargs: {"ok": True, "data": {"id": "event-new"}},
    )

    def conflict(*args: Any, **kwargs: Any) -> None:
        raise CalendarDraftConflictError("version mismatch")

    monkeypatch.setattr(store, "delete_submitted_calendar_draft", conflict)

    result = dispatch_calendar_draft_api(
        "calendar/drafts/submit",
        store,
        record.id,
        {"confirm": True, "draft_id": draft["id"], "version": draft["version"]},
        client=GraphClient.__new__(GraphClient),
    )

    assert result["ok"] is True
    assert result["data"]["submitted"] is True
    assert result["data"]["draft_cleanup"] == "pending"
