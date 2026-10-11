"""Server-side calendar drafts for the HTTP WebUI."""

from __future__ import annotations

import hashlib
import json
import logging
from datetime import date, datetime, timedelta, timezone
from typing import Any
from uuid import uuid4
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .credential_store import (
    CalendarDraftConflictError,
    CalendarDraftNotFoundError,
    CalendarDraftRecord,
    CalendarDraftStateError,
    CredentialStore,
)
from .tools.calendar import calendar_create, calendar_get, calendar_update
from .tools.common import DEFAULT_CALENDAR_TIMEZONE


_DRAFT_PATH_PREFIX = "calendar/drafts/"
_MAX_SUBJECT_LENGTH = 256
_MAX_LOCATION_LENGTH = 256
_MAX_BODY_LENGTH = 900_000

_logger = logging.getLogger(__name__)


def is_calendar_draft_api(api_path: str) -> bool:
    return api_path.startswith(_DRAFT_PATH_PREFIX)


def dispatch_calendar_draft_api(
    api_path: str,
    store: CredentialStore,
    credential_id: int,
    params: dict[str, Any],
    *,
    client: Any = None,
    enabled_tools: frozenset[str] | None = None,
) -> dict[str, Any]:
    action = api_path[len(_DRAFT_PATH_PREFIX):]
    try:
        if action == "list":
            return _list_drafts(store, credential_id, params)
        if action == "get":
            record = _required_draft(store, credential_id, params.get("draft_id"))
            return _ok({"draft": serialize_calendar_draft(record)})
        if action == "save":
            _require_confirm(params, f"calendar/drafts/{action}")
            return _save_draft(store, credential_id, params)
        if action == "delete":
            _require_confirm(params, f"calendar/drafts/{action}")
            store.delete_calendar_draft(
                credential_id,
                _required_text(params.get("draft_id"), "draft_id"),
                _required_version(params.get("version")),
            )
            return _ok({"deleted": True})
        if action == "convert-to-create":
            _require_confirm(params, f"calendar/drafts/{action}")
            record = store.convert_calendar_draft_to_create(
                credential_id,
                _required_text(params.get("draft_id"), "draft_id"),
                _required_version(params.get("version")),
            )
            return _ok({"draft": serialize_calendar_draft(record)})
        if action == "resolve-outcome":
            _require_confirm(params, f"calendar/drafts/{action}")
            record = store.resolve_calendar_draft_outcome(
                credential_id,
                _required_text(params.get("draft_id"), "draft_id"),
                _required_version(params.get("version")),
            )
            return _ok({"draft": serialize_calendar_draft(record)})
        if action == "settings/get":
            return _ok(store.get_calendar_draft_settings(credential_id))
        if action == "settings/preview":
            retention_days = _retention_days(params.get("retention_days"))
            return _ok(store.preview_calendar_draft_retention(credential_id, retention_days))
        if action == "settings/update":
            _require_confirm(params, f"calendar/drafts/{action}")
            retention_days = _retention_days(params.get("retention_days"))
            return _ok(store.set_calendar_draft_retention(credential_id, retention_days))
        if action == "submit":
            if client is None:
                return _error("auth_required", "需要重新授权后才能提交草稿。", 403)
            _require_confirm(params, "calendar/drafts/submit")
            return _submit_draft(
                store,
                credential_id,
                params,
                client=client,
                enabled_tools=enabled_tools,
            )
        return _error("not_found", f"Unknown calendar draft API path: {api_path}", 404)
    except CalendarDraftNotFoundError:
        return _error("draft_not_found", "草稿不存在或已经过期。", 404)
    except CalendarDraftConflictError as exc:
        data = {} if exc.record is None else {"draft": serialize_calendar_draft(exc.record)}
        return _error("draft_conflict", "草稿已在另一设备更新，请重新加载。", 409, data)
    except CalendarDraftStateError as exc:
        return _error(
            "draft_state_conflict",
            f"草稿当前状态为 {exc.status}，不能执行此操作。",
            409,
        )
    except (TypeError, ValueError) as exc:
        return _error("validation_error", str(exc), 400)


def serialize_calendar_draft(
    record: CalendarDraftRecord,
    *,
    include_payload: bool = True,
) -> dict[str, Any]:
    data = {
        "id": record.id,
        "draft_type": record.draft_type,
        "source_event_id": record.source_event_id,
        "subject": record.subject,
        "is_all_day": record.is_all_day,
        "start_at": record.start_at,
        "end_at": record.end_at,
        "start_date": record.start_date,
        "end_date": record.end_date,
        "timezone": record.timezone,
        "status": record.status,
        "version": record.version,
        "expires_at": record.expires_at,
        "last_error_code": record.last_error_code,
        "last_error_message": record.last_error_message,
        "created_at": record.created_at,
        "updated_at": record.updated_at,
    }
    if include_payload:
        data["payload"] = record.payload
        data["base_snapshot"] = record.base_snapshot
    return data


def canonical_event_snapshot(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": str(item.get("id") or ""),
        "changekey": str(item.get("changekey") or ""),
        "subject": str(item.get("subject") or ""),
        "body": str(item.get("body") or ""),
        "start": item.get("start_local") or item.get("start") or "",
        "end": item.get("end_local") or item.get("end") or "",
        "location": str(item.get("location") or ""),
        "is_all_day": bool(item.get("is_all_day")),
        "is_cancelled": bool(item.get("is_cancelled")),
        "required_attendees": sorted(str(value) for value in item.get("required_attendees") or []),
        "optional_attendees": sorted(str(value) for value in item.get("optional_attendees") or []),
    }


def event_snapshot_fingerprint(snapshot: dict[str, Any] | None) -> str | None:
    if not snapshot:
        return None
    changekey = str(snapshot.get("changekey") or "")
    canonical = {"changekey": changekey} if changekey else canonical_event_snapshot(snapshot)
    encoded = json.dumps(canonical, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _list_drafts(
    store: CredentialStore,
    credential_id: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    limit = _bounded_int(params.get("limit", 100), "limit", 1, 100)
    offset = _bounded_int(params.get("offset", 0), "offset", 0, 1_000_000)
    range_start = params.get("start")
    range_end = params.get("end")
    start_at = end_at = None
    start_date = end_date = None
    if range_start is not None or range_end is not None:
        start_text = _required_text(range_start, "start")
        end_text = _required_text(range_end, "end")
        start_at = _parse_instant(start_text).timestamp()
        end_at = _parse_instant(end_text).timestamp()
        if end_at <= start_at:
            raise ValueError("end must be later than start")
        start_date = start_text[:10]
        end_date = end_text[:10]
    records, total = store.list_calendar_drafts(
        credential_id,
        limit=limit,
        offset=offset,
        range_start_at=start_at,
        range_end_at=end_at,
        range_start_date=start_date,
        range_end_date=end_date,
    )
    return _ok({
        "items": [serialize_calendar_draft(record, include_payload=False) for record in records],
        "total": total,
        "has_more": offset + len(records) < total,
    })


def _save_draft(
    store: CredentialStore,
    credential_id: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    draft_id = str(params.get("draft_id") or "").strip()
    expected_version = params.get("version")
    existing = store.get_calendar_draft(credential_id, draft_id) if draft_id else None
    if draft_id and existing is None:
        raise CalendarDraftNotFoundError(draft_id)
    if existing is not None:
        version = _required_version(expected_version)
        draft_type = existing.draft_type
        source_event_id = existing.source_event_id
        base_snapshot = existing.base_snapshot
    else:
        version = None
        draft_id = str(uuid4())
        draft_type = str(params.get("draft_type") or "create").strip()
        source_event_id = _optional_text(params.get("source_event_id"))
        base_snapshot_raw = params.get("base_snapshot")
        if base_snapshot_raw is not None and not isinstance(base_snapshot_raw, dict):
            raise ValueError("base_snapshot must be an object")
        base_snapshot = canonical_event_snapshot(base_snapshot_raw) if base_snapshot_raw else None
    if draft_type == "update" and not source_event_id:
        raise ValueError("source_event_id is required for update drafts")
    if draft_type == "update" and base_snapshot is None:
        raise ValueError("base_snapshot is required for update drafts")
    payload = _normalize_payload(params.get("payload"))
    projection = _draft_projection(payload)
    record = store.save_calendar_draft(
        credential_id,
        draft_id=draft_id,
        expected_version=version,
        draft_type=draft_type,
        source_event_id=source_event_id,
        payload=payload,
        base_snapshot=base_snapshot,
        base_fingerprint=event_snapshot_fingerprint(base_snapshot),
        subject=payload["subject"],
        timezone=payload["timezone"],
        **projection,
    )
    return _ok({"draft": serialize_calendar_draft(record)})


def _submit_draft(
    store: CredentialStore,
    credential_id: int,
    params: dict[str, Any],
    *,
    client: Any,
    enabled_tools: frozenset[str] | None,
) -> dict[str, Any]:
    draft_id = _required_text(params.get("draft_id"), "draft_id")
    version = _required_version(params.get("version"))
    record = _required_draft(store, credential_id, draft_id)
    tool_name = "calendar_create" if record.draft_type == "create" else "calendar_update"
    if enabled_tools is not None and tool_name not in enabled_tools:
        return _error(
            "tool_disabled",
            f"MCP tool is disabled by current MCP tool selection: {tool_name}",
            403,
        )
    create_payload = _create_payload(record.payload) if record.draft_type == "create" else None
    update_payload = _update_payload(record.payload) if record.draft_type == "update" else None
    update_send_invites = (
        bool(update_payload.pop("send_invites", True)) if update_payload is not None else True
    )
    submitting = store.begin_calendar_draft_submission(credential_id, draft_id, version)
    if submitting.draft_type == "update":
        remote_result = calendar_get(
            client,
            item_id=str(submitting.source_event_id),
            timezone=submitting.timezone,
        )
        if not remote_result.get("ok"):
            code = str(remote_result.get("error") or "source_unavailable")
            message = str(remote_result.get("message") or "无法读取原日程。")
            if _looks_not_found(message):
                code = "source_missing"
                message = "原日程已不存在，可将草稿转为新建日程。"
            active = store.finish_calendar_draft_submission(
                credential_id,
                draft_id,
                submitting.version,
                status="active",
                error_code=code,
                error_message=message,
            )
            return _error(code, message, 404 if code == "source_missing" else 400, {
                "draft": serialize_calendar_draft(active),
            })
        remote = canonical_event_snapshot(remote_result.get("data") or {})
        if event_snapshot_fingerprint(remote) != submitting.base_fingerprint:
            active = store.finish_calendar_draft_submission(
                credential_id,
                draft_id,
                submitting.version,
                status="active",
                error_code="remote_conflict",
                error_message="原日程已在 Outlook 中发生变化。",
            )
            return _error(
                "remote_conflict",
                "原日程已发生变化，请核对远端内容后再处理草稿。",
                409,
                {"draft": serialize_calendar_draft(active), "remote": remote},
            )
        result = calendar_update(
            client,
            item_id=str(submitting.source_event_id),
            updates=update_payload or {},
            send_invites=update_send_invites,
            confirm=True,
            timezone=submitting.timezone,
            html=True,
        )
    else:
        result = calendar_create(client, confirm=True, html=True, **(create_payload or {}))
    if result.get("ok"):
        cleanup_pending = False
        try:
            store.delete_submitted_calendar_draft(
                credential_id,
                draft_id,
                submitting.version,
            )
        except CalendarDraftConflictError:
            cleanup_pending = True
            _logger.warning(
                "Calendar draft %s for credential %s was submitted but local cleanup hit a conflict",
                draft_id,
                credential_id,
            )
        payload = {
            "draft_id": draft_id,
            "submitted": True,
            "calendar": result.get("data") or {},
        }
        if cleanup_pending:
            payload["draft_cleanup"] = "pending"
        return _ok(payload)
    error_code = str(result.get("error") or "calendar_submit_failed")
    error_message = str(result.get("message") or "提交草稿失败。")
    next_status = "outcome_unknown" if error_code == "outcome_unknown" else "active"
    saved = store.finish_calendar_draft_submission(
        credential_id,
        draft_id,
        submitting.version,
        status=next_status,
        error_code=error_code,
        error_message=error_message,
    )
    status_code = 409 if next_status == "outcome_unknown" else 400
    if error_code == "auth_required":
        status_code = 403
    return _error(error_code, error_message, status_code, {
        "draft": serialize_calendar_draft(saved),
        **({"auth": result.get("data")} if result.get("data") else {}),
    })


def _normalize_payload(raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise ValueError("payload must be an object")
    timezone_name = str(raw.get("timezone") or DEFAULT_CALENDAR_TIMEZONE).strip()
    try:
        ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError as exc:
        raise ValueError(f"unknown timezone: {timezone_name}") from exc
    subject = str(raw.get("subject") or "")
    location = str(raw.get("location") or "")
    body = str(raw.get("body") or "")
    if len(subject) > _MAX_SUBJECT_LENGTH:
        raise ValueError(f"subject must be at most {_MAX_SUBJECT_LENGTH} characters")
    if len(location) > _MAX_LOCATION_LENGTH:
        raise ValueError(f"location must be at most {_MAX_LOCATION_LENGTH} characters")
    if len(body) > _MAX_BODY_LENGTH:
        raise ValueError(f"body must be at most {_MAX_BODY_LENGTH} characters")
    return {
        "subject": subject,
        "location": location,
        "body": body,
        "start": str(raw.get("start") or ""),
        "end": str(raw.get("end") or ""),
        "is_all_day": bool(raw.get("is_all_day")),
        "timezone": timezone_name,
        "send_invites": bool(raw.get("send_invites", True)),
        "required_attendees": _string_list(raw.get("required_attendees")),
        "optional_attendees": _string_list(raw.get("optional_attendees")),
        "reminder_minutes_before_start": raw.get("reminder_minutes_before_start"),
    }


def _draft_projection(payload: dict[str, Any]) -> dict[str, Any]:
    if payload["is_all_day"]:
        try:
            start_date = date.fromisoformat(payload["start"])
            inclusive_end = date.fromisoformat(payload["end"])
        except ValueError:
            return {
                "is_all_day": True,
                "start_at": None,
                "end_at": None,
                "start_date": None,
                "end_date": None,
            }
        if inclusive_end < start_date:
            return {
                "is_all_day": True,
                "start_at": None,
                "end_at": None,
                "start_date": None,
                "end_date": None,
            }
        return {
            "is_all_day": True,
            "start_at": None,
            "end_at": None,
            "start_date": start_date.isoformat(),
            "end_date": (inclusive_end + timedelta(days=1)).isoformat(),
        }
    start = _parse_local_optional(payload["start"], payload["timezone"])
    end = _parse_local_optional(payload["end"], payload["timezone"])
    if start is None or end is None or end <= start:
        start = end = None
    return {
        "is_all_day": False,
        "start_at": None if start is None else start.timestamp(),
        "end_at": None if end is None else end.timestamp(),
        "start_date": None,
        "end_date": None,
    }


def _create_payload(payload: dict[str, Any]) -> dict[str, Any]:
    subject = payload["subject"].strip()
    if not subject:
        raise ValueError("subject is required")
    start, end = _submission_times(payload)
    reminder = payload.get("reminder_minutes_before_start")
    if reminder in {"", None}:
        reminder = None
    else:
        reminder = int(reminder)
    return {
        "subject": subject,
        "body": payload["body"],
        "start": start,
        "end": end,
        "location": payload["location"],
        "required_attendees": payload["required_attendees"],
        "optional_attendees": payload["optional_attendees"],
        "is_all_day": payload["is_all_day"],
        "send_invites": payload["send_invites"],
        "timezone": payload["timezone"],
        "reminder_minutes_before_start": reminder,
    }


def _update_payload(payload: dict[str, Any]) -> dict[str, Any]:
    start, end = _submission_times(payload)
    if not payload["subject"].strip():
        raise ValueError("subject is required")
    return {
        "subject": payload["subject"].strip(),
        "body": payload["body"],
        "start": start,
        "end": end,
        "location": payload["location"] or None,
        "is_all_day": payload["is_all_day"],
        "required_attendees": payload["required_attendees"],
        "optional_attendees": payload["optional_attendees"],
        "send_invites": bool(payload.get("send_invites", True)),
    }


def _submission_times(payload: dict[str, Any]) -> tuple[str, str]:
    projection = _draft_projection(payload)
    if payload["is_all_day"]:
        if projection["start_date"] is None:
            raise ValueError("start and end must be valid dates")
        return str(projection["start_date"]), str(projection["end_date"])
    if projection["start_at"] is None:
        raise ValueError("start and end must be valid and end must be later than start")
    return payload["start"], payload["end"]


def _required_draft(
    store: CredentialStore,
    credential_id: int,
    draft_id: Any,
) -> CalendarDraftRecord:
    normalized_id = _required_text(draft_id, "draft_id")
    record = store.get_calendar_draft(credential_id, normalized_id)
    if record is None:
        raise CalendarDraftNotFoundError(normalized_id)
    return record


def _parse_local_optional(value: str, timezone_name: str) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo(timezone_name))
    return parsed.astimezone(timezone.utc)


def _parse_instant(value: str) -> datetime:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return parsed.replace(tzinfo=parsed.tzinfo or timezone.utc).astimezone(timezone.utc)


def _string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError("attendees must be arrays")
    return [str(item).strip() for item in value if str(item).strip()]


def _required_text(value: Any, name: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"{name} is required")
    return text


def _optional_text(value: Any) -> str | None:
    text = str(value or "").strip()
    return text or None


def _required_version(value: Any) -> int:
    return _bounded_int(value, "version", 1, 2_147_483_647)


def _bounded_int(value: Any, name: str, minimum: int, maximum: int) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} must be an integer") from exc
    if not minimum <= parsed <= maximum:
        raise ValueError(f"{name} must be between {minimum} and {maximum}")
    return parsed


def _retention_days(value: Any) -> int | None:
    if value is None or value == "":
        return None
    return _bounded_int(value, "retention_days", 1, 3650)


def _require_confirm(params: dict[str, Any], operation: str) -> None:
    if params.get("confirm") is not True:
        raise ValueError(f"{operation} requires confirm=true")


def _looks_not_found(message: str) -> bool:
    normalized = message.casefold()
    return "404" in normalized or "not found" in normalized or "erroritemnotfound" in normalized


def _ok(data: dict[str, Any]) -> dict[str, Any]:
    return {"ok": True, "data": data}


def _error(
    code: str,
    message: str,
    status_code: int,
    data: dict[str, Any] | None = None,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "ok": False,
        "error": code,
        "message": message,
        "_status_code": status_code,
    }
    if data:
        result["data"] = data
    return result
