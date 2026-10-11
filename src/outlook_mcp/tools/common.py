"""Shared validation, serialization and error helpers for MCP tools."""

from __future__ import annotations

import logging
import os
from datetime import date, datetime, timezone
import binascii
import traceback
from decimal import Decimal, InvalidOperation
from typing import Any, Callable, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from outlook_mcp.client import OutlookClientError
from outlook_mcp.config import EMAIL_RE, sensitive_values

DEFAULT_CALENDAR_TIMEZONE = "Asia/Shanghai"
# 3.5MB total per message/event; Graph /me/sendMail rejects request bodies over ~4MB.
MAX_ATTACHMENTS_TOTAL_BYTES = 3 * 1024 * 1024 + 512 * 1024

# exchangelib PhoneNumber indexed-property labels; shared by EWS and Graph contact writers.
PHONE_LABELS = frozenset({"BusinessPhone", "BusinessPhone2", "HomePhone", "HomePhone2", "MobilePhone"})


class ToolValidationError(ValueError):
    """Raised when a tool call has invalid or unsafe input."""


class RequestBodyTooLargeError(ToolValidationError):
    """Raised when an HTTP request body exceeds the configured byte limit."""


class AuthRequiredError(RuntimeError):
    """Raised when a delegated Graph login must be completed by the user."""

    def __init__(self, message: str, data: dict[str, Any]):
        super().__init__(message)
        self.data = data


class MethodUnsupportedError(RuntimeError):
    """Raised when an HTTP endpoint method exists but is not implemented."""


class OutcomeUnknownError(RuntimeError):
    """Raised when a remote write may have completed but cannot be confirmed."""


def ok(data: Any = None, message: str | None = None) -> dict[str, Any]:
    result: dict[str, Any] = {"ok": True, "data": data if data is not None else {}}
    if message:
        result["message"] = message
    return result


def fail(error: str, message: str) -> dict[str, Any]:
    return {"ok": False, "error": error, "message": redact_secrets(message)}


def run_safely(action: Callable[[], Any]) -> dict[str, Any]:
    """Execute a tool action and normalize its errors."""
    try:
        result = ok(action())
    except AuthRequiredError as exc:
        return {"ok": False, "error": "auth_required", "message": redact_secrets(str(exc)), "data": exc.data}
    except MethodUnsupportedError as exc:
        return fail("method_unsupported", str(exc))
    except OutcomeUnknownError as exc:
        return fail("outcome_unknown", str(exc))
    except ToolValidationError as exc:
        return fail("validation_error", str(exc))
    except OutlookClientError as exc:
        return fail("outlook_error", str(exc))
    except Exception as exc:
        logging.getLogger("outlook_mcp").error(
            "unexpected error in tool action: %s\n%s",
            redact_secrets(str(exc) or exc.__class__.__name__),
            redact_secrets("".join(traceback.format_exception(exc))),
        )
        return fail("exchange_error", str(exc) or exc.__class__.__name__)
    return result


def require_confirm(confirm: bool, operation: str) -> None:
    if confirm is not True:
        raise ToolValidationError(f"{operation} requires confirm=true")


def normalize_limit(limit: int, *, default: int = 20, maximum: int = 100) -> int:
    if limit is None:
        return default
    if limit < 1:
        raise ToolValidationError("limit must be greater than 0")
    if limit > maximum:
        raise ToolValidationError(f"limit must be less than or equal to {maximum}")
    return limit


def normalize_offset(offset: int) -> int:
    if offset is None:
        return 0
    if offset < 0:
        raise ToolValidationError("offset must be greater than or equal to 0")
    return offset


def normalize_reminder_minutes(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ToolValidationError("reminder_minutes_before_start must be a non-negative integer")
    return value


def parse_datetime(value: str | datetime | None, *, default: datetime | None = None, timezone_name: str | None = None) -> datetime:
    if value is None:
        if default is None:
            raise ToolValidationError("datetime value is required")
        return default
    if isinstance(value, datetime):
        parsed = value
    else:
        normalized = value.strip().replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError as exc:
            raise ToolValidationError(f"Invalid datetime: {value}") from exc
    if parsed.tzinfo is None:
        if timezone_name is not None:
            parsed = parsed.replace(tzinfo=ZoneInfo(normalize_timezone_name(timezone_name)))
        else:
            parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def parse_calendar_datetime(
    value: str | datetime | None,
    *,
    timezone_name: str | None,
    default: datetime | None = None,
) -> datetime:
    return parse_datetime(value, default=default, timezone_name=timezone_name)


def normalize_timezone_name(timezone_name: str | None) -> str:
    value = (timezone_name or DEFAULT_CALENDAR_TIMEZONE).strip() or DEFAULT_CALENDAR_TIMEZONE
    try:
        ZoneInfo(value)
    except ZoneInfoNotFoundError as exc:
        raise ToolValidationError(f"Invalid timezone: {timezone_name}") from exc
    return value


def serialize_datetime_utc(value: datetime | None) -> str | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.astimezone(timezone.utc).isoformat()


def serialize_datetime_local(value: datetime | None, timezone_name: str | None = None) -> str | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    normalized = normalize_timezone_name(timezone_name)
    return value.astimezone(ZoneInfo(normalized)).isoformat()


def parse_date(value: str | date | None) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise ToolValidationError(f"Invalid date: {value}") from exc


def parse_percent(value: int | float | str | Decimal | None) -> Decimal | None:
    if value is None:
        return None
    try:
        percent = Decimal(str(value))
    except (ValueError, InvalidOperation) as exc:
        raise ToolValidationError("percent_complete must be a number") from exc
    if percent < Decimal(0) or percent > Decimal(100):
        raise ToolValidationError("percent_complete must be between 0 and 100")
    return percent


def to_page(queryset: Any, *, offset: int, limit: int, order_by: str | None = None) -> list[Any]:
    if order_by and hasattr(queryset, "order_by"):
        queryset = queryset.order_by(order_by)
    try:
        return list(queryset[offset : offset + limit])
    except TypeError:
        items = queryset.all() if hasattr(queryset, "all") else queryset
        return list(items)[offset : offset + limit]


def to_page_with_more(
    queryset: Any,
    *,
    offset: int,
    limit: int,
    order_by: str | None = None,
) -> tuple[list[Any], bool]:
    """Fetch one extra item so ``has_more`` is exact without counting all rows."""
    candidates = to_page(
        queryset,
        offset=offset,
        limit=limit + 1,
        order_by=order_by,
    )
    return candidates[:limit], len(candidates) > limit


def query_filter(queryset: Any, **criteria: Any) -> Any:
    if hasattr(queryset, "filter"):
        return queryset.filter(**criteria)
    return queryset


def apply_updates(
    item: Any,
    updates: dict[str, Any],
    allowed: set[str],
    *,
    field_parsers: dict[str, Callable[[Any], Any]] | None = None,
    field_aliases: dict[str, str] | None = None,
) -> list[str]:
    """Apply validated updates to an item. Returns list of attribute names that were modified."""
    if not updates:
        raise ToolValidationError("updates cannot be empty")
    parsers = field_parsers or {}
    aliases = field_aliases or {}
    fields: list[str] = []
    for field, value in updates.items():
        if field not in allowed:
            raise ToolValidationError(f"Unsupported update field: {field}")
        attr = aliases.get(field, field)
        if field in parsers:
            value = parsers[field](value)
        setattr(item, attr, value)
        fields.append(attr)
    return fields


def serialize_item(
    item: Any,
    fields: Iterable[str],
    *,
    kind: str | None = None,
    alias_namespace: str | None = None,
) -> dict[str, Any]:
    raw_id = serialize_value(getattr(item, "id", getattr(item, "item_id", None)))
    if kind and isinstance(raw_id, str) and raw_id:
        from .alias_registry import register as _register_alias

        if not alias_namespace:
            raise ValueError("alias_namespace is required when serializing an aliased item")
        out_id = _register_alias(raw_id, kind, namespace=alias_namespace) or raw_id
    else:
        out_id = raw_id
    data: dict[str, Any] = {
        "id": out_id,
        "changekey": serialize_value(getattr(item, "changekey", None)),
    }
    for field in fields:
        data[field] = serialize_value(getattr(item, field, None))
    return data


def serialize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (list, tuple, set)):
        return [serialize_value(item) for item in value]
    if hasattr(value, "email_address"):
        return getattr(value, "email_address")
    if hasattr(value, "email"):
        return {"label": getattr(value, "label", None), "email": getattr(value, "email", None)}
    if hasattr(value, "phone_number"):
        return {"label": getattr(value, "label", None), "phone_number": getattr(value, "phone_number", None)}
    if hasattr(value, "mailbox"):
        return serialize_value(getattr(value, "mailbox"))
    return str(value)


def msal_error_message(payload: dict[str, Any] | None) -> str:
    """Extract a human-readable error from an MSAL response dict."""
    if not payload:
        return "Microsoft authentication failed"
    return str(
        payload.get("error_description")
        or payload.get("message")
        or payload.get("error")
        or "Microsoft authentication failed"
    )


def graph_error_message(status_code: int, payload: dict[str, Any]) -> str:
    """Extract a human-readable error from a Microsoft Graph error response."""
    error = payload.get("error") if isinstance(payload, dict) else None
    if isinstance(error, dict):
        message = error.get("message") or error.get("code")
        code = error.get("code")
        if message and code:
            return f"Microsoft Graph error {status_code} ({code}): {message}"
        if message:
            return f"Microsoft Graph error {status_code}: {message}"
    return f"Microsoft Graph error {status_code}"


# CLIENT_ID/TENANT_ID are intentionally excluded: they are identifiers, not secrets.
_SECRET_ENV_NAMES = ("OUTLOOK_PASSWORD", "OUTLOOK_CLIENT_SECRET")


def redact_secrets(message: str) -> str:
    redacted = message
    for name in _SECRET_ENV_NAMES:
        value = os.environ.get(name)
        if value:
            redacted = redacted.replace(value, "[redacted]")
    for value in sensitive_values():
        redacted = redacted.replace(value, "[redacted]")
    return redacted


def validate_email(address: str, *, label: str = "address") -> str:
    """Validate and clean a single email address. Returns the cleaned address or raises ToolValidationError."""
    value = address.strip()
    if not value or not EMAIL_RE.match(value):
        raise ToolValidationError(f"invalid {label}: {address}")
    return value


def file_attachment(raw: dict[str, Any], *, html: bool = False) -> Any:
    """Build an EWS FileAttachment from a raw attachment dict."""
    import base64
    from exchangelib import FileAttachment
    from outlook_mcp.uploads import DEFAULT_UPLOAD_MAX_BYTES, attachment_from_upload

    if raw.get("relative_path"):
        attachment = attachment_from_upload(raw, html=html)
        return FileAttachment(
            name=attachment.name,
            content=attachment.content,
            content_type=attachment.content_type,
            content_id=attachment.content_id,
            is_inline=attachment.disposition == "inline",
        )

    name = raw.get("name", "").strip()
    content_base64 = raw.get("content_base64", "")
    if not name or not content_base64:
        raise ToolValidationError("attachments require name and content_base64")
    try:
        content = base64.b64decode(content_base64, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ToolValidationError(f"Invalid base64 attachment content for {name}") from exc
    if len(content) > DEFAULT_UPLOAD_MAX_BYTES:
        raise ToolValidationError(
            f"attachment {name} is {len(content)} bytes, exceeding the "
            f"{DEFAULT_UPLOAD_MAX_BYTES} bytes per-attachment limit"
        )
    return FileAttachment(name=name, content=content)


def attach_with_budget(item: Any, raw: dict[str, Any], *, html: bool, total_bytes: int) -> int:
    """Attach ``raw`` to ``item`` and return the new running total of attachment bytes."""
    attachment = file_attachment(raw, html=html)
    total_bytes += len(getattr(attachment, "content", None) or b"")
    if total_bytes > MAX_ATTACHMENTS_TOTAL_BYTES:
        raise ToolValidationError(
            f"attachments total {total_bytes} bytes exceeds the "
            f"{MAX_ATTACHMENTS_TOTAL_BYTES} bytes per-message limit"
        )
    item.attach(attachment)
    return total_bytes
