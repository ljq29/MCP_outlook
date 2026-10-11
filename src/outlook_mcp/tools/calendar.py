"""Calendar tools for Outlook MCP."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone as dt_timezone
from typing import Any

from outlook_mcp.client import OutlookClient
from outlook_mcp.graph_client import GraphClient
from outlook_mcp.uploads import normalize_inline_image_html

from .alias_registry import register as _register_alias
from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    apply_updates,
    attach_with_budget,
    normalize_limit,
    normalize_offset,
    normalize_reminder_minutes,
    normalize_timezone_name,
    parse_calendar_datetime,
    require_confirm,
    run_safely,
    serialize_datetime_local,
    serialize_datetime_utc,
    serialize_item,
    serialize_value,
    to_page_with_more,
)


CALENDAR_FIELDS = (
    "subject",
    "body",
    "start",
    "end",
    "location",
    "is_all_day",
    "is_cancelled",
    "is_meeting",
    "organizer",
    "required_attendees",
    "optional_attendees",
)


def calendar_list(
    client: OutlookClient,
    start: str | None = None,
    end: str | None = None,
    limit: int = 20,
    offset: int = 0,
    timezone: str | None = None,
) -> dict[str, Any]:
    """List calendar items in a time range."""

    def action() -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        if isinstance(client, GraphClient):
            return client.calendar_list(start=start, end=end, limit=limit, offset=offset, timezone=timezone_name)
        now = datetime.now(dt_timezone.utc)
        start_dt = parse_calendar_datetime(start, timezone_name=timezone_name, default=now)
        end_dt = parse_calendar_datetime(end, timezone_name=timezone_name, default=start_dt + timedelta(days=7))
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        view = client.account.calendar.view(start=start_dt, end=end_dt)
        alias_namespace = _alias_namespace(client)
        page, has_more = to_page_with_more(
            view,
            offset=normalize_offset(offset),
            limit=normalize_limit(limit),
            order_by="start",
        )
        return {
            "items": [
                serialize_calendar_item(
                    item,
                    timezone_name=timezone_name,
                    alias_namespace=alias_namespace,
                )
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def calendar_create(
    client: OutlookClient,
    subject: str,
    start: str,
    end: str,
    body: str = "",
    location: str | None = None,
    required_attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    is_all_day: bool = False,
    send_invites: bool = True,
    html: bool = False,
    attachments: list[dict[str, Any]] | None = None,
    confirm: bool = False,
    timezone: str | None = None,
    reminder_minutes_before_start: int | None = None,
) -> dict[str, Any]:
    """Create a calendar item or meeting."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "calendar_create")
        timezone_name = normalize_timezone_name(timezone)
        reminder_minutes = normalize_reminder_minutes(reminder_minutes_before_start)
        if not subject:
            raise ToolValidationError("subject is required")
        normalized_body = normalize_inline_image_html(body, attachments) if html else body
        if isinstance(client, GraphClient):
            return client.calendar_create(
                subject=subject,
                start=start,
                end=end,
                body=normalized_body,
                location=location,
                required_attendees=required_attendees,
                optional_attendees=optional_attendees,
                is_all_day=is_all_day,
                send_invites=send_invites,
                html=html,
                attachments=attachments,
                timezone=timezone_name,
                reminder_minutes_before_start=reminder_minutes,
            )
        start_dt = parse_calendar_datetime(start, timezone_name=timezone_name)
        end_dt = parse_calendar_datetime(end, timezone_name=timezone_name)
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        from exchangelib import CalendarItem, HTMLBody
        from exchangelib.items import SEND_TO_ALL_AND_SAVE_COPY, SEND_TO_NONE
        reminder_kwargs: dict[str, Any] = {"reminder_is_set": reminder_minutes is not None}
        if reminder_minutes is not None:
            reminder_kwargs["reminder_minutes_before_start"] = reminder_minutes
        item = CalendarItem(
            account=client.account,
            folder=client.account.calendar,
            subject=subject,
            body=HTMLBody(normalized_body) if html else normalized_body,
            start=start_dt,
            end=end_dt,
            location=location,
            required_attendees=required_attendees or [],
            optional_attendees=optional_attendees or [],
            is_all_day=is_all_day,
            **reminder_kwargs,
        )
        total_bytes = 0
        for attachment in attachments or []:
            total_bytes = attach_with_budget(item, attachment, html=html, total_bytes=total_bytes)
        item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY if send_invites else SEND_TO_NONE)
        return serialize_calendar_item(
            item,
            timezone_name=timezone_name,
            alias_namespace=_alias_namespace(client),
        )

    return run_safely(action)


def calendar_update(
    client: OutlookClient,
    item_id: str,
    updates: dict[str, Any],
    send_invites: bool = True,
    confirm: bool = False,
    timezone: str | None = None,
    html: bool = False,
) -> dict[str, Any]:
    """Update a calendar item."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "calendar_update")
        timezone_name = normalize_timezone_name(timezone)
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=alias_namespace
        )
        normalized_updates = dict(updates)
        if html and "body" in updates:
            normalized_updates["body"] = normalize_inline_image_html(updates["body"], None)
        if isinstance(client, GraphClient):
            return client.calendar_update(item_id=real_id, updates=normalized_updates, timezone=timezone_name, send_invites=send_invites, html=html)
        item = client.get_item(real_id, "calendar")
        if html and "body" in normalized_updates:
            from exchangelib import HTMLBody

            normalized_updates["body"] = HTMLBody(normalized_updates["body"])
        fields = _apply_calendar_updates(item, normalized_updates, timezone_name=timezone_name)
        from exchangelib.items import SEND_ONLY_TO_CHANGED, SEND_TO_NONE
        item.save(
            update_fields=fields,
            send_meeting_invitations=SEND_ONLY_TO_CHANGED if send_invites else SEND_TO_NONE,
        )
        return serialize_calendar_item(
            item,
            timezone_name=timezone_name,
            alias_namespace=alias_namespace,
        )

    return run_safely(action)


def calendar_delete(
    client: OutlookClient,
    item_id: str,
    send_cancellations: bool = True,
    confirm: bool = False,
) -> dict[str, Any]:
    """Delete a calendar item."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "calendar_delete")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.calendar_delete(item_id=real_id, send_cancellations=send_cancellations)
        item = client.get_item(real_id, "calendar")
        if send_cancellations:
            from exchangelib.items import SEND_ONLY_TO_ALL
            item.delete(send_meeting_cancellations=SEND_ONLY_TO_ALL)
        else:
            item.delete()
        raw_id = serialize_value(getattr(item, "id", None))
        out_id = (
            _register_alias(raw_id, "cal", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"deleted": True, "id": out_id or raw_id}

    return run_safely(action)


def serialize_calendar_item(
    item: Any,
    timezone_name: str | None = None,
    *,
    alias_namespace: str,
) -> dict[str, Any]:
    normalized_timezone = normalize_timezone_name(timezone_name)
    data = serialize_item(
        item, CALENDAR_FIELDS, kind="cal", alias_namespace=alias_namespace
    )
    start = getattr(item, "start", None)
    end = getattr(item, "end", None)
    if isinstance(start, datetime):
        data["start"] = serialize_datetime_utc(start)
        data["start_local"] = serialize_datetime_local(start, normalized_timezone)
    else:
        data["start_local"] = None
    if isinstance(end, datetime):
        data["end"] = serialize_datetime_utc(end)
        data["end_local"] = serialize_datetime_local(end, normalized_timezone)
    else:
        data["end_local"] = None
    data["local_timezone"] = normalized_timezone
    return data


def calendar_get(
    client: OutlookClient,
    item_id: str,
    timezone: str | None = None,
) -> dict[str, Any]:
    """Get a single calendar item by ID."""

    def action() -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.calendar_get(item_id=real_id, timezone=timezone_name)
        item = client.get_item(real_id, "calendar")
        return serialize_calendar_item(
            item,
            timezone_name=timezone_name,
            alias_namespace=alias_namespace,
        )

    return run_safely(action)


def calendar_search(
    client: OutlookClient,
    query: str,
    start: str | None = None,
    end: str | None = None,
    limit: int = 20,
    offset: int = 0,
    timezone: str | None = None,
) -> dict[str, Any]:
    """Search calendar items by subject."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        timezone_name = normalize_timezone_name(timezone)
        real_offset = normalize_offset(offset)
        real_limit = normalize_limit(limit)
        if isinstance(client, GraphClient):
            return client.calendar_search(
                query=query, start=start, end=end,
                limit=real_limit, offset=real_offset, timezone=timezone_name,
            )
        now = datetime.now(dt_timezone.utc)
        start_dt = parse_calendar_datetime(start, timezone_name=timezone_name, default=now)
        end_dt = parse_calendar_datetime(end, timezone_name=timezone_name, default=start_dt + timedelta(days=30))
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        view = client.account.calendar.view(start=start_dt, end=end_dt)
        filtered = view.filter(subject__contains=query.strip())
        alias_namespace = _alias_namespace(client)
        page, has_more = to_page_with_more(
            filtered,
            offset=real_offset,
            limit=real_limit,
            order_by="start",
        )
        return {
            "items": [
                serialize_calendar_item(
                    item,
                    timezone_name=timezone_name,
                    alias_namespace=alias_namespace,
                )
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def _apply_calendar_updates(item: Any, updates: dict[str, Any], *, timezone_name: str | None = None) -> list[str]:
    fields = apply_updates(
        item, updates,
        allowed={"subject", "body", "start", "end", "location",
                 "required_attendees", "optional_attendees", "is_all_day"},
        field_parsers={
            "start": lambda v: parse_calendar_datetime(v, timezone_name=timezone_name),
            "end": lambda v: parse_calendar_datetime(v, timezone_name=timezone_name),
        },
    )
    if getattr(item, "start", None) and getattr(item, "end", None) and item.end <= item.start:
        raise ToolValidationError("end must be later than start")
    return fields
