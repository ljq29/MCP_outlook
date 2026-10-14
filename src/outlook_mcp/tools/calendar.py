"""Calendar tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.graph_client import GraphClient
from outlook_mcp.uploads import normalize_inline_image_html

from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    normalize_limit,
    normalize_offset,
    normalize_reminder_minutes,
    normalize_timezone_name,
    require_confirm,
    run_safely,
)


def calendar_list(
    client: GraphClient,
    start: str | None = None,
    end: str | None = None,
    limit: int = 20,
    offset: int = 0,
    timezone: str | None = None,
) -> dict[str, Any]:
    """List calendar items in a time range."""

    def action() -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        return client.calendar_list(start=start, end=end, limit=limit, offset=offset, timezone=timezone_name)

    return run_safely(action)


def calendar_create(
    client: GraphClient,
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

    return run_safely(action)


def calendar_update(
    client: GraphClient,
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
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=_alias_namespace(client)
        )
        normalized_updates = dict(updates)
        if html and "body" in updates:
            normalized_updates["body"] = normalize_inline_image_html(updates["body"], None)
        return client.calendar_update(
            item_id=real_id, updates=normalized_updates, timezone=timezone_name, send_invites=send_invites, html=html
        )

    return run_safely(action)


def calendar_delete(
    client: GraphClient,
    item_id: str,
    send_cancellations: bool = True,
    confirm: bool = False,
) -> dict[str, Any]:
    """Delete a calendar item."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "calendar_delete")
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=_alias_namespace(client)
        )
        return client.calendar_delete(item_id=real_id, send_cancellations=send_cancellations)

    return run_safely(action)


def calendar_get(
    client: GraphClient,
    item_id: str,
    timezone: str | None = None,
) -> dict[str, Any]:
    """Get a single calendar item by ID."""

    def action() -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        real_id = _resolve_alias(
            item_id, kind="cal", tool_name="calendar", namespace=_alias_namespace(client)
        )
        return client.calendar_get(item_id=real_id, timezone=timezone_name)

    return run_safely(action)


def calendar_search(
    client: GraphClient,
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
        return client.calendar_search(
            query=query,
            start=start,
            end=end,
            limit=normalize_limit(limit),
            offset=normalize_offset(offset),
            timezone=timezone_name,
        )

    return run_safely(action)
