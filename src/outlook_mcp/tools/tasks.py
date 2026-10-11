"""Task tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client import OutlookClient
from outlook_mcp.graph_client import GraphClient

from .alias_registry import register as _register_alias
from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    apply_updates,
    normalize_limit,
    normalize_offset,
    parse_date,
    parse_percent,
    query_filter,
    require_confirm,
    run_safely,
    serialize_item,
    serialize_value,
    to_page_with_more,
)


TASK_FIELDS = (
    "subject",
    "body",
    "start_date",
    "due_date",
    "status",
    "percent_complete",
    "is_complete",
    "owner",
)


def tasks_list(client: OutlookClient, status: str | None = None, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List tasks."""

    def action() -> dict[str, Any]:
        if isinstance(client, GraphClient):
            return client.tasks_list(status=status, limit=limit, offset=offset)
        items = client.account.tasks.all()
        alias_namespace = _alias_namespace(client)
        if status:
            items = query_filter(items, status=status)
        page, has_more = to_page_with_more(
            items,
            offset=normalize_offset(offset),
            limit=normalize_limit(limit),
            order_by="due_date",
        )
        return {
            "items": [
                serialize_task(item, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def task_create(
    client: OutlookClient,
    subject: str,
    body: str | None = None,
    start_date: str | None = None,
    due_date: str | None = None,
    status: str = "NotStarted",
    percent_complete: int | float | str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Create a task."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "task_create")
        if not subject:
            raise ToolValidationError("subject is required")
        if isinstance(client, GraphClient):
            return client.task_create(
                subject=subject,
                body=body,
                start_date=start_date,
                due_date=due_date,
                status=status,
                percent_complete=percent_complete,
            )
        from exchangelib import Task
        task = Task(
            account=client.account,
            folder=client.account.tasks,
            subject=subject,
            body=body,
            start_date=parse_date(start_date),
            due_date=parse_date(due_date),
            status=status,
            percent_complete=parse_percent(percent_complete),
        )
        task.save()
        return serialize_task(task, alias_namespace=_alias_namespace(client))

    return run_safely(action)


def task_update(client: OutlookClient, item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a task."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "task_update")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="task", tool_name="tasks", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.task_update(item_id=real_id, updates=updates)
        item = client.get_item(real_id, "tasks")
        if updates.get("complete") is True:
            other = {k: v for k, v in updates.items() if k != "complete"}
            if other:
                _apply_task_updates(item, other)
                item.complete()
                item.save()
            else:
                item.complete()
                item.save()
            return serialize_task(item, alias_namespace=alias_namespace)
        if updates.get("complete") is False:
            other = {k: v for k, v in updates.items() if k != "complete"}
            fields = _apply_task_updates(item, other) if other else []
            item.status = "NotStarted"
            item.percent_complete = parse_percent(0)
            item.save(update_fields=[*fields, "status", "percent_complete"])
            return serialize_task(item, alias_namespace=alias_namespace)
        fields = _apply_task_updates(item, updates)
        item.save(update_fields=fields)
        return serialize_task(item, alias_namespace=alias_namespace)

    return run_safely(action)


def task_delete(client: OutlookClient, item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a task."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "task_delete")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="task",
            tool_name="tasks",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.task_delete(item_id=real_id)
        item = client.get_item(real_id, "tasks")
        item.delete()
        raw_id = serialize_value(getattr(item, "id", None)) or real_id
        out_id = (
            _register_alias(raw_id, "task", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"deleted": True, "id": out_id or raw_id}

    return run_safely(action)


def serialize_task(item: Any, *, alias_namespace: str) -> dict[str, Any]:
    return serialize_item(
        item,
        TASK_FIELDS,
        kind="task",
        alias_namespace=alias_namespace,
    )


def task_get(client: OutlookClient, item_id: str) -> dict[str, Any]:
    """Get a single task by ID."""

    def action() -> dict[str, Any]:
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="task", tool_name="tasks", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.task_get(item_id=real_id)
        item = client.get_item(real_id, "tasks")
        return serialize_task(item, alias_namespace=alias_namespace)

    return run_safely(action)


def task_search(
    client: OutlookClient,
    query: str,
    status: str | None = None,
    limit: int = 20,
    offset: int = 0,
) -> dict[str, Any]:
    """Search tasks by subject."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        if isinstance(client, GraphClient):
            return client.task_search(query=query, status=status, limit=limit, offset=offset)
        items = client.get_folder("tasks").all()
        alias_namespace = _alias_namespace(client)
        items = items.filter(subject__contains=query.strip())
        if status:
            items = items.filter(status=status)
        real_limit = normalize_limit(limit)
        real_offset = normalize_offset(offset)
        page, has_more = to_page_with_more(
            items,
            offset=real_offset,
            limit=real_limit,
            order_by="due_date",
        )
        return {
            "items": [
                serialize_task(item, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def _apply_task_updates(item: Any, updates: dict[str, Any]) -> list[str]:
    filtered = {k: v for k, v in updates.items() if k != "complete"}
    return apply_updates(
        item, filtered,
        allowed={"subject", "body", "start_date", "due_date", "status", "percent_complete"},
        field_parsers={
            "start_date": parse_date,
            "due_date": parse_date,
            "percent_complete": parse_percent,
        },
    )
