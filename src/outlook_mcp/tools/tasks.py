"""Task tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.graph_client import GraphClient

from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    require_confirm,
    run_safely,
)


def tasks_list(client: GraphClient, status: str | None = None, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List tasks."""

    def action() -> dict[str, Any]:
        return client.tasks_list(status=status, limit=limit, offset=offset)

    return run_safely(action)


def task_create(
    client: GraphClient,
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
        return client.task_create(
            subject=subject,
            body=body,
            start_date=start_date,
            due_date=due_date,
            status=status,
            percent_complete=percent_complete,
        )

    return run_safely(action)


def task_update(client: GraphClient, item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a task."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "task_update")
        real_id = _resolve_alias(
            item_id, kind="task", tool_name="tasks", namespace=_alias_namespace(client)
        )
        return client.task_update(item_id=real_id, updates=updates)

    return run_safely(action)


def task_delete(client: GraphClient, item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a task."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "task_delete")
        real_id = _resolve_alias(
            item_id,
            kind="task",
            tool_name="tasks",
            namespace=_alias_namespace(client),
        )
        return client.task_delete(item_id=real_id)

    return run_safely(action)


def task_get(client: GraphClient, item_id: str) -> dict[str, Any]:
    """Get a single task by ID."""

    def action() -> dict[str, Any]:
        real_id = _resolve_alias(
            item_id, kind="task", tool_name="tasks", namespace=_alias_namespace(client)
        )
        return client.task_get(item_id=real_id)

    return run_safely(action)


def task_search(
    client: GraphClient,
    query: str,
    status: str | None = None,
    limit: int = 20,
    offset: int = 0,
) -> dict[str, Any]:
    """Search tasks by subject."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        return client.task_search(query=query, status=status, limit=limit, offset=offset)

    return run_safely(action)
