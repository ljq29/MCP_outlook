"""Contact tools for Outlook MCP."""

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


def contacts_list(client: GraphClient, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List contacts."""

    def action() -> dict[str, Any]:
        return client.contacts_list(limit=limit, offset=offset)

    return run_safely(action)


def contacts_search(client: GraphClient, query: str, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search contacts by display name."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        return client.contacts_search(query=query, limit=limit, offset=offset)

    return run_safely(action)


def contact_create(
    client: GraphClient,
    display_name: str,
    given_name: str | None = None,
    surname: str | None = None,
    company_name: str | None = None,
    job_title: str | None = None,
    department: str | None = None,
    email_addresses: list[str] | None = None,
    phone_numbers: list[dict[str, str]] | None = None,
    notes: str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Create a contact."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "contact_create")
        if not display_name:
            raise ToolValidationError("display_name is required")
        return client.contact_create(
            display_name=display_name,
            given_name=given_name,
            surname=surname,
            company_name=company_name,
            job_title=job_title,
            department=department,
            email_addresses=email_addresses,
            phone_numbers=phone_numbers,
            notes=notes,
        )

    return run_safely(action)


def contact_update(client: GraphClient, item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a contact."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "contact_update")
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=_alias_namespace(client),
        )
        return client.contact_update(item_id=real_id, updates=updates)

    return run_safely(action)


def contact_delete(client: GraphClient, item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a contact."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "contact_delete")
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=_alias_namespace(client),
        )
        return client.contact_delete(item_id=real_id)

    return run_safely(action)


def contact_get(client: GraphClient, item_id: str) -> dict[str, Any]:
    """Get a single contact by ID."""

    def action() -> dict[str, Any]:
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=_alias_namespace(client),
        )
        return client.contact_get(item_id=real_id)

    return run_safely(action)
