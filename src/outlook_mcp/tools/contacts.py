"""Contact tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client import OutlookClient
from outlook_mcp.graph_client import GraphClient

from .alias_registry import register as _register_alias
from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    PHONE_LABELS,
    ToolValidationError,
    apply_updates,
    normalize_limit,
    normalize_offset,
    query_filter,
    require_confirm,
    run_safely,
    serialize_item,
    serialize_value,
    to_page_with_more,
    validate_email,
)


CONTACT_FIELDS = (
    "display_name",
    "given_name",
    "surname",
    "company_name",
    "job_title",
    "department",
    "email_addresses",
    "phone_numbers",
    "body",
)


def contacts_list(client: OutlookClient, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List contacts."""

    def action() -> dict[str, Any]:
        if isinstance(client, GraphClient):
            return client.contacts_list(limit=limit, offset=offset)
        alias_namespace = _alias_namespace(client)
        page, has_more = to_page_with_more(
            client.account.contacts,
            offset=normalize_offset(offset),
            limit=normalize_limit(limit),
            order_by="display_name",
        )
        return {
            "items": [
                serialize_contact(item, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def contacts_search(client: OutlookClient, query: str, limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search contacts by display name."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        if isinstance(client, GraphClient):
            return client.contacts_search(query=query, limit=limit, offset=offset)
        alias_namespace = _alias_namespace(client)
        items = query_filter(client.account.contacts.all(), display_name__contains=query.strip())
        page, has_more = to_page_with_more(
            items,
            offset=normalize_offset(offset),
            limit=normalize_limit(limit),
            order_by="display_name",
        )
        return {
            "items": [
                serialize_contact(item, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def contact_create(
    client: OutlookClient,
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
        if isinstance(client, GraphClient):
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
        from exchangelib import Contact
        contact = Contact(
            account=client.account,
            folder=client.account.contacts,
            display_name=display_name,
            given_name=given_name,
            surname=surname,
            company_name=company_name,
            job_title=job_title,
            department=department,
            email_addresses=_email_addresses(email_addresses or []),
            phone_numbers=_phone_numbers(phone_numbers or []),
            body=notes,
        )
        contact.save()
        return serialize_contact(contact, alias_namespace=_alias_namespace(client))

    return run_safely(action)


def contact_update(client: OutlookClient, item_id: str, updates: dict[str, Any], confirm: bool = False) -> dict[str, Any]:
    """Update a contact."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "contact_update")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.contact_update(item_id=real_id, updates=updates)
        item = client.get_item(real_id, "contacts")
        fields = _apply_contact_updates(item, updates)
        item.save(update_fields=fields)
        return serialize_contact(item, alias_namespace=alias_namespace)

    return run_safely(action)


def contact_delete(client: OutlookClient, item_id: str, confirm: bool = False) -> dict[str, Any]:
    """Delete a contact."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "contact_delete")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.contact_delete(item_id=real_id)
        item = client.get_item(real_id, "contacts")
        item.delete()
        raw_id = serialize_value(getattr(item, "id", None)) or real_id
        out_id = (
            _register_alias(raw_id, "contact", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"deleted": True, "id": out_id or raw_id}

    return run_safely(action)


def contact_get(client: OutlookClient, item_id: str) -> dict[str, Any]:
    """Get a single contact by ID."""

    def action() -> dict[str, Any]:
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="contact",
            tool_name="contacts",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.contact_get(item_id=real_id)
        item = client.get_item(real_id, "contacts")
        return serialize_contact(item, alias_namespace=alias_namespace)

    return run_safely(action)


def serialize_contact(item: Any, *, alias_namespace: str) -> dict[str, Any]:
    return serialize_item(
        item,
        CONTACT_FIELDS,
        kind="contact",
        alias_namespace=alias_namespace,
    )


def _email_addresses(addresses: list[str]) -> list:
    from exchangelib.indexed_properties import EmailAddress
    return [EmailAddress(label=f"EmailAddress{idx}", email=validate_email(a, label="email address"))
            for idx, a in enumerate(addresses[:3], start=1)]


def _phone_numbers(numbers: list[dict[str, str]]) -> list:
    from exchangelib.indexed_properties import PhoneNumber
    result = []
    for raw in numbers:
        label = raw.get("label", "BusinessPhone")
        if label not in PHONE_LABELS:
            raise ToolValidationError(f"Unsupported phone number label: {label}")
        number = raw.get("phone_number", "").strip()
        if not number:
            raise ToolValidationError("phone_numbers entries require phone_number")
        result.append(PhoneNumber(label=label, phone_number=number))
    return result


def _apply_contact_updates(item: Any, updates: dict[str, Any]) -> list[str]:
    return apply_updates(
        item, updates,
        allowed={"display_name", "given_name", "surname", "company_name", "job_title",
                 "department", "email_addresses", "phone_numbers", "notes"},
        field_aliases={"notes": "body"},
        field_parsers={
            "email_addresses": lambda v: _email_addresses(v or []),
            "phone_numbers": lambda v: _phone_numbers(v or []),
        },
    )
