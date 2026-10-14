"""Mail tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.graph_client import GraphClient
from outlook_mcp.uploads import normalize_inline_image_html

from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    require_confirm,
    run_safely,
)


def mail_list(client: GraphClient, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List messages from a folder."""

    def action() -> dict[str, Any]:
        return client.mail_list(folder=folder, limit=limit, offset=offset)

    return run_safely(action)


def mail_search(client: GraphClient, query: str, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search messages by subject within a folder."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        return client.mail_search(query=query, folder=folder, limit=limit, offset=offset)

    return run_safely(action)


def mail_get(
    client: GraphClient,
    item_id: str,
    folder: str | None = None,
    include_body: bool = True,
    body_max_chars: int = 4000,
) -> dict[str, Any]:
    """Read a single mail item."""

    def action() -> dict[str, Any]:
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=_alias_namespace(client)
        )
        return client.mail_get(
            item_id=real_id,
            folder=folder,
            include_body=include_body,
            body_max_chars=body_max_chars,
        )

    return run_safely(action)


def mail_send(
    client: GraphClient,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[dict[str, Any]] | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Send a new message and save it to Sent Items."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_send")
        if not to:
            raise ToolValidationError("to must contain at least one recipient")
        if not subject:
            raise ToolValidationError("subject is required")
        normalized_body = normalize_inline_image_html(body, attachments) if html else body
        return client.mail_send(
            to=to,
            subject=subject,
            body=normalized_body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
        )

    return run_safely(action)


def mail_reply(
    client: GraphClient,
    item_id: str,
    body: str,
    subject: str | None = None,
    folder: str | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    confirm: bool = False,
) -> dict[str, Any]:
    """Reply to an existing message."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_reply")
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=_alias_namespace(client)
        )
        return client.mail_reply(
            item_id=real_id,
            body=body,
            subject=subject,
            to=to,
            cc=cc,
            bcc=bcc,
            html=html,
        )

    return run_safely(action)


def mail_move(
    client: GraphClient,
    item_id: str,
    target_folder: str,
    folder: str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Move a mail item to another folder."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_move")
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=_alias_namespace(client)
        )
        return client.mail_move(item_id=real_id, target_folder=target_folder)

    return run_safely(action)


def mail_delete(
    client: GraphClient,
    item_id: str,
    folder: str | None = None,
    mode: str = "move_to_trash",
    confirm: bool = False,
) -> dict[str, Any]:
    """Delete a mail item using move_to_trash, soft_delete or hard_delete."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_delete")
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        return client.mail_delete(item_id=real_id, mode=mode)

    return run_safely(action)


def mail_forward(
    client: GraphClient,
    item_id: str,
    to: list[str],
    body: str = "",
    folder: str | None = None,
    html: bool = False,
    confirm: bool = False,
) -> dict[str, Any]:
    """Forward a message to recipients."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_forward")
        if not to:
            raise ToolValidationError("to must contain at least one recipient")
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        return client.mail_forward(item_id=real_id, to=to, body=body, folder=folder, html=html)

    return run_safely(action)


def mail_mark_read(client: GraphClient, item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as read."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_mark_read")
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        return client.mail_mark_read(item_id=real_id, folder=folder)

    return run_safely(action)


def mail_mark_unread(client: GraphClient, item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as unread."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_mark_unread")
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        return client.mail_mark_unread(item_id=real_id, folder=folder)

    return run_safely(action)


def mail_attachments(client: GraphClient, item_id: str, folder: str | None = None) -> dict[str, Any]:
    """List attachments on a message."""

    def action() -> dict[str, Any]:
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        return client.mail_attachments(item_id=real_id, folder=folder)

    return run_safely(action)


def mail_folders(client: GraphClient) -> dict[str, Any]:
    """List available mail folders."""

    def action() -> dict[str, Any]:
        return client.mail_folders()

    return run_safely(action)
