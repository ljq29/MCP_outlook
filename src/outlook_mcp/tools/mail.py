"""Mail tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client import OutlookClient
from outlook_mcp.graph_client import GraphClient
from outlook_mcp.uploads import normalize_inline_image_html

from .alias_registry import register as _register_alias
from .alias_registry import resolve_or_passthrough as _resolve_alias
from .alias_registry import namespace_for_client as _alias_namespace
from .common import (
    ToolValidationError,
    attach_with_budget,
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


MAIL_FIELDS = (
    "subject",
    "author",
    "sender",
    "datetime_received",
    "datetime_sent",
    "display_to",
    "display_cc",
    "importance",
    "is_read",
    "has_attachments",
    "size",
)


def mail_list(client: OutlookClient, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """List messages from a folder."""

    def action() -> dict[str, Any]:
        if isinstance(client, GraphClient):
            return client.mail_list(folder=folder, limit=limit, offset=offset)
        real_limit = normalize_limit(limit)
        alias_namespace = _alias_namespace(client)
        page, has_more = to_page_with_more(
            client.get_folder(folder),
            offset=normalize_offset(offset),
            limit=real_limit,
            order_by="-datetime_received",
        )
        return {
            "items": [
                serialize_mail(item, include_body=False, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def mail_search(client: OutlookClient, query: str, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
    """Search messages by subject within a folder."""

    def action() -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        if isinstance(client, GraphClient):
            return client.mail_search(query=query, folder=folder, limit=limit, offset=offset)
        items = client.get_folder(folder).all()
        alias_namespace = _alias_namespace(client)
        filtered = query_filter(items, subject__contains=query.strip())
        page, has_more = to_page_with_more(
            filtered,
            offset=normalize_offset(offset),
            limit=normalize_limit(limit),
            order_by="-datetime_received",
        )
        return {
            "items": [
                serialize_mail(item, include_body=False, alias_namespace=alias_namespace)
                for item in page
            ],
            "has_more": has_more,
        }

    return run_safely(action)


def mail_get(
    client: OutlookClient,
    item_id: str,
    folder: str | None = None,
    include_body: bool = True,
    body_max_chars: int = 4000,
) -> dict[str, Any]:
    """Read a single mail item."""

    def action() -> dict[str, Any]:
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.mail_get(
                item_id=real_id,
                folder=folder,
                include_body=include_body,
                body_max_chars=body_max_chars,
            )
        item = client.get_item(real_id, folder)
        return serialize_mail(
            item,
            include_body=include_body,
            body_max_chars=body_max_chars,
            alias_namespace=alias_namespace,
        )

    return run_safely(action)


def mail_send(
    client: OutlookClient,
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
        if isinstance(client, GraphClient):
            return client.mail_send(
                to=to,
                subject=subject,
                body=normalized_body,
                cc=cc,
                bcc=bcc,
                html=html,
                attachments=attachments,
            )
        from exchangelib import HTMLBody, Message
        message = Message(
            account=client.account,
            folder=getattr(client.account, "sent", None),
            subject=subject,
            body=HTMLBody(normalized_body) if html else normalized_body,
            to_recipients=_mailboxes(to),
            cc_recipients=_mailboxes(cc or []),
            bcc_recipients=_mailboxes(bcc or []),
        )
        total_bytes = 0
        for attachment in attachments or []:
            total_bytes = attach_with_budget(message, attachment, html=html, total_bytes=total_bytes)
        message.send_and_save()
        return {"sent": True}

    return run_safely(action)


def mail_reply(
    client: OutlookClient,
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
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.mail_reply(
                item_id=real_id,
                body=body,
                subject=subject,
                to=to,
                cc=cc,
                bcc=bcc,
                html=html,
            )
        item = client.get_item(real_id, folder)
        if html:
            from exchangelib import HTMLBody

            reply_body: str | HTMLBody = HTMLBody(body)
        else:
            reply_body = body
        reply_subject = subject or f"Re: {getattr(item, 'subject', '')}".strip()
        item.reply(
            subject=reply_subject,
            body=reply_body,
            to_recipients=_mailboxes(to or []),
            cc_recipients=_mailboxes(cc or []),
            bcc_recipients=_mailboxes(bcc or []),
        )
        raw_id = serialize_value(getattr(item, "id", None))
        out_id = (
            _register_alias(raw_id, "mail", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"replied": True, "id": out_id or raw_id}

    return run_safely(action)


def mail_move(
    client: OutlookClient,
    item_id: str,
    target_folder: str,
    folder: str | None = None,
    confirm: bool = False,
) -> dict[str, Any]:
    """Move a mail item to another folder."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_move")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id, kind="mail", tool_name="mail", namespace=alias_namespace
        )
        if isinstance(client, GraphClient):
            return client.mail_move(item_id=real_id, target_folder=target_folder)
        item = client.get_item(real_id, folder)
        moved = item.move(client.get_folder(target_folder))
        raw_id = serialize_value(getattr(moved, "id", getattr(item, "id", None)))
        out_id = (
            _register_alias(raw_id, "mail", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {
            "moved": True,
            "id": out_id or raw_id,
            "target_folder": target_folder,
        }

    return run_safely(action)


def mail_delete(
    client: OutlookClient,
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
        if isinstance(client, GraphClient):
            return client.mail_delete(item_id=real_id, mode=mode)
        if mode not in {"move_to_trash", "soft_delete", "hard_delete"}:
            raise ToolValidationError("mode must be move_to_trash, soft_delete or hard_delete")
        item = client.get_item(real_id, folder)
        if mode == "move_to_trash":
            item.move_to_trash()
        elif mode == "soft_delete":
            item.soft_delete()
        elif mode == "hard_delete":
            item.delete()
        return {"deleted": True, "mode": mode}

    return run_safely(action)


def serialize_mail(
    item: Any,
    *,
    include_body: bool,
    body_max_chars: int = 4000,
    alias_namespace: str,
) -> dict[str, Any]:
    data = serialize_item(
        item, MAIL_FIELDS, kind="mail", alias_namespace=alias_namespace
    )
    if include_body:
        raw_body = getattr(item, "body", "") or ""
        body = str(raw_body)
        if body_max_chars < 0:
            raise ToolValidationError("body_max_chars must be greater than or equal to 0")
        data["body"] = body[:body_max_chars]
        data["body_truncated"] = len(body) > body_max_chars
        data["body_content_type"] = "html" if raw_body.__class__.__name__ == "HTMLBody" else "text"
    return data


def mail_forward(
    client: OutlookClient,
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
        if isinstance(client, GraphClient):
            return client.mail_forward(item_id=real_id, to=to, body=body, folder=folder, html=html)
        item = client.get_item(real_id, folder)
        if html:
            from exchangelib import HTMLBody

            forward_body: str | HTMLBody = HTMLBody(body)
        else:
            forward_body = body
        fwd = item.create_forward(
            subject=f"Fwd: {getattr(item, 'subject', '')}",
            body=forward_body,
            to_recipients=_mailboxes(to),
        )
        if hasattr(fwd, "send"):
            fwd.send()
        return {"forwarded": True}

    return run_safely(action)


def mail_mark_read(client: OutlookClient, item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as read."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_mark_read")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.mail_mark_read(item_id=real_id, folder=folder)
        item = client.get_item(real_id, folder)
        item.is_read = True
        item.save(update_fields=["is_read"])
        raw_id = serialize_value(getattr(item, "id", None)) or real_id
        out_id = (
            _register_alias(raw_id, "mail", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"marked": "read", "id": out_id or raw_id}

    return run_safely(action)


def mail_mark_unread(client: OutlookClient, item_id: str, folder: str | None = None, confirm: bool = False) -> dict[str, Any]:
    """Mark a message as unread."""

    def action() -> dict[str, Any]:
        require_confirm(confirm, "mail_mark_unread")
        alias_namespace = _alias_namespace(client)
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=alias_namespace,
        )
        if isinstance(client, GraphClient):
            return client.mail_mark_unread(item_id=real_id, folder=folder)
        item = client.get_item(real_id, folder)
        item.is_read = False
        item.save(update_fields=["is_read"])
        raw_id = serialize_value(getattr(item, "id", None)) or real_id
        out_id = (
            _register_alias(raw_id, "mail", namespace=alias_namespace)
            if isinstance(raw_id, str) and raw_id
            else raw_id
        )
        return {"marked": "unread", "id": out_id or raw_id}

    return run_safely(action)


def mail_attachments(client: OutlookClient, item_id: str, folder: str | None = None) -> dict[str, Any]:
    """List attachments on a message."""

    def action() -> dict[str, Any]:
        real_id = _resolve_alias(
            item_id,
            kind="mail",
            tool_name="mail",
            namespace=_alias_namespace(client),
        )
        if isinstance(client, GraphClient):
            return client.mail_attachments(item_id=real_id, folder=folder)
        item = client.get_item(real_id, folder)
        atts = getattr(item, "attachments", []) or []
        items = []
        for a in atts:
            attachment_id = getattr(a, "attachment_id", None)
            items.append({
                "name": getattr(a, "name", ""),
                "size": getattr(a, "size", 0),
                "content_type": getattr(a, "content_type", ""),
                "attachment_id": serialize_value(
                    getattr(attachment_id, "id", attachment_id)
                ),
                "is_inline": getattr(a, "is_inline", False),
            })
        return {"items": items, "has_more": False, "total": len(items)}

    return run_safely(action)


def mail_folders(client: OutlookClient) -> dict[str, Any]:
    """List available mail folders."""

    def action() -> dict[str, Any]:
        if isinstance(client, GraphClient):
            return client.mail_folders()
        folders = client.account.root.walk()
        items = []
        for f in folders:
            folder_class = getattr(f, "folder_class", "") or ""
            if folder_class and str(folder_class).casefold() != "ipf.note":
                continue
            items.append({
                "id": serialize_value(getattr(f, "id", None)) or getattr(f, "name", ""),
                "name": getattr(f, "name", ""),
                "folder_class": folder_class,
                "total_count": getattr(f, "total_count", 0),
                "unread_count": getattr(f, "unread_count", 0),
                "child_folder_count": getattr(f, "child_folder_count", 0),
            })
        return {"items": items, "has_more": False, "total": len(items)}

    return run_safely(action)


def _mailboxes(addresses: list[str]) -> list:
    from exchangelib import Mailbox
    return [Mailbox(email_address=validate_email(a, label="recipient address")) for a in addresses]
