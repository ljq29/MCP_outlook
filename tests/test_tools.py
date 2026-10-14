from __future__ import annotations

import hashlib

import pytest

from outlook_mcp.config import register_sensitive_value
from outlook_mcp.tools import alias_registry, calendar, common, contacts, mail, tasks
from outlook_mcp.tools import account as account_tools
from outlook_mcp.uploads import reset_upload_credential_id, save_uploaded_file, set_upload_credential_id

from conftest import FakeGraphClient, PNG_BYTES


def _alias(item_id: str, kind: str) -> str:
    return f"{kind}:" + hashlib.sha1(item_id.encode("utf-8")).hexdigest()[:16]


@pytest.fixture(autouse=True)
def _clear_alias_registry():
    alias_registry.clear()
    yield
    alias_registry.clear()


def test_mail_tools_passthrough() -> None:
    client = FakeGraphClient()

    assert mail.mail_list(client)["data"]["items"][0]["subject"] == "Quarterly report"
    assert client.last_call("mail_list") == {"folder": "inbox", "limit": 20, "offset": 0}
    assert mail.mail_search(client, "report")["ok"] is True
    assert client.last_call("mail_search")["query"] == "report"
    assert mail.mail_get(client, "mail-1")["data"]["body"] == "hello body"

    blocked = mail.mail_send(client, ["a@example.com"], "Hi", "Body")
    assert blocked["ok"] is False
    assert blocked["error"] == "validation_error"

    sent = mail.mail_send(client, ["a@example.com"], "Hi", "Body", confirm=True)
    assert sent["ok"] is True
    send_kwargs = client.last_call("mail_send")
    assert send_kwargs["to"] == ["a@example.com"]
    assert send_kwargs["subject"] == "Hi"

    assert mail.mail_reply(client, "mail-1", "Thanks", confirm=True)["ok"] is True
    assert client.last_call("mail_reply")["item_id"] == "mail-1"

    assert mail.mail_move(client, "mail-1", "trash", confirm=True)["ok"] is True
    assert client.last_call("mail_move") == {"item_id": "mail-1", "target_folder": "trash"}

    deleted = mail.mail_delete(client, "mail-1", mode="move_to_trash", confirm=True)
    assert deleted["ok"] is True
    assert client.last_call("mail_delete") == {"item_id": "mail-1", "mode": "move_to_trash"}

    blocked_read = mail.mail_mark_read(client, "mail-1")
    assert blocked_read["ok"] is False
    assert blocked_read["error"] == "validation_error"
    assert mail.mail_mark_read(client, "mail-1", confirm=True)["data"]["marked"] == "read"

    blocked_unread = mail.mail_mark_unread(client, "mail-1")
    assert blocked_unread["ok"] is False
    assert blocked_unread["error"] == "validation_error"
    assert mail.mail_mark_unread(client, "mail-1", confirm=True)["data"]["marked"] == "unread"

    assert mail.mail_forward(client, "mail-1", ["a@example.com"], confirm=True)["ok"] is True
    assert mail.mail_attachments(client, "mail-1")["ok"] is True
    assert mail.mail_folders(client)["data"]["total"] == 1


def test_calendar_tools_passthrough() -> None:
    client = FakeGraphClient()
    start = "2026-01-02T00:00:00+00:00"
    end = "2026-01-03T00:00:00+00:00"

    assert calendar.calendar_list(client, start=start, end=end)["ok"] is True
    list_kwargs = client.last_call("calendar_list")
    assert list_kwargs["timezone"] == "Asia/Shanghai"

    blocked = calendar.calendar_create(client, "Meet", start, end)
    assert blocked["ok"] is False

    created = calendar.calendar_create(
        client, "Meet", start, end, confirm=True,
        timezone="America/New_York", reminder_minutes_before_start=10,
    )
    assert created["ok"] is True
    create_kwargs = client.last_call("calendar_create")
    assert create_kwargs["timezone"] == "America/New_York"
    assert create_kwargs["reminder_minutes_before_start"] == 10

    assert calendar.calendar_update(client, "cal-1", {"location": "Room 2"}, confirm=True)["ok"] is True
    assert client.last_call("calendar_update")["updates"] == {"location": "Room 2"}
    assert calendar.calendar_delete(client, "cal-1", confirm=True)["data"]["deleted"] is True
    assert calendar.calendar_get(client, "cal-1")["ok"] is True
    assert calendar.calendar_search(client, "Plan")["ok"] is True


def test_contact_tools_passthrough() -> None:
    client = FakeGraphClient()

    assert contacts.contacts_list(client)["data"]["items"][0]["display_name"] == "Ada Lovelace"
    assert contacts.contacts_search(client, "Ada")["ok"] is True

    blocked = contacts.contact_create(client, "Grace Hopper")
    assert blocked["ok"] is False

    created = contacts.contact_create(
        client,
        "Grace Hopper",
        email_addresses=["grace@example.com"],
        phone_numbers=[{"label": "BusinessPhone", "phone_number": "123"}],
        confirm=True,
    )
    assert created["ok"] is True
    assert client.last_call("contact_create")["display_name"] == "Grace Hopper"

    assert contacts.contact_update(client, "contact-1", {"job_title": "Engineer"}, confirm=True)["ok"] is True
    assert contacts.contact_delete(client, "contact-1", confirm=True)["data"]["deleted"] is True
    assert contacts.contact_get(client, "contact-1")["ok"] is True


def test_task_tools_passthrough() -> None:
    client = FakeGraphClient()

    assert tasks.tasks_list(client)["data"]["items"][0]["subject"] == "Follow up"

    blocked = tasks.task_create(client, "Write notes")
    assert blocked["ok"] is False

    created = tasks.task_create(client, "Write notes", due_date="2026-01-02", percent_complete=10, confirm=True)
    assert created["ok"] is True
    assert client.last_call("task_create")["percent_complete"] == 10

    assert tasks.task_update(client, "task-1", {"complete": True}, confirm=True)["ok"] is True
    assert client.last_call("task_update")["updates"] == {"complete": True}
    assert tasks.task_delete(client, "task-1", confirm=True)["data"]["deleted"] is True
    assert tasks.task_get(client, "task-1")["ok"] is True
    assert tasks.task_search(client, "Follow")["ok"] is True


def test_account_info_passthrough() -> None:
    client = FakeGraphClient()

    result = account_tools.account_info(client)

    assert result["data"] == {
        "primary_smtp_address": "person@outlook.com",
        "display_name": "Example Person",
    }


def test_validation_edges() -> None:
    client = FakeGraphClient()

    assert mail.mail_search(client, "  ")["error"] == "validation_error"
    assert contacts.contacts_search(client, "")["error"] == "validation_error"
    assert tasks.task_search(client, "")["error"] == "validation_error"
    assert calendar.calendar_search(client, "  ")["error"] == "validation_error"
    assert mail.mail_send(client, [], "Hi", "Body", confirm=True)["error"] == "validation_error"
    assert mail.mail_send(client, ["a@example.com"], "", "Body", confirm=True)["error"] == "validation_error"
    assert mail.mail_forward(client, "mail-1", [], confirm=True)["error"] == "validation_error"
    assert contacts.contact_create(client, "", confirm=True)["error"] == "validation_error"
    assert tasks.task_create(client, "", confirm=True)["error"] == "validation_error"
    assert calendar.calendar_create(
        client, "", "2026-01-02T09:00:00+00:00", "2026-01-02T10:00:00+00:00", confirm=True
    )["error"] == "validation_error"
    assert calendar.calendar_list(client, timezone="Bad/Zone")["error"] == "validation_error"
    assert calendar.calendar_create(
        client,
        "Bad reminder",
        "2026-01-02T09:00:00+00:00",
        "2026-01-02T10:00:00+00:00",
        confirm=True,
        reminder_minutes_before_start=-1,
    )["error"] == "validation_error"
    assert calendar.calendar_create(
        client,
        "Bad reminder",
        "2026-01-02T09:00:00+00:00",
        "2026-01-02T10:00:00+00:00",
        confirm=True,
        reminder_minutes_before_start=True,
    )["error"] == "validation_error"


def test_calendar_update_has_no_attachments_parameter() -> None:
    import inspect

    assert "attachments" not in inspect.signature(calendar.calendar_update).parameters


def test_mail_send_normalizes_inline_upload_image_before_send(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    client = FakeGraphClient()
    try:
        sent = mail.mail_send(
            client,
            ["a@example.com"],
            "Hi",
            '<img src="cid:photo1">',
            html=True,
            attachments=[{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "photo1"}],
            confirm=True,
        )
    finally:
        reset_upload_credential_id(token)

    assert sent["ok"] is True
    send_kwargs = client.last_call("mail_send")
    assert send_kwargs["body"] == '<img src="cid:photo1" width="320" height="180">'
    assert send_kwargs["html"] is True


# --- alias resolution at the tools layer ---


def test_alias_round_trip_through_tools() -> None:
    client = FakeGraphClient()
    alias = alias_registry.register("mail-1", "mail", namespace="graph:person@outlook.com")

    fetched = mail.mail_get(client, alias)

    assert fetched["ok"] is True
    assert client.last_call("mail_get")["item_id"] == "mail-1"


def test_alias_miss_returns_validation_error() -> None:
    client = FakeGraphClient()
    result = mail.mail_get(client, "mail:000000")
    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "mail_list" in result["message"]
    assert client.calls == []


def test_mail_delete_rejects_calendar_alias() -> None:
    client = FakeGraphClient()
    calendar_alias = alias_registry.register("cal-1", "cal", namespace="graph:person@outlook.com")

    result = mail.mail_delete(client, calendar_alias, confirm=True)

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "cannot be used as a mail item" in result["message"]
    assert client.calls == []


def test_contact_alias_miss_points_to_contacts_list() -> None:
    client = FakeGraphClient()
    result = contacts.contact_get(client, "contact:000000")

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "contacts_list" in result["message"]
    assert client.calls == []


def test_long_id_passthrough_is_backward_compatible() -> None:
    client = FakeGraphClient()
    fetched = mail.mail_get(client, "AQMkADAwAT-long-base64-id==")
    assert fetched["ok"] is True
    assert client.last_call("mail_get")["item_id"] == "AQMkADAwAT-long-base64-id=="


# --- validate_email ---


def test_validate_email_accepts_valid_addresses() -> None:
    for address in [
        "user@example.com",
        "a@b.co",
        "first.last@sub.example.org",
        "user+tag@domain.com",
        "123@numbers.net",
    ]:
        assert common.validate_email(address) == address


def test_validate_email_rejects_invalid_addresses() -> None:
    for address in [
        "",
        "   ",
        "@",
        "not-an-email",
        "missing@tld",
        "@no-local.com",
        "spaces in@domain.com",
        "missing.domain@",
    ]:
        with pytest.raises(common.ToolValidationError, match="invalid"):
            common.validate_email(address)


def test_validate_email_strips_whitespace() -> None:
    assert common.validate_email("  user@example.com  ") == "user@example.com"


# --- redact_secrets ---


def test_redact_secrets_covers_registered_sensitive_values() -> None:
    register_sensitive_value("SuperSecret123")

    result = common.redact_secrets("auth with SuperSecret123 failed")

    assert "SuperSecret123" not in result
    assert "[redacted]" in result


def test_redact_secrets_leaves_client_id_and_tenant_id_visible() -> None:
    message = "client client-id-value tenant tenant-id-value"

    assert common.redact_secrets(message) == message


# --- run_safely log redaction ---


def test_run_safely_log_output_is_redacted(caplog) -> None:
    import logging

    register_sensitive_value("LogSecret456")

    def action():
        raise RuntimeError("boom LogSecret456")

    with caplog.at_level(logging.ERROR, logger="outlook_mcp"):
        result = common.run_safely(action)

    assert result["ok"] is False
    assert "LogSecret456" not in result["message"]
    assert "LogSecret456" not in caplog.text
