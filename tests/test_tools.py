from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from typing import Any

import exchangelib
import pytest

from outlook_mcp.client import OutlookClient
from outlook_mcp.tools import alias_registry, calendar, common, contacts, mail, tasks
from outlook_mcp.uploads import reset_upload_credential_id, save_uploaded_file, set_upload_credential_id

from conftest import FakeClient, FakeCreatedItem, PNG_BYTES


def _alias(item_id: str, kind: str) -> str:
    return f"{kind}:" + hashlib.sha1(item_id.encode("utf-8")).hexdigest()[:16]


@pytest.fixture(autouse=True)
def _clear_alias_registry():
    alias_registry.clear()
    yield
    alias_registry.clear()


def test_mail_tools_success_paths(monkeypatch) -> None:
    client = FakeClient()
    monkeypatch.setattr(exchangelib, "Message", FakeCreatedItem)

    assert mail.mail_list(client)["data"]["items"][0]["subject"] == "Quarterly report"
    assert mail.mail_search(client, "report")["data"]["items"][0]["id"] == _alias("mail-1", "mail")
    assert mail.mail_get(client, "mail-1")["data"]["body"] == "hello body"
    assert mail.mail_get(client, "mail-1")["data"]["body_content_type"] == "text"

    blocked = mail.mail_send(client, ["a@example.com"], "Hi", "Body")
    assert blocked["ok"] is False
    assert blocked["error"] == "validation_error"

    sent = mail.mail_send(client, ["a@example.com"], "Hi", "Body", confirm=True)
    assert sent["ok"] is True
    assert FakeCreatedItem.created[-1].sent is True

    replied = mail.mail_reply(client, "mail-1", "<p>Thanks</p>", html=True, confirm=True)
    assert replied["ok"] is True
    assert isinstance(client.account.root.get("mail-1").replied["body"], exchangelib.HTMLBody)
    assert str(client.account.root.get("mail-1").replied["body"]) == "<p>Thanks</p>"

    moved = mail.mail_move(client, "mail-1", "trash", confirm=True)
    assert moved["data"]["target_folder"] == "trash"

    deleted = mail.mail_delete(client, "mail-1", mode="move_to_trash", confirm=True)
    assert deleted["data"]["mode"] == "move_to_trash"

    blocked_read = mail.mail_mark_read(client, "mail-1")
    assert blocked_read["ok"] is False
    assert blocked_read["error"] == "validation_error"
    marked_read = mail.mail_mark_read(client, "mail-1", confirm=True)
    assert marked_read["data"]["marked"] == "read"
    assert marked_read["data"]["id"] == _alias("mail-1", "mail")

    blocked_unread = mail.mail_mark_unread(client, "mail-1")
    assert blocked_unread["ok"] is False
    assert blocked_unread["error"] == "validation_error"
    marked_unread = mail.mail_mark_unread(client, "mail-1", confirm=True)
    assert marked_unread["data"]["marked"] == "unread"
    assert marked_unread["data"]["id"] == _alias("mail-1", "mail")


def test_mail_send_ews_supports_upload_inline_attachment(monkeypatch, tmp_path) -> None:
    class CapturedFileAttachment:
        def __init__(self, **kwargs: Any):
            self.__dict__.update(kwargs)

    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(exchangelib, "Message", FakeCreatedItem)
    monkeypatch.setattr(exchangelib, "FileAttachment", CapturedFileAttachment)
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    try:
        sent = mail.mail_send(
            FakeClient(),
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
    assert str(FakeCreatedItem.created[-1].body) == '<img src="cid:photo1" width="320" height="180">'
    attachment = FakeCreatedItem.created[-1].attachments[0]
    assert attachment.name == "photo.png"
    assert attachment.content == PNG_BYTES
    assert attachment.content_type == "image/png"
    assert attachment.content_id == "photo1"
    assert attachment.is_inline is True


def test_mail_send_ews_rejects_inline_upload_without_content_id(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(exchangelib, "Message", FakeCreatedItem)
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    try:
        result = mail.mail_send(
            FakeClient(),
            ["a@example.com"],
            "Hi",
            "<p>Body</p>",
            html=True,
            attachments=[{"relative_path": uploaded.relative_path, "disposition": "inline"}],
            confirm=True,
        )
    finally:
        reset_upload_credential_id(token)

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "content_id" in result["message"]


def test_calendar_tools_success_paths(monkeypatch) -> None:
    client = FakeClient()
    monkeypatch.setattr(exchangelib, "CalendarItem", FakeCreatedItem)
    start = "2026-01-02T00:00:00+00:00"
    end = "2026-01-03T00:00:00+00:00"

    listed = calendar.calendar_list(client, start=start, end=end)
    assert listed["data"]["items"][0]["subject"] == "Planning"
    assert listed["data"]["items"][0]["start"] == "2026-01-02T09:00:00+00:00"
    assert listed["data"]["items"][0]["start_local"] == "2026-01-02T17:00:00+08:00"
    assert listed["data"]["items"][0]["local_timezone"] == "Asia/Shanghai"

    blocked = calendar.calendar_create(client, "Meet", start, end)
    assert blocked["ok"] is False

    created = calendar.calendar_create(client, "Meet", start, end, confirm=True, timezone="America/New_York")
    assert created["data"]["subject"] == "Meet"
    assert created["data"]["start"] == "2026-01-02T00:00:00+00:00"
    assert created["data"]["start_local"] == "2026-01-01T19:00:00-05:00"
    assert created["data"]["local_timezone"] == "America/New_York"
    assert FakeCreatedItem.created[-1].saved is True
    assert FakeCreatedItem.created[-1].reminder_is_set is False

    with_reminder = calendar.calendar_create(
        client,
        "Meet with reminder",
        start,
        end,
        confirm=True,
        reminder_minutes_before_start=10,
    )
    assert with_reminder["data"]["subject"] == "Meet with reminder"
    assert FakeCreatedItem.created[-1].reminder_is_set is True
    assert FakeCreatedItem.created[-1].reminder_minutes_before_start == 10

    with_attachment = calendar.calendar_create(
        client,
        "Meet with file",
        start,
        end,
        html=True,
        attachments=[{"name": "note.txt", "content_base64": "SGk="}],
        confirm=True,
    )
    assert with_attachment["data"]["subject"] == "Meet with file"
    assert FakeCreatedItem.created[-1].attachments[0].name == "note.txt"

    updated = calendar.calendar_update(client, "cal-1", {"location": "Room 2"}, confirm=True)
    assert updated["data"]["location"] == "Room 2"

    deleted = calendar.calendar_delete(client, "cal-1", confirm=True)
    assert deleted["data"]["deleted"] is True


def test_calendar_create_ews_normalizes_inline_upload_image(monkeypatch, tmp_path) -> None:
    class CapturedFileAttachment:
        def __init__(self, **kwargs: Any):
            self.__dict__.update(kwargs)

    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    monkeypatch.setattr(exchangelib, "CalendarItem", FakeCreatedItem)
    monkeypatch.setattr(exchangelib, "FileAttachment", CapturedFileAttachment)
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    try:
        created = calendar.calendar_create(
            FakeClient(),
            "Meet",
            "2026-01-02T00:00:00+00:00",
            "2026-01-03T00:00:00+00:00",
            body='<p><img src="cid:photo1"></p>',
            html=True,
            attachments=[{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "photo1"}],
            confirm=True,
        )
    finally:
        reset_upload_credential_id(token)

    assert created["ok"] is True
    item = FakeCreatedItem.created[-1]
    assert isinstance(item.body, exchangelib.HTMLBody)
    assert item.body == '<p><img src="cid:photo1" width="320" height="180"></p>'
    assert item.attachments[0].is_inline is True


def test_calendar_update_has_no_attachments_parameter() -> None:
    import inspect

    assert "attachments" not in inspect.signature(calendar.calendar_update).parameters


def test_calendar_update_ews_uses_html_body() -> None:
    client = FakeClient()

    result = calendar.calendar_update(
        client,
        "cal-1",
        {"body": "<b>Changed</b>"},
        html=True,
        confirm=True,
    )

    assert result["ok"] is True
    assert isinstance(client.account.calendar.get("cal-1").body, exchangelib.HTMLBody)


def test_mail_attachments_serializes_real_attachment_id() -> None:
    from exchangelib.attachments import AttachmentId

    client = FakeClient()
    item = client.account.root.get("mail-1")
    item.attachments = [
        type(
            "Attachment",
            (),
            {
                "name": "note.txt",
                "size": 2,
                "content_type": "text/plain",
                "attachment_id": AttachmentId(id="attachment-1", root_id="mail-1", root_changekey="ck-1"),
                "is_inline": False,
            },
        )()
    ]

    result = mail.mail_attachments(client, "mail-1")

    assert result["data"]["items"][0]["attachment_id"] == "attachment-1"
    json.dumps(result)


def test_mail_folders_uses_walk_and_returns_round_trip_id() -> None:
    client = OutlookClient(config=None, _account=FakeClient().account)

    result = mail.mail_folders(client)

    inbox = next(item for item in result["data"]["items"] if item["name"] == "inbox")
    trash = next(item for item in result["data"]["items"] if item["name"] == "trash")
    assert inbox["id"] == client.account.inbox.id
    assert client.get_folder(inbox["id"]) is client.account.inbox
    assert mail.mail_list(client, folder=inbox["id"])["ok"] is True
    assert mail.mail_get(client, "mail-1", folder=inbox["id"])["ok"] is True
    assert mail.mail_move(client, "mail-1", trash["id"], confirm=True)["ok"] is True


def test_account_info_reads_real_identity_attributes() -> None:
    from exchangelib import Identity
    from outlook_mcp.tools import account as account_tools

    client = FakeClient()
    client.account.identity = Identity(primary_smtp_address="person@example.com")
    client.account.fullname = "Example Person"

    result = account_tools.account_info(client)

    assert result["data"] == {
        "primary_smtp_address": "person@example.com",
        "display_name": "Example Person",
    }


def test_contact_tools_success_paths(monkeypatch) -> None:
    client = FakeClient()
    monkeypatch.setattr(exchangelib, "Contact", FakeCreatedItem)

    assert contacts.contacts_list(client)["data"]["items"][0]["display_name"] == "Ada Lovelace"
    assert contacts.contacts_search(client, "Ada")["data"]["items"][0]["id"] == _alias("contact-1", "contact")

    blocked = contacts.contact_create(client, "Grace Hopper")
    assert blocked["ok"] is False

    created = contacts.contact_create(
        client,
        "Grace Hopper",
        email_addresses=["grace@example.com"],
        phone_numbers=[{"label": "BusinessPhone", "phone_number": "123"}],
        confirm=True,
    )
    assert created["data"]["display_name"] == "Grace Hopper"

    updated = contacts.contact_update(client, "contact-1", {"job_title": "Engineer"}, confirm=True)
    assert updated["data"]["job_title"] == "Engineer"

    deleted = contacts.contact_delete(client, "contact-1", confirm=True)
    assert deleted["data"]["deleted"] is True
    assert deleted["data"]["id"] == _alias("contact-1", "contact")


def test_task_tools_success_paths(monkeypatch) -> None:
    client = FakeClient()
    monkeypatch.setattr(exchangelib, "Task", FakeCreatedItem)

    assert tasks.tasks_list(client)["data"]["items"][0]["subject"] == "Follow up"

    blocked = tasks.task_create(client, "Write notes")
    assert blocked["ok"] is False

    created = tasks.task_create(
        client,
        "Write notes",
        start_date="2026-01-01",
        due_date="2026-01-02",
        percent_complete=10,
        confirm=True,
    )
    assert created["data"]["subject"] == "Write notes"

    updated = tasks.task_update(client, "task-1", {"percent_complete": 50}, confirm=True)
    assert updated["data"]["percent_complete"] == 50.0

    completed = tasks.task_update(client, "task-1", {"complete": True}, confirm=True)
    assert completed["data"]["status"] == "Completed"

    deleted = tasks.task_delete(client, "task-1", confirm=True)
    assert deleted["data"]["deleted"] is True
    assert deleted["data"]["id"] == _alias("task-1", "task")


def test_validation_edges() -> None:
    client = FakeClient()

    assert mail.mail_list(client, limit=0)["error"] == "validation_error"
    assert calendar.calendar_create(
        client,
        "Bad",
        "2026-01-02T10:00:00+00:00",
        "2026-01-02T09:00:00+00:00",
        confirm=True,
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
    assert tasks.task_create(client, "Bad", percent_complete=101, confirm=True)["error"] == "validation_error"


def test_calendar_tools_parse_naive_datetime_with_calendar_timezone(monkeypatch) -> None:
    client = FakeClient()
    monkeypatch.setattr(exchangelib, "CalendarItem", FakeCreatedItem)

    listed = calendar.calendar_list(
        client,
        start="2026-01-02T16:00:00",
        end="2026-01-02T18:00:00",
        timezone="Asia/Shanghai",
    )
    assert listed["data"]["items"][0]["id"] == _alias("cal-1", "cal")

    created = calendar.calendar_create(
        client,
        "Local meet",
        "2026-03-16T20:00:00",
        "2026-03-16T21:00:00",
        confirm=True,
    )
    assert created["data"]["start"] == "2026-03-16T12:00:00+00:00"
    assert created["data"]["start_local"] == "2026-03-16T20:00:00+08:00"
    assert FakeCreatedItem.created[-1].start.isoformat() == "2026-03-16T20:00:00+08:00"

    offset_created = calendar.calendar_create(
        client,
        "Offset meet",
        "2026-03-16T20:00:00+08:00",
        "2026-03-16T21:00:00+08:00",
        confirm=True,
        timezone="America/New_York",
    )
    assert offset_created["data"]["start"] == "2026-03-16T12:00:00+00:00"
    assert offset_created["data"]["start_local"] == "2026-03-16T08:00:00-04:00"

    updated = calendar.calendar_update(
        client,
        "cal-1",
        {"start": "2026-03-16T20:00:00", "end": "2026-03-16T21:00:00"},
        confirm=True,
        timezone="America/New_York",
    )
    assert updated["data"]["start"] == "2026-03-17T00:00:00+00:00"
    assert updated["data"]["start_local"] == "2026-03-16T20:00:00-04:00"


def test_calendar_datetime_output_is_iso() -> None:
    item = FakeCreatedItem(
        subject="Standalone",
        start=datetime(2026, 1, 1, 12, tzinfo=timezone.utc),
        end=datetime(2026, 1, 1, 13, tzinfo=timezone.utc),
    )

    serialized = calendar.serialize_calendar_item(
        item,
        timezone_name="Asia/Shanghai",
        alias_namespace="ews:test@example.com",
    )
    assert serialized["start"] == "2026-01-01T12:00:00+00:00"
    assert serialized["start_local"] == "2026-01-01T20:00:00+08:00"
    assert serialized["local_timezone"] == "Asia/Shanghai"


def test_alias_round_trip_through_tools() -> None:
    client = FakeClient()
    listed = mail.mail_list(client)
    listed_alias = listed["data"]["items"][0]["id"]
    assert listed_alias == _alias("mail-1", "mail")

    fetched = mail.mail_get(client, listed_alias)
    assert fetched["ok"] is True
    assert fetched["data"]["body"] == "hello body"
    assert fetched["data"]["id"] == listed_alias


def test_alias_mappings_are_isolated_by_account(monkeypatch) -> None:
    from exchangelib import Identity

    class _Hash:
        def hexdigest(self) -> str:
            return "c" * 40

    monkeypatch.setattr(alias_registry.hashlib, "sha1", lambda *args, **kwargs: _Hash())
    first = FakeClient()
    first.account.identity = Identity(primary_smtp_address="first@example.com")
    second = FakeClient()
    second.account.identity = Identity(primary_smtp_address="second@example.com")
    second_mail = second.account.inbox.items[0]
    second_mail.id = "mail-2"
    second_mail.subject = "Second account"

    first_alias = mail.mail_list(first)["data"]["items"][0]["id"]
    second_alias = mail.mail_list(second)["data"]["items"][0]["id"]

    assert first_alias == second_alias
    assert mail.mail_get(first, first_alias)["data"]["subject"] == "Quarterly report"
    assert mail.mail_get(second, second_alias)["data"]["subject"] == "Second account"


def test_alias_miss_returns_validation_error() -> None:
    client = FakeClient()
    result = mail.mail_get(client, "mail:000000")
    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "mail_list" in result["message"]


def test_mail_delete_rejects_calendar_alias() -> None:
    client = FakeClient()
    calendar_alias = calendar.calendar_list(
        client,
        start="2026-01-01T00:00:00Z",
        end="2026-01-03T00:00:00Z",
    )["data"]["items"][0]["id"]
    event = client.account.calendar.get("cal-1")

    result = mail.mail_delete(client, calendar_alias, confirm=True)

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "cannot be used as a mail item" in result["message"]
    assert event.deleted is False


def test_contact_alias_miss_points_to_contacts_list() -> None:
    result = contacts.contact_get(FakeClient(), "contact:000000")

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "contacts_list" in result["message"]


def test_long_id_passthrough_is_backward_compatible() -> None:
    client = FakeClient()
    fetched = mail.mail_get(client, "mail-1")
    assert fetched["ok"] is True
    assert fetched["data"]["body"] == "hello body"
    assert fetched["data"]["id"] == _alias("mail-1", "mail")


# --- to_page optimization ---

def test_to_page_applies_order_by_before_slicing() -> None:
    """order_by should be applied to the queryset directly, not after .all()."""

    class _FakeQueryset:
        def __init__(self):
            self._items = [{"id": 3}, {"id": 1}, {"id": 2}]
            self._ordered = False

        def order_by(self, field):
            self._ordered = True
            self._items = sorted(self._items, key=lambda x: x[field])
            return self

        def __getitem__(self, key):
            return self._items[key]

    qs = _FakeQueryset()
    result = common.to_page(qs, offset=0, limit=2, order_by="id")
    assert result == [{"id": 1}, {"id": 2}]
    assert qs._ordered  # order_by was called on the queryset


def test_to_page_slices_queryset_directly() -> None:
    """Slicing should be applied directly to the queryset, .all() is fallback only."""

    class _TrackingQueryset:
        def __init__(self):
            self.all_called = False
            self.slice_args: tuple | None = None

        def all(self):
            self.all_called = True
            return self

        def order_by(self, field):
            return self

        def __getitem__(self, key):
            self.slice_args = (key.start, key.stop)
            return [{"x": i} for i in range(key.start, key.stop)]

    qs = _TrackingQueryset()
    result = common.to_page(qs, offset=2, limit=3)
    assert len(result) == 3
    assert qs.slice_args == (2, 5)
    assert not qs.all_called  # .all() was NOT called


def test_to_page_falls_back_to_all_on_typeerror() -> None:
    """When queryset doesn't support slicing, fall back to .all() + Python slice."""

    class _NoSliceQueryset:
        def __init__(self):
            self._items = [{"id": i} for i in range(10)]
            self.all_called = False

        def all(self):
            self.all_called = True
            return list(self._items)

        def order_by(self, field):
            return self

        def __getitem__(self, key):
            raise TypeError("slicing not supported")

    qs = _NoSliceQueryset()
    result = common.to_page(qs, offset=3, limit=2)
    assert len(result) == 2
    assert qs.all_called


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


# --- EWS server-side search (Fix 4 & 5) ---


def test_calendar_search_uses_serverside_filter() -> None:
    """EWS calendar_search must use .filter() not in-memory list comprehension."""
    client = FakeClient()
    result = calendar.calendar_search(client, query="Plan", start="2026-01-01T00:00:00Z")
    assert result["ok"] is True
    assert len(result["data"]["items"]) == 1  # Only "Planning" item matches
    assert result["data"]["has_more"] is False
    assert "total" not in result["data"]


def test_calendar_search_no_results() -> None:
    """EWS calendar_search returns empty results for non-matching query."""
    client = FakeClient()
    result = calendar.calendar_search(client, query="NonExistent", start="2026-01-01T00:00:00Z")
    assert result["ok"] is True
    assert result["data"]["items"] == []
    assert result["data"]["has_more"] is False


def test_task_search_uses_serverside_filter() -> None:
    """EWS task_search must use .filter() not in-memory list comprehension."""
    client = FakeClient()
    result = tasks.task_search(client, query="Follow")
    assert result["ok"] is True
    assert len(result["data"]["items"]) >= 1


def test_task_search_with_status_filter() -> None:
    """EWS task_search combined with status filter must use server-side .filter()."""
    client = FakeClient()
    result = tasks.task_search(client, query="Follow", status="NotStarted")
    assert result["ok"] is True
    assert all(it["status"] == "NotStarted" for it in result["data"]["items"])


def test_task_search_no_results() -> None:
    """EWS task_search returns empty results for non-matching query."""
    client = FakeClient()
    result = tasks.task_search(client, query="NonExistentTask")
    assert result["ok"] is True
    assert result["data"]["items"] == []
    assert result["data"]["has_more"] is False


# --- redact_secrets ---


def test_redact_secrets_reads_env_at_call_time(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_PASSWORD", "SuperSecret123")

    result = common.redact_secrets("auth with SuperSecret123 failed")

    assert "SuperSecret123" not in result
    assert "[redacted]" in result


def test_redact_secrets_leaves_client_id_and_tenant_id_visible(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_CLIENT_ID", "client-id-value")
    monkeypatch.setenv("OUTLOOK_TENANT_ID", "tenant-id-value")

    message = "client client-id-value tenant tenant-id-value"

    assert common.redact_secrets(message) == message


# --- mail_forward EWS parity ---


def test_mail_forward_ews_wraps_html_body(monkeypatch) -> None:
    client = FakeClient()
    item = client.account.root.get("mail-1")
    captured: dict[str, Any] = {}

    def fake_create_forward(**kwargs: Any):
        captured.update(kwargs)

        class _Forward:
            def send(self) -> None:
                captured["sent"] = True

        return _Forward()

    monkeypatch.setattr(item, "create_forward", fake_create_forward)

    result = mail.mail_forward(
        client, "mail-1", ["a@example.com"], body="<p>see below</p>", html=True, confirm=True
    )

    assert result["ok"] is True
    assert isinstance(captured["body"], exchangelib.HTMLBody)
    assert str(captured["body"]) == "<p>see below</p>"
    assert captured["sent"] is True


def test_mail_forward_ews_validates_recipients() -> None:
    client = FakeClient()

    result = mail.mail_forward(client, "mail-1", ["not-an-email"], confirm=True)

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "invalid recipient address" in result["message"]


# --- mail_delete mode validation order ---


def test_mail_delete_validates_mode_before_fetching() -> None:
    class NoFetchClient(FakeClient):
        def get_item(self, item_id: str, folder: str | None = None):
            raise AssertionError("get_item must not be called for an invalid mode")

    result = mail.mail_delete(NoFetchClient(), "mail-1", mode="bogus", confirm=True)

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "mode must be" in result["message"]


# --- attachment size limits ---


def test_file_attachment_rejects_oversize_base64() -> None:
    import base64

    content = b"x" * (3 * 1024 * 1024 + 1)
    with pytest.raises(common.ToolValidationError, match="per-attachment limit"):
        common.file_attachment(
            {"name": "big.bin", "content_base64": base64.b64encode(content).decode("ascii")}
        )


def test_file_attachment_accepts_exactly_three_mb() -> None:
    import base64

    content = b"x" * (3 * 1024 * 1024)
    attachment = common.file_attachment(
        {"name": "ok.bin", "content_base64": base64.b64encode(content).decode("ascii")}
    )
    assert attachment.content == content


def test_mail_send_rejects_attachments_over_total_budget(monkeypatch) -> None:
    import base64

    monkeypatch.setattr(exchangelib, "Message", FakeCreatedItem)
    two_mb = base64.b64encode(b"y" * (2 * 1024 * 1024)).decode("ascii")

    result = mail.mail_send(
        FakeClient(),
        ["a@example.com"],
        "Hi",
        "Body",
        attachments=[
            {"name": "a.bin", "content_base64": two_mb},
            {"name": "b.bin", "content_base64": two_mb},
        ],
        confirm=True,
    )

    assert result["ok"] is False
    assert result["error"] == "validation_error"
    assert "per-message limit" in result["message"]
    assert str(2 * 2 * 1024 * 1024) in result["message"]
    assert str(common.MAX_ATTACHMENTS_TOTAL_BYTES) in result["message"]


# --- task_update complete=false ---


def test_task_update_uncompletes_task() -> None:
    client = FakeClient()
    item = client.account.tasks.get("task-1")
    item.complete()

    result = tasks.task_update(client, "task-1", {"complete": False}, confirm=True)

    assert result["ok"] is True
    assert result["data"]["status"] == "NotStarted"
    assert result["data"]["percent_complete"] == 0.0
    assert set(item.save_kwargs["update_fields"]) == {"status", "percent_complete"}


def test_task_update_uncomplete_keeps_other_updates() -> None:
    client = FakeClient()
    item = client.account.tasks.get("task-1")
    item.complete()

    result = tasks.task_update(
        client, "task-1", {"complete": False, "subject": "Reopened"}, confirm=True
    )

    assert result["ok"] is True
    assert result["data"]["subject"] == "Reopened"
    assert result["data"]["status"] == "NotStarted"


# --- run_safely log redaction ---


def test_run_safely_log_output_is_redacted(monkeypatch, caplog) -> None:
    import logging

    monkeypatch.setenv("OUTLOOK_PASSWORD", "LogSecret456")

    def action():
        raise RuntimeError("boom LogSecret456")

    with caplog.at_level(logging.ERROR, logger="outlook_mcp"):
        result = common.run_safely(action)

    assert result["ok"] is False
    assert "LogSecret456" not in result["message"]
    assert "LogSecret456" not in caplog.text


# --- to_page optimization (canonical copy; also see test_credential_store) ---


def test_to_page_optimization_verified() -> None:
    """Sanity check that to_page avoids .all() for slicing-capable querysets."""

    class _SlicingQS:
        def __init__(self, items):
            self._items = items
            self.all_called = False

        def all(self):
            self.all_called = True
            return self._items

        def __getitem__(self, key):
            return self._items[key]

    qs = _SlicingQS(list(range(20)))
    result = common.to_page(qs, offset=5, limit=3)
    assert result == [5, 6, 7]
    assert not qs.all_called


def test_phone_numbers_rejects_unknown_label() -> None:
    with pytest.raises(common.ToolValidationError, match="Unsupported phone number label"):
        contacts._phone_numbers([{"label": "Other", "phone_number": "123"}])
