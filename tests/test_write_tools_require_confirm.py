from __future__ import annotations

from typing import Any

import pytest

from outlook_mcp.tools import calendar, contacts, mail, tasks

from conftest import FakeGraphClient


class NoTouchClient(FakeGraphClient):
    def _record(self, name: str, kwargs: dict[str, Any]) -> Any:
        raise AssertionError(f"client.{name} must not be called without confirm")


WRITE_CASES = [
    ("mail_send", mail.mail_send, (["a@example.com"], "Hi", "Body")),
    ("mail_reply", mail.mail_reply, ("mail-1", "Thanks")),
    ("mail_forward", mail.mail_forward, ("mail-1", ["a@example.com"])),
    ("mail_move", mail.mail_move, ("mail-1", "trash")),
    ("mail_delete", mail.mail_delete, ("mail-1",)),
    ("mail_mark_read", mail.mail_mark_read, ("mail-1",)),
    ("mail_mark_unread", mail.mail_mark_unread, ("mail-1",)),
    ("calendar_create", calendar.calendar_create, ("Meeting", "2026-01-02 09:00", "2026-01-02 10:00")),
    ("calendar_update", calendar.calendar_update, ("cal-1", {"subject": "New"})),
    ("calendar_delete", calendar.calendar_delete, ("cal-1",)),
    ("contact_create", contacts.contact_create, ("Ada",)),
    ("contact_update", contacts.contact_update, ("contact-1", {"display_name": "Ada"})),
    ("contact_delete", contacts.contact_delete, ("contact-1",)),
    ("task_create", tasks.task_create, ("Follow up",)),
    ("task_update", tasks.task_update, ("task-1", {"subject": "New"})),
    ("task_delete", tasks.task_delete, ("task-1",)),
]


@pytest.mark.parametrize("tool_name,tool,args", WRITE_CASES, ids=[case[0] for case in WRITE_CASES])
def test_write_tool_requires_confirm_by_default(tool_name: str, tool: Any, args: tuple) -> None:
    result = tool(NoTouchClient(), *args)
    assert result["ok"] is False, tool_name
    assert result["error"] == "validation_error", tool_name
    assert "confirm" in result["message"], tool_name


@pytest.mark.parametrize("tool_name,tool,args", WRITE_CASES, ids=[case[0] for case in WRITE_CASES])
def test_write_tool_rejects_confirm_false(tool_name: str, tool: Any, args: tuple) -> None:
    result = tool(NoTouchClient(), *args, confirm=False)
    assert result["ok"] is False, tool_name
    assert result["error"] == "validation_error", tool_name
