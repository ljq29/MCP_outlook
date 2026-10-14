from __future__ import annotations

from types import SimpleNamespace
from typing import Any

PNG_BYTES = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (320).to_bytes(4, "big") + (180).to_bytes(4, "big") + b"\x08\x02\x00\x00\x00"


class FakeGraphClient:
    """In-memory stand-in for GraphClient: returns preset dicts and records call kwargs."""

    def __init__(self, email: str = "person@outlook.com", **responses: Any):
        self.config = SimpleNamespace(email=email)
        self.calls: list[tuple[str, dict[str, Any]]] = []
        self._responses = dict(responses)

    def _record(self, name: str, kwargs: dict[str, Any]) -> Any:
        self.calls.append((name, kwargs))
        response = self._responses.get(name)
        if callable(response):
            return response(**kwargs)
        if response is not None:
            return response
        return dict(self._DEFAULT_RESPONSES[name])

    def last_call(self, name: str) -> dict[str, Any]:
        for method, kwargs in reversed(self.calls):
            if method == name:
                return kwargs
        raise AssertionError(f"{name} was not called")

    def mail_list(self, **kwargs: Any) -> Any:
        return self._record("mail_list", kwargs)

    def mail_search(self, **kwargs: Any) -> Any:
        return self._record("mail_search", kwargs)

    def mail_get(self, **kwargs: Any) -> Any:
        return self._record("mail_get", kwargs)

    def mail_send(self, **kwargs: Any) -> Any:
        return self._record("mail_send", kwargs)

    def mail_reply(self, **kwargs: Any) -> Any:
        return self._record("mail_reply", kwargs)

    def mail_move(self, **kwargs: Any) -> Any:
        return self._record("mail_move", kwargs)

    def mail_delete(self, **kwargs: Any) -> Any:
        return self._record("mail_delete", kwargs)

    def mail_forward(self, **kwargs: Any) -> Any:
        return self._record("mail_forward", kwargs)

    def mail_mark_read(self, **kwargs: Any) -> Any:
        return self._record("mail_mark_read", kwargs)

    def mail_mark_unread(self, **kwargs: Any) -> Any:
        return self._record("mail_mark_unread", kwargs)

    def mail_attachments(self, **kwargs: Any) -> Any:
        return self._record("mail_attachments", kwargs)

    def mail_folders(self, **kwargs: Any) -> Any:
        return self._record("mail_folders", kwargs)

    def calendar_list(self, **kwargs: Any) -> Any:
        return self._record("calendar_list", kwargs)

    def calendar_get(self, **kwargs: Any) -> Any:
        return self._record("calendar_get", kwargs)

    def calendar_create(self, **kwargs: Any) -> Any:
        return self._record("calendar_create", kwargs)

    def calendar_update(self, **kwargs: Any) -> Any:
        return self._record("calendar_update", kwargs)

    def calendar_delete(self, **kwargs: Any) -> Any:
        return self._record("calendar_delete", kwargs)

    def calendar_search(self, **kwargs: Any) -> Any:
        return self._record("calendar_search", kwargs)

    def contacts_list(self, **kwargs: Any) -> Any:
        return self._record("contacts_list", kwargs)

    def contacts_search(self, **kwargs: Any) -> Any:
        return self._record("contacts_search", kwargs)

    def contact_get(self, **kwargs: Any) -> Any:
        return self._record("contact_get", kwargs)

    def contact_create(self, **kwargs: Any) -> Any:
        return self._record("contact_create", kwargs)

    def contact_update(self, **kwargs: Any) -> Any:
        return self._record("contact_update", kwargs)

    def contact_delete(self, **kwargs: Any) -> Any:
        return self._record("contact_delete", kwargs)

    def tasks_list(self, **kwargs: Any) -> Any:
        return self._record("tasks_list", kwargs)

    def task_get(self, **kwargs: Any) -> Any:
        return self._record("task_get", kwargs)

    def task_create(self, **kwargs: Any) -> Any:
        return self._record("task_create", kwargs)

    def task_update(self, **kwargs: Any) -> Any:
        return self._record("task_update", kwargs)

    def task_delete(self, **kwargs: Any) -> Any:
        return self._record("task_delete", kwargs)

    def task_search(self, **kwargs: Any) -> Any:
        return self._record("task_search", kwargs)

    def account_info(self, **kwargs: Any) -> Any:
        return self._record("account_info", kwargs)

    _DEFAULT_RESPONSES: dict[str, dict[str, Any]] = {
        "mail_list": {"items": [{"id": "mail-1", "subject": "Quarterly report"}], "has_more": False},
        "mail_search": {"items": [{"id": "mail-1", "subject": "Quarterly report"}], "has_more": False},
        "mail_get": {"id": "mail-1", "subject": "Quarterly report", "body": "hello body"},
        "mail_send": {"sent": True},
        "mail_reply": {"replied": True, "id": "mail-1"},
        "mail_move": {"moved": True, "id": "mail-1"},
        "mail_delete": {"deleted": True},
        "mail_forward": {"forwarded": True},
        "mail_mark_read": {"marked": "read", "id": "mail-1"},
        "mail_mark_unread": {"marked": "unread", "id": "mail-1"},
        "mail_attachments": {"items": [], "has_more": False, "total": 0},
        "mail_folders": {"items": [{"id": "folder-inbox", "name": "inbox"}], "has_more": False, "total": 1},
        "calendar_list": {"items": [{"id": "cal-1", "subject": "Planning"}], "has_more": False},
        "calendar_get": {"id": "cal-1", "subject": "Planning"},
        "calendar_create": {"id": "cal-new", "subject": "Meet"},
        "calendar_update": {"id": "cal-1", "subject": "Planning"},
        "calendar_delete": {"deleted": True, "id": "cal-1"},
        "calendar_search": {"items": [{"id": "cal-1", "subject": "Planning"}], "has_more": False},
        "contacts_list": {"items": [{"id": "contact-1", "display_name": "Ada Lovelace"}], "has_more": False},
        "contacts_search": {"items": [{"id": "contact-1", "display_name": "Ada Lovelace"}], "has_more": False},
        "contact_get": {"id": "contact-1", "display_name": "Ada Lovelace"},
        "contact_create": {"id": "contact-new", "display_name": "Grace Hopper"},
        "contact_update": {"id": "contact-1", "display_name": "Ada Lovelace"},
        "contact_delete": {"deleted": True, "id": "contact-1"},
        "tasks_list": {"items": [{"id": "task-1", "subject": "Follow up"}], "has_more": False},
        "task_get": {"id": "task-1", "subject": "Follow up"},
        "task_create": {"id": "task-new", "subject": "Write notes"},
        "task_update": {"id": "task-1", "subject": "Follow up"},
        "task_delete": {"deleted": True, "id": "task-1"},
        "task_search": {"items": [{"id": "task-1", "subject": "Follow up"}], "has_more": False},
        "account_info": {"primary_smtp_address": "person@outlook.com", "display_name": "Example Person"},
    }


import pytest


@pytest.fixture(autouse=True)
def _reset_global_mutable_state():
    yield
    try:
        from outlook_mcp import http_server

        http_server._UPLOAD_BUCKETS.clear()
        http_server._MAIL_RESPONSE_CACHE.clear()
    except Exception:
        pass
    try:
        from outlook_mcp import server

        with server._http_lock:
            server._close_all_clients_unsafe()
            server._http_clients.clear()
    except Exception:
        pass
