from __future__ import annotations

# NOTE: FakeItem/FakeQuery 等 fake 对象的方法签名（reply/save/complete/create_forward 等）
# 是对照 src/outlook_mcp/tools/ 的调用方式手写的，并未验证 exchangelib 的真实契约。
# 升级 exchangelib 版本时需要同步核对这些签名。
from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Any

PNG_BYTES = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + (320).to_bytes(4, "big") + (180).to_bytes(4, "big") + b"\x08\x02\x00\x00\x00"


class FakeQuery:
    def __init__(self, items: list[Any]):
        self.items = list(items)

    def all(self) -> "FakeQuery":
        return self

    def order_by(self, field: str) -> "FakeQuery":
        reverse = field.startswith("-")
        attr = field[1:] if reverse else field
        return FakeQuery(sorted(self.items, key=lambda item: getattr(item, attr, None) or "", reverse=reverse))

    def filter(self, **criteria: Any) -> "FakeQuery":
        result = self.items
        for key, expected in criteria.items():
            if key.endswith("__contains"):
                attr = key.removesuffix("__contains")
                result = [item for item in result if expected.lower() in str(getattr(item, attr, "")).lower()]
            else:
                result = [item for item in result if getattr(item, key, None) == expected]
        return FakeQuery(result)

    def __getitem__(self, index: Any) -> Any:
        return self.items[index]

    def __iter__(self):
        return iter(self.items)


class FakeFolder(FakeQuery):
    def __init__(
        self,
        name: str,
        items: list[Any] | None = None,
        *,
        folder_id: str | None = None,
        folder_class: str = "IPF.Note",
        child_folders: list["FakeFolder"] | None = None,
    ):
        super().__init__(items or [])
        self.name = name
        self.id = folder_id or f"folder-{name}"
        self.folder_class = folder_class
        self.total_count = len(self.items)
        self.unread_count = 0
        self.child_folder_count = len(child_folders or [])
        self.child_folders = list(child_folders or [])

    def get(self, id: str) -> Any:
        for item in self.items:
            if item.id == id:
                return item
        raise KeyError(id)

    def view(self, start: datetime, end: datetime) -> FakeQuery:
        return FakeQuery([item for item in self.items if start <= item.start < end])

    def walk(self) -> list["FakeFolder"]:
        return list(self.child_folders)


class FakeItem:
    def __init__(self, **kwargs: Any):
        self.id = kwargs.pop("id", "item-1")
        self.changekey = kwargs.pop("changekey", "ck-1")
        self.saved = False
        self.deleted = False
        self.moved_to = None
        self.replied = None
        self.completed = False
        for key, value in kwargs.items():
            setattr(self, key, value)

    def save(self, **kwargs: Any) -> None:
        self.saved = True
        self.save_kwargs = kwargs

    def delete(self, **kwargs: Any) -> None:
        self.deleted = True
        self.delete_kwargs = kwargs

    def soft_delete(self) -> None:
        self.deleted = "soft"

    def move_to_trash(self) -> None:
        self.deleted = "trash"

    def move(self, folder: FakeFolder) -> "FakeItem":
        self.moved_to = folder.name
        return self

    def reply(self, **kwargs: Any) -> None:
        self.replied = kwargs

    def complete(self) -> None:
        self.completed = True
        self.status = "Completed"
        self.percent_complete = Decimal(100)

    def create_forward(self, **kwargs: Any) -> "FakeItem":
        fwd = FakeItem(id="fwd-1", subject=kwargs.get("subject", ""))
        fwd.sent = False
        return fwd

    def send(self) -> None:
        self.sent = True


class FakeCreatedItem(FakeItem):
    created: list["FakeCreatedItem"] = []

    def __init__(self, **kwargs: Any):
        super().__init__(id="new-id", **kwargs)
        self.attachments = []
        FakeCreatedItem.created.append(self)

    def attach(self, attachment: Any) -> None:
        self.attachments.append(attachment)

    def send_and_save(self) -> None:
        self.sent = True


class FakeAccount:
    def __init__(self):
        mail = FakeItem(
            id="mail-1",
            subject="Quarterly report",
            body="hello body",
            datetime_received=datetime(2026, 1, 1, tzinfo=timezone.utc),
            is_read=False,
            has_attachments=False,
        )
        event = FakeItem(
            id="cal-1",
            subject="Planning",
            start=datetime(2026, 1, 2, 9, tzinfo=timezone.utc),
            end=datetime(2026, 1, 2, 10, tzinfo=timezone.utc),
            location="Room 1",
        )
        contact = FakeItem(id="contact-1", display_name="Ada Lovelace", email_addresses=[])
        task = FakeItem(
            id="task-1",
            subject="Follow up",
            due_date=date(2026, 1, 3),
            status="NotStarted",
            percent_complete=Decimal(0),
        )
        self.inbox = FakeFolder("inbox", [mail])
        self.sent = FakeFolder("sent", [])
        self.trash = FakeFolder("trash", [])
        self.calendar = FakeFolder("calendar", [event], folder_class="IPF.Appointment")
        self.contacts = FakeFolder("contacts", [contact], folder_class="IPF.Contact")
        self.tasks = FakeFolder("tasks", [task], folder_class="IPF.Task")
        self.root = FakeFolder(
            "root",
            [mail, event, contact, task],
            folder_class="IPF.Root",
            child_folders=[
                self.inbox,
                self.sent,
                self.trash,
                self.calendar,
                self.contacts,
                self.tasks,
            ],
        )


class FakeClient:
    def __init__(self):
        self.account = FakeAccount()

    def get_folder(self, name: str):
        return getattr(self.account, {"deleted": "trash"}.get(name, name))

    def get_item(self, item_id: str, folder: str | None = None):
        if folder:
            return self.get_folder(folder).get(item_id)
        return self.account.root.get(item_id)


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
