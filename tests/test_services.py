"""Tests for pagination in run_safely (formerly services layer)."""

from __future__ import annotations

from typing import Any

from outlook_mcp.tools import mail as mail_tools
from outlook_mcp.tools import account as account_tools
from outlook_mcp.tools.common import OutcomeUnknownError, run_safely


# --- paginate via run_safely ---


def test_run_safely_does_not_guess_pagination_metadata() -> None:
    def action():
        return {"items": [{"id": 1}, {"id": 2}, {"id": 3}]}

    out = run_safely(action)
    assert out["data"] == {"items": [{"id": 1}, {"id": 2}, {"id": 3}]}


def test_run_safely_preserves_backend_pagination_metadata() -> None:
    def action():
        return {"items": [{"id": 1}, {"id": 2}], "has_more": False}

    out = run_safely(action)
    assert out["data"]["has_more"] is False


def test_run_safely_does_not_invent_total_for_empty_items() -> None:
    def action():
        return {"items": []}

    out = run_safely(action)
    assert out["data"] == {"items": []}


def test_run_safely_paginate_reuses_existing_total() -> None:
    def action():
        return {"items": [1, 2], "total": 42}

    out = run_safely(action)
    assert out["data"]["total"] == 42


def test_run_safely_paginate_skips_error_responses() -> None:
    def action():
        raise ValueError("test error")

    out = run_safely(action)
    assert out["ok"] is False
    assert "has_more" not in out


def test_run_safely_preserves_unknown_write_outcome() -> None:
    def action():
        raise OutcomeUnknownError("sendMail may have completed")

    out = run_safely(action)

    assert out == {
        "ok": False,
        "error": "outcome_unknown",
        "message": "sendMail may have completed",
    }


def test_run_safely_no_paginate_when_limit_none() -> None:
    def action():
        return {"items": [1, 2, 3]}

    out = run_safely(action)
    assert out["data"]["items"] == [1, 2, 3]
    assert "has_more" not in out["data"]


# --- tool integration: paginate built-in via run_safely ---


def test_mail_list_returns_backend_has_more_without_fake_total(monkeypatch) -> None:
    """Verify the EWS adapter computes has_more by fetching one extra row."""
    client = _FakeEWSClient()
    result = mail_tools.mail_list(client, limit=2)
    assert result["ok"] is True
    assert "has_more" in result["data"]
    assert "total" not in result["data"]


def test_mail_list_exactly_full_page_has_no_more_items() -> None:
    client = _FakeEWSClient()

    result = mail_tools.mail_list(client, limit=1)

    assert len(result["data"]["items"]) == 1
    assert result["data"]["has_more"] is False


def test_mail_list_detects_an_extra_item() -> None:
    from datetime import datetime, timezone
    from conftest import FakeItem

    client = _FakeEWSClient()
    client.account.inbox.items.append(
        FakeItem(
            id="mail-2",
            subject="Second message",
            datetime_received=datetime(2026, 1, 2, tzinfo=timezone.utc),
        )
    )

    result = mail_tools.mail_list(client, limit=1)

    assert len(result["data"]["items"]) == 1
    assert result["data"]["has_more"] is True


def test_account_info_no_pagination(monkeypatch) -> None:
    """Verify account_info does not add pagination metadata."""
    client = _FakeEWSClient()
    result = account_tools.account_info(client)
    assert result["ok"] is True
    assert "has_more" not in result.get("data", result)


# --- helper ---

class _FakeEWSClient:
    """A minimal EWS-like client for service tests."""

    def __init__(self):
        from conftest import FakeAccount  # noqa: PLC0415

        self.account = FakeAccount()
        self._folder_cache: dict[str, Any] = {}

    def get_folder(self, name: str) -> Any:
        return getattr(self.account, {"deleted": "trash"}.get(name, name))

    def get_item(self, item_id: str, folder: str | None = None) -> Any:
        if folder:
            return self.get_folder(folder).get(item_id)
        return self.account.root.get(item_id)
