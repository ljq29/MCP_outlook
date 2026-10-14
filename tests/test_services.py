"""Tests for pagination in run_safely (formerly services layer)."""

from __future__ import annotations

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
