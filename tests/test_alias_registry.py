"""Tests for tools.alias_registry."""

from __future__ import annotations

import hashlib
import re
import threading

import pytest

from outlook_mcp.tools import alias_registry
from outlook_mcp.tools.alias_registry import AliasRegistry
from outlook_mcp.tools.common import ToolValidationError

NS = "ews:user@example.com"


@pytest.fixture(autouse=True)
def _clear_registry():
    alias_registry.clear()
    yield
    alias_registry.clear()


def _expected(item_id: str, kind: str) -> str:
    return f"{kind}:" + hashlib.sha1(item_id.encode("utf-8")).hexdigest()[:16]


def test_alias_format_matches_regex_for_each_kind():
    pattern = re.compile(r"^(?:cal|mail|task|contact):[0-9a-f]{16}$")
    for kind in ("cal", "mail", "task", "contact"):
        alias = alias_registry.register(f"id-{kind}", kind, namespace=NS)
        assert alias is not None
        assert pattern.match(alias)


def test_alias_is_deterministic():
    a1 = alias_registry.register("AAAA=long+id/value==", "mail", namespace=NS)
    a2 = alias_registry.register("AAAA=long+id/value==", "mail", namespace=NS)
    assert a1 == a2


def test_alias_differs_across_kinds_for_same_id():
    a_mail = alias_registry.register("same-id", "mail", namespace=NS)
    a_cal = alias_registry.register("same-id", "cal", namespace=NS)
    assert a_mail != a_cal
    assert a_mail.split(":")[1] == a_cal.split(":")[1]  # same hex


def test_register_returns_none_for_empty_or_none():
    assert alias_registry.register(None, "mail", namespace=NS) is None
    assert alias_registry.register("", "cal", namespace=NS) is None


def test_register_rejects_unknown_kind():
    with pytest.raises(ValueError):
        alias_registry.register("anything", "bogus", namespace=NS)


def test_is_alias_truth_table():
    assert alias_registry.is_alias("cal:abcdef0123456789")
    assert alias_registry.is_alias("mail:0000000000000000")
    assert alias_registry.is_alias("task:0a1b2c3d4e5f6789")
    assert alias_registry.is_alias("contact:ffffffffffffffff")
    assert alias_registry.is_alias("cal:abcdef")  # stale six-char aliases fail closed
    assert not alias_registry.is_alias("cal:ABCDEF0123456789")  # uppercase rejected
    assert not alias_registry.is_alias("cal:abcde")  # too short
    assert not alias_registry.is_alias("cal:abcdefg")  # unsupported length
    assert not alias_registry.is_alias("foo:abcdef0123456789")  # unknown prefix
    assert not alias_registry.is_alias("AQMkADAwAT...long-base64==")  # raw EWS id
    assert not alias_registry.is_alias("")
    assert not alias_registry.is_alias(None)
    assert not alias_registry.is_alias(123)


def test_resolve_alias_round_trip():
    alias = alias_registry.register("real-id-1", "mail", namespace=NS)
    assert alias_registry.resolve_or_passthrough(
        alias, kind="mail", tool_name="mail", namespace=NS
    ) == "real-id-1"


def test_resolve_passthrough_for_non_alias():
    raw = "AQMkADAwAT-long-base64-id=="
    assert alias_registry.resolve_or_passthrough(
        raw, kind="mail", tool_name="mail", namespace=NS
    ) == raw


def test_resolve_alias_miss_raises_validation_error():
    with pytest.raises(ToolValidationError) as exc:
        alias_registry.resolve_or_passthrough(
            "cal:000000", kind="cal", tool_name="calendar", namespace=NS
        )
    assert "calendar_list" in str(exc.value)
    assert "cal:000000" in str(exc.value)


def test_resolve_rejects_alias_from_another_kind():
    alias = alias_registry.register("calendar-item", "cal", namespace=NS)

    with pytest.raises(ToolValidationError, match="cannot be used as a mail item"):
        alias_registry.resolve_or_passthrough(
            alias, kind="mail", tool_name="mail", namespace=NS
        )


def test_collision_refuses_to_overwrite_existing_mapping(monkeypatch):
    class _Hash:
        def hexdigest(self) -> str:
            return "a" * 40

    monkeypatch.setattr(alias_registry.hashlib, "sha1", lambda *args, **kwargs: _Hash())
    reg = AliasRegistry()
    alias = reg.register("first", "mail", namespace=NS)

    with pytest.raises(alias_registry.AliasCollisionError, match="refusing to replace"):
        reg.register("second", "mail", namespace=NS)

    assert reg.resolve(alias, namespace=NS) == "first"


def test_resolve_handles_empty_value():
    assert alias_registry.resolve_or_passthrough(
        "", kind="mail", tool_name="mail", namespace=NS
    ) == ""


def test_lru_eviction_at_capacity():
    reg = AliasRegistry(capacity=3)
    a1 = reg.register("id-1", "mail", namespace=NS)
    a2 = reg.register("id-2", "mail", namespace=NS)
    a3 = reg.register("id-3", "mail", namespace=NS)
    assert len(reg) == 3
    reg.register("id-4", "mail", namespace=NS)
    assert len(reg) == 3
    with pytest.raises(KeyError):
        reg.resolve(a1, namespace=NS)
    assert reg.resolve(a2, namespace=NS) == "id-2"
    assert reg.resolve(a3, namespace=NS) == "id-3"


def test_lru_resolve_refreshes_recency():
    reg = AliasRegistry(capacity=3)
    a1 = reg.register("id-1", "mail", namespace=NS)
    reg.register("id-2", "mail", namespace=NS)
    reg.register("id-3", "mail", namespace=NS)
    reg.resolve(a1, namespace=NS)  # touch id-1
    reg.register("id-4", "mail", namespace=NS)  # should evict id-2 now (oldest)
    assert reg.resolve(a1, namespace=NS) == "id-1"


def test_namespaces_have_independent_collision_and_lru_state(monkeypatch):
    class _Hash:
        def hexdigest(self) -> str:
            return "b" * 40

    monkeypatch.setattr(alias_registry.hashlib, "sha1", lambda *args, **kwargs: _Hash())
    reg = AliasRegistry(capacity=1)
    first_alias = reg.register("first", "mail", namespace="ews:first@example.com")
    second_alias = reg.register("second", "mail", namespace="ews:second@example.com")

    assert first_alias == second_alias
    assert reg.resolve(first_alias, namespace="ews:first@example.com") == "first"
    assert reg.resolve(second_alias, namespace="ews:second@example.com") == "second"

    reg.register("replacement", "cal", namespace="ews:first@example.com")
    with pytest.raises(KeyError):
        reg.resolve(first_alias, namespace="ews:first@example.com")
    assert reg.resolve(second_alias, namespace="ews:second@example.com") == "second"


def test_capacity_rejects_zero_and_negative():
    with pytest.raises(ValueError):
        AliasRegistry(capacity=0)
    with pytest.raises(ValueError):
        AliasRegistry(capacity=-1)


def test_thread_safety_smoke():
    reg = AliasRegistry(capacity=10000)
    errors: list[Exception] = []
    aliases: list[tuple[str, str]] = []
    aliases_lock = threading.Lock()

    def worker(start: int) -> None:
        try:
            local: list[tuple[str, str]] = []
            for i in range(start, start + 200):
                real = f"id-{i}"
                alias = reg.register(real, "mail", namespace=NS)
                local.append((alias, real))
            with aliases_lock:
                aliases.extend(local)
        except Exception as exc:
            errors.append(exc)

    threads = [threading.Thread(target=worker, args=(i * 200,)) for i in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert not errors
    # Each unique alias maps to exactly one of its colliders' real_ids.
    # Just verify no crashes and final size matches unique alias count.
    unique_aliases = {alias for alias, _ in aliases}
    assert len(reg) == len(unique_aliases)


def test_clear_empties_registry():
    alias_registry.register("id-x", "mail", namespace=NS)
    assert alias_registry.size() == 1
    alias_registry.clear()
    assert alias_registry.size() == 0


def test_expected_hash_value():
    alias = alias_registry.register("real-id-1", "mail", namespace=NS)
    assert alias == _expected("real-id-1", "mail")


# --- legacy alias hint ---


def test_resolve_legacy_alias_miss_explains_expiration():
    with pytest.raises(ToolValidationError) as exc:
        alias_registry.resolve_or_passthrough(
            "cal:000000", kind="cal", tool_name="calendar", namespace=NS
        )
    message = str(exc.value)
    assert "retired 6-character format" in message
    assert "calendar_list" in message
    assert "cal:000000" in message
