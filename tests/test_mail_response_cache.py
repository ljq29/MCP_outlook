from __future__ import annotations

from outlook_mcp.mail_cache import MailResponseCache


def test_mail_response_cache_returns_deep_copies_and_expires() -> None:
    now = [100.0]
    cache = MailResponseCache(max_entries=4, clock=lambda: now[0])
    params = {"folder": "inbox", "limit": 20, "offset": 0}
    original = {"ok": True, "data": {"items": [{"id": "mail-1"}]}}

    cache.put(11, "mail/list", params, original, ttl_seconds=30)
    cached = cache.get(11, "mail/list", params)

    assert cached == original
    assert cached is not original
    cached["data"]["items"].append({"id": "mutated"})
    assert cache.get(11, "mail/list", params) == original

    now[0] = 130.0
    assert cache.get(11, "mail/list", params) is None


def test_mail_response_cache_isolates_credentials_and_invalidates_namespace() -> None:
    cache = MailResponseCache(max_entries=4)
    params = {"item_id": "mail-1", "include_body": True}
    first = {"ok": True, "data": {"body": "first"}}
    second = {"ok": True, "data": {"body": "second"}}

    cache.put(11, "mail/get", params, first, ttl_seconds=60)
    cache.put(22, "mail/get", params, second, ttl_seconds=60)

    assert cache.get(11, "mail/get", params) == first
    assert cache.get(22, "mail/get", params) == second

    cache.invalidate(11)

    assert cache.get(11, "mail/get", params) is None
    assert cache.get(22, "mail/get", params) == second


def test_mail_response_cache_evicts_least_recently_used_entry() -> None:
    cache = MailResponseCache(max_entries=2)

    cache.put(11, "mail/get", {"item_id": "one"}, {"data": {"id": "one"}}, ttl_seconds=60)
    cache.put(11, "mail/get", {"item_id": "two"}, {"data": {"id": "two"}}, ttl_seconds=60)
    assert cache.get(11, "mail/get", {"item_id": "one"}) is not None

    cache.put(11, "mail/get", {"item_id": "three"}, {"data": {"id": "three"}}, ttl_seconds=60)

    assert cache.get(11, "mail/get", {"item_id": "one"}) is not None
    assert cache.get(11, "mail/get", {"item_id": "two"}) is None
    assert cache.get(11, "mail/get", {"item_id": "three"}) is not None


def test_mail_response_cache_enforces_entry_and_total_byte_limits() -> None:
    cache = MailResponseCache(max_entries=10, max_bytes=100, max_entry_bytes=60)
    oversized = {"data": {"body": "x" * 100}}

    cache.put(11, "mail/get", {"item_id": "oversized"}, oversized, ttl_seconds=60)

    assert cache.get(11, "mail/get", {"item_id": "oversized"}) is None

    cache.put(11, "mail/get", {"item_id": "one"}, {"data": {"body": "1" * 20}}, ttl_seconds=60)
    cache.put(11, "mail/get", {"item_id": "two"}, {"data": {"body": "2" * 20}}, ttl_seconds=60)
    cache.put(11, "mail/get", {"item_id": "three"}, {"data": {"body": "3" * 20}}, ttl_seconds=60)

    assert cache.get(11, "mail/get", {"item_id": "one"}) is None
    assert cache.get(11, "mail/get", {"item_id": "two"}) is not None
    assert cache.get(11, "mail/get", {"item_id": "three"}) is not None
