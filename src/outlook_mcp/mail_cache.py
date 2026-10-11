"""Short-lived response cache for the HTTP Mail WebUI."""

from __future__ import annotations

import json
import threading
import time
from collections import OrderedDict
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable


CacheKey = tuple[int, str, str]


@dataclass(frozen=True)
class _CacheEntry:
    expires_at: float
    size_bytes: int
    value: dict[str, Any]


class MailResponseCache:
    """Thread-safe LRU cache isolated by HTTP credential id."""

    def __init__(
        self,
        *,
        max_entries: int = 256,
        max_bytes: int = 32 * 1024 * 1024,
        max_entry_bytes: int = 512 * 1024,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        if max_entries <= 0:
            raise ValueError("max_entries must be greater than 0")
        if max_bytes <= 0 or max_entry_bytes <= 0:
            raise ValueError("cache byte limits must be greater than 0")
        self._max_entries = max_entries
        self._max_bytes = max_bytes
        self._max_entry_bytes = max_entry_bytes
        self._clock = clock
        self._entries: OrderedDict[CacheKey, _CacheEntry] = OrderedDict()
        self._total_bytes = 0
        self._lock = threading.Lock()

    @staticmethod
    def _key(namespace: int, operation: str, params: dict[str, Any]) -> CacheKey:
        canonical_params = json.dumps(
            params,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        return namespace, operation, canonical_params

    def get(
        self,
        namespace: int,
        operation: str,
        params: dict[str, Any],
    ) -> dict[str, Any] | None:
        key = self._key(namespace, operation, params)
        now = self._clock()
        with self._lock:
            entry = self._entries.get(key)
            if entry is None:
                return None
            if entry.expires_at <= now:
                self._total_bytes -= entry.size_bytes
                del self._entries[key]
                return None
            self._entries.move_to_end(key)
            return deepcopy(entry.value)

    def put(
        self,
        namespace: int,
        operation: str,
        params: dict[str, Any],
        value: dict[str, Any],
        *,
        ttl_seconds: float,
    ) -> None:
        if ttl_seconds <= 0:
            return
        size_bytes = len(
            json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        )
        if size_bytes > self._max_entry_bytes or size_bytes > self._max_bytes:
            return
        key = self._key(namespace, operation, params)
        now = self._clock()
        with self._lock:
            expired = [cache_key for cache_key, entry in self._entries.items() if entry.expires_at <= now]
            for cache_key in expired:
                self._total_bytes -= self._entries.pop(cache_key).size_bytes
            previous = self._entries.pop(key, None)
            if previous is not None:
                self._total_bytes -= previous.size_bytes
            self._entries[key] = _CacheEntry(now + ttl_seconds, size_bytes, deepcopy(value))
            self._total_bytes += size_bytes
            self._entries.move_to_end(key)
            while len(self._entries) > self._max_entries or self._total_bytes > self._max_bytes:
                _, evicted = self._entries.popitem(last=False)
                self._total_bytes -= evicted.size_bytes

    def invalidate(self, namespace: int) -> None:
        with self._lock:
            stale_keys = [key for key in self._entries if key[0] == namespace]
            for key in stale_keys:
                self._total_bytes -= self._entries.pop(key).size_bytes

    def clear(self) -> None:
        with self._lock:
            self._entries.clear()
            self._total_bytes = 0
