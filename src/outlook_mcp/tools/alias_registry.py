"""Module-level LRU registry mapping short aliases to long Graph item IDs.

Long Graph IDs (150+ chars) get truncated or mangled when
LLMs copy them across tool calls. This registry replaces them with short,
deterministic aliases like ``cal:a3f29b1c07d3e529`` so LLM clients see a
fixed-width string they can round-trip without errors.

Aliases are deterministic: ``sha1(item_id)[:16]`` prefixed by kind. Same
``item_id`` always produces the same alias, so a server restart followed by a
fresh ``*_list`` call rebuilds the same mapping.

Tools accept BOTH aliases and raw long IDs. Raw IDs (anything not matching the
alias regex) pass through unchanged for backward compatibility.
"""

from __future__ import annotations

import hashlib
import re
import threading
from collections import OrderedDict

KINDS: tuple[str, ...] = ("cal", "mail", "task", "contact")
_ALIAS_RE = re.compile(r"^(?P<kind>cal|mail|task|contact):[0-9a-f]{16}$")
_LEGACY_ALIAS_RE = re.compile(r"^(?P<kind>cal|mail|task|contact):[0-9a-f]{6}$")
_DEFAULT_CAPACITY = 1000


class AliasCollisionError(RuntimeError):
    """Raised when two item IDs produce the same alias."""


class AliasRegistry:
    """Thread-safe, per-account LRU maps of alias -> raw item_id."""

    def __init__(self, capacity: int = _DEFAULT_CAPACITY) -> None:
        if capacity <= 0:
            raise ValueError("capacity must be positive")
        self._capacity = capacity
        self._stores: dict[str, OrderedDict[str, str]] = {}
        self._lock = threading.Lock()

    def register(self, item_id: str | None, kind: str, *, namespace: str) -> str | None:
        if kind not in KINDS:
            raise ValueError(f"Unsupported alias kind: {kind!r}")
        _validate_namespace(namespace)
        if not item_id or not isinstance(item_id, str):
            return None
        digest = hashlib.sha1(item_id.encode("utf-8"), usedforsecurity=False).hexdigest()[:16]
        alias = f"{kind}:{digest}"
        with self._lock:
            store = self._stores.setdefault(namespace, OrderedDict())
            existing = store.get(alias)
            if existing is not None and existing != item_id:
                raise AliasCollisionError(
                    f"Alias collision for {alias!r}; refusing to replace the existing item mapping"
                )
            store[alias] = item_id
            store.move_to_end(alias)
            while len(store) > self._capacity:
                store.popitem(last=False)
        return alias

    def resolve(self, alias: str, *, namespace: str) -> str:
        _validate_namespace(namespace)
        with self._lock:
            store = self._stores.get(namespace)
            value = store.get(alias) if store is not None else None
            if value is None:
                raise KeyError(alias)
            store.move_to_end(alias)
            return value

    def is_alias(self, value: object) -> bool:
        return isinstance(value, str) and (
            _ALIAS_RE.fullmatch(value) is not None
            or _LEGACY_ALIAS_RE.fullmatch(value) is not None
        )

    def clear(self, *, namespace: str | None = None) -> None:
        with self._lock:
            if namespace is None:
                self._stores.clear()
            else:
                self._stores.pop(namespace, None)

    def __len__(self) -> int:
        with self._lock:
            return sum(len(store) for store in self._stores.values())


_registry = AliasRegistry()


def register(item_id: str | None, kind: str, *, namespace: str) -> str | None:
    """Compute and store the alias for ``item_id``. Returns the alias or ``None``."""
    return _registry.register(item_id, kind, namespace=namespace)


def is_alias(value: object) -> bool:
    """True iff ``value`` matches the alias regex."""
    return _registry.is_alias(value)


def resolve_or_passthrough(
    value: str,
    *,
    kind: str,
    tool_name: str,
    namespace: str,
) -> str:
    """Resolve an alias to the raw item_id, or pass through non-alias values.

    - Alias shape and present in cache -> return raw item_id.
    - Alias shape but missing from cache -> raise ``ToolValidationError`` with a
      hint to refresh via the corresponding ``*_list`` tool.
    - Anything else -> return unchanged (backward-compatible passthrough for
      callers that already hold the long Graph id).
    """
    if kind not in KINDS:
        raise ValueError(f"Unsupported alias kind: {kind!r}")
    _validate_namespace(namespace)
    if not isinstance(value, str) or not value:
        return value
    match = _ALIAS_RE.fullmatch(value) or _LEGACY_ALIAS_RE.fullmatch(value)
    if match is None:
        return value
    alias_kind = match.group("kind")
    if alias_kind != kind:
        from .common import ToolValidationError

        raise ToolValidationError(
            f"Alias {value!r} identifies a {alias_kind} item and cannot be used as a {kind} item"
        )
    try:
        return _registry.resolve(value, namespace=namespace)
    except KeyError:
        from .common import ToolValidationError

        if _LEGACY_ALIAS_RE.fullmatch(value):
            raise ToolValidationError(
                f"Alias {value!r} uses the retired 6-character format and has expired; "
                f"call {tool_name}_list again and use the new id it returns"
            ) from None
        raise ToolValidationError(
            f"Alias {value!r} not found, please call {tool_name}_list to refresh"
        ) from None


def clear() -> None:
    """Clear all entries. Intended for tests."""
    _registry.clear()


def size() -> int:
    """Return current number of cached aliases. Intended for tests."""
    return len(_registry)


def namespace_for_client(client: object) -> str:
    """Return a stable backend/account namespace for alias mappings."""
    config = getattr(client, "config", None)
    configured_email = getattr(config, "email", None)
    backend = "graph"
    if configured_email:
        return f"{backend}:{str(configured_email).strip().casefold()}"

    account = getattr(client, "account", None)
    identity = getattr(account, "identity", None)
    account_email = (
        getattr(identity, "primary_smtp_address", None)
        or getattr(account, "primary_smtp_address", None)
    )
    if account_email:
        return f"{backend}:{str(account_email).strip().casefold()}"
    return f"{backend}:object:{id(account if account is not None else client)}"


def _validate_namespace(namespace: str) -> None:
    if not isinstance(namespace, str) or not namespace.strip():
        raise ValueError("alias namespace must be a non-empty string")
