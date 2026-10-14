"""Configuration loading for the Outlook MCP server."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping


class ConfigError(ValueError):
    """Raised when required Outlook MCP configuration is missing or invalid."""


EMAIL_RE = re.compile(
    r"^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?"
    r"(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$"
)


_SENSITIVE_VALUES: set[str] = set()


def register_sensitive_value(value: str | None) -> None:
    """Register a loaded secret so error messages can redact it even outside env vars."""
    if value and len(value) >= 8:
        _SENSITIVE_VALUES.add(value)


def sensitive_values() -> tuple[str, ...]:
    return tuple(_SENSITIVE_VALUES)


DEFAULT_GRAPH_SCOPES = (
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
)


@dataclass(frozen=True)
class GraphConfig:
    email: str
    client_id: str
    tenant: str = "consumers"
    token_cache_path: str = "~/.outlook-mcp/msal-token-cache.json"
    scopes: tuple[str, ...] = DEFAULT_GRAPH_SCOPES
    provider: str = "graph"
    graph_base_url: str = "https://graph.microsoft.com/v1.0"

    @classmethod
    def from_json(cls, payload: Mapping[str, object], env: Mapping[str, str] | None = None) -> "GraphConfig":
        source = env or os.environ
        graph = _section(payload, "graph")
        email = _required_json(payload, "email")
        client_id = source.get("OUTLOOK_GRAPH_CLIENT_ID", "").strip() or _optional_json(graph, "client_id") or ""
        scopes_value = graph.get("scopes", DEFAULT_GRAPH_SCOPES)
        if not isinstance(scopes_value, list | tuple):
            raise ConfigError("graph.scopes must be a list of strings")
        scopes = tuple(str(scope).strip() for scope in scopes_value if str(scope).strip())
        if not scopes:
            raise ConfigError("graph.scopes cannot be empty")
        config = cls(
            email=email,
            client_id=client_id,
            tenant=(graph.get("tenant") or "consumers").strip() or "consumers",
            token_cache_path=str(graph.get("token_cache_path", "~/.outlook-mcp/msal-token-cache.json")).strip()
            or "~/.outlook-mcp/msal-token-cache.json",
            scopes=scopes,
            graph_base_url=str(graph.get("graph_base_url", "https://graph.microsoft.com/v1.0")).rstrip("/"),
        )
        config.validate()
        return config

    def validate(self) -> None:
        if not self.email:
            raise ConfigError("email is required")
        if not EMAIL_RE.match(self.email):
            raise ConfigError(f"invalid email: {self.email}")
        if not self.client_id:
            raise ConfigError("graph.client_id is required")
        if not self.tenant.strip():
            raise ConfigError("graph.tenant is required")

    @property
    def cache_path(self) -> Path:
        path = Path(self.token_cache_path).expanduser()
        return path

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant}"

    @property
    def redacted(self) -> dict[str, object]:
        return {
            "provider": self.provider,
            "email": self.email,
            "client_id": self.client_id,
            "tenant": self.tenant,
            "token_cache_path": self.token_cache_path,
            "scopes": list(self.scopes),
            "graph_base_url": self.graph_base_url,
        }


RuntimeConfig = GraphConfig


def load_config(env: Mapping[str, str] | None = None, config_path: str | Path | None = None) -> RuntimeConfig:
    source = env or os.environ
    candidate = _config_path(source, config_path)
    if candidate is not None:
        return load_config_file(candidate, env=source)
    searched = Path.cwd() / "outlook-mcp.json"
    raise ConfigError(
        f"No config file found. Searched for outlook-mcp.json in: {searched}; "
        "set OUTLOOK_CONFIG_FILE or create outlook-mcp.json from outlook-mcp.example.json "
        'with provider "graph"'
    )


def load_config_file(path: str | Path, env: Mapping[str, str] | None = None) -> RuntimeConfig:
    config_path = Path(path).expanduser()
    if not config_path.exists():
        raise ConfigError(f"Config file does not exist: {config_path}")
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Invalid JSON config file: {config_path}") from exc
    if not isinstance(payload, dict):
        raise ConfigError("Config file must contain a JSON object")
    provider = str(payload.get("provider", "")).strip().lower()
    if provider == "graph":
        return GraphConfig.from_json(payload, env=env)
    if provider == "ews":
        raise ConfigError("EWS support was removed; use provider=graph or pin an older version")
    raise ConfigError("provider must be 'graph'")


def _config_path(env: Mapping[str, str], explicit: str | Path | None) -> Path | None:
    if explicit is not None:
        return Path(explicit)
    env_path = env.get("OUTLOOK_CONFIG_FILE", "").strip()
    if env_path:
        return Path(env_path)
    default_path = Path.cwd() / "outlook-mcp.json"
    return default_path if default_path.exists() else None


def _section(payload: Mapping[str, object], name: str) -> Mapping[str, object]:
    value = payload.get(name, {})
    if not isinstance(value, dict):
        raise ConfigError(f"{name} must be a JSON object")
    return value


def _required_json(payload: Mapping[str, object], name: str) -> str:
    value = payload.get(name)
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{name} is required")
    return value.strip()


def _optional_json(payload: Mapping[str, object], name: str) -> str | None:
    value = payload.get(name)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ConfigError(f"{name} must be a string")
    return value.strip() or None
