"""Shared Graph authorization identity and binding primitives."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from typing import Any, Protocol

import requests

from .config import GraphConfig
from .tools.common import OutlookClientError, graph_error_message


GRAPH_HTTP_TIMEOUT_SECONDS = 30.0


class TimeoutSession(requests.Session):
    """requests.Session with a default timeout so MSAL token calls cannot hang forever."""

    def __init__(self, timeout: float = GRAPH_HTTP_TIMEOUT_SECONDS) -> None:
        super().__init__()
        self._default_timeout = timeout

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        kwargs.setdefault("timeout", self._default_timeout)
        return super().request(method, url, **kwargs)


@dataclass(frozen=True)
class GraphAuthorizationBinding:
    """The Microsoft identity bound to one Graph authorization."""

    expected_account: str
    home_account_id: str | None = None
    graph_user_id: str | None = None
    identities: tuple[str, ...] = ()
    status: str = "unbound"


@dataclass(frozen=True)
class GraphIdentity:
    graph_user_id: str | None
    identities: tuple[str, ...]


class GraphIdentityError(OutlookClientError):
    """Raised when a Graph token cannot be bound to the expected mailbox."""


class GraphIdentityMismatchError(GraphIdentityError):
    """Raised when Graph confirms that the token belongs to another mailbox."""

    def __init__(self, message: str, identities: tuple[str, ...] = ()):
        super().__init__(message)
        self.identities = identities


class GraphIdentityTransientError(GraphIdentityError):
    """Raised when Graph cannot temporarily verify the token identity."""


class GraphIdentityVerifier(Protocol):
    def verify(self, token: str, expected_account: str) -> GraphIdentity:
        ...


@dataclass
class RequestsGraphIdentityVerifier:
    config: GraphConfig
    session: requests.Session

    def verify(self, token: str, expected_account: str) -> GraphIdentity:
        try:
            response = self.session.get(
                f"{self.config.graph_base_url}/me",
                params={"$select": "id,mail,userPrincipalName"},
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                timeout=GRAPH_HTTP_TIMEOUT_SECONDS,
            )
        except requests.RequestException as exc:
            raise GraphIdentityTransientError(
                f"Temporarily unable to verify the Microsoft account: {exc}"
            ) from exc
        try:
            try:
                payload = response.json()
            except ValueError:
                payload = {}
        finally:
            response.close()
        if response.status_code == 429 or response.status_code >= 500:
            raise GraphIdentityTransientError(graph_error_message(response.status_code, payload))
        if not 200 <= response.status_code < 300:
            raise GraphIdentityError(graph_error_message(response.status_code, payload))

        identities = graph_account_identities(payload)
        if not any(same_account(identity, expected_account) for identity in identities):
            actual = ", ".join(identities) or "unknown"
            raise GraphIdentityMismatchError(
                f"Authorized Microsoft account {actual} does not match configured Outlook account {expected_account}",
                tuple(identities),
            )
        graph_user_id = str(payload.get("id") or "").strip() or None
        return GraphIdentity(graph_user_id=graph_user_id, identities=tuple(identities))


def graph_config_fingerprint(config: GraphConfig) -> str:
    payload = {
        "authority": config.authority,
        "client_id": config.client_id,
        "scopes": sorted(set(config.scopes)),
    }
    encoded = json.dumps(payload, ensure_ascii=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def account_home_id(account: dict[str, Any]) -> str | None:
    value = account.get("home_account_id") or account.get("homeAccountId")
    normalized = str(value or "").strip()
    return normalized or None


def graph_account_identities(payload: dict[str, Any]) -> list[str]:
    identities: list[str] = []
    for key in ("mail", "userPrincipalName"):
        value = payload.get(key)
        identity = str(value or "").strip()
        if identity and identity not in identities:
            identities.append(identity)
    return identities


def same_account(actual: str | None, expected: str) -> bool:
    return bool(actual) and actual.strip().casefold() == expected.strip().casefold()


def binding_to_json(binding: GraphAuthorizationBinding) -> dict[str, Any]:
    return {
        "expected_account": binding.expected_account,
        "home_account_id": binding.home_account_id,
        "graph_user_id": binding.graph_user_id,
        "identities": list(binding.identities),
        "status": binding.status,
    }


def binding_from_json(value: Any) -> GraphAuthorizationBinding | None:
    if not isinstance(value, dict):
        return None
    expected_account = str(value.get("expected_account") or "").strip()
    if not expected_account:
        return None
    raw_identities = value.get("identities", [])
    if not isinstance(raw_identities, list | tuple):
        raw_identities = []
    identities = tuple(
        str(item).strip()
        for item in raw_identities
        if str(item).strip()
    )
    return GraphAuthorizationBinding(
        expected_account=expected_account,
        home_account_id=str(value.get("home_account_id") or "").strip() or None,
        graph_user_id=str(value.get("graph_user_id") or "").strip() or None,
        identities=identities,
        status=str(value.get("status") or "unbound").strip() or "unbound",
    )
