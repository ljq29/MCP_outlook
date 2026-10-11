"""Microsoft Graph adapter for personal Outlook.com accounts."""

from __future__ import annotations

import base64
import binascii
import html
import json
import logging
import os
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone as dt_timezone
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.policy import SMTP
from email.utils import formatdate, make_msgid
from pathlib import Path
from typing import Any, Protocol
from uuid import uuid4
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import msal
import requests
from requests.adapters import HTTPAdapter

from .client import OutlookClientError
from .config import GraphConfig
from .graph_auth import (
    GRAPH_HTTP_TIMEOUT_SECONDS,
    GraphAuthorizationBinding,
    GraphIdentity,
    GraphIdentityError,
    GraphIdentityMismatchError,
    GraphIdentityVerifier,
    RequestsGraphIdentityVerifier,
    account_home_id,
    binding_from_json,
    binding_to_json,
    same_account,
    TimeoutSession,
)
from .tools.common import (
    AuthRequiredError,
    MAX_ATTACHMENTS_TOTAL_BYTES,
    OutcomeUnknownError,
    PHONE_LABELS,
    ToolValidationError,
    graph_error_message,
    msal_error_message,
    normalize_limit,
    normalize_offset,
    normalize_reminder_minutes,
    normalize_timezone_name,
    parse_date,
    parse_calendar_datetime,
    parse_datetime,
    parse_percent,
    serialize_datetime_local,
    serialize_datetime_utc,
)
from .uploads import DEFAULT_UPLOAD_MAX_BYTES, attachment_from_upload, ResolvedAttachment


_logger = logging.getLogger(__name__)


class GraphClientError(OutlookClientError):
    """Raised when Microsoft Graph returns an error."""


class GraphOutcomeUnknownError(GraphClientError, OutcomeUnknownError):
    """Raised when a write may have reached Graph but no outcome was received."""


GRAPH_TIMEZONE_ALIASES = {
    "UTC": "UTC",
    "China Standard Time": "Asia/Shanghai",
}
_GRAPH_UTC_HEADERS = {"Prefer": 'outlook.timezone="UTC"'}

_HTML_BLOCK_BREAK_RE = re.compile(r"(?is)<\s*br\s*/?\s*>")
_HTML_CLOSING_BREAK_RE = re.compile(r"(?is)<\s*/\s*(p|div|li|ul|ol|tr|h[1-6])\s*>")
_HTML_LIST_ITEM_OPEN_RE = re.compile(r"(?is)<\s*li\b[^>]*>")
_HTML_SCRIPT_STYLE_RE = re.compile(r"(?is)<\s*(script|style)\b[^>]*>.*?<\s*/\s*\1\s*>")
_HTML_TAG_RE = re.compile(r"(?is)<[^>]+>")
_HTML_DOCUMENT_MARKER_RE = re.compile(r"(?is)<\s*(?:!\s*doctype|/?html|head|body)\b")
_HTML_BODY_RE = re.compile(r"(?is)<\s*body\b[^>]*>(.*?)<\s*/\s*body\s*>")
_HTML_COMMENT_RE = re.compile(r"(?is)<!--.*?-->")
_HTML_DOCUMENT_SECTION_RE = re.compile(r"(?is)<\s*(head|script|style|title)\b[^>]*>.*?<\s*/\s*\1\s*>")
_HTML_DOCUMENT_VOID_TAG_RE = re.compile(r"(?is)<\s*(?:!doctype|meta|link|base)\b[^>]*>")
_HTML_DOCUMENT_WRAPPER_TAG_RE = re.compile(r"(?is)<\s*/?\s*(?:html|body)\b[^>]*>")
_HTML_NON_TEXT_CONTENT_RE = re.compile(r"(?is)<\s*(?:img|table|hr|iframe|video|audio|ul|ol|blockquote|pre|object|embed|svg|canvas)\b")


class GraphTokenStorage(Protocol):
    def load_cache(self) -> str | None:
        ...

    def save_cache(self, cache_json: str) -> None:
        ...

    def load_pending_flow(self) -> dict[str, Any] | None:
        ...

    def save_pending_flow(self, flow: dict[str, Any]) -> None:
        ...

    def delete_pending_flow(self) -> None:
        ...

    def load_binding(self) -> GraphAuthorizationBinding | None:
        ...

    def save_binding(self, binding: GraphAuthorizationBinding | None) -> None:
        ...

    def clear_binding(self) -> None:
        ...


@dataclass
class FileGraphTokenStorage:
    cache_path: Path

    def load_cache(self) -> str | None:
        if not self.cache_path.exists():
            return None
        return self.cache_path.read_text(encoding="utf-8")

    def save_cache(self, cache_json: str) -> None:
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        _chmod_quiet(self.cache_path.parent, 0o700)
        tmp_path = self.cache_path.with_name(f"{self.cache_path.name}.{uuid4().hex}.tmp")
        try:
            tmp_path.write_text(cache_json, encoding="utf-8")
            os.replace(tmp_path, self.cache_path)
        finally:
            tmp_path.unlink(missing_ok=True)
        _chmod_quiet(self.cache_path, 0o600)

    def load_pending_flow(self) -> dict[str, Any] | None:
        path = self._pending_path
        if not path.exists():
            return None
        try:
            flow = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            path.unlink(missing_ok=True)
            return None
        return flow if isinstance(flow, dict) else None

    def save_pending_flow(self, flow: dict[str, Any]) -> None:
        path = self._pending_path
        path.parent.mkdir(parents=True, exist_ok=True)
        _chmod_quiet(path.parent, 0o700)
        path.write_text(json.dumps(flow), encoding="utf-8")
        _chmod_quiet(path, 0o600)

    def delete_pending_flow(self) -> None:
        self._pending_path.unlink(missing_ok=True)

    def load_binding(self) -> GraphAuthorizationBinding | None:
        path = self._binding_path
        if not path.exists():
            return None
        try:
            return binding_from_json(json.loads(path.read_text(encoding="utf-8")))
        except json.JSONDecodeError:
            path.unlink(missing_ok=True)
            return None

    def save_binding(self, binding: GraphAuthorizationBinding | None) -> None:
        if binding is None:
            self._binding_path.unlink(missing_ok=True)
            return
        self._binding_path.parent.mkdir(parents=True, exist_ok=True)
        _chmod_quiet(self._binding_path.parent, 0o700)
        self._binding_path.write_text(json.dumps(binding_to_json(binding)), encoding="utf-8")
        _chmod_quiet(self._binding_path, 0o600)

    def clear_binding(self) -> None:
        self.save_binding(None)

    @property
    def _pending_path(self) -> Path:
        return self.cache_path.with_suffix(self.cache_path.suffix + ".pending.json")

    @property
    def _binding_path(self) -> Path:
        return self.cache_path.with_suffix(self.cache_path.suffix + ".identity.json")


@dataclass
class GraphTokenProvider:
    config: GraphConfig
    storage: GraphTokenStorage | None = None
    identity_verifier: GraphIdentityVerifier | None = None
    _cache: msal.SerializableTokenCache = field(init=False)
    _app: msal.PublicClientApplication = field(init=False)
    _pending_flow: dict[str, Any] | None = None
    _pending_flow_saved: dict[str, Any] | None = field(default=None, init=False)
    _binding: GraphAuthorizationBinding | None = field(default=None, init=False)

    def __post_init__(self) -> None:
        self._token_lock = threading.Lock()
        self._flow_event: threading.Event | None = None
        self._flow_error: BaseException | None = None
        self._cache = msal.SerializableTokenCache()
        if self.storage is None:
            self.storage = FileGraphTokenStorage(self.config.cache_path)
        cache_json = self.storage.load_cache()
        if cache_json:
            try:
                self._cache.deserialize(cache_json)
            except Exception:
                _logger.warning("Graph token cache could not be deserialized; starting with an empty cache")
                _quarantine_corrupt_cache(self.storage)
        self._binding = self.storage.load_binding()
        if self._binding and not same_account(self._binding.expected_account, self.config.email):
            self.storage.clear_binding()
            self._binding = None
        self._pending_flow = self._load_pending_flow()
        self._app = msal.PublicClientApplication(
            self.config.client_id,
            authority=self.config.authority,
            token_cache=self._cache,
            http_client=TimeoutSession(),
        )

    def get_access_token(self) -> str:
        while True:
            with self._token_lock:
                token = self._silent_token()
                if token:
                    return token
                if self._binding and self._binding.home_account_id and not self._candidate_accounts():
                    self._clear_binding_for_reauthorization()
                if self._pending_flow is None:
                    self._initiate_device_flow_locked()
                flow = self._pending_flow
                assert flow is not None
                event = self._flow_event
                if event is None:
                    event = threading.Event()
                    self._flow_event = event
                    self._flow_error = None
                    leader = True
                else:
                    leader = False
            if leader:
                try:
                    return self._poll_device_flow(flow)
                except BaseException as exc:
                    with self._token_lock:
                        self._flow_error = exc
                    raise
                finally:
                    with self._token_lock:
                        self._flow_event = None
                    event.set()
            if not event.wait(timeout=_FLOW_WAIT_TIMEOUT_SECONDS):
                raise self._auth_required(flow)
            with self._token_lock:
                error = self._flow_error
            if error is not None:
                raise error

    def _initiate_device_flow_locked(self) -> None:
        flow = self._app.initiate_device_flow(scopes=list(self.config.scopes))
        if "user_code" not in flow:
            raise GraphClientError(msal_error_message(flow))
        flow["_outlook_mcp_known_home_account_ids"] = [
            home_account_id
            for account in self._app.get_accounts()
            if isinstance(account, dict)
            if (home_account_id := account_home_id(account)) is not None
        ]
        self._pending_flow = flow
        self._save_pending_flow(flow)
        raise self._auth_required(flow)

    def _poll_device_flow(self, flow: dict[str, Any]) -> str:
        deadline = time.time() + 30
        result = None
        poll_interval = _device_flow_interval(flow)
        while time.time() < deadline:
            result = self._app.acquire_token_by_device_flow(
                flow,
                exit_condition=lambda pending: True,
            )
            if result and "access_token" in result:
                known_home_account_ids = _known_device_flow_accounts(flow)
                with self._token_lock:
                    self._pending_flow = None
                    self._delete_pending_flow()
                    self._bind_token_after_device_flow(
                        result["access_token"],
                        known_home_account_ids=known_home_account_ids,
                    )
                    self._save_cache()
                return result["access_token"]
            error = (result or {}).get("error")
            if error == "slow_down":
                poll_interval += 5
            elif error != "authorization_pending":
                break
            time.sleep(poll_interval)
        error = (result or {}).get("error")
        if error in {"authorization_pending", "slow_down"}:
            with self._token_lock:
                self._save_pending_flow(flow)
            raise self._auth_required(flow)
        if error in ("expired_token", "access_denied", "bad_verification_code"):
            with self._token_lock:
                self._pending_flow = None
                self._delete_pending_flow()
                if error == "expired_token":
                    self._initiate_device_flow_locked()
            raise GraphClientError(
                f"Device flow {error.replace('_', ' ')} — clear auth state and re-authorize"
            )
        if error:
            raise GraphClientError(msal_error_message(result))
        raise GraphClientError(
            msal_error_message(result) if result else "Device flow returned an empty response"
        )

    def _silent_token(self) -> str | None:
        for account in self._candidate_accounts():
            result = self._app.acquire_token_silent(list(self.config.scopes), account=account)
            if result and "access_token" in result:
                self._ensure_identity_binding(result["access_token"], account)
                self._save_cache()
                return result["access_token"]
            if result and result.get("error"):
                continue
        return None

    def _candidate_accounts(self) -> list[dict[str, Any]]:
        accounts = [account for account in self._app.get_accounts() if isinstance(account, dict)]
        if self._binding and self._binding.home_account_id:
            return [
                account
                for account in accounts
                if account_home_id(account) == self._binding.home_account_id
            ]
        exact = [
            account
            for account in accounts
            if str(account.get("username") or "").strip().casefold() == self.config.email.strip().casefold()
        ]
        if exact:
            return exact
        filtered = [
            account
            for account in self._app.get_accounts(username=self.config.email)
            if isinstance(account, dict)
        ]
        if filtered:
            return filtered
        if len(accounts) == 1 and self.identity_verifier is not None:
            return accounts
        return []

    def _bind_token_after_device_flow(
        self,
        token: str,
        *,
        known_home_account_ids: frozenset[str],
    ) -> None:
        if self.identity_verifier is None:
            raise GraphIdentityError("Device authorization requires Microsoft account identity verification")
        accounts = [account for account in self._app.get_accounts() if isinstance(account, dict)]
        try:
            identity = self.identity_verifier.verify(token, self.config.email)
        except GraphIdentityMismatchError as exc:
            mismatched_accounts = [
                account
                for account in accounts
                if any(
                    same_account(str(account.get("username") or ""), identity_value)
                    for identity_value in exc.identities
                )
            ]
            if not mismatched_accounts:
                mismatched_accounts = [
                    account
                    for account in accounts
                    if account_home_id(account) not in known_home_account_ids
                ]
            if not mismatched_accounts and len(accounts) == 1:
                mismatched_accounts = accounts
            for account in mismatched_accounts:
                self._app.remove_account(account)
            assert self.storage is not None
            self.storage.clear_binding()
            self._binding = None
            self._save_cache()
            raise AuthRequiredError(
                str(exc),
                {"reason": "authorized_account_mismatch", "account": self.config.email},
            ) from exc
        matching_accounts = [
            account
            for account in accounts
            if any(
                same_account(str(account.get("username") or ""), identity_value)
                for identity_value in identity.identities
            )
        ]
        if len(matching_accounts) == 1:
            accounts = matching_accounts
        else:
            new_accounts = [
                account
                for account in accounts
                if account_home_id(account) not in known_home_account_ids
            ]
            if len(new_accounts) == 1:
                accounts = new_accounts
        if len(accounts) != 1:
            raise GraphIdentityError("Unable to identify the Microsoft account returned by device authorization")
        self._save_identity_binding(accounts[0], identity)

    def _ensure_identity_binding(self, token: str, account: dict[str, Any]) -> None:
        home_account_id = account_home_id(account)
        if not home_account_id:
            if self.identity_verifier is None:
                return
            raise GraphIdentityError("MSAL did not return a stable home_account_id for the authorized account")
        if self._binding and self._binding.home_account_id:
            if not same_account(self._binding.expected_account, self.config.email):
                raise GraphIdentityMismatchError(
                    "The bound Microsoft account does not match the configured Outlook account"
                )
            if self._binding.home_account_id != home_account_id:
                raise GraphIdentityMismatchError("The cached Microsoft account does not match the bound authorization")
            if self._binding.status == "authorized":
                return
        if self.identity_verifier is None:
            return
        try:
            identity = self.identity_verifier.verify(token, self.config.email)
        except GraphIdentityMismatchError as exc:
            self._remove_account(account)
            assert self.storage is not None
            self.storage.clear_binding()
            self._binding = None
            raise AuthRequiredError(
                str(exc),
                {"reason": "authorized_account_mismatch", "account": self.config.email},
            ) from exc
        self._save_identity_binding(account, identity)

    def _save_identity_binding(self, account: dict[str, Any], identity: GraphIdentity) -> None:
        home_account_id = account_home_id(account)
        if not home_account_id:
            raise GraphIdentityError("MSAL did not return a stable home_account_id for the authorized account")
        binding = GraphAuthorizationBinding(
            expected_account=self.config.email,
            home_account_id=home_account_id,
            graph_user_id=identity.graph_user_id,
            identities=identity.identities,
            status="authorized",
        )
        assert self.storage is not None
        self.storage.save_binding(binding)
        self._binding = binding

    def _remove_account(self, account: dict[str, Any]) -> None:
        self._app.remove_account(account)
        self._save_cache()

    def _clear_binding_for_reauthorization(self) -> None:
        assert self.storage is not None
        self.storage.clear_binding()
        self._binding = None

    def _clear_cached_accounts(self) -> None:
        accounts = self._candidate_accounts()
        for account in accounts:
            self._app.remove_account(account)
        assert self.storage is not None
        self.storage.clear_binding()
        self._binding = None
        self._save_cache()

    def _save_cache(self) -> None:
        if not self._cache.has_state_changed:
            return
        assert self.storage is not None
        self.storage.save_cache(self._cache.serialize())

    def _load_pending_flow(self) -> dict[str, Any] | None:
        assert self.storage is not None
        flow = self.storage.load_pending_flow()
        if flow is None:
            return None
        if not isinstance(flow, dict) or flow.get("expires_at", 0) < time.time():
            self.storage.delete_pending_flow()
            return None
        return flow

    def _save_pending_flow(self, flow: dict[str, Any]) -> None:
        if flow is self._pending_flow_saved:
            return
        assert self.storage is not None
        self.storage.save_pending_flow(_pending_flow_with_expiry(flow))
        self._pending_flow_saved = flow

    def _delete_pending_flow(self) -> None:
        assert self.storage is not None
        self.storage.delete_pending_flow()
        self._pending_flow_saved = None

    def _auth_required(self, flow: dict[str, Any]) -> AuthRequiredError:
        data = {
            "verification_uri": flow.get("verification_uri") or flow.get("verification_url"),
            "verification_uri_complete": flow.get("verification_uri_complete"),
            "user_code": flow.get("user_code"),
            "expires_at": _flow_expires_at(flow),
        }
        message = flow.get("message") or "Open the verification URI, enter the user code, then retry the MCP tool."
        return AuthRequiredError(message, data)


_SENTINEL: Any = object()

_FLOW_WAIT_TIMEOUT_SECONDS = 35.0


@dataclass
class GraphClient:
    config: GraphConfig
    session: requests.Session | None = None
    token_provider: GraphTokenProvider | None = None
    _default_task_list_id: Any = _SENTINEL

    def __post_init__(self) -> None:
        if self.session is None:
            self.session = requests.Session()
            adapter = HTTPAdapter(
                pool_connections=5,
                pool_maxsize=20,
                max_retries=0,
                pool_block=True,
            )
            self.session.mount("https://", adapter)
        if self.token_provider is None:
            self.token_provider = GraphTokenProvider(
                self.config,
                identity_verifier=RequestsGraphIdentityVerifier(self.config, self.session),
            )
        elif isinstance(self.token_provider, GraphTokenProvider) and self.token_provider.identity_verifier is None:
            self.token_provider.identity_verifier = RequestsGraphIdentityVerifier(self.config, self.session)

    def close(self) -> None:
        if self.session is not None:
            self.session.close()

    def request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, Any] | None = None,
        json: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        token = self.token_provider.get_access_token()
        return self._request_inner(method, path, token, params, json, headers, can_retry=True)

    _RETRYABLE_STATUSES = frozenset({429, 500, 502, 503, 504})
    _AMBIGUOUS_WRITE_STATUSES = frozenset({500, 502, 503, 504})
    _RETRYABLE_METHODS = frozenset({"GET", "HEAD"})
    _MAX_RETRIES = 4

    def _retry_request(self, perform: Any, *, allow_retry: bool, operation: str) -> Any:
        retries = 0
        delay = 1.0
        while True:
            try:
                response = perform()
            except (requests.ConnectionError, requests.Timeout) as exc:
                if allow_retry and retries < self._MAX_RETRIES:
                    retries += 1
                    time.sleep(delay)
                    delay *= 2
                    continue
                if allow_retry:
                    raise GraphClientError(f"Network error connecting to Microsoft Graph: {exc}") from exc
                raise GraphOutcomeUnknownError(
                    f"Microsoft Graph {operation} did not return an outcome; the write may have completed. "
                    "Do not retry blindly."
                ) from exc

            if (
                allow_retry
                and response.status_code in self._RETRYABLE_STATUSES
                and retries < self._MAX_RETRIES
            ):
                retry_after = response.headers.get("Retry-After")
                response.close()
                retries += 1
                if retry_after and retry_after.isdigit():
                    time.sleep(min(int(retry_after), 60))
                else:
                    time.sleep(delay)
                    delay *= 2
                continue
            return response

    def _parse_or_clear_cache(
        self,
        response: Any,
        *,
        can_retry: bool,
        method: str,
        operation: str,
    ) -> dict[str, Any] | None:
        """Close *response*, parse JSON body. Returns ``None`` when a 401 was
        handled by clearing cached accounts (caller must retry with a fresh token).
        Raises :class:`GraphClientError` on non-2xx statuses."""
        try:
            if response.status_code == 204:
                return {}
            payload = response.json()
        except ValueError:
            payload = {}
        finally:
            response.close()
        if response.status_code == 401 and can_retry:
            self.token_provider._clear_cached_accounts()
            return None
        if not 200 <= response.status_code < 300:
            message = graph_error_message(response.status_code, payload)
            if method not in self._RETRYABLE_METHODS and response.status_code in self._AMBIGUOUS_WRITE_STATUSES:
                raise GraphOutcomeUnknownError(
                    f"{message}. Microsoft Graph {operation} may have completed; do not retry blindly."
                )
            raise GraphClientError(message)
        return payload

    def _request_inner(
        self,
        method: str,
        path: str,
        token: str,
        params: dict[str, Any] | None,
        json: dict[str, Any] | None,
        headers: dict[str, str] | None,
        *,
        can_retry: bool,
    ) -> dict[str, Any]:
        method = method.upper()
        url = f"{self.config.graph_base_url}{path}"
        request_headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        if headers:
            request_headers.update(headers)

        def _perform() -> Any:
            return self.session.request(
                method, url,
                headers=request_headers,
                params=params, json=json, timeout=GRAPH_HTTP_TIMEOUT_SECONDS,
            )

        operation = f"{method} {path}"
        response = self._retry_request(
            _perform,
            allow_retry=method in self._RETRYABLE_METHODS,
            operation=operation,
        )
        payload = self._parse_or_clear_cache(
            response,
            can_retry=can_retry,
            method=method,
            operation=operation,
        )
        if payload is None:
            fresh_token = self.token_provider.get_access_token()
            return self._request_inner(
                method,
                path,
                fresh_token,
                params,
                json,
                headers,
                can_retry=False,
            )
        return payload

    def _list_items(self, endpoint: str, params: dict[str, Any], item_mapper: Any, *,
                    limit: int | None = None, offset: int | None = None) -> dict[str, Any]:
        target: int | None = None
        if limit is not None:
            limit = normalize_limit(limit)
            target = limit + 1
            params["$top"] = min(target, 100)
        if offset is not None:
            params["$skip"] = normalize_offset(offset)
        payload = self.request("GET", endpoint, params=params)
        items: list[dict[str, Any]] = [item_mapper(item) for item in payload.get("value", [])]
        while target is not None and len(items) < target:
            next_link = payload.get("@odata.nextLink")
            if not next_link:
                break
            next_path = self._next_link_path(next_link)
            payload = self.request("GET", next_path, params={})
            items.extend(item_mapper(item) for item in payload.get("value", []))
        return {
            "items": items[:limit] if limit is not None else items,
            "has_more": limit is not None and len(items) > limit,
        }

    def mail_list(self, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
        return self._list_items(
            f"/me/mailFolders/{_mail_folder_id(folder)}/messages",
            {"$orderby": "receivedDateTime desc", "$select": _MAIL_SELECT},
            lambda item: _mail_item(item, include_body=False),
            limit=limit, offset=offset,
        )

    def mail_search(self, query: str, folder: str = "inbox", limit: int = 20, offset: int = 0) -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        return self._list_items(
            f"/me/mailFolders/{_mail_folder_id(folder)}/messages",
            {"$filter": f"contains(subject,'{_odata_string(query.strip())}')", "$select": _MAIL_SELECT},
            lambda item: _mail_item(item, include_body=False),
            limit=limit, offset=offset,
        )

    def mail_get(
        self,
        item_id: str,
        folder: str | None = None,
        include_body: bool = True,
        body_max_chars: int = 4000,
    ) -> dict[str, Any]:
        if body_max_chars < 0:
            raise ToolValidationError("body_max_chars must be greater than or equal to 0")
        payload = self.request("GET", f"/me/messages/{item_id}", params={"$select": _MAIL_GET_SELECT})
        return _mail_item(payload, include_body=include_body, body_max_chars=body_max_chars)

    def mail_send(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        if not to:
            raise ToolValidationError("to must contain at least one recipient")
        if not subject:
            raise ToolValidationError("subject is required")
        message: dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "HTML" if html else "Text", "content": body},
            "toRecipients": _graph_recipients(to),
            "ccRecipients": _graph_recipients(cc or []),
            "bccRecipients": _graph_recipients(bcc or []),
        }
        if attachments:
            message["attachments"] = [
                _graph_attachment(attachment)
                for attachment in _resolve_attachments_with_budget(attachments, html=html)
            ]
        self.request("POST", "/me/sendMail", json={"message": message, "saveToSentItems": True})
        return {"sent": True}

    def mail_reply(
        self,
        item_id: str,
        body: str,
        subject: str | None = None,
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
    ) -> dict[str, Any]:
        message: dict[str, Any] = {}
        if subject:
            message["subject"] = subject
        if to:
            message["toRecipients"] = _graph_recipients(to)
        if cc:
            message["ccRecipients"] = _graph_recipients(cc)
        if bcc:
            message["bccRecipients"] = _graph_recipients(bcc)
        payload: dict[str, Any] = {}
        if html:
            message["body"] = {"contentType": "HTML", "content": body}
        else:
            payload["comment"] = body
        if message:
            payload["message"] = message
        self.request("POST", f"/me/messages/{item_id}/reply", json=payload)
        return {"replied": True, "id": item_id}

    def mail_move(self, item_id: str, target_folder: str) -> dict[str, Any]:
        payload = self.request(
            "POST",
            f"/me/messages/{item_id}/move",
            json={"destinationId": _mail_folder_id(target_folder)},
        )
        return {"moved": True, "id": payload.get("id", item_id), "target_folder": target_folder}

    def mail_delete(self, item_id: str, mode: str = "move_to_trash") -> dict[str, Any]:
        if mode == "move_to_trash":
            self.mail_move(item_id, "deleteditems")
        elif mode == "soft_delete":
            self.request("DELETE", f"/me/messages/{item_id}")
        elif mode == "hard_delete":
            self.request("POST", f"/me/messages/{item_id}/permanentDelete")
        else:
            raise ToolValidationError("mode must be move_to_trash, soft_delete or hard_delete")
        return {"deleted": True, "mode": mode}

    def calendar_list(
        self,
        start: str | None = None,
        end: str | None = None,
        limit: int = 20,
        offset: int = 0,
        timezone: str | None = None,
    ) -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        page_limit = normalize_limit(limit)
        page_offset = normalize_offset(offset)
        now = datetime.now(dt_timezone.utc)
        start_dt = parse_calendar_datetime(start, timezone_name=timezone_name, default=now)
        if end is None:
            end_dt = start_dt + _days(7)
        else:
            end_dt = parse_calendar_datetime(end, timezone_name=timezone_name)
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        params: dict[str, Any] = {
            "startDateTime": serialize_datetime_utc(start_dt),
            "endDateTime": serialize_datetime_utc(end_dt),
            "$top": min(page_limit + 1, 100),
            "$skip": page_offset,
            "$orderby": "start/dateTime",
        }
        payload = self.request(
            "GET",
            "/me/calendarView",
            params=params,
            headers=_GRAPH_UTC_HEADERS,
        )
        raw_items: list[dict[str, Any]] = list(payload.get("value", []))
        while len(raw_items) < page_limit + 1 and payload.get("@odata.nextLink"):
            next_path = self._next_link_path(payload["@odata.nextLink"])
            payload = self.request("GET", next_path, params={}, headers=_GRAPH_UTC_HEADERS)
            raw_items.extend(payload.get("value", []))
        return {
            "items": [
                _event_item(item, timezone_name=timezone_name)
                for item in raw_items[:page_limit]
            ],
            "has_more": len(raw_items) > page_limit,
        }

    def _next_link_path(self, next_link: Any) -> str:
        value = str(next_link or "")
        if value.startswith(self.config.graph_base_url):
            return value[len(self.config.graph_base_url):]
        if value.startswith("/"):
            return value
        raise GraphClientError("Microsoft Graph returned an invalid @odata.nextLink")

    def calendar_create(
        self,
        subject: str,
        start: str,
        end: str,
        body: str = "",
        location: str | None = None,
        required_attendees: list[str] | None = None,
        optional_attendees: list[str] | None = None,
        is_all_day: bool = False,
        send_invites: bool = True,
        html: bool = False,
        attachments: list[dict[str, Any]] | None = None,
        timezone: str | None = None,
        reminder_minutes_before_start: int | None = None,
    ) -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        reminder_minutes = normalize_reminder_minutes(reminder_minutes_before_start)
        if not subject:
            raise ToolValidationError("subject is required")
        if is_all_day:
            start_dt = _parse_all_day_datetime(start, timezone_name)
            end_dt = _parse_all_day_datetime(end, timezone_name)
        else:
            start_dt = parse_calendar_datetime(start, timezone_name=timezone_name)
            end_dt = parse_calendar_datetime(end, timezone_name=timezone_name)
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        required = [_clean_address(address) for address in required_attendees or []]
        optional = [_clean_address(address) for address in optional_attendees or []]
        if required or optional:
            if send_invites is False:
                raise ToolValidationError("Graph calendar_create with attendees does not support send_invites=false")
            return self._calendar_send_invite(
                subject=subject,
                start=start_dt,
                end=end_dt,
                body=body,
                location=location,
                required_attendees=required,
                optional_attendees=optional,
                is_all_day=is_all_day,
                html=html,
                attachments=attachments,
                timezone_name=timezone_name,
                reminder_minutes_before_start=reminder_minutes,
            )
        attachment_payloads = [
            _graph_attachment(attachment)
            for attachment in _resolve_attachments_with_budget(attachments, html=html)
        ]
        payload = self.request(
            "POST",
            "/me/events",
            json=_event_payload(
                subject=subject,
                body=body,
                start=start_dt,
                end=end_dt,
                location=location,
                required_attendees=required_attendees or [],
                optional_attendees=optional_attendees or [],
                is_all_day=is_all_day,
                timezone_name=timezone_name,
                html=html,
                reminder_minutes_before_start=reminder_minutes,
            ),
            headers=_GRAPH_UTC_HEADERS,
        )
        event_id = str(payload.get("id") or "").strip()
        if not event_id:
            raise GraphOutcomeUnknownError(
                "Microsoft Graph created an event response without an id; the remote outcome is unknown"
            )
        attach_errors: list[str] = []
        uploaded_attachment_indexes: list[int] = []
        for i, attachment_payload in enumerate(attachment_payloads):
            try:
                self.request(
                    "POST",
                    f"/me/events/{event_id}/attachments",
                    json=attachment_payload,
                )
            except GraphClientError as exc:
                attach_errors.append(f"attachment #{i}: {exc}")
            else:
                uploaded_attachment_indexes.append(i)
        result = _event_item(payload, timezone_name=timezone_name)
        if attach_errors:
            result["status"] = "partial_success"
            result["partial_success"] = True
            result["event_id"] = event_id
            result["attachment_errors"] = attach_errors
            result["uploaded_attachment_indexes"] = uploaded_attachment_indexes
        return result

    def _calendar_send_invite(
        self,
        *,
        subject: str,
        start: datetime,
        end: datetime,
        body: str,
        location: str | None,
        required_attendees: list[str],
        optional_attendees: list[str],
        is_all_day: bool,
        html: bool,
        attachments: list[dict[str, Any]] | None,
        timezone_name: str,
        reminder_minutes_before_start: int | None,
    ) -> dict[str, Any]:
        uid = f"{uuid4()}@outlook-mcp.local"
        recipients = required_attendees + optional_attendees
        ics = _calendar_invite_ics(
            uid=uid,
            organizer=self.config.email,
            subject=subject,
            body=body,
            start=start,
            end=end,
            location=location,
            required_attendees=required_attendees,
            optional_attendees=optional_attendees,
            is_all_day=is_all_day,
            reminder_minutes_before_start=reminder_minutes_before_start,
        )
        message = _calendar_invite_message(
            sender=self.config.email,
            recipients=recipients,
            subject=subject,
            body=body,
            ics=ics,
            html=html,
            attachments=attachments,
        )
        self._send_mime_message(message.as_bytes(policy=SMTP))
        if is_all_day:
            start_value = start.date().isoformat()
            end_value = end.date().isoformat()
            start_local = start_value
            end_local = end_value
        else:
            start_value = serialize_datetime_utc(start)
            end_value = serialize_datetime_utc(end)
            start_local = serialize_datetime_local(start, timezone_name)
            end_local = serialize_datetime_local(end, timezone_name)
        return {
            "id": None,
            "ics_uid": uid,
            "changekey": None,
            "subject": subject,
            "body": body,
            "body_content_type": "html" if html else "text",
            "body_plain": _html_to_plain_text(body) if html else body,
            "start": start_value,
            "end": end_value,
            "start_local": start_local,
            "end_local": end_local,
            "local_timezone": timezone_name,
            "location": location,
            "is_all_day": is_all_day,
            "is_cancelled": False,
            "is_meeting": True,
            "organizer": self.config.email,
            "required_attendees": required_attendees,
            "optional_attendees": optional_attendees,
            "sent": True,
        }

    def _send_mime_message(self, mime_bytes: bytes) -> None:
        self._send_mime_inner(mime_bytes, can_retry=True)

    def _send_mime_inner(self, mime_bytes: bytes, *, can_retry: bool) -> None:
        token = self.token_provider.get_access_token()
        encoded = base64.b64encode(mime_bytes).decode("ascii")

        def _perform() -> Any:
            return self.session.request(
                "POST", f"{self.config.graph_base_url}/me/sendMail",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Accept": "application/json",
                    "Content-Type": "text/plain",
                },
                data=encoded, timeout=GRAPH_HTTP_TIMEOUT_SECONDS,
            )

        operation = "POST /me/sendMail (MIME)"
        response = self._retry_request(_perform, allow_retry=False, operation=operation)
        payload = self._parse_or_clear_cache(
            response,
            can_retry=can_retry,
            method="POST",
            operation=operation,
        )
        if payload is None:
            return self._send_mime_inner(mime_bytes, can_retry=False)

    def calendar_update(self, item_id: str, updates: dict[str, Any], timezone: str | None = None, send_invites: bool = True, html: bool = False) -> dict[str, Any]:
        if not updates:
            raise ToolValidationError("updates cannot be empty")
        if send_invites is False:
            raise ToolValidationError("Graph calendar_update does not support send_invites=false")
        timezone_name = normalize_timezone_name(timezone)
        existing_attendees, existing_is_all_day = self._calendar_update_context(item_id, updates)
        payload = self.request(
            "PATCH",
            f"/me/events/{item_id}",
            json=_event_update_payload(
                updates,
                timezone_name=timezone_name,
                html=html,
                existing_attendees=existing_attendees,
                existing_is_all_day=existing_is_all_day,
            ),
            headers=_GRAPH_UTC_HEADERS,
        )
        return _event_item(payload, timezone_name=timezone_name)

    def _calendar_update_context(
        self,
        item_id: str,
        updates: dict[str, Any],
    ) -> tuple[list[dict[str, Any]] | None, bool | None]:
        attendee_fields = {"required_attendees", "optional_attendees"} & set(updates)
        needs_attendees = len(attendee_fields) == 1
        needs_all_day = bool({"start", "end"} & set(updates)) and "is_all_day" not in updates
        selected_fields: list[str] = []
        if needs_attendees:
            selected_fields.append("attendees")
        if needs_all_day:
            selected_fields.append("isAllDay")
        if not selected_fields:
            return None, None
        payload = self.request(
            "GET",
            f"/me/events/{item_id}",
            params={"$select": ",".join(selected_fields)},
            headers=_GRAPH_UTC_HEADERS,
        )
        attendees = payload.get("attendees", [])
        existing_attendees = attendees if needs_attendees and isinstance(attendees, list) else None
        existing_is_all_day = bool(payload.get("isAllDay")) if needs_all_day else None
        return existing_attendees, existing_is_all_day

    def calendar_delete(self, item_id: str, send_cancellations: bool = True) -> dict[str, Any]:
        if not send_cancellations:
            raise ToolValidationError("Graph calendar_delete does not support suppressing meeting cancellations")
        if send_cancellations and self._event_needs_cancellation(item_id):
            self.request("POST", f"/me/events/{item_id}/cancel", json={"comment": ""})
        else:
            self.request("DELETE", f"/me/events/{item_id}")
        return {"deleted": True, "id": item_id}

    def _event_needs_cancellation(self, item_id: str) -> bool:
        payload = self.request("GET", f"/me/events/{item_id}", params={"$select": "isOrganizer,attendees"})
        return bool(payload.get("isOrganizer")) and bool(payload.get("attendees"))

    def contacts_list(self, limit: int = 20, offset: int = 0) -> dict[str, Any]:
        return self._list_items(
            "/me/contacts",
            {"$orderby": "displayName"},
            _contact_item,
            limit=limit, offset=offset,
        )

    def contacts_search(self, query: str, limit: int = 20, offset: int = 0) -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        return self._list_items(
            "/me/contacts",
            {"$filter": f"contains(displayName,'{_odata_string(query.strip())}')", "$orderby": "displayName"},
            _contact_item,
            limit=limit, offset=offset,
        )

    def contact_create(
        self,
        display_name: str,
        given_name: str | None = None,
        surname: str | None = None,
        company_name: str | None = None,
        job_title: str | None = None,
        department: str | None = None,
        email_addresses: list[str] | None = None,
        phone_numbers: list[dict[str, str]] | None = None,
        notes: str | None = None,
    ) -> dict[str, Any]:
        if not display_name:
            raise ToolValidationError("display_name is required")
        payload = self.request(
            "POST",
            "/me/contacts",
            json=_contact_payload(
                display_name=display_name,
                given_name=given_name,
                surname=surname,
                company_name=company_name,
                job_title=job_title,
                department=department,
                email_addresses=email_addresses or [],
                phone_numbers=phone_numbers or [],
                notes=notes,
            ),
        )
        return _contact_item(payload)

    def contact_update(self, item_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        if not updates:
            raise ToolValidationError("updates cannot be empty")
        payload = self.request("PATCH", f"/me/contacts/{item_id}", json=_contact_update_payload(updates))
        return _contact_item(payload)

    def contact_delete(self, item_id: str) -> dict[str, Any]:
        self.request("DELETE", f"/me/contacts/{item_id}")
        return {"deleted": True, "id": item_id}

    def tasks_list(self, status: str | None = None, limit: int = 20, offset: int = 0) -> dict[str, Any]:
        params: dict[str, Any] = {}
        if status:
            params["$filter"] = f"status eq '{_graph_task_status(status)}'"
        return self._list_items(
            f"/me/todo/lists/{self.default_task_list_id}/tasks",
            params, _task_item,
            limit=limit, offset=offset,
        )

    def task_create(
        self,
        subject: str,
        body: str | None = None,
        start_date: str | None = None,
        due_date: str | None = None,
        status: str = "NotStarted",
        percent_complete: int | float | str | None = None,
    ) -> dict[str, Any]:
        if not subject:
            raise ToolValidationError("subject is required")
        payload = self.request(
            "POST",
            f"/me/todo/lists/{self.default_task_list_id}/tasks",
            json=_task_payload(
                subject=subject,
                body=body,
                start_date=parse_date(start_date),
                due_date=parse_date(due_date),
                status=status,
                percent_complete=parse_percent(percent_complete),
            ),
        )
        return _task_item(payload)

    def task_update(self, item_id: str, updates: dict[str, Any]) -> dict[str, Any]:
        if not updates:
            raise ToolValidationError("updates cannot be empty")
        payload = self.request(
            "PATCH",
            f"/me/todo/lists/{self.default_task_list_id}/tasks/{item_id}",
            json=_task_update_payload(updates),
        )
        return _task_item(payload)

    def task_delete(self, item_id: str) -> dict[str, Any]:
        self.request("DELETE", f"/me/todo/lists/{self.default_task_list_id}/tasks/{item_id}")
        return {"deleted": True, "id": item_id}

    # -- new operations --

    def mail_forward(self, item_id: str, to: list[str], body: str = "",
                     folder: str | None = None, html: bool = False) -> dict[str, Any]:
        self.request(
            "POST",
            f"/me/messages/{item_id}/forward",
            json={
                "comment": _html_to_plain_text(body) if html else body,
                "toRecipients": _graph_recipients(to),
            },
        )
        return {"forwarded": True}

    def mail_mark_read(self, item_id: str, folder: str | None = None) -> dict[str, Any]:
        self.request("PATCH", f"/me/messages/{item_id}", json={"isRead": True})
        return {"marked": "read", "id": item_id}

    def mail_mark_unread(self, item_id: str, folder: str | None = None) -> dict[str, Any]:
        self.request("PATCH", f"/me/messages/{item_id}", json={"isRead": False})
        return {"marked": "unread", "id": item_id}

    def mail_attachments(self, item_id: str, folder: str | None = None) -> dict[str, Any]:
        payload = self.request("GET", f"/me/messages/{item_id}/attachments", params={"$top": 100})
        raw_items: list[dict[str, Any]] = list(payload.get("value", []))
        while payload.get("@odata.nextLink"):
            payload = self.request("GET", self._next_link_path(payload["@odata.nextLink"]), params={})
            raw_items.extend(payload.get("value", []))
        items = [
            {
                "name": attachment.get("name", ""),
                "size": attachment.get("size", 0),
                "content_type": attachment.get("contentType", ""),
                "attachment_id": attachment.get("id"),
                "is_inline": attachment.get("isInline", False),
            }
            for attachment in raw_items
        ]
        return {"items": items, "has_more": False, "total": len(items)}

    def mail_folders(self) -> dict[str, Any]:
        payload = self.request("GET", "/me/mailFolders", params={"$top": 50})
        raw_items: list[dict[str, Any]] = list(payload.get("value", []))
        while payload.get("@odata.nextLink"):
            payload = self.request("GET", self._next_link_path(payload["@odata.nextLink"]), params={})
            raw_items.extend(payload.get("value", []))
        items = []
        for folder_item in raw_items:
            items.append({
                "id": folder_item.get("id"),
                "name": folder_item.get("displayName", ""),
                "folder_class": "IPF.Note",
                "total_count": folder_item.get("totalItemCount", 0),
                "unread_count": folder_item.get("unreadItemCount", 0),
                "child_folder_count": folder_item.get("childFolderCount", 0),
            })
        return {"items": items, "has_more": False, "total": len(items)}

    def calendar_get(self, item_id: str, timezone: str | None = None) -> dict[str, Any]:
        timezone_name = normalize_timezone_name(timezone)
        payload = self.request("GET", f"/me/events/{item_id}", headers=_GRAPH_UTC_HEADERS)
        return _event_item(payload, timezone_name=timezone_name)

    def calendar_search(self, query: str, start: str | None = None,
                        end: str | None = None, limit: int = 20, offset: int = 0,
                        timezone: str | None = None) -> dict[str, Any]:
        if not query or not query.strip():
            raise ToolValidationError("query is required")
        timezone_name = normalize_timezone_name(timezone)
        limit = normalize_limit(limit)
        offset = normalize_offset(offset)
        now = datetime.now(dt_timezone.utc)
        start_dt = parse_calendar_datetime(start, timezone_name=timezone_name, default=now)
        if end is None:
            end_dt = start_dt + _days(30)
        else:
            end_dt = parse_calendar_datetime(end, timezone_name=timezone_name)
        if end_dt <= start_dt:
            raise ToolValidationError("end must be later than start")
        params: dict[str, Any] = {
            "startDateTime": serialize_datetime_utc(start_dt),
            "endDateTime": serialize_datetime_utc(end_dt),
            "$top": 100,
            "$orderby": "start/dateTime",
        }
        needle = query.strip().lower()
        needed = offset + limit + 1
        matches: list[dict[str, Any]] = []
        endpoint = "/me/calendarView"
        while True:
            payload = self.request("GET", endpoint, params=params, headers=_GRAPH_UTC_HEADERS)
            for item in payload.get("value", []):
                subject = str(item.get("subject") or "")
                if needle in subject.lower():
                    matches.append(_event_item(item, timezone_name=timezone_name))
                    if len(matches) >= needed:
                        break
            if len(matches) >= needed:
                break
            next_link = payload.get("@odata.nextLink")
            if not next_link:
                break
            endpoint = self._next_link_path(next_link)
            params = {}
        page_end = offset + limit
        return {
            "items": matches[offset:page_end],
            "has_more": len(matches) > page_end,
        }

    def contact_get(self, item_id: str) -> dict[str, Any]:
        payload = self.request("GET", f"/me/contacts/{item_id}")
        return _contact_item(payload)

    def task_get(self, item_id: str) -> dict[str, Any]:
        payload = self.request("GET", f"/me/todo/lists/{self.default_task_list_id}/tasks/{item_id}")
        return _task_item(payload)

    def task_search(self, query: str, status: str | None = None,
                    limit: int = 20, offset: int = 0) -> dict[str, Any]:
        params: dict[str, Any] = {}
        filters = []
        if status:
            filters.append(f"status eq '{_graph_task_status(status)}'")
        if query and query.strip():
            filters.append(f"contains(title,'{_odata_string(query.strip())}')")
        if filters:
            params["$filter"] = " and ".join(filters)
        return self._list_items(
            f"/me/todo/lists/{self.default_task_list_id}/tasks",
            params, _task_item,
            limit=limit, offset=offset,
        )

    def account_info(self) -> dict[str, Any]:
        payload = self.request("GET", "/me")
        return {
            "primary_smtp_address": payload.get("mail") or payload.get("userPrincipalName", ""),
            "display_name": payload.get("displayName", ""),
        }

    @property
    def default_task_list_id(self) -> str:
        if self._default_task_list_id is not _SENTINEL:
            if self._default_task_list_id is None:
                raise GraphClientError("No Microsoft To Do task list is available")
            return self._default_task_list_id
        payload = self.request("GET", "/me/todo/lists", params={"$top": 50})
        lists = payload.get("value", [])
        for item in lists:
            if item.get("wellknownListName") == "defaultList":
                self._default_task_list_id = item["id"]
                return self._default_task_list_id
        if lists:
            self._default_task_list_id = lists[0]["id"]
            return self._default_task_list_id
        self._default_task_list_id = None
        raise GraphClientError("No Microsoft To Do task list is available")


_MAIL_SELECT = (
    "id,changeKey,subject,from,sender,receivedDateTime,sentDateTime,toRecipients,ccRecipients,"
    "importance,isRead,hasAttachments,bodyPreview"
)
_MAIL_GET_SELECT = f"{_MAIL_SELECT},body"
_MAIL_FOLDERS = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sentitems": "sentitems",
    "sent_items": "sentitems",
    "deleted": "deleteditems",
    "deleteditems": "deleteditems",
    "deleted_items": "deleteditems",
    "trash": "deleteditems",
    "drafts": "drafts",
    "archive": "archive",
    "junk": "junkemail",
    "junkemail": "junkemail",
}
_TASK_STATUS_TO_GRAPH = {
    "NotStarted": "notStarted",
    "InProgress": "inProgress",
    "Completed": "completed",
    "WaitingOnOthers": "waitingOnOthers",
    "Deferred": "deferred",
}
_TASK_STATUS_FROM_GRAPH = {value: key for key, value in _TASK_STATUS_TO_GRAPH.items()}


def _mail_item(item: dict[str, Any], *, include_body: bool, body_max_chars: int = 4000) -> dict[str, Any]:
    raw_body = item.get("body")
    raw_content = raw_body.get("content") if isinstance(raw_body, dict) else None
    body = raw_content or item.get("bodyPreview", "")
    data = {
        "id": item.get("id"),
        "changekey": item.get("changeKey"),
        "subject": item.get("subject"),
        "author": _mail_address(item.get("from")),
        "sender": _mail_address(item.get("sender")),
        "datetime_received": item.get("receivedDateTime"),
        "datetime_sent": item.get("sentDateTime"),
        "display_to": ", ".join(_recipient_addresses(item.get("toRecipients", []))),
        "display_cc": ", ".join(_recipient_addresses(item.get("ccRecipients", []))),
        "importance": item.get("importance"),
        "is_read": item.get("isRead"),
        "has_attachments": item.get("hasAttachments"),
        "size": None,
    }
    if include_body:
        body_text = str(body or "")
        data["body"] = body_text[:body_max_chars]
        data["body_truncated"] = len(body_text) > body_max_chars
        content_type = raw_body.get("contentType") if raw_content and isinstance(raw_body, dict) else None
        data["body_content_type"] = "html" if str(content_type or "").lower() == "html" else "text"
    return data


def _event_item(item: dict[str, Any], timezone_name: str | None = None) -> dict[str, Any]:
    normalized_timezone = normalize_timezone_name(timezone_name)
    attendees = item.get("attendees", [])
    is_all_day = bool(item.get("isAllDay"))
    if is_all_day:
        start_value = _event_date_value(item.get("start"))
        end_value = _event_date_value(item.get("end"))
        start_local = start_value
        end_local = end_value
    else:
        start = _event_datetime(item.get("start"))
        end = _event_datetime(item.get("end"))
        start_value = serialize_datetime_utc(start)
        end_value = serialize_datetime_utc(end)
        start_local = serialize_datetime_local(start, normalized_timezone)
        end_local = serialize_datetime_local(end, normalized_timezone)
    body, body_content_type = _calendar_body_content(item)
    return {
        "id": item.get("id"),
        "changekey": item.get("changeKey"),
        "subject": item.get("subject"),
        "body": body,
        "body_content_type": body_content_type,
        "body_plain": _body_plain_content(body, body_content_type),
        "start": start_value,
        "end": end_value,
        "start_local": start_local,
        "end_local": end_local,
        "local_timezone": normalized_timezone,
        "location": (item.get("location") or {}).get("displayName"),
        "is_all_day": is_all_day,
        "is_cancelled": item.get("isCancelled"),
        "is_meeting": bool(attendees),
        "organizer": _mail_address(item.get("organizer")),
        "required_attendees": _attendee_addresses(attendees, "required"),
        "optional_attendees": _attendee_addresses(attendees, "optional"),
    }


def _contact_item(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": item.get("id"),
        "changekey": item.get("@odata.etag"),
        "display_name": item.get("displayName"),
        "given_name": item.get("givenName"),
        "surname": item.get("surname"),
        "company_name": item.get("companyName"),
        "job_title": item.get("jobTitle"),
        "department": item.get("department"),
        "email_addresses": [
            {"label": address.get("name") or f"EmailAddress{idx}", "email": address.get("address")}
            for idx, address in enumerate(item.get("emailAddresses", []), start=1)
        ],
        "phone_numbers": _contact_phone_numbers(item),
        "body": item.get("personalNotes"),
    }


def _task_item(item: dict[str, Any]) -> dict[str, Any]:
    status = _TASK_STATUS_FROM_GRAPH.get(item.get("status"), item.get("status"))
    return {
        "id": item.get("id"),
        "changekey": item.get("@odata.etag"),
        "subject": item.get("title"),
        "body": _body_content(item),
        "start_date": _date_only(item.get("startDateTime")),
        "due_date": _date_only(item.get("dueDateTime")),
        "status": status,
        "percent_complete": 100.0 if status == "Completed" else 0.0,
        "is_complete": status == "Completed",
        "owner": None,
    }


def _event_payload(
    *,
    subject: str,
    body: str,
    start: datetime,
    end: datetime,
    location: str | None,
    required_attendees: list[str],
    optional_attendees: list[str],
    is_all_day: bool,
    timezone_name: str | None = None,
    html: bool = False,
    reminder_minutes_before_start: int | None = None,
) -> dict[str, Any]:
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML" if html else "Text", "content": body},
        "start": _graph_datetime(start, timezone_name, all_day=is_all_day),
        "end": _graph_datetime(end, timezone_name, all_day=is_all_day),
        "isAllDay": is_all_day,
        "attendees": _event_attendees(required_attendees, "required") + _event_attendees(optional_attendees, "optional"),
        "isReminderOn": reminder_minutes_before_start is not None,
    }
    if reminder_minutes_before_start is not None:
        payload["reminderMinutesBeforeStart"] = reminder_minutes_before_start
    if location:
        payload["location"] = {"displayName": location}
    return payload


def _event_update_payload(
    updates: dict[str, Any],
    *,
    timezone_name: str | None = None,
    html: bool = False,
    existing_attendees: list[dict[str, Any]] | None = None,
    existing_is_all_day: bool | None = None,
) -> dict[str, Any]:
    all_day = bool(updates.get("is_all_day", existing_is_all_day or False))
    payload: dict[str, Any] = {}
    has_required_update = "required_attendees" in updates
    has_optional_update = "optional_attendees" in updates
    for fld, value in updates.items():
        if fld == "subject":
            payload["subject"] = value
        elif fld == "body":
            payload["body"] = {"contentType": "HTML" if html else "Text", "content": value}
        elif fld == "start":
            parsed_start = (
                _parse_all_day_datetime(value, normalize_timezone_name(timezone_name))
                if all_day
                else parse_calendar_datetime(value, timezone_name=timezone_name)
            )
            payload["start"] = _graph_datetime(parsed_start, timezone_name, all_day=all_day)
        elif fld == "end":
            parsed_end = (
                _parse_all_day_datetime(value, normalize_timezone_name(timezone_name))
                if all_day
                else parse_calendar_datetime(value, timezone_name=timezone_name)
            )
            payload["end"] = _graph_datetime(parsed_end, timezone_name, all_day=all_day)
        elif fld == "location":
            payload["location"] = {"displayName": value}
        elif fld == "is_all_day":
            payload["isAllDay"] = value
        elif fld == "required_attendees":
            continue
        elif fld == "optional_attendees":
            continue
        else:
            raise ToolValidationError(f"Unsupported calendar update field: {fld}")
    if has_required_update or has_optional_update:
        current_required = _attendee_addresses(existing_attendees or [], "required")
        current_optional = _attendee_addresses(existing_attendees or [], "optional")
        required = updates.get("required_attendees", current_required) or []
        optional = updates.get("optional_attendees", current_optional) or []
        payload["attendees"] = _event_attendees(required, "required") + _event_attendees(optional, "optional")
    return payload


def _contact_payload(
    *,
    display_name: str,
    given_name: str | None,
    surname: str | None,
    company_name: str | None,
    job_title: str | None,
    department: str | None,
    email_addresses: list[str],
    phone_numbers: list[dict[str, str]],
    notes: str | None,
) -> dict[str, Any]:
    payload = {
        "displayName": display_name,
        "givenName": given_name,
        "surname": surname,
        "companyName": company_name,
        "jobTitle": job_title,
        "department": department,
        "emailAddresses": [
            {"name": f"EmailAddress{idx}", "address": email.strip()} for idx, email in enumerate(email_addresses[:3], start=1)
        ],
        "personalNotes": notes,
    }
    _apply_phone_numbers(payload, phone_numbers)
    return {key: value for key, value in payload.items() if value is not None}


def _contact_update_payload(updates: dict[str, Any]) -> dict[str, Any]:
    mapping = {
        "display_name": "displayName",
        "given_name": "givenName",
        "surname": "surname",
        "company_name": "companyName",
        "job_title": "jobTitle",
        "department": "department",
        "notes": "personalNotes",
    }
    payload: dict[str, Any] = {}
    for fld, value in updates.items():
        if fld == "email_addresses":
            payload["emailAddresses"] = [
                {"name": f"EmailAddress{idx}", "address": email.strip()} for idx, email in enumerate((value or [])[:3], start=1)
            ]
        elif fld == "phone_numbers":
            _apply_phone_numbers(payload, value or [], include_empty=True)
        elif fld in mapping:
            payload[mapping[fld]] = value
        else:
            raise ToolValidationError(f"Unsupported contact update field: {fld}")
    return payload


def _task_payload(
    *,
    subject: str,
    body: str | None,
    start_date: date | None,
    due_date: date | None,
    status: str,
    percent_complete: Any,
) -> dict[str, Any]:
    graph_status = _task_status_from_percent(percent_complete) if percent_complete is not None else _graph_task_status(status)
    payload: dict[str, Any] = {"title": subject, "status": graph_status}
    if body:
        payload["body"] = {"content": body, "contentType": "text"}
    if start_date:
        payload["startDateTime"] = _graph_date(start_date)
    if due_date:
        payload["dueDateTime"] = _graph_date(due_date)
    return payload


def _task_update_payload(updates: dict[str, Any]) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for fld, value in updates.items():
        if fld == "complete" and value is True:
            payload["status"] = "completed"
        elif fld == "complete" and value is False:
            payload["status"] = "notStarted"
        elif fld == "subject":
            payload["title"] = value
        elif fld == "body":
            payload["body"] = {"content": value, "contentType": "text"}
        elif fld == "start_date":
            parsed = parse_date(value)
            payload["startDateTime"] = _graph_date(parsed) if parsed else None
        elif fld == "due_date":
            parsed = parse_date(value)
            payload["dueDateTime"] = _graph_date(parsed) if parsed else None
        elif fld == "status":
            payload["status"] = _graph_task_status(value)
        elif fld == "percent_complete":
            payload["status"] = _task_status_from_percent(parse_percent(value))
        else:
            raise ToolValidationError(f"Unsupported task update field: {fld}")
    return payload


def _graph_datetime(value: datetime, timezone_name: str | None = None, *, all_day: bool = False) -> dict[str, str]:
    if timezone_name:
        local = value.astimezone(ZoneInfo(timezone_name))
        if all_day:
            local = local.replace(hour=0, minute=0, second=0, microsecond=0)
        return {"dateTime": local.replace(tzinfo=None).isoformat(), "timeZone": timezone_name}
    utc_value = value.astimezone(dt_timezone.utc).replace(tzinfo=None)
    return {"dateTime": utc_value.isoformat(), "timeZone": "UTC"}


def _parse_all_day_datetime(value: str | datetime, timezone_name: str) -> datetime:
    zone = ZoneInfo(timezone_name)
    if isinstance(value, datetime):
        parsed = parse_calendar_datetime(value, timezone_name=timezone_name)
        raw_date = parsed.astimezone(zone).date()
    else:
        raw_value = str(value).strip()
        try:
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw_value):
                raw_date = date.fromisoformat(raw_value)
            else:
                parsed = parse_calendar_datetime(raw_value, timezone_name=timezone_name)
                raw_date = parsed.astimezone(zone).date()
        except (TypeError, ValueError) as exc:
            raise ToolValidationError(f"Invalid all-day date: {value}") from exc
    return datetime.combine(raw_date, datetime.min.time(), tzinfo=zone)


def _graph_date(value: date) -> dict[str, str]:
    return {"dateTime": value.isoformat(), "timeZone": "UTC"}


def _date_time_value(value: Any) -> str | None:
    if not isinstance(value, dict):
        return None
    return value.get("dateTime")


def _event_date_value(value: Any) -> str | None:
    raw = _date_time_value(value)
    if not raw:
        return None
    raw_date = str(raw).strip()[:10]
    try:
        return date.fromisoformat(raw_date).isoformat()
    except ValueError as exc:
        raise GraphClientError(f"Microsoft Graph returned an invalid all-day event date: {raw}") from exc


def _event_datetime(value: Any) -> datetime | None:
    raw = _date_time_value(value)
    if not raw:
        return None
    raw_text = str(raw).strip()
    parsed = parse_datetime(raw_text)
    if parsed.tzinfo is not None and raw_text.endswith(("Z", "z")):
        return parsed
    if parsed.tzinfo is not None and ("+" in raw_text[10:] or "-" in raw_text[10:]):
        return parsed
    time_zone = value.get("timeZone") if isinstance(value, dict) else None
    if isinstance(time_zone, str) and time_zone.strip():
        zone_name = GRAPH_TIMEZONE_ALIASES.get(time_zone.strip(), time_zone.strip())
        try:
            return parsed.replace(tzinfo=ZoneInfo(zone_name))
        except ZoneInfoNotFoundError as exc:
            raise GraphClientError(
                f"Microsoft Graph returned unsupported timezone {time_zone!r}; the event time was not guessed"
            ) from exc
    return parsed


def _date_only(value: Any) -> str | None:
    raw = _date_time_value(value)
    return raw[:10] if raw else None


def _body_content(item: dict[str, Any]) -> str | None:
    body = item.get("body")
    if isinstance(body, dict):
        return body.get("content")
    return None


def _body_content_type(item: dict[str, Any]) -> str | None:
    body = item.get("body")
    if not isinstance(body, dict):
        return None
    content_type = str(body.get("contentType") or "").strip().lower()
    return content_type or None


def _calendar_body_content(item: dict[str, Any]) -> tuple[str | None, str | None]:
    body = _body_content(item)
    content_type = _body_content_type(item)
    if body is None:
        return None, content_type
    fragment = _html_document_body_fragment(body)
    if fragment is None:
        return body, content_type
    return fragment, "html"


def _html_document_body_fragment(value: str) -> str | None:
    normalized = value.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not normalized:
        return None
    if "<" not in normalized and "&lt;" in normalized.lower():
        unescaped = html.unescape(normalized)
        if _HTML_DOCUMENT_MARKER_RE.search(unescaped):
            normalized = unescaped
    if not _HTML_DOCUMENT_MARKER_RE.search(normalized):
        return None

    body_match = _HTML_BODY_RE.search(normalized)
    fragment = body_match.group(1) if body_match else normalized
    fragment = _HTML_COMMENT_RE.sub("", fragment)
    fragment = _HTML_DOCUMENT_SECTION_RE.sub("", fragment)
    fragment = _HTML_DOCUMENT_VOID_TAG_RE.sub("", fragment)
    fragment = _HTML_DOCUMENT_WRAPPER_TAG_RE.sub("", fragment)
    fragment = fragment.strip()
    if not _html_to_plain_text(fragment) and not _HTML_NON_TEXT_CONTENT_RE.search(fragment):
        return ""
    return fragment


def _body_plain_content(body: str | None, content_type: str | None) -> str | None:
    if body is None:
        return None
    if content_type == "html":
        return _html_to_plain_text(body)
    return body


def _mail_address(value: Any) -> str | None:
    if not isinstance(value, dict):
        return None
    address = value.get("emailAddress")
    if isinstance(address, dict):
        return address.get("address")
    return None


def _recipient_addresses(values: list[dict[str, Any]]) -> list[str]:
    return [address for value in values if (address := _mail_address(value))]


def _attendee_addresses(values: list[dict[str, Any]], attendee_type: str) -> list[str]:
    return [
        address
        for value in values
        if value.get("type") == attendee_type and (address := _mail_address(value))
    ]


def _event_attendees(addresses: list[str], attendee_type: str) -> list[dict[str, Any]]:
    return [
        {
            "emailAddress": {"address": cleaned, "name": cleaned},
            "type": attendee_type,
        }
        for address in addresses
        if (cleaned := _clean_address(address))
    ]


def _graph_recipients(addresses: list[str]) -> list[dict[str, Any]]:
    return [{"emailAddress": {"address": _clean_address(address)}} for address in addresses]


def _calendar_invite_message(
    *,
    sender: str,
    recipients: list[str],
    subject: str,
    body: str,
    ics: str,
    html: bool = False,
    attachments: list[dict[str, Any]] | None = None,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    message["Date"] = formatdate(localtime=False)
    message["Message-ID"] = make_msgid(domain="outlook-mcp.local")
    message["Content-Class"] = "urn:content-classes:calendarmessage"
    if html:
        message.set_content(_html_to_plain_text(body), charset="utf-8")
        message.add_alternative(body, subtype="html", charset="utf-8")
    else:
        message.set_content(body, charset="utf-8")
    message.add_alternative(
        ics,
        subtype="calendar",
        charset="utf-8",
        params={"method": "REQUEST", "name": "invite.ics"},
    )
    if attachments:
        message.make_mixed()
    for resolved_attachment in _resolve_attachments_with_budget(attachments, html=html):
        _add_mime_attachment(message, resolved_attachment)
    return message


def _calendar_invite_ics(
    *,
    uid: str,
    organizer: str,
    subject: str,
    body: str,
    start: datetime,
    end: datetime,
    location: str | None,
    required_attendees: list[str],
    optional_attendees: list[str],
    is_all_day: bool,
    reminder_minutes_before_start: int | None,
) -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//OutlookMCP//MIME ICS Invite//EN",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:REQUEST",
        "BEGIN:VEVENT",
        f"UID:{_ics_value(uid)}",
        f"DTSTAMP:{datetime.now(dt_timezone.utc).strftime('%Y%m%dT%H%M%SZ')}",
    ]
    if is_all_day:
        lines.append(f"DTSTART;VALUE=DATE:{start.date().strftime('%Y%m%d')}")
        lines.append(f"DTEND;VALUE=DATE:{end.date().strftime('%Y%m%d')}")
    else:
        lines.append(f"DTSTART:{start.astimezone(dt_timezone.utc).strftime('%Y%m%dT%H%M%SZ')}")
        lines.append(f"DTEND:{end.astimezone(dt_timezone.utc).strftime('%Y%m%dT%H%M%SZ')}")
    lines.extend(
        [
            f"SUMMARY:{_ics_value(subject)}",
            f"DESCRIPTION:{_ics_value(body)}",
            "STATUS:CONFIRMED",
            "TRANSP:OPAQUE",
            f"ORGANIZER;CN={_ics_param_value(organizer)}:mailto:{organizer}",
        ]
    )
    if location:
        lines.append(f"LOCATION:{_ics_value(location)}")
    for address in required_attendees:
        lines.append(
            "ATTENDEE;"
            f"CN={_ics_param_value(address)};"
            "ROLE=REQ-PARTICIPANT;"
            "PARTSTAT=NEEDS-ACTION;"
            "RSVP=TRUE:"
            f"mailto:{address}"
        )
    for address in optional_attendees:
        lines.append(
            "ATTENDEE;"
            f"CN={_ics_param_value(address)};"
            "ROLE=OPT-PARTICIPANT;"
            "PARTSTAT=NEEDS-ACTION;"
            "RSVP=TRUE:"
            f"mailto:{address}"
        )
    if reminder_minutes_before_start is not None:
        trigger = "PT0M" if reminder_minutes_before_start == 0 else f"-PT{reminder_minutes_before_start}M"
        lines.extend(
            [
                "BEGIN:VALARM",
                f"TRIGGER:{trigger}",
                "ACTION:DISPLAY",
                f"DESCRIPTION:{_ics_value(subject)}",
                "END:VALARM",
            ]
        )
    lines.extend(["SEQUENCE:0", "END:VEVENT", "END:VCALENDAR"])
    return "\r\n".join(_ics_fold_line(line) for line in lines) + "\r\n"


_ICS_LINE_OCTET_LIMIT = 75


def _ics_fold_line(line: str) -> str:
    if len(line.encode("utf-8")) <= _ICS_LINE_OCTET_LIMIT:
        return line
    folded: list[str] = []
    current = ""
    current_octets = 0
    limit = _ICS_LINE_OCTET_LIMIT
    for char in line:
        char_octets = len(char.encode("utf-8"))
        if current_octets + char_octets > limit:
            folded.append(current)
            current = " " + char
            current_octets = 1 + char_octets
            limit = _ICS_LINE_OCTET_LIMIT - 1
        else:
            current += char
            current_octets += char_octets
    folded.append(current)
    return "\r\n".join(folded)


def _chmod_quiet(path: Path, mode: int) -> None:
    try:
        os.chmod(path, mode)
    except OSError:
        pass


def _quarantine_corrupt_cache(storage: GraphTokenStorage) -> None:
    if not isinstance(storage, FileGraphTokenStorage):
        return
    path = storage.cache_path
    if not path.exists():
        return
    backup = path.with_name(f"{path.name}.corrupt-{int(time.time())}")
    try:
        os.replace(path, backup)
        _chmod_quiet(backup, 0o600)
        _logger.warning("Corrupt Graph token cache backed up to %s", backup)
    except OSError:
        _logger.warning("Corrupt Graph token cache at %s could not be backed up", path)


def _clean_address(address: str) -> str:
    from .tools.common import validate_email
    return validate_email(address, label="recipient address")


def _ics_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n")


def _ics_param_value(value: str) -> str:
    cleaned = _ics_value(value).replace('"', "'")
    if any(char in cleaned for char in (",", ";", ":")):
        return f'"{cleaned}"'
    return cleaned


def _html_to_plain_text(value: str) -> str:
    normalized = value.replace("\r\n", "\n").replace("\r", "\n")
    if (
        "<" not in normalized
        and "&lt;" in normalized.lower()
        and re.search(r"&lt;/?(?:p|div|span|br|ul|ol|li|table|thead|tbody|tr|td|th|strong|em|b|i|u)\b", normalized, re.IGNORECASE)
    ):
        normalized = html.unescape(normalized)
    normalized = _HTML_SCRIPT_STYLE_RE.sub("", normalized)
    normalized = _HTML_BLOCK_BREAK_RE.sub("\n", normalized)
    normalized = _HTML_CLOSING_BREAK_RE.sub("\n", normalized)
    normalized = _HTML_LIST_ITEM_OPEN_RE.sub("- ", normalized)
    normalized = _HTML_TAG_RE.sub("", normalized)
    normalized = html.unescape(normalized).replace("\xa0", " ")
    normalized = re.sub(r"[ \t]+\n", "\n", normalized)
    normalized = re.sub(r"\n[ \t]+", "\n", normalized)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.strip()


def _resolved_attachment(raw: dict[str, Any], *, html: bool = False) -> ResolvedAttachment:
    if raw.get("relative_path"):
        return attachment_from_upload(raw, html=html)

    name = raw.get("name", "").strip()
    content_base64 = raw.get("content_base64", "")
    if not name or not content_base64:
        raise ToolValidationError("attachments require name and content_base64")
    try:
        content = base64.b64decode(content_base64, validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ToolValidationError(f"Invalid base64 attachment content for {name}") from exc
    if len(content) > DEFAULT_UPLOAD_MAX_BYTES:
        raise ToolValidationError(
            f"attachment {name} is {len(content)} bytes, exceeding the "
            f"{DEFAULT_UPLOAD_MAX_BYTES} bytes per-attachment limit"
        )
    return ResolvedAttachment(
        name=name,
        content=content,
        content_base64=content_base64,
        content_type=str(raw.get("content_type") or ""),
        disposition="attachment",
        content_id=None,
    )


def _resolve_attachments_with_budget(raws: list[dict[str, Any]] | None, *, html: bool = False) -> list[ResolvedAttachment]:
    resolved: list[ResolvedAttachment] = []
    total_bytes = 0
    for raw in raws or []:
        attachment = _resolved_attachment(raw, html=html)
        total_bytes += len(attachment.content)
        if total_bytes > MAX_ATTACHMENTS_TOTAL_BYTES:
            raise ToolValidationError(
                f"attachments total {total_bytes} bytes exceeds the "
                f"{MAX_ATTACHMENTS_TOTAL_BYTES} bytes per-message limit"
            )
        resolved.append(attachment)
    return resolved


def _add_mime_attachment(message: EmailMessage, attachment: ResolvedAttachment) -> None:
    content_type = attachment.content_type or "application/octet-stream"
    maintype, _, subtype = content_type.partition("/")
    subtype = subtype or "octet-stream"
    if maintype == "image":
        part = MIMEImage(attachment.content, _subtype=subtype)
    else:
        part = MIMEApplication(attachment.content, _subtype=subtype)
    part.add_header("Content-Disposition", "inline" if attachment.disposition == "inline" else "attachment", filename=attachment.name)
    if attachment.content_id:
        part.add_header("Content-ID", f"<{attachment.content_id}>")
    part.add_header("Content-Type", content_type, name=attachment.name)
    message.attach(part)


def _graph_attachment(attachment: ResolvedAttachment) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": attachment.name,
        "contentBytes": attachment.content_base64,
    }
    if attachment.content_type:
        payload["contentType"] = attachment.content_type
    if attachment.disposition == "inline":
        payload["isInline"] = True
        payload["contentId"] = attachment.content_id
    return payload


def _contact_phone_numbers(item: dict[str, Any]) -> list[dict[str, str]]:
    result = [{"label": "BusinessPhone", "phone_number": number} for number in item.get("businessPhones", [])]
    if item.get("mobilePhone"):
        result.append({"label": "MobilePhone", "phone_number": item["mobilePhone"]})
    for number in item.get("homePhones", []):
        result.append({"label": "HomePhone", "phone_number": number})
    return result


def _apply_phone_numbers(
    payload: dict[str, Any],
    numbers: list[dict[str, str]],
    *,
    include_empty: bool = False,
) -> None:
    business: list[str] = []
    home: list[str] = []
    mobile: str | None = None
    for raw in numbers:
        label = raw.get("label", "BusinessPhone")
        number = raw.get("phone_number", "").strip()
        if not number:
            raise ToolValidationError("phone_numbers entries require phone_number")
        if label not in PHONE_LABELS:
            raise ToolValidationError(f"Unsupported phone number label: {label}")
        if label == "MobilePhone":
            mobile = number
        elif label in {"HomePhone", "HomePhone2"}:
            home.append(number)
        else:
            business.append(number)
    if business or include_empty:
        payload["businessPhones"] = business
    if home or include_empty:
        payload["homePhones"] = home
    if mobile or include_empty:
        payload["mobilePhone"] = mobile


def _mail_folder_id(name: str | None) -> str:
    value = (name or "inbox").strip()
    key = value.lower().replace(" ", "_").replace("-", "_")
    return _MAIL_FOLDERS.get(key, value)


def _graph_task_status(status: str) -> str:
    if status in _TASK_STATUS_TO_GRAPH:
        return _TASK_STATUS_TO_GRAPH[status]
    if status in _TASK_STATUS_FROM_GRAPH:
        return status
    raise ToolValidationError(f"Unsupported task status: {status}")


def _task_status_from_percent(percent: Any) -> str:
    if percent is None:
        return "notStarted"
    if percent >= 100:
        return "completed"
    if percent <= 0:
        return "notStarted"
    return "inProgress"


_MAX_QUERY_LEN = 256


def _odata_string(value: str) -> str:
    """Escape single-quotes and backslashes, strip double-quotes, truncate to _MAX_QUERY_LEN."""
    return value.replace('"', "").replace("\\", "\\\\").replace("'", "''")[:_MAX_QUERY_LEN]


def _days(count: int) -> timedelta:
    return timedelta(days=count)


def _device_flow_interval(flow: dict[str, Any]) -> float:
    raw_interval = flow.get("interval", 5)
    if isinstance(raw_interval, bool) or not isinstance(raw_interval, int | float):
        return 5.0
    return min(max(float(raw_interval), 1.0), 60.0)


def _known_device_flow_accounts(flow: dict[str, Any]) -> frozenset[str]:
    raw_accounts = flow.get("_outlook_mcp_known_home_account_ids", [])
    if not isinstance(raw_accounts, list):
        return frozenset()
    return frozenset(str(value) for value in raw_accounts if str(value).strip())


def _flow_expires_at(flow: dict[str, Any]) -> str | None:
    expires_at = flow.get("expires_at")
    if isinstance(expires_at, int | float):
        return datetime.fromtimestamp(expires_at, tz=dt_timezone.utc).isoformat()
    expires_in = flow.get("expires_in")
    if isinstance(expires_in, int | float):
        return datetime.fromtimestamp(time.time() + expires_in, tz=dt_timezone.utc).isoformat()
    return None


def _pending_flow_with_expiry(flow: dict[str, Any]) -> dict[str, Any]:
    if isinstance(flow.get("expires_at"), int | float):
        return flow
    expires_in = flow.get("expires_in")
    if not isinstance(expires_in, int | float):
        return flow
    return {**flow, "expires_at": time.time() + expires_in}
