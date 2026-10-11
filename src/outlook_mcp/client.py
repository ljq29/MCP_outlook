"""Thin exchangelib adapter used by MCP tools."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from exchangelib import (
    DELEGATE,
    IMPERSONATION,
    Account,
    Configuration,
    Credentials,
    Identity,
    OAuth2Credentials,
)
from exchangelib.errors import DoesNotExist, ErrorItemNotFound

from .config import GraphConfig, OutlookConfig, RuntimeConfig


class OutlookClientError(RuntimeError):
    """Raised when an Exchange account or folder cannot be accessed."""


def create_client(config: RuntimeConfig) -> Any:
    if isinstance(config, GraphConfig):
        from .graph_client import GraphClient

        return GraphClient(config)
    return OutlookClient(config)


@dataclass
class OutlookClient:
    config: OutlookConfig
    _account: Any | None = None

    @property
    def account(self) -> Any:
        if self._account is None:
            self._account = self._build_account()
        return self._account

    def _build_account(self) -> Any:
        try:
            credentials = self._build_credentials()
            exchange_config = Configuration(server=self.config.server, credentials=credentials)
            return Account(
                primary_smtp_address=self.config.email,
                config=exchange_config,
                autodiscover=False,
                access_type=self._access_type(),
            )
        except Exception as exc:  # pragma: no cover - covered through public error handling
            raise OutlookClientError(f"Failed to connect to Outlook mailbox: {_safe_exception_message(exc)}") from exc

    def _build_credentials(self) -> Any:
        if self.config.auth_mode == "oauth2":
            return OAuth2Credentials(
                client_id=self.config.client_id,
                client_secret=self.config.client_secret,
                tenant_id=self.config.tenant_id,
                identity=Identity(primary_smtp_address=self.config.email),
            )
        return Credentials(username=self.config.username, password=self.config.password)

    def _access_type(self) -> str:
        return IMPERSONATION if self.config.access_type == "impersonation" else DELEGATE

    def get_folder(self, name: str) -> Any:
        requested = (name or "inbox").strip()
        folder_name = normalize_folder_name(requested)
        attr = _FOLDER_ATTRS.get(folder_name)
        if attr is not None:
            try:
                return getattr(self.account, attr)
            except AttributeError:
                pass

        folders = list(self.account.root.walk())
        id_matches = [folder for folder in folders if getattr(folder, "id", None) == requested]
        if len(id_matches) == 1:
            return id_matches[0]
        if len(id_matches) > 1:
            raise OutlookClientError(f"Ambiguous Outlook folder ID: {requested}")

        requested_name = requested.casefold()
        name_matches = [
            folder
            for folder in folders
            if str(getattr(folder, "name", "") or "").strip().casefold() == requested_name
        ]
        if len(name_matches) == 1:
            return name_matches[0]
        if len(name_matches) > 1:
            raise OutlookClientError(
                f"Ambiguous Outlook folder name: {requested}; use the folder id returned by mail_folders"
            )
        raise OutlookClientError(f"Unknown Outlook folder: {name}")

    def get_item(self, item_id: str, folder: str | None = None) -> Any:
        if not item_id:
            raise OutlookClientError("item_id is required")

        target = self.get_folder(folder) if folder else getattr(self.account, "root", None)
        if target is None:
            raise OutlookClientError("Account root folder is not available; pass a folder name")
        try:
            return target.get(id=item_id)
        except (DoesNotExist, ErrorItemNotFound) as exc:
            raise OutlookClientError(f"Outlook item not found: {item_id}") from exc


_FOLDER_ATTRS = {
    "inbox": "inbox",
    "deleted": "trash",
    "deleteditems": "trash",
    "deleted_items": "trash",
    "trash": "trash",
    "sent": "sent",
    "sentitems": "sent",
    "sent_items": "sent",
    "drafts": "drafts",
    "outbox": "outbox",
    "junk": "junk",
    "junkemail": "junk",
    "junk_email": "junk",
    "calendar": "calendar",
    "contacts": "contacts",
    "tasks": "tasks",
}


def normalize_folder_name(name: str | None) -> str:
    value = (name or "inbox").strip().lower().replace(" ", "_").replace("-", "_")
    if not value:
        return "inbox"
    return value


def _safe_exception_message(exc: Exception) -> str:
    message = str(exc) or exc.__class__.__name__
    return message.replace("\n", " ")
