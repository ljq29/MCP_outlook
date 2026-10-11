from __future__ import annotations

import pytest

from outlook_mcp import client as client_module
from outlook_mcp.client import OutlookClient, OutlookClientError, normalize_folder_name
from outlook_mcp.config import OutlookConfig

from conftest import FakeAccount


def test_get_folder_uses_known_alias() -> None:
    client = OutlookClient(_password_config(), _account=FakeAccount())

    assert client.get_folder("deleted").name == "trash"


def test_get_folder_rejects_unknown_folder() -> None:
    client = OutlookClient(_password_config(), _account=FakeAccount())

    with pytest.raises(OutlookClientError, match="Unknown Outlook folder"):
        client.get_folder("nope")


def test_get_folder_resolves_deleted_items_and_folder_id() -> None:
    account = FakeAccount()
    client = OutlookClient(_password_config(), _account=account)

    assert client.get_folder("Deleted Items") is account.trash
    assert client.get_folder(account.inbox.id) is account.inbox


def test_get_folder_resolves_unique_custom_name() -> None:
    account = FakeAccount()
    custom = type("Folder", (), {"id": "custom-id", "name": "Project Alpha"})()
    account.root.child_folders.append(custom)
    client = OutlookClient(_password_config(), _account=account)

    assert client.get_folder("project alpha") is custom


def test_get_folder_rejects_ambiguous_custom_name() -> None:
    account = FakeAccount()
    account.root.child_folders.extend([
        type("Folder", (), {"id": "custom-1", "name": "Project"})(),
        type("Folder", (), {"id": "custom-2", "name": "project"})(),
    ])
    client = OutlookClient(_password_config(), _account=account)

    with pytest.raises(OutlookClientError, match="Ambiguous Outlook folder name"):
        client.get_folder("Project")


def test_get_item_from_specific_folder() -> None:
    client = OutlookClient(_password_config(), _account=FakeAccount())

    assert client.get_item("mail-1", "inbox").subject == "Quarterly report"


def test_get_item_preserves_transport_errors() -> None:
    class BrokenRoot:
        def get(self, **kwargs):
            raise TimeoutError("Exchange timed out")

    account = FakeAccount()
    account.root = BrokenRoot()
    client = OutlookClient(_password_config(), _account=account)

    with pytest.raises(TimeoutError, match="Exchange timed out"):
        client.get_item("mail-1")


def test_get_item_maps_real_exchangelib_not_found() -> None:
    from exchangelib.errors import DoesNotExist

    class MissingRoot:
        def get(self, **kwargs):
            raise DoesNotExist()

    account = FakeAccount()
    account.root = MissingRoot()
    client = OutlookClient(_password_config(), _account=account)

    with pytest.raises(OutlookClientError, match="Outlook item not found"):
        client.get_item("missing")


def test_normalize_folder_name_defaults_to_inbox() -> None:
    assert normalize_folder_name(None) == "inbox"
    assert normalize_folder_name("Sent Items") == "sent_items"


def test_password_account_builds_exchangelib_account(monkeypatch) -> None:
    calls = {}

    class Credentials:
        def __init__(self, **kwargs):
            calls["credentials"] = kwargs

    class Configuration:
        def __init__(self, **kwargs):
            calls["configuration"] = kwargs

    class Account:
        def __init__(self, **kwargs):
            calls["account"] = kwargs

    monkeypatch.setattr(client_module, "Credentials", Credentials)
    monkeypatch.setattr(client_module, "Configuration", Configuration)
    monkeypatch.setattr(client_module, "Account", Account)

    account = OutlookClient(_password_config()).account

    assert account is not None
    assert calls["credentials"] == {"username": "user", "password": "secret"}
    assert calls["configuration"]["server"] == "outlook.office365.com"
    assert calls["account"]["autodiscover"] is False


def test_oauth_account_builds_identity(monkeypatch) -> None:
    calls = {}

    class Identity:
        def __init__(self, **kwargs):
            calls["identity"] = kwargs

    class OAuth2Credentials:
        def __init__(self, **kwargs):
            calls["oauth"] = kwargs

    class Configuration:
        def __init__(self, **kwargs):
            calls["configuration"] = kwargs

    class Account:
        def __init__(self, **kwargs):
            calls["account"] = kwargs

    monkeypatch.setattr(client_module, "Identity", Identity)
    monkeypatch.setattr(client_module, "OAuth2Credentials", OAuth2Credentials)
    monkeypatch.setattr(client_module, "Configuration", Configuration)
    monkeypatch.setattr(client_module, "Account", Account)

    OutlookClient(
        OutlookConfig(
            email="user@example.com",
            auth_mode="oauth2",
            client_id="client",
            client_secret="secret",
            tenant_id="tenant",
            access_type="impersonation",
        )
    ).account

    assert calls["identity"] == {"primary_smtp_address": "user@example.com"}
    assert calls["oauth"]["tenant_id"] == "tenant"
    assert calls["account"]["access_type"] == client_module.IMPERSONATION


def _password_config() -> OutlookConfig:
    return OutlookConfig(
        email="user@example.com",
        auth_mode="password",
        username="user",
        password="secret",
    )
