from __future__ import annotations

import base64
import json
import os
from email import policy
from email.parser import BytesParser
from pathlib import Path
from typing import Any

import pytest

from outlook_mcp.config import GraphConfig
from outlook_mcp.credential_store import CredentialStore, SQLiteGraphTokenStorage
from outlook_mcp.graph_auth import (
    GraphAuthorizationBinding,
    GraphIdentity,
    GraphIdentityMismatchError,
    GraphIdentityTransientError,
    RequestsGraphIdentityVerifier,
    graph_config_fingerprint,
)
from outlook_mcp.tools.common import graph_error_message
from outlook_mcp.graph_client import (
    FileGraphTokenStorage,
    GraphClient,
    GraphClientError,
    GraphOutcomeUnknownError,
    GraphTokenProvider,
    _MAX_QUERY_LEN,
    _html_to_plain_text,
    _ics_param_value,
    _ics_value,
    _odata_string,
)
from outlook_mcp.tools.common import AuthRequiredError, ToolValidationError
from conftest import PNG_BYTES
from outlook_mcp.uploads import (
    reset_upload_credential_id,
    save_uploaded_file,
    set_upload_credential_id,
)


class StaticTokenProvider:
    def get_access_token(self) -> str:
        return "access-token"


class FakeResponse:
    def __init__(self, status_code: int, payload: dict[str, Any] | None = None, headers: dict[str, str] | None = None):
        self.status_code = status_code
        self.payload = payload or {}
        self.headers = headers or {}

    def json(self) -> dict[str, Any]:
        return self.payload

    def close(self) -> None:
        pass


class FakeSession:
    def __init__(self, responses: list[FakeResponse]):
        self.responses = responses
        self.calls: list[dict[str, Any]] = []

    def request(self, method: str, url: str, **kwargs: Any) -> FakeResponse:
        self.calls.append({"method": method, "url": url, **kwargs})
        return self.responses.pop(0)


def _unfold_ics(text: str) -> str:
    return text.replace("\r\n ", "").replace("\r\n\t", "").replace("\n ", "").replace("\n\t", "")


class _RetryingTokenProvider:
    def __init__(self, tokens: list[str]):
        self._tokens = iter(tokens)
        self.cleared = 0

    def get_access_token(self) -> str:
        return next(self._tokens)

    def _clear_cached_accounts(self) -> None:
        self.cleared += 1


def test_graph_auth_required_is_returned(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = False

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "message": "Use ABCD-EFGH",
            }

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    provider = GraphTokenProvider(_graph_config(tmp_path))

    with pytest.raises(AuthRequiredError) as exc:
        provider.get_access_token()

    assert exc.value.data["user_code"] == "ABCD-EFGH"
    assert exc.value.data["verification_uri"] == "https://microsoft.com/devicelogin"
    pending_path = Path(_graph_config(tmp_path).token_cache_path + ".pending.json")
    assert pending_path.exists()
    assert json.loads(pending_path.read_text(encoding="utf-8"))["expires_at"] > 0


def test_graph_pending_flow_success_saves_cache(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return '{"token": true}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"home_account_id": "home-1", "username": "person@outlook.com"}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> None:
            return None

        def acquire_token_by_device_flow(self, flow: dict[str, Any], exit_condition: Any) -> dict[str, Any]:
            return {"access_token": "new-token"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    Path(config.token_cache_path + ".pending.json").write_text('{"user_code": "ABCD", "expires_at": 9999999999}', encoding="utf-8")
    storage = FileGraphTokenStorage(config.cache_path)
    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account=config.email,
            home_account_id="home-1",
            status="authorized",
        )
    )

    class FakeVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            assert token == "new-token"
            return GraphIdentity(graph_user_id="graph-user-1", identities=(expected_account,))

    provider = GraphTokenProvider(config, storage=storage, identity_verifier=FakeVerifier())

    assert provider.get_access_token() == "new-token"
    assert Path(config.token_cache_path).read_text(encoding="utf-8") == '{"token": true}'
    assert not Path(config.token_cache_path + ".pending.json").exists()


def test_graph_pending_flow_honors_interval_and_slow_down(monkeypatch, tmp_path) -> None:
    class FakeTime:
        now = 0.0
        sleeps: list[float] = []

        @classmethod
        def time(cls) -> float:
            return cls.now

        @classmethod
        def sleep(cls, seconds: float) -> None:
            cls.sleeps.append(seconds)
            cls.now += seconds

    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return '{"token": true}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.results = [
                {"error": "slow_down"},
                {"error": "authorization_pending"},
                {"access_token": "new-token"},
            ]

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"home_account_id": "home-1", "username": "person@outlook.com"}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> None:
            return None

        def acquire_token_by_device_flow(self, flow: dict[str, Any], exit_condition: Any) -> dict[str, Any]:
            return self.results.pop(0)

    class FakeVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            return GraphIdentity(graph_user_id="graph-user-1", identities=(expected_account,))

    monkeypatch.setattr("outlook_mcp.graph_client.time", FakeTime)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    Path(config.token_cache_path + ".pending.json").write_text(
        '{"user_code": "ABCD", "expires_at": 999, "interval": 2}',
        encoding="utf-8",
    )

    provider = GraphTokenProvider(config, identity_verifier=FakeVerifier())

    assert provider.get_access_token() == "new-token"
    assert FakeTime.sleeps == [7.0, 7.0]


def test_graph_device_flow_mismatch_removes_wrong_account_and_clears_pending(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return '{"accounts": []}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.accounts = [{"home_account_id": "wrong-home", "username": "wrong@example.com"}]

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return list(self.accounts)

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> None:
            return None

        def acquire_token_by_device_flow(self, flow: dict[str, Any], exit_condition: Any) -> dict[str, str]:
            return {"access_token": "wrong-token"}

        def remove_account(self, account: dict[str, Any]) -> None:
            self.accounts.remove(account)

    class MismatchVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            raise GraphIdentityMismatchError(
                "Authorized Microsoft account wrong@example.com does not match configured account",
                ("wrong@example.com",),
            )

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    pending_path = Path(config.token_cache_path + ".pending.json")
    pending_path.write_text(
        '{"user_code": "ABCD", "expires_at": 9999999999, "_outlook_mcp_known_home_account_ids": []}',
        encoding="utf-8",
    )
    provider = GraphTokenProvider(config, identity_verifier=MismatchVerifier())
    saved_account_snapshots: list[list[str]] = []
    original_save_cache = provider._save_cache

    def save_cache_after_identity_check() -> None:
        saved_account_snapshots.append(
            [str(account["username"]) for account in provider._app.accounts]
        )
        original_save_cache()

    provider._save_cache = save_cache_after_identity_check

    with pytest.raises(AuthRequiredError, match="wrong@example.com"):
        provider.get_access_token()

    assert provider._app.accounts == []
    assert saved_account_snapshots == [[]]
    assert provider.storage.load_binding() is None
    assert provider.storage.load_cache() == '{"accounts": []}'
    assert not pending_path.exists()


def test_graph_silent_token_loads_and_saves_cache(monkeypatch, tmp_path) -> None:
    cache_path = tmp_path / "cache.json"
    cache_path.write_text('{"old": true}', encoding="utf-8")

    class FakeCache:
        has_state_changed = True

        def deserialize(self, value: str) -> None:
            self.loaded = value

        def serialize(self) -> str:
            return '{"new": true}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"username": username}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "silent-token"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(cache_path))

    assert GraphTokenProvider(config).get_access_token() == "silent-token"
    assert cache_path.read_text(encoding="utf-8") == '{"new": true}'


def test_graph_silent_token_does_not_fall_back_to_unrelated_account(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = False

        def deserialize(self, value: str) -> None:
            self.loaded = value

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.account_queries: list[str | None] = []

        def get_accounts(self, username: str | None = None) -> list[Any]:
            self.account_queries.append(username)
            if username == "person@outlook.com":
                return []
            return [{"username": "other@outlook.com"}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            raise AssertionError("silent auth must not use unrelated cached accounts")

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
            }

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)

    with pytest.raises(AuthRequiredError):
        GraphTokenProvider(config).get_access_token()


def test_graph_silent_token_binds_home_account_when_username_differs(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return '{"AccessToken": {}, "Account": {}}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            if username is not None:
                return []
            return [{"home_account_id": "home-1", "username": "login@live.com"}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            assert account["home_account_id"] == "home-1"
            return {"access_token": "silent-token"}

    class FakeVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            assert token == "silent-token"
            assert expected_account == "person@outlook.com"
            return GraphIdentity(graph_user_id="graph-user-1", identities=(expected_account,))

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)

    provider = GraphTokenProvider(config, identity_verifier=FakeVerifier())

    assert provider.get_access_token() == "silent-token"
    binding = provider.storage.load_binding()
    assert binding is not None
    assert binding.home_account_id == "home-1"
    assert binding.graph_user_id == "graph-user-1"
    assert binding.status == "authorized"


def test_graph_silent_token_uses_bound_home_account_without_username_filter(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = False

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            assert username is None
            return [
                {"home_account_id": "other", "username": "other@example.com"},
                {"home_account_id": "home-1", "username": "login@live.com"},
            ]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            assert account["home_account_id"] == "home-1"
            return {"access_token": "bound-token"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    storage = FileGraphTokenStorage(config.cache_path)
    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account=config.email,
            home_account_id="home-1",
            graph_user_id="graph-user-1",
            identities=(config.email,),
            status="authorized",
        )
    )

    assert GraphTokenProvider(config, storage=storage).get_access_token() == "bound-token"


def test_graph_binding_for_another_configured_account_is_not_reused(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.accounts = [{"home_account_id": "old-home", "username": "old@example.com"}]

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [] if username is not None else list(self.accounts)

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, str]:
            return {"access_token": "old-token"}

        def remove_account(self, account: dict[str, Any]) -> None:
            self.accounts.remove(account)

    class MismatchVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            raise GraphIdentityMismatchError(
                f"Authorized Microsoft account old@example.com does not match {expected_account}"
            )

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    storage = FileGraphTokenStorage(config.cache_path)
    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account="old@example.com",
            home_account_id="old-home",
            status="authorized",
        )
    )

    provider = GraphTokenProvider(config, storage=storage, identity_verifier=MismatchVerifier())

    assert storage.load_binding() is None
    with pytest.raises(AuthRequiredError, match="old@example.com"):
        provider.get_access_token()
    assert provider._app.accounts == []


def test_graph_identity_transient_failure_preserves_cache_and_binding(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = False

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.accounts = [{"home_account_id": "home-1", "username": "person@outlook.com"}]

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return list(self.accounts)

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, str]:
            return {"access_token": "cached-token"}

        def remove_account(self, account: dict[str, Any]) -> None:
            raise AssertionError("transient identity failures must not remove cached accounts")

    class TransientVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            raise GraphIdentityTransientError("Microsoft Graph error 503")

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    storage = FileGraphTokenStorage(config.cache_path)
    binding = GraphAuthorizationBinding(
        expected_account=config.email,
        home_account_id="home-1",
        status="unbound",
    )
    storage.save_binding(binding)
    provider = GraphTokenProvider(config, storage=storage, identity_verifier=TransientVerifier())

    with pytest.raises(GraphIdentityTransientError, match="503"):
        provider.get_access_token()

    assert storage.load_binding() == binding
    assert provider._app.accounts[0]["home_account_id"] == "home-1"


def test_graph_identity_verifier_marks_service_failure_as_transient(tmp_path) -> None:
    class IdentitySession:
        def get(self, url: str, **kwargs: Any) -> FakeResponse:
            return FakeResponse(503, {"error": {"code": "ServiceUnavailable", "message": "Try later"}})

    verifier = RequestsGraphIdentityVerifier(_graph_config(tmp_path), IdentitySession())

    with pytest.raises(GraphIdentityTransientError, match="ServiceUnavailable"):
        verifier.verify("token", "person@outlook.com")


def test_graph_401_clears_only_the_bound_account(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            self.accounts = [
                {"home_account_id": "other", "username": "other@example.com"},
                {"home_account_id": "home-1", "username": "login@live.com"},
            ]

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return list(self.accounts)

        def remove_account(self, account: dict[str, Any]) -> None:
            self.accounts.remove(account)

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    storage = FileGraphTokenStorage(config.cache_path)
    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account=config.email,
            home_account_id="home-1",
            graph_user_id="graph-user-1",
            identities=(config.email,),
            status="authorized",
        )
    )
    provider = GraphTokenProvider(config, storage=storage)

    provider._clear_cached_accounts()

    assert [account["home_account_id"] for account in provider._app.accounts] == ["other"]
    assert storage.load_binding() is None


def test_graph_token_provider_uses_sqlite_storage(monkeypatch, tmp_path) -> None:
    class FakeCache:
        has_state_changed = True

        def deserialize(self, value: str) -> None:
            self.loaded = value

        def serialize(self) -> str:
            return '{"sqlite": true}'

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"username": username}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "sqlite-token"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "123456")
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(tmp_path / "ignored.json"))
    authorization_id = store.resolve_graph_authorization(
        record.id,
        config.email,
        graph_config_fingerprint(config),
    )
    storage = SQLiteGraphTokenStorage(store, authorization_id)

    assert GraphTokenProvider(config, storage=storage).get_access_token() == "sqlite-token"
    assert store.load_msal_cache(record.id) == '{"sqlite": true}'


def test_graph_mail_operations() -> None:
    session = FakeSession(
        [
            FakeResponse(200, {"value": [_message("m1")]}),
            FakeResponse(200, _message("m1", include_body=True)),
            FakeResponse(202, {}),
            FakeResponse(202, {}),
            FakeResponse(200, {"id": "moved-id"}),
            FakeResponse(204, {}),
        ]
    )
    client = _graph_client(session)

    assert client.mail_list()["items"][0]["subject"] == "Hello"
    mail_detail = client.mail_get("m1")
    assert mail_detail["body"] == "<p>Hello</p>"
    assert mail_detail["body_content_type"] == "html"
    assert client.mail_send(["a@example.com"], "Hi", "Body") == {"sent": True}
    assert client.mail_reply("m1", "<p>Thanks</p>", html=True)["replied"] is True
    assert client.mail_move("m1", "deleteditems")["id"] == "moved-id"
    assert client.mail_delete("m1", mode="hard_delete") == {"deleted": True, "mode": "hard_delete"}
    assert session.calls[2]["json"]["message"]["toRecipients"][0]["emailAddress"]["address"] == "a@example.com"
    assert session.calls[3]["json"] == {
        "message": {"body": {"contentType": "HTML", "content": "<p>Thanks</p>"}}
    }
    assert session.calls[5]["method"] == "POST"
    assert session.calls[5]["url"].endswith("/me/messages/m1/permanentDelete")


def test_graph_mail_get_falls_back_to_preview_when_body_content_is_empty() -> None:
    message = _message("m1", include_body=True)
    message["body"] = {"contentType": "html", "content": ""}
    message["bodyPreview"] = "Preview fallback"
    client = _graph_client(FakeSession([FakeResponse(200, message)]))

    result = client.mail_get("m1")

    assert result["body"] == "Preview fallback"
    assert result["body_content_type"] == "text"


def test_graph_mail_send_supports_content_base64_attachment() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    assert client.mail_send(
        ["a@example.com"],
        "Hi",
        "Body",
        attachments=[{"name": "note.txt", "content_base64": "SGk="}],
    ) == {"sent": True}

    attachment = session.calls[0]["json"]["message"]["attachments"][0]
    assert attachment == {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": "note.txt",
        "contentBytes": "SGk=",
    }


def test_graph_mail_send_supports_upload_attachment_and_inline(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    attached = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    inline = save_uploaded_file(
        credential_id=12,
        filename="inline.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)
    token = set_upload_credential_id(12)
    try:
        client.mail_send(
            ["a@example.com"],
            "Hi",
            '<img src="cid:photo1">',
            html=True,
            attachments=[
                {"relative_path": attached.relative_path, "disposition": "attachment"},
                {"relative_path": inline.relative_path, "disposition": "inline", "content_id": "photo1"},
            ],
        )
    finally:
        reset_upload_credential_id(token)

    attachments = session.calls[0]["json"]["message"]["attachments"]
    assert attachments[0]["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert attachments[0]["name"] == "photo.png"
    assert attachments[0]["contentBytes"] == base64.b64encode(PNG_BYTES).decode("ascii")
    assert attachments[0]["contentType"] == "image/png"
    assert "isInline" not in attachments[0]
    assert attachments[1]["isInline"] is True
    assert attachments[1]["contentId"] == "photo1"
    assert attachments[1]["contentType"] == "image/png"
    body = session.calls[0]["json"]["message"]["body"]["content"]
    assert body == '<img src="cid:photo1">'


def test_graph_mail_send_rejects_inline_upload_without_html(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    uploaded = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    client = _graph_client(FakeSession([]))
    token = set_upload_credential_id(12)
    try:
        with pytest.raises(ToolValidationError, match="html=true"):
            client.mail_send(
                ["a@example.com"],
                "Hi",
                "Body",
                html=False,
                attachments=[{"relative_path": uploaded.relative_path, "disposition": "inline", "content_id": "photo1"}],
            )
    finally:
        reset_upload_credential_id(token)


def test_graph_mail_search_trash_delete_and_errors(_skip_sleep) -> None:
    session = FakeSession(
        [
            FakeResponse(200, {"value": [_message("m2")]}),
            FakeResponse(200, {"id": "trash-id"}),
            FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow down"}}),
            FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow down"}}),
            FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow down"}}),
            FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow down"}}),
            FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow down"}}),
        ]
    )
    client = _graph_client(session)

    assert client.mail_search("Hello", limit=1, offset=1)["items"][0]["id"] == "m2"
    assert client.mail_delete("m1", mode="move_to_trash") == {"deleted": True, "mode": "move_to_trash"}
    with pytest.raises(GraphClientError, match="TooManyRequests"):
        client.request("GET", "/me/messages")
    with pytest.raises(ToolValidationError):
        client.mail_search("")


def test_graph_calendar_contact_and_task_operations() -> None:
    session = FakeSession(
        [
            FakeResponse(200, {"value": [_event("e1")]}),
            FakeResponse(201, _event("e2")),
            FakeResponse(200, _event("e2", subject="Updated")),
            FakeResponse(200, _event("e2", subject="Updated")),
            FakeResponse(204, {}),
            FakeResponse(200, {"value": [_contact("c1")]}),
            FakeResponse(201, _contact("c2")),
            FakeResponse(200, _contact("c2", job_title="Engineer")),
            FakeResponse(204, {}),
            FakeResponse(200, {"value": [{"id": "task-list", "wellknownListName": "defaultList"}]}),
            FakeResponse(200, {"value": [_task("t1")]}),
            FakeResponse(201, _task("t2")),
            FakeResponse(200, _task("t2", status="completed")),
            FakeResponse(204, {}),
        ]
    )
    client = _graph_client(session)

    listed = client.calendar_list("2026-01-01T00:00:00+00:00", "2026-01-03T00:00:00+00:00", timezone="Asia/Shanghai")
    created = client.calendar_create(
        "Meet",
        "2026-01-01T09:00:00+00:00",
        "2026-01-01T10:00:00+00:00",
        timezone="Asia/Shanghai",
    )
    assert listed["items"][0]["id"] == "e1"
    assert listed["items"][0]["start"] == "2026-01-01T09:00:00+00:00"
    assert listed["items"][0]["start_local"] == "2026-01-01T17:00:00+08:00"
    assert listed["items"][0]["local_timezone"] == "Asia/Shanghai"
    assert listed["items"][0]["body_plain"] == "Agenda"
    assert created["id"] == "e2"
    assert created["start"] == "2026-01-01T09:00:00+00:00"
    assert created["start_local"] == "2026-01-01T17:00:00+08:00"
    assert created["local_timezone"] == "Asia/Shanghai"
    assert created["body_content_type"] == "text"
    assert client.calendar_update("e2", {"subject": "Updated"})["subject"] == "Updated"
    assert client.calendar_delete("e2") == {"deleted": True, "id": "e2"}

    assert client.contacts_list()["items"][0]["display_name"] == "Ada Lovelace"
    assert client.contact_create("Grace Hopper", email_addresses=["grace@example.com"])["id"] == "c2"
    assert client.contact_update("c2", {"job_title": "Engineer"})["job_title"] == "Engineer"
    assert client.contact_delete("c2") == {"deleted": True, "id": "c2"}

    assert client.tasks_list()["items"][0]["subject"] == "Follow up"
    assert client.task_create("Write notes")["id"] == "t2"
    assert client.task_update("t2", {"complete": True})["is_complete"] is True
    assert client.task_delete("t2") == {"deleted": True, "id": "t2"}
    event_payload = session.calls[1]["json"]
    assert event_payload["attendees"] == []
    assert event_payload["isReminderOn"] is False
    assert "reminderMinutesBeforeStart" not in event_payload


def test_graph_calendar_event_respects_graph_timezone() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_event("e1", time_zone="China Standard Time")]})])
    client = _graph_client(session)

    listed = client.calendar_list("2026-01-01T00:00:00+00:00", "2026-01-03T00:00:00+00:00")

    assert listed["items"][0]["start"] == "2026-01-01T01:00:00+00:00"
    assert listed["items"][0]["start_local"] == "2026-01-01T09:00:00+08:00"
    assert listed["items"][0]["local_timezone"] == "Asia/Shanghai"


def test_graph_calendar_list_follows_next_link_and_reports_exact_has_more() -> None:
    next_link = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=next"
    session = FakeSession([
        FakeResponse(200, {"value": [_event("e1"), _event("e2")], "@odata.nextLink": next_link}),
        FakeResponse(200, {"value": [_event("e3")]}),
    ])
    client = _graph_client(session)

    result = client.calendar_list(
        "2026-01-01T00:00:00+00:00",
        "2026-01-03T00:00:00+00:00",
        limit=2,
    )

    assert [item["id"] for item in result["items"]] == ["e1", "e2"]
    assert result["has_more"] is True
    assert len(session.calls) == 2
    assert session.calls[0]["headers"]["Prefer"] == 'outlook.timezone="UTC"'


def test_graph_calendar_full_last_page_reports_no_more_items() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_event("e1"), _event("e2")]})])
    client = _graph_client(session)

    result = client.calendar_list(
        "2026-01-01T00:00:00+00:00",
        "2026-01-03T00:00:00+00:00",
        limit=2,
    )

    assert len(result["items"]) == 2
    assert result["has_more"] is False


def test_graph_all_day_event_preserves_graph_dates_without_timezone_conversion() -> None:
    event = _event("e1", time_zone="Pacific Standard Time")
    event["isAllDay"] = True
    event["start"]["dateTime"] = "2026-01-01T00:00:00.0000000"
    event["end"]["dateTime"] = "2026-01-02T00:00:00.0000000"
    session = FakeSession([FakeResponse(200, {"value": [event]})])
    client = _graph_client(session)

    listed = client.calendar_list(
        "2026-01-01T00:00:00+00:00",
        "2026-01-03T00:00:00+00:00",
        timezone="Asia/Shanghai",
    )

    item = listed["items"][0]
    assert item["start"] == "2026-01-01"
    assert item["end"] == "2026-01-02"
    assert item["start_local"] == "2026-01-01"


def test_graph_timed_event_with_unknown_timezone_is_not_guessed_as_utc() -> None:
    session = FakeSession([
        FakeResponse(200, {"value": [_event("e1", time_zone="Unsupported Standard Time")]})
    ])
    client = _graph_client(session)

    with pytest.raises(GraphClientError, match="was not guessed"):
        client.calendar_list(
            "2026-01-01T00:00:00+00:00",
            "2026-01-03T00:00:00+00:00",
        )


def test_graph_calendar_naive_datetime_uses_calendar_timezone() -> None:
    session = FakeSession(
        [
            FakeResponse(200, {"value": []}),
            FakeResponse(201, _event("e1")),
            FakeResponse(200, {"isAllDay": False}),
            FakeResponse(200, _event("e1")),
        ]
    )
    client = _graph_client(session)

    client.calendar_list("2026-03-16T00:00:00", "2026-03-17T00:00:00", timezone="Asia/Shanghai")
    client.calendar_create(
        "Meet",
        "2026-03-16T20:00:00",
        "2026-03-16T21:00:00",
        timezone="Asia/Shanghai",
    )
    client.calendar_update(
        "e1",
        {"start": "2026-03-16T20:00:00", "end": "2026-03-16T21:00:00"},
        timezone="Asia/Shanghai",
    )

    list_params = session.calls[0]["params"]
    assert list_params["startDateTime"] == "2026-03-15T16:00:00+00:00"
    assert list_params["endDateTime"] == "2026-03-16T16:00:00+00:00"

    create_payload = session.calls[1]["json"]
    assert create_payload["start"] == {"dateTime": "2026-03-16T20:00:00", "timeZone": "Asia/Shanghai"}
    assert create_payload["end"] == {"dateTime": "2026-03-16T21:00:00", "timeZone": "Asia/Shanghai"}

    update_payload = session.calls[3]["json"]
    assert update_payload["start"] == {"dateTime": "2026-03-16T20:00:00", "timeZone": "Asia/Shanghai"}
    assert update_payload["end"] == {"dateTime": "2026-03-16T21:00:00", "timeZone": "Asia/Shanghai"}


def test_graph_calendar_offset_datetime_wins_over_timezone_argument() -> None:
    session = FakeSession([FakeResponse(201, _event("e1"))])
    client = _graph_client(session)

    client.calendar_create(
        "Meet",
        "2026-03-16T20:00:00+08:00",
        "2026-03-16T21:00:00+08:00",
        timezone="America/New_York",
    )

    event_payload = session.calls[0]["json"]
    assert event_payload["start"] == {"dateTime": "2026-03-16T08:00:00", "timeZone": "America/New_York"}
    assert event_payload["end"] == {"dateTime": "2026-03-16T09:00:00", "timeZone": "America/New_York"}


def test_graph_calendar_create_can_enable_reminder() -> None:
    session = FakeSession([FakeResponse(201, _event("e1"))])
    client = _graph_client(session)

    client.calendar_create(
        "Meet",
        "2026-03-16T20:00:00+08:00",
        "2026-03-16T21:00:00+08:00",
        reminder_minutes_before_start=10,
    )

    event_payload = session.calls[0]["json"]
    assert event_payload["isReminderOn"] is True
    assert event_payload["reminderMinutesBeforeStart"] == 10


def test_graph_calendar_create_supports_upload_attachment_and_inline(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    attached = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    inline = save_uploaded_file(
        credential_id=12,
        filename="inline.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    session = FakeSession([FakeResponse(201, _event("e1")), FakeResponse(201, {"id": "a1"}), FakeResponse(201, {"id": "a2"})])
    client = _graph_client(session)
    try:
        result = client.calendar_create(
            "Meet",
            "2026-03-16T20:00:00+08:00",
            "2026-03-16T21:00:00+08:00",
            body='<p>Agenda <img src="cid:photo1"></p>',
            html=True,
            attachments=[
                {"relative_path": attached.relative_path, "disposition": "attachment"},
                {"relative_path": inline.relative_path, "disposition": "inline", "content_id": "photo1"},
            ],
        )
    finally:
        reset_upload_credential_id(token)

    assert result["id"] == "e1"
    assert session.calls[0]["url"].endswith("/me/events")
    assert session.calls[0]["headers"]["Prefer"] == 'outlook.timezone="UTC"'
    assert session.calls[1]["url"].endswith("/me/events/e1/attachments")
    assert session.calls[2]["url"].endswith("/me/events/e1/attachments")
    body_payload = session.calls[0]["json"]["body"]
    assert body_payload["contentType"] == "HTML"
    assert body_payload["content"] == '<p>Agenda <img src="cid:photo1"></p>'
    attachment_payload = session.calls[1]["json"]
    inline_payload = session.calls[2]["json"]
    assert attachment_payload["name"] == "photo.png"
    assert attachment_payload["contentBytes"] == base64.b64encode(PNG_BYTES).decode("ascii")
    assert "isInline" not in attachment_payload
    assert inline_payload["name"] == "inline.png"
    assert inline_payload["isInline"] is True
    assert inline_payload["contentId"] == "photo1"


def test_graph_calendar_create_validates_all_attachments_before_remote_write() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="Invalid base64"):
        client.calendar_create(
            "Meet",
            "2026-03-16T20:00:00+08:00",
            "2026-03-16T21:00:00+08:00",
            attachments=[{"name": "bad.txt", "content_base64": "@@@@"}],
        )

    assert session.calls == []


def test_graph_calendar_create_reports_partial_attachment_failure_with_event_id() -> None:
    session = FakeSession([
        FakeResponse(201, _event("e1")),
        FakeResponse(400, {"error": {"code": "InvalidAttachment", "message": "Rejected"}}),
    ])
    client = _graph_client(session)

    result = client.calendar_create(
        "Meet",
        "2026-03-16T20:00:00+08:00",
        "2026-03-16T21:00:00+08:00",
        attachments=[{"name": "note.txt", "content_base64": "SGk="}],
    )

    assert result["id"] == "e1"
    assert result["event_id"] == "e1"
    assert result["status"] == "partial_success"
    assert result["partial_success"] is True
    assert result["attachment_errors"] == ["attachment #0: Microsoft Graph error 400 (InvalidAttachment): Rejected"]


def test_graph_calendar_update_preserves_html_payload_when_requested() -> None:
    session = FakeSession([FakeResponse(200, _event("e1", body="<p>Agenda</p>", body_content_type="html"))])
    client = _graph_client(session)

    result = client.calendar_update("e1", {"body": "<p>Agenda</p>"}, html=True)

    assert result["body_content_type"] == "html"
    assert result["body_plain"] == "Agenda"
    body_payload = session.calls[0]["json"]["body"]
    assert body_payload["contentType"] == "HTML"
    assert body_payload["content"] == "<p>Agenda</p>"


def test_graph_calendar_partial_update_preserves_existing_all_day_semantics() -> None:
    from outlook_mcp.tools import calendar

    updated_event = _event("e1")
    updated_event["isAllDay"] = True
    updated_event["start"] = {"dateTime": "2026-03-17T00:00:00", "timeZone": "UTC"}
    updated_event["end"] = {"dateTime": "2026-03-18T00:00:00", "timeZone": "UTC"}
    session = FakeSession([
        FakeResponse(200, {"isAllDay": True}),
        FakeResponse(200, updated_event),
    ])
    client = _graph_client(session)

    result = calendar.calendar_update(
        client,
        "e1",
        {"start": "2026-03-17", "end": "2026-03-18"},
        confirm=True,
        timezone="Asia/Shanghai",
    )

    assert result["ok"] is True
    assert result["data"]["start"] == "2026-03-17"
    assert session.calls[0]["params"] == {"$select": "isAllDay"}
    assert session.calls[0]["headers"]["Prefer"] == 'outlook.timezone="UTC"'
    assert session.calls[1]["json"]["start"] == {
        "dateTime": "2026-03-17T00:00:00",
        "timeZone": "Asia/Shanghai",
    }
    assert session.calls[1]["headers"]["Prefer"] == 'outlook.timezone="UTC"'


def test_graph_all_day_create_converts_offset_datetime_in_requested_timezone() -> None:
    response = _event("e1")
    response["isAllDay"] = True
    response["start"] = {"dateTime": "2026-03-17T00:00:00", "timeZone": "UTC"}
    response["end"] = {"dateTime": "2026-03-18T00:00:00", "timeZone": "UTC"}
    session = FakeSession([FakeResponse(201, response)])
    client = _graph_client(session)

    result = client.calendar_create(
        "All day",
        "2026-03-16T16:00:00+00:00",
        "2026-03-17T16:00:00+00:00",
        is_all_day=True,
        timezone="Asia/Shanghai",
    )

    assert session.calls[0]["json"]["start"] == {
        "dateTime": "2026-03-17T00:00:00",
        "timeZone": "Asia/Shanghai",
    }
    assert result["start"] == "2026-03-17"


def test_graph_calendar_create_with_attendees_sends_mime_invite() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    result = client.calendar_create(
        "Meet",
        "2026-01-01T09:00:00+00:00",
        "2026-01-01T10:00:00+00:00",
        body="Agenda",
        location="Online",
        required_attendees=["a@example.com"],
        optional_attendees=["b@example.com"],
        timezone="America/New_York",
    )

    assert result["sent"] is True
    assert result["id"] is None
    assert result["ics_uid"].endswith("@outlook-mcp.local")
    assert result["required_attendees"] == ["a@example.com"]
    assert result["optional_attendees"] == ["b@example.com"]
    assert result["start"] == "2026-01-01T09:00:00+00:00"
    assert result["start_local"] == "2026-01-01T04:00:00-05:00"
    assert result["local_timezone"] == "America/New_York"
    assert result["body_content_type"] == "text"
    assert result["body_plain"] == "Agenda"
    assert session.calls[0]["method"] == "POST"
    assert session.calls[0]["url"].endswith("/me/sendMail")
    assert session.calls[0]["headers"]["Content-Type"] == "text/plain"

    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    calendar_text = _unfold_ics(calendar_part.get_content())

    assert message["To"] == "a@example.com, b@example.com"
    assert "METHOD:REQUEST" in calendar_text
    assert "VALARM" not in calendar_text
    assert "ATTENDEE;CN=a@example.com;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:a@example.com" in calendar_text
    assert "ATTENDEE;CN=b@example.com;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:b@example.com" in calendar_text


def test_graph_calendar_create_rejects_suppressed_invites_with_attendees() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="send_invites=false"):
        client.calendar_create(
            "Meet",
            "2026-01-01T09:00:00+00:00",
            "2026-01-01T10:00:00+00:00",
            required_attendees=["a@example.com"],
            send_invites=False,
        )

    assert session.calls == []


def test_graph_calendar_update_rejects_suppressed_invites() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="send_invites=false"):
        client.calendar_update("e1", {"subject": "Updated"}, send_invites=False)

    assert session.calls == []


def test_graph_calendar_invite_can_enable_reminder() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    client.calendar_create(
        "Meet",
        "2026-01-01T09:00:00+00:00",
        "2026-01-01T10:00:00+00:00",
        required_attendees=["a@example.com"],
        reminder_minutes_before_start=10,
    )

    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    calendar_text = _unfold_ics(calendar_part.get_content())

    assert "BEGIN:VALARM" in calendar_text
    assert "TRIGGER:-PT10M" in calendar_text
    assert "ATTENDEE;CN=a@example.com;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:a@example.com" in calendar_text


def test_graph_calendar_invite_supports_mime_upload_attachment_and_inline(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_UPLOAD_DIR", str(tmp_path / "uploads"))
    attached = save_uploaded_file(
        credential_id=12,
        filename="photo.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    inline = save_uploaded_file(
        credential_id=12,
        filename="inline.png",
        mime_type="image/png",
        content=PNG_BYTES,
    )
    token = set_upload_credential_id(12)
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)
    try:
        result = client.calendar_create(
            "Meet",
            "2026-01-01T09:00:00+00:00",
            "2026-01-01T10:00:00+00:00",
            body='<p>Agenda <img src="cid:photo1"></p>',
            required_attendees=["a@example.com"],
            html=True,
            attachments=[
                {"relative_path": attached.relative_path, "disposition": "attachment"},
                {"relative_path": inline.relative_path, "disposition": "inline", "content_id": "photo1"},
            ],
        )
    finally:
        reset_upload_credential_id(token)

    assert result["sent"] is True
    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    html_part = next(part for part in message.walk() if part.get_content_type() == "text/html")
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    attachment_part = next(part for part in message.walk() if part.get_filename() == "photo.png")
    inline_part = next(part for part in message.walk() if part.get_filename() == "inline.png")

    assert '<img src="cid:photo1">' in html_part.get_content()
    assert "METHOD:REQUEST" in calendar_part.get_content()
    assert attachment_part.get_content_disposition() == "attachment"
    assert attachment_part.get_content_type() == "image/png"
    assert attachment_part.get_payload(decode=True) == PNG_BYTES
    assert inline_part.get_content_disposition() == "inline"
    assert inline_part["Content-ID"] == "<photo1>"
    assert inline_part.get_content_type() == "image/png"
    assert inline_part.get_payload(decode=True) == PNG_BYTES


def test_graph_calendar_invite_naive_datetime_uses_timezone_in_ics() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    result = client.calendar_create(
        "Meet",
        "2026-03-16T20:00:00",
        "2026-03-16T21:00:00",
        required_attendees=["a@example.com"],
        timezone="Asia/Shanghai",
    )

    assert result["start"] == "2026-03-16T12:00:00+00:00"
    assert result["start_local"] == "2026-03-16T20:00:00+08:00"
    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    calendar_text = _unfold_ics(calendar_part.get_content())
    assert "DTSTART:20260316T120000Z" in calendar_text
    assert "DTEND:20260316T130000Z" in calendar_text


def test_graph_payload_edges_and_filters() -> None:
    session = FakeSession(
        [
            FakeResponse(200, {"value": [_contact("c1")]}),
            FakeResponse(200, _contact("c1", job_title="Architect")),
            FakeResponse(200, {"value": [{"id": "task-list", "wellknownListName": "defaultList"}]}),
            FakeResponse(200, {"value": [_task("t1", status="inProgress")]}),
            FakeResponse(201, _task("t2", status="inProgress")),
            FakeResponse(
                200,
                {
                    "attendees": [
                        {"type": "optional", "emailAddress": {"address": "b@example.com"}},
                    ]
                },
            ),
            FakeResponse(200, _event("e1")),
        ]
    )
    client = _graph_client(session)

    assert client.contacts_search("Ada's")["items"][0]["id"] == "c1"
    updated = client.contact_update(
        "c1",
        {
            "phone_numbers": [
                {"label": "MobilePhone", "phone_number": "555"},
                {"label": "HomePhone", "phone_number": "777"},
            ],
            "job_title": "Architect",
        },
    )
    assert updated["job_title"] == "Architect"
    assert client.tasks_list(status="InProgress")["items"][0]["status"] == "InProgress"
    assert client.task_create("Half done", percent_complete=50)["status"] == "InProgress"
    assert client.calendar_update("e1", {"required_attendees": ["a@example.com"], "is_all_day": True})["id"] == "e1"

    contact_patch = session.calls[1]["json"]
    assert contact_patch["mobilePhone"] == "555"
    assert contact_patch["homePhones"] == ["777"]
    assert "contains(displayName,'Ada''s')" in session.calls[0]["params"]["$filter"]
    assert session.calls[5]["method"] == "GET"
    assert session.calls[5]["params"] == {"$select": "attendees"}
    event_patch = session.calls[6]["json"]
    assert event_patch["attendees"][0]["emailAddress"] == {"address": "a@example.com", "name": "a@example.com"}
    assert event_patch["attendees"][1]["emailAddress"] == {"address": "b@example.com", "name": "b@example.com"}


def test_graph_contact_phone_update_replaces_and_can_clear_all_categories() -> None:
    session = FakeSession([
        FakeResponse(200, _contact("c1")),
        FakeResponse(200, _contact("c1")),
    ])
    client = _graph_client(session)

    client.contact_update("c1", {"phone_numbers": [{"label": "MobilePhone", "phone_number": "555"}]})
    client.contact_update("c1", {"phone_numbers": []})

    assert session.calls[0]["json"] == {
        "businessPhones": [],
        "homePhones": [],
        "mobilePhone": "555",
    }
    assert session.calls[1]["json"] == {
        "businessPhones": [],
        "homePhones": [],
        "mobilePhone": None,
    }


def test_graph_contact_phone_update_rejects_unknown_label_before_write() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="Unsupported phone number label"):
        client.contact_update("c1", {"phone_numbers": [{"label": "MobliePhone", "phone_number": "555"}]})

    assert session.calls == []


def test_graph_calendar_update_with_full_attendee_lists_does_not_fetch_existing_event() -> None:
    session = FakeSession([FakeResponse(200, _event("e1"))])
    client = _graph_client(session)

    client.calendar_update(
        "e1",
        {
            "required_attendees": ["a@example.com"],
            "optional_attendees": [],
        },
    )

    assert len(session.calls) == 1
    assert session.calls[0]["method"] == "PATCH"


def test_graph_default_task_list_retry_after_transient_failure() -> None:
    session = FakeSession(
        [
            FakeResponse(503, {"error": {"message": "temporary"}}),
            FakeResponse(200, {"value": [{"id": "list-1", "wellknownListName": "defaultList"}]}),
            FakeResponse(200, {"value": [_task("t1")]}),
        ]
    )
    client = _graph_client(session)
    client._MAX_RETRIES = 0

    with pytest.raises(GraphClientError):
        _ = client.default_task_list_id

    assert client.default_task_list_id == "list-1"
    assert client.tasks_list()["items"][0]["id"] == "t1"


def test_graph_rejects_invalid_base64_attachment_content() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="Invalid base64"):
        client.mail_send(["a@example.com"], "Hi", "Body", attachments=[{"name": "bad.txt", "content_base64": "@@@@"}])


def test_graph_tool_auth_required_response() -> None:
    from outlook_mcp.tools import mail

    client = _graph_client(
        FakeSession([FakeResponse(200, {"value": []})]),
    )
    # Patch mail_list on the GraphClient instance to simulate auth error

    def raising_mail_list(**kwargs: Any) -> dict[str, Any]:
        raise AuthRequiredError("Login required", {"user_code": "ABCD"})

    client.mail_list = raising_mail_list  # type: ignore[method-assign]
    result = mail.mail_list(client)
    assert result["ok"] is False
    assert result["error"] == "auth_required"
    assert result["data"]["user_code"] == "ABCD"


def test_graph_tool_delegation_success() -> None:
    from outlook_mcp.tools import calendar, contacts, mail, tasks

    client = _graph_client(
        FakeSession([
            FakeResponse(200, {"value": [{"id": "m1"}]}),
            FakeResponse(200, {"value": [{"id": "e1"}]}),
            FakeResponse(200, {"value": [{"id": "c1"}]}),
            FakeResponse(200, {"value": [{"id": "t1"}]}),
        ]),
    )
    client._default_task_list_id = "list-abc"

    assert mail.mail_list(client)["data"]["items"][0]["id"] == "m1"
    assert calendar.calendar_list(client)["data"]["items"][0]["id"] == "e1"
    assert contacts.contacts_list(client)["data"]["items"][0]["id"] == "c1"
    assert tasks.tasks_list(client)["data"]["items"][0]["id"] == "t1"


def _graph_config(tmp_path: Path) -> GraphConfig:
    return GraphConfig(
        email="person@outlook.com",
        client_id="client-id",
        token_cache_path=str(tmp_path / "cache.json"),
    )


def _graph_client(session: FakeSession) -> GraphClient:
    return GraphClient(_graph_config(Path(".")), session=session, token_provider=StaticTokenProvider())


def _message(item_id: str, include_body: bool = False) -> dict[str, Any]:
    payload = {
        "id": item_id,
        "changeKey": "ck",
        "subject": "Hello",
        "from": {"emailAddress": {"address": "from@example.com"}},
        "sender": {"emailAddress": {"address": "sender@example.com"}},
        "receivedDateTime": "2026-01-01T00:00:00Z",
        "sentDateTime": "2026-01-01T00:01:00Z",
        "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
        "ccRecipients": [],
        "importance": "normal",
        "isRead": False,
        "hasAttachments": False,
    }
    if include_body:
        payload["body"] = {"contentType": "html", "content": "<p>Hello</p>"}
    return payload


def _event(
    item_id: str,
    subject: str = "Planning",
    time_zone: str | None = None,
    *,
    body: str = "Agenda",
    body_content_type: str | None = "text",
) -> dict[str, Any]:
    start = {"dateTime": "2026-01-01T09:00:00"}
    end = {"dateTime": "2026-01-01T10:00:00"}
    if time_zone is not None:
        start["timeZone"] = time_zone
        end["timeZone"] = time_zone
    payload_body: dict[str, Any] = {"content": body}
    if body_content_type is not None:
        payload_body["contentType"] = body_content_type
    return {
        "id": item_id,
        "changeKey": "ck",
        "subject": subject,
        "body": payload_body,
        "start": start,
        "end": end,
        "location": {"displayName": "Room 1"},
        "isAllDay": False,
        "isCancelled": False,
        "attendees": [{"type": "required", "emailAddress": {"address": "a@example.com"}}],
    }


def _contact(item_id: str, job_title: str | None = None) -> dict[str, Any]:
    return {
        "id": item_id,
        "@odata.etag": "etag",
        "displayName": "Ada Lovelace",
        "givenName": "Ada",
        "surname": "Lovelace",
        "jobTitle": job_title,
        "emailAddresses": [{"name": "EmailAddress1", "address": "ada@example.com"}],
        "businessPhones": ["123"],
    }


def _task(item_id: str, status: str = "notStarted") -> dict[str, Any]:
    return {
        "id": item_id,
        "@odata.etag": "etag",
        "title": "Follow up",
        "status": status,
        "body": {"content": "Notes"},
        "dueDateTime": {"dateTime": "2026-01-02"},
    }


def test_graph_mail_search_kql_escaping() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_message("m1")]})])
    client = _graph_client(session)
    client.mail_search("test 'injection' query", limit=1)
    filter_val = session.calls[0]["params"]["$filter"]
    assert "''" in filter_val  # single quotes doubled for OData

def test_graph_mail_search_query_truncated() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_message("m1")]})])
    client = _graph_client(session)
    long_query = "x" * 500
    client.mail_search(long_query, limit=1)
    filter_val = session.calls[0]["params"]["$filter"]
    assert len(filter_val) <= _MAX_QUERY_LEN + len("contains(subject,'')")

def test_graph_contacts_search_query_truncated() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_contact("c1")]})])
    client = _graph_client(session)
    long_query = "y" * 500
    client.contacts_search(long_query, limit=1)
    filter_val = session.calls[0]["params"]["$filter"]
    assert len(filter_val) <= _MAX_QUERY_LEN + len("contains(displayName,'')")

def test_odata_string_strips_double_quotes() -> None:
    assert _odata_string('hello "world"') == "hello world"


def test_odata_string_escapes_backslashes() -> None:
    result = _odata_string(r"path\to\file")
    assert "\\\\" in result


def test_odata_string_truncates() -> None:
    result = _odata_string("a" * 500)
    assert len(result) == _MAX_QUERY_LEN


def test_odata_string_single_quote_escaping() -> None:
    assert _odata_string("O'Brien") == "O''Brien"
    assert _odata_string("it's") == "it''s"
    assert _odata_string("'hello'") == "''hello''"
    assert _odata_string("a'b'c") == "a''b''c"
    assert _odata_string("noquotes") == "noquotes"


def test_odata_string_edge_cases() -> None:
    assert _odata_string("''''''") == "''''''''''''"
    assert _odata_string("") == ""
    assert _odata_string("你好'world") == "你好''world"
    assert _odata_string(r"test\backslash") == r"test\\backslash"
    assert _odata_string('""""') == ""
    assert _odata_string("中文搜索") == "中文搜索"


def test_odata_string_combined_escape_and_truncate() -> None:
    base = "a" * (_MAX_QUERY_LEN - 2) + "'" + "b" * 20
    result = _odata_string(base)
    assert len(result) == _MAX_QUERY_LEN


class TestListItemsPagination:
    def test_follows_odata_next_link(self) -> None:
        page1 = {
            "value": [{"id": "a"}, {"id": "b"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$top=2&$skip=2",
        }
        page2 = {"value": [{"id": "c"}, {"id": "d"}]}
        session = FakeSession([FakeResponse(200, page1), FakeResponse(200, page2)])
        client = _graph_client(session)
        result = client.contacts_list(limit=4, offset=0)
        assert len(result["items"]) == 4
        assert [item["id"] for item in result["items"]] == ["a", "b", "c", "d"]
        assert result["has_more"] is False
        assert len(session.calls) == 2

    def test_stops_when_limit_reached_across_pages(self) -> None:
        page1 = {
            "value": [{"id": "a"}, {"id": "b"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$top=3&$skip=3",
        }
        page2 = {"value": [{"id": "c"}, {"id": "d"}]}
        session = FakeSession([FakeResponse(200, page1), FakeResponse(200, page2)])
        client = _graph_client(session)
        result = client.contacts_list(limit=3, offset=0)
        assert len(result["items"]) == 3
        assert [item["id"] for item in result["items"]] == ["a", "b", "c"]
        assert result["has_more"] is True

    def test_no_next_link_returns_single_page(self) -> None:
        page1 = {"value": [{"id": "x"}]}
        session = FakeSession([FakeResponse(200, page1)])
        client = _graph_client(session)
        result = client.contacts_list(limit=1)
        assert len(result["items"]) == 1
        assert result["has_more"] is False
        assert len(session.calls) == 1


def test_ics_value_escapes_special_characters() -> None:
    assert _ics_value(r"a\b") == "a\\\\b"
    assert _ics_value("a;b") == "a\\;b"
    assert _ics_value("a,b") == "a\\,b"
    assert _ics_value("a\nb") == "a\\nb"
    assert _ics_value("a\r\nb") == "a\\nb"


def test_ics_param_value_wraps_when_special_chars_present() -> None:
    assert _ics_param_value("plain") == "plain"
    assert _ics_param_value("has,comma") == '"has\\,comma"'
    assert _ics_param_value("has;semi") == '"has\\;semi"'
    assert _ics_param_value('say "hi",ok') == '"say \'hi\'\\,ok"'


def test_html_to_plain_text_replaces_br_variants() -> None:
    assert _html_to_plain_text("line1<br>line2") == "line1\nline2"
    assert _html_to_plain_text("line1<br/>line2") == "line1\nline2"
    assert _html_to_plain_text("line1<br />line2") == "line1\nline2"
    assert _html_to_plain_text("no break") == "no break"
    assert _html_to_plain_text("<p>Hello&nbsp;world</p><p>Again</p>") == "Hello world\nAgain"
    assert _html_to_plain_text("<ul><li>One</li><li>Two</li></ul>") == "- One\n- Two"
    assert _html_to_plain_text("&lt;p&gt;Hello&lt;br&gt;World&lt;/p&gt;") == "Hello\nWorld"


def testgraph_error_message_formats_all_branches() -> None:
    assert graph_error_message(400, {"error": {"code": "BadRequest", "message": "Bad"}}) == "Microsoft Graph error 400 (BadRequest): Bad"
    assert graph_error_message(429, {"error": {"message": "Slow"}}) == "Microsoft Graph error 429: Slow"
    assert graph_error_message(500, {"error": {"code": "InternalError"}}) == "Microsoft Graph error 500 (InternalError): InternalError"
    assert graph_error_message(503, {"error": "plain"}) == "Microsoft Graph error 503"
    assert graph_error_message(404, {}) == "Microsoft Graph error 404"


# -- new mail operations (Graph path) --


def test_graph_mail_forward() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)
    result = client.mail_forward("m1", to=["a@example.com"], body="Fwd")
    assert result == {"forwarded": True}
    call = session.calls[0]
    assert call["url"].endswith("/me/messages/m1/forward")
    assert call["json"] == {
        "comment": "Fwd",
        "toRecipients": [
            {"emailAddress": {"address": "a@example.com"}},
        ],
    }


def test_graph_mail_forward_converts_html_comment_to_plain_text() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    result = client.mail_forward(
        "m1",
        to=["a@example.com"],
        body="<p>Hello&nbsp;world</p><ul><li>One</li><li>Two</li></ul>",
        html=True,
    )

    assert result == {"forwarded": True}
    assert session.calls[0]["json"] == {
        "comment": "Hello world\n- One\n- Two",
        "toRecipients": [
            {"emailAddress": {"address": "a@example.com"}},
        ],
    }


def test_graph_mail_mark_read() -> None:
    session = FakeSession([FakeResponse(200, {})])
    client = _graph_client(session)
    result = client.mail_mark_read("m1")
    assert result == {"marked": "read", "id": "m1"}
    assert session.calls[0]["json"]["isRead"] is True


def test_graph_mail_mark_unread() -> None:
    session = FakeSession([FakeResponse(200, {})])
    client = _graph_client(session)
    result = client.mail_mark_unread("m1")
    assert result == {"marked": "unread", "id": "m1"}
    assert session.calls[0]["json"]["isRead"] is False


def test_graph_mail_attachments() -> None:
    next_link = "https://graph.microsoft.com/v1.0/me/messages/m1/attachments?$skiptoken=next"
    session = FakeSession([
        FakeResponse(200, {
            "value": [{"name": "photo.png", "size": 1234, "contentType": "image/png", "id": "att1"}],
            "@odata.nextLink": next_link,
        }),
        FakeResponse(200, {
            "value": [{"name": "logo.jpg", "size": 567, "contentType": "image/jpeg", "id": "att2", "isInline": True}],
        }),
    ])
    client = _graph_client(session)
    result = client.mail_attachments("m1")
    assert result["items"][0]["name"] == "photo.png"
    assert result["items"][0]["size"] == 1234
    assert result["items"][1]["is_inline"] is True
    assert result["total"] == 2
    assert result["has_more"] is False
    assert len(session.calls) == 2
    assert session.calls[0]["url"].endswith("/me/messages/m1/attachments")


def test_graph_mail_folders() -> None:
    next_link = "https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=next"
    session = FakeSession([
        FakeResponse(200, {
            "value": [{"id": "custom-id", "displayName": "Projects", "totalItemCount": 10}],
            "@odata.nextLink": next_link,
        }),
        FakeResponse(200, {
            "value": [{"id": "sentitems", "displayName": "Sent", "unreadItemCount": 0}],
        }),
        FakeResponse(200, {"value": [_message("m1")]}),
        FakeResponse(200, {"id": "moved-id"}),
    ])
    client = _graph_client(session)
    result = client.mail_folders()
    assert result["items"][0] == {
        "id": "custom-id",
        "name": "Projects",
        "folder_class": "IPF.Note",
        "total_count": 10,
        "unread_count": 0,
        "child_folder_count": 0,
    }
    assert result["items"][1]["unread_count"] == 0
    assert result["total"] == 2
    assert result["has_more"] is False

    assert client.mail_list(folder=result["items"][0]["id"])["items"][0]["id"] == "m1"
    assert client.mail_move("m1", result["items"][0]["id"])["id"] == "moved-id"
    assert session.calls[2]["url"].endswith("/me/mailFolders/custom-id/messages")
    assert session.calls[3]["json"]["destinationId"] == "custom-id"


# -- new calendar / contact / task / account Graph-path operations --


def test_graph_calendar_get() -> None:
    session = FakeSession([FakeResponse(200, _event("e1", body="<p>Hello<br>World</p>", body_content_type="html"))])
    client = _graph_client(session)
    result = client.calendar_get("e1")
    assert result["id"] == "e1"
    assert result["subject"] == "Planning"
    assert result["body"] == "<p>Hello<br>World</p>"
    assert result["body_content_type"] == "html"
    assert result["body_plain"] == "Hello\nWorld"


def test_graph_calendar_get_extracts_full_html_document_body() -> None:
    body = (
        "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">"
        "<style>p { color: red; }</style></head><body><p>Hello<br>World</p></body></html>"
    )
    session = FakeSession([FakeResponse(200, _event("e1", body=body, body_content_type="html"))])
    client = _graph_client(session)

    result = client.calendar_get("e1")

    assert result["body"] == "<p>Hello<br>World</p>"
    assert result["body_content_type"] == "html"
    assert result["body_plain"] == "Hello\nWorld"


def test_graph_calendar_get_normalizes_exchange_empty_html_document_from_text_body() -> None:
    body = (
        "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">"
        "<meta name=\"Generator\" content=\"Microsoft Exchange Server\">"
        "<!-- converted from text -->"
        "<style><!-- .EmailQuote { margin-left: 1pt; padding-left: 4pt; border-left: #800000 2px solid; } --></style>"
        "</head><body><font size=\"2\"><span style=\"font-size:11pt;\">"
        "<div class=\"PlainText\">&nbsp;</div></span></font></body></html>"
    )
    session = FakeSession([FakeResponse(200, _event("e1", body=body, body_content_type="text"))])
    client = _graph_client(session)

    result = client.calendar_get("e1")

    assert result["body"] == ""
    assert result["body_content_type"] == "html"
    assert result["body_plain"] == ""


def test_graph_calendar_search() -> None:
    session = FakeSession([FakeResponse(200, {"value": [
        _event("e1", "Planning"),
        _event("e2", "Budget"),
    ]})])
    client = _graph_client(session)
    result = client.calendar_search(
        "Plan",
        start="2026-01-01T00:00:00+00:00",
        end="2026-01-03T00:00:00+00:00",
    )
    assert len(result["items"]) == 1
    assert result["items"][0]["subject"] == "Planning"
    assert session.calls[0]["url"].endswith("/me/calendarView")
    params = session.calls[0]["params"]
    assert params["startDateTime"] == "2026-01-01T00:00:00+00:00"
    assert params["endDateTime"] == "2026-01-03T00:00:00+00:00"
    assert "$filter" not in params


def test_graph_contact_get() -> None:
    session = FakeSession([FakeResponse(200, _contact("c1"))])
    client = _graph_client(session)
    result = client.contact_get("c1")
    assert result["id"] == "c1"
    assert result["display_name"] == "Ada Lovelace"


def test_graph_task_get() -> None:
    session = FakeSession(
        [FakeResponse(200, {"value": [{"id": "list-1", "wellknownListName": "defaultList"}]}),
         FakeResponse(200, _task("t1"))],
    )
    client = _graph_client(session)
    result = client.task_get("t1")
    assert result["id"] == "t1"
    assert result["subject"] == "Follow up"


def test_graph_task_search_server_side_filter() -> None:
    session = FakeSession([FakeResponse(200, {"value": [
        _task("t1", status="inProgress"), _task("t2", status="completed"),
    ]})])
    client = _graph_client(session)
    # Set cached task list so we skip the /me/todo/lists call
    client._default_task_list_id = "list-abc"
    result = client.task_search("Follow", limit=10)
    assert len(result["items"]) == 2
    params = session.calls[0]["params"]
    assert "contains(title,'Follow')" in params["$filter"]
    assert params["$top"] == 11


def test_graph_task_search_with_status() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_task("t1", status="inProgress")]})])
    client = _graph_client(session)
    client._default_task_list_id = "list-abc"
    client.task_search("Follow", status="InProgress")
    params = session.calls[0]["params"]
    assert "status eq 'inProgress'" in params["$filter"]
    assert "contains(title,'Follow')" in params["$filter"]


def test_graph_task_search_empty_query() -> None:
    session = FakeSession([FakeResponse(200, {"value": [_task("t1", status="completed")]})])
    client = _graph_client(session)
    client._default_task_list_id = "list-abc"
    result = client.task_search("  ", status="Completed")
    # empty query should not add a contains filter, only status
    assert "contains" not in session.calls[0]["params"].get("$filter", "")
    assert result["items"][0]["status"] == "Completed"


def test_graph_task_search_pagination_uses_skip() -> None:
    """task_search must use $skip for offset, not client-side slicing."""
    session = FakeSession([FakeResponse(200, {"value": [_task("t1"), _task("t2"), _task("t3")]})])
    client = _graph_client(session)
    client._default_task_list_id = "list-abc"
    result = client.task_search("Task", limit=10, offset=5)
    params = session.calls[0]["params"]
    assert params["$top"] == 11
    assert params["$skip"] == 5
    # Result should return what the server gave, not client-sliced
    assert len(result["items"]) == 3


def test_graph_account_info() -> None:
    session = FakeSession([FakeResponse(200, {"mail": "me@outlook.com", "displayName": "My Name"})])
    client = _graph_client(session)
    result = client.account_info()
    assert result["primary_smtp_address"] == "me@outlook.com"
    assert result["display_name"] == "My Name"
    assert session.calls[0]["url"] == "https://graph.microsoft.com/v1.0/me"


def test_graph_account_info_with_user_principal_name() -> None:
    session = FakeSession([FakeResponse(200, {"userPrincipalName": "me@outlook.com", "displayName": "Me"})])
    client = _graph_client(session)
    result = client.account_info()
    assert result["primary_smtp_address"] == "me@outlook.com"


# --- retry logic ---


@pytest.fixture
def _skip_sleep(monkeypatch) -> None:
    """Skip time.sleep in graph_client to make retry tests fast."""
    import outlook_mcp.graph_client as gc_mod

    monkeypatch.setattr(gc_mod, "time", _FakeTime())


class _FakeTime:
    """Fake time module: time() returns 0, sleep() is a no-op."""

    @staticmethod
    def sleep(seconds: float) -> None:
        pass

    @staticmethod
    def time() -> float:
        return 0.0


def test_retry_on_429_raises_after_exhaustion(_skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
    ])
    client = _graph_client(session)
    with pytest.raises(GraphClientError, match="TooManyRequests"):
        client.request("GET", "/me/messages")


def test_retry_on_429_succeeds_after_recovery(_skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(200, {"value": []}),
    ])
    client = _graph_client(session)
    result = client.request("GET", "/me/messages")
    assert result == {"value": []}
    assert len(session.calls) == 2


def test_retry_on_5xx_with_backoff(_skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(503, {"error": {"code": "ServiceUnavailable", "message": "Down"}}),
        FakeResponse(503, {"error": {"code": "ServiceUnavailable", "message": "Down"}}),
        FakeResponse(200, {"value": [{"id": "ok"}]}),
    ])
    client = _graph_client(session)
    result = client.request("GET", "/me/events")
    assert result == {"value": [{"id": "ok"}]}
    assert len(session.calls) == 3


def test_retry_respects_retry_after_header(_skip_sleep) -> None:
    """Retry-After header (seconds) should override exponential backoff."""
    session = FakeSession([
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}, headers={"Retry-After": "1"}),
        FakeResponse(200, {"value": []}),
    ])
    client = _graph_client(session)
    result = client.request("GET", "/me/messages")
    assert result == {"value": []}


def test_no_retry_on_4xx(_skip_sleep) -> None:
    """400, 403, 404 are not retryable — should raise immediately."""
    session = FakeSession([FakeResponse(400, {"error": {"code": "BadRequest", "message": "Invalid"}})])
    client = _graph_client(session)
    with pytest.raises(GraphClientError, match="BadRequest"):
        client.request("GET", "/me/messages")


def test_no_retry_on_403(_skip_sleep) -> None:
    session = FakeSession([FakeResponse(403, {"error": {"code": "Forbidden", "message": "No access"}})])
    client = _graph_client(session)
    with pytest.raises(GraphClientError, match="Forbidden"):
        client.request("GET", "/me/messages")


def test_retry_on_connection_error(_skip_sleep) -> None:
    import requests as req

    class _FailingSession:
        retries = 0

        def request(self, method: str, url: str, **kwargs: Any) -> FakeResponse:
            self.retries += 1
            if self.retries < 3:
                raise req.ConnectionError("connection refused")
            return FakeResponse(200, {"value": []})

    client = _graph_client(_FailingSession())
    result = client.request("GET", "/me/messages")
    assert result == {"value": []}


@pytest.mark.parametrize("method", ["POST", "PATCH", "DELETE"])
def test_write_5xx_is_not_replayed_and_reports_unknown_outcome(method: str, _skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(503, {"error": {"code": "ServiceUnavailable", "message": "Down"}}),
    ])
    client = _graph_client(session)

    with pytest.raises(GraphOutcomeUnknownError, match="do not retry blindly"):
        client.request(method, "/me/messages/action", json={"value": True})

    assert len(session.calls) == 1


def test_post_transport_failure_is_not_replayed_and_reports_unknown_outcome(_skip_sleep) -> None:
    import requests as req

    class FailingWriteSession:
        calls = 0

        def request(self, method: str, url: str, **kwargs: Any) -> FakeResponse:
            self.calls += 1
            raise req.Timeout("response timeout")

    session = FailingWriteSession()
    client = _graph_client(session)

    with pytest.raises(GraphOutcomeUnknownError, match="write may have completed"):
        client.request("POST", "/me/sendMail", json={"message": {}})

    assert session.calls == 1


def test_send_mime_inner_does_not_retry_on_429(_skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(202, {}),
    ])
    client = _graph_client(session)
    with pytest.raises(GraphClientError, match="TooManyRequests"):
        client._send_mime_message(b"raw mime bytes")
    assert len(session.calls) == 1


def test_send_mime_inner_raises_after_retry_exhaustion(_skip_sleep) -> None:
    session = FakeSession([
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
        FakeResponse(429, {"error": {"code": "TooManyRequests", "message": "Slow"}}),
    ])
    client = _graph_client(session)
    with pytest.raises(GraphClientError, match="TooManyRequests"):
        client._send_mime_message(b"raw mime bytes")


# --- calendar_delete cancellation path ---


def test_graph_calendar_delete_cancels_organized_meeting() -> None:
    session = FakeSession([
        FakeResponse(200, {
            "isOrganizer": True,
            "attendees": [{"type": "required", "emailAddress": {"address": "a@example.com"}}],
        }),
        FakeResponse(202, {}),
    ])
    client = _graph_client(session)

    result = client.calendar_delete("e1")

    assert result == {"deleted": True, "id": "e1"}
    assert session.calls[0]["method"] == "GET"
    assert session.calls[0]["url"].endswith("/me/events/e1")
    assert session.calls[0]["params"] == {"$select": "isOrganizer,attendees"}
    assert session.calls[1]["method"] == "POST"
    assert session.calls[1]["url"].endswith("/me/events/e1/cancel")


def test_graph_calendar_delete_removes_non_meeting_event() -> None:
    session = FakeSession([
        FakeResponse(200, {"isOrganizer": False, "attendees": []}),
        FakeResponse(204),
    ])
    client = _graph_client(session)

    result = client.calendar_delete("e1")

    assert result == {"deleted": True, "id": "e1"}
    assert session.calls[1]["method"] == "DELETE"
    assert session.calls[1]["url"].endswith("/me/events/e1")


def test_graph_calendar_delete_rejects_suppressed_cancellations() -> None:
    session = FakeSession([])
    client = _graph_client(session)

    with pytest.raises(ToolValidationError, match="cancellations"):
        client.calendar_delete("e1", send_cancellations=False)

    assert session.calls == []


# --- all-day meeting invite (MIME/ICS path) ---


def test_graph_calendar_create_all_day_with_attendees_sends_mime_invite() -> None:
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    result = client.calendar_create(
        "Offsite",
        "2026-01-01",
        "2026-01-03",
        is_all_day=True,
        required_attendees=["a@example.com"],
        timezone="Asia/Shanghai",
    )

    assert result["sent"] is True
    assert result["is_all_day"] is True
    assert result["start"] == "2026-01-01"
    assert result["end"] == "2026-01-03"
    assert session.calls[0]["url"].endswith("/me/sendMail")

    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    calendar_text = _unfold_ics(calendar_part.get_content())
    assert "DTSTART;VALUE=DATE:20260101" in calendar_text
    assert "DTEND;VALUE=DATE:20260103" in calendar_text
    assert "METHOD:REQUEST" in calendar_text


# --- ICS line folding (RFC 5545 75-octet limit) ---


def test_calendar_invite_ics_folds_long_lines_at_75_octets() -> None:
    subject = "季度评审会议" * 20
    body = "第一行议程\n" + "讨论" * 60
    session = FakeSession([FakeResponse(202, {})])
    client = _graph_client(session)

    client.calendar_create(
        subject,
        "2026-01-01T09:00:00+00:00",
        "2026-01-01T10:00:00+00:00",
        body=body,
        required_attendees=["a@example.com"],
    )

    mime_bytes = base64.b64decode(session.calls[0]["data"])
    message = BytesParser(policy=policy.default).parsebytes(mime_bytes)
    calendar_part = next(part for part in message.walk() if part.get_content_type() == "text/calendar")
    raw = calendar_part.get_content()

    for line in raw.replace("\r\n", "\n").split("\n"):
        assert len(line.encode("utf-8")) <= 75

    unfolded = _unfold_ics(raw)
    assert f"SUMMARY:{subject}" in unfolded
    assert "DESCRIPTION:第一行议程\\n" + "讨论" * 60 in unfolded


# --- 401 clears cache and retries with a fresh token ---


def test_graph_401_clears_cache_and_retries_with_fresh_token() -> None:
    provider = _RetryingTokenProvider(["stale-token", "fresh-token"])
    session = FakeSession([
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
        FakeResponse(200, {"value": []}),
    ])
    client = GraphClient(_graph_config(Path(".")), session=session, token_provider=provider)

    result = client.request("GET", "/me/messages")

    assert result == {"value": []}
    assert provider.cleared == 1
    assert len(session.calls) == 2
    assert session.calls[0]["headers"]["Authorization"] == "Bearer stale-token"
    assert session.calls[1]["headers"]["Authorization"] == "Bearer fresh-token"


def test_graph_401_retry_is_not_recursive() -> None:
    provider = _RetryingTokenProvider(["stale-token", "fresh-token"])
    session = FakeSession([
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
    ])
    client = GraphClient(_graph_config(Path(".")), session=session, token_provider=provider)

    with pytest.raises(GraphClientError, match="InvalidAuthenticationToken"):
        client.request("GET", "/me/messages")

    assert provider.cleared == 1
    assert len(session.calls) == 2


def test_send_mime_401_clears_cache_and_retries_once() -> None:
    provider = _RetryingTokenProvider(["stale-token", "fresh-token"])
    session = FakeSession([
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
        FakeResponse(202, {}),
    ])
    client = GraphClient(_graph_config(Path(".")), session=session, token_provider=provider)

    client._send_mime_message(b"raw mime bytes")

    assert provider.cleared == 1
    assert len(session.calls) == 2
    assert session.calls[1]["headers"]["Authorization"] == "Bearer fresh-token"


def test_send_mime_401_retry_is_not_recursive() -> None:
    provider = _RetryingTokenProvider(["stale-token", "fresh-token"])
    session = FakeSession([
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
        FakeResponse(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Expired"}}),
    ])
    client = GraphClient(_graph_config(Path(".")), session=session, token_provider=provider)

    with pytest.raises(GraphClientError, match="InvalidAuthenticationToken"):
        client._send_mime_message(b"raw mime bytes")

    assert provider.cleared == 1
    assert len(session.calls) == 2


# --- FileGraphTokenStorage corrupt file recovery ---


def test_file_token_storage_drops_corrupt_pending_flow(tmp_path) -> None:
    storage = FileGraphTokenStorage(tmp_path / "cache.json")
    pending = tmp_path / "cache.json.pending.json"
    pending.write_text("{not-json", encoding="utf-8")

    assert storage.load_pending_flow() is None
    assert not pending.exists()


def test_file_token_storage_drops_corrupt_binding(tmp_path) -> None:
    storage = FileGraphTokenStorage(tmp_path / "cache.json")
    binding = tmp_path / "cache.json.identity.json"
    binding.write_text("{not-json", encoding="utf-8")

    assert storage.load_binding() is None
    assert not binding.exists()


def test_graph_token_provider_quarantines_corrupt_cache(monkeypatch, tmp_path) -> None:
    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    cache_path = Path(config.token_cache_path)
    cache_path.write_text("{corrupt-json", encoding="utf-8")

    GraphTokenProvider(config)

    assert not cache_path.exists()
    backups = list(tmp_path.glob("cache.json.corrupt-*"))
    assert len(backups) == 1
    assert backups[0].read_text(encoding="utf-8") == "{corrupt-json"


def test_file_token_storage_save_cache_is_atomic_and_private(tmp_path) -> None:
    storage = FileGraphTokenStorage(tmp_path / "tokens" / "cache.json")

    storage.save_cache('{"cache": true}')

    cache_path = tmp_path / "tokens" / "cache.json"
    assert cache_path.read_text(encoding="utf-8") == '{"cache": true}'
    assert list((tmp_path / "tokens").glob("*.tmp")) == []
    if os.name != "nt":
        import stat

        assert stat.S_IMODE(os.stat(cache_path).st_mode) == 0o600
        assert stat.S_IMODE(os.stat(tmp_path / "tokens").st_mode) == 0o700


def test_timeout_session_applies_default_timeout(monkeypatch) -> None:
    from outlook_mcp.graph_auth import GRAPH_HTTP_TIMEOUT_SECONDS, TimeoutSession

    sent: dict[str, Any] = {}

    def fake_send(self: Any, request: Any, **kwargs: Any) -> Any:
        sent.update(kwargs)

        class Resp:
            status_code = 200

        return Resp()

    monkeypatch.setattr("requests.Session.send", fake_send)
    session = TimeoutSession()
    session.get("https://example.com/")
    assert sent["timeout"] == GRAPH_HTTP_TIMEOUT_SECONDS
    session.get("https://example.com/", timeout=5)
    assert sent["timeout"] == 5


def test_graph_token_provider_passes_timeout_http_client_to_msal(monkeypatch, tmp_path) -> None:
    from outlook_mcp.graph_auth import TimeoutSession

    captured: dict[str, Any] = {}

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            captured.update(kwargs)

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    GraphTokenProvider(_graph_config(tmp_path))

    assert isinstance(captured.get("http_client"), TimeoutSession)


def test_graph_pending_flow_concurrent_callers_share_single_poll(monkeypatch, tmp_path) -> None:
    import threading
    import time

    class FakeCache:
        has_state_changed = True

        def deserialize(self, _: str) -> None:
            pass

        def serialize(self) -> str:
            return "{}"

    poll_calls: list[int] = []

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"home_account_id": "home-1", "username": "person@outlook.com"}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any] | None:
            if poll_calls:
                return {"access_token": "shared-token"}
            return None

        def acquire_token_by_device_flow(self, flow: dict[str, Any], exit_condition: Any) -> dict[str, Any]:
            poll_calls.append(1)
            time.sleep(0.2)
            return {"access_token": "shared-token"}

    monkeypatch.setattr("outlook_mcp.graph_client.msal.SerializableTokenCache", FakeCache)
    monkeypatch.setattr("outlook_mcp.graph_client.msal.PublicClientApplication", FakeApp)
    config = _graph_config(tmp_path)
    Path(config.token_cache_path + ".pending.json").write_text(
        '{"user_code": "ABCD", "expires_at": 9999999999}', encoding="utf-8"
    )
    storage = FileGraphTokenStorage(config.cache_path)
    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account=config.email,
            home_account_id="home-1",
            status="authorized",
        )
    )

    class FakeVerifier:
        def verify(self, token: str, expected_account: str) -> GraphIdentity:
            return GraphIdentity(graph_user_id="graph-user-1", identities=(expected_account,))

    provider = GraphTokenProvider(config, storage=storage, identity_verifier=FakeVerifier())

    results: list[str] = []
    errors: list[BaseException] = []

    def worker() -> None:
        try:
            results.append(provider.get_access_token())
        except BaseException as exc:
            errors.append(exc)

    threads = [threading.Thread(target=worker) for _ in range(3)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=10)

    assert not errors
    assert results == ["shared-token", "shared-token", "shared-token"]
    assert len(poll_calls) == 1


def test_graph_task_update_supports_complete_false() -> None:
    from outlook_mcp.graph_client import _task_update_payload

    assert _task_update_payload({"complete": True}) == {"status": "completed"}
    assert _task_update_payload({"complete": False}) == {"status": "notStarted"}


def test_graph_attachment_per_file_limit() -> None:
    from outlook_mcp.graph_client import _resolved_attachment
    from outlook_mcp.uploads import DEFAULT_UPLOAD_MAX_BYTES

    oversized = base64.b64encode(b"x" * (DEFAULT_UPLOAD_MAX_BYTES + 1)).decode("ascii")
    with pytest.raises(ToolValidationError, match="per-attachment limit"):
        _resolved_attachment({"name": "big.bin", "content_base64": oversized})


def test_graph_attachments_total_budget() -> None:
    from outlook_mcp.graph_client import _resolve_attachments_with_budget
    from outlook_mcp.tools.common import MAX_ATTACHMENTS_TOTAL_BYTES

    chunk = base64.b64encode(b"x" * (MAX_ATTACHMENTS_TOTAL_BYTES // 2 + 1)).decode("ascii")
    with pytest.raises(ToolValidationError, match="per-message limit"):
        _resolve_attachments_with_budget(
            [
                {"name": "a.bin", "content_base64": chunk},
                {"name": "b.bin", "content_base64": chunk},
            ]
        )


@pytest.mark.skipif(os.name == "nt", reason="POSIX permissions")
def test_pending_flow_and_binding_file_permissions(tmp_path) -> None:
    import stat

    storage = FileGraphTokenStorage(tmp_path / "msal-token-cache.json")
    storage.save_pending_flow({"user_code": "ABCDEF"})
    pending = tmp_path / "msal-token-cache.json.pending.json"
    assert stat.S_IMODE(os.stat(pending).st_mode) == 0o600

    storage.save_binding(
        GraphAuthorizationBinding(
            expected_account="person@outlook.com",
            home_account_id=None,
            graph_user_id=None,
            identities=(),
            status="unbound",
        )
    )
    binding = tmp_path / "msal-token-cache.json.identity.json"
    assert stat.S_IMODE(os.stat(binding).st_mode) == 0o600
    assert stat.S_IMODE(os.stat(tmp_path).st_mode) == 0o700
