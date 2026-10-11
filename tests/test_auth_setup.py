from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from outlook_mcp.auth_setup import AuthSetupError, authorize_graph, setup_logging
from outlook_mcp.config import GraphConfig


REPO_ROOT = Path(__file__).resolve().parents[1]


class FakeCache:
    def __init__(self, serialized: str = '{"cache": true}'):
        self.serialized = serialized
        self.loaded: str | None = None

    def deserialize(self, value: str) -> None:
        self.loaded = value

    def serialize(self) -> str:
        return self.serialized


class FakeGraphResponse:
    def __init__(self, status_code: int, payload: dict[str, Any]):
        self.status_code = status_code
        self.payload = payload

    def json(self) -> dict[str, Any]:
        return self.payload


class FakeSession:
    def __init__(self, response: FakeGraphResponse):
        self.response = response
        self.headers: dict[str, str] | None = None

    def get(self, url: str, **kwargs: Any) -> FakeGraphResponse:
        self.headers = kwargs["headers"]
        return self.response


def test_authorize_graph_device_flow_writes_cache_and_verifies_me(tmp_path) -> None:
    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://www.microsoft.com/link",
                "expires_at": 9999999999,
                "message": "Use ABCD-EFGH",
            }

        def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "secret-access-token"}

    log_file = tmp_path / "auth.log"
    logger = setup_logging(str(log_file), debug=True)
    session = FakeSession(FakeGraphResponse(200, {"userPrincipalName": "person@outlook.com"}))
    config = _config(tmp_path)

    result = authorize_graph(
        config,
        logger,
        app_factory=FakeApp,
        cache_factory=lambda: FakeCache(),
        session=session,
    )

    assert result.account == "person@outlook.com"
    assert config.cache_path.read_text(encoding="utf-8") == '{"cache": true}'
    assert session.headers and session.headers["Authorization"] == "Bearer secret-access-token"
    log_content = log_file.read_text(encoding="utf-8")
    assert "secret-access-token" not in log_content
    assert "Authorization" not in log_content
    assert "ABCD-EFGH" in log_content


def test_authorize_graph_uses_existing_silent_token(tmp_path) -> None:
    cache_path = tmp_path / "cache.json"
    cache_path.write_text('{"existing": true}', encoding="utf-8")

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"username": username}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "silent-token"}

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            raise AssertionError("device flow should not start")

    logger = setup_logging(str(tmp_path / "silent.log"), debug=True)
    session = FakeSession(FakeGraphResponse(200, {"mail": "person@outlook.com"}))
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(cache_path))

    result = authorize_graph(
        config,
        logger,
        app_factory=FakeApp,
        cache_factory=lambda: FakeCache(serialized='{"existing": true}'),
        session=session,
    )

    assert result.account == "person@outlook.com"
    assert cache_path.read_text(encoding="utf-8") == '{"existing": true}'


def test_authorize_graph_mismatched_single_cached_account_starts_device_flow(tmp_path) -> None:
    cache_path = tmp_path / "cache.json"
    cache_path.write_text('{"existing": true}', encoding="utf-8")

    class FakeApp:
        instance: "FakeApp"

        def __init__(self, *args: Any, **kwargs: Any):
            self.wrong_account = {"username": "wrong@example.com"}
            self.removed: list[dict[str, str]] = []
            self.device_flow_started = False
            FakeApp.instance = self

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [] if username is not None else [self.wrong_account]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, str]:
            return {"access_token": "wrong-token"}

        def remove_account(self, account: dict[str, Any]) -> None:
            self.removed.append(account)

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            self.device_flow_started = True
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://www.microsoft.com/link",
                "expires_at": 9999999999,
            }

        def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, str]:
            return {"access_token": "correct-token"}

    class SequenceSession:
        def __init__(self):
            self.responses = [
                FakeGraphResponse(200, {"mail": "wrong@example.com"}),
                FakeGraphResponse(200, {"mail": "person@outlook.com"}),
            ]

        def get(self, url: str, **kwargs: Any) -> FakeGraphResponse:
            return self.responses.pop(0)

    logger = setup_logging(str(tmp_path / "mismatch.log"), debug=True)
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(cache_path))

    result = authorize_graph(
        config,
        logger,
        app_factory=FakeApp,
        cache_factory=lambda: FakeCache(serialized='{"updated": true}'),
        session=SequenceSession(),
    )

    assert result.account == "person@outlook.com"
    assert FakeApp.instance.removed == [FakeApp.instance.wrong_account]
    assert FakeApp.instance.device_flow_started is True
    assert cache_path.read_text(encoding="utf-8") == '{"updated": true}'


def test_authorize_graph_reports_access_denied(tmp_path) -> None:
    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {"user_code": "ABCD", "verification_uri": "https://www.microsoft.com/link"}

        def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, Any]:
            return {"error": "access_denied", "error_description": "User canceled authorization"}

    logger = setup_logging(str(tmp_path / "denied.log"), debug=True)

    with pytest.raises(AuthSetupError, match="User canceled authorization"):
        authorize_graph(
            _config(tmp_path),
            logger,
            app_factory=FakeApp,
            cache_factory=lambda: FakeCache(),
            session=FakeSession(FakeGraphResponse(200, {})),
        )


def test_authorize_graph_reports_graph_validation_failure(tmp_path) -> None:
    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            pass

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return []

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            return {"user_code": "ABCD", "verification_uri": "https://www.microsoft.com/link"}

        def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "secret-access-token"}

    logger = setup_logging(str(tmp_path / "graph-error.log"), debug=True)

    with pytest.raises(AuthSetupError, match="Graph /me validation failed"):
        authorize_graph(
            _config(tmp_path),
            logger,
            app_factory=FakeApp,
            cache_factory=lambda: FakeCache(),
            session=FakeSession(FakeGraphResponse(403, {"error": {"code": "Forbidden", "message": "No permission"}})),
        )


def test_bat_file_points_to_auth_setup() -> None:
    content = (REPO_ROOT / "authorize_outlook.bat").read_text(encoding="utf-8")

    assert "python -m outlook_mcp.auth_setup" in content
    assert "logs\\outlook-auth-" in content
    assert "OUTLOOK_CONFIG_FILE" in content
    assert "chcp 65001" in content
    assert "PYTHONUTF8=1" in content
    assert "PYTHONIOENCODING=utf-8" in content


def test_mcp_json_sets_utf8_env() -> None:
    content = (REPO_ROOT / ".mcp.json").read_text(encoding="utf-8")

    assert '"PYTHONUTF8": "1"' in content
    assert '"PYTHONIOENCODING": "utf-8"' in content


def _config(tmp_path: Path) -> GraphConfig:
    return GraphConfig(
        email="person@outlook.com",
        client_id="9966331f-d9e2-42b2-85bf-5a54280ee60e",
        token_cache_path=str(tmp_path / "msal-token-cache.json"),
    )


def test_authorize_graph_passes_timeout_http_client(tmp_path) -> None:
    captured: dict[str, Any] = {}

    class FakeApp:
        def __init__(self, *args: Any, **kwargs: Any):
            captured.update(kwargs)

        def get_accounts(self, username: str | None = None) -> list[Any]:
            return [{"username": username}]

        def acquire_token_silent(self, scopes: list[str], account: dict[str, Any]) -> dict[str, Any]:
            return {"access_token": "silent-token"}

        def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
            raise AssertionError("device flow should not start")

    logger = setup_logging(str(tmp_path / "auth.log"))
    session = FakeSession(FakeGraphResponse(200, {"mail": "person@outlook.com"}))
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(tmp_path / "cache.json"))

    result = authorize_graph(
        config,
        logger,
        app_factory=FakeApp,
        cache_factory=lambda: FakeCache(),
        session=session,
    )

    from outlook_mcp.graph_auth import TimeoutSession

    assert result.account == "person@outlook.com"
    assert isinstance(captured.get("http_client"), TimeoutSession)


def test_verify_graph_token_closes_owned_session(monkeypatch, tmp_path) -> None:
    from outlook_mcp.auth_setup import verify_graph_token

    closed = {"value": False}

    class ClosingSession(FakeSession):
        def close(self) -> None:
            closed["value"] = True

    session = ClosingSession(FakeGraphResponse(200, {"mail": "person@outlook.com"}))
    monkeypatch.setattr("outlook_mcp.auth_setup.requests.Session", lambda: session)
    config = GraphConfig(email="person@outlook.com", client_id="client-id", token_cache_path=str(tmp_path / "cache.json"))
    logger = setup_logging(str(tmp_path / "verify.log"))

    account, _ = verify_graph_token(config, "token", logger)

    assert account == "person@outlook.com"
    assert closed["value"] is True


def test_save_cache_uses_atomic_replace(monkeypatch, tmp_path) -> None:
    import os

    import outlook_mcp.auth_setup as auth_setup

    replaced: list[tuple[Any, Any]] = []
    real_replace = os.replace

    def spy_replace(src: Any, dst: Any) -> None:
        replaced.append((src, dst))
        real_replace(src, dst)

    monkeypatch.setattr(auth_setup.os, "replace", spy_replace)
    cache_path = tmp_path / "cache.json"
    logger = setup_logging(str(tmp_path / "save.log"))

    auth_setup._save_cache(FakeCache(), cache_path, logger)

    assert replaced and replaced[0][1] == cache_path
    assert cache_path.read_text(encoding="utf-8") == '{"cache": true}'
    assert not list(tmp_path.glob("*.tmp"))


def test_http_bat_skip_prompt_requires_value_1() -> None:
    content = (REPO_ROOT / "start_outlook_mcp_http.bat").read_text(encoding="utf-8")
    assert 'if not "%OUTLOOK_MCP_SKIP_PROMPT%"=="1"' in content


def test_redact_text_masks_json_secret_values() -> None:
    from outlook_mcp.auth_setup import redact_text

    redacted = redact_text('{"access_token": "opaque-token-value", "client_secret": "sekret-12345"}')
    assert "opaque-token-value" not in redacted
    assert "sekret-12345" not in redacted
