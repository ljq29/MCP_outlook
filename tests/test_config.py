from __future__ import annotations

import pytest

from outlook_mcp.config import ConfigError, GraphConfig, OutlookConfig, load_config, load_config_file


def test_oauth2_config_loads_required_values() -> None:
    config = OutlookConfig.from_env(
        {
            "OUTLOOK_EMAIL": "user@example.com",
            "OUTLOOK_AUTH_MODE": "oauth2",
            "OUTLOOK_CLIENT_ID": "client-id",
            "OUTLOOK_CLIENT_SECRET": "secret",
            "OUTLOOK_TENANT_ID": "tenant",
        }
    )

    assert config.email == "user@example.com"
    assert config.auth_mode == "oauth2"
    assert config.server == "outlook.office365.com"
    assert config.redacted["client_secret"] == "[redacted]"


def test_password_config_loads_required_values() -> None:
    config = OutlookConfig.from_env(
        {
            "OUTLOOK_EMAIL": "user@example.com",
            "OUTLOOK_AUTH_MODE": "password",
            "OUTLOOK_USERNAME": "DOMAIN\\user",
            "OUTLOOK_PASSWORD": "secret",
        }
    )

    assert config.username == "DOMAIN\\user"
    assert config.password == "secret"
    assert config.redacted["password"] == "[redacted]"


def test_missing_auth_mode_is_rejected() -> None:
    with pytest.raises(ConfigError, match="OUTLOOK_AUTH_MODE"):
        OutlookConfig.from_env({"OUTLOOK_EMAIL": "user@example.com"})


def test_invalid_access_type_is_rejected() -> None:
    with pytest.raises(ConfigError, match="OUTLOOK_ACCESS_TYPE"):
        OutlookConfig.from_env(
            {
                "OUTLOOK_EMAIL": "user@example.com",
                "OUTLOOK_AUTH_MODE": "password",
                "OUTLOOK_USERNAME": "user",
                "OUTLOOK_PASSWORD": "secret",
                "OUTLOOK_ACCESS_TYPE": "owner",
            }
        )


def test_graph_json_config_loads(tmp_path) -> None:
    config_path = tmp_path / "outlook-mcp.json"
    config_path.write_text(
        """
        {
          "provider": "graph",
          "email": "person@outlook.com",
          "graph": {
            "client_id": "client-id",
            "tenant": "consumers",
            "token_cache_path": "~/.outlook-mcp/test-cache.json",
            "scopes": ["User.Read", "Mail.ReadWrite"]
          }
        }
        """,
        encoding="utf-8",
    )

    config = load_config_file(config_path)

    assert isinstance(config, GraphConfig)
    assert config.email == "person@outlook.com"
    assert config.client_id == "client-id"
    assert config.scopes == ("User.Read", "Mail.ReadWrite")


def test_load_config_uses_outlook_config_file_env(tmp_path) -> None:
    config_path = tmp_path / "graph.json"
    config_path.write_text(
        '{"provider":"graph","email":"person@outlook.com","graph":{"client_id":"client-id"}}',
        encoding="utf-8",
    )

    config = load_config({"OUTLOOK_CONFIG_FILE": str(config_path)})

    assert isinstance(config, GraphConfig)
    assert config.tenant == "consumers"


def test_graph_client_id_can_come_from_env(tmp_path) -> None:
    config_path = tmp_path / "graph.json"
    config_path.write_text(
        '{"provider":"graph","email":"person@outlook.com","graph":{"client_id":"file-client-id"}}',
        encoding="utf-8",
    )

    config = load_config(
        {
            "OUTLOOK_CONFIG_FILE": str(config_path),
            "OUTLOOK_GRAPH_CLIENT_ID": "env-client-id",
        }
    )

    assert isinstance(config, GraphConfig)
    assert config.client_id == "env-client-id"


def test_graph_client_id_falls_back_to_config_file(tmp_path) -> None:
    config_path = tmp_path / "graph.json"
    config_path.write_text(
        '{"provider":"graph","email":"person@outlook.com","graph":{"client_id":"file-client-id"}}',
        encoding="utf-8",
    )

    config = load_config({"OUTLOOK_CONFIG_FILE": str(config_path)})

    assert isinstance(config, GraphConfig)
    assert config.client_id == "file-client-id"


def test_graph_client_id_is_required_when_env_and_file_are_missing(tmp_path) -> None:
    config_path = tmp_path / "graph.json"
    config_path.write_text(
        '{"provider":"graph","email":"person@outlook.com","graph":{}}',
        encoding="utf-8",
    )

    with pytest.raises(ConfigError, match="graph.client_id"):
        load_config({"OUTLOOK_CONFIG_FILE": str(config_path)})


def test_ews_json_config_loads(tmp_path) -> None:
    config_path = tmp_path / "ews.json"
    config_path.write_text(
        """
        {
          "provider": "ews",
          "email": "user@example.com",
          "ews": {
            "auth_mode": "password",
            "username": "user@example.com",
            "password": "secret"
          }
        }
        """,
        encoding="utf-8",
    )

    config = load_config_file(config_path)

    assert isinstance(config, OutlookConfig)
    assert config.auth_mode == "password"


# --- config discovery diagnostics and email format ---


def test_load_config_without_file_reports_searched_paths(tmp_path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)

    with pytest.raises(ConfigError, match="Missing required environment variable: OUTLOOK_EMAIL"):
        load_config({"UNRELATED": "1"})

    with pytest.raises(ConfigError) as exc:
        load_config({"UNRELATED": "1"})
    message = str(exc.value)
    assert "Searched for outlook-mcp.json in:" in message
    assert str(tmp_path) in message
    assert "OUTLOOK_CONFIG_FILE" in message


def test_graph_config_rejects_invalid_email() -> None:
    with pytest.raises(ConfigError, match="invalid email"):
        GraphConfig(email="not-an-email", client_id="client-id").validate()


def test_graph_config_accepts_valid_email() -> None:
    GraphConfig(email="person@outlook.com", client_id="client-id").validate()


def test_ews_config_rejects_invalid_email() -> None:
    with pytest.raises(ConfigError, match="invalid email"):
        OutlookConfig.from_env(
            {
                "OUTLOOK_EMAIL": "not-an-email",
                "OUTLOOK_AUTH_MODE": "password",
                "OUTLOOK_USERNAME": "user",
                "OUTLOOK_PASSWORD": "secret",
            }
        )


def test_redact_secrets_covers_config_loaded_password() -> None:
    from outlook_mcp.tools.common import redact_secrets

    OutlookConfig.from_env(
        {
            "OUTLOOK_EMAIL": "user@example.com",
            "OUTLOOK_AUTH_MODE": "password",
            "OUTLOOK_USERNAME": "user",
            "OUTLOOK_PASSWORD": "CfgSecret-987654",
        }
    )
    redacted = redact_secrets("EWS auth failed for CfgSecret-987654")
    assert "CfgSecret-987654" not in redacted
    assert "[redacted]" in redacted
