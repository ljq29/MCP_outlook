from __future__ import annotations

import pytest

from outlook_mcp.config import ConfigError, GraphConfig, load_config, load_config_file


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


def test_ews_provider_is_rejected_with_migration_hint(tmp_path) -> None:
    config_path = tmp_path / "ews.json"
    config_path.write_text(
        '{"provider":"ews","email":"user@example.com","ews":{"auth_mode":"password"}}',
        encoding="utf-8",
    )

    with pytest.raises(ConfigError, match="EWS support was removed"):
        load_config_file(config_path)


def test_unknown_provider_is_rejected(tmp_path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text('{"provider":"caldav","email":"user@example.com"}', encoding="utf-8")

    with pytest.raises(ConfigError, match="provider must be 'graph'"):
        load_config_file(config_path)


# --- config discovery diagnostics and email format ---


def test_load_config_without_file_reports_searched_paths(tmp_path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)

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
