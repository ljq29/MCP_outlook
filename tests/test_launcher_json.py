from __future__ import annotations

import json

import pytest

from outlook_mcp.launcher_json import build_endpoint_path, build_mcp_json, validate_secret

TEST_SECRET = "test-secret-at-least-24-chars"


def test_launcher_json_requires_explicit_bearer_token() -> None:
    with pytest.raises(ValueError, match="bearer token is required"):
        build_mcp_json()


def test_launcher_json_uses_explicit_debug_endpoint() -> None:
    payload = json.loads(build_mcp_json(bearer_token=TEST_SECRET))

    assert payload == {
        "mcpServers": {
            "outlook-http": {
                "type": "streamable-http",
                "url": "http://127.0.0.1:8000/graph/you@example.com",
                "headers": {
                    "Authorization": f"Bearer {TEST_SECRET}",
                },
            }
        }
    }


def test_launcher_json_allows_host_and_port_override() -> None:
    payload = json.loads(build_mcp_json(host="0.0.0.0", port=9000, bearer_token=TEST_SECRET))

    server = payload["mcpServers"]["outlook-http"]
    assert server["url"] == "http://0.0.0.0:9000/graph/you@example.com"
    assert server["headers"]["Authorization"] == f"Bearer {TEST_SECRET}"


def test_launcher_json_includes_custom_tools_header() -> None:
    payload = json.loads(build_mcp_json(tools="calendar_list,account_info", bearer_token=TEST_SECRET))

    assert payload["mcpServers"]["outlook-http"]["headers"] == {
        "Authorization": f"Bearer {TEST_SECRET}",
        "X-Outlook-MCP-Tools": "calendar_list,account_info",
    }


def test_launcher_json_includes_mode_header() -> None:
    payload = json.loads(build_mcp_json(mode="calendar", bearer_token=TEST_SECRET))

    assert payload["mcpServers"]["outlook-http"]["headers"] == {
        "Authorization": f"Bearer {TEST_SECRET}",
        "X-Outlook-MCP-Mode": "calendar",
    }


def test_launcher_json_accepts_bearer_token_alias() -> None:
    payload = json.loads(build_mcp_json(bearer_token="test-bearer-token-at-least-24-chars"))

    assert payload["mcpServers"]["outlook-http"]["headers"]["Authorization"] == "Bearer test-bearer-token-at-least-24-chars"


def test_launcher_json_empty_custom_mode_disables_all_tools() -> None:
    payload = json.loads(build_mcp_json(mode="custom", bearer_token=TEST_SECRET))

    assert payload["mcpServers"]["outlook-http"]["headers"] == {
        "Authorization": f"Bearer {TEST_SECRET}",
        "X-Outlook-MCP-Mode": "custom",
    }


def test_endpoint_path_preserves_account_at_sign_and_safe_secret() -> None:
    assert build_endpoint_path("graph", "person@example.com") == "/graph/person@example.com"


@pytest.mark.parametrize("secret", ["a/b", "a b", "a?b", "a%b", "中文", ""])
def test_validate_secret_rejects_url_escaped_characters(secret: str) -> None:
    with pytest.raises(ValueError, match="secret may only contain"):
        validate_secret(secret)


def test_validate_secret_uses_sixteen_character_default() -> None:
    validate_secret("a" * 16)

    with pytest.raises(ValueError, match="at least 16 characters"):
        validate_secret("a" * 15)


def test_validate_secret_uses_configured_minimum(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SECRET_MIN_LENGTH", "8")

    validate_secret("abcdefgh")
    with pytest.raises(ValueError, match="at least 8 characters"):
        validate_secret("abcdefg")


@pytest.mark.parametrize(
    ("value", "message"),
    [
        ("invalid", "must be an integer"),
        ("0", "must be greater than 0"),
        ("-1", "must be greater than 0"),
    ],
)
def test_validate_secret_rejects_invalid_configured_minimum(
    monkeypatch,
    value: str,
    message: str,
) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SECRET_MIN_LENGTH", value)

    with pytest.raises(ValueError, match=message):
        validate_secret("valid-secret")


def test_validate_secret_explicit_minimum_ignores_environment(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SECRET_MIN_LENGTH", "8")

    validate_secret("x", minimum_length=1)
