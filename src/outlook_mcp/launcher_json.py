"""Generate MCP client JSON for the HTTP Outlook MCP launcher."""

from __future__ import annotations

import argparse
import json
import os
import re
from typing import Any
from urllib.parse import quote


DEFAULT_METHOD = "graph"
DEFAULT_ACCOUNT = "you@example.com"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8000
DEFAULT_SERVER_NAME = "outlook-http"
TOOL_MODES = {"full", "mail", "calendar", "contacts", "tasks", "custom"}
SECRET_ALLOWED_PATTERN = re.compile(r"^[A-Za-z0-9._~-]+$")
INVALID_SECRET_MESSAGE = "secret may only contain A-Z, a-z, 0-9, dot, underscore, hyphen, and tilde"
DEFAULT_SECRET_MIN_LENGTH = 16
SECRET_MIN_LENGTH_ENV = "OUTLOOK_MCP_SECRET_MIN_LENGTH"


def build_mcp_config(
    *,
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    method: str = DEFAULT_METHOD,
    account: str = DEFAULT_ACCOUNT,
    secret: str | None = None,
    bearer_token: str | None = None,
    mode: str = "full",
    tools: str | None = None,
    scheme: str = "http",
    server_name: str = DEFAULT_SERVER_NAME,
) -> dict[str, Any]:
    token = bearer_token if bearer_token is not None else secret
    if token is None:
        raise ValueError("bearer token is required")
    validate_secret(token)
    normalized_mode = mode.strip().lower() or "full"
    validate_mode(normalized_mode)
    headers = {"Authorization": f"Bearer {token}"}
    if tools:
        headers["X-Outlook-MCP-Tools"] = tools
    elif normalized_mode != "full":
        headers["X-Outlook-MCP-Mode"] = normalized_mode
    return {
        "mcpServers": {
            server_name: {
                "type": "streamable-http",
                "url": f"{scheme}://{host}:{port}{build_endpoint_path(method, account)}",
                "headers": headers,
            }
        }
    }


def build_mcp_json(**kwargs: Any) -> str:
    return json.dumps(build_mcp_config(**kwargs), ensure_ascii=False, indent=2)


def build_endpoint_path(method: str, account: str) -> str:
    return "/" + "/".join(
        (
            _path_segment(method),
            _account_segment(account),
        )
    )


def _configured_secret_min_length() -> int:
    raw = os.environ.get(SECRET_MIN_LENGTH_ENV)
    if raw is None:
        return DEFAULT_SECRET_MIN_LENGTH
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{SECRET_MIN_LENGTH_ENV} must be an integer") from exc
    if value < 1:
        raise ValueError(f"{SECRET_MIN_LENGTH_ENV} must be greater than 0")
    return value


def validate_secret(secret: str, *, minimum_length: int | None = None) -> None:
    if not secret or not SECRET_ALLOWED_PATTERN.fullmatch(secret):
        raise ValueError(INVALID_SECRET_MESSAGE)
    required_length = (
        minimum_length if minimum_length is not None else _configured_secret_min_length()
    )
    if len(secret) < required_length:
        raise ValueError(f"secret must be at least {required_length} characters")


def validate_mode(mode: str) -> None:
    if mode not in TOOL_MODES:
        raise ValueError(f"Unknown MCP tools mode: {mode}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Print HTTP MCP JSON for Outlook MCP.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--method", default=DEFAULT_METHOD)
    parser.add_argument("--account", default=DEFAULT_ACCOUNT)
    parser.add_argument("--secret", default=None, help="Deprecated alias for --bearer-token.")
    parser.add_argument("--bearer-token", default=None)
    parser.add_argument("--mode", default="full", choices=sorted(TOOL_MODES))
    parser.add_argument("--tools", default=None, help="Comma-separated X-Outlook-MCP-Tools value.")
    parser.add_argument("--scheme", default="http")
    parser.add_argument("--server-name", default=DEFAULT_SERVER_NAME)
    args = parser.parse_args(argv)
    token = args.bearer_token or args.secret
    if token is None:
        parser.error("--bearer-token is required")
    print(
        build_mcp_json(
            host=args.host,
            port=args.port,
            method=args.method,
            account=args.account,
            secret=args.secret,
            bearer_token=args.bearer_token,
            mode=args.mode,
            tools=args.tools,
            scheme=args.scheme,
            server_name=args.server_name,
        )
    )
    return 0


def _path_segment(value: str) -> str:
    return quote(str(value), safe="")


def _account_segment(value: str) -> str:
    return quote(str(value), safe="@")


if __name__ == "__main__":
    raise SystemExit(main())
