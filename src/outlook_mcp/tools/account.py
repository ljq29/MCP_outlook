"""Account tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.graph_client import GraphClient

from .common import run_safely


def account_info(client: GraphClient) -> dict[str, Any]:
    """Return the current account email and display name."""

    def action() -> dict[str, Any]:
        return client.account_info()

    return run_safely(action)
