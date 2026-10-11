"""Account tools for Outlook MCP."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client import OutlookClient
from outlook_mcp.graph_client import GraphClient

from .common import run_safely


def account_info(client: OutlookClient) -> dict[str, Any]:
    """Return the current account email and display name."""

    def action() -> dict[str, Any]:
        if isinstance(client, GraphClient):
            return client.account_info()
        account = client.account
        identity = getattr(account, "identity", None)
        return {
            "primary_smtp_address": (
                getattr(identity, "primary_smtp_address", None)
                or getattr(account, "primary_smtp_address", "")
            ),
            "display_name": (
                getattr(account, "fullname", None)
                or getattr(identity, "display_name", "")
            ),
        }

    return run_safely(action)
