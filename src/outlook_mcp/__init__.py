"""Outlook MCP server package."""

import sys

__version__ = "0.1.48"


def _configure_utf8_stdio() -> None:
    for stream_name in ("stdout", "stderr"):
        try:
            stream = getattr(sys, stream_name)
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


_configure_utf8_stdio()
