#!/usr/bin/env sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$SCRIPT_DIR"

PORT="${1:-${OUTLOOK_MCP_HTTP_PORT:-8000}}"
BIND_HOST="${2:-${OUTLOOK_MCP_HTTP_HOST:-127.0.0.1}}"

OUTLOOK_CONFIG_FILE="${OUTLOOK_CONFIG_FILE:-$SCRIPT_DIR/outlook-mcp.json}"
OUTLOOK_MCP_DB="${OUTLOOK_MCP_DB:-$HOME/.outlook-mcp/outlook-mcp.sqlite3}"
OUTLOOK_MCP_METHOD="${OUTLOOK_MCP_METHOD:-graph}"
OUTLOOK_MCP_ACCOUNT="${OUTLOOK_MCP_ACCOUNT:-}"
OUTLOOK_MCP_BEARER_TOKEN="${OUTLOOK_MCP_BEARER_TOKEN:-}"
OUTLOOK_MCP_MODE="${OUTLOOK_MCP_MODE:-full}"
OUTLOOK_MCP_TOOLS="${OUTLOOK_MCP_TOOLS:-}"
OUTLOOK_MCP_SCHEME="${OUTLOOK_MCP_SCHEME:-http}"
OUTLOOK_MCP_SERVER_NAME="${OUTLOOK_MCP_SERVER_NAME:-outlook-http}"

export OUTLOOK_CONFIG_FILE
export OUTLOOK_MCP_DB
export OUTLOOK_MCP_HTTP_PORT="$PORT"
export OUTLOOK_MCP_HTTP_HOST="$BIND_HOST"
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8

if [ -z "$OUTLOOK_MCP_ACCOUNT" ]; then
  echo "OUTLOOK_MCP_ACCOUNT is required (for example: you@outlook.com)." >&2
  exit 2
fi
if [ -z "$OUTLOOK_MCP_BEARER_TOKEN" ]; then
  echo "OUTLOOK_MCP_BEARER_TOKEN is required." >&2
  exit 2
fi

echo "Note: the first uv run will automatically install dependencies (uv sync)."
echo "Registering default debug credential: $OUTLOOK_MCP_METHOD/$OUTLOOK_MCP_ACCOUNT"
uv run python -m outlook_mcp.credential_store register "$OUTLOOK_MCP_METHOD" "$OUTLOOK_MCP_ACCOUNT" "$OUTLOOK_MCP_BEARER_TOKEN"

echo
echo "Full MCP JSON:"
if [ -n "$OUTLOOK_MCP_TOOLS" ]; then
  uv run python -m outlook_mcp.launcher_json \
    --host "$BIND_HOST" \
    --port "$PORT" \
    --method "$OUTLOOK_MCP_METHOD" \
    --account "$OUTLOOK_MCP_ACCOUNT" \
    --bearer-token "$OUTLOOK_MCP_BEARER_TOKEN" \
    --mode "$OUTLOOK_MCP_MODE" \
    --tools "$OUTLOOK_MCP_TOOLS" \
    --scheme "$OUTLOOK_MCP_SCHEME" \
    --server-name "$OUTLOOK_MCP_SERVER_NAME"
else
  uv run python -m outlook_mcp.launcher_json \
    --host "$BIND_HOST" \
    --port "$PORT" \
    --method "$OUTLOOK_MCP_METHOD" \
    --account "$OUTLOOK_MCP_ACCOUNT" \
    --bearer-token "$OUTLOOK_MCP_BEARER_TOKEN" \
    --mode "$OUTLOOK_MCP_MODE" \
    --scheme "$OUTLOOK_MCP_SCHEME" \
    --server-name "$OUTLOOK_MCP_SERVER_NAME"
fi

echo
echo "Service base URL: $OUTLOOK_MCP_SCHEME://$BIND_HOST:$PORT"
echo "MCP URL:          $OUTLOOK_MCP_SCHEME://$BIND_HOST:$PORT/$OUTLOOK_MCP_METHOD/$OUTLOOK_MCP_ACCOUNT"
echo "Auth/debug URL:   $OUTLOOK_MCP_SCHEME://$BIND_HOST:$PORT/$OUTLOOK_MCP_METHOD/$OUTLOOK_MCP_ACCOUNT/$OUTLOOK_MCP_BEARER_TOKEN"
echo

if [ "${OUTLOOK_MCP_SKIP_PROMPT:-0}" != "1" ]; then
  printf "Copy the MCP JSON above, then press Enter to start Outlook MCP HTTP..."
  read _start || true
fi

exec uv run outlook-mcp-http
