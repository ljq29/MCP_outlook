"""Standalone Microsoft Graph authorization helper."""

from __future__ import annotations

import argparse
import logging
import os
import platform
import re
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import msal
import requests

from .config import ConfigError, GraphConfig, load_config
from .graph_auth import GRAPH_HTTP_TIMEOUT_SECONDS, TimeoutSession, graph_account_identities, same_account
from .tools.common import graph_error_message as _graph_error_message, msal_error_message as _msal_error_message


class AuthSetupError(RuntimeError):
    """Raised when the standalone authorization flow cannot complete."""


class AuthAccountMismatchError(AuthSetupError):
    """Raised when a valid Graph token belongs to a different mailbox."""


@dataclass(frozen=True)
class AuthSetupResult:
    account: str | None
    cache_path: Path
    graph_status_code: int


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Authorize Outlook MCP Microsoft Graph access.")
    parser.add_argument("--config-file", default=os.environ.get("OUTLOOK_CONFIG_FILE"))
    parser.add_argument("--log-file")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=900)
    args = parser.parse_args(argv)

    logger = setup_logging(args.log_file, debug=args.debug)
    try:
        logger.info("Outlook MCP Graph authorization started")
        logger.debug("Python: %s", sys.version.replace("\n", " "))
        logger.debug("Platform: %s", platform.platform())
        config = load_config(config_path=args.config_file)
        if not isinstance(config, GraphConfig):
            raise AuthSetupError("authorize_outlook.bat only supports provider='graph'")
        result = authorize_graph(config, logger, timeout_seconds=args.timeout_seconds)
        logger.info("Authorization succeeded for account: %s", result.account or config.email)
        logger.info("Token cache written: %s", result.cache_path)
        logger.info("Graph /me status: %s", result.graph_status_code)
        return 0
    except (AuthSetupError, ConfigError) as exc:
        logger.error("Authorization failed: %s", redact_text(str(exc)))
        return 1
    except KeyboardInterrupt:
        logger.error("Authorization interrupted by user")
        return 130
    except Exception as exc:  # pragma: no cover - defensive final barrier
        logger.exception("Unexpected authorization failure: %s", redact_text(str(exc)))
        return 1


def setup_logging(log_file: str | None, *, debug: bool = False) -> logging.Logger:
    logger = logging.getLogger("outlook_mcp.auth_setup")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG if debug else logging.INFO)
    console.setFormatter(formatter)
    console.addFilter(RedactingFilter())
    logger.addHandler(console)

    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        file_handler.addFilter(RedactingFilter())
        logger.addHandler(file_handler)
    return logger


def authorize_graph(
    config: GraphConfig,
    logger: logging.Logger,
    *,
    timeout_seconds: int = 900,
    app_factory: Callable[..., Any] = msal.PublicClientApplication,
    cache_factory: Callable[[], Any] = msal.SerializableTokenCache,
    session: Any | None = None,
) -> AuthSetupResult:
    if timeout_seconds < 30:
        raise AuthSetupError("timeout_seconds must be at least 30")

    cache_path = config.cache_path
    pending_path = _pending_path(cache_path)
    pending_path.unlink(missing_ok=True)
    cache = cache_factory()
    if cache_path.exists():
        logger.info("Existing token cache found: %s (%s bytes)", cache_path, cache_path.stat().st_size)
        cache.deserialize(cache_path.read_text(encoding="utf-8"))
    else:
        logger.info("No existing token cache found: %s", cache_path)

    logger.info("Config file provider: graph")
    logger.info("Email: %s", config.email)
    logger.info("Tenant: %s", config.tenant)
    logger.info("Client ID: %s", _mask_client_id(config.client_id))
    logger.info("Scopes: %s", ", ".join(config.scopes))
    logger.info("Token cache path: %s", cache_path)

    app = app_factory(config.client_id, authority=config.authority, token_cache=cache, http_client=TimeoutSession())
    silent_result = _acquire_silent_token(app, config, logger)
    if silent_result:
        token, cached_account = silent_result
        logger.info("Silent token acquisition succeeded")
        try:
            account, status_code = verify_graph_token(config, token, logger, session=session)
        except AuthAccountMismatchError as exc:
            logger.warning("Cached account does not match this configuration: %s", exc)
            if hasattr(app, "remove_account"):
                app.remove_account(cached_account)
            token = _interactive_device_flow(app, config, logger, timeout_seconds=timeout_seconds)
            account, status_code = verify_graph_token(config, token, logger, session=session)
    else:
        token = _interactive_device_flow(app, config, logger, timeout_seconds=timeout_seconds)
        account, status_code = verify_graph_token(config, token, logger, session=session)
    _save_cache(cache, cache_path, logger)
    pending_path.unlink(missing_ok=True)
    return AuthSetupResult(account=account, cache_path=cache_path, graph_status_code=status_code)


def verify_graph_token(
    config: GraphConfig,
    token: str,
    logger: logging.Logger,
    *,
    session: Any | None = None,
) -> tuple[str | None, int]:
    owns_session = session is None
    http = session or requests.Session()
    try:
        return _verify_graph_token_with(config, token, logger, http)
    finally:
        if owns_session:
            http.close()


def _verify_graph_token_with(
    config: GraphConfig,
    token: str,
    logger: logging.Logger,
    http: Any,
) -> tuple[str | None, int]:
    url = f"{config.graph_base_url}/me"
    logger.info("Verifying token with Graph /me")
    response = http.get(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        timeout=GRAPH_HTTP_TIMEOUT_SECONDS,
    )
    status_code = response.status_code
    try:
        payload = response.json()
    except ValueError:
        payload = {}
    logger.debug("Graph /me response status: %s", status_code)
    if not 200 <= status_code < 300:
        raise AuthSetupError(f"Graph /me validation failed: {_graph_error_message(status_code, payload)}")
    account = next(
        (identity for identity in graph_account_identities(payload) if same_account(identity, config.email)),
        None,
    )
    if account is None:
        actual = ", ".join(graph_account_identities(payload)) or "unknown"
        raise AuthAccountMismatchError(
            f"Authorized Microsoft account {actual} does not match configured Outlook account {config.email}"
        )
    return account, status_code


def _acquire_silent_token(
    app: Any,
    config: GraphConfig,
    logger: logging.Logger,
) -> tuple[str, dict[str, Any]] | None:
    all_accounts = app.get_accounts()
    accounts = app.get_accounts(username=config.email)
    if not accounts and len(all_accounts) == 1:
        accounts = all_accounts
    logger.debug("MSAL accounts available for silent auth: %s", len(accounts))
    for account in accounts:
        result = app.acquire_token_silent(list(config.scopes), account=account)
        if result and "access_token" in result:
            return result["access_token"], account
        if result and result.get("error"):
            logger.debug("Silent auth skipped account due to: %s", result.get("error"))
    return None


def _interactive_device_flow(app: Any, config: GraphConfig, logger: logging.Logger, *, timeout_seconds: int) -> str:
    flow = app.initiate_device_flow(scopes=list(config.scopes))
    if "user_code" not in flow:
        raise AuthSetupError(_msal_error_message(flow))
    expires_at = flow.get("expires_at")
    if isinstance(expires_at, int | float):
        flow["expires_at"] = min(expires_at, time.time() + timeout_seconds)
    else:
        flow["expires_at"] = time.time() + timeout_seconds

    verification_uri = flow.get("verification_uri") or flow.get("verification_url")
    user_code = flow.get("user_code")
    logger.info("Open this URL in a browser: %s", verification_uri)
    logger.info("Enter this code: %s", user_code)
    logger.info("Waiting up to %s seconds for browser authorization", timeout_seconds)
    message = flow.get("message")
    if message:
        logger.info("%s", message)

    result = app.acquire_token_by_device_flow(flow)
    if result and "access_token" in result:
        return result["access_token"]
    raise AuthSetupError(_msal_error_message(result))


def _save_cache(cache: Any, cache_path: Path, logger: logging.Logger) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    serialized = cache.serialize()
    if not serialized:
        raise AuthSetupError("MSAL token cache was empty after authorization")
    fd, tmp_name = tempfile.mkstemp(dir=cache_path.parent, prefix=cache_path.name + ".", suffix=".tmp")
    tmp_path = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(serialized)
        try:
            os.chmod(tmp_path, 0o600)
        except OSError:
            pass
        os.replace(tmp_path, cache_path)
    finally:
        tmp_path.unlink(missing_ok=True)
    logger.info("Token cache saved: %s (%s bytes)", cache_path, cache_path.stat().st_size)


def _pending_path(cache_path: Path) -> Path:
    return cache_path.with_suffix(cache_path.suffix + ".pending.json")


def _mask_client_id(client_id: str) -> str:
    if len(client_id) <= 8:
        return "[redacted]"
    return f"{client_id[:4]}...{client_id[-6:]}"



_REDACT_FIELDS = ("access_token", "refresh_token", "id_token", "Authorization")
_REDACT_TOKEN_RE = re.compile(r"\beyJ[a-zA-Z0-9_-]{20,}\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\b")
_REDACT_KV_RE = re.compile(
    r'("(?:access_token|refresh_token|id_token|client_secret)"\s*:\s*")[^"]+',
    re.IGNORECASE,
)


def redact_text(value: str) -> str:
    redacted = _REDACT_KV_RE.sub(r"\1[redacted]", value)
    for key in _REDACT_FIELDS:
        redacted = redacted.replace(key, "[redacted]")
    redacted = _REDACT_TOKEN_RE.sub(
        lambda m: m.group(0)[:12] + "…[redacted]",
        redacted,
    )
    return redacted


class RedactingFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        record.msg = redact_text(str(record.msg))
        if record.args:
            record.args = tuple(
                redact_text(arg.decode("utf-8", errors="replace")) if isinstance(arg, bytes)
                else redact_text(str(arg)) if isinstance(arg, str)
                else arg
                for arg in record.args
            )
        return True


if __name__ == "__main__":
    raise SystemExit(main())
