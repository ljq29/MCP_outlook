"""SQLite-backed credential and Graph token storage for HTTP MCP."""

from __future__ import annotations

import argparse
import atexit
import hashlib
import hmac
import logging
import json
import os
import sqlite3
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Generator

from .graph_auth import GraphAuthorizationBinding
from .launcher_json import validate_secret


DEFAULT_DB_PATH = "~/.outlook-mcp/outlook-mcp.sqlite3"

_logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class CredentialRecord:
    id: int
    method: str
    account: str
    secret_hash: str


@dataclass(frozen=True)
class CalendarDraftRecord:
    id: str
    credential_id: int
    draft_type: str
    source_event_id: str | None
    payload: dict[str, Any]
    base_snapshot: dict[str, Any] | None
    base_fingerprint: str | None
    subject: str
    is_all_day: bool
    start_at: float | None
    end_at: float | None
    start_date: str | None
    end_date: str | None
    timezone: str
    status: str
    version: int
    submission_started_at: float | None
    expires_at: float | None
    last_error_code: str | None
    last_error_message: str | None
    created_at: float
    updated_at: float


class CalendarDraftNotFoundError(LookupError):
    pass


class CalendarDraftConflictError(RuntimeError):
    def __init__(self, record: CalendarDraftRecord | None = None):
        super().__init__("calendar draft version conflict")
        self.record = record


class CalendarDraftStateError(RuntimeError):
    def __init__(self, status: str):
        super().__init__(f"calendar draft cannot be changed while status is {status}")
        self.status = status


def hash_secret(secret: str) -> str:
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()


class CredentialStore:
    _LOG_FLUSH_INTERVAL = 5.0
    _LOG_FLUSH_BATCH = 100
    _LOG_RETENTION_DAYS = 30
    _DRAFT_SUBMISSION_STALE_SECONDS = 120.0

    def __init__(self, path: str | Path = DEFAULT_DB_PATH):
        self.path = Path(path).expanduser()
        self._lock = threading.Lock()
        self._log_buffer: list[tuple] = []
        self._log_buffer_lock = threading.Lock()
        self._log_buffer_event = threading.Event()
        self._flush_thread: threading.Thread | None = None
        self._flush_running = False
        self._permissions_ready = False
        self._atexit_registered = False
        self._closed = False
        self.initialize()

    def initialize(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS http_credentials (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    method TEXT NOT NULL,
                    account TEXT NOT NULL,
                    secret_hash TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL,
                    UNIQUE(method, account, secret_hash)
                );

                CREATE TABLE IF NOT EXISTS graph_authorizations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    expected_account TEXT NOT NULL,
                    account_key TEXT NOT NULL,
                    home_account_id TEXT,
                    graph_user_id TEXT,
                    identities_json TEXT NOT NULL DEFAULT '[]',
                    config_fingerprint TEXT,
                    msal_cache_json TEXT,
                    pending_flow_json TEXT,
                    pending_expires_at REAL,
                    status TEXT NOT NULL DEFAULT 'unbound',
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_graph_authorizations_account_config
                ON graph_authorizations(account_key, config_fingerprint);

                CREATE TABLE IF NOT EXISTS http_operation_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    credential_id INTEGER NOT NULL,
                    created_at REAL NOT NULL,
                    operation TEXT NOT NULL,
                    http_method TEXT NOT NULL,
                    status_code INTEGER,
                    ok INTEGER NOT NULL,
                    error TEXT,
                    duration_ms REAL NOT NULL,
                    FOREIGN KEY(credential_id) REFERENCES http_credentials(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_http_operation_logs_credential_created
                ON http_operation_logs(credential_id, created_at DESC);

                CREATE INDEX IF NOT EXISTS idx_http_operation_logs_created_at
                ON http_operation_logs(created_at);

                CREATE TABLE IF NOT EXISTS calendar_draft_settings (
                    credential_id INTEGER PRIMARY KEY,
                    retention_days INTEGER,
                    updated_at REAL NOT NULL,
                    CHECK(retention_days IS NULL OR retention_days BETWEEN 1 AND 3650),
                    FOREIGN KEY(credential_id) REFERENCES http_credentials(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS calendar_drafts (
                    id TEXT PRIMARY KEY,
                    credential_id INTEGER NOT NULL,
                    draft_type TEXT NOT NULL CHECK(draft_type IN ('create', 'update')),
                    source_event_id TEXT,
                    payload_json TEXT NOT NULL,
                    base_snapshot_json TEXT,
                    base_fingerprint TEXT,
                    subject TEXT NOT NULL,
                    is_all_day INTEGER NOT NULL,
                    start_at REAL,
                    end_at REAL,
                    start_date TEXT,
                    end_date TEXT,
                    timezone TEXT NOT NULL,
                    status TEXT NOT NULL CHECK(status IN ('active', 'submitting', 'outcome_unknown')),
                    version INTEGER NOT NULL,
                    submission_started_at REAL,
                    expires_at REAL,
                    last_error_code TEXT,
                    last_error_message TEXT,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL,
                    CHECK(
                        (draft_type = 'create' AND source_event_id IS NULL)
                        OR (draft_type = 'update' AND source_event_id IS NOT NULL)
                    ),
                    FOREIGN KEY(credential_id) REFERENCES http_credentials(id) ON DELETE CASCADE
                );

                CREATE UNIQUE INDEX IF NOT EXISTS idx_calendar_drafts_update_source
                ON calendar_drafts(credential_id, source_event_id)
                WHERE draft_type = 'update';

                CREATE INDEX IF NOT EXISTS idx_calendar_drafts_credential_updated
                ON calendar_drafts(credential_id, updated_at DESC);

                CREATE INDEX IF NOT EXISTS idx_calendar_drafts_expires_at
                ON calendar_drafts(expires_at)
                WHERE expires_at IS NOT NULL;
                """
            )
            columns = {str(row[1]) for row in conn.execute("PRAGMA table_info(http_credentials)")}
            if "graph_authorization_id" not in columns:
                conn.execute(
                    "ALTER TABLE http_credentials ADD COLUMN graph_authorization_id INTEGER"
                )
            self._migrate_legacy_graph_state(conn)
            self._detach_shared_graph_authorizations(conn)
        self._ensure_private_permissions()
        self._prune_old_logs()
        self._start_flush_thread()

    def _migrate_legacy_graph_state(self, conn: sqlite3.Connection) -> None:
        legacy_table = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'graph_oauth_state'"
        ).fetchone()
        if legacy_table is None:
            return
        rows = conn.execute(
            """
            SELECT c.id, c.account, s.msal_cache_json, s.pending_flow_json, s.pending_expires_at
            FROM http_credentials c
            JOIN graph_oauth_state s ON s.credential_id = c.id
            WHERE c.graph_authorization_id IS NULL
            ORDER BY c.id
            """
        ).fetchall()
        now = time.time()
        for row in rows:
            cache_json = row["msal_cache_json"]
            pending_json = row["pending_flow_json"]
            home_account_id = _legacy_home_account_id(cache_json)
            status = "legacy_unverified" if cache_json else "unbound"
            cursor = conn.execute(
                """
                INSERT INTO graph_authorizations(
                    expected_account, account_key, home_account_id,
                    msal_cache_json, pending_flow_json, pending_expires_at,
                    status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(row["account"]),
                    _normalize_account_key(str(row["account"])),
                    home_account_id,
                    cache_json,
                    pending_json,
                    row["pending_expires_at"],
                    status,
                    now,
                    now,
                ),
            )
            conn.execute(
                "UPDATE http_credentials SET graph_authorization_id = ? WHERE id = ?",
                (cursor.lastrowid, row["id"]),
            )
        # The legacy table is a one-time migration source. Drop it so no shadow
        # token rows can resurrect stale OAuth state on the next process start.
        conn.execute("DROP TABLE graph_oauth_state")

    def _detach_shared_graph_authorizations(self, conn: sqlite3.Connection) -> None:
        """Fail closed for credentials linked by the former implicit reuse behavior."""
        conn.execute(
            """
            UPDATE http_credentials
            SET graph_authorization_id = NULL
            WHERE graph_authorization_id IN (
                SELECT graph_authorization_id
                FROM http_credentials
                WHERE graph_authorization_id IS NOT NULL
                GROUP BY graph_authorization_id
                HAVING COUNT(*) > 1
            )
            """
        )
        conn.execute(
            """
            DELETE FROM graph_authorizations
            WHERE id NOT IN (
                SELECT graph_authorization_id
                FROM http_credentials
                WHERE graph_authorization_id IS NOT NULL
            )
            """
        )

    def lookup_endpoint(self, method: str, account: str, secret: str) -> CredentialRecord | None:
        return self.lookup_endpoint_hash(method, account, hash_secret(secret))

    def lookup_endpoint_hash(self, method: str, account: str, secret_hash: str) -> CredentialRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT id, method, account, secret_hash
                FROM http_credentials
                WHERE method = ? AND account = ? AND secret_hash = ?
                """,
                (_normalize_method(method), account, secret_hash),
            ).fetchone()
        return _credential_record(row)

    def get_or_create_endpoint(self, method: str, account: str, secret: str) -> CredentialRecord:
        method = _normalize_method(method)
        secret_hash = hash_secret(secret)
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                conn.execute(
                    """
                    INSERT INTO http_credentials(method, account, secret_hash, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(method, account, secret_hash) DO UPDATE SET updated_at = excluded.updated_at
                    """,
                    (method, account, secret_hash, now, now),
                )
                row = conn.execute(
                    """
                    SELECT id, method, account, secret_hash
                    FROM http_credentials
                    WHERE method = ? AND account = ? AND secret_hash = ?
                    """,
                    (method, account, secret_hash),
                ).fetchone()
        record = _credential_record(row)
        if record is None:
            raise RuntimeError("Failed to create HTTP credential record")
        return record

    def rotate_endpoint(
        self,
        method: str,
        account: str,
        old_secret: str,
        new_secret: str,
    ) -> CredentialRecord:
        """Atomically replace a registered secret without changing its credential id."""
        validate_secret(old_secret)
        validate_secret(new_secret)
        if hmac.compare_digest(old_secret, new_secret):
            raise ValueError("new secret must differ from old secret")
        method = _normalize_method(method)
        old_hash = hash_secret(old_secret)
        new_hash = hash_secret(new_secret)
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                row = conn.execute(
                    """
                    SELECT id, method, account, secret_hash
                    FROM http_credentials
                    WHERE method = ? AND account = ? AND secret_hash = ?
                    """,
                    (method, account, old_hash),
                ).fetchone()
                if row is None:
                    raise ValueError("old credential was not found")
                conflict = conn.execute(
                    """
                    SELECT id FROM http_credentials
                    WHERE method = ? AND account = ? AND secret_hash = ?
                    """,
                    (method, account, new_hash),
                ).fetchone()
                if conflict is not None:
                    raise ValueError("new secret is already registered for this endpoint")
                conn.execute(
                    """
                    UPDATE http_credentials
                    SET secret_hash = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (new_hash, now, row["id"]),
                )
                updated = conn.execute(
                    """
                    SELECT id, method, account, secret_hash
                    FROM http_credentials WHERE id = ?
                    """,
                    (row["id"],),
                ).fetchone()
        record = _credential_record(updated)
        if record is None:
            raise RuntimeError("Failed to rotate HTTP credential")
        return record

    def get_calendar_draft_settings(self, credential_id: int) -> dict[str, Any]:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT retention_days, updated_at FROM calendar_draft_settings WHERE credential_id = ?",
                    (credential_id,),
                ).fetchone()
        return {
            "retention_days": None if row is None or row["retention_days"] is None else int(row["retention_days"]),
            "updated_at": None if row is None else float(row["updated_at"]),
        }

    def set_calendar_draft_retention(
        self,
        credential_id: int,
        retention_days: int | None,
    ) -> dict[str, Any]:
        if retention_days is not None and not 1 <= retention_days <= 3650:
            raise ValueError("retention_days must be between 1 and 3650 days")
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                conn.execute(
                    """
                    INSERT INTO calendar_draft_settings(credential_id, retention_days, updated_at)
                    VALUES (?, ?, ?)
                    ON CONFLICT(credential_id) DO UPDATE SET
                        retention_days = excluded.retention_days,
                        updated_at = excluded.updated_at
                    """,
                    (credential_id, retention_days, now),
                )
                if retention_days is None:
                    conn.execute(
                        "UPDATE calendar_drafts SET expires_at = NULL WHERE credential_id = ?",
                        (credential_id,),
                    )
                else:
                    conn.execute(
                        "UPDATE calendar_drafts SET expires_at = updated_at + ? WHERE credential_id = ?",
                        (retention_days * 86400.0, credential_id),
                    )
                expired_count = int(
                    conn.execute(
                        "SELECT COUNT(*) FROM calendar_drafts WHERE credential_id = ? AND expires_at <= ?",
                        (credential_id, now),
                    ).fetchone()[0]
                )
                conn.execute(
                    "DELETE FROM calendar_drafts WHERE credential_id = ? AND expires_at <= ?",
                    (credential_id, now),
                )
        return {"retention_days": retention_days, "expired_count": expired_count, "updated_at": now}

    def preview_calendar_draft_retention(
        self,
        credential_id: int,
        retention_days: int | None,
    ) -> dict[str, Any]:
        if retention_days is not None and not 1 <= retention_days <= 3650:
            raise ValueError("retention_days must be between 1 and 3650 days")
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                if retention_days is None:
                    expired_count = 0
                else:
                    expired_count = int(conn.execute(
                        """
                        SELECT COUNT(*) FROM calendar_drafts
                        WHERE credential_id = ? AND updated_at + ? <= ?
                        """,
                        (credential_id, retention_days * 86400.0, now),
                    ).fetchone()[0])
        return {"retention_days": retention_days, "expired_count": expired_count}

    def save_calendar_draft(
        self,
        credential_id: int,
        *,
        draft_id: str,
        expected_version: int | None,
        draft_type: str,
        source_event_id: str | None,
        payload: dict[str, Any],
        base_snapshot: dict[str, Any] | None,
        base_fingerprint: str | None,
        subject: str,
        is_all_day: bool,
        start_at: float | None,
        end_at: float | None,
        start_date: str | None,
        end_date: str | None,
        timezone: str,
    ) -> CalendarDraftRecord:
        if draft_type not in {"create", "update"}:
            raise ValueError("draft_type must be create or update")
        if (draft_type == "create") != (source_event_id is None):
            raise ValueError("source_event_id is required only for update drafts")
        now = time.time()
        payload_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        base_json = None if base_snapshot is None else json.dumps(
            base_snapshot, ensure_ascii=False, separators=(",", ":")
        )
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                retention_days = self._calendar_draft_retention_days(conn, credential_id)
                expires_at = None if retention_days is None else now + retention_days * 86400.0
                if expected_version is None:
                    try:
                        conn.execute(
                            """
                            INSERT INTO calendar_drafts(
                                id, credential_id, draft_type, source_event_id,
                                payload_json, base_snapshot_json, base_fingerprint,
                                subject, is_all_day, start_at, end_at, start_date, end_date,
                                timezone, status, version, expires_at, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', 1, ?, ?, ?)
                            """,
                            (
                                draft_id, credential_id, draft_type, source_event_id,
                                payload_json, base_json, base_fingerprint,
                                subject, 1 if is_all_day else 0, start_at, end_at,
                                start_date, end_date, timezone, expires_at, now, now,
                            ),
                        )
                    except sqlite3.IntegrityError as exc:
                        conflict = self._calendar_draft_conflict_row(
                            conn, credential_id, draft_id, source_event_id
                        )
                        raise CalendarDraftConflictError(_calendar_draft_record(conflict)) from exc
                else:
                    row = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    if row is None:
                        raise CalendarDraftNotFoundError(draft_id)
                    current = _calendar_draft_record(row)
                    if current.status != "active":
                        raise CalendarDraftStateError(current.status)
                    if current.version != expected_version:
                        raise CalendarDraftConflictError(current)
                    if current.draft_type != draft_type or current.source_event_id != source_event_id:
                        raise CalendarDraftConflictError(current)
                    cursor = conn.execute(
                        """
                        UPDATE calendar_drafts SET
                            payload_json = ?, base_snapshot_json = ?, base_fingerprint = ?,
                            subject = ?, is_all_day = ?, start_at = ?, end_at = ?,
                            start_date = ?, end_date = ?, timezone = ?, version = version + 1,
                            expires_at = ?, last_error_code = NULL, last_error_message = NULL,
                            updated_at = ?
                        WHERE credential_id = ? AND id = ? AND version = ?
                        """,
                        (
                            payload_json, base_json, base_fingerprint, subject,
                            1 if is_all_day else 0, start_at, end_at, start_date, end_date,
                            timezone, expires_at, now, credential_id, draft_id, expected_version,
                        ),
                    )
                    if cursor.rowcount != 1:
                        conflict = conn.execute(
                            "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                            (credential_id, draft_id),
                        ).fetchone()
                        raise CalendarDraftConflictError(_calendar_draft_record(conflict))
                saved = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        record = _calendar_draft_record(saved)
        if record is None:
            raise RuntimeError("failed to save calendar draft")
        return record

    def get_calendar_draft(self, credential_id: int, draft_id: str) -> CalendarDraftRecord | None:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        return _calendar_draft_record(row)

    def list_calendar_drafts(
        self,
        credential_id: int,
        *,
        limit: int = 100,
        offset: int = 0,
        range_start_at: float | None = None,
        range_end_at: float | None = None,
        range_start_date: str | None = None,
        range_end_date: str | None = None,
    ) -> tuple[list[CalendarDraftRecord], int]:
        limit = max(1, min(int(limit), 100))
        offset = max(0, int(offset))
        clauses = ["credential_id = ?"]
        params: list[Any] = [credential_id]
        if range_start_at is not None and range_end_at is not None:
            clauses.append(
                "((is_all_day = 0 AND start_at < ? AND end_at > ?) "
                "OR (is_all_day = 1 AND start_date < ? AND end_date > ?))"
            )
            params.extend([range_end_at, range_start_at, range_end_date, range_start_date])
        where_sql = " AND ".join(clauses)
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                total = int(conn.execute(
                    "SELECT COUNT(*) FROM calendar_drafts WHERE " + where_sql, params
                ).fetchone()[0])
                rows = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE " + where_sql
                    + " ORDER BY updated_at DESC LIMIT ? OFFSET ?",
                    [*params, limit, offset],
                ).fetchall()
        return [_calendar_draft_record(row) for row in rows], total

    def delete_calendar_draft(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
    ) -> None:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    raise CalendarDraftNotFoundError(draft_id)
                current = _calendar_draft_record(row)
                if current.status == "submitting":
                    raise CalendarDraftStateError(current.status)
                if current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                cursor = conn.execute(
                    "DELETE FROM calendar_drafts WHERE credential_id = ? AND id = ? AND version = ?",
                    (credential_id, draft_id, expected_version),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))

    def convert_calendar_draft_to_create(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
    ) -> CalendarDraftRecord:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    raise CalendarDraftNotFoundError(draft_id)
                current = _calendar_draft_record(row)
                if current.status != "active":
                    raise CalendarDraftStateError(current.status)
                if current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                if current.draft_type != "update":
                    raise ValueError("only update drafts can be converted to create drafts")
                retention_days = self._calendar_draft_retention_days(conn, credential_id)
                expires_at = None if retention_days is None else now + retention_days * 86400.0
                cursor = conn.execute(
                    """
                    UPDATE calendar_drafts SET
                        draft_type = 'create', source_event_id = NULL,
                        base_snapshot_json = NULL, base_fingerprint = NULL,
                        version = version + 1, last_error_code = NULL,
                        last_error_message = NULL, expires_at = ?, updated_at = ?
                    WHERE credential_id = ? AND id = ? AND version = ?
                    """,
                    (expires_at, now, credential_id, draft_id, expected_version),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))
                updated = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        record = _calendar_draft_record(updated)
        if record is None:
            raise RuntimeError("failed to convert calendar draft")
        return record

    def begin_calendar_draft_submission(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
    ) -> CalendarDraftRecord:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    raise CalendarDraftNotFoundError(draft_id)
                current = _calendar_draft_record(row)
                if current.status != "active":
                    raise CalendarDraftStateError(current.status)
                if current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                cursor = conn.execute(
                    """
                    UPDATE calendar_drafts SET
                        status = 'submitting', version = version + 1,
                        submission_started_at = ?, last_error_code = NULL,
                        last_error_message = NULL, updated_at = ?
                    WHERE credential_id = ? AND id = ? AND version = ?
                    """,
                    (now, now, credential_id, draft_id, expected_version),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))
                updated = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        record = _calendar_draft_record(updated)
        if record is None:
            raise RuntimeError("failed to begin calendar draft submission")
        return record

    def finish_calendar_draft_submission(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
        *,
        status: str,
        error_code: str | None = None,
        error_message: str | None = None,
    ) -> CalendarDraftRecord:
        if status not in {"active", "outcome_unknown"}:
            raise ValueError("submission result status must be active or outcome_unknown")
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    raise CalendarDraftNotFoundError(draft_id)
                current = _calendar_draft_record(row)
                if current.status != "submitting":
                    raise CalendarDraftStateError(current.status)
                if current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                retention_days = self._calendar_draft_retention_days(conn, credential_id)
                expires_at = None if retention_days is None else now + retention_days * 86400.0
                cursor = conn.execute(
                    """
                    UPDATE calendar_drafts SET
                        status = ?, version = version + 1, submission_started_at = NULL,
                        last_error_code = ?, last_error_message = ?, expires_at = ?, updated_at = ?
                    WHERE credential_id = ? AND id = ? AND version = ?
                    """,
                    (
                        status, error_code, error_message, expires_at, now,
                        credential_id, draft_id, expected_version,
                    ),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))
                updated = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        record = _calendar_draft_record(updated)
        if record is None:
            raise RuntimeError("failed to finish calendar draft submission")
        return record

    def delete_submitted_calendar_draft(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
    ) -> None:
        with self._lock:
            with self._connect() as conn:
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    return
                current = _calendar_draft_record(row)
                if current.status != "submitting" or current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                cursor = conn.execute(
                    "DELETE FROM calendar_drafts WHERE credential_id = ? AND id = ? AND version = ?",
                    (credential_id, draft_id, expected_version),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))

    def resolve_calendar_draft_outcome(
        self,
        credential_id: int,
        draft_id: str,
        expected_version: int,
    ) -> CalendarDraftRecord:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                self._prepare_calendar_drafts(conn, credential_id, now)
                row = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
                if row is None:
                    raise CalendarDraftNotFoundError(draft_id)
                current = _calendar_draft_record(row)
                if current.status != "outcome_unknown":
                    raise CalendarDraftStateError(current.status)
                if current.version != expected_version:
                    raise CalendarDraftConflictError(current)
                retention_days = self._calendar_draft_retention_days(conn, credential_id)
                expires_at = None if retention_days is None else now + retention_days * 86400.0
                cursor = conn.execute(
                    """
                    UPDATE calendar_drafts SET
                        status = 'active', version = version + 1,
                        last_error_code = NULL, last_error_message = NULL,
                        expires_at = ?, updated_at = ?
                    WHERE credential_id = ? AND id = ? AND version = ?
                    """,
                    (expires_at, now, credential_id, draft_id, expected_version),
                )
                if cursor.rowcount != 1:
                    conflict = conn.execute(
                        "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                        (credential_id, draft_id),
                    ).fetchone()
                    raise CalendarDraftConflictError(_calendar_draft_record(conflict))
                updated = conn.execute(
                    "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
                    (credential_id, draft_id),
                ).fetchone()
        record = _calendar_draft_record(updated)
        if record is None:
            raise RuntimeError("failed to resolve calendar draft outcome")
        return record

    def _prepare_calendar_drafts(
        self,
        conn: sqlite3.Connection,
        credential_id: int,
        now: float,
    ) -> None:
        stale_before = now - self._DRAFT_SUBMISSION_STALE_SECONDS
        retention_days = self._calendar_draft_retention_days(conn, credential_id)
        expires_at = None if retention_days is None else now + retention_days * 86400.0
        conn.execute(
            """
            UPDATE calendar_drafts SET
                status = 'outcome_unknown', version = version + 1,
                submission_started_at = NULL, last_error_code = 'outcome_unknown',
                last_error_message = '提交进程中断，远端结果未知。',
                expires_at = ?, updated_at = ?
            WHERE credential_id = ? AND status = 'submitting'
              AND submission_started_at <= ?
            """,
            (expires_at, now, credential_id, stale_before),
        )
        conn.execute(
            "DELETE FROM calendar_drafts WHERE credential_id = ? AND expires_at IS NOT NULL AND expires_at <= ?",
            (credential_id, now),
        )

    @staticmethod
    def _calendar_draft_retention_days(
        conn: sqlite3.Connection,
        credential_id: int,
    ) -> int | None:
        row = conn.execute(
            "SELECT retention_days FROM calendar_draft_settings WHERE credential_id = ?",
            (credential_id,),
        ).fetchone()
        return None if row is None or row["retention_days"] is None else int(row["retention_days"])

    @staticmethod
    def _calendar_draft_conflict_row(
        conn: sqlite3.Connection,
        credential_id: int,
        draft_id: str,
        source_event_id: str | None,
    ) -> sqlite3.Row | None:
        if source_event_id is not None:
            row = conn.execute(
                "SELECT * FROM calendar_drafts WHERE credential_id = ? AND source_event_id = ?",
                (credential_id, source_event_id),
            ).fetchone()
            if row is not None:
                return row
        return conn.execute(
            "SELECT * FROM calendar_drafts WHERE credential_id = ? AND id = ?",
            (credential_id, draft_id),
        ).fetchone()

    def get_graph_authorization_id(self, credential_id: int) -> int | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT graph_authorization_id FROM http_credentials WHERE id = ?",
                (credential_id,),
            ).fetchone()
        if row is None or row["graph_authorization_id"] is None:
            return None
        return int(row["graph_authorization_id"])

    def _ensure_graph_authorization_for_credential(self, credential_id: int) -> int:
        existing = self.get_graph_authorization_id(credential_id)
        if existing is not None:
            return existing
        with self._lock:
            with self._connect() as conn:
                credential = conn.execute(
                    "SELECT account FROM http_credentials WHERE id = ?",
                    (credential_id,),
                ).fetchone()
                if credential is None:
                    raise RuntimeError(f"Unknown HTTP credential: {credential_id}")
                existing_row = conn.execute(
                    "SELECT graph_authorization_id FROM http_credentials WHERE id = ?",
                    (credential_id,),
                ).fetchone()
                if existing_row["graph_authorization_id"] is not None:
                    return int(existing_row["graph_authorization_id"])
                now = time.time()
                cursor = conn.execute(
                    """
                    INSERT INTO graph_authorizations(
                        expected_account, account_key, status, created_at, updated_at
                    ) VALUES (?, ?, 'unbound', ?, ?)
                    """,
                    (
                        str(credential["account"]),
                        _normalize_account_key(str(credential["account"])),
                        now,
                        now,
                    ),
                )
                authorization_id = int(cursor.lastrowid)
                conn.execute(
                    "UPDATE http_credentials SET graph_authorization_id = ? WHERE id = ?",
                    (authorization_id, credential_id),
                )
                return authorization_id

    def resolve_graph_authorization(
        self,
        credential_id: int,
        expected_account: str,
        config_fingerprint: str,
    ) -> int:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                credential = conn.execute(
                    "SELECT graph_authorization_id FROM http_credentials WHERE id = ?",
                    (credential_id,),
                ).fetchone()
                if credential is None:
                    raise RuntimeError(f"Unknown HTTP credential: {credential_id}")
                authorization_id = credential["graph_authorization_id"]
                if authorization_id is not None:
                    current = conn.execute(
                        """
                        SELECT expected_account, config_fingerprint
                        FROM graph_authorizations WHERE id = ?
                        """,
                        (authorization_id,),
                    ).fetchone()
                    if current is not None and (
                        _normalize_account_key(str(current["expected_account"]))
                        == _normalize_account_key(expected_account)
                        and current["config_fingerprint"] in (None, config_fingerprint)
                    ):
                        if current["config_fingerprint"] is None:
                            conn.execute(
                                "UPDATE graph_authorizations SET config_fingerprint = ?, updated_at = ? WHERE id = ?",
                                (config_fingerprint, now, authorization_id),
                            )
                        return int(authorization_id)

                cursor = conn.execute(
                    """
                    INSERT INTO graph_authorizations(
                        expected_account, account_key, config_fingerprint,
                        status, created_at, updated_at
                    ) VALUES (?, ?, ?, 'unbound', ?, ?)
                    """,
                    (
                        expected_account,
                        _normalize_account_key(expected_account),
                        config_fingerprint,
                        now,
                        now,
                    ),
                )
                authorization_id = int(cursor.lastrowid)
                conn.execute(
                    "UPDATE http_credentials SET graph_authorization_id = ? WHERE id = ?",
                    (authorization_id, credential_id),
                )
                previous_authorization_id = credential["graph_authorization_id"]
                if previous_authorization_id is not None:
                    conn.execute(
                        """
                        DELETE FROM graph_authorizations
                        WHERE id = ?
                          AND NOT EXISTS (
                              SELECT 1 FROM http_credentials
                              WHERE graph_authorization_id = graph_authorizations.id
                          )
                        """,
                        (previous_authorization_id,),
                    )
                return authorization_id

    def load_graph_authorization_binding(self, authorization_id: int) -> GraphAuthorizationBinding | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT expected_account, home_account_id, graph_user_id, identities_json, status
                FROM graph_authorizations
                WHERE id = ?
                """,
                (authorization_id,),
            ).fetchone()
        if row is None:
            return None
        try:
            identities = json.loads(row["identities_json"] or "[]")
        except json.JSONDecodeError:
            identities = []
        return GraphAuthorizationBinding(
            expected_account=str(row["expected_account"]),
            home_account_id=str(row["home_account_id"] or "").strip() or None,
            graph_user_id=str(row["graph_user_id"] or "").strip() or None,
            identities=tuple(str(item).strip() for item in identities if str(item).strip()),
            status=str(row["status"] or "unbound"),
        )

    def save_graph_authorization_binding(
        self,
        authorization_id: int,
        binding: GraphAuthorizationBinding,
    ) -> None:
        self._update_graph_authorization(
            authorization_id,
            expected_account=binding.expected_account,
            account_key=_normalize_account_key(binding.expected_account),
            home_account_id=binding.home_account_id,
            graph_user_id=binding.graph_user_id,
            identities_json=json.dumps(list(binding.identities), ensure_ascii=True),
            status=binding.status,
        )

    def clear_graph_authorization_binding(self, authorization_id: int) -> None:
        self._update_graph_authorization(
            authorization_id,
            home_account_id=None,
            graph_user_id=None,
            identities_json="[]",
            pending_flow_json=None,
            pending_expires_at=None,
            status="unbound",
        )

    def clear_graph_authorization(self, authorization_id: int) -> None:
        self._update_graph_authorization(
            authorization_id,
            home_account_id=None,
            graph_user_id=None,
            identities_json="[]",
            msal_cache_json=None,
            pending_flow_json=None,
            pending_expires_at=None,
            status="unbound",
        )

    def load_msal_cache_for_authorization(self, authorization_id: int) -> str | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT msal_cache_json FROM graph_authorizations WHERE id = ?",
                (authorization_id,),
            ).fetchone()
        return None if row is None else row["msal_cache_json"]

    def save_msal_cache_for_authorization(self, authorization_id: int, cache_json: str) -> None:
        self._update_graph_authorization(authorization_id, msal_cache_json=cache_json)

    def load_pending_flow_for_authorization(self, authorization_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT pending_flow_json, pending_expires_at
                FROM graph_authorizations
                WHERE id = ?
                """,
                (authorization_id,),
            ).fetchone()
        if row is None or not row["pending_flow_json"]:
            return None
        if row["pending_expires_at"] is not None and row["pending_expires_at"] < time.time():
            self._update_graph_authorization(
                authorization_id,
                pending_flow_json=None,
                pending_expires_at=None,
            )
            return None
        try:
            flow = json.loads(row["pending_flow_json"])
        except json.JSONDecodeError:
            self._update_graph_authorization(
                authorization_id,
                pending_flow_json=None,
                pending_expires_at=None,
            )
            return None
        return flow if isinstance(flow, dict) else None

    def save_pending_flow_for_authorization(self, authorization_id: int, flow: dict[str, Any]) -> None:
        expires_at = flow.get("expires_at")
        if not isinstance(expires_at, (int, float)):
            expires_in = flow.get("expires_in")
            expires_at = time.time() + float(expires_in) if isinstance(expires_in, (int, float)) else 0.0
        self._update_graph_authorization(
            authorization_id,
            pending_flow_json=json.dumps(flow),
            pending_expires_at=float(expires_at or 0),
        )

    def delete_pending_flow_for_authorization(self, authorization_id: int) -> None:
        self._update_graph_authorization(
            authorization_id,
            pending_flow_json=None,
            pending_expires_at=None,
        )

    def load_msal_cache(self, credential_id: int) -> str | None:
        authorization_id = self.get_graph_authorization_id(credential_id)
        if authorization_id is None:
            return None
        return self.load_msal_cache_for_authorization(authorization_id)

    def save_msal_cache(self, credential_id: int, cache_json: str) -> None:
        authorization_id = self._ensure_graph_authorization_for_credential(credential_id)
        self.save_msal_cache_for_authorization(authorization_id, cache_json)

    def load_pending_flow(self, credential_id: int) -> dict[str, Any] | None:
        authorization_id = self.get_graph_authorization_id(credential_id)
        if authorization_id is None:
            return None
        return self.load_pending_flow_for_authorization(authorization_id)

    def save_pending_flow(self, credential_id: int, flow: dict[str, Any]) -> None:
        authorization_id = self._ensure_graph_authorization_for_credential(credential_id)
        self.save_pending_flow_for_authorization(authorization_id, flow)

    def delete_pending_flow(self, credential_id: int) -> None:
        authorization_id = self.get_graph_authorization_id(credential_id)
        if authorization_id is not None:
            self.delete_pending_flow_for_authorization(authorization_id)

    def clear_graph_oauth_state(self, credential_id: int) -> None:
        authorization_id = self.get_graph_authorization_id(credential_id)
        if authorization_id is not None:
            self.clear_graph_authorization(authorization_id)

    def append_operation_log(
        self,
        credential_id: int,
        *,
        operation: str,
        http_method: str,
        status_code: int | None,
        ok: bool,
        error: str | None,
        duration_ms: float,
        created_at: float | None = None,
    ) -> None:
        now = time.time() if created_at is None else created_at
        entry = (
            credential_id,
            now,
            operation,
            http_method.upper(),
            status_code,
            1 if ok else 0,
            error,
            duration_ms,
        )
        with self._log_buffer_lock:
            self._log_buffer.append(entry)
            should_signal = len(self._log_buffer) >= self._LOG_FLUSH_BATCH
        if should_signal:
            self._log_buffer_event.set()

    def flush_logs(self) -> None:
        """Flush buffered log entries to disk."""
        with self._log_buffer_lock:
            if not self._log_buffer:
                return
            batch = self._log_buffer[:]
            self._log_buffer.clear()
        with self._lock:
            with self._connect() as conn:
                conn.executemany(
                    """
                    INSERT INTO http_operation_logs(
                        credential_id, created_at, operation,
                        http_method, status_code, ok, error, duration_ms
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    batch,
                )

    def _start_flush_thread(self) -> None:
        self._flush_running = True
        self._flush_thread = threading.Thread(target=self._flush_loop, daemon=True)
        self._flush_thread.start()
        if not self._atexit_registered:
            atexit.register(self._flush_exit)
            self._atexit_registered = True

    def _flush_loop(self) -> None:
        while self._flush_running:
            self._log_buffer_event.wait(timeout=self._LOG_FLUSH_INTERVAL)
            self._log_buffer_event.clear()
            try:
                self.flush_logs()
            except Exception:
                _logger.exception("Failed to flush HTTP operation logs")

    def _flush_exit(self) -> None:
        self._flush_running = False
        self._log_buffer_event.set()
        if self._flush_thread is not None:
            self._flush_thread.join(timeout=10)
        if self._flush_thread is not None and self._flush_thread.is_alive():
            try:
                self.flush_logs()
            except Exception:
                _logger.exception("Failed to flush HTTP operation logs during shutdown")
            return
        self.flush_logs()

    def close(self) -> None:
        """Stop the flush thread, unregister the atexit hook and flush buffered logs."""
        if self._closed:
            return
        self._closed = True
        self._flush_running = False
        self._log_buffer_event.set()
        thread = self._flush_thread
        if thread is not None and thread is not threading.current_thread():
            thread.join(timeout=2)
        if self._atexit_registered:
            atexit.unregister(self._flush_exit)
            self._atexit_registered = False
        self.flush_logs()

    def _prune_old_logs(self) -> None:
        cutoff = time.time() - self._LOG_RETENTION_DAYS * 24 * 3600
        while True:
            with self._lock:
                with self._connect() as conn:
                    cur = conn.execute(
                        "DELETE FROM http_operation_logs WHERE id IN ("
                        "SELECT id FROM http_operation_logs WHERE created_at < ? ORDER BY id LIMIT 1000"
                        ")",
                        (cutoff,),
                    )
                    if cur.rowcount < 1000:
                        break

    def list_operation_logs(
        self,
        credential_id: int,
        *,
        start_at: float | None = None,
        end_at: float | None = None,
        limit: int = 500,
    ) -> list[dict[str, Any]]:
        clauses = ["credential_id = ?"]
        params: list[Any] = [credential_id]
        if start_at is not None:
            clauses.append("created_at >= ?")
            params.append(start_at)
        if end_at is not None:
            clauses.append("created_at <= ?")
            params.append(end_at)
        params.append(max(1, min(int(limit), 1000)))
        where_sql = " AND ".join(clauses)
        query = (
            "SELECT id, credential_id, created_at, operation, http_method, status_code, ok, error, duration_ms "
            "FROM http_operation_logs "
            "WHERE " + where_sql + " "
            "ORDER BY created_at DESC "
            "LIMIT ?"
        )
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [_operation_log(row) for row in rows]

    def _update_graph_authorization(self, authorization_id: int, **values: Any) -> None:
        now = time.time()
        with self._lock:
            with self._connect() as conn:
                row = conn.execute(
                    "SELECT * FROM graph_authorizations WHERE id = ?",
                    (authorization_id,),
                ).fetchone()
                if row is None:
                    raise RuntimeError(f"Unknown Graph authorization: {authorization_id}")
                current = dict(row)
                current.update(values)
                conn.execute(
                    """
                    UPDATE graph_authorizations
                    SET expected_account = ?, account_key = ?, home_account_id = ?,
                        graph_user_id = ?, identities_json = ?, config_fingerprint = ?,
                        msal_cache_json = ?, pending_flow_json = ?, pending_expires_at = ?,
                        status = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        current["expected_account"],
                        current["account_key"],
                        current["home_account_id"],
                        current["graph_user_id"],
                        current["identities_json"],
                        current["config_fingerprint"],
                        current["msal_cache_json"],
                        current["pending_flow_json"],
                        current["pending_expires_at"],
                        current["status"],
                        now,
                        authorization_id,
                    ),
                )

    @contextmanager
    def _connect(self) -> Generator[sqlite3.Connection, None, None]:
        conn = sqlite3.connect(str(self.path))
        self._ensure_private_permissions()
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA foreign_keys = ON")
            conn.execute("PRAGMA journal_mode = WAL")
            conn.execute("PRAGMA wal_autocheckpoint = 100")
            conn.execute("PRAGMA synchronous = NORMAL")
            conn.execute("PRAGMA busy_timeout = 5000")
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _ensure_private_permissions(self) -> None:
        try:
            if not self._permissions_ready:
                os.chmod(self.path.parent, 0o700)
                self._permissions_ready = True
            for candidate in (self.path, Path(f"{self.path}-wal"), Path(f"{self.path}-shm")):
                if candidate.exists():
                    os.chmod(candidate, 0o600)
        except OSError as exc:
            _logger.warning("could not restrict permissions on %s: %s", self.path, exc)


def register_endpoint(method: str, account: str, secret: str, db_path: str | Path | None = None) -> CredentialRecord:
    validate_secret(secret)
    store = CredentialStore(db_path or os.environ.get("OUTLOOK_MCP_DB") or DEFAULT_DB_PATH)
    try:
        return store.get_or_create_endpoint(method, account, secret)
    finally:
        store.close()


def rotate_endpoint(
    method: str,
    account: str,
    old_secret: str,
    new_secret: str,
    db_path: str | Path | None = None,
) -> CredentialRecord:
    store = CredentialStore(db_path or os.environ.get("OUTLOOK_MCP_DB") or DEFAULT_DB_PATH)
    try:
        return store.rotate_endpoint(method, account, old_secret, new_secret)
    finally:
        store.close()


class SQLiteGraphTokenStorage:
    def __init__(self, store: CredentialStore, authorization_id: int):
        self.store = store
        self.authorization_id = authorization_id

    def load_cache(self) -> str | None:
        return self.store.load_msal_cache_for_authorization(self.authorization_id)

    def save_cache(self, cache_json: str) -> None:
        self.store.save_msal_cache_for_authorization(self.authorization_id, cache_json)

    def load_pending_flow(self) -> dict[str, Any] | None:
        return self.store.load_pending_flow_for_authorization(self.authorization_id)

    def save_pending_flow(self, flow: dict[str, Any]) -> None:
        self.store.save_pending_flow_for_authorization(self.authorization_id, flow)

    def delete_pending_flow(self) -> None:
        self.store.delete_pending_flow_for_authorization(self.authorization_id)

    def load_binding(self) -> GraphAuthorizationBinding | None:
        return self.store.load_graph_authorization_binding(self.authorization_id)

    def save_binding(self, binding: GraphAuthorizationBinding | None) -> None:
        if binding is None:
            self.clear_binding()
            return
        self.store.save_graph_authorization_binding(self.authorization_id, binding)

    def clear_binding(self) -> None:
        self.store.clear_graph_authorization_binding(self.authorization_id)


def _credential_record(row: sqlite3.Row | None) -> CredentialRecord | None:
    if row is None:
        return None
    return CredentialRecord(
        id=int(row["id"]),
        method=str(row["method"]),
        account=str(row["account"]),
        secret_hash=str(row["secret_hash"]),
    )


def _calendar_draft_record(row: sqlite3.Row | None) -> CalendarDraftRecord | None:
    if row is None:
        return None
    try:
        payload = json.loads(str(row["payload_json"]))
        base_raw = row["base_snapshot_json"]
        base_snapshot = None if base_raw is None else json.loads(str(base_raw))
    except (json.JSONDecodeError, TypeError) as exc:
        raise RuntimeError(f"invalid calendar draft JSON for {row['id']}") from exc
    if not isinstance(payload, dict) or (base_snapshot is not None and not isinstance(base_snapshot, dict)):
        raise RuntimeError(f"invalid calendar draft payload for {row['id']}")
    return CalendarDraftRecord(
        id=str(row["id"]),
        credential_id=int(row["credential_id"]),
        draft_type=str(row["draft_type"]),
        source_event_id=None if row["source_event_id"] is None else str(row["source_event_id"]),
        payload=payload,
        base_snapshot=base_snapshot,
        base_fingerprint=None if row["base_fingerprint"] is None else str(row["base_fingerprint"]),
        subject=str(row["subject"]),
        is_all_day=bool(row["is_all_day"]),
        start_at=None if row["start_at"] is None else float(row["start_at"]),
        end_at=None if row["end_at"] is None else float(row["end_at"]),
        start_date=None if row["start_date"] is None else str(row["start_date"]),
        end_date=None if row["end_date"] is None else str(row["end_date"]),
        timezone=str(row["timezone"]),
        status=str(row["status"]),
        version=int(row["version"]),
        submission_started_at=(
            None if row["submission_started_at"] is None else float(row["submission_started_at"])
        ),
        expires_at=None if row["expires_at"] is None else float(row["expires_at"]),
        last_error_code=None if row["last_error_code"] is None else str(row["last_error_code"]),
        last_error_message=(
            None if row["last_error_message"] is None else str(row["last_error_message"])
        ),
        created_at=float(row["created_at"]),
        updated_at=float(row["updated_at"]),
    )


def _operation_log(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": int(row["id"]),
        "credential_id": int(row["credential_id"]),
        "created_at": float(row["created_at"]),
        "operation": str(row["operation"]),
        "http_method": str(row["http_method"]),
        "status_code": None if row["status_code"] is None else int(row["status_code"]),
        "ok": bool(row["ok"]),
        "error": row["error"],
        "duration_ms": float(row["duration_ms"]),
    }


def _normalize_method(method: str) -> str:
    return method.strip().lower()


def _normalize_account_key(account: str) -> str:
    return account.strip().casefold()


def _legacy_home_account_id(cache_json: str | None) -> str | None:
    if not cache_json:
        return None
    try:
        cache = json.loads(cache_json)
    except json.JSONDecodeError:
        return None
    accounts = cache.get("Account") if isinstance(cache, dict) else None
    if not isinstance(accounts, dict) or len(accounts) != 1:
        return None
    raw = next(iter(accounts.values()))
    if not isinstance(raw, dict):
        return None
    value = raw.get("home_account_id") or raw.get("homeAccountId")
    normalized = str(value or "").strip()
    return normalized or None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Manage Outlook MCP HTTP credentials.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    register = subparsers.add_parser("register", help="Register an HTTP MCP endpoint credential.")
    register.add_argument("method")
    register.add_argument("account")
    register.add_argument("secret")
    register.add_argument("--db", default=os.environ.get("OUTLOOK_MCP_DB") or DEFAULT_DB_PATH)

    rotate = subparsers.add_parser("rotate", help="Atomically rotate an HTTP MCP endpoint credential.")
    rotate.add_argument("method")
    rotate.add_argument("account")
    rotate.add_argument("old_secret")
    rotate.add_argument("new_secret")
    rotate.add_argument("--db", default=os.environ.get("OUTLOOK_MCP_DB") or DEFAULT_DB_PATH)

    args = parser.parse_args(argv)
    if args.command == "register":
        try:
            record = register_endpoint(args.method, args.account, args.secret, args.db)
        except ValueError as exc:
            parser.error(str(exc))
        print(
            json.dumps(
                {
                    "ok": True,
                    "id": record.id,
                    "method": record.method,
                    "account": record.account,
                    "db_path": str(Path(args.db).expanduser()),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    if args.command == "rotate":
        try:
            record = rotate_endpoint(
                args.method,
                args.account,
                args.old_secret,
                args.new_secret,
                args.db,
            )
        except ValueError as exc:
            parser.error(str(exc))
        print(
            json.dumps(
                {
                    "ok": True,
                    "id": record.id,
                    "method": record.method,
                    "account": record.account,
                    "db_path": str(Path(args.db).expanduser()),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
