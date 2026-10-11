from __future__ import annotations

import json
import os
import sqlite3
import time

import pytest

from outlook_mcp.credential_store import (
    CalendarDraftConflictError,
    CalendarDraftStateError,
    CredentialStore,
    hash_secret,
    main as credential_store_main,
)
from outlook_mcp.config import GraphConfig
from outlook_mcp.graph_auth import GraphAuthorizationBinding, graph_config_fingerprint


def test_credential_store_creates_and_looks_up_records(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")

    assert store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars") is None

    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    again = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    found = store.lookup_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")

    assert record.id == again.id == found.id
    assert record.secret_hash == hash_secret("test-secret-at-least-24-chars")


def test_credential_store_manages_graph_state(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")

    store.save_msal_cache(record.id, '{"cache": true}')
    assert store.load_msal_cache(record.id) == '{"cache": true}'

    store.save_pending_flow(record.id, {"user_code": "ABCD", "expires_at": time.time() + 10})
    pending = store.load_pending_flow(record.id)
    assert pending is not None
    assert pending["user_code"] == "ABCD"

    store.save_pending_flow(record.id, {"user_code": "EXPIRED", "expires_at": time.time() - 10})
    assert store.load_pending_flow(record.id) is None


def test_credential_store_clears_only_one_graph_state(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    other = store.get_or_create_endpoint("graph", "other@outlook.com", "test-secret-at-least-24-chars")

    store.save_msal_cache(record.id, '{"cache": true}')
    store.save_pending_flow(record.id, {"user_code": "ABCD", "expires_at": time.time() + 10})
    store.save_msal_cache(other.id, '{"other": true}')
    store.save_pending_flow(other.id, {"user_code": "WXYZ", "expires_at": time.time() + 10})

    store.clear_graph_oauth_state(record.id)

    assert store.load_msal_cache(record.id) is None
    assert store.load_pending_flow(record.id) is None
    assert store.load_msal_cache(other.id) == '{"other": true}'
    assert store.load_pending_flow(other.id)["user_code"] == "WXYZ"


def test_credential_store_migrates_legacy_graph_state(tmp_path) -> None:
    db_path = tmp_path / "legacy.sqlite3"
    conn = sqlite3.connect(db_path)
    conn.executescript(
        """
        CREATE TABLE http_credentials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            method TEXT NOT NULL,
            account TEXT NOT NULL,
            secret_hash TEXT NOT NULL,
            created_at REAL NOT NULL,
            updated_at REAL NOT NULL,
            UNIQUE(method, account, secret_hash)
        );
        CREATE TABLE graph_oauth_state (
            credential_id INTEGER PRIMARY KEY,
            msal_cache_json TEXT,
            pending_flow_json TEXT,
            pending_expires_at REAL
        );
        """
    )
    conn.execute(
        "INSERT INTO http_credentials(method, account, secret_hash, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
        ("graph", "person@outlook.com", hash_secret("test-secret-at-least-24-chars"), 1, 1),
    )
    conn.execute(
        "INSERT INTO graph_oauth_state(credential_id, msal_cache_json, pending_flow_json, pending_expires_at) VALUES (?, ?, ?, ?)",
        (
            1,
            json.dumps({"Account": {"home": {"home_account_id": "home-1"}}, "AccessToken": {"token": {}}}),
            json.dumps({"user_code": "ABCD", "expires_at": time.time() + 60}),
            time.time() + 60,
        ),
    )
    conn.commit()
    conn.close()

    store = CredentialStore(db_path)

    authorization_id = store.get_graph_authorization_id(1)
    assert authorization_id is not None
    assert store.load_msal_cache_for_authorization(authorization_id) is not None
    binding = store.load_graph_authorization_binding(authorization_id)
    assert binding is not None
    assert binding.home_account_id == "home-1"
    assert binding.status == "legacy_unverified"
    assert store.load_pending_flow_for_authorization(authorization_id)["user_code"] == "ABCD"
    with store._connect() as conn:
        assert conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'graph_oauth_state'"
        ).fetchone() is None


def test_new_credential_does_not_reuse_existing_authorization(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    config = GraphConfig(email="person@outlook.com", client_id="client-id")
    first = store.get_or_create_endpoint("graph", config.email, "test-secret-at-least-24-chars")
    first_authorization = store.resolve_graph_authorization(first.id, config.email, graph_config_fingerprint(config))
    store.save_graph_authorization_binding(
        first_authorization,
        GraphAuthorizationBinding(
            expected_account=config.email,
            home_account_id="home-1",
            graph_user_id="graph-user-1",
            identities=(config.email,),
            status="authorized",
        ),
    )

    second = store.get_or_create_endpoint("graph", config.email, "second-secret-at-least-24-chars")
    second_authorization = store.resolve_graph_authorization(
        second.id,
        config.email,
        graph_config_fingerprint(config),
    )

    assert second_authorization != first_authorization
    assert store.load_msal_cache_for_authorization(second_authorization) is None


def test_config_change_requires_new_authorization_and_removes_old_tokens(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    first_authorization = store.resolve_graph_authorization(record.id, record.account, "config-1")
    store.save_msal_cache_for_authorization(first_authorization, '{"AccessToken": {}}')

    second_authorization = store.resolve_graph_authorization(record.id, record.account, "config-2")

    assert second_authorization != first_authorization
    assert store.load_msal_cache_for_authorization(second_authorization) is None
    assert store.load_msal_cache_for_authorization(first_authorization) is None


def test_credential_store_rotates_secret_atomically_and_preserves_authorization(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    old_secret = "old"
    new_secret = "rotated-secret-at-least-24-chars"
    record = store.get_or_create_endpoint("graph", "person@outlook.com", old_secret)
    authorization_id = store.resolve_graph_authorization(record.id, record.account, "config-1")
    store.save_msal_cache_for_authorization(authorization_id, '{"AccessToken": {}}')

    rotated = store.rotate_endpoint("graph", record.account, old_secret, new_secret)

    assert rotated.id == record.id
    assert store.lookup_endpoint("graph", record.account, old_secret) is None
    assert store.lookup_endpoint("graph", record.account, new_secret) == rotated
    assert store.get_graph_authorization_id(rotated.id) == authorization_id
    assert store.load_msal_cache(rotated.id) == '{"AccessToken": {}}'


def test_credential_store_rotation_rejects_registered_target_without_partial_change(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    account = "person@outlook.com"
    old_secret = "first-secret-at-least-24-chars"
    existing_secret = "existing-secret-at-least-24-chars"
    old = store.get_or_create_endpoint("graph", account, old_secret)
    existing = store.get_or_create_endpoint("graph", account, existing_secret)

    with pytest.raises(ValueError, match="already registered"):
        store.rotate_endpoint("graph", account, old_secret, existing_secret)

    assert store.lookup_endpoint("graph", account, old_secret) == old
    assert store.lookup_endpoint("graph", account, existing_secret) == existing


def test_shared_authorization_migration_revokes_all_linked_credentials(tmp_path) -> None:
    db_path = tmp_path / "outlook.sqlite3"
    store = CredentialStore(db_path)
    first = store.get_or_create_endpoint("graph", "person@outlook.com", "first-secret-at-least-24-chars")
    second = store.get_or_create_endpoint("graph", "person@outlook.com", "second-secret-at-least-24-chars")
    authorization_id = store.resolve_graph_authorization(first.id, first.account, "config-1")
    store.save_msal_cache_for_authorization(authorization_id, '{"AccessToken": {}}')
    with store._connect() as conn:
        conn.execute(
            "UPDATE http_credentials SET graph_authorization_id = ? WHERE id = ?",
            (authorization_id, second.id),
        )
        conn.executescript(
            """
            CREATE TABLE graph_oauth_state (
                credential_id INTEGER PRIMARY KEY,
                msal_cache_json TEXT,
                pending_flow_json TEXT,
                pending_expires_at REAL,
                updated_at REAL NOT NULL
            );
            """
        )
        conn.execute(
            """
            INSERT INTO graph_oauth_state(
                credential_id, msal_cache_json, pending_flow_json,
                pending_expires_at, updated_at
            ) VALUES (?, ?, NULL, NULL, ?)
            """,
            (first.id, '{"AccessToken": {"legacy": {}}}', time.time()),
        )

    migrated = CredentialStore(db_path)

    assert migrated.get_graph_authorization_id(first.id) is None
    assert migrated.get_graph_authorization_id(second.id) is None
    with migrated._connect() as conn:
        assert conn.execute(
            "SELECT 1 FROM graph_authorizations WHERE id = ?",
            (authorization_id,),
        ).fetchone() is None
        assert conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'graph_oauth_state'"
        ).fetchone() is None

    reopened = CredentialStore(db_path)
    assert reopened.get_graph_authorization_id(first.id) is None
    assert reopened.get_graph_authorization_id(second.id) is None


def test_credential_store_filters_operation_logs(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    other = store.get_or_create_endpoint("graph", "other@outlook.com", "test-secret-at-least-24-chars")

    store.append_operation_log(
        record.id,
        operation="mail_list",
        http_method="post",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=10,
        created_at=100,
    )
    store.append_operation_log(
        record.id,
        operation="calendar_create",
        http_method="POST",
        status_code=400,
        ok=False,
        error="validation_error",
        duration_ms=20,
        created_at=200,
    )
    store.append_operation_log(
        other.id,
        operation="mail_delete",
        http_method="POST",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=30,
        created_at=200,
    )

    store.flush_logs()
    logs = store.list_operation_logs(record.id, start_at=150, end_at=250)

    assert len(logs) == 1
    assert logs[0]["operation"] == "calendar_create"
    assert logs[0]["http_method"] == "POST"
    assert logs[0]["ok"] is False
    assert logs[0]["error"] == "validation_error"


def test_credential_store_register_cli(tmp_path, capsys) -> None:
    db_path = tmp_path / "outlook.sqlite3"

    assert credential_store_main(
        ["register", "graph", "you@outlook.com", "cli-secret-at-least-24-chars", "--db", str(db_path)]
    ) == 0

    output = capsys.readouterr().out
    assert '"ok": true' in output
    store = CredentialStore(db_path)
    assert store.lookup_endpoint("graph", "you@outlook.com", "cli-secret-at-least-24-chars") is not None


def test_credential_store_register_cli_uses_configured_minimum(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SECRET_MIN_LENGTH", "8")
    db_path = tmp_path / "outlook.sqlite3"

    with pytest.raises(SystemExit):
        credential_store_main(
            ["register", "graph", "you@outlook.com", "abcd", "--db", str(db_path)]
        )

    store = CredentialStore(db_path)
    assert store.lookup_endpoint("graph", "you@outlook.com", "abcd") is None


def test_credential_store_register_cli_rejects_invalid_secret(tmp_path) -> None:
    db_path = tmp_path / "outlook.sqlite3"

    with pytest.raises(SystemExit):
        credential_store_main(["register", "graph", "you@outlook.com", "a/b", "--db", str(db_path)])

    store = CredentialStore(db_path)
    assert store.lookup_endpoint("graph", "you@outlook.com", "a/b") is None


def test_credential_store_rotate_cli_revokes_old_secret(tmp_path, capsys) -> None:
    db_path = tmp_path / "outlook.sqlite3"
    old_secret = "old-secret-at-least-24-characters"
    new_secret = "new-secret-at-least-24-characters"
    assert credential_store_main(
        ["register", "graph", "you@outlook.com", old_secret, "--db", str(db_path)]
    ) == 0
    capsys.readouterr()

    assert credential_store_main(
        ["rotate", "graph", "you@outlook.com", old_secret, new_secret, "--db", str(db_path)]
    ) == 0

    store = CredentialStore(db_path)
    assert store.lookup_endpoint("graph", "you@outlook.com", old_secret) is None
    assert store.lookup_endpoint("graph", "you@outlook.com", new_secret) is not None


def test_credential_store_wal_pragmas_are_set(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    with store._connect() as conn:
        journal = conn.execute("PRAGMA journal_mode").fetchone()[0]
        synchronous = conn.execute("PRAGMA synchronous").fetchone()[0]
        busy_timeout = conn.execute("PRAGMA busy_timeout").fetchone()[0]
    assert journal.upper() == "WAL"
    assert synchronous == 1  # NORMAL
    assert busy_timeout == 5000


def test_credential_store_flush_persists_buffered_entries(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.append_operation_log(
        record.id,
        operation="mail_list",
        http_method="POST",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=10,
    )
    store.flush_logs()
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 1
    assert logs[0]["operation"] == "mail_list"


def test_credential_store_prune_removes_old_logs(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    old_time = time.time() - 40 * 24 * 3600
    store.append_operation_log(
        record.id,
        operation="old_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
        created_at=old_time,
    )
    store.flush_logs()
    store._prune_old_logs()
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 0


def test_credential_store_auto_flushes_at_100_entry_boundary(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    for i in range(99):
        store.append_operation_log(
            record.id,
            operation=f"op_{i}",
            http_method="POST",
            status_code=200,
            ok=True,
            error=None,
            duration_ms=1,
        )
    assert len(store._log_buffer) == 99

    store.append_operation_log(
        record.id,
        operation="op_99",
        http_method="POST",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=1,
    )
    deadline = time.time() + 5
    logs = store.list_operation_logs(record.id)
    while len(logs) < 100 and time.time() < deadline:
        time.sleep(0.05)
        logs = store.list_operation_logs(record.id)
    assert len(logs) == 100
    assert len(store._log_buffer) == 0


def test_credential_store_flush_thread_is_daemon_and_registers_atexit(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    assert store._flush_thread is not None
    assert store._flush_thread.daemon is True
    assert store._flush_running is True

    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.append_operation_log(
        record.id,
        operation="pending_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
    )
    store._flush_exit()
    assert store._flush_running is False
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 1
    assert logs[0]["operation"] == "pending_op"


def test_credential_store_retention_boundary(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    now = time.time()
    store.append_operation_log(
        record.id,
        operation="kept_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
        created_at=now - 29 * 24 * 3600,
    )
    store.append_operation_log(
        record.id,
        operation="pruned_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
        created_at=now - 31 * 24 * 3600,
    )
    store.flush_logs()
    store._prune_old_logs()
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 1
    assert logs[0]["operation"] == "kept_op"


def test_credential_store_concurrent_append_does_not_race(tmp_path) -> None:
    import threading

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store._flush_exit()
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    errors: list[str] = []

    def worker(worker_id: int) -> None:
        try:
            for i in range(25):
                store.append_operation_log(
                    record.id,
                    operation=f"w{worker_id}_op{i}",
                    http_method="POST",
                    status_code=200,
                    ok=True,
                    error=None,
                    duration_ms=1,
                )
        except Exception as exc:
            errors.append(str(exc))

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert not errors
    store.flush_logs()
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 100


def test_credential_store_flush_exit_final_flush(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")
    store.append_operation_log(
        record.id,
        operation="exit_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
    )
    store._flush_exit()
    assert store._flush_running is False
    logs = store.list_operation_logs(record.id)
    assert len(logs) == 1
    assert logs[0]["operation"] == "exit_op"


def test_prune_index_created_at_exists(tmp_path) -> None:
    """Verify idx_http_operation_logs_created_at index exists and is used by prune query."""
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    with store._connect() as conn:
        rows = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='index' AND name='idx_http_operation_logs_created_at'"
        ).fetchall()
    assert len(rows) == 1
    assert rows[0]["name"] == "idx_http_operation_logs_created_at"


def _calendar_draft_values(*, source_event_id: str | None = None) -> dict:
    return {
        "draft_id": "draft-1",
        "expected_version": None,
        "draft_type": "update" if source_event_id else "create",
        "source_event_id": source_event_id,
        "payload": {
            "subject": "跨设备草稿",
            "start": "2026-07-16T09:00",
            "end": "2026-07-16T10:00",
        },
        "base_snapshot": {"id": source_event_id, "changekey": "base-1"} if source_event_id else None,
        "base_fingerprint": "fingerprint-1" if source_event_id else None,
        "subject": "跨设备草稿",
        "is_all_day": False,
        "start_at": 1_768_529_600.0,
        "end_at": 1_768_533_200.0,
        "start_date": None,
        "end_date": None,
        "timezone": "Asia/Shanghai",
    }


def test_calendar_draft_persists_across_store_instances_and_secret_rotation(tmp_path) -> None:
    db_path = tmp_path / "outlook.sqlite3"
    store = CredentialStore(db_path)
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )

    created = store.save_calendar_draft(record.id, **_calendar_draft_values())
    assert created.version == 1
    assert created.status == "active"
    assert created.expires_at is None

    reopened = CredentialStore(db_path)
    loaded = reopened.get_calendar_draft(record.id, created.id)
    assert loaded is not None
    assert loaded.payload["subject"] == "跨设备草稿"

    rotated = reopened.rotate_endpoint(
        "graph",
        "person@outlook.com",
        "test-secret-at-least-24-chars",
        "replacement-secret-at-least-24-chars",
    )
    assert rotated.id == record.id
    assert reopened.get_calendar_draft(rotated.id, created.id) is not None


def test_calendar_draft_uses_optimistic_version_and_one_update_per_source(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    values = _calendar_draft_values(source_event_id="event-1")
    created = store.save_calendar_draft(record.id, **values)

    updated_values = dict(values)
    updated_values["expected_version"] = created.version
    updated_values["subject"] = "第二版"
    updated_values["payload"] = {**values["payload"], "subject": "第二版"}
    updated = store.save_calendar_draft(record.id, **updated_values)
    assert updated.version == 2
    assert updated.subject == "第二版"

    with pytest.raises(CalendarDraftConflictError) as stale:
        store.save_calendar_draft(record.id, **updated_values)
    assert stale.value.record.version == 2

    duplicate_values = _calendar_draft_values(source_event_id="event-1")
    duplicate_values["draft_id"] = "draft-2"
    with pytest.raises(CalendarDraftConflictError) as duplicate:
        store.save_calendar_draft(record.id, **duplicate_values)
    assert duplicate.value.record.id == "draft-1"


def test_calendar_draft_cross_store_concurrent_writes_have_one_winner(tmp_path) -> None:
    import threading

    db_path = tmp_path / "outlook.sqlite3"
    first_store = CredentialStore(db_path)
    record = first_store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    created = first_store.save_calendar_draft(record.id, **_calendar_draft_values())
    second_store = CredentialStore(db_path)
    barrier = threading.Barrier(2)
    outcomes: list[str] = []

    def write(store: CredentialStore, subject: str) -> None:
        values = _calendar_draft_values()
        values["expected_version"] = created.version
        values["subject"] = subject
        values["payload"] = {**values["payload"], "subject": subject}
        barrier.wait()
        try:
            store.save_calendar_draft(record.id, **values)
            outcomes.append("saved")
        except CalendarDraftConflictError:
            outcomes.append("conflict")

    threads = [
        threading.Thread(target=write, args=(first_store, "设备一")),
        threading.Thread(target=write, args=(second_store, "设备二")),
    ]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert sorted(outcomes) == ["conflict", "saved"]
    assert first_store.get_calendar_draft(record.id, created.id).version == 2


def test_calendar_draft_retention_recomputes_and_prunes(monkeypatch, tmp_path) -> None:
    now = 1_800_000_000.0
    monkeypatch.setattr("outlook_mcp.credential_store.time.time", lambda: now)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    store.set_calendar_draft_retention(record.id, 1)
    created = store.save_calendar_draft(record.id, **_calendar_draft_values())

    settings = store.set_calendar_draft_retention(record.id, 30)
    assert settings["expired_count"] == 0
    retained = store.get_calendar_draft(record.id, created.id)
    assert retained.expires_at == now + 30 * 86400

    now += 31 * 86400
    assert store.get_calendar_draft(record.id, created.id) is None


def test_calendar_draft_submission_state_blocks_reentry_and_recovers_unknown(
    monkeypatch,
    tmp_path,
) -> None:
    now = 1_800_000_000.0
    monkeypatch.setattr("outlook_mcp.credential_store.time.time", lambda: now)
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    record = store.get_or_create_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars"
    )
    store.set_calendar_draft_retention(record.id, 1)
    created = store.save_calendar_draft(record.id, **_calendar_draft_values())
    submitting = store.begin_calendar_draft_submission(record.id, created.id, created.version)
    assert submitting.status == "submitting"

    with pytest.raises(CalendarDraftStateError):
        store.begin_calendar_draft_submission(record.id, created.id, submitting.version)

    now += 86401
    recovered = store.get_calendar_draft(record.id, created.id)
    assert recovered.status == "outcome_unknown"
    assert recovered.last_error_code == "outcome_unknown"
    assert recovered.expires_at == now + 86400




def test_credential_store_db_file_has_private_permissions(tmp_path) -> None:
    if os.name == "nt":
        pytest.skip("chmod semantics are weak on Windows")
    import stat

    db_path = tmp_path / "outlook.sqlite3"
    CredentialStore(db_path)
    assert stat.S_IMODE(os.stat(db_path).st_mode) == 0o600
    assert stat.S_IMODE(os.stat(tmp_path).st_mode) == 0o700


def test_credential_store_flush_loop_survives_flush_errors(monkeypatch, tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    store._LOG_FLUSH_INTERVAL = 0.05
    record = store.get_or_create_endpoint("graph", "person@outlook.com", "test-secret-at-least-24-chars")

    original_flush = CredentialStore.flush_logs
    failures = {"count": 0}

    def flaky_flush(self: CredentialStore) -> None:
        if self is store and failures["count"] < 1:
            failures["count"] += 1
            raise RuntimeError("simulated flush failure")
        return original_flush(self)

    monkeypatch.setattr(CredentialStore, "flush_logs", flaky_flush)
    store.append_operation_log(
        record.id,
        operation="resilient_op",
        http_method="GET",
        status_code=200,
        ok=True,
        error=None,
        duration_ms=5,
    )
    store._log_buffer_event.set()

    deadline = time.time() + 5
    logs = store.list_operation_logs(record.id)
    while not logs and time.time() < deadline:
        time.sleep(0.05)
        logs = store.list_operation_logs(record.id)

    assert failures["count"] == 1
    assert [log["operation"] for log in logs] == ["resilient_op"]
    assert store._flush_thread is not None and store._flush_thread.is_alive()


def test_credential_store_close_is_idempotent_and_stops_flush_thread(tmp_path) -> None:
    store = CredentialStore(tmp_path / "outlook.sqlite3")
    assert store._flush_thread is not None and store._flush_thread.is_alive()

    store.close()

    assert not store._flush_thread.is_alive()
    store.close()


def test_module_level_register_endpoint_closes_its_store(monkeypatch, tmp_path) -> None:
    import outlook_mcp.credential_store as credential_store_module

    created: list[CredentialStore] = []
    real_init = CredentialStore.__init__

    def tracking_init(self: CredentialStore, path: object = None) -> None:
        created.append(self)
        real_init(self, path if path is not None else tmp_path / "outlook.sqlite3")

    monkeypatch.setattr(CredentialStore, "__init__", tracking_init)

    credential_store_module.register_endpoint(
        "graph", "person@outlook.com", "test-secret-at-least-24-chars", db_path=tmp_path / "outlook.sqlite3"
    )

    assert created and created[0]._closed is True


@pytest.mark.skipif(os.name == "nt", reason="POSIX permissions")
def test_credential_store_restricts_db_file_permissions(tmp_path) -> None:
    import stat

    store = CredentialStore(tmp_path / "outlook.sqlite3")
    try:
        assert stat.S_IMODE(os.stat(tmp_path / "outlook.sqlite3").st_mode) == 0o600
        assert stat.S_IMODE(os.stat(tmp_path).st_mode) == 0o700
    finally:
        store.close()
