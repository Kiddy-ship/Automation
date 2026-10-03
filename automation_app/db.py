"""Small, thread-safe SQLite layer with idempotent migrations."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterator, Mapping, Sequence


SCHEMA_VERSION = 5


MIGRATIONS: dict[int, str] = {
    1: r"""
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_key TEXT NOT NULL UNIQUE,
            business_date TEXT NOT NULL,
            source_path TEXT NOT NULL,
            source_name TEXT NOT NULL,
            source_sha256 TEXT NOT NULL,
            source_size INTEGER NOT NULL,
            source_mtime_ns INTEGER NOT NULL,
            parser_metadata_json TEXT NOT NULL DEFAULT '{}',
            status TEXT NOT NULL,
            order_count INTEGER NOT NULL DEFAULT 0,
            line_count INTEGER NOT NULL DEFAULT 0,
            duplicate_count INTEGER NOT NULL DEFAULT 0,
            duplicate_file INTEGER NOT NULL DEFAULT 0 CHECK (duplicate_file IN (0, 1)),
            duplicates_allowed INTEGER NOT NULL DEFAULT 0 CHECK (duplicates_allowed IN (0, 1)),
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            completed_at TEXT,
            error_message TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_runs_business_date
            ON runs (business_date, created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_runs_source_hash
            ON runs (business_date, source_sha256);
        CREATE INDEX IF NOT EXISTS idx_runs_status
            ON runs (status, updated_at DESC);

        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
            sequence INTEGER NOT NULL,
            order_code TEXT NOT NULL,
            payload_sha256 TEXT NOT NULL,
            source_start_row INTEGER,
            source_end_row INTEGER,
            customer_code TEXT NOT NULL DEFAULT '',
            customer_name TEXT NOT NULL DEFAULT '',
            customer_address TEXT NOT NULL DEFAULT '',
            order_date TEXT,
            status TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE (run_id, order_code),
            UNIQUE (run_id, sequence)
        );

        CREATE INDEX IF NOT EXISTS idx_orders_business_key
            ON orders (order_code, run_id);
        CREATE INDEX IF NOT EXISTS idx_orders_status
            ON orders (run_id, status, sequence);

        CREATE TABLE IF NOT EXISTS line_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
            sequence INTEGER NOT NULL,
            source_row INTEGER,
            product_code TEXT NOT NULL,
            product_name TEXT NOT NULL DEFAULT '',
            quantity TEXT NOT NULL,
            unit_price TEXT NOT NULL,
            gross_amount TEXT NOT NULL,
            discount_percent TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            UNIQUE (order_id, sequence)
        );

        CREATE TABLE IF NOT EXISTS batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
            batch_number INTEGER NOT NULL,
            workbook_path TEXT,
            workbook_sha256 TEXT,
            sheet_count INTEGER NOT NULL CHECK (sheet_count BETWEEN 1 AND 19),
            line_count INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            imported_at TEXT,
            review_ready_at TEXT,
            user_confirmed_at TEXT,
            error_message TEXT,
            UNIQUE (run_id, batch_number)
        );

        CREATE INDEX IF NOT EXISTS idx_batches_status
            ON batches (run_id, status, batch_number);

        CREATE TABLE IF NOT EXISTS batch_orders (
            batch_id INTEGER NOT NULL REFERENCES batches(id) ON DELETE CASCADE,
            order_id INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
            sheet_index INTEGER NOT NULL,
            sheet_name TEXT,
            expected_tab_label TEXT,
            actual_tab_label TEXT,
            note_verified INTEGER NOT NULL DEFAULT 0 CHECK (note_verified IN (0, 1)),
            status TEXT NOT NULL DEFAULT 'pending',
            updated_at TEXT NOT NULL,
            PRIMARY KEY (batch_id, sheet_index),
            UNIQUE (batch_id, order_id)
        );

        CREATE TABLE IF NOT EXISTS duplicate_decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
            order_id INTEGER NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
            prior_run_id INTEGER REFERENCES runs(id) ON DELETE SET NULL,
            prior_order_id INTEGER REFERENCES orders(id) ON DELETE SET NULL,
            duplicate_kind TEXT NOT NULL,
            decision TEXT NOT NULL,
            decided_at TEXT,
            created_at TEXT NOT NULL,
            UNIQUE (run_id, order_id, prior_order_id)
        );

        CREATE INDEX IF NOT EXISTS idx_duplicate_decisions_run
            ON duplicate_decisions (run_id, decision);

        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER REFERENCES runs(id) ON DELETE CASCADE,
            batch_id INTEGER REFERENCES batches(id) ON DELETE CASCADE,
            order_id INTEGER REFERENCES orders(id) ON DELETE SET NULL,
            event_type TEXT NOT NULL,
            level TEXT NOT NULL DEFAULT 'info',
            message TEXT NOT NULL DEFAULT '',
            payload_json TEXT NOT NULL DEFAULT '{}',
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_events_run
            ON events (run_id, id DESC);
        CREATE INDEX IF NOT EXISTS idx_events_batch
            ON events (batch_id, id DESC);
    """,
    2: r"""
        -- A row in this ledger is created only after a person confirms that
        -- every final KiotViet order in the batch was handled.  It is the
        -- sole source of truth for automatic same-day de-duplication.
        CREATE TABLE IF NOT EXISTS confirmed_order_ledger (
            business_date TEXT NOT NULL,
            order_code TEXT NOT NULL,
            payload_sha256 TEXT NOT NULL,
            source_order_id INTEGER NOT NULL REFERENCES orders(id),
            run_id INTEGER NOT NULL REFERENCES runs(id),
            batch_id INTEGER NOT NULL REFERENCES batches(id),
            confirmation_method TEXT NOT NULL
                CHECK (confirmation_method = 'manual_user_confirmation'),
            confirmed_at TEXT NOT NULL,
            PRIMARY KEY (business_date, order_code)
        );

        CREATE INDEX IF NOT EXISTS idx_confirmed_order_ledger_date
            ON confirmed_order_ledger (business_date, confirmed_at DESC);

        -- Backfill only previously *confirmed* batches.  A batch which was
        -- merely imported or is awaiting a review must never be treated as
        -- safe to skip.
        INSERT INTO confirmed_order_ledger(
            business_date, order_code, payload_sha256, source_order_id,
            run_id, batch_id, confirmation_method, confirmed_at
        )
        SELECT r.business_date, o.order_code, o.payload_sha256, o.id,
               r.id, b.id, 'manual_user_confirmation',
               COALESCE(b.user_confirmed_at, r.completed_at, r.updated_at)
        FROM batches b
        JOIN runs r ON r.id = b.run_id
        JOIN batch_orders bo ON bo.batch_id = b.id
        JOIN orders o ON o.id = bo.order_id
        WHERE b.status = 'confirmed' AND o.status = 'user_confirmed';

        CREATE TABLE IF NOT EXISTS report_scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_key TEXT NOT NULL UNIQUE,
            business_date TEXT NOT NULL,
            trigger TEXT NOT NULL DEFAULT 'manual' CHECK (trigger = 'manual'),
            source_path TEXT NOT NULL,
            source_name TEXT NOT NULL,
            source_sha256 TEXT NOT NULL,
            source_size INTEGER NOT NULL,
            source_mtime_ns INTEGER NOT NULL,
            source_order_count INTEGER NOT NULL,
            source_line_count INTEGER NOT NULL,
            new_order_count INTEGER NOT NULL,
            new_line_count INTEGER NOT NULL,
            skipped_confirmed_count INTEGER NOT NULL,
            unresolved_count INTEGER NOT NULL,
            changed_confirmed_count INTEGER NOT NULL DEFAULT 0,
            audit_path TEXT,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL,
            completed_at TEXT,
            error_message TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_report_scans_date
            ON report_scans (business_date, id DESC);

        CREATE TABLE IF NOT EXISTS scan_order_decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id INTEGER NOT NULL REFERENCES report_scans(id) ON DELETE CASCADE,
            order_code TEXT NOT NULL,
            source_sequence INTEGER NOT NULL,
            source_start_row INTEGER,
            source_end_row INTEGER,
            payload_sha256 TEXT NOT NULL,
            decision TEXT NOT NULL
                CHECK (decision IN ('new', 'skipped_confirmed', 'unresolved', 'changed_confirmed')),
            prior_order_id INTEGER REFERENCES orders(id),
            prior_run_id INTEGER REFERENCES runs(id),
            prior_batch_id INTEGER REFERENCES batches(id),
            reason TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            UNIQUE (scan_id, order_code)
        );

        CREATE INDEX IF NOT EXISTS idx_scan_order_decisions_scan
            ON scan_order_decisions (scan_id, decision, source_sequence);

    """,
    3: r"""
        -- The direct Edge controller can now verify the final KiotViet
        -- submission itself.  Preserve manual confirmations while recording
        -- the distinct, machine-verifiable completion method.
        ALTER TABLE batches ADD COLUMN automation_confirmed_at TEXT;

        -- SQLite cannot extend a CHECK constraint in place.  Rebuild this
        -- small audit table without changing any existing ledger rows.
        DROP INDEX IF EXISTS idx_confirmed_order_ledger_date;
        ALTER TABLE confirmed_order_ledger RENAME TO confirmed_order_ledger_legacy;

        CREATE TABLE confirmed_order_ledger (
            business_date TEXT NOT NULL,
            order_code TEXT NOT NULL,
            payload_sha256 TEXT NOT NULL,
            source_order_id INTEGER NOT NULL REFERENCES orders(id),
            run_id INTEGER NOT NULL REFERENCES runs(id),
            batch_id INTEGER NOT NULL REFERENCES batches(id),
            confirmation_method TEXT NOT NULL
                CHECK (confirmation_method IN (
                    'manual_user_confirmation',
                    'automation_verified_submission'
                )),
            confirmed_at TEXT NOT NULL,
            PRIMARY KEY (business_date, order_code)
        );

        INSERT INTO confirmed_order_ledger(
            business_date, order_code, payload_sha256, source_order_id,
            run_id, batch_id, confirmation_method, confirmed_at
        )
        SELECT business_date, order_code, payload_sha256, source_order_id,
               run_id, batch_id, confirmation_method, confirmed_at
        FROM confirmed_order_ledger_legacy;

        DROP TABLE confirmed_order_ledger_legacy;

        CREATE INDEX IF NOT EXISTS idx_confirmed_order_ledger_date
            ON confirmed_order_ledger (business_date, confirmed_at DESC);
    """,
    4: r"""
        -- A user-selected historical report is allowed to remain active
        -- after midnight.  Older unfinished runs keep the previous
        -- fail-safe behaviour and are still expired when a new day starts.
        ALTER TABLE runs ADD COLUMN historical_run INTEGER NOT NULL DEFAULT 0
            CHECK (historical_run IN (0, 1));
    """,
    5: r"""
        -- Scheduled scans are first-class audit records.  SQLite cannot
        -- alter a CHECK constraint in place, and scan_order_decisions has a
        -- foreign key to this table, so rebuild the two small audit tables
        -- together while preserving every existing id and decision.
        DROP INDEX IF EXISTS idx_scan_order_decisions_scan;
        DROP INDEX IF EXISTS idx_report_scans_date;
        ALTER TABLE scan_order_decisions RENAME TO scan_order_decisions_legacy;
        ALTER TABLE report_scans RENAME TO report_scans_legacy;

        CREATE TABLE report_scans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_key TEXT NOT NULL UNIQUE,
            business_date TEXT NOT NULL,
            trigger TEXT NOT NULL DEFAULT 'manual'
                CHECK (trigger IN ('manual', 'scheduled')),
            source_path TEXT NOT NULL,
            source_name TEXT NOT NULL,
            source_sha256 TEXT NOT NULL,
            source_size INTEGER NOT NULL,
            source_mtime_ns INTEGER NOT NULL,
            source_order_count INTEGER NOT NULL,
            source_line_count INTEGER NOT NULL,
            new_order_count INTEGER NOT NULL,
            new_line_count INTEGER NOT NULL,
            skipped_confirmed_count INTEGER NOT NULL,
            unresolved_count INTEGER NOT NULL,
            changed_confirmed_count INTEGER NOT NULL DEFAULT 0,
            audit_path TEXT,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL,
            completed_at TEXT,
            error_message TEXT
        );

        INSERT INTO report_scans(
            id, scan_key, business_date, trigger, source_path, source_name,
            source_sha256, source_size, source_mtime_ns, source_order_count,
            source_line_count, new_order_count, new_line_count,
            skipped_confirmed_count, unresolved_count, changed_confirmed_count,
            audit_path, status, created_at, completed_at, error_message
        )
        SELECT id, scan_key, business_date, trigger, source_path, source_name,
               source_sha256, source_size, source_mtime_ns, source_order_count,
               source_line_count, new_order_count, new_line_count,
               skipped_confirmed_count, unresolved_count, changed_confirmed_count,
               audit_path, status, created_at, completed_at, error_message
        FROM report_scans_legacy;

        CREATE TABLE scan_order_decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scan_id INTEGER NOT NULL REFERENCES report_scans(id) ON DELETE CASCADE,
            order_code TEXT NOT NULL,
            source_sequence INTEGER NOT NULL,
            source_start_row INTEGER,
            source_end_row INTEGER,
            payload_sha256 TEXT NOT NULL,
            decision TEXT NOT NULL
                CHECK (decision IN ('new', 'skipped_confirmed', 'unresolved', 'changed_confirmed')),
            prior_order_id INTEGER REFERENCES orders(id),
            prior_run_id INTEGER REFERENCES runs(id),
            prior_batch_id INTEGER REFERENCES batches(id),
            reason TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            UNIQUE (scan_id, order_code)
        );

        INSERT INTO scan_order_decisions(
            id, scan_id, order_code, source_sequence, source_start_row,
            source_end_row, payload_sha256, decision, prior_order_id,
            prior_run_id, prior_batch_id, reason, created_at
        )
        SELECT id, scan_id, order_code, source_sequence, source_start_row,
               source_end_row, payload_sha256, decision, prior_order_id,
               prior_run_id, prior_batch_id, reason, created_at
        FROM scan_order_decisions_legacy;

        DROP TABLE scan_order_decisions_legacy;
        DROP TABLE report_scans_legacy;
        CREATE INDEX idx_report_scans_date
            ON report_scans (business_date, id DESC);
        CREATE INDEX idx_scan_order_decisions_scan
            ON scan_order_decisions (scan_id, decision, source_sequence);
    """,
}


def json_default(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, Path):
        return str(value)
    if hasattr(value, "to_dict"):
        return value.to_dict()
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def json_dumps(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=json_default,
    )


class Database:
    """Connection-per-operation SQLite wrapper.

    Connections never cross thread boundaries.  WAL mode permits the Flask UI
    to read progress while a worker records state transitions.
    """

    def __init__(self, path: str | Path, *, busy_timeout_ms: int = 15_000) -> None:
        self.path = Path(path).expanduser().resolve()
        self.busy_timeout_ms = int(busy_timeout_ms)

    def connect(self) -> sqlite3.Connection:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(
            self.path,
            timeout=max(self.busy_timeout_ms / 1000, 1),
            isolation_level=None,
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute(f"PRAGMA busy_timeout = {self.busy_timeout_ms}")
        connection.execute("PRAGMA journal_mode = WAL")
        connection.execute("PRAGMA synchronous = NORMAL")
        return connection

    def initialize(self) -> None:
        """Apply each migration exactly once; safe under concurrent startup."""

        with self.transaction(immediate=True) as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS schema_migrations (
                    version INTEGER PRIMARY KEY,
                    applied_at TEXT NOT NULL
                )
                """
            )
            applied = {
                int(row["version"])
                for row in connection.execute(
                    "SELECT version FROM schema_migrations"
                ).fetchall()
            }
            for version in sorted(MIGRATIONS):
                if version in applied:
                    continue
                connection.executescript(MIGRATIONS[version])
                connection.execute(
                    "INSERT OR IGNORE INTO schema_migrations(version, applied_at) VALUES (?, ?)",
                    (version, datetime.now().astimezone().isoformat(timespec="seconds")),
                )

        current = self.scalar("SELECT MAX(version) FROM schema_migrations") or 0
        if int(current) != SCHEMA_VERSION:
            raise RuntimeError(
                f"Unsupported database schema {current}; expected {SCHEMA_VERSION}"
            )

    @contextmanager
    def transaction(self, *, immediate: bool = False) -> Iterator[sqlite3.Connection]:
        connection = self.connect()
        try:
            connection.execute("BEGIN IMMEDIATE" if immediate else "BEGIN")
            yield connection
            # ``executescript`` may commit implicitly on SQLite.  Avoid a
            # second COMMIT against a connection that is already out of a
            # transaction while preserving normal transactional behavior.
            if connection.in_transaction:
                connection.execute("COMMIT")
        except Exception:
            if connection.in_transaction:
                connection.execute("ROLLBACK")
            raise
        finally:
            connection.close()

    def execute(self, sql: str, parameters: Sequence[Any] = ()) -> int:
        with self.transaction(immediate=True) as connection:
            cursor = connection.execute(sql, parameters)
            return int(cursor.lastrowid or 0)

    def executemany(self, sql: str, rows: Sequence[Sequence[Any]]) -> None:
        with self.transaction(immediate=True) as connection:
            connection.executemany(sql, rows)

    def fetch_one(
        self, sql: str, parameters: Sequence[Any] = ()
    ) -> dict[str, Any] | None:
        connection = self.connect()
        try:
            row = connection.execute(sql, parameters).fetchone()
            return dict(row) if row is not None else None
        finally:
            connection.close()

    def fetch_all(
        self, sql: str, parameters: Sequence[Any] = ()
    ) -> list[dict[str, Any]]:
        connection = self.connect()
        try:
            return [dict(row) for row in connection.execute(sql, parameters).fetchall()]
        finally:
            connection.close()

    def scalar(self, sql: str, parameters: Sequence[Any] = ()) -> Any:
        connection = self.connect()
        try:
            row = connection.execute(sql, parameters).fetchone()
            return None if row is None else row[0]
        finally:
            connection.close()

    def integrity_check(self) -> str:
        return str(self.scalar("PRAGMA integrity_check"))


def row_dict(row: sqlite3.Row | Mapping[str, Any] | None) -> dict[str, Any] | None:
    if row is None:
        return None
    return dict(row)


__all__ = [
    "Database",
    "SCHEMA_VERSION",
    "json_default",
    "json_dumps",
    "row_dict",
]
