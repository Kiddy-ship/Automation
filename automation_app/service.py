"""Transactional workflow service for the local automation dashboard.

The service never controls Edge itself.  It records either a manual
confirmation or an independently verified result returned by the direct Edge
controller after KiotViet has accepted every final order in a batch.
"""

from __future__ import annotations

import hashlib
import inspect
import json
import re
import sqlite3
import unicodedata
import uuid
from dataclasses import asdict, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

from .config import AppConfig, BANGKOK_TZ
from .db import Database, json_dumps
from .workbook_builder import (
    BuildResult,
    WorkbookBuildError,
    WorkbookBuilder,
    sha256_file,
)
from .report_parser import SUPPORTED_REPORT_EXTENSIONS, SUPPORTED_REPORT_FORMATS


ACTIVE_RUN_STATUSES = {
    "created",
    "parsed",
    "duplicate_review",
    "preparing_batches",
    "batches_ready",
    "batch_in_progress",
    "waiting_user_confirmation",
    "automation_attention",
}
TERMINAL_RUN_STATUSES = {"completed", "cancelled", "failed", "expired"}

MUTATING_EXTENSION_EVENTS = {
    "batch.import_started",
    "batch.imported",
    "order.note_filled",
    "batch.notes_filled",
    "batch.review_ready",
    "batch.waiting_user_confirmation",
    "batch.error",
}

PROHIBITED_EVENT_FRAGMENTS = {
    "auto_submit",
    "autosubmit",
    "order.submit",
    "order.placed",
    "final_order_click",
    "click_order_button",
}

MANUAL_CONFIRMATION_METHOD = "manual_user_confirmation"
AUTOMATION_CONFIRMATION_METHOD = "automation_verified_submission"


class WorkflowError(RuntimeError):
    """Base user-facing workflow error."""

    code = "workflow_error"

    def __init__(self, message: str, *, details: Mapping[str, Any] | None = None):
        super().__init__(message)
        self.message = message
        self.details = dict(details or {})

    def to_dict(self) -> dict[str, Any]:
        return {"ok": False, "error": self.code, "message": self.message, "details": self.details}


class ActiveRunError(WorkflowError):
    code = "active_run_exists"


class DuplicateReviewRequired(WorkflowError):
    code = "duplicate_review_required"


class InvalidTransitionError(WorkflowError):
    code = "invalid_transition"


class MidnightGuardError(WorkflowError):
    code = "business_date_expired"


class SourceValidationError(WorkflowError):
    code = "invalid_source"


def _readable_json(raw: str | None, fallback: Any) -> Any:
    if not raw:
        return fallback
    try:
        return json.loads(raw)
    except (TypeError, json.JSONDecodeError):
        return fallback


def _mapping(value: Any) -> dict[str, Any]:
    if isinstance(value, Mapping):
        return dict(value)
    if hasattr(value, "to_dict"):
        result = value.to_dict()
        if isinstance(result, Mapping):
            return dict(result)
    if is_dataclass(value):
        return asdict(value)
    if hasattr(value, "__dict__"):
        return dict(vars(value))
    raise SourceValidationError(f"Unsupported parsed object: {type(value).__name__}")


def _date_value(value: Any, *, field: str) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for date_format in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            pass
    raise SourceValidationError(f"Invalid {field}: {text or '(blank)'}")


def _decimal_text(value: Any) -> str:
    if isinstance(value, Decimal):
        return format(value, "f")
    if value is None or value == "":
        return "0"
    try:
        return format(Decimal(str(value)), "f")
    except Exception as exc:
        raise SourceValidationError(f"Invalid numeric value: {value}") from exc


def _json_number(value: Any) -> int | float:
    number = Decimal(_decimal_text(value))
    if number == number.to_integral_value():
        return int(number)
    return float(number)


def _optional_json_number(value: Any) -> int | float | None:
    """Return a JSON-safe number only when the source actually supplied one."""

    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        return _json_number(value)
    except SourceValidationError:
        return None


def _payment_amount_from_payload(payload: Any) -> int | float | None:
    """Pick the report total that KiotViet displays as the order amount due."""

    if not isinstance(payload, Mapping):
        return None
    for field in ("payment_amount", "calculated_net_total", "revenue_ex_vat"):
        amount = _optional_json_number(payload.get(field))
        if amount is not None:
            return amount
    return None


def _line_net_amount(gross_amount: Any, discount_percent: Any) -> int | float | None:
    """Calculate the post-line-discount value visible on a KiotViet cart row."""

    try:
        gross = Decimal(_decimal_text(gross_amount))
        discount = Decimal(_decimal_text(discount_percent))
    except SourceValidationError:
        return None
    return _json_number(gross * (Decimal("100") - discount) / Decimal("100"))


class AutomationService:
    """Persistence-backed application service used by Flask route handlers."""

    def __init__(
        self,
        config: AppConfig | None = None,
        *,
        database: Database | None = None,
        parser: Callable[..., Any] | None = None,
        workbook_builder: WorkbookBuilder | None = None,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.config = config or AppConfig()
        self.config.ensure_directories()
        self.db = database or Database(
            self.config.database_path,
            busy_timeout_ms=self.config.sqlite_busy_timeout_ms,
        )
        self.db.initialize()
        self._parser = parser
        self.workbook_builder = workbook_builder or WorkbookBuilder(self.config)
        self._clock = clock

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def state(self, run_id: int | None = None) -> dict[str, Any]:
        """Return a JSON-safe snapshot for Flask/API/UI use."""

        active = self._active_run_row()
        if run_id is None:
            run = active or self.db.fetch_one(
                "SELECT * FROM runs ORDER BY id DESC LIMIT 1"
            )
        else:
            run = self.db.fetch_one("SELECT * FROM runs WHERE id = ?", (run_id,))

        base: dict[str, Any] = {
            "ok": True,
            # Final clicks remain guarded by exact tab/note/result checks, but
            # the direct Edge controller is now authorized to complete them.
            "safe_mode": False,
            "auto_submit_enabled": True,
            "automation_mode": "verified_auto_submit",
            "expected_branch": self.config.expected_branch,
            "today": self._today().isoformat(),
            "max_orders_per_batch": self.config.max_orders_per_batch,
            "phase": "idle",
            "can_start": active is None,
            "run": None,
            "last_scan": self.scan_state(),
        }
        if run is None:
            return base

        run_state = self._run_state(int(run["id"]))
        base["run"] = run_state
        base["phase"] = str(run_state["status"])
        base["can_start"] = active is None
        base["requires_duplicate_confirmation"] = (
            run_state["status"] == "duplicate_review"
        )
        base["waiting_user_confirmation"] = (
            run_state["status"] == "waiting_user_confirmation"
        )
        return base

    def ledger_orders(self, business_date: str | date | None = None) -> dict[str, Any]:
        """List the order de-duplication ledger for the dashboard.

        This is intentionally a separate read model from ``state()``: the
        ledger can contain orders from several completed runs, while the main
        workflow view only exposes the latest run.
        """

        day = self._today() if business_date in (None, "") else _date_value(
            business_date, field="business_date"
        )
        rows = self.db.fetch_all(
            """
            SELECT l.business_date, l.order_code, l.confirmed_at,
                   l.confirmation_method, l.run_id, l.batch_id,
                   l.source_order_id, r.status AS run_status,
                   r.source_name, b.batch_number, b.status AS batch_status,
                   o.customer_code, o.customer_name, o.order_date,
                   o.status AS order_status
            FROM confirmed_order_ledger l
            LEFT JOIN runs r ON r.id = l.run_id
            LEFT JOIN batches b ON b.id = l.batch_id
            LEFT JOIN orders o ON o.id = l.source_order_id
            WHERE l.business_date = ?
            ORDER BY l.confirmed_at DESC, l.order_code
            """,
            (day.isoformat(),),
        )
        return {
            "ok": True,
            "business_date": day.isoformat(),
            "count": len(rows),
            "orders": rows,
        }

    def delete_selected_ledger_orders(
        self,
        business_date: str | date,
        order_codes: Sequence[str],
        *,
        confirm: bool = False,
        acknowledge_external_order_risk: bool = False,
    ) -> dict[str, Any]:
        """Forget explicitly selected confirmed codes so a later scan can retry.

        This is a deliberately separate, stronger action than deleting a
        failed run.  It removes the de-duplication ledger and persisted source
        rows, but never touches KiotViet.  The UI therefore requires both a
        checkbox selection and an explicit acknowledgement that rerunning may
        create a duplicate external order.
        """

        if confirm is not True or acknowledge_external_order_risk is not True:
            raise WorkflowError(
                "Xóa mã đã xác nhận cần tích chọn và xác nhận rõ rủi ro chạy lại đơn trên KiotViet."
            )
        day = _date_value(business_date, field="business_date")
        if not isinstance(order_codes, Sequence) or isinstance(order_codes, (str, bytes)):
            raise WorkflowError("Danh sách mã cần xóa không hợp lệ.")
        selected = sorted({str(code or "").strip() for code in order_codes if str(code or "").strip()})
        if not selected:
            raise WorkflowError("Hãy tích chọn ít nhất một mã trước khi xóa.")

        placeholders = ",".join("?" for _ in selected)
        now = self._iso_now()
        affected_runs: set[int] = set()
        affected_batches: set[int] = set()
        deleted_order_ids: list[int] = []
        with self.db.transaction(immediate=True) as connection:
            active_placeholders = ",".join("?" for _ in ACTIVE_RUN_STATUSES)
            active = connection.execute(
                f"SELECT id, business_date, status FROM runs WHERE status IN ({active_placeholders}) LIMIT 1",
                tuple(sorted(ACTIVE_RUN_STATUSES)),
            ).fetchone()
            if active is not None:
                raise InvalidTransitionError(
                    "Không thể xóa ledger khi còn workflow đang chạy.",
                    details={"run_id": active["id"], "status": active["status"]},
                )

            ledger_rows = connection.execute(
                f"""
                SELECT business_date, order_code, source_order_id, run_id, batch_id
                FROM confirmed_order_ledger
                WHERE business_date = ? AND order_code IN ({placeholders})
                """,
                (day.isoformat(), *selected),
            ).fetchall()
            found = {str(row["order_code"]) for row in ledger_rows}
            missing = [code for code in selected if code not in found]
            if missing:
                raise InvalidTransitionError(
                    "Một hoặc nhiều mã không còn trong SQLite ledger.",
                    details={"missing_codes": missing},
                )

            deleted_order_ids = [int(row["source_order_id"]) for row in ledger_rows]
            affected_runs = {int(row["run_id"]) for row in ledger_rows}
            affected_batches = {int(row["batch_id"]) for row in ledger_rows}
            order_placeholders = ",".join("?" for _ in deleted_order_ids)

            # scan_order_decisions keeps historical prior_order references
            # without ON DELETE SET NULL; clear them before removing source
            # rows so the selected codes can be scanned again safely.
            connection.execute(
                f"""
                UPDATE scan_order_decisions
                SET prior_order_id = NULL, prior_run_id = NULL, prior_batch_id = NULL
                WHERE prior_order_id IN ({order_placeholders})
                """,
                tuple(deleted_order_ids),
            )
            connection.execute(
                f"DELETE FROM confirmed_order_ledger WHERE business_date = ? AND order_code IN ({placeholders})",
                (day.isoformat(), *selected),
            )
            connection.execute(
                f"DELETE FROM orders WHERE id IN ({order_placeholders})",
                tuple(deleted_order_ids),
            )

            for batch_id in sorted(affected_batches):
                remaining = connection.execute(
                    """
                    SELECT bo.order_id, o.sequence,
                           (SELECT COUNT(*) FROM line_items li WHERE li.order_id = o.id) AS line_count
                    FROM batch_orders bo
                    JOIN orders o ON o.id = bo.order_id
                    WHERE bo.batch_id = ? ORDER BY o.sequence
                    """,
                    (batch_id,),
                ).fetchall()
                if not remaining:
                    connection.execute("DELETE FROM batches WHERE id = ?", (batch_id,))
                    continue
                # Reindex the surviving sheets without colliding with the
                # existing PRIMARY KEY while shifting 2..n down to 1..n-1.
                connection.execute(
                    "UPDATE batch_orders SET sheet_index = sheet_index + 1000 WHERE batch_id = ?",
                    (batch_id,),
                )
                for index, row in enumerate(remaining, 1):
                    connection.execute(
                        "UPDATE batch_orders SET sheet_index = ?, updated_at = ? WHERE batch_id = ? AND order_id = ?",
                        (index, now, batch_id, row["order_id"]),
                    )
                connection.execute(
                    """
                    UPDATE batches
                    SET sheet_count = ?, line_count = ?, updated_at = ?,
                        error_message = ?
                    WHERE id = ?
                    """,
                    (
                        len(remaining),
                        sum(int(row["line_count"] or 0) for row in remaining),
                        now,
                        "Một số mã đã được user xóa khỏi ledger để chạy lại.",
                        batch_id,
                    ),
                )

            for run_id in sorted(affected_runs):
                remaining_run = connection.execute(
                    """
                    SELECT COUNT(*) AS order_count,
                           COALESCE(SUM((SELECT COUNT(*) FROM line_items li WHERE li.order_id = o.id)), 0) AS line_count
                    FROM orders o WHERE o.run_id = ?
                    """,
                    (run_id,),
                ).fetchone()
                order_count = int(remaining_run["order_count"] or 0)
                line_count = int(remaining_run["line_count"] or 0)
                connection.execute(
                    """
                    UPDATE runs
                    SET order_count = ?, line_count = ?, updated_at = ?,
                        error_message = CASE WHEN ? = 0
                            THEN 'User đã xóa toàn bộ mã đã chọn khỏi SQLite để chạy lại.'
                            ELSE 'User đã xóa một phần mã đã chọn khỏi SQLite để chạy lại.'
                        END,
                        status = CASE WHEN ? = 0 THEN 'cancelled' ELSE status END
                    WHERE id = ?
                    """,
                    (order_count, line_count, now, order_count, order_count, run_id),
                )
                self._insert_event(
                    connection,
                    run_id=run_id,
                    event_type="orders.ledger_deleted_selected",
                    level="warning",
                    message="User đã xóa các mã đã xác nhận khỏi SQLite; KiotViet không bị thay đổi.",
                    payload={
                        "business_date": day.isoformat(),
                        "order_codes": selected,
                        "acknowledge_external_order_risk": True,
                    },
                    created_at=now,
                )

        return {
            "ok": True,
            "business_date": day.isoformat(),
            "deleted_count": len(selected),
            "deleted_order_codes": selected,
            "message": "Đã xóa mã khỏi SQLite. KiotViet không bị thay đổi; hãy quét report lại để tạo workflow mới.",
            "state": self.state(),
        }

    def scan_state(self, scan_id: int | None = None) -> dict[str, Any] | None:
        """Return the most recent auditable report scan.

        A scan is intentionally separate from a workflow run: downloading a
        report with no new orders is still useful evidence and must not create
        an empty KiotViet workbook.
        """

        if scan_id is None:
            scan = self.db.fetch_one(
                "SELECT * FROM report_scans WHERE business_date = ? ORDER BY id DESC LIMIT 1",
                (self._today().isoformat(),),
            )
        else:
            scan = self.db.fetch_one("SELECT * FROM report_scans WHERE id = ?", (int(scan_id),))
        if scan is None:
            return None
        decisions = self.db.fetch_all(
            """
            SELECT order_code, source_sequence, source_start_row, source_end_row,
                   decision, reason, prior_order_id, prior_run_id, prior_batch_id
            FROM scan_order_decisions WHERE scan_id = ?
            ORDER BY source_sequence
            """,
            (scan["id"],),
        )
        scan["decisions"] = decisions
        scan["blocked"] = scan["status"] in {
            "duplicate_review",
            "unresolved",
            "changed_confirmed",
            "active_workflow",
        }
        return scan

    def scan_and_create_incremental_run(
        self,
        source_path: str | Path,
        *,
        trigger: str = "manual",
        allow_confirmed_duplicates: bool = False,
        business_date: str | date | None = None,
    ) -> dict[str, Any]:
        """Safely scan a Sopoka report and create a run containing *only new* orders.

        The raw Excel report is always read-only.  Orders are skipped automatically
        only when they have a same-day row in ``confirmed_order_ledger``.
        Before that skip happens, the user must explicitly acknowledge the
        duplicate list. Any order found in an unfinished/failed workflow is
        reported as unresolved and prevents a new import run.
        """

        trigger = str(trigger or "manual").strip().casefold()
        if trigger not in {"manual", "scheduled"}:
            raise SourceValidationError("Report scan trigger must be manual or scheduled.")

        requested_date = (
            None
            if business_date in (None, "")
            else _date_value(business_date, field="business_date")
        )
        source_info = self._read_source_for_incremental_scan(
            source_path,
            expected_date=requested_date,
        )
        source = source_info["source"]
        stat = source_info["stat"]
        source_hash = source_info["source_hash"]
        business_date = source_info["business_date"]
        parser_metadata = source_info["parser_metadata"]
        orders = source_info["orders"]

        # A workflow created from an explicitly selected report date owns the
        # single browser queue until it is finished.  Do not let a new date
        # overtake it, even after midnight; the caller can resume the
        # historical run by its persisted run id.
        active_global = self._active_run_row()
        if active_global is not None:
            active_date = _date_value(active_global["business_date"], field="business_date")
            if active_date != business_date:
                raise ActiveRunError(
                    "Another selected-date workflow is still active.",
                    details={
                        "run_id": active_global["id"],
                        "business_date": active_global["business_date"],
                        "status": active_global["status"],
                    },
                )
        now = self._iso_now()
        scan_key = uuid.uuid4().hex
        source_line_count = sum(len(order["products"]) for order in orders)

        # Keep a complete classification in SQLite before any workbooks are
        # built.  This is the audit trail that proves why a code was included,
        # skipped, or deliberately held back.
        with self.db.transaction(immediate=True) as connection:
            active = self._active_run_in_connection(connection, business_date)
            decisions: list[dict[str, Any]] = []
            new_orders: list[dict[str, Any]] = []
            for order in orders:
                ledger = connection.execute(
                    """
                    SELECT business_date, order_code, payload_sha256, source_order_id,
                           run_id, batch_id, confirmed_at
                    FROM confirmed_order_ledger
                    WHERE business_date = ? AND order_code = ?
                    """,
                    (business_date.isoformat(), order["order_code"]),
                ).fetchone()
                if ledger is not None:
                    same_payload = ledger["payload_sha256"] == order["payload_sha256"]
                    decisions.append(
                        {
                            "order": order,
                            "decision": "skipped_confirmed" if same_payload else "changed_confirmed",
                            "reason": (
                                "Mã đã được người dùng xác nhận hoàn tất trong ngày."
                                if same_payload
                                else "Mã đã xác nhận nhưng dữ liệu sản phẩm/khách hàng đã thay đổi."
                            ),
                            "prior_order_id": ledger["source_order_id"],
                            "prior_run_id": ledger["run_id"],
                            "prior_batch_id": ledger["batch_id"],
                        }
                    )
                    continue

                pending = connection.execute(
                    """
                    SELECT o.id AS order_id, o.run_id AS run_id,
                           b.id AS batch_id, b.status AS batch_status,
                           r.status AS run_status
                    FROM orders o
                    JOIN runs r ON r.id = o.run_id
                    LEFT JOIN batch_orders bo ON bo.order_id = o.id
                    LEFT JOIN batches b ON b.id = bo.batch_id
                    WHERE r.business_date = ? AND o.order_code = ?
                      AND r.status NOT IN ('cancelled', 'expired')
                    ORDER BY o.id DESC
                    LIMIT 1
                    """,
                    (business_date.isoformat(), order["order_code"]),
                ).fetchone()
                if pending is not None:
                    decisions.append(
                        {
                            "order": order,
                            "decision": "unresolved",
                            "reason": (
                                "Mã có trong workflow chưa có xác nhận hoàn tất "
                                f"(run={pending['run_status']}, batch={pending['batch_status'] or 'chưa tạo'})."
                            ),
                            "prior_order_id": pending["order_id"],
                            "prior_run_id": pending["run_id"],
                            "prior_batch_id": pending["batch_id"],
                        }
                    )
                    continue

                decisions.append(
                    {
                        "order": order,
                        "decision": "new",
                        "reason": "Mã chưa từng được xác nhận hoặc đưa vào workflow trong ngày.",
                        "prior_order_id": None,
                        "prior_run_id": None,
                        "prior_batch_id": None,
                    }
                )
                new_orders.append(order)

            skipped_count = sum(item["decision"] == "skipped_confirmed" for item in decisions)
            unresolved_count = sum(item["decision"] == "unresolved" for item in decisions)
            changed_count = sum(item["decision"] == "changed_confirmed" for item in decisions)
            new_line_count = sum(len(order["products"]) for order in new_orders)
            if changed_count:
                scan_status = "changed_confirmed"
            elif unresolved_count:
                scan_status = "unresolved"
            elif active is not None:
                # Even a report containing entirely different codes must not
                # overtake a live workflow.  The user finishes or explicitly
                # reconciles that workflow first.
                scan_status = "active_workflow"
            elif skipped_count and not allow_confirmed_duplicates:
                scan_status = "duplicate_review"
            elif not new_orders:
                scan_status = "no_new"
            else:
                scan_status = "run_created"

            cursor = connection.execute(
                """
                INSERT INTO report_scans(
                    scan_key, business_date, trigger, source_path, source_name,
                    source_sha256, source_size, source_mtime_ns, source_order_count,
                    source_line_count, new_order_count, new_line_count,
                    skipped_confirmed_count, unresolved_count,
                    changed_confirmed_count, status, created_at, completed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    scan_key,
                    business_date.isoformat(),
                    trigger,
                    str(source),
                    source.name,
                    source_hash,
                    stat.st_size,
                    stat.st_mtime_ns,
                    len(orders),
                    source_line_count,
                    len(new_orders),
                    new_line_count,
                    skipped_count,
                    unresolved_count,
                    changed_count,
                    scan_status,
                    now,
                    now,
                ),
            )
            scan_id = int(cursor.lastrowid)
            for item in decisions:
                order = item["order"]
                connection.execute(
                    """
                    INSERT INTO scan_order_decisions(
                        scan_id, order_code, source_sequence, source_start_row,
                        source_end_row, payload_sha256, decision, prior_order_id,
                        prior_run_id, prior_batch_id, reason, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        scan_id,
                        order["order_code"],
                        int(order["sequence"]),
                        order.get("source_start_row"),
                        order.get("source_end_row"),
                        order["payload_sha256"],
                        item["decision"],
                        item["prior_order_id"],
                        item["prior_run_id"],
                        item["prior_batch_id"],
                        item["reason"],
                        now,
                    ),
                )

            run_id: int | None = None
            if scan_status == "run_created":
                run_id = self._insert_incremental_run(
                    connection,
                    source=source,
                    stat=stat,
                    source_hash=source_hash,
                    business_date=business_date,
                    parser_metadata=parser_metadata,
                    orders=new_orders,
                    scan_id=scan_id,
                    source_order_count=len(orders),
                    historical_run=requested_date is not None,
                    now=now,
                )

        self._write_scan_audit(scan_id)
        if run_id is not None:
            # The run contains only the newly classified orders.  The legacy
            # duplicate-review flow is intentionally not used here.
            state = self.prepare_batches(run_id)
        else:
            state = self.state()
        state["scan"] = self.scan_state(scan_id)
        state["requires_duplicate_confirmation"] = scan_status == "duplicate_review"
        return state

    def decide_scan_duplicates(
        self,
        scan_id: int,
        *,
        continue_scan: bool,
    ) -> dict[str, Any]:
        """Record the user's decision for same-day confirmed codes.

        Continuing never recreates a confirmed KiotViet order: the report is
        rescanned with the acknowledgement and only its genuinely new codes
        are placed into an import workbook.
        """

        scan = self.db.fetch_one("SELECT * FROM report_scans WHERE id = ?", (int(scan_id),))
        if scan is None:
            raise WorkflowError(f"Scan {scan_id} was not found.")
        if scan["status"] != "duplicate_review":
            raise InvalidTransitionError(
                "This scan is not waiting for a duplicate decision.",
                details={"scan_id": int(scan_id), "status": scan["status"]},
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                "UPDATE report_scans SET status = ?, completed_at = ?, error_message = ? WHERE id = ?",
                (
                    "continued" if continue_scan else "cancelled",
                    now,
                    "Người dùng đồng ý tiếp tục với các mã mới còn lại."
                    if continue_scan
                    else "Người dùng dừng do phát hiện mã đã xử lý trong ngày.",
                    int(scan_id),
                ),
            )

        self._write_scan_audit(int(scan_id))
        if not continue_scan:
            state = self.state()
            state["scan"] = self.scan_state(int(scan_id))
            return state

        return self.scan_and_create_incremental_run(
            scan["source_path"],
            trigger="manual",
            allow_confirmed_duplicates=True,
            business_date=scan["business_date"],
        )

    def create_run_from_source(
        self,
        source_path: str | Path,
        allow_duplicates: bool = False,
        business_date: str | date | None = None,
    ) -> dict[str, Any]:
        """Parse one today's report and persist an auditable run.

        Duplicate detection is scoped to ``(business_date, order_code)``.  A
        blocked run is idempotently reused; calling again with
        ``allow_duplicates=True`` records the user's explicit override.
        """

        source = Path(source_path).expanduser().resolve()
        if not source.is_file():
            raise SourceValidationError(f"Source report was not found: {source}")
        if source.suffix.casefold() not in SUPPORTED_REPORT_EXTENSIONS:
            raise SourceValidationError(
                f"The Sopoka source must be an Excel file ({SUPPORTED_REPORT_FORMATS})."
            )

        stat = source.stat()
        source_hash = sha256_file(source)
        requested_date = (
            None
            if business_date in (None, "")
            else _date_value(business_date, field="business_date")
        )
        parsed = self._parse_source(source, expected_date=requested_date)
        report = _mapping(parsed)
        business_date = _date_value(
            report.get("report_start_date") or getattr(parsed, "report_start_date", None),
            field="report_start_date",
        )
        report_end = _date_value(
            report.get("report_end_date") or getattr(parsed, "report_end_date", None),
            field="report_end_date",
        )
        if business_date != report_end:
            raise SourceValidationError("The source report must cover one day only.")
        if requested_date is None and business_date != self._today():
            raise MidnightGuardError(
                "The report is no longer for today's Thailand business date.",
                details={
                    "report_date": business_date.isoformat(),
                    "today": self._today().isoformat(),
                },
            )

        # A live workflow from a previous business day must not block a new
        # report.  Parse and validate the incoming source first, so we know
        # which Thailand business date is being started, then expire any stale
        # active run before applying the usual same-day guard.
        active = self._active_run_row()
        while active is not None:
            active_date = _date_value(active["business_date"], field="business_date")
            if active_date == business_date:
                break
            if bool(active.get("historical_run")):
                raise ActiveRunError(
                    "Another selected-date workflow is still active.",
                    details={
                        "run_id": active["id"],
                        "business_date": active["business_date"],
                        "status": active["status"],
                    },
                )
            self._expire_stale_run(
                active,
                run_date=active_date,
                today=business_date,
            )
            active = self._active_run_row()

        if active is not None:
            if active["source_sha256"] != source_hash:
                raise ActiveRunError(
                    "Another workflow is still active.",
                    details={"run_id": active["id"], "status": active["status"]},
                )
            if active["status"] == "duplicate_review" and allow_duplicates:
                self._guard_business_date(active)
                self._override_duplicates(int(active["id"]))
            return self.state(int(active["id"]))

        parsed_orders_value = report.get("orders") or getattr(parsed, "orders", None)
        if not isinstance(parsed_orders_value, Sequence) or isinstance(
            parsed_orders_value, (str, bytes)
        ):
            raise SourceValidationError("Parser returned no order list.")
        orders = [self._normalize_order(order, index) for index, order in enumerate(parsed_orders_value, 1)]
        if not orders:
            raise SourceValidationError("Today's report contains no orders.")
        if len({order["order_code"] for order in orders}) != len(orders):
            raise SourceValidationError("The source contains duplicate order codes.")

        now = self._iso_now()
        parser_metadata = {
            key: report.get(key)
            for key in (
                "source_file",
                "sheet_name",
                "header_row",
                "first_data_row",
                "last_data_row",
                "report_start_date",
                "report_end_date",
                "qa",
            )
            if key in report
        }
        run_key = uuid.uuid4().hex
        line_count = sum(len(order["products"]) for order in orders)

        with self.db.transaction(immediate=True) as connection:
            duplicate_file = connection.execute(
                """
                SELECT 1 FROM runs
                WHERE business_date = ? AND source_sha256 = ?
                  AND status NOT IN ('cancelled', 'expired')
                LIMIT 1
                """,
                (business_date.isoformat(), source_hash),
            ).fetchone() is not None

            duplicate_matches: dict[str, sqlite3.Row] = {}
            for order in orders:
                match = connection.execute(
                    """
                    SELECT o.id AS prior_order_id, o.run_id AS prior_run_id,
                           o.payload_sha256, o.status AS prior_order_status,
                           r.status AS prior_run_status, r.created_at
                    FROM orders o
                    JOIN runs r ON r.id = o.run_id
                    WHERE r.business_date = ? AND o.order_code = ?
                      AND r.status NOT IN ('cancelled', 'expired')
                    ORDER BY o.id DESC
                    LIMIT 1
                    """,
                    (business_date.isoformat(), order["order_code"]),
                ).fetchone()
                if match is not None:
                    duplicate_matches[order["order_code"]] = match

            duplicate_count = len(duplicate_matches)
            status = (
                "duplicate_review"
                if duplicate_count and not allow_duplicates
                else "parsed"
            )
            cursor = connection.execute(
                """
                INSERT INTO runs(
                    run_key, business_date, source_path, source_name,
                    source_sha256, source_size, source_mtime_ns,
                    parser_metadata_json, status, order_count, line_count,
                    duplicate_count, duplicate_file, duplicates_allowed,
                    historical_run, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_key,
                    business_date.isoformat(),
                    str(source),
                    source.name,
                    source_hash,
                    stat.st_size,
                    stat.st_mtime_ns,
                    json_dumps(parser_metadata),
                    status,
                    len(orders),
                    line_count,
                    duplicate_count,
                    int(duplicate_file),
                    int(bool(allow_duplicates and duplicate_count)),
                    int(requested_date is not None),
                    now,
                    now,
                ),
            )
            run_id = int(cursor.lastrowid)

            for order in orders:
                match = duplicate_matches.get(order["order_code"])
                order_status = (
                    "duplicate_pending"
                    if match is not None and not allow_duplicates
                    else "duplicate_override"
                    if match is not None
                    else "parsed"
                )
                order_cursor = connection.execute(
                    """
                    INSERT INTO orders(
                        run_id, sequence, order_code, payload_sha256,
                        source_start_row, source_end_row, customer_code,
                        customer_name, customer_address, order_date, status,
                        payload_json, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        order["sequence"],
                        order["order_code"],
                        order["payload_sha256"],
                        order.get("source_start_row"),
                        order.get("source_end_row"),
                        order["customer_code"],
                        order["customer_name"],
                        order["customer_address"],
                        order.get("order_date"),
                        order_status,
                        json_dumps(order["raw"]),
                        now,
                        now,
                    ),
                )
                order_id = int(order_cursor.lastrowid)
                for item_index, item in enumerate(order["products"], 1):
                    connection.execute(
                        """
                        INSERT INTO line_items(
                            order_id, sequence, source_row, product_code,
                            product_name, quantity, unit_price, gross_amount,
                            discount_percent, payload_json
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            order_id,
                            item_index,
                            item.get("source_row"),
                            item["product_code"],
                            item["product_name"],
                            _decimal_text(item["quantity"]),
                            _decimal_text(item["unit_price"]),
                            _decimal_text(item.get("gross_amount", 0)),
                            _decimal_text(item["discount_percent"]),
                            json_dumps(item["raw"]),
                        ),
                    )
                if match is not None:
                    kind = (
                        "same_payload"
                        if match["payload_sha256"] == order["payload_sha256"]
                        else "changed_payload"
                    )
                    connection.execute(
                        """
                        INSERT INTO duplicate_decisions(
                            run_id, order_id, prior_run_id, prior_order_id,
                            duplicate_kind, decision, decided_at, created_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            run_id,
                            order_id,
                            match["prior_run_id"],
                            match["prior_order_id"],
                            kind,
                            "continue" if allow_duplicates else "pending",
                            now if allow_duplicates else None,
                            now,
                        ),
                    )

            self._insert_event(
                connection,
                run_id=run_id,
                event_type="run.created",
                message="Source report parsed and stored.",
                payload={
                    "source_sha256": source_hash,
                    "order_count": len(orders),
                    "line_count": line_count,
                    "duplicate_count": duplicate_count,
                    "duplicate_file": duplicate_file,
                },
                created_at=now,
            )
            if duplicate_count:
                self._insert_event(
                    connection,
                    run_id=run_id,
                    event_type=(
                        "duplicates.overridden"
                        if allow_duplicates
                        else "duplicates.detected"
                    ),
                    level="warning",
                    message=(
                        "User explicitly continued same-day duplicate orders."
                        if allow_duplicates
                        else "Workflow paused for same-day duplicate review."
                    ),
                    payload={"count": duplicate_count},
                    created_at=now,
                )

        return self.state(run_id)

    def prepare_batches(self, run_id: int | None = None) -> dict[str, Any]:
        """Build deterministic workbooks of at most 19 orders each."""

        run = self._resolve_run(run_id)
        self._guard_business_date(run)
        if run["status"] == "duplicate_review":
            raise DuplicateReviewRequired(
                "Same-day duplicate orders require an explicit user decision.",
                details={"run_id": run["id"], "duplicate_count": run["duplicate_count"]},
            )
        if run["status"] in {"batch_in_progress", "waiting_user_confirmation", "completed"}:
            return self.state(int(run["id"]))
        if run["status"] not in {"parsed", "preparing_batches", "batches_ready", "failed"}:
            raise InvalidTransitionError(
                f"Cannot prepare batches while run status is {run['status']}."
            )

        run_id = int(run["id"])
        orders = self.db.fetch_all(
            "SELECT * FROM orders WHERE run_id = ? ORDER BY sequence", (run_id,)
        )
        if not orders:
            raise SourceValidationError("The run has no persisted orders.")

        existing = self.db.fetch_all(
            "SELECT * FROM batches WHERE run_id = ? ORDER BY batch_number", (run_id,)
        )
        if not existing:
            now = self._iso_now()
            with self.db.transaction(immediate=True) as connection:
                connection.execute(
                    "UPDATE runs SET status = 'preparing_batches', updated_at = ?, error_message = NULL WHERE id = ?",
                    (now, run_id),
                )
                for start in range(0, len(orders), self.config.max_orders_per_batch):
                    group = orders[start : start + self.config.max_orders_per_batch]
                    batch_number = (start // self.config.max_orders_per_batch) + 1
                    line_count = sum(
                        int(
                            connection.execute(
                                "SELECT COUNT(*) FROM line_items WHERE order_id = ?",
                                (order["id"],),
                            ).fetchone()[0]
                        )
                        for order in group
                    )
                    batch_cursor = connection.execute(
                        """
                        INSERT INTO batches(
                            run_id, batch_number, sheet_count, line_count,
                            status, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, 'pending', ?, ?)
                        """,
                        (run_id, batch_number, len(group), line_count, now, now),
                    )
                    batch_id = int(batch_cursor.lastrowid)
                    for sheet_index, order in enumerate(group, 1):
                        connection.execute(
                            """
                            INSERT INTO batch_orders(
                                batch_id, order_id, sheet_index,
                                expected_tab_label, status, updated_at
                            ) VALUES (?, ?, ?, ?, 'pending', ?)
                            """,
                            (
                                batch_id,
                                order["id"],
                                sheet_index,
                                # KiotViet keeps one blank base tab and appends
                                # one tab per imported worksheet.
                                f"Đặt hàng {sheet_index + 1}",
                                now,
                            ),
                        )
                self._insert_event(
                    connection,
                    run_id=run_id,
                    event_type="batches.planned",
                    message="Order batches were planned.",
                    payload={
                        "batch_count": (len(orders) + self.config.max_orders_per_batch - 1)
                        // self.config.max_orders_per_batch
                    },
                    created_at=now,
                )
            existing = self.db.fetch_all(
                "SELECT * FROM batches WHERE run_id = ? ORDER BY batch_number",
                (run_id,),
            )

        run_output_dir = (
            self.config.outputs_dir
            / str(run["business_date"])
            / str(run["run_key"])
        )
        for batch in existing:
            stored_path = Path(batch["workbook_path"]) if batch["workbook_path"] else None
            if batch["status"] in {
                "ready",
                "importing",
                "imported",
                "notes_in_progress",
                "waiting_user_confirmation",
                "confirmed",
            } and stored_path and stored_path.is_file():
                continue

            self._mark_batch_building(int(batch["id"]), run_id)
            build_orders = self._builder_orders_for_batch(int(batch["id"]))
            output = run_output_dir / (
                f"KiotViet_{run['business_date']}_batch_{int(batch['batch_number']):02d}.xlsx"
            )
            try:
                result = self.workbook_builder.build(build_orders, output)
                self._record_build_result(run_id, int(batch["id"]), result)
            except Exception as exc:
                self._record_build_failure(run_id, int(batch["id"]), exc)
                if isinstance(exc, WorkflowError):
                    raise
                if isinstance(exc, WorkbookBuildError):
                    raise WorkflowError(str(exc), details={"batch_id": batch["id"]}) from exc
                raise

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            not_ready = int(
                connection.execute(
                    "SELECT COUNT(*) FROM batches WHERE run_id = ? AND status != 'ready'",
                    (run_id,),
                ).fetchone()[0]
            )
            if not_ready == 0:
                connection.execute(
                    "UPDATE runs SET status = 'batches_ready', updated_at = ?, error_message = NULL WHERE id = ?",
                    (now, run_id),
                )
                self._insert_event(
                    connection,
                    run_id=run_id,
                    event_type="batches.ready",
                    message="All import workbooks are ready.",
                    payload={"batch_count": len(existing)},
                    created_at=now,
                )
        return self.state(run_id)

    def batch_execution_payload(self, batch_id: int) -> dict[str, Any]:
        """Return the immutable DB-backed input for one unlocked Edge batch.

        The browser worker must never reconstruct product/order mappings from
        the UI.  It asks this method for the workbook path and the exact
        KiotViet tab expectations that were persisted when the workbook was
        built.  A later batch cannot be read until all earlier batches have
        a durable confirmation.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "ready":
            raise InvalidTransitionError(
                "The batch is not ready for automated import.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )

        workbook_path = Path(str(batch.get("workbook_path") or "")).expanduser().resolve()
        if not workbook_path.is_file():
            raise SourceValidationError(
                "The generated KiotViet workbook is missing.",
                details={"batch_id": batch_id, "workbook_path": str(workbook_path)},
            )
        try:
            workbook_path.relative_to(self.config.outputs_dir.resolve())
        except ValueError as exc:
            raise SourceValidationError(
                "The generated workbook is outside the configured outputs directory.",
                details={"batch_id": batch_id, "workbook_path": str(workbook_path)},
            ) from exc

        rows = self.db.fetch_all(
            """
            SELECT bo.sheet_index, bo.expected_tab_label,
                   o.id AS order_id, o.order_code, o.customer_code, o.customer_name,
                   o.payload_json,
                   (SELECT COUNT(*) FROM line_items li WHERE li.order_id = o.id) AS line_count,
                   (SELECT li.product_code FROM line_items li
                    WHERE li.order_id = o.id ORDER BY li.sequence LIMIT 1) AS first_sku
            FROM batch_orders bo
            JOIN orders o ON o.id = bo.order_id
            WHERE bo.batch_id = ?
            ORDER BY bo.sheet_index
            """,
            (batch_id,),
        )
        if len(rows) != int(batch["sheet_count"]):
            raise WorkflowError(
                "The persisted batch mappings are incomplete.",
                details={
                    "batch_id": batch_id,
                    "expected": int(batch["sheet_count"]),
                    "actual": len(rows),
                },
            )
        order_ids = [int(row["order_id"]) for row in rows]
        product_codes_by_order: dict[int, list[str]] = {order_id: [] for order_id in order_ids}
        product_names_by_order: dict[int, list[str]] = {order_id: [] for order_id in order_ids}
        product_lines_by_order: dict[int, list[dict[str, Any]]] = {
            order_id: [] for order_id in order_ids
        }
        if order_ids:
            placeholders = ",".join("?" for _ in order_ids)
            item_rows = self.db.fetch_all(
                f"""
                SELECT order_id, product_code, product_name, quantity, gross_amount, discount_percent
                FROM line_items WHERE order_id IN ({placeholders})
                ORDER BY order_id, sequence
                """,
                tuple(order_ids),
            )
            for item in item_rows:
                code = str(item["product_code"] or "").strip()
                if code:
                    product_codes_by_order.setdefault(int(item["order_id"]), []).append(code)
                name = str(item["product_name"] or "").strip()
                if name:
                    product_names_by_order.setdefault(int(item["order_id"]), []).append(name)
                product_lines_by_order.setdefault(int(item["order_id"]), []).append(
                    {
                        "productCode": code,
                        "productName": name,
                        "quantity": _json_number(item["quantity"]),
                        "netAmount": _line_net_amount(
                            item["gross_amount"], item["discount_percent"]
                        ),
                    }
                )
        orders = []
        for row in rows:
            raw_payload = _readable_json(row.get("payload_json"), {})
            orders.append(
                {
                    "sheetIndex": int(row["sheet_index"]),
                    "orderCode": str(row["order_code"]),
                    "customerCode": str(row["customer_code"] or ""),
                    "customerName": str(row["customer_name"] or ""),
                    "firstSku": str(row["first_sku"] or ""),
                    "productCodes": product_codes_by_order.get(int(row["order_id"]), []),
                    "productNames": product_names_by_order.get(int(row["order_id"]), []),
                    "productLines": product_lines_by_order.get(int(row["order_id"]), []),
                    "lineCount": int(row["line_count"] or 0),
                    "paymentAmount": _payment_amount_from_payload(raw_payload),
                    "expectedTabLabel": str(row["expected_tab_label"] or ""),
                }
            )
        return {
            "batchId": int(batch["id"]),
            "runId": int(run["id"]),
            "batchNumber": int(batch["batch_number"]),
            "businessDate": str(run["business_date"]),
            "workbookPath": str(workbook_path),
            "orders": orders,
        }

    def batch_submission_payload(self, batch_id: int) -> dict[str, Any]:
        """Return trusted expectations to resume final submission only.

        This is intentionally separate from :meth:`batch_execution_payload`:
        it can be used only after a workbook was already imported and every
        note has been durably verified.  It lets an interrupted server restart
        finish the current KiotViet tabs without uploading that workbook a
        second time.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "waiting_user_confirmation":
            raise InvalidTransitionError(
                "The batch is not ready to resume final-order submission.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )
        with self.db.transaction() as connection:
            rows = self._require_verified_batch_mappings(
                connection,
                batch_id=batch_id,
                expected_count=int(batch["sheet_count"]),
            )
        return {
            "batchId": int(batch["id"]),
            "runId": int(run["id"]),
            "batchNumber": int(batch["batch_number"]),
            "businessDate": str(run["business_date"]),
            "orders": [
                {
                    "sheetIndex": int(row["sheet_index"]),
                    "orderCode": str(row["order_code"]),
                    "expectedTabLabel": str(row["expected_tab_label"] or ""),
                }
                for row in rows
            ],
        }

    def confirm_batch(
        self,
        batch_id: int | None = None,
        *,
        run_id: int | None = None,
        batch_number: int | None = None,
    ) -> dict[str, Any]:
        """Record the user's manual confirmation and unlock the next batch.

        This method only changes SQLite state.  It never invokes the browser or
        the KiotViet final-order action.
        """

        batch = self._resolve_batch(
            batch_id=batch_id, run_id=run_id, batch_number=batch_number
        )
        run = self._resolve_run(int(batch["run_id"]))
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] == "confirmed":
            return self.state(int(run["id"]))
        if batch["status"] != "waiting_user_confirmation":
            raise InvalidTransitionError(
                "The batch is not ready for manual confirmation.",
                details={"batch_id": batch["id"], "status": batch["status"]},
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            confirmed_orders = connection.execute(
                """
                SELECT o.id, o.order_code, o.payload_sha256
                FROM batch_orders bo
                JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ? ORDER BY bo.sheet_index
                """,
                (batch["id"],),
            ).fetchall()
            for order in confirmed_orders:
                existing = connection.execute(
                    """
                    SELECT source_order_id, run_id, batch_id
                    FROM confirmed_order_ledger
                    WHERE business_date = ? AND order_code = ?
                    """,
                    (run["business_date"], order["order_code"]),
                ).fetchone()
                if existing is not None and int(existing["source_order_id"]) != int(order["id"]):
                    raise InvalidTransitionError(
                        "A same-day order code is already confirmed by a different batch.",
                        details={
                            "order_code": order["order_code"],
                            "confirmed_run_id": existing["run_id"],
                            "confirmed_batch_id": existing["batch_id"],
                        },
                    )
                if existing is None:
                    connection.execute(
                        """
                        INSERT INTO confirmed_order_ledger(
                            business_date, order_code, payload_sha256,
                            source_order_id, run_id, batch_id,
                            confirmation_method, confirmed_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            run["business_date"],
                            order["order_code"],
                            order["payload_sha256"],
                            order["id"],
                            run["id"],
                            batch["id"],
                            MANUAL_CONFIRMATION_METHOD,
                            now,
                        ),
                    )
            connection.execute(
                """
                UPDATE batches
                SET status = 'confirmed', user_confirmed_at = ?, updated_at = ?
                WHERE id = ?
                """,
                (now, now, batch["id"]),
            )
            connection.execute(
                """
                UPDATE orders SET status = 'user_confirmed', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch["id"]),
            )
            next_batch = connection.execute(
                """
                SELECT * FROM batches
                WHERE run_id = ? AND batch_number > ?
                ORDER BY batch_number LIMIT 1
                """,
                (run["id"], batch["batch_number"]),
            ).fetchone()
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=int(batch["id"]),
                event_type="batch.user_confirmed",
                message="User confirmed that final orders were handled manually.",
                payload={"batch_number": batch["batch_number"]},
                created_at=now,
            )
            if next_batch is None:
                connection.execute(
                    """
                    UPDATE runs
                    SET status = 'completed', completed_at = ?, error_message = NULL, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, now, run["id"]),
                )
                self._insert_event(
                    connection,
                    run_id=int(run["id"]),
                    event_type="run.completed",
                    message="All batches were manually confirmed.",
                    created_at=now,
                )
            else:
                connection.execute(
                    "UPDATE runs SET status = 'batches_ready', error_message = NULL, updated_at = ? WHERE id = ?",
                    (now, run["id"]),
                )
                self._insert_event(
                    connection,
                    run_id=int(run["id"]),
                    batch_id=int(next_batch["id"]),
                    event_type="batch.next_ready",
                    message="The next batch is unlocked after manual confirmation.",
                    payload={"batch_number": next_batch["batch_number"]},
                    created_at=now,
                )
        return self.state(int(run["id"]))

    def reconcile_manual_batch_after_source_cancellation(
        self,
        batch_id: int,
        source_path: str | Path,
        *,
        cancelled_order_code: str,
        evidence: Sequence[Mapping[str, Any]],
    ) -> dict[str, Any]:
        """Close one stopped batch after an audited source cancellation.

        This narrow recovery path is for a batch whose remaining orders were
        submitted manually in KiotViet and whose one duplicate source order
        was subsequently removed in Sopoka.  It never clicks KiotViet or
        recreates a workbook.  Instead, it requires a newly validated Sopoka
        report whose complete code/payload set equals the existing ledger plus
        the manually verified orders, with exactly the supplied cancelled code
        absent.  The operator's KiotViet order references are retained in the
        immutable event log alongside the source fingerprint.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "automation_attention":
            raise InvalidTransitionError(
                "Source-cancellation reconciliation is allowed only for a stopped automated batch.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )
        cancelled_code = str(cancelled_order_code or "").strip()
        if not cancelled_code:
            raise WorkflowError("A cancelled Sopoka order code is required.")
        if not isinstance(evidence, Sequence) or isinstance(evidence, (str, bytes)):
            raise WorkflowError("KiotViet evidence must be an array.")

        business_date = _date_value(run["business_date"], field="business_date")
        source_info = self._read_source_for_incremental_scan(
            source_path,
            expected_date=business_date,
        )
        fresh_orders = source_info["orders"]
        fresh_by_code = {str(order["order_code"]): order for order in fresh_orders}
        if len(fresh_by_code) != len(fresh_orders):  # parser should reject this; keep the recovery strict.
            raise SourceValidationError("The replacement Sopoka report contains duplicate order codes.")

        evidence_by_code: dict[str, dict[str, Any]] = {}
        kiot_codes: set[str] = set()
        for item in evidence:
            if not isinstance(item, Mapping):
                raise WorkflowError("Every KiotViet evidence entry must be an object.")
            code = str(item.get("order_code") or item.get("orderCode") or "").strip()
            kiot_code = str(
                item.get("kiot_order_code") or item.get("kiotOrderCode") or ""
            ).strip()
            if not code or not kiot_code or code in evidence_by_code or kiot_code in kiot_codes:
                raise WorkflowError("Each manual order needs one unique Sopoka code and one unique KiotViet order code.")
            evidence_by_code[code] = dict(item)
            kiot_codes.add(kiot_code)

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            current_batch = connection.execute("SELECT * FROM batches WHERE id = ?", (batch_id,)).fetchone()
            current_run = connection.execute("SELECT * FROM runs WHERE id = ?", (run["id"],)).fetchone()
            if current_batch is None or current_run is None:  # pragma: no cover - FK invariant
                raise WorkflowError("The stopped batch no longer exists.")
            current_batch_dict = dict(current_batch)
            current_run_dict = dict(current_run)
            self._assert_run_not_cancelled(current_run_dict)
            self._assert_batch_unlocked_in_connection(connection, current_batch_dict)
            if current_batch_dict["status"] != "automation_attention":
                raise InvalidTransitionError(
                    "The batch changed state before its source-cancellation reconciliation.",
                    details={"batch_id": batch_id, "status": current_batch_dict["status"]},
                )

            rows = connection.execute(
                """
                SELECT bo.order_id, bo.sheet_index, o.order_code, o.payload_sha256,
                       o.customer_code, o.customer_name
                FROM batch_orders bo
                JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ?
                ORDER BY bo.sheet_index
                """,
                (batch_id,),
            ).fetchall()
            expected_by_code = {str(row["order_code"]): dict(row) for row in rows}
            if len(expected_by_code) != int(current_batch_dict["sheet_count"]):
                raise WorkflowError("The stopped batch does not contain its full persisted order set.")
            cancelled = expected_by_code.get(cancelled_code)
            if cancelled is None:
                raise WorkflowError("The cancelled Sopoka code is not part of this stopped batch.")
            if cancelled_code in fresh_by_code:
                raise WorkflowError("The cancelled Sopoka code is still present in the replacement report.")

            manual_codes = set(expected_by_code) - {cancelled_code}
            if set(evidence_by_code) != manual_codes:
                raise WorkflowError(
                    "KiotViet evidence must cover every remaining order in the stopped batch exactly once.",
                    details={
                        "missing_evidence": sorted(manual_codes - set(evidence_by_code)),
                        "unexpected_evidence": sorted(set(evidence_by_code) - manual_codes),
                    },
                )
            for code in manual_codes:
                persisted = expected_by_code[code]
                replacement = fresh_by_code.get(code)
                if replacement is None or replacement["payload_sha256"] != persisted["payload_sha256"]:
                    raise WorkflowError(
                        "A manually completed order is missing or changed in the replacement Sopoka report.",
                        details={"order_code": code},
                    )
                proof = evidence_by_code[code]
                if str(proof.get("customer_code") or proof.get("customerCode") or "").strip() != str(persisted["customer_code"] or "").strip():
                    raise WorkflowError("KiotViet evidence has a different customer code.", details={"order_code": code})
                if str(proof.get("customer_name") or proof.get("customerName") or "").strip() != str(persisted["customer_name"] or "").strip():
                    raise WorkflowError("KiotViet evidence has a different customer name.", details={"order_code": code})

            ledger_rows = connection.execute(
                """
                SELECT order_code, payload_sha256
                FROM confirmed_order_ledger
                WHERE business_date = ?
                """,
                (current_run_dict["business_date"],),
            ).fetchall()
            ledger_by_code = {str(row["order_code"]): str(row["payload_sha256"]) for row in ledger_rows}
            for code, payload_sha256 in ledger_by_code.items():
                replacement = fresh_by_code.get(code)
                if replacement is None or replacement["payload_sha256"] != payload_sha256:
                    raise WorkflowError(
                        "A previously confirmed order is missing or changed in the replacement Sopoka report.",
                        details={"order_code": code},
                    )
            expected_replacement_codes = set(ledger_by_code) | manual_codes
            if set(fresh_by_code) != expected_replacement_codes:
                raise WorkflowError(
                    "The replacement Sopoka report has unexpected or missing codes, so this batch cannot be reconciled.",
                    details={
                        "unexpected_codes": sorted(set(fresh_by_code) - expected_replacement_codes),
                        "missing_codes": sorted(expected_replacement_codes - set(fresh_by_code)),
                    },
                )

            for code in sorted(manual_codes):
                order = expected_by_code[code]
                existing = connection.execute(
                    """
                    SELECT source_order_id FROM confirmed_order_ledger
                    WHERE business_date = ? AND order_code = ?
                    """,
                    (current_run_dict["business_date"], code),
                ).fetchone()
                if existing is not None and int(existing["source_order_id"]) != int(order["order_id"]):
                    raise InvalidTransitionError(
                        "A manual Sopoka code is already confirmed by a different workflow.",
                        details={"order_code": code},
                    )
                if existing is None:
                    connection.execute(
                        """
                        INSERT INTO confirmed_order_ledger(
                            business_date, order_code, payload_sha256,
                            source_order_id, run_id, batch_id,
                            confirmation_method, confirmed_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            current_run_dict["business_date"],
                            code,
                            order["payload_sha256"],
                            order["order_id"],
                            current_run_dict["id"],
                            batch_id,
                            MANUAL_CONFIRMATION_METHOD,
                            now,
                        ),
                    )

            placeholders = ",".join("?" for _ in manual_codes)
            connection.execute(
                f"UPDATE batch_orders SET status='manual_reconciled', updated_at=? WHERE batch_id=? AND order_id IN ({placeholders})",
                (now, batch_id, *(expected_by_code[code]["order_id"] for code in sorted(manual_codes))),
            )
            connection.execute(
                f"UPDATE orders SET status='manual_reconciled', updated_at=? WHERE id IN ({placeholders})",
                (now, *(expected_by_code[code]["order_id"] for code in sorted(manual_codes))),
            )
            connection.execute(
                "UPDATE batch_orders SET status='source_cancelled', updated_at=? WHERE batch_id=? AND order_id=?",
                (now, batch_id, cancelled["order_id"]),
            )
            connection.execute(
                "UPDATE orders SET status='source_cancelled', updated_at=? WHERE id=?",
                (now, cancelled["order_id"]),
            )
            connection.execute(
                """
                UPDATE batches
                SET status='confirmed', user_confirmed_at=?, error_message=NULL, updated_at=?
                WHERE id=?
                """,
                (now, now, batch_id),
            )
            connection.execute(
                """
                UPDATE runs
                SET status='completed', completed_at=?, error_message=NULL, updated_at=?
                WHERE id=?
                """,
                (now, now, current_run_dict["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(current_run_dict["id"]),
                batch_id=batch_id,
                event_type="batch.manual_source_cancellation_reconciled",
                level="warning",
                message="Manual KiotViet orders were ledgered after one duplicate Sopoka source order was cancelled.",
                payload={
                    "cancelled_order_code": cancelled_code,
                    "manual_order_codes": sorted(manual_codes),
                    "evidence": [evidence_by_code[code] for code in sorted(manual_codes)],
                    "replacement_source": {
                        "path": str(source_info["source"]),
                        "name": source_info["source"].name,
                        "sha256": source_info["source_hash"],
                        "size": source_info["stat"].st_size,
                        "order_count": len(fresh_orders),
                    },
                },
                created_at=now,
            )
            self._insert_event(
                connection,
                run_id=int(current_run_dict["id"]),
                event_type="run.completed",
                message="Stopped run was completed through audited manual reconciliation and source cancellation.",
                payload={
                    "confirmation_method": MANUAL_CONFIRMATION_METHOD,
                    "cancelled_order_code": cancelled_code,
                },
                created_at=now,
            )
        return self.state(int(run["id"]))

    def complete_automated_batch(
        self,
        batch_id: int,
        submission_result: Mapping[str, Any],
    ) -> dict[str, Any]:
        """Persist a fully verified final-order result from direct Edge control.

        Edge has already performed the external action when this method is
        called.  SQLite is updated only when all of the following agree:

        * every persisted sheet has a verified note on its expected KiotViet
          tab;
        * the Edge result explicitly reports a final-order click;
        * the result contains each expected order code exactly once, paired
          with its expected tab label.

        The transaction inserts the same-day de-duplication ledger and either
        unlocks the next batch or closes the run.  This makes retries of the
        API idempotent without allowing an unverified browser response to
        advance the queue.
        """

        if not isinstance(submission_result, Mapping):
            raise WorkflowError("Automated batch completion requires an Edge result object.")

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)

        if batch["status"] == "confirmed":
            with self.db.transaction() as connection:
                self._assert_prior_automated_completion(batch, run, connection=connection)
                self._require_verified_batch_mappings(
                    connection,
                    batch_id=batch_id,
                    expected_count=int(batch["sheet_count"]),
                )
                self._require_automated_submission_result(
                    connection,
                    batch_id=batch_id,
                    submission_result=submission_result,
                )
            return self.state(int(run["id"]))

        allowed_statuses = {
            "importing",
            "imported",
            "notes_in_progress",
            "waiting_user_confirmation",
            # A browser import may have completed and all notes may have been
            # verified before a transient Edge callback/timeout quarantines
            # the batch.  Allow an explicit, fully verified submission proof
            # to reconcile that state without re-importing or duplicating it.
            "automation_attention",
        }
        if batch["status"] not in allowed_statuses:
            raise InvalidTransitionError(
                "The batch is not ready to record an automated final-order result.",
                details={
                    "batch_id": batch_id,
                    "status": batch["status"],
                    "allowed": sorted(allowed_statuses),
                },
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            # Refresh under the write lock so an old browser callback cannot
            # accidentally overwrite a completion that just won the race.
            current_batch = connection.execute(
                "SELECT * FROM batches WHERE id = ?", (batch_id,)
            ).fetchone()
            current_run = connection.execute(
                "SELECT * FROM runs WHERE id = ?", (run["id"],)
            ).fetchone()
            if current_batch is None or current_run is None:  # pragma: no cover - FK invariant
                raise WorkflowError("The batch or run disappeared during automated completion.")
            current_batch_dict = dict(current_batch)
            current_run_dict = dict(current_run)
            self._assert_run_not_cancelled(current_run_dict)
            self._assert_batch_unlocked_in_connection(connection, current_batch_dict)

            if current_batch_dict["status"] == "confirmed":
                self._assert_prior_automated_completion(
                    current_batch_dict, current_run_dict, connection=connection
                )
                raise InvalidTransitionError(
                    "The batch was confirmed concurrently; reload its current workflow state.",
                    details={"batch_id": batch_id},
                )
            if current_batch_dict["status"] not in allowed_statuses:
                raise InvalidTransitionError(
                    "The batch changed state before the automated result could be recorded.",
                    details={"batch_id": batch_id, "status": current_batch_dict["status"]},
                )

            # ``IMPORT_BATCH`` with auto-submit can return its verified notes
            # in the same response.  When imports and submission are invoked
            # separately, those notes are already persisted by /imported;
            # both paths end in the same strict DB validation below.
            if current_batch_dict["status"] == "importing":
                imported_payload = submission_result.get("imported")
                if not isinstance(imported_payload, Mapping) or imported_payload.get("ok") is not True:
                    raise InvalidTransitionError(
                        "An importing batch needs an explicit successful KiotViet import result.",
                        details={"batch_id": batch_id},
                    )
            notes_payload = submission_result.get("notes")
            if isinstance(notes_payload, Mapping):
                self._apply_verified_note_payload(connection, batch_id, notes_payload, now)
            elif current_batch_dict["status"] == "importing":
                raise InvalidTransitionError(
                    "An importing batch needs the verified note result before it can be completed.",
                    details={"batch_id": batch_id},
                )

            self._require_verified_batch_mappings(
                connection,
                batch_id=batch_id,
                expected_count=int(current_batch_dict["sheet_count"]),
            )
            self._require_automated_submission_result(
                connection,
                batch_id=batch_id,
                submission_result=submission_result,
            )

            confirmed_orders = connection.execute(
                """
                SELECT o.id, o.order_code, o.payload_sha256
                FROM batch_orders bo
                JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ?
                ORDER BY bo.sheet_index
                """,
                (batch_id,),
            ).fetchall()
            if len(confirmed_orders) != int(current_batch_dict["sheet_count"]):
                raise WorkflowError(
                    "The batch does not contain the expected number of orders.",
                    details={
                        "batch_id": batch_id,
                        "expected": int(current_batch_dict["sheet_count"]),
                        "actual": len(confirmed_orders),
                    },
                )
            for order in confirmed_orders:
                existing = connection.execute(
                    """
                    SELECT source_order_id, run_id, batch_id, confirmation_method
                    FROM confirmed_order_ledger
                    WHERE business_date = ? AND order_code = ?
                    """,
                    (current_run_dict["business_date"], order["order_code"]),
                ).fetchone()
                if existing is not None:
                    raise InvalidTransitionError(
                        "A same-day order code is already confirmed by a different workflow result.",
                        details={
                            "order_code": order["order_code"],
                            "confirmed_run_id": existing["run_id"],
                            "confirmed_batch_id": existing["batch_id"],
                            "confirmation_method": existing["confirmation_method"],
                        },
                    )
                connection.execute(
                    """
                    INSERT INTO confirmed_order_ledger(
                        business_date, order_code, payload_sha256,
                        source_order_id, run_id, batch_id,
                        confirmation_method, confirmed_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        current_run_dict["business_date"],
                        order["order_code"],
                        order["payload_sha256"],
                        order["id"],
                        current_run_dict["id"],
                        batch_id,
                        AUTOMATION_CONFIRMATION_METHOD,
                        now,
                    ),
                )

            connection.execute(
                """
                UPDATE batches
                SET status = 'confirmed', imported_at = COALESCE(imported_at, ?),
                    review_ready_at = COALESCE(review_ready_at, ?),
                    automation_confirmed_at = ?, error_message = NULL, updated_at = ?
                WHERE id = ?
                """,
                (now, now, now, now, batch_id),
            )
            connection.execute(
                """
                UPDATE batch_orders
                SET status = 'automation_confirmed', updated_at = ?
                WHERE batch_id = ?
                """,
                (now, batch_id),
            )
            connection.execute(
                """
                UPDATE orders SET status = 'automation_confirmed', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch_id),
            )

            next_batch = connection.execute(
                """
                SELECT * FROM batches
                WHERE run_id = ? AND batch_number > ?
                ORDER BY batch_number LIMIT 1
                """,
                (current_run_dict["id"], current_batch_dict["batch_number"]),
            ).fetchone()
            if next_batch is not None and next_batch["status"] != "ready":
                raise InvalidTransitionError(
                    "The next batch is not ready, so the automated queue cannot advance.",
                    details={
                        "batch_id": batch_id,
                        "next_batch_id": next_batch["id"],
                        "next_batch_status": next_batch["status"],
                    },
                )

            submitted_entries = self._submission_entries(submission_result)
            self._insert_event(
                connection,
                run_id=int(current_run_dict["id"]),
                batch_id=batch_id,
                event_type="batch.automation_verified_submission",
                message="Edge verified every final KiotViet order in this batch.",
                payload={
                    "batch_number": int(current_batch_dict["batch_number"]),
                    "order_count": len(confirmed_orders),
                    "confirmation_method": AUTOMATION_CONFIRMATION_METHOD,
                    "submitted": submitted_entries,
                },
                created_at=now,
            )
            if next_batch is None:
                connection.execute(
                    """
                    UPDATE runs
                    SET status = 'completed', completed_at = ?, error_message = NULL, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, now, current_run_dict["id"]),
                )
                self._insert_event(
                    connection,
                    run_id=int(current_run_dict["id"]),
                    event_type="run.completed",
                    message="All batches were completed by verified Edge automation.",
                    payload={"confirmation_method": AUTOMATION_CONFIRMATION_METHOD},
                    created_at=now,
                )
            else:
                connection.execute(
                    """
                    UPDATE runs
                    SET status = 'batches_ready', error_message = NULL, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, current_run_dict["id"]),
                )
                self._insert_event(
                    connection,
                    run_id=int(current_run_dict["id"]),
                    batch_id=int(next_batch["id"]),
                    event_type="batch.next_ready",
                    message="The next batch is unlocked after verified Edge automation.",
                    payload={"batch_number": int(next_batch["batch_number"])},
                    created_at=now,
                )
        return self.state(int(run["id"]))

    def record_automated_submission_failure(
        self,
        batch_id: int,
        failure: Mapping[str, Any] | str | None = None,
    ) -> dict[str, Any]:
        """Quarantine a partial/failed final-order attempt for manual review.

        A final click can be partially successful outside SQLite before Edge
        reports an error.  Such a batch must *not* be blindly retried or
        treated as deduplicated.  This state deliberately leaves the ledger
        empty, keeps all later batches locked, and requires an explicit
        operator recovery path rather than a retry button.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] == "confirmed":
            raise InvalidTransitionError(
                "A confirmed batch cannot be changed to automation attention.",
                details={"batch_id": batch_id},
            )
        allowed_statuses = {
            "importing",
            "imported",
            "notes_in_progress",
            "waiting_user_confirmation",
        }
        if batch["status"] not in allowed_statuses:
            raise InvalidTransitionError(
                "This batch cannot be quarantined from its current state.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )

        payload = dict(failure) if isinstance(failure, Mapping) else {"error": str(failure or "")}
        message = str(
            payload.get("error") or payload.get("message") or "Automated final-order submission was incomplete."
        ).strip()
        if len(message) > 2_000:
            message = message[:2_000]
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            current = connection.execute("SELECT * FROM batches WHERE id = ?", (batch_id,)).fetchone()
            if current is None:  # pragma: no cover - FK invariant
                raise WorkflowError("The batch disappeared while recording its automation failure.")
            current_dict = dict(current)
            self._assert_batch_unlocked_in_connection(connection, current_dict)
            if current_dict["status"] not in allowed_statuses:
                raise InvalidTransitionError(
                    "The batch changed state before the automation failure could be recorded.",
                    details={"batch_id": batch_id, "status": current_dict["status"]},
                )
            connection.execute(
                """
                UPDATE batches
                SET status = 'automation_attention', error_message = ?, updated_at = ?
                WHERE id = ?
                """,
                (message, now, batch_id),
            )
            connection.execute(
                """
                UPDATE batch_orders
                SET status = 'automation_attention', updated_at = ?
                WHERE batch_id = ?
                """,
                (now, batch_id),
            )
            connection.execute(
                """
                UPDATE orders SET status = 'automation_attention', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch_id),
            )
            connection.execute(
                """
                UPDATE runs SET status = 'automation_attention', error_message = ?, updated_at = ?
                WHERE id = ?
                """,
                (message, now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.automation_attention",
                level="error",
                message="Automated final-order result was partial or failed; no ledger rows were written.",
                payload={"error": message, "result": payload},
                created_at=now,
            )
        return self.state(int(run["id"]))

    def reconcile_external_partial_submission(
        self,
        batch_id: int,
        submission_result: Mapping[str, Any],
        *,
        evidence: Sequence[Mapping[str, Any]],
    ) -> dict[str, Any]:
        """Record verified KiotViet history evidence without re-clicking carts.

        This is deliberately *not* a completion operation: it stores a
        durable partial proof in the event log so a stopped batch can resume
        only its still-open tabs.  The same-day ledger is written later by
        ``complete_automated_batch`` only after every worksheet has a final
        order proof.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "automation_attention":
            raise InvalidTransitionError(
                "External reconciliation is allowed only for a stopped automated batch.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )
        if not isinstance(submission_result, Mapping):
            raise WorkflowError("External reconciliation needs a result object.")
        container = submission_result.get("submitted")
        entries = container.get("submitted") if isinstance(container, Mapping) else None
        if not isinstance(entries, Sequence) or isinstance(entries, (str, bytes)) or not entries:
            raise WorkflowError("External reconciliation needs one or more submitted order entries.")
        if not isinstance(evidence, Sequence) or isinstance(evidence, (str, bytes)) or not evidence:
            raise WorkflowError("External reconciliation needs KiotViet history evidence.")

        evidence_by_code = {
            str(item.get("orderCode") or item.get("order_code") or "").strip(): dict(item)
            for item in evidence
            if isinstance(item, Mapping)
            and str(item.get("orderCode") or item.get("order_code") or "").strip()
        }
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            expected_rows = connection.execute(
                """
                SELECT bo.sheet_index, bo.expected_tab_label, o.order_code
                FROM batch_orders bo JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ? ORDER BY bo.sheet_index
                """,
                (batch_id,),
            ).fetchall()
            expected = {str(row["order_code"]): dict(row) for row in expected_rows}
            reconciled: list[dict[str, Any]] = []
            seen: set[str] = set()
            for item in entries:
                if not isinstance(item, Mapping) or item.get("success") is not True:
                    raise WorkflowError("Every reconciled order must be marked as successful.")
                code = str(item.get("orderCode") or item.get("order_code") or "").strip()
                if not code or code in seen or code not in expected or code not in evidence_by_code:
                    raise WorkflowError("External evidence does not match this persisted KiotViet batch.")
                seen.add(code)
                row = expected[code]
                actual_label = str(item.get("actualTabLabel") or item.get("tabLabel") or "").strip()
                self._require_expected_tab_label(
                    batch_id=batch_id,
                    sheet_index=int(row["sheet_index"]),
                    expected_tab_label=str(row["expected_tab_label"] or ""),
                    actual_tab_label=actual_label,
                )
                reconciled.append(dict(item))
            if len(reconciled) >= int(batch["sheet_count"]):
                raise WorkflowError("Use normal automated completion for a complete batch result.")

            notes = submission_result.get("notes")
            if isinstance(notes, Mapping):
                self._apply_verified_note_payload(connection, batch_id, notes, now)
            message = "Verified KiotViet history was reconciled; only remaining open carts may resume."
            connection.execute(
                "UPDATE batches SET status='automation_attention', error_message=?, updated_at=? WHERE id=?",
                (message, now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status='automation_attention', error_message=?, updated_at=? WHERE id=?",
                (message, now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.external_reconciliation",
                level="warning",
                message=message,
                payload={
                    "result": dict(submission_result),
                    "evidence": [evidence_by_code[code] for code in sorted(seen)],
                },
                created_at=now,
            )
        return self.state(int(run["id"]))

    def prepare_partial_automation_recovery(self, batch_id: int) -> dict[str, Any] | None:
        """Re-open an interrupted batch without re-uploading or duplicating.

        Edge returns an exact list of tabs it already verified as submitted
        before a later tab fails.  If those imported tabs are still open, the
        next UI run may resume only the remaining tabs.  This method consumes
        that persisted proof, re-applies the verified note mappings, and
        returns the remaining trusted payload.  It deliberately returns
        ``None`` when no partial proof exists; callers must then use the
        explicit cleanup/retry flow.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "automation_attention":
            return None
        events = self.db.fetch_all(
            """
            SELECT payload_json FROM events
            WHERE batch_id = ? AND event_type IN ('batch.automation_attention', 'batch.external_reconciliation')
            ORDER BY id DESC LIMIT 20
            """,
            (batch_id,),
        )
        # A retry can verify additional tabs before its transport response is
        # lost.  Union trusted successes across recent attention events so a
        # later proof never causes an earlier externally completed order to
        # be submitted again.
        result = None
        entries_by_code: dict[str, dict[str, Any]] = {}
        note_verified_by_code: dict[str, dict[str, Any]] = {}
        for event in events:
            try:
                envelope = json.loads(str(event["payload_json"] or "{}"))
            except (TypeError, ValueError):
                continue
            candidate = envelope.get("result") if isinstance(envelope, Mapping) else None
            container = candidate.get("submitted") if isinstance(candidate, Mapping) else None
            candidate_entries = container.get("submitted") if isinstance(container, Mapping) else None
            if not isinstance(candidate_entries, Sequence) or isinstance(candidate_entries, (str, bytes)):
                continue
            if result is None:
                result = dict(candidate)
            for item in candidate_entries:
                if not isinstance(item, Mapping) or item.get("success") is not True:
                    continue
                code = str(item.get("orderCode") or item.get("order_code") or "").strip()
                if code and code not in entries_by_code:
                    entries_by_code[code] = dict(item)
            candidate_notes = candidate.get("notes") if isinstance(candidate, Mapping) else None
            candidate_verified = candidate_notes.get("verified") if isinstance(candidate_notes, Mapping) else None
            if isinstance(candidate_verified, Sequence) and not isinstance(candidate_verified, (str, bytes)):
                for item in candidate_verified:
                    if not isinstance(item, Mapping):
                        continue
                    code = str(item.get("orderCode") or item.get("order_code") or "").strip()
                    if code and code not in note_verified_by_code:
                        note_verified_by_code[code] = dict(item)
        if not entries_by_code:
            return None
        successful = list(entries_by_code.values())
        if not successful:
            return None
        success_codes = {
            str(item.get("orderCode") or item.get("order_code") or "").strip()
            for item in successful
        }
        if not success_codes or "" in success_codes:
            return None
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            rows = connection.execute(
                """
                SELECT bo.order_id, bo.sheet_index, bo.expected_tab_label,
                       o.order_code
                FROM batch_orders bo JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ? ORDER BY bo.sheet_index
                """,
                (batch_id,),
            ).fetchall()
            known = {str(row["order_code"]) for row in rows}
            # Ignore a malformed/stale proof from a prior transport failure;
            # only codes that belong to this persisted batch may influence
            # recovery.  A single unknown entry must not hide valid proofs
            # for other orders.
            success_codes &= known
            successful = [
                item
                for item in successful
                if str(item.get("orderCode") or item.get("order_code") or "").strip() in success_codes
            ]
            if not success_codes or len(successful) >= int(batch["sheet_count"]):
                return None
            if any(
                connection.execute(
                    "SELECT 1 FROM confirmed_order_ledger WHERE business_date = ? AND order_code = ?",
                    (run["business_date"], code),
                ).fetchone()
                for code in success_codes
            ):
                return None
            notes = dict(result.get("notes") or {}) if isinstance(result, Mapping) else {}
            if note_verified_by_code:
                notes["ok"] = True
                notes["verified"] = [
                    # Preserve note proof only for carts that were also
                    # independently proven as submitted.  Remaining carts
                    # must be refilled after a browser/session recovery; an
                    # old tab number is not durable proof for a new import.
                    item for code, item in note_verified_by_code.items() if code in success_codes
                ]
            if isinstance(notes, Mapping):
                self._apply_verified_note_payload(connection, batch_id, notes, now)
            # These carts were already imported and their notes were
            # independently verified before the transport/session failure.
            # Mark the resumed state as submit-ready, not as a fresh import:
            # ``complete_automated_batch`` must not demand an impossible
            # second import result after the remaining tabs are submitted.
            connection.execute(
                """
                UPDATE batches
                SET status='waiting_user_confirmation',
                    imported_at=COALESCE(imported_at, ?),
                    review_ready_at=COALESCE(review_ready_at, ?),
                    error_message=NULL, updated_at=?
                WHERE id=?
                """,
                (now, now, now, batch_id),
            )
            connection.execute(
                "UPDATE batch_orders SET status=CASE WHEN note_verified=1 THEN 'note_filled' ELSE 'importing' END, updated_at=? WHERE batch_id=?",
                (now, batch_id),
            )
            connection.execute(
                "UPDATE orders SET status=CASE WHEN status='note_filled' THEN status ELSE 'importing' END, updated_at=? WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id=?)",
                (now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status='batch_in_progress', error_message=NULL, updated_at=? WHERE id=?",
                (now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.partial_recovery_started",
                message="Resuming only the remaining tabs from a verified partial Edge result.",
                payload={"already_submitted": sorted(success_codes)},
                created_at=now,
            )
            remaining_rows = [row for row in rows if str(row["order_code"]) not in success_codes]
            remaining = []
            for row in remaining_rows:
                item = connection.execute(
                    """
                    SELECT bo.sheet_index, bo.expected_tab_label, o.order_code,
                           o.customer_name,
                           (SELECT li.product_code FROM line_items li WHERE li.order_id=o.id ORDER BY li.sequence LIMIT 1) AS first_sku
                    FROM batch_orders bo JOIN orders o ON o.id=bo.order_id
                    WHERE bo.batch_id=? AND o.order_code=?
                    """,
                    (batch_id, row["order_code"]),
                ).fetchone()
                remaining.append({
                    "sheetIndex": int(item["sheet_index"]),
                    "orderCode": str(item["order_code"]),
                    "customerName": str(item["customer_name"] or ""),
                    "firstSku": str(item["first_sku"] or ""),
                    "expectedTabLabel": str(item["expected_tab_label"] or ""),
                })
        # ``result`` is the newest attention event, which can contain only a
        # single later reconciliation.  The queue must receive the complete
        # union here; otherwise completing the remaining carts would omit
        # earlier, independently verified submissions from its final atomic
        # ledger write.
        prior_result = dict(result) if isinstance(result, Mapping) else {}
        prior_submitted = prior_result.get("submitted")
        merged_submitted = dict(prior_submitted) if isinstance(prior_submitted, Mapping) else {}
        merged_submitted.update(
            {"ok": True, "allSubmitted": False, "submitted": successful}
        )
        prior_result["submitted"] = merged_submitted
        prior_result["notes"] = notes
        return {
            "batchId": int(batch_id),
            "runId": int(run["id"]),
            "remainingOrders": remaining,
            "alreadySubmitted": successful,
            "priorResult": prior_result,
        }

    def build_partial_recovery_import(
        self,
        batch_id: int,
        remaining_order_codes: Sequence[str],
    ) -> dict[str, Any]:
        """Build a new workbook only for carts lost with a browser session.

        This is used after partial completion has been independently proven
        in KiotViet history but Edge has restarted and no longer retains the
        imported carts.  It never rebuilds the already-submitted sheets, so
        it cannot re-create their external orders.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] != "waiting_user_confirmation":
            raise InvalidTransitionError(
                "A partial recovery workbook is allowed only after verified recovery preparation.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )
        if not isinstance(remaining_order_codes, Sequence) or isinstance(
            remaining_order_codes, (str, bytes)
        ):
            raise WorkflowError("Partial recovery needs a list of remaining order codes.")
        requested = [str(code or "").strip() for code in remaining_order_codes]
        if not requested or len(set(requested)) != len(requested):
            raise WorkflowError("Partial recovery order codes are invalid.")

        rows = self.db.fetch_all(
            """
            SELECT bo.sheet_index, bo.order_id, o.order_code, o.customer_name,
                   (SELECT COUNT(*) FROM line_items li WHERE li.order_id=o.id) AS line_count,
                   (SELECT li.product_code FROM line_items li WHERE li.order_id=o.id
                    ORDER BY li.sequence LIMIT 1) AS first_sku
            FROM batch_orders bo JOIN orders o ON o.id=bo.order_id
            WHERE bo.batch_id=? ORDER BY bo.sheet_index
            """,
            (batch_id,),
        )
        by_code = {str(row["order_code"]): row for row in rows}
        if any(code not in by_code for code in requested):
            raise WorkflowError("Partial recovery contains an order outside its persisted batch.")
        selected_rows = [row for row in rows if str(row["order_code"]) in set(requested)]
        if not selected_rows or len(selected_rows) >= int(batch["sheet_count"]):
            raise WorkflowError("Partial recovery must rebuild only a non-empty proper subset of the batch.")

        builder_rows = {
            str(item["order_code"]): item
            for item in self._builder_orders_for_batch(batch_id)
        }
        build_orders = [builder_rows[str(row["order_code"])] for row in selected_rows]
        output = (
            self.config.outputs_dir
            / str(run["business_date"])
            / str(run["run_key"])
            / f"KiotViet_{run['business_date']}_batch_{int(batch['batch_number']):02d}_recovery_{len(build_orders):02d}.xlsx"
        )
        try:
            build_result = self.workbook_builder.build(build_orders, output)
        except WorkbookBuildError as exc:
            raise WorkflowError(str(exc), details={"batch_id": batch_id}) from exc

        now = self._iso_now()
        orders: list[dict[str, Any]] = []
        with self.db.transaction(immediate=True) as connection:
            current = connection.execute("SELECT * FROM batches WHERE id=?", (batch_id,)).fetchone()
            if current is None or current["status"] != "waiting_user_confirmation":
                raise InvalidTransitionError(
                    "The batch changed before its partial recovery workbook was recorded.",
                    details={"batch_id": batch_id},
                )
            for index, row in enumerate(selected_rows, start=1):
                # Unicode escapes avoid accidental mojibake if the source is
                # later touched by an ANSI-configured Windows editor.
                expected_tab_label = f"\u0110\u1eb7t h\u00e0ng {index + 1}"
                connection.execute(
                    """
                    UPDATE batch_orders
                    SET expected_tab_label=?, actual_tab_label=NULL, note_verified=0,
                        status='importing', updated_at=?
                    WHERE batch_id=? AND order_id=?
                    """,
                    (expected_tab_label, now, batch_id, int(row["order_id"])),
                )
                orders.append(
                    {
                        "sheetIndex": int(row["sheet_index"]),
                        "orderCode": str(row["order_code"]),
                        "customerName": str(row["customer_name"] or ""),
                        "firstSku": str(row["first_sku"] or ""),
                        "lineCount": int(row["line_count"] or 0),
                        "expectedTabLabel": expected_tab_label,
                    }
                )
            connection.execute(
                "UPDATE batches SET status='importing', error_message=NULL, updated_at=? WHERE id=?",
                (now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status='batch_in_progress', error_message=NULL, updated_at=? WHERE id=?",
                (now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.partial_reimport_ready",
                message="Built a recovery workbook only for carts not confirmed in KiotViet.",
                payload={
                    "workbook_path": str(build_result.output_path),
                    "already_submitted": [
                        str(row["order_code"])
                        for row in rows
                        if str(row["order_code"]) not in {str(item["order_code"]) for item in selected_rows}
                    ],
                    "remaining_order_codes": [str(row["order_code"]) for row in selected_rows],
                    "sheet_count": len(orders),
                },
                created_at=now,
            )
        return {
            "batchId": int(batch_id),
            "runId": int(run["id"]),
            "batchNumber": int(batch["batch_number"]),
            "businessDate": str(run["business_date"]),
            "workbookPath": str(build_result.output_path),
            "orders": orders,
            "recovery": True,
        }

    def retry_batch(self, batch_id: int) -> dict[str, Any]:
        """Explicitly reset a failed browser attempt after the user cleans tabs."""

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._assert_run_not_cancelled(run)
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] == "ready":
            return self.state(int(run["id"]))
        if batch["status"] not in {"failed", "automation_attention"}:
            raise InvalidTransitionError(
                "Only a failed or automation-attention batch can be retried.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )
        workbook_path = Path(str(batch.get("workbook_path") or ""))
        if not workbook_path.is_file():
            return self.prepare_batches(int(run["id"]))

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                """
                UPDATE batches
                SET status = 'ready', imported_at = NULL,
                    review_ready_at = NULL, user_confirmed_at = NULL,
                    error_message = NULL, updated_at = ?
                WHERE id = ?
                """,
                (now, batch_id),
            )
            connection.execute(
                """
                UPDATE batch_orders
                SET actual_tab_label = NULL, note_verified = 0,
                    status = 'ready', updated_at = ?
                WHERE batch_id = ?
                """,
                (now, batch_id),
            )
            connection.execute(
                """
                UPDATE orders SET status = 'batch_ready', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status = 'batches_ready', error_message = NULL, updated_at = ? WHERE id = ?",
                (now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.retry_authorized",
                level="warning",
                message="User confirmed cleanup and authorized one retry.",
                created_at=now,
            )
        return self.state(int(run["id"]))

    def abort_run(
        self,
        run_id: int,
        *,
        confirm: bool = False,
        kiot_tabs_cleared: bool = False,
        acknowledge_external_order_risk: bool = False,
    ) -> dict[str, Any]:
        """Terminally stop a workflow without deleting its audit records.

        Closing KiotViet carts does not prove that an earlier final-order click
        did not succeed.  The operator must therefore explicitly attest that
        the carts were cleared and accept that external orders are outside the
        SQLite ledger's visibility.  Confirmed batches and ledger rows remain
        untouched; only the unconfirmed part becomes terminally cancelled.
        """

        if confirm is not True:
            raise WorkflowError("Aborting a workflow requires explicit confirmation.")
        if kiot_tabs_cleared is not True:
            raise WorkflowError("Confirm that the unfinished KiotViet tabs were cleared before aborting.")
        if acknowledge_external_order_risk is not True:
            raise WorkflowError(
                "Acknowledge that cancelling SQLite cannot undo an external KiotViet order."
            )

        run = self._resolve_run(int(run_id))
        allowed_statuses = set(ACTIVE_RUN_STATUSES) | {"failed"}
        if str(run.get("status") or "") == "cancelled":
            return self.state(int(run["id"]))
        if str(run.get("status") or "") not in allowed_statuses:
            raise InvalidTransitionError(
                "Only an unfinished or failed workflow can be aborted.",
                details={"run_id": int(run_id), "status": run.get("status")},
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            current = connection.execute("SELECT * FROM runs WHERE id = ?", (int(run_id),)).fetchone()
            if current is None:  # pragma: no cover - FK invariant
                raise WorkflowError("The workflow disappeared while it was being aborted.")
            current_run = dict(current)
            current_status = str(current_run.get("status") or "")
            if current_status == "cancelled":
                raise InvalidTransitionError(
                    "The workflow was already cancelled by another request.",
                    details={"run_id": int(run_id), "status": current_status},
                )
            if current_status not in allowed_statuses:
                raise InvalidTransitionError(
                    "The workflow changed state before it could be aborted.",
                    details={"run_id": int(run_id), "status": current_status},
                )

            unconfirmed_batch_count = int(
                connection.execute(
                    "SELECT COUNT(*) FROM batches WHERE run_id = ? AND status != 'confirmed'",
                    (int(run_id),),
                ).fetchone()[0]
            )
            unconfirmed_order_count = int(
                connection.execute(
                    """
                    SELECT COUNT(*)
                    FROM orders o
                    WHERE o.run_id = ?
                      AND NOT EXISTS (
                          SELECT 1
                          FROM batch_orders bo
                          JOIN batches b ON b.id = bo.batch_id
                          WHERE bo.order_id = o.id AND b.status = 'confirmed'
                      )
                    """,
                    (int(run_id),),
                ).fetchone()[0]
            )

            connection.execute(
                """
                UPDATE batch_orders
                SET status = 'cancelled', updated_at = ?
                WHERE batch_id IN (
                    SELECT id FROM batches WHERE run_id = ? AND status != 'confirmed'
                )
                """,
                (now, int(run_id)),
            )
            connection.execute(
                """
                UPDATE batches
                SET status = 'cancelled', updated_at = ?
                WHERE run_id = ? AND status != 'confirmed'
                """,
                (now, int(run_id)),
            )
            connection.execute(
                """
                UPDATE orders
                SET status = 'cancelled', updated_at = ?
                WHERE run_id = ?
                  AND NOT EXISTS (
                      SELECT 1
                      FROM batch_orders bo
                      JOIN batches b ON b.id = bo.batch_id
                      WHERE bo.order_id = orders.id AND b.status = 'confirmed'
                  )
                """,
                (now, int(run_id)),
            )
            connection.execute(
                """
                UPDATE runs
                SET status = 'cancelled',
                    error_message = 'Workflow aborted by operator after KiotViet cleanup.',
                    updated_at = ?
                WHERE id = ?
                """,
                (now, int(run_id)),
            )
            self._insert_event(
                connection,
                run_id=int(run_id),
                event_type="run.aborted",
                level="warning",
                message="Workflow was aborted after the operator cleared unfinished KiotViet tabs.",
                payload={
                    "unconfirmed_batch_count": unconfirmed_batch_count,
                    "unconfirmed_order_count": unconfirmed_order_count,
                    "orders_deleted": False,
                    "ledger_changed": False,
                },
                created_at=now,
            )
        return self.state(int(run_id))

    def delete_selected_run_orders(
        self,
        run_id: int,
        order_codes: Sequence[str],
        *,
        confirm: bool = False,
    ) -> dict[str, Any]:
        """Forget every explicitly selected order in a failed run.

        Deletion is deliberately all-or-nothing for one run.  This keeps the
        persisted batch/workbook mapping coherent: a user must tick every
        code in the blocked run before its order rows and batches are removed.
        Confirmed ledger rows are never deletable through this UI action.
        """

        if confirm is not True:
            raise WorkflowError("Xóa mã cần xác nhận rõ ràng từ giao diện.")
        if not isinstance(order_codes, Sequence) or isinstance(order_codes, (str, bytes)):
            raise WorkflowError("Danh sách mã cần xóa không hợp lệ.")
        selected = sorted({str(code or "").strip() for code in order_codes if str(code or "").strip()})
        if not selected:
            raise WorkflowError("Hãy tích chọn ít nhất một mã trước khi xóa.")

        run = self._resolve_run(int(run_id))
        if str(run.get("status") or "") not in {"automation_attention", "failed", "cancelled", "expired"}:
            raise InvalidTransitionError(
                "Chỉ được xóa mã trong workflow lỗi hoặc đã dừng.",
                details={"run_id": int(run_id), "status": run.get("status")},
            )
        rows = self.db.fetch_all(
            "SELECT id, order_code, status FROM orders WHERE run_id = ? ORDER BY sequence",
            (int(run_id),),
        )
        actual = [str(row["order_code"]) for row in rows]
        if set(selected) != set(actual):
            missing = [code for code in actual if code not in selected]
            extra = [code for code in selected if code not in set(actual)]
            raise InvalidTransitionError(
                "Để giữ nguyên cấu trúc lô, hãy tích chọn đủ toàn bộ mã của workflow này.",
                details={"run_id": int(run_id), "missing_codes": missing, "unknown_codes": extra},
            )

        ledger = self.db.fetch_all(
            """
            SELECT order_code FROM confirmed_order_ledger
            WHERE run_id = ? AND order_code IN ({})
            """.format(",".join("?" for _ in selected)),
            (int(run_id), *selected),
        )
        if ledger:
            protected = [str(row["order_code"]) for row in ledger]
            raise InvalidTransitionError(
                "Có mã đã được ghi nhận hoàn tất; tool từ chối xóa để tránh chạy trùng.",
                details={"protected_codes": protected},
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            # Keep the run row as an audit marker while removing the actual
            # order/batch records.  Existing workbooks remain on disk for
            # audit; no filesystem deletion is performed here.
            connection.execute("DELETE FROM batches WHERE run_id = ?", (int(run_id),))
            connection.execute("DELETE FROM orders WHERE run_id = ?", (int(run_id),))
            connection.execute(
                """
                UPDATE runs
                SET status = 'cancelled', order_count = 0, line_count = 0,
                    duplicate_count = 0, error_message = ?, updated_at = ?
                WHERE id = ?
                """,
                ("User deleted all explicitly selected order codes for a retry.", now, int(run_id)),
            )
            self._insert_event(
                connection,
                run_id=int(run_id),
                event_type="orders.deleted_selected",
                level="warning",
                message="User deleted all selected order codes from SQLite.",
                payload={"order_codes": selected, "count": len(selected)},
                created_at=now,
            )
        return {
            "ok": True,
            "run_id": int(run_id),
            "deleted_order_codes": selected,
            "deleted_count": len(selected),
            "state": self.state(int(run_id)),
        }

    def reconcile_completed_import(
        self, batch_id: int, notes_payload: Mapping[str, Any]
    ) -> dict[str, Any]:
        """Record a verified KiotViet import after a local timeout.

        This accepts only a failed batch after the extension has re-visited
        every expected tab and verified its note.  It never uploads a file.
        """

        batch = self._resolve_batch(batch_id=batch_id)
        run = self._resolve_run(int(batch["run_id"]))
        self._guard_business_date(run)
        self._assert_batch_unlocked(batch)
        if batch["status"] not in {"failed", "automation_attention"}:
            raise InvalidTransitionError(
                "Only a failed or automation-attention batch can be reconciled.",
                details={"batch_id": batch_id, "status": batch["status"]},
            )

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            self._apply_verified_note_payload(connection, batch_id, notes_payload, now)
            verified = int(
                connection.execute(
                    "SELECT COUNT(*) FROM batch_orders WHERE batch_id = ? AND note_verified = 1",
                    (batch_id,),
                ).fetchone()[0]
            )
            if verified != int(batch["sheet_count"]):
                raise InvalidTransitionError(
                    "Cannot reconcile until every imported tab has a verified note.",
                    details={
                        "batch_id": batch_id,
                        "verified": verified,
                        "expected": batch["sheet_count"],
                    },
                )
            connection.execute(
                """
                UPDATE batches
                SET status = 'waiting_user_confirmation', imported_at = ?,
                    review_ready_at = ?, error_message = NULL, updated_at = ?
                WHERE id = ?
                """,
                (now, now, now, batch_id),
            )
            connection.execute(
                """
                UPDATE orders SET status = 'waiting_user_confirmation', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status = 'waiting_user_confirmation', error_message = NULL, updated_at = ? WHERE id = ?",
                (now, run["id"]),
            )
            self._insert_event(
                connection,
                run_id=int(run["id"]),
                batch_id=batch_id,
                event_type="batch.import_reconciled",
                level="warning",
                message="Verified KiotViet tabs reconciled after a local timeout; no second upload was performed.",
                payload={"verified_notes": verified},
                created_at=now,
            )
        return self.state(int(run["id"]))

    def mark_extension_event(
        self,
        event_type: str,
        payload: Mapping[str, Any] | None = None,
        *,
        run_id: int | None = None,
        batch_id: int | None = None,
        order_code: str | None = None,
        level: str = "info",
        message: str = "",
    ) -> dict[str, Any]:
        """Persist a browser-extension observation and apply safe transitions."""

        event_type = str(event_type or "").strip().casefold()
        if not re.fullmatch(r"[a-z0-9_.-]{1,80}", event_type):
            raise WorkflowError("Invalid extension event type.")
        if any(fragment in event_type for fragment in PROHIBITED_EVENT_FRAGMENTS):
            raise WorkflowError(
                "Final order submission events are prohibited in safe mode."
            )
        payload_dict = dict(payload or {})
        canonical = event_type.removeprefix("extension.")

        run = self._resolve_run(run_id) if (run_id is not None or self._active_run_row()) else None
        if run is not None and canonical in MUTATING_EXTENSION_EVENTS:
            self._assert_run_not_cancelled(run)
            self._guard_business_date(run)

        if batch_id is None and payload_dict.get("batch_id") is not None:
            batch_id = int(payload_dict["batch_id"])
        if batch_id is None and payload_dict.get("batch_number") is not None and run is not None:
            batch = self._resolve_batch(
                run_id=int(run["id"]), batch_number=int(payload_dict["batch_number"])
            )
        elif batch_id is not None:
            batch = self._resolve_batch(batch_id=int(batch_id))
        else:
            batch = None
        if batch is not None:
            batch_id = int(batch["id"])
            if run is None:
                run = self._resolve_run(int(batch["run_id"]))
            elif int(batch["run_id"]) != int(run["id"]):
                raise WorkflowError("Batch does not belong to the selected run.")
            if canonical in MUTATING_EXTENSION_EVENTS:
                self._assert_run_not_cancelled(run)
                self._guard_business_date(run)
            if canonical in {
                "batch.import_started",
                "batch.imported",
                "order.note_filled",
                "batch.notes_filled",
                "batch.review_ready",
                "batch.waiting_user_confirmation",
                "batch.error",
            }:
                self._assert_batch_unlocked(batch)

        if order_code is None:
            order_code = payload_dict.get("order_code") or payload_dict.get("orderCode")
        order = None
        if order_code and run is not None:
            order = self.db.fetch_one(
                "SELECT * FROM orders WHERE run_id = ? AND order_code = ?",
                (run["id"], str(order_code)),
            )
            if order is None:
                raise WorkflowError(f"Unknown order code: {order_code}")

        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            if run is not None and (
                canonical in MUTATING_EXTENSION_EVENTS
                or (canonical == "workflow.error" and bool(payload_dict.get("fatal")))
            ):
                current_run = connection.execute(
                    "SELECT * FROM runs WHERE id = ?", (int(run["id"]),)
                ).fetchone()
                if current_run is None:  # pragma: no cover - FK invariant
                    raise WorkflowError("The workflow disappeared before the extension event was recorded.")
                self._assert_run_not_cancelled(dict(current_run))
            if canonical == "batch.import_started":
                self._require_batch_status(batch, {"ready"}, canonical)
                connection.execute(
                    "UPDATE batches SET status = 'importing', updated_at = ? WHERE id = ?",
                    (now, batch_id),
                )
                connection.execute(
                    "UPDATE runs SET status = 'batch_in_progress', error_message = NULL, updated_at = ? WHERE id = ?",
                    (now, run["id"]),
                )
            elif canonical == "batch.imported":
                self._require_batch_status(batch, {"ready", "importing"}, canonical)
                connection.execute(
                    "UPDATE batches SET status = 'imported', imported_at = ?, updated_at = ? WHERE id = ?",
                    (now, now, batch_id),
                )
                connection.execute(
                    """
                    UPDATE orders SET status = 'imported', updated_at = ?
                    WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                    """,
                    (now, batch_id),
                )
                connection.execute(
                    "UPDATE runs SET status = 'batch_in_progress', error_message = NULL, updated_at = ? WHERE id = ?",
                    (now, run["id"]),
                )
            elif canonical == "order.note_filled":
                self._require_batch_status(
                    batch, {"imported", "notes_in_progress"}, canonical
                )
                if order is None:
                    raise WorkflowError("order.note_filled requires order_code.")
                actual_tab = str(
                    payload_dict.get("tab_label")
                    or payload_dict.get("tabLabel")
                    or ""
                ).strip()
                mapping = connection.execute(
                    """
                    SELECT sheet_index, expected_tab_label
                    FROM batch_orders
                    WHERE batch_id = ? AND order_id = ?
                    """,
                    (batch_id, order["id"]),
                ).fetchone()
                if mapping is None:
                    raise WorkflowError("Order is not assigned to the selected batch.")
                self._require_expected_tab_label(
                    batch_id=int(batch_id),
                    sheet_index=int(mapping["sheet_index"]),
                    expected_tab_label=str(mapping["expected_tab_label"] or ""),
                    actual_tab_label=actual_tab,
                )
                connection.execute(
                    "UPDATE orders SET status = 'note_filled', updated_at = ? WHERE id = ?",
                    (now, order["id"]),
                )
                connection.execute(
                    """
                    UPDATE batch_orders
                    SET note_verified = 1, status = 'note_filled',
                        actual_tab_label = COALESCE(NULLIF(?, ''), actual_tab_label),
                        updated_at = ?
                    WHERE batch_id = ? AND order_id = ?
                    """,
                    (actual_tab, now, batch_id, order["id"]),
                )
                connection.execute(
                    "UPDATE batches SET status = 'notes_in_progress', updated_at = ? WHERE id = ?",
                    (now, batch_id),
                )
            elif canonical in {
                "batch.notes_filled",
                "batch.review_ready",
                "batch.waiting_user_confirmation",
            }:
                self._require_batch_status(
                    batch, {"imported", "notes_in_progress"}, canonical
                )
                self._apply_verified_note_payload(connection, batch_id, payload_dict, now)
                verified = int(
                    connection.execute(
                        "SELECT COUNT(*) FROM batch_orders WHERE batch_id = ? AND note_verified = 1",
                        (batch_id,),
                    ).fetchone()[0]
                )
                if verified != int(batch["sheet_count"]):
                    raise InvalidTransitionError(
                        "Not every imported tab has a verified order note.",
                        details={
                            "batch_id": batch_id,
                            "verified": verified,
                            "expected": batch["sheet_count"],
                        },
                    )
                connection.execute(
                    """
                    UPDATE batches
                    SET status = 'waiting_user_confirmation', review_ready_at = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, now, batch_id),
                )
                connection.execute(
                    """
                    UPDATE orders SET status = 'waiting_user_confirmation', updated_at = ?
                    WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                    """,
                    (now, batch_id),
                )
                connection.execute(
                    "UPDATE runs SET status = 'waiting_user_confirmation', error_message = NULL, updated_at = ? WHERE id = ?",
                    (now, run["id"]),
                )
            elif canonical == "batch.error":
                self._require_batch_status(
                    batch,
                    {"ready", "importing", "imported", "notes_in_progress"},
                    canonical,
                )
                error = str(payload_dict.get("error") or message or "Extension batch error")
                connection.execute(
                    "UPDATE batches SET status = 'failed', error_message = ?, updated_at = ? WHERE id = ?",
                    (error, now, batch_id),
                )
                connection.execute(
                    "UPDATE runs SET status = 'failed', error_message = ?, updated_at = ? WHERE id = ?",
                    (error, now, run["id"]),
                )
                level = "error"
            elif canonical == "workflow.error" and run is not None:
                error = str(payload_dict.get("error") or message or "Extension workflow error")
                if bool(payload_dict.get("fatal")):
                    connection.execute(
                        "UPDATE runs SET status = 'failed', error_message = ?, updated_at = ? WHERE id = ?",
                        (error, now, run["id"]),
                    )
                level = "error"

            self._insert_event(
                connection,
                run_id=int(run["id"]) if run is not None else None,
                batch_id=batch_id,
                order_id=int(order["id"]) if order is not None else None,
                event_type=event_type,
                level=level,
                message=message,
                payload=payload_dict,
                created_at=now,
            )
        return self.state(int(run["id"])) if run is not None else self.state()

    # Compatibility alias for route code that uses a pluralized name.
    mark_extension_events = mark_extension_event

    # ------------------------------------------------------------------
    # Incremental report scanning and confirmed-order ledger
    # ------------------------------------------------------------------
    def _read_source_for_incremental_scan(
        self,
        source_path: str | Path,
        *,
        expected_date: date | None = None,
    ) -> dict[str, Any]:
        source = Path(source_path).expanduser().resolve()
        if not source.is_file():
            raise SourceValidationError(f"Source report was not found: {source}")
        if source.suffix.casefold() not in SUPPORTED_REPORT_EXTENSIONS:
            raise SourceValidationError(
                f"The Sopoka source must be an Excel file ({SUPPORTED_REPORT_FORMATS})."
            )

        stat = source.stat()
        source_hash = sha256_file(source)
        parsed = self._parse_source(source, expected_date=expected_date)
        report = _mapping(parsed)
        business_date = _date_value(
            report.get("report_start_date") or getattr(parsed, "report_start_date", None),
            field="report_start_date",
        )
        report_end = _date_value(
            report.get("report_end_date") or getattr(parsed, "report_end_date", None),
            field="report_end_date",
        )
        if business_date != report_end:
            raise SourceValidationError("The source report must cover one day only.")
        if expected_date is None and business_date != self._today():
            raise MidnightGuardError(
                "The report is no longer for today's Thailand business date.",
                details={
                    "report_date": business_date.isoformat(),
                    "today": self._today().isoformat(),
                },
            )
        parsed_orders_value = report.get("orders") or getattr(parsed, "orders", None)
        if not isinstance(parsed_orders_value, Sequence) or isinstance(
            parsed_orders_value, (str, bytes)
        ):
            raise SourceValidationError("Parser returned no order list.")
        orders = [
            self._normalize_order(order, index)
            for index, order in enumerate(parsed_orders_value, 1)
        ]
        if not orders:
            raise SourceValidationError("Today's report contains no orders.")
        if len({order["order_code"] for order in orders}) != len(orders):
            raise SourceValidationError("The source contains duplicate order codes.")
        parser_metadata = {
            key: report.get(key)
            for key in (
                "source_file",
                "sheet_name",
                "header_row",
                "first_data_row",
                "last_data_row",
                "report_start_date",
                "report_end_date",
                "qa",
            )
            if key in report
        }
        return {
            "source": source,
            "stat": stat,
            "source_hash": source_hash,
            "business_date": business_date,
            "parser_metadata": parser_metadata,
            "orders": orders,
        }

    def _active_run_in_connection(
        self, connection: sqlite3.Connection, business_date: date
    ) -> sqlite3.Row | None:
        placeholders = ",".join("?" for _ in ACTIVE_RUN_STATUSES)
        return connection.execute(
            f"""
            SELECT * FROM runs
            WHERE business_date = ? AND status IN ({placeholders})
            ORDER BY id DESC LIMIT 1
            """,
            (business_date.isoformat(), *tuple(sorted(ACTIVE_RUN_STATUSES))),
        ).fetchone()

    def _insert_incremental_run(
        self,
        connection: sqlite3.Connection,
        *,
        source: Path,
        stat: Any,
        source_hash: str,
        business_date: date,
        parser_metadata: Mapping[str, Any],
        orders: Sequence[Mapping[str, Any]],
        scan_id: int,
        source_order_count: int,
        historical_run: bool,
        now: str,
    ) -> int:
        """Persist only the new orders selected by one report scan."""

        run_key = uuid.uuid4().hex
        metadata = dict(parser_metadata)
        metadata["incremental_scan_id"] = scan_id
        metadata["filtered_from_source_order_count"] = int(source_order_count)
        line_count = sum(len(order["products"]) for order in orders)
        cursor = connection.execute(
            """
            INSERT INTO runs(
                run_key, business_date, source_path, source_name,
                source_sha256, source_size, source_mtime_ns,
                parser_metadata_json, status, order_count, line_count,
                duplicate_count, duplicate_file, duplicates_allowed,
                historical_run, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'parsed', ?, ?, 0, 0, 0, ?, ?, ?)
            """,
            (
                run_key,
                business_date.isoformat(),
                str(source),
                source.name,
                source_hash,
                int(stat.st_size),
                int(stat.st_mtime_ns),
                json_dumps(metadata),
                len(orders),
                line_count,
                int(historical_run),
                now,
                now,
            ),
        )
        run_id = int(cursor.lastrowid)
        for filtered_sequence, order_value in enumerate(orders, 1):
            order = dict(order_value)
            order["sequence"] = filtered_sequence
            self._insert_normalized_order(
                connection, run_id=run_id, order=order, status="parsed", now=now
            )
        self._insert_event(
            connection,
            run_id=run_id,
            event_type="run.incremental_created",
            message="New-only workflow created from an audited SOPoka scan.",
            payload={
                "scan_id": scan_id,
                "new_order_count": len(orders),
                "line_count": line_count,
            },
            created_at=now,
        )
        return run_id

    def _insert_normalized_order(
        self,
        connection: sqlite3.Connection,
        *,
        run_id: int,
        order: Mapping[str, Any],
        status: str,
        now: str,
    ) -> int:
        order_cursor = connection.execute(
            """
            INSERT INTO orders(
                run_id, sequence, order_code, payload_sha256,
                source_start_row, source_end_row, customer_code,
                customer_name, customer_address, order_date, status,
                payload_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                int(order["sequence"]),
                order["order_code"],
                order["payload_sha256"],
                order.get("source_start_row"),
                order.get("source_end_row"),
                order["customer_code"],
                order["customer_name"],
                order["customer_address"],
                order.get("order_date"),
                status,
                json_dumps(order["raw"]),
                now,
                now,
            ),
        )
        order_id = int(order_cursor.lastrowid)
        for item_index, item in enumerate(order["products"], 1):
            connection.execute(
                """
                INSERT INTO line_items(
                    order_id, sequence, source_row, product_code,
                    product_name, quantity, unit_price, gross_amount,
                    discount_percent, payload_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    order_id,
                    item_index,
                    item.get("source_row"),
                    item["product_code"],
                    item["product_name"],
                    _decimal_text(item["quantity"]),
                    _decimal_text(item["unit_price"]),
                    _decimal_text(item.get("gross_amount", 0)),
                    _decimal_text(item["discount_percent"]),
                    json_dumps(item["raw"]),
                ),
            )
        return order_id

    @staticmethod
    def _scan_status_message(
        status: str,
        new_count: int,
        skipped_count: int,
        unresolved_count: int,
        changed_count: int,
    ) -> str:
        if status == "run_created":
            return f"Đã tạo workflow mới cho {new_count} mã chưa xử lý; bỏ qua {skipped_count} mã đã xác nhận."
        if status == "no_new":
            return f"Không có mã mới; đã bỏ qua an toàn {skipped_count} mã đã xác nhận."
        if status == "unresolved":
            return f"Dừng an toàn: {unresolved_count} mã đang ở workflow chưa được xác nhận."
        if status == "changed_confirmed":
            return f"Dừng an toàn: {changed_count} mã đã xác nhận có dữ liệu khác với report mới."
        if status == "active_workflow":
            return "Dừng an toàn: còn workflow đang hoạt động cần hoàn tất hoặc đối soát."
        return status

    def _write_scan_audit(self, scan_id: int) -> None:
        """Write a read-only audit companion; never mutate Sopoka's source file."""

        scan = self.scan_state(scan_id)
        if scan is None:
            return
        audit_dir = self.config.outputs_dir / str(scan["business_date"]) / "scans" / str(scan["scan_key"])
        audit_dir.mkdir(parents=True, exist_ok=True)
        audit_path = audit_dir / "scan_audit.json"
        audit_payload = {
            "note": (
                "Raw Sopoka report is retained unchanged. This audit defines the new-only working set "
                "used for KiotViet import."
            ),
            "scan": scan,
        }
        audit_path.write_text(
            json.dumps(
                json.loads(json_dumps(audit_payload)),
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                "UPDATE report_scans SET audit_path = ? WHERE id = ?",
                (str(audit_path), scan_id),
            )

    # ------------------------------------------------------------------
    # Parser and normalization
    # ------------------------------------------------------------------
    def _parse_source(self, source: Path, *, expected_date: date | None = None) -> Any:
        parser = self._parser
        if parser is None:
            try:
                from .report_parser import parse_report
            except ImportError as exc:  # pragma: no cover - defensive adapter
                raise SourceValidationError(
                    "automation_app.report_parser.parse_report is unavailable."
                ) from exc
            parser = parse_report
        try:
            signature = inspect.signature(parser)
            kwargs: dict[str, Any] = {}
            if "expected_date" in signature.parameters:
                # Preserve the original safety default for normal UI runs;
                # selected historical runs explicitly provide their date.
                kwargs["expected_date"] = expected_date or self._today()
            if "max_orders_per_batch" in signature.parameters:
                kwargs["max_orders_per_batch"] = self.config.max_orders_per_batch
            return parser(source, **kwargs)
        except WorkflowError:
            raise
        except Exception as exc:
            raise SourceValidationError(str(exc)) from exc

    def _normalize_order(self, value: Any, fallback_sequence: int) -> dict[str, Any]:
        raw = _mapping(value)
        code = str(
            raw.get("related_order_code")
            or raw.get("order_code")
            or raw.get("Mã đặt đơn hàng liên quan")
            or ""
        ).strip()
        if not code:
            raise SourceValidationError("A parsed order is missing its related order code.")
        items_value = raw.get("items") or raw.get("products") or raw.get("line_items")
        if not isinstance(items_value, Sequence) or isinstance(items_value, (str, bytes)):
            raise SourceValidationError(f"Order {code} has no product list.")
        products = [self._normalize_item(item) for item in items_value]
        if not 1 <= len(products) <= self.config.max_products_per_order:
            raise SourceValidationError(
                f"Order {code} has {len(products)} products; KiotViet permits at most {self.config.max_products_per_order}."
            )
        customer_code = str(raw.get("customer_code") or "").strip()
        customer_name = str(raw.get("customer_name") or "").strip()
        builder_payload = {
            "order_code": code,
            "customer_code": customer_code,
            "customer_name": customer_name,
            "products": [
                {
                    "product_code": item["product_code"],
                    "product_name": item["product_name"],
                    "unit_price": _json_number(item["unit_price"]),
                    "discount_percent": _json_number(item["discount_percent"]),
                    "quantity": _json_number(item["quantity"]),
                }
                for item in products
            ],
        }
        payload_sha256 = hashlib.sha256(
            json_dumps(builder_payload).encode("utf-8")
        ).hexdigest()
        order_date = raw.get("related_order_date") or raw.get("order_date")
        if isinstance(order_date, (date, datetime)):
            order_date = order_date.isoformat()
        elif order_date:
            order_date = str(order_date)
        return {
            "sequence": int(raw.get("sequence") or fallback_sequence),
            "order_code": code,
            "payload_sha256": payload_sha256,
            "source_start_row": raw.get("source_start_row"),
            "source_end_row": raw.get("source_end_row"),
            "customer_code": customer_code,
            "customer_name": customer_name,
            "customer_address": str(raw.get("customer_address") or "").strip(),
            "order_date": order_date,
            "products": products,
            "builder_payload": builder_payload,
            "raw": raw,
        }

    @staticmethod
    def _normalize_item(value: Any) -> dict[str, Any]:
        raw = _mapping(value)
        code = str(
            raw.get("product_code") or raw.get("sku") or raw.get("Mã sản phẩm") or ""
        ).strip()
        if not code:
            raise SourceValidationError("A product line is missing product_code.")
        return {
            "source_row": raw.get("source_row"),
            "product_code": code,
            "product_name": str(raw.get("product_name") or "").strip(),
            "quantity": raw.get("quantity", 0),
            "unit_price": raw.get("unit_price", 0),
            "gross_amount": raw.get("gross_amount", 0),
            "discount_percent": raw.get("discount_percent", 0),
            "raw": raw,
        }

    # ------------------------------------------------------------------
    # Persistence helpers
    # ------------------------------------------------------------------
    def _run_state(self, run_id: int) -> dict[str, Any]:
        run = self.db.fetch_one("SELECT * FROM runs WHERE id = ?", (run_id,))
        if run is None:
            raise WorkflowError(f"Run {run_id} was not found.")
        run["duplicate_file"] = bool(run["duplicate_file"])
        run["duplicates_allowed"] = bool(run["duplicates_allowed"])
        run["parser_metadata"] = _readable_json(
            run.pop("parser_metadata_json", "{}"), {}
        )

        orders = self.db.fetch_all(
            """
            SELECT o.*,
                   (SELECT COUNT(*) FROM line_items li WHERE li.order_id = o.id) AS product_count,
                   (SELECT li.product_code FROM line_items li WHERE li.order_id = o.id ORDER BY li.sequence LIMIT 1) AS first_sku
            FROM orders o WHERE o.run_id = ? ORDER BY o.sequence
            """,
            (run_id,),
        )
        for order in orders:
            order["payload"] = _readable_json(order.pop("payload_json", "{}"), {})

        batches = self.db.fetch_all(
            "SELECT * FROM batches WHERE run_id = ? ORDER BY batch_number",
            (run_id,),
        )
        unlocked_batch_number = next(
            (
                int(item["batch_number"])
                for item in batches
                if item["status"] != "confirmed"
            ),
            None,
        )
        for batch in batches:
            mappings = self.db.fetch_all(
                """
                SELECT bo.sheet_index, bo.sheet_name, bo.expected_tab_label,
                       bo.actual_tab_label, bo.note_verified, bo.status,
                       o.id AS order_id, o.order_code, o.customer_code,
                       o.customer_name,
                       (SELECT COUNT(*) FROM line_items li WHERE li.order_id = o.id) AS product_count,
                       (SELECT li.product_code FROM line_items li WHERE li.order_id = o.id ORDER BY li.sequence LIMIT 1) AS first_sku
                FROM batch_orders bo
                JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ? ORDER BY bo.sheet_index
                """,
                (batch["id"],),
            )
            for mapping in mappings:
                mapping["note_verified"] = bool(mapping["note_verified"])
            batch["orders"] = mappings
            batch["order_count"] = len(mappings)
            batch["line_count"] = sum(int(item.get("product_count") or 0) for item in mappings)
            is_unlocked = int(batch["batch_number"]) == unlocked_batch_number
            batch["is_unlocked"] = is_unlocked
            batch["can_import"] = is_unlocked and batch["status"] == "ready"
            batch["can_confirm"] = is_unlocked and batch["status"] == "waiting_user_confirmation"
            batch["can_retry"] = is_unlocked and batch["status"] in {"failed", "automation_attention"}
            batch["can_resume_partial"] = False
            if is_unlocked and batch["status"] == "automation_attention":
                partial_events = self.db.fetch_all(
                    """
                    SELECT payload_json FROM events
                    WHERE batch_id = ?
                      AND event_type IN ('batch.automation_attention', 'batch.external_reconciliation')
                    ORDER BY id DESC LIMIT 20
                    """,
                    (batch["id"],),
                )
                for partial_event in partial_events:
                    envelope = _readable_json(partial_event.get("payload_json", "{}"), {})
                    result = envelope.get("result") if isinstance(envelope, dict) else None
                    container = result.get("submitted") if isinstance(result, dict) else None
                    entries = container.get("submitted") if isinstance(container, dict) else None
                    if isinstance(entries, list) and any(isinstance(item, dict) and item.get("success") is True for item in entries):
                        batch["can_resume_partial"] = True
                        break
            batch["display_status"] = batch["status"] if is_unlocked or batch["status"] == "confirmed" else "blocked"
            batch["is_current"] = is_unlocked

        duplicates = self.db.fetch_all(
            """
            SELECT d.id, d.duplicate_kind, d.decision, d.decided_at,
                   o.order_code, o.payload_sha256,
                   d.prior_run_id, d.prior_order_id,
                   po.payload_sha256 AS prior_payload_sha256,
                   pr.created_at AS prior_run_created_at,
                   pr.status AS prior_run_status
            FROM duplicate_decisions d
            JOIN orders o ON o.id = d.order_id
            LEFT JOIN orders po ON po.id = d.prior_order_id
            LEFT JOIN runs pr ON pr.id = d.prior_run_id
            WHERE d.run_id = ? ORDER BY o.sequence
            """,
            (run_id,),
        )
        events = self.db.fetch_all(
            "SELECT * FROM events WHERE run_id = ? ORDER BY id DESC LIMIT 100",
            (run_id,),
        )
        for event in events:
            event["payload"] = _readable_json(event.pop("payload_json", "{}"), {})
        events.reverse()

        run["orders"] = orders
        run["batches"] = batches
        run["duplicates"] = duplicates
        run["events"] = events
        run["requires_duplicate_confirmation"] = run["status"] == "duplicate_review"
        run["waiting_user_confirmation"] = run["status"] == "waiting_user_confirmation"
        run["safe_mode"] = False
        run["auto_submit_enabled"] = True
        run["automation_mode"] = "verified_auto_submit"
        return run

    def _active_run_row(self) -> dict[str, Any] | None:
        placeholders = ",".join("?" for _ in ACTIVE_RUN_STATUSES)
        # Keep an explicitly selected historical workflow visible and
        # resumable.  Legacy implicit runs still expire when the business day
        # changes, so they cannot take over a fresh dashboard indefinitely.
        while True:
            candidate = self.db.fetch_one(
                f"SELECT * FROM runs WHERE status IN ({placeholders}) ORDER BY id DESC LIMIT 1",
                tuple(sorted(ACTIVE_RUN_STATUSES)),
            )
            if candidate is None:
                return None
            if (
                _date_value(candidate["business_date"], field="business_date") == self._today()
                or bool(candidate["historical_run"])
            ):
                return candidate
            self._expire_stale_run(candidate)

    @staticmethod
    def _assert_run_not_cancelled(run: Mapping[str, Any]) -> None:
        """Reject every resume path after the operator terminally aborts a run."""

        if str(run.get("status") or "") == "cancelled":
            raise InvalidTransitionError(
                "This workflow was cancelled and cannot be resumed.",
                details={"run_id": run.get("id"), "status": "cancelled"},
            )

    def _resolve_run(self, run_id: int | None) -> dict[str, Any]:
        if run_id is None:
            run = self._active_run_row()
            if run is None:
                run = self.db.fetch_one("SELECT * FROM runs ORDER BY id DESC LIMIT 1")
        else:
            run = self.db.fetch_one("SELECT * FROM runs WHERE id = ?", (int(run_id),))
        if run is None:
            raise WorkflowError("No workflow run is available.")
        return run

    def _resolve_batch(
        self,
        *,
        batch_id: int | None = None,
        run_id: int | None = None,
        batch_number: int | None = None,
    ) -> dict[str, Any]:
        if batch_id is not None:
            batch = self.db.fetch_one("SELECT * FROM batches WHERE id = ?", (int(batch_id),))
        else:
            run = self._resolve_run(run_id)
            if batch_number is None:
                batch = self.db.fetch_one(
                    """
                    SELECT * FROM batches
                    WHERE run_id = ? AND status != 'confirmed'
                    ORDER BY batch_number LIMIT 1
                    """,
                    (run["id"],),
                )
            else:
                batch = self.db.fetch_one(
                    "SELECT * FROM batches WHERE run_id = ? AND batch_number = ?",
                    (run["id"], int(batch_number)),
                )
        if batch is None:
            raise WorkflowError("The requested batch was not found.")
        return batch

    def _assert_batch_unlocked(self, batch: Mapping[str, Any]) -> None:
        with self.db.transaction() as connection:
            self._assert_batch_unlocked_in_connection(connection, batch)

    @staticmethod
    def _assert_batch_unlocked_in_connection(
        connection: sqlite3.Connection,
        batch: Mapping[str, Any],
    ) -> None:
        prior = connection.execute(
            """
            SELECT id, batch_number, status FROM batches
            WHERE run_id = ? AND batch_number < ? AND status != 'confirmed'
            ORDER BY batch_number LIMIT 1
            """,
            (batch["run_id"], batch["batch_number"]),
        ).fetchone()
        if prior is not None:
            raise InvalidTransitionError(
                "The previous batch must be confirmed before this batch can start.",
                details={
                    "batch_id": batch["id"],
                    "blocked_by_batch_id": prior["id"],
                    "blocked_by_batch_number": prior["batch_number"],
                    "blocked_by_status": prior["status"],
                },
            )

    def _override_duplicates(self, run_id: int) -> None:
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            pending = int(
                connection.execute(
                    "SELECT COUNT(*) FROM duplicate_decisions WHERE run_id = ? AND decision = 'pending'",
                    (run_id,),
                ).fetchone()[0]
            )
            connection.execute(
                """
                UPDATE duplicate_decisions
                SET decision = 'continue', decided_at = ?
                WHERE run_id = ? AND decision = 'pending'
                """,
                (now, run_id),
            )
            connection.execute(
                "UPDATE orders SET status = 'duplicate_override', updated_at = ? WHERE run_id = ? AND status = 'duplicate_pending'",
                (now, run_id),
            )
            connection.execute(
                """
                UPDATE runs SET status = 'parsed', duplicates_allowed = 1, updated_at = ?
                WHERE id = ? AND status = 'duplicate_review'
                """,
                (now, run_id),
            )
            self._insert_event(
                connection,
                run_id=run_id,
                event_type="duplicates.overridden",
                level="warning",
                message="User explicitly continued same-day duplicate orders.",
                payload={"count": pending},
                created_at=now,
            )

    def _guard_business_date(self, run: Mapping[str, Any]) -> None:
        run_date = _date_value(run["business_date"], field="business_date")
        today = self._today()
        if run_date == today:
            return
        # A report is still constrained to one source business date when it
        # is created, but a run already in progress is explicitly allowed to
        # finish after midnight.  Do not rewrite its business date: the
        # confirmed-order ledger must retain the source date for safe
        # per-day de-duplication.
        return

    def _expire_stale_run(
        self,
        run: Mapping[str, Any],
        *,
        run_date: date | None = None,
        today: date | None = None,
    ) -> None:
        """Expire an unfinished prior-day run without blocking a new run."""

        run_date = run_date or _date_value(run["business_date"], field="business_date")
        today = today or self._today()
        # A run created from an explicitly selected report date is an
        # intentional historical workflow.  Keep it resumable across
        # midnight; its ledger remains keyed to the source business date.
        if run_date == today or bool(run.get("historical_run")):
            return
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            updated = connection.execute(
                """
                UPDATE runs SET status = 'expired', updated_at = ?,
                                error_message = 'Business date changed before workflow completion.'
                WHERE id = ? AND status NOT IN ('completed', 'cancelled', 'expired')
                """,
                (now, run["id"]),
            )
            if updated.rowcount:
                self._insert_event(
                    connection,
                    run_id=int(run["id"]),
                    event_type="run.expired_at_midnight",
                    level="warning",
                    message="Workflow stopped because the Thailand business date changed.",
                    payload={"run_date": run_date.isoformat(), "today": today.isoformat()},
                    created_at=now,
                )

    def _builder_orders_for_batch(self, batch_id: int) -> list[dict[str, Any]]:
        rows = self.db.fetch_all(
            """
            SELECT bo.sheet_index, o.*
            FROM batch_orders bo
            JOIN orders o ON o.id = bo.order_id
            WHERE bo.batch_id = ? ORDER BY bo.sheet_index
            """,
            (batch_id,),
        )
        result: list[dict[str, Any]] = []
        for order in rows:
            products = self.db.fetch_all(
                "SELECT * FROM line_items WHERE order_id = ? ORDER BY sequence",
                (order["id"],),
            )
            result.append(
                {
                    "order_code": order["order_code"],
                    "customer_code": order["customer_code"],
                    "customer_name": order["customer_name"],
                    "products": [
                        {
                            "product_code": item["product_code"],
                            "product_name": item["product_name"],
                            "unit_price": _json_number(item["unit_price"]),
                            "discount_percent": _json_number(item["discount_percent"]),
                            "quantity": _json_number(item["quantity"]),
                        }
                        for item in products
                    ],
                }
            )
        return result

    def _mark_batch_building(self, batch_id: int, run_id: int) -> None:
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                "UPDATE batches SET status = 'building', error_message = NULL, updated_at = ? WHERE id = ?",
                (now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status = 'preparing_batches', error_message = NULL, updated_at = ? WHERE id = ?",
                (now, run_id),
            )
            self._insert_event(
                connection,
                run_id=run_id,
                batch_id=batch_id,
                event_type="batch.build_started",
                created_at=now,
            )

    def _record_build_result(self, run_id: int, batch_id: int, result: BuildResult) -> None:
        now = self._iso_now()
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                """
                UPDATE batches SET status = 'ready', workbook_path = ?,
                                   workbook_sha256 = ?, updated_at = ?, error_message = NULL
                WHERE id = ?
                """,
                (result.output_path, result.sha256, now, batch_id),
            )
            for mapping in result.mappings:
                connection.execute(
                    """
                    UPDATE batch_orders
                    SET sheet_name = ?, status = 'ready', updated_at = ?
                    WHERE batch_id = ? AND sheet_index = ?
                    """,
                    (
                        str(mapping.get("sheet_name") or ""),
                        now,
                        batch_id,
                        int(mapping["sheet_index"]),
                    ),
                )
            connection.execute(
                """
                UPDATE orders SET status = 'batch_ready', updated_at = ?
                WHERE id IN (SELECT order_id FROM batch_orders WHERE batch_id = ?)
                """,
                (now, batch_id),
            )
            self._insert_event(
                connection,
                run_id=run_id,
                batch_id=batch_id,
                event_type="batch.workbook_ready",
                message="KiotViet import workbook built and verified.",
                payload=result.to_dict(),
                created_at=now,
            )

    def _record_build_failure(self, run_id: int, batch_id: int, error: Exception) -> None:
        now = self._iso_now()
        message = str(error)
        with self.db.transaction(immediate=True) as connection:
            connection.execute(
                "UPDATE batches SET status = 'failed', error_message = ?, updated_at = ? WHERE id = ?",
                (message, now, batch_id),
            )
            connection.execute(
                "UPDATE runs SET status = 'failed', error_message = ?, updated_at = ? WHERE id = ?",
                (message, now, run_id),
            )
            self._insert_event(
                connection,
                run_id=run_id,
                batch_id=batch_id,
                event_type="batch.build_failed",
                level="error",
                message=message,
                created_at=now,
            )

    @classmethod
    def _require_verified_batch_mappings(
        cls,
        connection: sqlite3.Connection,
        *,
        batch_id: int,
        expected_count: int,
    ) -> list[dict[str, Any]]:
        """Fail closed unless every persisted sheet has the expected tab/note."""

        rows = [
            dict(row)
            for row in connection.execute(
                """
                SELECT bo.order_id, bo.sheet_index, bo.expected_tab_label,
                       bo.actual_tab_label, bo.note_verified, bo.status,
                       o.order_code
                FROM batch_orders bo
                JOIN orders o ON o.id = bo.order_id
                WHERE bo.batch_id = ?
                ORDER BY bo.sheet_index
                """,
                (batch_id,),
            ).fetchall()
        ]
        if len(rows) != expected_count:
            raise WorkflowError(
                "The persisted batch mappings are incomplete.",
                details={
                    "batch_id": batch_id,
                    "expected": expected_count,
                    "actual": len(rows),
                },
            )
        for row in rows:
            if int(row["note_verified"] or 0) != 1:
                raise InvalidTransitionError(
                    "Every imported KiotViet tab must have a verified order note.",
                    details={
                        "batch_id": batch_id,
                        "order_code": row["order_code"],
                        "sheet_index": int(row["sheet_index"]),
                    },
                )
            cls._require_expected_tab_label(
                batch_id=batch_id,
                sheet_index=int(row["sheet_index"]),
                expected_tab_label=str(row["expected_tab_label"] or ""),
                actual_tab_label=str(row["actual_tab_label"] or ""),
            )
        return rows

    @staticmethod
    def _submission_entries(submission_result: Mapping[str, Any]) -> list[dict[str, Any]]:
        """Extract the precise direct-Edge final-order proof envelope."""

        if submission_result.get("ok") is not True:
            raise InvalidTransitionError(
                "Edge did not report a successful automated batch result."
            )
        if submission_result.get("finalOrderClicked") is not True:
            raise InvalidTransitionError(
                "Edge did not report that the final KiotViet order action completed."
            )
        submitted_container = submission_result.get("submitted")
        if not isinstance(submitted_container, Mapping):
            raise InvalidTransitionError(
                "The Edge result is missing its final-order verification payload."
            )
        if submitted_container.get("ok") is not True:
            raise InvalidTransitionError(
                "KiotViet did not report a successful final-order verification."
            )
        entries = submitted_container.get("submitted")
        if not isinstance(entries, Sequence) or isinstance(entries, (str, bytes)):
            raise InvalidTransitionError(
                "The Edge final-order verification payload is malformed."
            )
        result: list[dict[str, Any]] = []
        for entry in entries:
            if not isinstance(entry, Mapping):
                raise InvalidTransitionError(
                    "The Edge final-order verification contains an invalid order entry."
                )
            result.append(dict(entry))
        return result

    @classmethod
    def _require_automated_submission_result(
        cls,
        connection: sqlite3.Connection,
        *,
        batch_id: int,
        submission_result: Mapping[str, Any],
    ) -> None:
        """Require one successful direct-Edge proof for every exact mapping."""

        expected_rows = cls._require_verified_batch_mappings(
            connection,
            batch_id=batch_id,
            expected_count=int(
                connection.execute(
                    "SELECT sheet_count FROM batches WHERE id = ?", (batch_id,)
                ).fetchone()[0]
            ),
        )
        expected = {
            str(row["order_code"]): {
                "sheet_index": int(row["sheet_index"]),
                "tab_label": str(row["expected_tab_label"] or ""),
            }
            for row in expected_rows
        }
        entries = cls._submission_entries(submission_result)
        if len(entries) != len(expected):
            raise InvalidTransitionError(
                "Edge did not verify every final order in this batch.",
                details={
                    "batch_id": batch_id,
                    "expected": len(expected),
                    "actual": len(entries),
                },
            )
        seen: set[str] = set()
        for entry in entries:
            order_code = str(entry.get("orderCode") or entry.get("order_code") or "").strip()
            tab_label = str(entry.get("tabLabel") or entry.get("tab_label") or "").strip()
            if not order_code or order_code not in expected:
                raise InvalidTransitionError(
                    "Edge reported an unexpected final-order code.",
                    details={"batch_id": batch_id, "order_code": order_code},
                )
            if order_code in seen:
                raise InvalidTransitionError(
                    "Edge reported the same final-order code more than once.",
                    details={"batch_id": batch_id, "order_code": order_code},
                )
            seen.add(order_code)
            mapping = expected[order_code]
            cls._require_expected_tab_label(
                batch_id=batch_id,
                sheet_index=int(mapping["sheet_index"]),
                expected_tab_label=str(mapping["tab_label"]),
                actual_tab_label=tab_label,
            )
            supplied_sheet_index = entry.get("sheetIndex")
            if supplied_sheet_index is not None:
                try:
                    actual_sheet_index = int(supplied_sheet_index)
                except (TypeError, ValueError) as exc:
                    raise InvalidTransitionError(
                        "Edge reported an invalid sheet index for a final order.",
                        details={"batch_id": batch_id, "order_code": order_code},
                    ) from exc
                if actual_sheet_index != int(mapping["sheet_index"]):
                    raise InvalidTransitionError(
                        "Edge reported a final order from an unexpected worksheet.",
                        details={
                            "batch_id": batch_id,
                            "order_code": order_code,
                            "expected_sheet_index": int(mapping["sheet_index"]),
                            "actual_sheet_index": actual_sheet_index,
                        },
                    )
        if seen != set(expected):  # Defensive clarity if the counts ever change.
            raise InvalidTransitionError(
                "Edge final-order verification does not match the persisted batch.",
                details={
                    "batch_id": batch_id,
                    "missing_order_codes": sorted(set(expected) - seen),
                },
            )

    def _assert_prior_automated_completion(
        self,
        batch: Mapping[str, Any],
        run: Mapping[str, Any],
        *,
        connection: sqlite3.Connection | None = None,
    ) -> None:
        """Only an already automated completion is safely idempotent."""

        query = """
            SELECT source_order_id, batch_id, confirmation_method
            FROM confirmed_order_ledger
            WHERE run_id = ? AND batch_id = ?
            ORDER BY order_code
        """
        rows = (
            connection.execute(query, (run["id"], batch["id"])).fetchall()
            if connection is not None
            else self.db.fetch_all(query, (run["id"], batch["id"]))
        )
        if len(rows) != int(batch["sheet_count"]) or any(
            str(row["confirmation_method"]) != AUTOMATION_CONFIRMATION_METHOD
            for row in rows
        ):
            raise InvalidTransitionError(
                "This batch was already confirmed by a different method.",
                details={"batch_id": batch["id"]},
            )

    @staticmethod
    def _require_batch_status(
        batch: Mapping[str, Any] | None,
        allowed: set[str],
        event_type: str,
    ) -> None:
        if batch is None:
            raise WorkflowError(f"{event_type} requires a batch.")
        if batch["status"] not in allowed:
            raise InvalidTransitionError(
                f"Event {event_type} is not valid while batch status is {batch['status']}.",
                details={"batch_id": batch["id"], "allowed": sorted(allowed)},
            )

    @staticmethod
    def _normalized_tab_label(value: str) -> str:
        return " ".join(unicodedata.normalize("NFC", str(value or "")).split()).casefold()

    @classmethod
    def _require_expected_tab_label(
        cls,
        *,
        batch_id: int,
        sheet_index: int,
        expected_tab_label: str,
        actual_tab_label: str,
    ) -> None:
        if actual_tab_label and cls._normalized_tab_label(actual_tab_label) == cls._normalized_tab_label(
            expected_tab_label
        ):
            return
        raise InvalidTransitionError(
            "The verified note came from an unexpected KiotViet tab.",
            details={
                "batch_id": batch_id,
                "sheet_index": sheet_index,
                "expected_tab_label": expected_tab_label,
                "actual_tab_label": actual_tab_label,
            },
        )

    @classmethod
    def _apply_verified_note_payload(
        cls,
        connection: sqlite3.Connection,
        batch_id: int,
        payload: Mapping[str, Any],
        now: str,
    ) -> None:
        verified_value = payload.get("verified")
        if verified_value is None and isinstance(payload.get("notes"), Mapping):
            verified_value = payload["notes"].get("verified")
        if not isinstance(verified_value, Sequence) or isinstance(
            verified_value, (str, bytes)
        ):
            return
        for item_value in verified_value:
            if not isinstance(item_value, Mapping):
                continue
            order_code = str(
                item_value.get("order_code") or item_value.get("orderCode") or ""
            ).strip()
            sheet_index = item_value.get("sheet_index") or item_value.get("sheetIndex")
            tab_label = str(
                item_value.get("tab_label") or item_value.get("tabLabel") or ""
            ).strip()
            if order_code:
                row = connection.execute(
                    """
                    SELECT bo.order_id, bo.sheet_index, bo.expected_tab_label
                    FROM batch_orders bo
                    JOIN orders o ON o.id = bo.order_id
                    WHERE bo.batch_id = ? AND o.order_code = ?
                    """,
                    (batch_id, order_code),
                ).fetchone()
            elif sheet_index is not None:
                row = connection.execute(
                    """
                    SELECT order_id, sheet_index, expected_tab_label
                    FROM batch_orders WHERE batch_id = ? AND sheet_index = ?
                    """,
                    (batch_id, int(sheet_index)),
                ).fetchone()
            else:
                row = None
            if row is None:
                continue
            order_id = int(row["order_id"])
            expected_tab_label = str(row["expected_tab_label"] or "")
            if cls._normalized_tab_label(tab_label) != cls._normalized_tab_label(expected_tab_label):
                # KiotViet can shift numeric cart labels when a stale blank
                # cart is still open. Edge verified SKU/customer content before
                # returning this mapping, so persist only a strict order-tab
                # label for subsequent final-order verification.
                # A stale blank cart can shift the visible numeric labels.
                # Accept only a real KiotViet "Đặt hàng N" label (not an
                # arbitrary tab) and persist that verified label below.
                order_tab_prefix = "\u0111\u1eb7t h\u00e0ng"
                suffix = tab_label[len(order_tab_prefix):].strip()
                is_order_tab = (
                    tab_label.casefold().startswith(order_tab_prefix)
                    and suffix.isdigit()
                    and int(suffix) > 0
                )
                if not is_order_tab:
                    cls._require_expected_tab_label(
                        batch_id=batch_id,
                        sheet_index=int(row["sheet_index"]),
                        expected_tab_label=expected_tab_label,
                        actual_tab_label=tab_label,
                    )
            connection.execute(
                "UPDATE orders SET status = 'note_filled', updated_at = ? WHERE id = ?",
                (now, order_id),
            )
            connection.execute(
                """
                UPDATE batch_orders
                SET expected_tab_label = COALESCE(NULLIF(?, ''), expected_tab_label),
                    note_verified = 1, status = 'note_filled',
                    actual_tab_label = COALESCE(NULLIF(?, ''), actual_tab_label),
                    updated_at = ?
                WHERE batch_id = ? AND order_id = ?
                """,
                (tab_label, tab_label, now, batch_id, order_id),
            )

    @staticmethod
    def _insert_event(
        connection: sqlite3.Connection,
        *,
        run_id: int | None = None,
        batch_id: int | None = None,
        order_id: int | None = None,
        event_type: str,
        level: str = "info",
        message: str = "",
        payload: Mapping[str, Any] | None = None,
        created_at: str,
    ) -> None:
        connection.execute(
            """
            INSERT INTO events(
                run_id, batch_id, order_id, event_type, level,
                message, payload_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                run_id,
                batch_id,
                order_id,
                event_type,
                level,
                message,
                json_dumps(dict(payload or {})),
                created_at,
            ),
        )

    def _now(self) -> datetime:
        value = self._clock() if self._clock else self.config.now()
        if value.tzinfo is None:
            return value.replace(tzinfo=BANGKOK_TZ)
        return value.astimezone(BANGKOK_TZ)

    def _today(self) -> date:
        return self._now().date()

    def _iso_now(self) -> str:
        return self._now().isoformat(timespec="seconds")


__all__ = [
    "ACTIVE_RUN_STATUSES",
    "ActiveRunError",
    "AutomationService",
    "DuplicateReviewRequired",
    "InvalidTransitionError",
    "MidnightGuardError",
    "SourceValidationError",
    "TERMINAL_RUN_STATUSES",
    "WorkflowError",
]
