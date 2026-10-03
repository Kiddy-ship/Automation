"""Local dashboard/API for the Sopoka to KiotViet safe workflow."""

from __future__ import annotations

import argparse
import json
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Callable

from flask import Flask, jsonify, render_template, request, send_file
from waitress import serve

from automation_app.config import AppConfig
from automation_app.db import json_dumps
from automation_app.edge_automation import EdgeAutomation, EdgeAutomationError
from automation_app.scheduler import ScheduledWorkflowRunner
from automation_app.workbook_builder import PurchaseWorkbookBuilder, WorkbookBuildError
from automation_app.service import (
    AutomationService,
    DuplicateReviewRequired,
    InvalidTransitionError,
    MidnightGuardError,
    SourceValidationError,
    WorkflowError,
)


config = AppConfig()
service = AutomationService(config)
edge_automation = EdgeAutomation(config)
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["JSON_AS_ASCII"] = False


SAFE_CHROME_ACTIONS = {
    "PING",
    "CANCEL_ACTIVE",
    "EXPORT_TODAY",
    "EXPORT_DATE",
    "INSPECT_SOPKA",
    "INSPECT_KIOT",
    "PREPARE_KIOT",
    "CLEAN_KIOT_TABS",
    "IMPORT_BATCH",
    "FILL_NOTES",
    "IMPORT_PURCHASE",
}

QUEUED_COMMAND_TIMEOUT_SECONDS = 30.0
DISPATCHED_COMMAND_TIMEOUT_SECONDS = 300.0


class ChromeCommandBroker:
    """A localhost-only, one-at-a-time bridge to the loaded Edge extension.

    The extension polls this broker from the visible OrderBridge tab.  This
    lets the local automation service request the same *safe* extension
    actions exposed in the UI without needing a remote browser session.
    No final order-submit action exists in the allow-list.
    """

    def __init__(
        self,
        *,
        queued_timeout_seconds: float = QUEUED_COMMAND_TIMEOUT_SECONDS,
        dispatched_timeout_seconds: float = DISPATCHED_COMMAND_TIMEOUT_SECONDS,
        clock: Callable[[], float] | None = None,
    ) -> None:
        self._lock = threading.Lock()
        self._commands: dict[str, dict[str, Any]] = {}
        self._queued_timeout_seconds = max(float(queued_timeout_seconds), 0.001)
        self._dispatched_timeout_seconds = max(float(dispatched_timeout_seconds), 0.001)
        self._clock = clock or time.time

    def _expire_locked(self, now: float) -> None:
        for command in self._commands.values():
            status = command["status"]
            if status == "queued":
                started_at = float(command["created_at"])
                timeout = self._queued_timeout_seconds
                phase = "before dispatch"
            elif status == "dispatched":
                started_at = float(command["dispatched_at"] or command["created_at"])
                timeout = self._dispatched_timeout_seconds
                phase = "while awaiting a result"
            else:
                continue
            if now - started_at < timeout:
                continue
            command["status"] = "expired"
            command["expired_at"] = now
            command["response"] = {
                "ok": False,
                "error": f"Edge command timed out {phase}.",
            }

    def submit(self, action: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        action = str(action or "").upper().strip()
        if action not in SAFE_CHROME_ACTIONS:
            raise WorkflowError("Edge action is not allowed in safe mode.")
        with self._lock:
            now = self._clock()
            self._expire_locked(now)
            active = next(
                (
                    command
                    for command in self._commands.values()
                    if command["status"] in {"queued", "dispatched"}
                ),
                None,
            )
            if active is not None:
                raise InvalidTransitionError(
                    "An Edge action is already in progress.",
                    details={"command_id": active["id"], "action": active["action"]},
                )
            command = {
                "id": uuid.uuid4().hex,
                "action": action,
                "payload": dict(payload or {}),
                "status": "queued",
                "created_at": now,
                "dispatched_at": None,
                "completed_at": None,
                "expired_at": None,
                "response": None,
            }
            self._commands[command["id"]] = command
            return dict(command)

    def poll(self) -> dict[str, Any] | None:
        with self._lock:
            now = self._clock()
            self._expire_locked(now)
            command = next(
                (item for item in self._commands.values() if item["status"] == "queued"),
                None,
            )
            if command is None:
                return None
            command["status"] = "dispatched"
            command["dispatched_at"] = now
            return {
                "id": command["id"],
                "action": command["action"],
                "payload": command["payload"],
            }

    def complete(self, command_id: str, response: dict[str, Any]) -> dict[str, Any]:
        with self._lock:
            now = self._clock()
            self._expire_locked(now)
            command = self._commands.get(command_id)
            if command is None:
                raise WorkflowError("Unknown Edge command.")
            if command["status"] == "expired":
                raise InvalidTransitionError(
                    "Edge command expired before its result arrived.",
                    details={"command_id": command_id, "action": command["action"]},
                )
            command["response"] = dict(response)
            command["completed_at"] = now
            command["status"] = "completed" if response.get("ok") is not False and not response.get("error") else "failed"
            return dict(command)

    def get(self, command_id: str) -> dict[str, Any] | None:
        with self._lock:
            self._expire_locked(self._clock())
            command = self._commands.get(command_id)
            return dict(command) if command is not None else None


chrome_broker = ChromeCommandBroker()

# Only one server-side queue may control the visible KiotViet window at a
# time.  This is deliberately separate from EdgeAutomation's WebDriver lock:
# the queue owns the business workflow while EdgeAutomation owns an individual
# browser action.  A second UI click must never upload the next workbook while
# the first request is still awaiting a result.
automation_queue_lock = threading.Lock()
scheduler_tick_lock = threading.Lock()


def _payload() -> dict[str, Any]:
    value = request.get_json(silent=True)
    return value if isinstance(value, dict) else {}


def _error_response(exc: Exception, status: int | None = None):
    if isinstance(exc, WorkflowError):
        body: dict[str, Any] = {"ok": False, "error": exc.__class__.__name__, "message": str(exc)}
        body.update(exc.to_dict())
        if status is None:
            status = 409 if isinstance(exc, (DuplicateReviewRequired, InvalidTransitionError)) else 400
    elif isinstance(exc, EdgeAutomationError):
        body = {"ok": False, "error": exc.__class__.__name__, "message": str(exc)}
        status = status or 400
    else:
        body = {"ok": False, "error": exc.__class__.__name__, "message": str(exc)}
        status = status or 500
    return jsonify(body), status


def _current_automated_batch(run_state: dict[str, Any]) -> dict[str, Any] | None:
    """Return the one unlocked batch that the verified queue may advance.

    The database exposes ``is_current`` based on the first non-confirmed
    batch.  Using that persisted ordering, rather than a client supplied
    batch id, is what prevents a later Excel file from being uploaded before
    the earlier batch has a durable KiotViet result in SQLite.
    """

    batches = run_state.get("batches")
    if not isinstance(batches, list):
        return None
    for batch in batches:
        if not isinstance(batch, dict):
            continue
        if batch.get("is_current") is True:
            return batch
    return None


def _automation_attention_error(run_state: dict[str, Any], batch: dict[str, Any] | None) -> InvalidTransitionError:
    """Produce a fail-closed error after a partial or unverified submission."""

    detail_batch = batch or {}
    message = str(
        detail_batch.get("error_message")
        or run_state.get("error_message")
        or "Lô tự động cần được kiểm tra vì kết quả ĐẶT HÀNG chưa thể xác minh đầy đủ."
    )
    return InvalidTransitionError(
        message,
        details={
            "run_id": run_state.get("id"),
            "batch_id": detail_batch.get("id"),
            "batch_status": detail_batch.get("status"),
            "state": "automation_attention",
        },
    )


def _record_automation_failure(batch_id: int, result: dict[str, Any]) -> None:
    """Quarantine a batch after an external action could be incomplete.

    This helper intentionally ignores a secondary persistence failure: the
    original browser failure is still returned to the UI, and a log entry
    remains available for diagnosis.  It never retries an order click.
    """

    try:
        service.record_automated_submission_failure(batch_id, result)
    except Exception as exc:  # pragma: no cover - defensive audit path
        app.logger.exception(
            "Could not record automated batch failure batch_id=%s: %s",
            batch_id,
            exc,
        )


def _run_automated_queue(run_id: int) -> dict[str, Any]:
    """Upload/submit one persisted batch at a time until the run is done.

    Every payload comes from SQLite, so the browser cannot be driven by stale
    UI data.  The next batch is selected only after
    ``complete_automated_batch`` has atomically recorded all final-order
    proofs and ledger rows for the preceding batch.
    """

    if not automation_queue_lock.acquire(blocking=False):
        raise InvalidTransitionError(
            "Một workflow tự động khác đang chạy trên Edge. Vui lòng chờ tác vụ hiện tại hoàn tất.",
            details={"run_id": run_id},
        )

    completed_batches: list[dict[str, Any]] = []
    active_batch_id: int | None = None
    try:
        while True:
            snapshot = service.state(run_id)
            run = snapshot.get("run")
            if not isinstance(run, dict) or int(run.get("id") or 0) != int(run_id):
                raise WorkflowError(f"Không tìm thấy workflow {run_id}.")

            run_status = str(run.get("status") or "")
            if run_status == "completed":
                return {
                    "ok": True,
                    "run_id": run_id,
                    "completed_batches": completed_batches,
                    "state": snapshot,
                }
            if run_status == "cancelled":
                raise InvalidTransitionError(
                    "Workflow was cancelled and cannot be resumed.",
                    details={"run_id": run_id, "status": run_status},
                )

            batch = _current_automated_batch(run)
            partial_recovery = None
            if run_status == "automation_attention" or (
                isinstance(batch, dict) and str(batch.get("status") or "") == "automation_attention"
            ):
                if not isinstance(batch, dict):
                    raise _automation_attention_error(run, None)
                partial_recovery = service.prepare_partial_automation_recovery(int(batch["id"]))
                if partial_recovery is None:
                    raise _automation_attention_error(run, batch)
            if not isinstance(batch, dict):
                raise InvalidTransitionError(
                    "Không tìm thấy lô đang được mở khóa để tự động xử lý.",
                    details={"run_id": run_id, "run_status": run_status},
                )

            try:
                active_batch_id = int(batch["id"])
            except (KeyError, TypeError, ValueError) as exc:
                raise WorkflowError("Lô workflow thiếu mã nhận diện hợp lệ.") from exc
            batch_status = str(batch.get("status") or "")
            batch_number = batch.get("batch_number")

            if partial_recovery is not None:
                # A browser restart can discard the imported carts.  Inspect
                # before touching anything: only one *blank* order tab means
                # the old carts are gone.  In that exact state create a small
                # replacement workbook for the unconfirmed codes only.  The
                # five externally proven orders never appear in that file.
                live = edge_automation.execute("INSPECT_KIOT_CURRENT", {})
                if not isinstance(live, dict) or live.get("ok") is not True:
                    resumed = live if isinstance(live, dict) else {
                        "ok": False,
                        "error": "Could not inspect the current KiotViet sales canvas.",
                    }
                else:
                    order_tabs = live.get("orderTabs") if isinstance(live.get("orderTabs"), list) else []
                    invoice_tabs = live.get("invoiceTabs") if isinstance(live.get("invoiceTabs"), list) else []
                    blank_canvas = (
                        len(order_tabs) == 1
                        and not invoice_tabs
                        and not str(live.get("noteValue") or "").strip()
                        and int(live.get("cartItemCount") or 0) == 0
                    )
                    if blank_canvas:
                        recovery_import = service.build_partial_recovery_import(
                            active_batch_id,
                            [
                                str(order.get("orderCode") or "").strip()
                                for order in partial_recovery["remainingOrders"]
                            ],
                        )
                        resumed = edge_automation.execute(
                            "IMPORT_BATCH",
                            {
                                "filePath": recovery_import["workbookPath"],
                                "orders": recovery_import["orders"],
                                "batchId": active_batch_id,
                                "batchNumber": batch.get("batch_number"),
                                "autoSubmit": True,
                            },
                        )
                        if isinstance(resumed, dict) and resumed.get("ok") is True:
                            prior = partial_recovery["priorResult"]
                            prior_notes = prior.get("notes") if isinstance(prior, dict) else {}
                            prior_verified = prior_notes.get("verified", []) if isinstance(prior_notes, dict) else []
                            current_notes = resumed.get("notes") if isinstance(resumed.get("notes"), dict) else {}
                            current_verified = current_notes.get("verified", []) if isinstance(current_notes.get("verified"), list) else []
                            resumed = {
                                **resumed,
                                "notes": {"ok": True, "verified": prior_verified + current_verified},
                            }
                    elif len(order_tabs) >= len(partial_recovery["remainingOrders"]) + 1:
                        # The previous import is still visible. Refill only
                        # the remaining cart notes, then continue submitting.
                        refilled = edge_automation.execute(
                            "FILL_NOTES",
                            {"orders": partial_recovery["remainingOrders"], "batchId": active_batch_id},
                        )
                        if not isinstance(refilled, dict) or refilled.get("ok") is not True:
                            resumed = refilled if isinstance(refilled, dict) else {"ok": False, "error": "Could not refill remaining KiotViet notes."}
                        else:
                            refilled_notes = refilled.get("notes") if isinstance(refilled.get("notes"), dict) else {}
                            refilled_verified = refilled_notes.get("verified") if isinstance(refilled_notes.get("verified"), list) else []
                            tab_by_code = {
                                str(item.get("orderCode") or item.get("order_code") or "").strip(): str(item.get("tabLabel") or item.get("tab_label") or "").strip()
                                for item in refilled_verified if isinstance(item, dict)
                            }
                            submit_remaining = [
                                {**order, "expectedTabLabel": tab_by_code.get(str(order.get("orderCode") or "").strip(), order.get("expectedTabLabel"))}
                                for order in partial_recovery["remainingOrders"]
                            ]
                            resumed = edge_automation.execute(
                                "SUBMIT_BATCH",
                                {
                                    "orders": submit_remaining,
                                    "batchId": active_batch_id,
                                    "batchNumber": batch.get("batch_number"),
                                },
                            )
                            if isinstance(resumed, dict) and resumed.get("ok") is True:
                                prior = partial_recovery["priorResult"]
                                prior_notes = prior.get("notes") if isinstance(prior, dict) else {}
                                prior_verified = prior_notes.get("verified", []) if isinstance(prior_notes, dict) else []
                                resumed = {
                                    **resumed,
                                    "notes": {"ok": True, "verified": prior_verified + refilled_verified},
                                }
                    else:
                        resumed = {
                            "ok": False,
                            "error": "KiotViet has an ambiguous cart state; no tabs were changed.",
                            "orderTabs": order_tabs,
                            "invoiceTabs": invoice_tabs,
                        }
                if isinstance(resumed, dict) and resumed.get("ok") is True:
                    prior = partial_recovery["priorResult"]
                    prior_container = prior.get("submitted") if isinstance(prior, dict) else {}
                    prior_entries = (
                        prior_container.get("submitted", [])
                        if isinstance(prior_container, dict)
                        else []
                    )
                    current_container = resumed.get("submitted")
                    current_entries = (
                        current_container.get("submitted", [])
                        if isinstance(current_container, dict)
                        else []
                    )
                    if not isinstance(prior_entries, list) or not isinstance(current_entries, list):
                        resumed = {"ok": False, "error": "Partial Edge result had an invalid submission envelope."}
                    else:
                        merged_container = dict(current_container or {})
                        merged_container.update(
                            {
                                "ok": True,
                                "allSubmitted": True,
                                "submitted": prior_entries + current_entries,
                            }
                        )
                        result = {
                            **prior,
                            "ok": True,
                            "finalOrderClicked": True,
                            "notes": resumed.get("notes") or prior.get("notes"),
                            "submitted": merged_container,
                        }
                        if isinstance(resumed.get("imported"), dict):
                            result["imported"] = resumed["imported"]
                if not isinstance(resumed, dict) or resumed.get("ok") is not True:
                    failure = resumed if isinstance(resumed, dict) else {"error": "Edge partial recovery failed."}
                    _record_automation_failure(active_batch_id, failure)
                    raise InvalidTransitionError(
                        str(failure.get("message") or failure.get("error") or "Edge partial recovery failed."),
                        details={"run_id": run_id, "batch_id": active_batch_id, "state": "automation_attention"},
                    )
            elif batch_status == "ready":
                payload = service.batch_execution_payload(active_batch_id)
                service.mark_extension_event(
                    "batch.import_started",
                    {"source": "verified_auto_queue", "auto_submit": True},
                    batch_id=active_batch_id,
                    message="Workflow tự động bắt đầu upload và hoàn tất lô trên KiotViet.",
                )
                result = edge_automation.execute(
                    "IMPORT_BATCH",
                    {
                        "filePath": payload["workbookPath"],
                        "orders": payload["orders"],
                        "batchId": payload["batchId"],
                        "batchNumber": payload["batchNumber"],
                        "autoSubmit": True,
                    },
                )
            elif batch_status == "waiting_user_confirmation":
                # This recovery path is for a batch that was imported and had
                # every note verified before the server/UI was upgraded.  It
                # does not upload its workbook again.
                payload = service.batch_submission_payload(active_batch_id)
                result = edge_automation.execute(
                    "SUBMIT_BATCH",
                    {
                        "orders": payload["orders"],
                        "batchId": payload["batchId"],
                        "batchNumber": payload["batchNumber"],
                    },
                )
            else:
                raise InvalidTransitionError(
                    "Lô hiện tại không ở trạng thái an toàn để tự động xử lý.",
                    details={
                        "run_id": run_id,
                        "batch_id": active_batch_id,
                        "batch_number": batch_number,
                        "batch_status": batch_status,
                    },
                )

            if not isinstance(result, dict) or result.get("ok") is not True:
                failure = result if isinstance(result, dict) else {"error": "Edge returned an invalid result."}
                _record_automation_failure(active_batch_id, failure)
                raise InvalidTransitionError(
                    str(failure.get("message") or failure.get("error") or "Edge chưa xác minh hoàn tất toàn bộ lô."),
                    details={
                        "run_id": run_id,
                        "batch_id": active_batch_id,
                        "batch_number": batch_number,
                        "state": "automation_attention",
                    },
                )

            try:
                service.complete_automated_batch(active_batch_id, result)
            except Exception as exc:
                # The orders may already have been accepted externally.  Stop
                # here and mark the batch for review rather than guessing or
                # moving on to another workbook.
                _record_automation_failure(
                    active_batch_id,
                    {
                        "error": str(exc),
                        "message": "Kết quả ĐẶT HÀNG cần kiểm tra trước khi chạy lô tiếp theo.",
                        "edge_result": result,
                    },
                )
                raise

            completed_batches.append(
                {
                    "batch_id": active_batch_id,
                    "batch_number": batch_number,
                    # Partial recovery can submit a replacement workbook and
                    # therefore has no ordinary ``payload`` local.  The
                    # persisted batch count is the complete, verified count.
                    "order_count": int(batch.get("sheet_count") or 0),
                    "message": str(result.get("message") or "Đã hoàn tất lô trên KiotViet."),
                }
            )
            active_batch_id = None
    except EdgeAutomationError as exc:
        if active_batch_id is not None:
            _record_automation_failure(active_batch_id, {"error": str(exc)})
        raise
    except Exception as exc:
        # A low-level WebDriver/urllib3 failure can escape the narrower
        # EdgeAutomationError wrapper (for example when msedgedriver stops
        # responding while a native print preview is open).  Never leave the
        # persisted batch in ``importing``/``batch_in_progress`` in that case:
        # quarantine it with the same explicit recovery contract used by all
        # other partial submissions.  This keeps the UI truthful and prevents
        # a second click from treating an unknown external result as fresh.
        if active_batch_id is not None and not isinstance(exc, InvalidTransitionError):
            try:
                _record_automation_failure(
                    active_batch_id,
                    {
                        "error": str(exc),
                        "source": "unhandled_automation_exception",
                    },
                )
            except Exception:
                app.logger.exception(
                    "Could not quarantine batch after unhandled automation exception batch_id=%s",
                    active_batch_id,
                )
        raise
    finally:
        automation_queue_lock.release()


def _run_scheduled_tick() -> dict[str, Any]:
    """Run one Task Scheduler pass through the normal audited workflow."""

    runner = ScheduledWorkflowRunner(
        config,
        service,
        edge_automation,
        _run_automated_queue,
    )
    return runner.run_once()


@app.get("/")
def index():
    return render_template("index.html", default_purchase_date=config.today().isoformat())


@app.get("/api/state")
def api_state():
    run_id = request.args.get("run_id", type=int)
    try:
        return jsonify(service.state(run_id))
    except Exception as exc:  # pragma: no cover - defensive API boundary
        return _error_response(exc)


@app.get("/api/ledger/orders")
def api_ledger_orders():
    """Read the confirmed-order ledger for the SQLite dashboard tab."""

    try:
        return jsonify(service.ledger_orders(request.args.get("business_date")))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/ledger/orders/delete")
def api_delete_ledger_orders():
    """Forget explicitly selected confirmed codes after a risk acknowledgement."""

    data = _payload()
    try:
        if data.get("confirm") is not True:
            raise WorkflowError("Xóa mã cần xác nhận rõ ràng từ giao diện.")
        if data.get("acknowledge_external_order_risk") is not True:
            raise WorkflowError(
                "Cần xác nhận rằng xóa SQLite không hoàn tác đơn đã đặt trên KiotViet."
            )
        codes = data.get("order_codes") or data.get("orderCodes") or []
        return jsonify(
            service.delete_selected_ledger_orders(
                data.get("business_date") or data.get("businessDate") or service._today().isoformat(),
                codes,
                confirm=True,
                acknowledge_external_order_risk=True,
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.get("/api/browser/status")
def api_browser_status():
    try:
        return jsonify(edge_automation.status())
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/browser/commands")
def api_browser_command():
    """Run one direct Selenium action in the visible Edge automation window."""

    data = _payload()
    try:
        action = str(data.get("action") or "")
        payload = data.get("payload") if isinstance(data.get("payload"), dict) else {}
        return jsonify(edge_automation.execute(action, payload))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/purchase/import")
def api_purchase_import():
    """Build and import one KiotViet purchase workbook.

    The source workbook is read-only.  The generated workbook is retained in
    ``outputs/purchase`` for audit/retry.  The Edge action is run only after
    the workbook has been validated and the single browser queue is acquired.
    """

    data = _payload()
    source_path = str(data.get("source_path") or data.get("sourcePath") or "").strip()
    baseline_path = str(
        data.get("baseline_source_path")
        or data.get("baselineSourcePath")
        or ""
    ).strip()
    try:
        if not source_path:
            raise WorkflowError("Vui lòng dán đường dẫn file Đồng Nai 1 cần nhập hàng.")
        source = Path(source_path).expanduser().resolve()
        if not source.is_file() or source.name.startswith("~$"):
            raise WorkflowError(f"Không tìm thấy file nguồn: {source}")
        # Receipt date is explicit rather than inferred from a filename or a
        # potentially stale KiotViet form.  Validate before the workbook is
        # built so a typo can never proceed toward an upload.
        purchase_date = EdgeAutomation.parse_purchase_date(
            data.get("purchase_date") or data.get("purchaseDate")
        ).isoformat()
        baseline: Path | None = None
        if baseline_path:
            baseline = Path(baseline_path).expanduser().resolve()
            if not baseline.is_file() or baseline.name.startswith("~$"):
                raise WorkflowError(f"Không tìm thấy file đối chiếu: {baseline}")
            if baseline == source:
                raise WorkflowError("File đối chiếu phải khác file nguồn cần nhập.")
        output_dir = config.outputs_dir / "purchase"
        output_dir.mkdir(parents=True, exist_ok=True)
        output = output_dir / f"PurchaseOrder_{config.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if not automation_queue_lock.acquire(blocking=False):
            raise InvalidTransitionError("Một tác vụ Edge khác đang chạy. Vui lòng chờ hoàn tất.")
        try:
            build = PurchaseWorkbookBuilder(config).build(
                source,
                output,
                baseline_path=baseline,
            )
            # An incremental comparison can correctly produce no positive
            # increase.  Retain the audit workbook but do not open KiotViet
            # or create a zero-line purchase receipt.
            if int(build.get("item_count") or 0) == 0:
                return jsonify(
                    {
                        "ok": True,
                        "no_incremental_items": True,
                        "purchase_date": purchase_date,
                        "build": build,
                        "edge": {
                            "ok": True,
                            "skipped": True,
                            "reason": "Không có số lượng giao tăng thêm so với file đối chiếu.",
                        },
                    }
                )
            imported = edge_automation.execute(
                "IMPORT_PURCHASE",
                {
                    "filePath": str(output),
                    # The builder is the sole source of code/name/row mapping
                    # for an optional missing-SKU recovery.  Never let the
                    # browser guess this data from an error banner.
                    "products": build.get("items") or [],
                    "autoCreateMissingProducts": data.get("auto_create_missing_products") is True,
                    "purchaseDate": purchase_date,
                },
            )
        finally:
            automation_queue_lock.release()
        if imported.get("ok") is not True:
            return jsonify({"ok": False, "build": build, "edge": imported}), 400
        return jsonify(
            {
                "ok": True,
                "purchase_date": purchase_date,
                "build": build,
                "edge": imported,
            }
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/runs")
def api_create_run():
    data = _payload()
    source_path = str(data.get("source_path") or data.get("sourcePath") or "").strip()
    business_date = data.get("business_date") or data.get("businessDate") or None
    try:
        if not source_path:
            raise SourceValidationError(
                "Vui lòng chọn đường dẫn file Excel đã tải về (.xls, .xlsx, .xlsm, .xltx hoặc .xltm)."
            )
        return jsonify(service.scan_and_create_incremental_run(
            source_path,
            trigger="manual",
            allow_confirmed_duplicates=data.get("allow_duplicates") is True,
            business_date=business_date,
        ))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/scans/<int:scan_id>/duplicate-decision")
def api_scan_duplicate_decision(scan_id: int):
    data = _payload()
    try:
        return jsonify(
            service.decide_scan_duplicates(
                scan_id,
                continue_scan=data.get("continue") is True,
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/runs/<int:run_id>/duplicate-decision")
def api_duplicate_decision(run_id: int):
    data = _payload()
    try:
        continue_run = bool(data.get("continue"))
        current = service.state(run_id).get("run")
        if not current:
            raise WorkflowError(f"Run {run_id} không tồn tại.")
        if not continue_run:
            now = config.now().isoformat(timespec="seconds")
            with service.db.transaction(immediate=True) as connection:
                connection.execute(
                    "UPDATE runs SET status='cancelled', updated_at=?, error_message=? WHERE id=? AND status='duplicate_review'",
                    (now, now, run_id),
                )
                connection.execute(
                    "UPDATE duplicate_decisions SET decision='cancel', decided_at=? WHERE run_id=? AND decision='pending'",
                    (now, run_id),
                )
                service._insert_event(
                    connection,
                    run_id=run_id,
                    event_type="duplicates.cancelled",
                    level="warning",
                    message="Người dùng dừng vì dữ liệu trùng.",
                    created_at=now,
                )
            return jsonify(service.state(run_id))
        state = service.create_run_from_source(current["source_path"], allow_duplicates=True)
        return jsonify(service.prepare_batches(run_id))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/runs/<int:run_id>/auto-run")
def api_run_automated_queue(run_id: int):
    """Run all unlocked Excel batches sequentially through visible Edge.

    This endpoint intentionally accepts no workbook path, tab name, or order
    list from the browser.  Those values are reconstructed from the already
    audited SQLite workflow, which makes a page refresh or a stale UI unable
    to reorder batches or submit an unintended order.
    """

    try:
        return jsonify(_run_automated_queue(run_id))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/scheduler/tick")
def api_scheduler_tick():
    """Local-only endpoint invoked by the Windows scheduled-task worker."""

    try:
        if request.remote_addr not in {"127.0.0.1", "::1"}:
            raise WorkflowError("Scheduled workflow is available only from this computer.")
        if not scheduler_tick_lock.acquire(blocking=False):
            raise InvalidTransitionError("A scheduled workflow tick is already running.")
        try:
            if automation_queue_lock.locked():
                return jsonify(
                    {
                        "ok": False,
                        "outcome": "blocked",
                        "message": "A manual or automated workflow already owns the Edge queue.",
                    }
                )
            browser = edge_automation.status()
            if bool(browser.get("progress", {}).get("active")):
                return jsonify(
                    {
                        "ok": False,
                        "outcome": "blocked",
                        "message": "Another Edge action is already in progress.",
                    }
                )
            return jsonify(_run_scheduled_tick())
        finally:
            scheduler_tick_lock.release()
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/events")
def api_event():
    data = _payload()
    try:
        event_type = str(data.get("event_type") or data.get("type") or "ui.event")
        payload = data.get("payload") if isinstance(data.get("payload"), dict) else data
        return jsonify(
            service.mark_extension_event(
                event_type,
                payload,
                run_id=data.get("run_id"),
                batch_id=data.get("batch_id"),
                order_code=data.get("order_code"),
                level=str(data.get("level") or "info"),
                message=str(data.get("message") or ""),
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.get("/api/bridge/poll")
def api_bridge_poll():
    """Called only by the OrderBridge dashboard content script."""
    return jsonify({"ok": True, "command": chrome_broker.poll()})


@app.post("/api/bridge/results/<command_id>")
def api_bridge_result(command_id: str):
    try:
        data = _payload()
        response = data.get("response") if isinstance(data.get("response"), dict) else data
        return jsonify({"ok": True, "command": chrome_broker.complete(command_id, response)})
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/bridge/commands")
def api_bridge_command():
    """Local controller endpoint; only safe, visible Edge actions are queued."""
    try:
        data = _payload()
        action = str(data.get("action") or "")
        payload = data.get("payload") if isinstance(data.get("payload"), dict) else {}
        return jsonify({"ok": True, "command": chrome_broker.submit(action, payload)})
    except Exception as exc:
        return _error_response(exc)


@app.get("/api/bridge/commands/<command_id>")
def api_bridge_command_status(command_id: str):
    command = chrome_broker.get(command_id)
    if command is None:
        return jsonify({"ok": False, "message": "Unknown Edge command."}), 404
    return jsonify({"ok": True, "command": command})


@app.post("/api/batches/<int:batch_id>/imported")
def api_batch_imported(batch_id: int):
    data = _payload()
    try:
        result = data.get("extension_result") if isinstance(data.get("extension_result"), dict) else data
        state = service.mark_extension_event(
            "batch.imported",
            {"extension_result": result},
            batch_id=batch_id,
            message="KiotViet đã nhận file import; chưa bấm ĐẶT HÀNG.",
        )
        notes = result.get("notes") if isinstance(result, dict) else None
        if isinstance(notes, dict) and notes.get("ok") and isinstance(notes.get("verified"), list):
            state = service.mark_extension_event(
                "batch.notes_filled",
                notes,
                batch_id=batch_id,
                message="Đã điền và xác minh ghi chú cho từng tab.",
            )
        return jsonify(state)
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/import-started")
def api_batch_import_started(batch_id: int):
    """Claim the sequential batch before Edge performs any upload action."""
    try:
        return jsonify(
            service.mark_extension_event(
                "batch.import_started",
                {"source": "dashboard"},
                batch_id=batch_id,
                message="Người dùng bắt đầu import lô đã được mở khóa.",
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/failed")
def api_batch_failed(batch_id: int):
    data = _payload()
    try:
        return jsonify(
            service.mark_extension_event(
                "batch.error",
                {"error": str(data.get("error") or data.get("message") or "Extension error")},
                batch_id=batch_id,
                level="error",
                message=str(data.get("message") or data.get("error") or "Extension error"),
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/reconcile-import")
def api_batch_reconcile_import(batch_id: int):
    """Reconcile a verified browser import which timed out before its result."""

    data = _payload()
    try:
        if data.get("user_confirmed_import_present") is not True:
            raise WorkflowError("Import reconciliation requires explicit confirmation.")
        notes = data.get("notes") if isinstance(data.get("notes"), dict) else data
        return jsonify(service.reconcile_completed_import(batch_id, notes))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/reconcile-external-submissions")
def api_batch_reconcile_external_submissions(batch_id: int):
    """Store explicitly reviewed KiotViet history before resuming a batch."""

    data = _payload()
    try:
        if data.get("user_confirmed") is not True:
            raise WorkflowError("External KiotViet reconciliation requires explicit confirmation.")
        result = data.get("submission_result")
        evidence = data.get("evidence")
        if not isinstance(result, dict) or not isinstance(evidence, list):
            raise WorkflowError("External reconciliation requires result and evidence arrays.")
        return jsonify(
            service.reconcile_external_partial_submission(
                batch_id,
                result,
                evidence=evidence,
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/reconcile-manual-source-cancellation")
def api_batch_reconcile_manual_source_cancellation(batch_id: int):
    """Close a stopped batch only after audited manual/source reconciliation."""

    data = _payload()
    try:
        if data.get("user_confirmed") is not True:
            raise WorkflowError(
                "Manual source-cancellation reconciliation requires explicit confirmation."
            )
        source_path = data.get("source_path") or data.get("sourcePath")
        cancelled_order_code = (
            data.get("cancelled_order_code") or data.get("cancelledOrderCode")
        )
        evidence = data.get("evidence")
        if not isinstance(source_path, str) or not source_path.strip():
            raise WorkflowError("A freshly exported Sopoka report path is required.")
        if not isinstance(cancelled_order_code, str) or not cancelled_order_code.strip():
            raise WorkflowError("The cancelled Sopoka order code is required.")
        if not isinstance(evidence, list):
            raise WorkflowError("KiotViet evidence must be an array.")
        return jsonify(
            service.reconcile_manual_batch_after_source_cancellation(
                batch_id,
                source_path,
                cancelled_order_code=cancelled_order_code,
                evidence=evidence,
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/retry")
def api_batch_retry(batch_id: int):
    data = _payload()
    try:
        if data.get("user_confirmed_cleanup") is not True:
            raise WorkflowError(
                "Retry requires explicit JSON confirmation that the failed tabs were cleaned up."
            )
        return jsonify(service.retry_batch(batch_id))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/runs/<int:run_id>/abort")
def api_abort_run(run_id: int):
    """Terminally stop a cleaned-up workflow while retaining its audit data."""

    data = _payload()
    try:
        if data.get("confirm") is not True:
            raise WorkflowError("Aborting a workflow requires explicit confirmation.")
        if data.get("kiot_tabs_cleared") is not True:
            raise WorkflowError("Confirm that unfinished KiotViet tabs were cleared first.")
        if data.get("acknowledge_external_order_risk") is not True:
            raise WorkflowError(
                "Acknowledge that cancelling SQLite cannot undo an external KiotViet order."
            )
        if not automation_queue_lock.acquire(blocking=False):
            raise InvalidTransitionError(
                "An Edge workflow is still running; it must finish or stop before this run can be aborted.",
                details={"run_id": run_id},
            )
        try:
            return jsonify(
                service.abort_run(
                    run_id,
                    confirm=True,
                    kiot_tabs_cleared=True,
                    acknowledge_external_order_risk=True,
                )
            )
        finally:
            automation_queue_lock.release()
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/runs/<int:run_id>/delete-selected")
def api_delete_selected_run_orders(run_id: int):
    """Delete only a fully selected failed run after explicit UI confirmation."""

    data = _payload()
    try:
        if data.get("confirm") is not True:
            raise WorkflowError("Xóa mã cần xác nhận rõ ràng từ giao diện.")
        codes = data.get("order_codes") or data.get("orderCodes") or []
        return jsonify(
            service.delete_selected_run_orders(
                run_id,
                codes,
                confirm=data.get("confirm") is True,
            )
        )
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/confirm")
def api_batch_confirm(batch_id: int):
    data = _payload()
    try:
        if data.get("user_confirmed") is not True:
            raise WorkflowError(
                "Batch confirmation requires explicit JSON user confirmation."
            )
        return jsonify(service.confirm_batch(batch_id=batch_id))
    except Exception as exc:
        return _error_response(exc)


@app.post("/api/batches/<int:batch_id>/automation-complete")
def api_batch_automation_complete(batch_id: int):
    """Commit a direct-Edge final-order result after strict DB verification."""

    data = _payload()
    try:
        result = data.get("submission_result")
        if not isinstance(result, dict):
            raise WorkflowError(
                "Automated batch completion requires a submission_result object from Edge."
            )
        return jsonify(service.complete_automated_batch(batch_id, result))
    except Exception as exc:
        return _error_response(exc)


@app.get("/api/files/<int:batch_id>")
def api_download_batch(batch_id: int):
    batch = service.db.fetch_one("SELECT workbook_path FROM batches WHERE id=?", (batch_id,))
    path = Path(str(batch.get("workbook_path") or "")) if batch else None
    if not path or not path.is_file():
        return jsonify({"ok": False, "message": "Không tìm thấy file workbook."}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


def main() -> None:
    parser = argparse.ArgumentParser(description="OrderBridge DN2 local dashboard")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    serve(app, host=args.host, port=args.port, threads=6)


if __name__ == "__main__":
    main()
