"""One safe, externally-triggered scheduled OrderBridge tick.

Windows Task Scheduler decides *when* to call this module.  This module owns
only the business workflow: it never retries an uncertain run, clears a cart,
or changes the ledger outside the normal audited service methods.
"""

from __future__ import annotations

from typing import Any, Callable

from .config import AppConfig
from .service import ACTIVE_RUN_STATUSES, WorkflowError


class ScheduledWorkflowRunner:
    """Run at most one fully verified Sopoka-to-KiotViet pass."""

    def __init__(
        self,
        config: AppConfig,
        service: Any,
        edge_automation: Any,
        auto_run: Callable[[int], dict[str, Any]],
    ) -> None:
        self.config = config
        self.service = service
        self.edge_automation = edge_automation
        self.auto_run = auto_run

    @staticmethod
    def _edge_ok(result: Any, label: str) -> dict[str, Any]:
        if not isinstance(result, dict) or result.get("ok") is not True:
            message = result.get("error") if isinstance(result, dict) else None
            raise WorkflowError(str(message or f"{label} was not verified."))
        return result

    def _close_after_success(self) -> dict[str, Any]:
        """Best-effort cleanup that never changes a completed run's result."""

        try:
            result = self.edge_automation.execute(
                "CLOSE_SCHEDULED_TABS", {"afterSuccess": True}
            )
            return result if isinstance(result, dict) else {"ok": False, "error": "Invalid cleanup result."}
        except Exception as exc:  # pragma: no cover - defensive cleanup path
            return {"ok": False, "error": str(exc)}

    def run_once(self) -> dict[str, Any]:
        """Export, scan, and run only new orders for Bangkok's current date.

        A non-success status is intentionally returned as a completed task
        result.  It is evidence for the operator, not a reason to send a
        second browser request or to erase an uncertain cart.
        """

        business_date = self.config.today().isoformat()
        existing_state = self.service.state()
        existing_run = existing_state.get("run") if isinstance(existing_state, dict) else None
        existing_status = str(existing_run.get("status") or "") if isinstance(existing_run, dict) else ""
        if existing_status in ACTIVE_RUN_STATUSES or existing_status in {"failed", "duplicate_review"}:
            return {
                "ok": False,
                "outcome": "blocked",
                "business_date": business_date,
                "run_id": existing_run.get("id") if isinstance(existing_run, dict) else None,
                "message": "A previous workflow still requires operator review; scheduled run did not touch Edge.",
            }
        preflight = self._edge_ok(
            self.edge_automation.execute(
                "SCHEDULE_PREFLIGHT",
                {
                    "businessDate": business_date,
                    "maxClockSkewSeconds": self.config.scheduler_max_clock_skew_seconds,
                },
            ),
            "Scheduled browser preflight",
        )
        exported = self._edge_ok(
            self.edge_automation.execute(
                "EXPORT_DATE", {"business_date": business_date}
            ),
            "Sopoka export",
        )
        if str(exported.get("business_date") or "") != business_date:
            raise WorkflowError("Sopoka export did not prove the scheduled business date.")
        source_path = str(exported.get("source_path") or exported.get("sourcePath") or "").strip()
        if not source_path:
            raise WorkflowError("Sopoka export did not return a source report path.")

        state = self.service.scan_and_create_incremental_run(
            source_path,
            trigger="scheduled",
            allow_confirmed_duplicates=True,
            business_date=business_date,
        )
        scan = state.get("scan") if isinstance(state, dict) else None
        scan_status = str(scan.get("status") or "") if isinstance(scan, dict) else ""
        base = {
            "ok": True,
            "business_date": business_date,
            "preflight": preflight,
            "export": {
                "source_path": source_path,
                "order_count": exported.get("order_count"),
            },
            "scan": scan,
        }

        if scan_status == "no_new":
            return {**base, "outcome": "no_new", "cleanup": self._close_after_success()}
        if scan_status != "run_created":
            return {
                **base,
                "ok": False,
                "outcome": "blocked",
                "message": str(scan.get("error_message") or f"Scheduled scan stopped: {scan_status or 'unknown'}"),
            }

        run = state.get("run") if isinstance(state, dict) else None
        run_id = int(run.get("id") or 0) if isinstance(run, dict) else 0
        if run_id <= 0:
            raise WorkflowError("Scheduled scan created no runnable workflow.")
        queued = self.auto_run(run_id)
        final_state = self.service.state(run_id)
        final_run = final_state.get("run") if isinstance(final_state, dict) else None
        batch_statuses = [
            str(batch.get("status") or "")
            for batch in (final_run.get("batches") or [])
            if isinstance(batch, dict)
        ] if isinstance(final_run, dict) else []
        if not isinstance(final_run, dict) or final_run.get("status") != "completed" or any(
            status != "confirmed" for status in batch_statuses
        ):
            return {
                **base,
                "ok": False,
                "outcome": "blocked",
                "run_id": run_id,
                "message": "Scheduled workflow did not reach a fully confirmed completion.",
                "state": final_state,
                "queue": queued,
            }
        return {
            **base,
            "outcome": "completed",
            "run_id": run_id,
            "state": final_state,
            "queue": queued,
            "cleanup": self._close_after_success(),
        }


__all__ = ["ScheduledWorkflowRunner"]
