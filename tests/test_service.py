from __future__ import annotations

import hashlib
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from automation_app.config import AppConfig, BANGKOK_TZ
from automation_app.service import (
    AutomationService,
    InvalidTransitionError,
    SourceValidationError,
)
from automation_app.workbook_builder import BuildResult


def make_orders(count: int) -> list[dict]:
    return [
        {
            "sequence": index,
            "order_code": f"DH-{index:03d}",
            "related_order_date": "2026-08-06",
            "customer_code": f"KH{index:06d}",
            "customer_name": f"Khách {index}",
            "products": [
                {
                    "product_code": f"SKU-{index:03d}",
                    "product_name": f"Sản phẩm {index}",
                    "quantity": 1,
                    "unit_price": 10_000,
                    "gross_amount": 10_000,
                    "discount_percent": 0,
                }
            ],
        }
        for index in range(1, count + 1)
    ]


def verified_notes(batch: dict) -> list[dict]:
    return [
        {
            "order_code": item["order_code"],
            "tab_label": item["expected_tab_label"],
        }
        for item in batch["orders"]
    ]


def automated_submission(batch: dict, *, include_notes: bool = False) -> dict:
    """Shape returned by EdgeAutomation.submit_batch/import_batch(autoSubmit)."""

    result = {
        "ok": True,
        "finalOrderClicked": True,
        "submitted": {
            "ok": True,
            "submitted": [
                {
                    "orderCode": item["order_code"],
                    "tabLabel": item["expected_tab_label"],
                    "confirmationClicked": False,
                }
                for item in reversed(batch["orders"])
            ],
        },
    }
    if include_notes:
        result["imported"] = {"ok": True, "tabLabels": ["Đặt hàng 1"]}
        result["notes"] = {"ok": True, "verified": verified_notes(batch)}
    return result


class FakeWorkbookBuilder:
    def build(self, orders, output_path):
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        payload = "\n".join(order["order_code"] for order in orders).encode("utf-8")
        output.write_bytes(payload)
        return BuildResult(
            output_path=str(output),
            sha256=hashlib.sha256(payload).hexdigest(),
            file_size=len(payload),
            order_count=len(orders),
            visible_sheet_count=len(orders),
            mappings=[
                {
                    "sheet_index": index,
                    "sheet_name": "ImportProductTemplate" if index == 1 else f"ImportProductTemplate ({index})",
                    "order_code": order["order_code"],
                    "product_count": len(order["products"]),
                }
                for index, order in enumerate(orders, 1)
            ],
        )


class AutomationServiceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="orderbridge-tests-")
        self.root = Path(self.temp.name)
        self.now = datetime(2026, 8, 6, 10, 0, tzinfo=BANGKOK_TZ)

    def tearDown(self):
        self.temp.cleanup()

    def service(self, orders: list[dict]) -> AutomationService:
        def parser(_source, *, expected_date=None, max_orders_per_batch=19):
            business_date = expected_date.isoformat()
            normalized = []
            for order in orders:
                item = dict(order)
                item["related_order_date"] = business_date
                normalized.append(item)
            return {
                "source_file": str(_source),
                "sheet_name": "Báo cáo",
                "header_row": 4,
                "first_data_row": 5,
                "last_data_row": 4 + len(normalized),
                "report_start_date": business_date,
                "report_end_date": business_date,
                "orders": normalized,
                "qa": {},
            }

        config = AppConfig(
            database_path=self.root / "automation.sqlite3",
            outputs_dir=self.root / "outputs",
            temp_dir=self.root / "tmp",
        )
        return AutomationService(
            config,
            parser=parser,
            workbook_builder=FakeWorkbookBuilder(),
            clock=lambda: self.now,
        )

    def source(self, name: str, content: bytes = b"xls") -> Path:
        path = self.root / name
        path.write_bytes(content)
        return path

    def test_second_batch_stays_locked_until_manual_confirmation(self):
        service = self.service(make_orders(20))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        first, second = state["run"]["batches"]

        self.assertEqual((first["order_count"], second["order_count"]), (19, 1))
        self.assertTrue(first["can_import"])
        self.assertEqual(second["display_status"], "blocked")
        self.assertFalse(second["can_import"])
        self.assertEqual(first["orders"][0]["expected_tab_label"], "Đặt hàng 2")
        self.assertEqual(first["orders"][0]["first_sku"], "SKU-001")

        with self.assertRaises(InvalidTransitionError):
            service.mark_extension_event("batch.import_started", batch_id=second["id"])

        service.mark_extension_event("batch.import_started", batch_id=first["id"])
        service.mark_extension_event("batch.imported", batch_id=first["id"])
        state = service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(first)},
            batch_id=first["id"],
        )
        self.assertTrue(state["run"]["batches"][0]["can_confirm"])

        state = service.confirm_batch(first["id"])
        self.assertTrue(state["run"]["batches"][1]["can_import"])

    def test_execution_payload_includes_full_cart_fingerprint(self):
        orders = make_orders(1)
        orders[0]["customer_code"] = "KH0001"
        orders[0]["payment_amount"] = 8_800
        orders[0]["products"][0].update(
            {"quantity": 2, "gross_amount": 10_000, "discount_percent": 12}
        )
        service = self.service(orders)
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])

        execution = service.batch_execution_payload(state["run"]["batches"][0]["id"])
        expectation = execution["orders"][0]

        self.assertEqual(expectation["customerCode"], "KH0001")
        self.assertEqual(expectation["paymentAmount"], 8_800)
        self.assertEqual(len(expectation["productLines"]), 1)
        self.assertEqual(expectation["productLines"][0]["productCode"], "SKU-001")
        self.assertEqual(expectation["productLines"][0]["quantity"], 2)
        self.assertEqual(expectation["productLines"][0]["netAmount"], 8_800)

    def test_automated_completion_writes_distinct_ledger_and_unlocks_next_batch(self):
        service = self.service(make_orders(20))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        first, second = state["run"]["batches"]

        execution = service.batch_execution_payload(first["id"])
        self.assertEqual(execution["batchId"], first["id"])
        self.assertTrue(Path(execution["workbookPath"]).is_file())
        self.assertEqual(
            [item["orderCode"] for item in execution["orders"]],
            [item["order_code"] for item in first["orders"]],
        )
        self.assertEqual(execution["orders"][0]["expectedTabLabel"], first["orders"][0]["expected_tab_label"])

        service.mark_extension_event("batch.import_started", batch_id=first["id"])
        service.mark_extension_event("batch.imported", batch_id=first["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(first)},
            batch_id=first["id"],
        )
        resume = service.batch_submission_payload(first["id"])
        self.assertEqual(
            [item["orderCode"] for item in resume["orders"]],
            [item["order_code"] for item in first["orders"]],
        )
        state = service.complete_automated_batch(first["id"], automated_submission(first))

        completed_first, unlocked_second = state["run"]["batches"]
        self.assertEqual(completed_first["status"], "confirmed")
        self.assertIsNotNone(completed_first["automation_confirmed_at"])
        self.assertTrue(unlocked_second["can_import"])
        self.assertEqual(state["run"]["status"], "batches_ready")
        self.assertEqual(
            service.db.scalar(
                "SELECT COUNT(*) FROM confirmed_order_ledger "
                "WHERE confirmation_method = 'automation_verified_submission'"
            ),
            19,
        )
        self.assertEqual(
            service.db.scalar(
                "SELECT COUNT(*) FROM orders WHERE status = 'automation_confirmed'"
            ),
            19,
        )

        # A delayed duplicate HTTP callback is idempotent only for the same
        # automation-confirmed batch and must not duplicate the ledger.
        retry = service.complete_automated_batch(first["id"], automated_submission(first))
        self.assertEqual(retry["run"]["status"], "batches_ready")
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM confirmed_order_ledger"), 19)

    def test_automated_completion_can_atomically_record_auto_import_notes_and_finish_run(self):
        service = self.service(make_orders(1))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])

        state = service.complete_automated_batch(
            batch["id"], automated_submission(batch, include_notes=True)
        )

        self.assertEqual(state["run"]["status"], "completed")
        completed = state["run"]["batches"][0]
        self.assertEqual(completed["status"], "confirmed")
        self.assertTrue(completed["orders"][0]["note_verified"])
        self.assertEqual(
            service.db.scalar("SELECT confirmation_method FROM confirmed_order_ledger"),
            "automation_verified_submission",
        )

    def test_automated_completion_rejects_a_partial_result_without_ledger_writes(self):
        service = self.service(make_orders(2))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(batch)},
            batch_id=batch["id"],
        )
        partial = automated_submission(batch)
        partial["submitted"]["submitted"] = partial["submitted"]["submitted"][:1]

        with self.assertRaises(InvalidTransitionError):
            service.complete_automated_batch(batch["id"], partial)

        current = service.state(state["run"]["id"])["run"]["batches"][0]
        self.assertEqual(current["status"], "waiting_user_confirmation")
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM confirmed_order_ledger"), 0)

    def test_partial_automatic_submission_is_quarantined_without_a_blind_retry(self):
        service = self.service(make_orders(20))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        first, second = state["run"]["batches"]
        service.mark_extension_event("batch.import_started", batch_id=first["id"])
        service.mark_extension_event("batch.imported", batch_id=first["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(first)},
            batch_id=first["id"],
        )

        state = service.record_automated_submission_failure(
            first["id"],
            {"error": "KiotViet confirmed only 18 of 19 final tabs."},
        )
        quarantined, blocked = state["run"]["batches"]
        self.assertEqual(state["run"]["status"], "automation_attention")
        self.assertEqual(quarantined["status"], "automation_attention")
        self.assertTrue(quarantined["can_retry"])
        self.assertEqual(blocked["display_status"], "blocked")
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM confirmed_order_ledger"), 0)
        retried = service.retry_batch(first["id"])
        self.assertTrue(retried["run"]["batches"][0]["can_import"])

    def test_abort_retains_audit_and_ledger_but_releases_unconfirmed_codes(self):
        service = self.service(make_orders(20))
        source = self.source("orders.xls")
        state = service.create_run_from_source(source)
        state = service.prepare_batches(state["run"]["id"])
        first, second = state["run"]["batches"]
        run_id = state["run"]["id"]

        # Preserve a completed first batch to prove that aborting the stopped
        # second batch never deletes existing audit rows or ledger evidence.
        service.mark_extension_event("batch.import_started", batch_id=first["id"])
        service.mark_extension_event("batch.imported", batch_id=first["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(first)},
            batch_id=first["id"],
        )
        service.confirm_batch(first["id"])
        service.mark_extension_event("batch.import_started", batch_id=second["id"])
        service.record_automated_submission_failure(second["id"], {"error": "tab closed"})

        aborted = service.abort_run(
            run_id,
            confirm=True,
            kiot_tabs_cleared=True,
            acknowledge_external_order_risk=True,
        )
        run = aborted["run"]
        self.assertEqual(run["status"], "cancelled")
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM orders WHERE run_id = ?", (run_id,)), 20)
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM batches WHERE run_id = ?", (run_id,)), 2)
        self.assertEqual(service.ledger_orders("2026-08-06")["count"], 19)
        self.assertEqual([batch["status"] for batch in run["batches"]], ["confirmed", "cancelled"])
        self.assertEqual(
            service.db.scalar(
                "SELECT COUNT(*) FROM events WHERE run_id = ? AND event_type = 'run.aborted'",
                (run_id,),
            ),
            1,
        )

        # A cancelled run can never be retried, and its remaining code is
        # deliberately available to a fresh report scan instead of unresolved.
        with self.assertRaises(InvalidTransitionError):
            service.retry_batch(second["id"])
        with self.assertRaises(InvalidTransitionError):
            service.mark_extension_event("batch.import_started", batch_id=second["id"])

        follow_up = service.scan_and_create_incremental_run(
            source,
            trigger="manual",
            allow_confirmed_duplicates=True,
            business_date="2026-08-06",
        )
        self.assertEqual(follow_up["scan"]["status"], "run_created")
        self.assertEqual(follow_up["scan"]["new_order_count"], 1)
        self.assertEqual(follow_up["scan"]["skipped_confirmed_count"], 19)

    def test_manual_source_cancellation_reconciliation_ledgers_only_verified_remaining_orders(self):
        orders = make_orders(3)
        service = self.service(orders)
        source = self.source("orders.xls")
        state = service.create_run_from_source(source)
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.record_automated_submission_failure(batch["id"], {"error": "stopped"})

        # The replacement report reflects Sopoka deleting the first duplicate.
        orders[:] = orders[1:]
        replacement = self.source("replacement.xls", b"replacement")
        evidence = [
            {
                "order_code": order["order_code"],
                "kiot_order_code": f"KV-{index}",
                "customer_code": order["customer_code"],
                "customer_name": order["customer_name"],
            }
            for index, order in enumerate(orders, 2)
        ]

        reconciled = service.reconcile_manual_batch_after_source_cancellation(
            batch["id"],
            replacement,
            cancelled_order_code="DH-001",
            evidence=evidence,
        )

        self.assertEqual(reconciled["run"]["status"], "completed")
        self.assertEqual(reconciled["run"]["batches"][0]["status"], "confirmed")
        self.assertEqual(service.ledger_orders("2026-08-06")["count"], 2)
        self.assertEqual(
            service.db.scalar(
                "SELECT COUNT(*) FROM confirmed_order_ledger WHERE confirmation_method = ?",
                ("manual_user_confirmation",),
            ),
            2,
        )
        self.assertEqual(
            service.db.scalar(
                "SELECT status FROM orders WHERE order_code = ?", ("DH-001",)
            ),
            "source_cancelled",
        )
        follow_up = service.scan_and_create_incremental_run(
            replacement,
            trigger="manual",
            allow_confirmed_duplicates=True,
            business_date="2026-08-06",
        )
        self.assertEqual(follow_up["scan"]["status"], "no_new")

    def test_delete_selected_orders_requires_every_code_and_removes_failed_run_rows(self):
        service = self.service(make_orders(2))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.record_automated_submission_failure(batch["id"], {"error": "timeout"})
        run_id = state["run"]["id"]

        with self.assertRaises(InvalidTransitionError):
            service.delete_selected_run_orders(run_id, ["DH-001"], confirm=True)

        deleted = service.delete_selected_run_orders(
            run_id, ["DH-001", "DH-002"], confirm=True
        )
        self.assertEqual(deleted["deleted_count"], 2)
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM orders WHERE run_id = ?", (run_id,)), 0)
        self.assertEqual(service.state(run_id)["run"]["status"], "cancelled")

    def test_ledger_listing_and_selected_delete_releases_codes_for_a_future_scan(self):
        service = self.service(make_orders(3))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(batch)},
            batch_id=batch["id"],
        )
        service.confirm_batch(batch["id"])

        ledger = service.ledger_orders("2026-08-06")
        self.assertEqual(ledger["count"], 3)
        self.assertEqual({item["order_code"] for item in ledger["orders"]}, {"DH-001", "DH-002", "DH-003"})

        with self.assertRaises(Exception):
            service.delete_selected_ledger_orders(
                "2026-08-06", ["DH-001"], confirm=True,
            )
        deleted = service.delete_selected_ledger_orders(
            "2026-08-06",
            ["DH-001"],
            confirm=True,
            acknowledge_external_order_risk=True,
        )
        self.assertEqual(deleted["deleted_count"], 1)
        self.assertEqual(service.ledger_orders("2026-08-06")["count"], 2)
        self.assertIsNone(service.db.fetch_one("SELECT id FROM orders WHERE order_code = ?", ("DH-001",)))
        batch_row = service.db.fetch_one("SELECT status, sheet_count FROM batches WHERE id = ?", (batch["id"],))
        self.assertEqual((batch_row["status"], batch_row["sheet_count"]), ("confirmed", 2))
        self.assertEqual(service.state(state["run"]["id"])["run"]["order_count"], 2)

    def test_failed_batch_requires_explicit_retry(self):
        service = self.service(make_orders(1))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        state = service.mark_extension_event(
            "batch.error", {"error": "upload failed"}, batch_id=batch["id"]
        )
        self.assertTrue(state["run"]["batches"][0]["can_retry"])
        state = service.retry_batch(batch["id"])
        self.assertTrue(state["run"]["batches"][0]["can_import"])

    def test_note_verification_rebases_a_shifted_order_tab_label(self):
        service = self.service(make_orders(1))
        state = service.create_run_from_source(self.source("orders.xls"))
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])

        service.mark_extension_event(
            "batch.notes_filled",
            {
                "verified": [
                    {"order_code": "DH-001", "tab_label": "Đặt hàng 1"}
                ]
            },
            batch_id=batch["id"],
        )

        current = service.state(state["run"]["id"])["run"]["batches"][0]
        self.assertEqual(current["status"], "waiting_user_confirmation")
        self.assertTrue(current["orders"][0]["note_verified"])
        self.assertEqual(current["orders"][0]["expected_tab_label"], "Đặt hàng 1")
        self.assertEqual(current["orders"][0]["actual_tab_label"], "Đặt hàng 1")

    def test_prior_day_active_run_expires_and_does_not_block_new_run(self):
        service = self.service(make_orders(1))
        old = service.create_run_from_source(self.source("old.xls", b"old"))
        old_run_id = old["run"]["id"]

        self.now = datetime(2026, 8, 7, 9, 0, tzinfo=BANGKOK_TZ)
        new = service.create_run_from_source(self.source("new.xls", b"new"))

        self.assertNotEqual(new["run"]["id"], old_run_id)
        self.assertEqual(new["run"]["business_date"], "2026-08-07")
        self.assertEqual(service.state(old_run_id)["run"]["status"], "expired")

    def test_selected_historical_scan_uses_requested_business_date(self):
        service = self.service(make_orders(2))
        state = service.scan_and_create_incremental_run(
            self.source("yesterday.xls", b"yesterday"),
            business_date="2026-08-05",
        )
        self.assertEqual(state["run"]["business_date"], "2026-08-05")
        self.assertEqual(state["run"]["order_count"], 2)
        self.assertEqual(state["run"]["status"], "batches_ready")
        self.assertTrue(state["run"]["historical_run"])

        # An explicitly selected historical workflow remains resumable after
        # the operational date advances; it must not be expired by a state
        # refresh before its Edge batches are completed.
        self.now = datetime(2026, 8, 7, 9, 0, tzinfo=BANGKOK_TZ)
        resumed = service.state(state["run"]["id"])
        self.assertEqual(resumed["run"]["status"], "batches_ready")

    def test_duplicate_check_resets_on_new_business_date(self):
        service = self.service(make_orders(1))
        first_source = self.source("first.xls", b"first")
        state = service.create_run_from_source(first_source)
        state = service.prepare_batches(state["run"]["id"])
        batch = state["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(batch)},
            batch_id=batch["id"],
        )
        service.confirm_batch(batch["id"])

        duplicate = service.create_run_from_source(self.source("same-day.xls", b"second"))
        self.assertEqual(duplicate["run"]["status"], "duplicate_review")
        self.assertEqual(duplicate["run"]["duplicate_count"], 1)
        service.create_run_from_source(
            duplicate["run"]["source_path"], allow_duplicates=True
        )
        duplicate_batch = service.prepare_batches(duplicate["run"]["id"])["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=duplicate_batch["id"])
        service.mark_extension_event("batch.imported", batch_id=duplicate_batch["id"])
        service.mark_extension_event(
            "batch.notes_filled",
            {"verified": verified_notes(duplicate_batch)},
            batch_id=duplicate_batch["id"],
        )
        # The legacy override route can still build a duplicate for backwards
        # compatibility, but the durable ledger must refuse to mark it as
        # confirmed.  New scan-based workflow never creates this batch.
        with self.assertRaises(InvalidTransitionError):
            service.confirm_batch(duplicate_batch["id"])

        self.now = datetime(2026, 8, 7, 9, 0, tzinfo=BANGKOK_TZ)
        next_day = service.create_run_from_source(self.source("next-day.xls", b"third"))
        self.assertEqual(next_day["run"]["status"], "parsed")
        self.assertEqual(next_day["run"]["duplicate_count"], 0)

    def test_incremental_scan_builds_a_run_for_new_orders_only(self):
        orders = make_orders(1)
        service = self.service(orders)
        first = service.create_run_from_source(self.source("first.xls", b"first"))
        first = service.prepare_batches(first["run"]["id"])
        batch = first["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled", {"verified": verified_notes(batch)}, batch_id=batch["id"]
        )
        service.confirm_batch(batch["id"])
        self.assertEqual(
            service.db.scalar("SELECT COUNT(*) FROM confirmed_order_ledger"), 1
        )

        orders.append(make_orders(2)[1])
        state = service.scan_and_create_incremental_run(
            self.source("latest.xls", b"latest"),
            allow_confirmed_duplicates=True,
        )

        scan = state["scan"]
        self.assertEqual(scan["status"], "run_created")
        self.assertEqual(scan["skipped_confirmed_count"], 1)
        self.assertEqual(scan["new_order_count"], 1)
        self.assertEqual([item["order_code"] for item in state["run"]["orders"]], ["DH-002"])
        self.assertEqual(state["run"]["batches"][0]["order_count"], 1)

    def test_incremental_scan_requires_decision_before_skipping_confirmed_codes(self):
        orders = make_orders(1)
        service = self.service(orders)
        first = service.create_run_from_source(self.source("first.xls", b"first"))
        first = service.prepare_batches(first["run"]["id"])
        batch = first["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled", {"verified": verified_notes(batch)}, batch_id=batch["id"]
        )
        service.confirm_batch(batch["id"])

        orders.append(make_orders(2)[1])
        held = service.scan_and_create_incremental_run(self.source("latest.xls", b"latest"))

        self.assertTrue(held["requires_duplicate_confirmation"])
        self.assertEqual(held["scan"]["status"], "duplicate_review")
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM runs"), 1)

        continued = service.decide_scan_duplicates(held["scan"]["id"], continue_scan=True)
        self.assertEqual(continued["scan"]["status"], "run_created")
        self.assertEqual([item["order_code"] for item in continued["run"]["orders"]], ["DH-002"])

    def test_incremental_scan_holds_unconfirmed_codes_and_creates_nothing(self):
        orders = make_orders(1)
        service = self.service(orders)
        first = service.create_run_from_source(self.source("first.xls", b"first"))
        first = service.prepare_batches(first["run"]["id"])
        batch = first["run"]["batches"][0]
        service.mark_extension_event("batch.import_started", batch_id=batch["id"])
        service.mark_extension_event("batch.imported", batch_id=batch["id"])
        service.mark_extension_event(
            "batch.notes_filled", {"verified": verified_notes(batch)}, batch_id=batch["id"]
        )

        orders.append(make_orders(2)[1])
        state = service.scan_and_create_incremental_run(self.source("latest.xls", b"latest"))

        scan = state["scan"]
        self.assertEqual(scan["status"], "unresolved")
        self.assertEqual(scan["unresolved_count"], 1)
        self.assertEqual(scan["new_order_count"], 1)
        self.assertEqual(service.db.scalar("SELECT COUNT(*) FROM runs"), 1)
        self.assertEqual(
            [item["decision"] for item in scan["decisions"]], ["unresolved", "new"]
        )

    def test_scan_accepts_the_audited_scheduler_trigger_without_stateful_scheduler(self):
        service = self.service(make_orders(1))

        self.assertNotIn("scheduler", service.state())
        result = service.scan_and_create_incremental_run(
            self.source("scheduled.xls"), trigger="scheduled"
        )
        self.assertEqual(result["scan"]["status"], "run_created")
        with self.assertRaises(SourceValidationError):
            service.scan_and_create_incremental_run(
                self.source("invalid-trigger.xls"), trigger="background"
            )


if __name__ == "__main__":
    unittest.main()
