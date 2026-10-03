from __future__ import annotations

import unittest
from unittest.mock import Mock, patch

import app as app_module
from automation_app.service import InvalidTransitionError


class MutableClock:
    def __init__(self) -> None:
        self.value = 1_000.0

    def __call__(self) -> float:
        return self.value


class ChromeCommandBrokerTests(unittest.TestCase):
    def test_queued_command_expires_and_releases_the_broker(self):
        clock = MutableClock()
        broker = app_module.ChromeCommandBroker(
            queued_timeout_seconds=5,
            dispatched_timeout_seconds=10,
            clock=clock,
        )
        stale = broker.submit("PING")

        clock.value += 5
        replacement = broker.submit("INSPECT_KIOT")

        self.assertEqual(broker.get(stale["id"])["status"], "expired")
        self.assertEqual(replacement["status"], "queued")

    def test_dispatched_command_expires_and_rejects_a_late_result(self):
        clock = MutableClock()
        broker = app_module.ChromeCommandBroker(
            queued_timeout_seconds=5,
            dispatched_timeout_seconds=10,
            clock=clock,
        )
        command = broker.submit("PING")
        broker.poll()

        clock.value += 10
        replacement = broker.submit("INSPECT_KIOT")

        self.assertEqual(broker.get(command["id"])["status"], "expired")
        self.assertEqual(replacement["status"], "queued")
        with self.assertRaises(InvalidTransitionError):
            broker.complete(command["id"], {"ok": True})


class ConfirmationEndpointTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def test_confirm_requires_literal_json_true(self):
        fake_service = Mock()
        fake_service.confirm_batch.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in ({}, {"user_confirmed": False}, {"user_confirmed": 1}):
                response = self.client.post("/api/batches/7/confirm", json=payload)
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                "/api/batches/7/confirm", json={"user_confirmed": True}
            )

        self.assertEqual(response.status_code, 200)
        fake_service.confirm_batch.assert_called_once_with(batch_id=7)

    def test_retry_requires_literal_json_cleanup_confirmation(self):
        fake_service = Mock()
        fake_service.retry_batch.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in (
                {},
                {"user_confirmed_cleanup": False},
                {"user_confirmed_cleanup": "true"},
            ):
                response = self.client.post("/api/batches/9/retry", json=payload)
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                "/api/batches/9/retry",
                json={"user_confirmed_cleanup": True},
            )

        self.assertEqual(response.status_code, 200)
        fake_service.retry_batch.assert_called_once_with(9)

    def test_abort_requires_all_explicit_safety_confirmations(self):
        fake_service = Mock()
        fake_service.abort_run.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in (
                {},
                {"confirm": True},
                {"confirm": True, "kiot_tabs_cleared": True},
                {
                    "confirm": True,
                    "kiot_tabs_cleared": True,
                    "acknowledge_external_order_risk": "true",
                },
            ):
                response = self.client.post("/api/runs/9/abort", json=payload)
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                "/api/runs/9/abort",
                json={
                    "confirm": True,
                    "kiot_tabs_cleared": True,
                    "acknowledge_external_order_risk": True,
                },
            )

        self.assertEqual(response.status_code, 200)
        fake_service.abort_run.assert_called_once_with(
            9,
            confirm=True,
            kiot_tabs_cleared=True,
            acknowledge_external_order_risk=True,
        )


class SchedulerEndpointTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def test_scheduler_tick_is_loopback_only_and_uses_one_runner(self):
        fake_edge = Mock()
        fake_edge.status.return_value = {"progress": {"active": False}}
        expected = {"ok": True, "outcome": "no_new", "business_date": "2026-09-24"}
        with patch.object(app_module, "edge_automation", fake_edge), patch.object(
            app_module, "_run_scheduled_tick", return_value=expected
        ) as runner:
            response = self.client.post("/api/scheduler/tick", json={})

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), expected)
        runner.assert_called_once_with()

        response = self.client.post(
            "/api/scheduler/tick",
            json={},
            environ_overrides={"REMOTE_ADDR": "192.168.1.23"},
        )
        self.assertEqual(response.status_code, 400)

    def test_scheduler_tick_does_not_compete_with_an_active_edge_action(self):
        fake_edge = Mock()
        fake_edge.status.return_value = {"progress": {"active": True}}
        with patch.object(app_module, "edge_automation", fake_edge), patch.object(
            app_module, "_run_scheduled_tick"
        ) as runner:
            response = self.client.post("/api/scheduler/tick", json={})

        self.assertEqual(response.status_code, 200)
        self.assertFalse(response.get_json()["ok"])
        self.assertEqual(response.get_json()["outcome"], "blocked")
        runner.assert_not_called()

    def test_manual_source_cancellation_reconciliation_requires_audited_payload(self):
        fake_service = Mock()
        fake_service.reconcile_manual_batch_after_source_cancellation.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in (
                {},
                {"user_confirmed": True},
                {
                    "user_confirmed": True,
                    "source_path": "report.xls",
                    "cancelled_order_code": "DH-001",
                    "evidence": {},
                },
            ):
                response = self.client.post(
                    "/api/batches/9/reconcile-manual-source-cancellation",
                    json=payload,
                )
                self.assertEqual(response.status_code, 400)
            evidence = [{"order_code": "DH-002", "kiot_order_code": "KV-2"}]
            response = self.client.post(
                "/api/batches/9/reconcile-manual-source-cancellation",
                json={
                    "user_confirmed": True,
                    "source_path": "report.xls",
                    "cancelled_order_code": "DH-001",
                    "evidence": evidence,
                },
            )

        self.assertEqual(response.status_code, 200)
        fake_service.reconcile_manual_batch_after_source_cancellation.assert_called_once_with(
            9,
            "report.xls",
            cancelled_order_code="DH-001",
            evidence=evidence,
        )

    def test_delete_selected_requires_literal_confirmation_and_forwards_codes(self):
        fake_service = Mock()
        fake_service.delete_selected_run_orders.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in (
                {},
                {"order_codes": ["DH-001"], "confirm": False},
                {"order_codes": ["DH-001"], "confirm": 1},
            ):
                response = self.client.post("/api/runs/5/delete-selected", json=payload)
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                "/api/runs/5/delete-selected",
                json={"order_codes": ["DH-001"], "confirm": True},
            )

        self.assertEqual(response.status_code, 200)
        fake_service.delete_selected_run_orders.assert_called_once_with(
            5, ["DH-001"], confirm=True
        )

    def test_ledger_delete_requires_external_order_risk_acknowledgement(self):
        fake_service = Mock()
        fake_service._today.return_value = "2026-08-10"
        fake_service.delete_selected_ledger_orders.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in (
                {},
                {"business_date": "2026-08-10", "order_codes": ["DH-001"], "confirm": True},
                {"business_date": "2026-08-10", "order_codes": ["DH-001"], "confirm": False, "acknowledge_external_order_risk": True},
            ):
                response = self.client.post("/api/ledger/orders/delete", json=payload)
                self.assertEqual(response.status_code, 400)
            response = self.client.post(
                "/api/ledger/orders/delete",
                json={
                    "business_date": "2026-08-10",
                    "order_codes": ["DH-001"],
                    "confirm": True,
                    "acknowledge_external_order_risk": True,
                },
            )

        self.assertEqual(response.status_code, 200)
        fake_service.delete_selected_ledger_orders.assert_called_once_with(
            "2026-08-10",
            ["DH-001"],
            confirm=True,
            acknowledge_external_order_risk=True,
        )

    def test_automation_completion_requires_an_edge_result_object(self):
        fake_service = Mock()
        fake_service.complete_automated_batch.return_value = {"ok": True}
        with patch.object(app_module, "service", fake_service):
            for payload in ({}, {"submission_result": []}, {"submission_result": "done"}):
                response = self.client.post(
                    "/api/batches/11/automation-complete", json=payload
                )
                self.assertEqual(response.status_code, 400)
            edge_result = {
                "ok": True,
                "finalOrderClicked": True,
                "submitted": {"ok": True, "submitted": []},
            }
            response = self.client.post(
                "/api/batches/11/automation-complete",
                json={"submission_result": edge_result},
            )

        self.assertEqual(response.status_code, 200)
        fake_service.complete_automated_batch.assert_called_once_with(11, edge_result)


class AutomatedQueueEndpointTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    @staticmethod
    def _state(status, batch):
        return {
            "ok": True,
            "run": {
                "id": 31,
                "status": status,
                "batches": [batch],
            },
        }

    def test_queue_rejects_a_cancelled_run_before_it_touches_edge(self):
        batch = {
            "id": 41,
            "batch_number": 1,
            "status": "cancelled",
            "is_current": False,
        }
        fake_service = Mock()
        fake_service.state.return_value = self._state("cancelled", batch)
        fake_edge = Mock()

        with patch.object(app_module, "service", fake_service), patch.object(
            app_module, "edge_automation", fake_edge
        ):
            response = self.client.post("/api/runs/31/auto-run", json={})

        self.assertEqual(response.status_code, 409)
        fake_edge.execute.assert_not_called()
        fake_service.prepare_partial_automation_recovery.assert_not_called()

    def test_queue_imports_and_submits_before_it_reads_the_next_batch(self):
        ready = {
            "id": 41,
            "batch_number": 1,
            "status": "ready",
            "is_current": True,
        }
        confirmed = {**ready, "status": "confirmed", "is_current": False}
        edge_result = {
            "ok": True,
            "finalOrderClicked": True,
            "submitted": {"ok": True, "submitted": []},
        }
        fake_service = Mock()
        fake_service.state.side_effect = [
            self._state("batches_ready", ready),
            self._state("completed", confirmed),
        ]
        fake_service.batch_execution_payload.return_value = {
            "batchId": 41,
            "batchNumber": 1,
            "workbookPath": r"C:\\output\\batch.xlsx",
            "orders": [{"orderCode": "DH_1", "expectedTabLabel": "Đặt hàng 2"}],
        }
        fake_edge = Mock()
        fake_edge.execute.return_value = edge_result

        with patch.object(app_module, "service", fake_service), patch.object(
            app_module, "edge_automation", fake_edge
        ):
            response = self.client.post("/api/runs/31/auto-run", json={})

        self.assertEqual(response.status_code, 200)
        body = response.get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["completed_batches"][0]["batch_id"], 41)
        fake_service.mark_extension_event.assert_called_once()
        fake_edge.execute.assert_called_once_with(
            "IMPORT_BATCH",
            {
                "filePath": r"C:\\output\\batch.xlsx",
                "orders": [{"orderCode": "DH_1", "expectedTabLabel": "Đặt hàng 2"}],
                "batchId": 41,
                "batchNumber": 1,
                "autoSubmit": True,
            },
        )
        fake_service.complete_automated_batch.assert_called_once_with(41, edge_result)

    def test_queue_resumes_verified_tabs_without_uploading_again(self):
        waiting = {
            "id": 42,
            "batch_number": 2,
            "status": "waiting_user_confirmation",
            "is_current": True,
        }
        confirmed = {**waiting, "status": "confirmed", "is_current": False}
        edge_result = {
            "ok": True,
            "finalOrderClicked": True,
            "submitted": {"ok": True, "submitted": []},
        }
        fake_service = Mock()
        fake_service.state.side_effect = [
            self._state("waiting_user_confirmation", waiting),
            self._state("completed", confirmed),
        ]
        fake_service.batch_submission_payload.return_value = {
            "batchId": 42,
            "batchNumber": 2,
            "orders": [{"orderCode": "DH_2", "expectedTabLabel": "Đặt hàng 3"}],
        }
        fake_edge = Mock()
        fake_edge.execute.return_value = edge_result

        with patch.object(app_module, "service", fake_service), patch.object(
            app_module, "edge_automation", fake_edge
        ):
            response = self.client.post("/api/runs/31/auto-run", json={})

        self.assertEqual(response.status_code, 200)
        fake_service.batch_execution_payload.assert_not_called()
        fake_service.mark_extension_event.assert_not_called()
        fake_edge.execute.assert_called_once_with(
            "SUBMIT_BATCH",
            {
                "orders": [{"orderCode": "DH_2", "expectedTabLabel": "Đặt hàng 3"}],
                "batchId": 42,
                "batchNumber": 2,
            },
        )

    def test_queue_reimports_only_lost_partial_carts_from_a_blank_canvas(self):
        attention = {
            "id": 44,
            "batch_number": 1,
            "sheet_count": 2,
            "status": "automation_attention",
            "is_current": True,
        }
        confirmed = {**attention, "status": "confirmed", "is_current": False}
        prior = {
            "submitted": {
                "ok": True,
                "submitted": [{"orderCode": "DH_DONE", "success": True}],
            },
            "notes": {"ok": True, "verified": [{"orderCode": "DH_DONE", "tabLabel": "Order 2"}]},
        }
        partial = {
            "remainingOrders": [{"orderCode": "DH_OPEN", "expectedTabLabel": "Order 3"}],
            "priorResult": prior,
        }
        recovery = {
            "workbookPath": r"C:\output\recovery.xlsx",
            "orders": [{"orderCode": "DH_OPEN", "expectedTabLabel": "Order 2"}],
        }
        imported = {
            "ok": True,
            "imported": {"ok": True},
            "notes": {"ok": True, "verified": [{"orderCode": "DH_OPEN", "tabLabel": "Order 2"}]},
            "submitted": {"ok": True, "submitted": [{"orderCode": "DH_OPEN", "success": True}]},
            "finalOrderClicked": True,
        }
        fake_service = Mock()
        fake_service.state.side_effect = [
            self._state("automation_attention", attention),
            self._state("completed", confirmed),
        ]
        fake_service.prepare_partial_automation_recovery.return_value = partial
        fake_service.build_partial_recovery_import.return_value = recovery
        fake_edge = Mock()
        fake_edge.execute.side_effect = [
            {"ok": True, "orderTabs": ["Order 2"], "invoiceTabs": [], "noteValue": ""},
            imported,
        ]

        with patch.object(app_module, "service", fake_service), patch.object(
            app_module, "edge_automation", fake_edge
        ):
            response = self.client.post("/api/runs/31/auto-run", json={})

        self.assertEqual(response.status_code, 200)
        fake_service.build_partial_recovery_import.assert_called_once_with(44, ["DH_OPEN"])
        self.assertEqual(fake_edge.execute.call_args_list[1].args[0], "IMPORT_BATCH")
        fake_service.complete_automated_batch.assert_called_once()
        completed = fake_service.complete_automated_batch.call_args.args[1]
        self.assertEqual(completed["submitted"]["submitted"], [
            {"orderCode": "DH_DONE", "success": True},
            {"orderCode": "DH_OPEN", "success": True},
        ])
        self.assertEqual(completed["imported"], {"ok": True})

    def test_queue_quarantines_unsuccessful_edge_result_and_does_not_continue(self):
        ready = {
            "id": 43,
            "batch_number": 1,
            "status": "ready",
            "is_current": True,
        }
        fake_service = Mock()
        fake_service.state.return_value = self._state("batches_ready", ready)
        fake_service.batch_execution_payload.return_value = {
            "batchId": 43,
            "batchNumber": 1,
            "workbookPath": r"C:\\output\\batch.xlsx",
            "orders": [{"orderCode": "DH_3", "expectedTabLabel": "Đặt hàng 2"}],
        }
        fake_edge = Mock()
        fake_edge.execute.return_value = {"ok": False, "message": "one tab could not be verified"}

        with patch.object(app_module, "service", fake_service), patch.object(
            app_module, "edge_automation", fake_edge
        ):
            response = self.client.post("/api/runs/31/auto-run", json={})

        self.assertEqual(response.status_code, 409)
        fake_service.record_automated_submission_failure.assert_called_once()
        fake_service.complete_automated_batch.assert_not_called()
        self.assertEqual(fake_edge.execute.call_count, 1)


if __name__ == "__main__":
    unittest.main()
