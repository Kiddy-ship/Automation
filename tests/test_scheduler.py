from __future__ import annotations

import unittest
from datetime import date
from types import SimpleNamespace
from unittest.mock import Mock

from automation_app.scheduler import ScheduledWorkflowRunner


class ScheduledWorkflowRunnerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.config = SimpleNamespace(
            today=lambda: date(2026, 9, 24),
            scheduler_max_clock_skew_seconds=120,
        )
        self.service = Mock()
        self.edge = Mock()
        self.auto_run = Mock()
        self.runner = ScheduledWorkflowRunner(
            self.config, self.service, self.edge, self.auto_run
        )

    def test_active_or_attention_run_does_not_touch_edge(self):
        self.service.state.return_value = {
            "run": {"id": 71, "status": "automation_attention"}
        }

        result = self.runner.run_once()

        self.assertFalse(result["ok"])
        self.assertEqual(result["outcome"], "blocked")
        self.assertEqual(result["run_id"], 71)
        self.edge.execute.assert_not_called()
        self.service.scan_and_create_incremental_run.assert_not_called()

    def test_no_new_orders_closes_only_after_verified_export(self):
        self.service.state.return_value = {"run": {"id": 70, "status": "cancelled"}}
        self.edge.execute.side_effect = [
            {"ok": True, "business_date": "2026-09-24"},
            {
                "ok": True,
                "business_date": "2026-09-24",
                "source_path": r"C:\reports\today.xls",
                "order_count": 28,
            },
            {"ok": True, "closed": ["http://dms.sopoka.com.vn/reports/summaryOrderSales"]},
        ]
        self.service.scan_and_create_incremental_run.return_value = {
            "scan": {"status": "no_new", "new_order_count": 0}
        }

        result = self.runner.run_once()

        self.assertTrue(result["ok"])
        self.assertEqual(result["outcome"], "no_new")
        self.service.scan_and_create_incremental_run.assert_called_once_with(
            r"C:\reports\today.xls",
            trigger="scheduled",
            allow_confirmed_duplicates=True,
            business_date="2026-09-24",
        )
        self.assertEqual(
            [item.args[0] for item in self.edge.execute.call_args_list],
            ["SCHEDULE_PREFLIGHT", "EXPORT_DATE", "CLOSE_SCHEDULED_TABS"],
        )
        self.auto_run.assert_not_called()

    def test_completed_new_run_is_closed_only_after_every_batch_is_confirmed(self):
        self.service.state.side_effect = [
            {"run": {"id": 70, "status": "cancelled"}},
            {
                "run": {
                    "id": 72,
                    "status": "completed",
                    "batches": [{"status": "confirmed"}],
                }
            },
        ]
        self.edge.execute.side_effect = [
            {"ok": True, "business_date": "2026-09-24"},
            {
                "ok": True,
                "business_date": "2026-09-24",
                "source_path": r"C:\reports\today.xls",
            },
            {"ok": True, "closed": ["https://nongsanlamhong.kiotviet.vn/sale/#/"]},
        ]
        self.service.scan_and_create_incremental_run.return_value = {
            "scan": {"status": "run_created", "new_order_count": 1},
            "run": {"id": 72},
        }
        self.auto_run.return_value = {"ok": True, "run_id": 72}

        result = self.runner.run_once()

        self.assertTrue(result["ok"])
        self.assertEqual(result["outcome"], "completed")
        self.auto_run.assert_called_once_with(72)
        self.assertEqual(self.edge.execute.call_args_list[-1].args[0], "CLOSE_SCHEDULED_TABS")

    def test_blocked_scan_keeps_browser_open_for_operator_review(self):
        self.service.state.return_value = {"run": {"id": 70, "status": "cancelled"}}
        self.edge.execute.side_effect = [
            {"ok": True, "business_date": "2026-09-24"},
            {
                "ok": True,
                "business_date": "2026-09-24",
                "source_path": r"C:\reports\today.xls",
            },
        ]
        self.service.scan_and_create_incremental_run.return_value = {
            "scan": {"status": "changed_confirmed", "error_message": "payload changed"}
        }

        result = self.runner.run_once()

        self.assertFalse(result["ok"])
        self.assertEqual(result["outcome"], "blocked")
        self.assertEqual(self.edge.execute.call_count, 2)
        self.auto_run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
