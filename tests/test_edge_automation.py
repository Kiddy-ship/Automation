from __future__ import annotations

import base64
import io
import json
import threading
import tempfile
import time
import unittest
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from urllib.parse import quote
from unittest.mock import MagicMock, Mock, call, patch

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from automation_app.config import AppConfig
from automation_app.edge_automation import (
    KIOT_NEW_ORDER_URL,
    KIOT_ORDERS_URL,
    KIOT_SALES_URL,
    SOPKA_URL,
    EdgeAutomation,
    EdgeAutomationError,
)


def order(code: str, tab: str) -> dict[str, str]:
    return {"orderCode": code, "expectedTabLabel": tab}


class EdgeAutomationSubmissionTests(unittest.TestCase):
    def test_remote_clock_uses_http_date_and_measures_skew(self):
        class Response:
            headers = {"Date": "Thu, 24 Sep 2026 02:00:01 GMT"}

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        now = datetime(2026, 9, 24, 2, 0, 0, tzinfo=timezone.utc)
        with patch("automation_app.edge_automation.urlopen", return_value=Response()):
            result = EdgeAutomation._remote_clock(SOPKA_URL, now, 120)

        self.assertEqual(result["skew_seconds"], 1.0)
        self.assertTrue(result["within_limit"])

    def test_scheduled_browser_lifecycle_is_fail_closed_and_targeted(self):
        import inspect

        preflight = inspect.getsource(EdgeAutomation.scheduled_preflight)
        cleanup = inspect.getsource(EdgeAutomation.close_scheduled_tabs)
        self.assertIn("inspect_sopoka()", preflight)
        self.assertIn("inspect_kiot()", preflight)
        self.assertIn("not exactly one blank order tab", preflight)
        self.assertIn("_remote_clock", preflight)
        self.assertIn("dms.sopoka.com.vn", cleanup)
        self.assertIn("nongsanlamhong.kiotviet.vn", cleanup)
        self.assertNotIn("quit()", cleanup)

    def test_safe_context_error_is_recoverable_before_any_click(self):
        self.assertTrue(EdgeAutomation._is_safe_submission_context_error(
            EdgeAutomationError("KiotViet is not on the required branch Đồng Nai 2 - ĐN2; final order was not clicked.")
        ))
        self.assertTrue(EdgeAutomation._is_safe_submission_context_error(
            EdgeAutomationError("KiotViet still has an Invoice tab; final order was not clicked.")
        ))
        self.assertFalse(EdgeAutomation._is_safe_submission_context_error(
            EdgeAutomationError("No KiotViet confirmation after final order; stopping without retry.")
        ))

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="orderbridge-edge-tests-")
        root = Path(self.temp.name)
        self.automation = EdgeAutomation(
            AppConfig(
                data_dir=root / "data",
                outputs_dir=root / "outputs",
                logs_dir=root / "logs",
                temp_dir=root / "temp",
            )
        )

    def tearDown(self) -> None:
        self._close_test_log_handler()
        self.temp.cleanup()

    def _close_test_log_handler(self) -> None:
        root = Path(self.temp.name).resolve()
        for handler in list(self.automation._logger.handlers):
            filename = getattr(handler, "baseFilename", None)
            try:
                is_test_handler = filename is not None and Path(filename).resolve().is_relative_to(root)
            except (OSError, ValueError):
                is_test_handler = False
            if is_test_handler:
                handler.close()
                self.automation._logger.removeHandler(handler)

    def _install_fake_browser(self, *, fail_code: str | None = None) -> list[tuple[str, bool]]:
        calls: list[tuple[str, bool]] = []
        self.automation._kiot_status = lambda: {
            "branchOk": True,
            "orderMode": True,
            "invoiceTabs": [],
            "orderTabs": ["Đặt hàng 1", "Đặt hàng 2", "Đặt hàng 3", "Đặt hàng 4"],
        }

        def run_task(name, expectations, _branch, allow_existing_invoice_tabs, **_kwargs):
            self.assertEqual(name, "kiotSubmitOrdersTask")
            self.assertEqual(len(expectations), 1)
            expected = expectations[0]
            code = expected["orderCode"]
            calls.append((code, bool(allow_existing_invoice_tabs)))
            if code == fail_code:
                return {"error": f"failed {code}"}
            return {
                "submitted": [
                    {
                        "orderCode": code,
                        "actualTabLabel": expected["expectedTabLabel"],
                        "success": True,
                        "completionProof": "tab_closed",
                    }
                ]
            }

        self.automation._run_task = run_task
        self.automation._dismiss_print_preview = lambda **_kwargs: None
        self.automation._watch_and_cancel_print_preview = (
            lambda stop, _state: stop.wait(0.05)
        )
        return calls

    def test_submission_uses_descending_tab_order_and_exact_outcomes(self):
        calls = self._install_fake_browser()
        result = self.automation._submit_verified_orders(
            [
                order("DH-002", "Đặt hàng 2"),
                order("DH-004", "Đặt hàng 4"),
                order("DH-003", "Đặt hàng 3"),
            ],
            batch_label="7",
        )

        self.assertTrue(result["ok"])
        self.assertTrue(result["allSubmitted"])
        self.assertEqual(calls, [("DH-004", False), ("DH-003", True), ("DH-002", True)])
        self.assertEqual(
            [item["actualTabLabel"] for item in result["submitted"]],
            ["Đặt hàng 4", "Đặt hàng 3", "Đặt hàng 2"],
        )
        self.assertTrue(all(item["success"] for item in result["submitted"]))

    def test_submission_stops_after_first_failed_order_and_keeps_partial_result(self):
        calls = self._install_fake_browser(fail_code="DH-003")
        result = self.automation._submit_verified_orders(
            [order("DH-002", "Đặt hàng 2"), order("DH-003", "Đặt hàng 3")]
        )

        self.assertFalse(result["ok"])
        self.assertFalse(result["allSubmitted"])
        self.assertEqual(calls, [("DH-003", False)])
        self.assertEqual(result["submitted"][0]["status"], "failed")
        self.assertEqual(result["submitted"][0]["orderCode"], "DH-003")

    def test_submission_plan_rejects_duplicate_or_invalid_tab_labels(self):
        with self.assertRaises(EdgeAutomationError):
            self.automation._submission_plan(
                [order("DH-001", "Đặt hàng 2"), order("DH-002", "Đặt hàng 2")]
            )
        with self.assertRaises(EdgeAutomationError):
            self.automation._submission_plan([order("DH-001", "Hóa đơn 2")])

    def test_print_watcher_retries_until_cancel_control_is_ready(self):
        preview = {"id": "print-target-1"}
        attempts: list[dict[str, float]] = []
        polls = {"count": 0}
        state: dict[str, object] = {"cancelled": 0, "methods": []}
        stop = threading.Event()

        def find_target(_capabilities):
            polls["count"] += 1
            return (None, True) if len(attempts) >= 2 else (preview, True)

        def dismiss(**kwargs):
            attempts.append(kwargs)
            return "devtools_cancel" if len(attempts) >= 2 else None

        self.automation._driver = SimpleNamespace(capabilities={})
        self.automation._find_edge_print_target = find_target
        self.automation._dismiss_print_preview = dismiss
        thread = threading.Thread(
            target=self.automation._watch_and_cancel_print_preview,
            args=(stop, state),
            daemon=True,
        )
        thread.start()
        time.sleep(0.95)
        stop.set()
        thread.join(timeout=2)

        self.assertGreaterEqual(len(attempts), 2)
        self.assertEqual(state["cancelled"], 1)
        self.assertEqual(state["methods"], ["devtools_cancel"])

    def test_dismiss_print_preview_prefers_the_real_cancel_control(self):
        """A DevTools target close can acknowledge without dismissing Edge UI."""

        preview = {"id": "print-target-1"}
        self.automation._driver = SimpleNamespace(capabilities={})
        self.automation._find_edge_print_target = lambda _capabilities: (preview, True)
        self.automation._click_print_cancel_via_devtools = Mock(return_value="devtools_cancel")
        self.automation._close_edge_print_target = Mock(return_value=(True, True))
        self.automation._click_visible_print_cancel = Mock(return_value=None)

        method = self.automation._dismiss_print_preview(wait_seconds=0.1)

        self.assertIn(method, {"devtools_cancel", "devtools_cancel+devtools_close"})
        self.automation._click_print_cancel_via_devtools.assert_called_once_with(preview)
        # Newer Edge builds keep the target alive after Cancel; the guarded
        # close fallback is expected in that case.

    def test_submission_refuses_a_partial_resume_with_existing_invoice_tabs(self):
        self.automation._kiot_status = lambda: {
            "branchOk": True,
            "orderMode": True,
            "invoiceTabs": ["Hóa đơn 3"],
            "orderTabs": ["Đặt hàng 1", "Đặt hàng 2"],
        }
        with self.assertRaisesRegex(EdgeAutomationError, "chưa có tab Hóa đơn"):
            self.automation._assert_submission_preflight(
                self.automation._submission_plan([order("DH-002", "Đặt hàng 2")])
            )

    def test_submit_batch_keeps_nested_completion_proof_for_backend(self):
        proof = {
            "ok": True,
            "allSubmitted": True,
            "submitted": [
                {
                    "orderCode": "DH-002",
                    "tabLabel": "Đặt hàng 2",
                    "success": True,
                }
            ],
            "failed": None,
            "message": "Đã xong",
        }
        self.automation._submit_verified_orders = lambda _orders, **_kwargs: proof

        result = self.automation.submit_batch(
            {"orders": [order("DH-002", "Đặt hàng 2")], "batchId": 7}
        )

        self.assertTrue(result["ok"])
        self.assertTrue(result["finalOrderClicked"])
        self.assertIs(result["submitted"], proof)
        self.assertEqual(result["submitted"]["submitted"][0]["tabLabel"], "Đặt hàng 2")

    def test_status_progress_does_not_wait_for_browser_action_lock(self):
        self.automation._set_progress(
            active=True,
            phase="submitting",
            message="Đang đặt đơn 1/2",
            current=1,
            total=2,
            batch="4",
            order_code="DH-001",
        )
        self.automation._lock.acquire()
        try:
            status = self.automation.status()
        finally:
            self.automation._lock.release()

        self.assertEqual(status["progress"]["phase"], "submitting")
        self.assertEqual(status["progress"]["current"], 1)
        self.assertEqual(status["progress"]["orderCode"], "DH-001")


class EdgeAutomationKiotReadinessTests(unittest.TestCase):
    """Regression tests for KiotViet's initial SPA loading spinner."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="orderbridge-kiot-ready-tests-")
        root = Path(self.temp.name)
        self.automation = EdgeAutomation(
            AppConfig(
                data_dir=root / "data",
                outputs_dir=root / "outputs",
                logs_dir=root / "logs",
                temp_dir=root / "temp",
            )
        )

    def tearDown(self) -> None:
        for handler in list(self.automation._logger.handlers):
            filename = getattr(handler, "baseFilename", None)
            try:
                is_test_handler = filename is not None and Path(filename).resolve().is_relative_to(
                    Path(self.temp.name).resolve()
                )
            except (OSError, ValueError):
                is_test_handler = False
            if is_test_handler:
                handler.close()
                self.automation._logger.removeHandler(handler)
        self.temp.cleanup()

    def test_status_retries_only_the_read_only_sync_inspection_after_timeout(self):
        expected = {"ok": True, "branchOk": True, "orderMode": True}
        self.automation._ensure_kiot_shell_ready = Mock()
        self.automation._run_sync_task = Mock(
            side_effect=[TimeoutException("script timeout"), expected]
        )

        result = self.automation._kiot_status()

        self.assertEqual(result, expected)
        self.automation._ensure_kiot_shell_ready.assert_has_calls([call(), call(force=True)])
        self.assertEqual(self.automation._run_sync_task.call_count, 2)
        self.assertEqual(
            self.automation._run_sync_task.call_args_list,
            [
                call("kiotInspectTask", self.automation.config.expected_branch),
                call("kiotInspectTask", self.automation.config.expected_branch),
            ],
        )

    def test_shell_waits_for_two_stable_interactive_samples_before_inspection(self):
        spinner = {
            "ready": False,
            "url": "https://nongsanlamhong.kiotviet.vn/sale/#/",
            "branchVisible": False,
            "tabVisible": False,
            "noteVisible": False,
        }
        ready = {
            "ready": True,
            "url": "https://nongsanlamhong.kiotviet.vn/sale/#/",
            "branchVisible": True,
            "tabVisible": True,
            "noteVisible": True,
        }
        self.automation._kiot_shell_snapshot = Mock(
            side_effect=[spinner, spinner, ready, ready]
        )

        with patch("automation_app.edge_automation.time.sleep") as sleep:
            self.automation._ensure_kiot_shell_ready(force=True)

        self.assertEqual(
            self.automation._last_kiot_shell_url,
            "https://nongsanlamhong.kiotviet.vn/sale/#/",
        )
        self.assertGreaterEqual(sleep.call_count, 2)

    def test_sync_inspection_uses_execute_script_without_async_callback(self):
        driver = Mock()
        driver.execute_script.return_value = {"ok": True, "branchOk": True}
        self.automation._ensure_driver = lambda: driver
        self.automation._task_source = lambda _name: "function kiotInspectTask() { return {ok: true}; }"

        result = self.automation._run_sync_task("kiotInspectTask", "Đồng Nai 2 - ĐN2")

        self.assertTrue(result["ok"])
        script = driver.execute_script.call_args.args[0]
        self.assertIn("return kiotInspectTask(...arguments);", script)
        self.assertFalse(hasattr(driver, "execute_async_script") and driver.execute_async_script.called)

    @staticmethod
    def _ready_kiot_state(*, branch_ok: bool = True) -> dict[str, object]:
        return {
            "ok": True,
            "url": KIOT_SALES_URL,
            "branch": "Đồng Nai 2 - ĐN2" if branch_ok else "Chi nhánh khác",
            "branchOk": branch_ok,
            "orderMode": True,
            "invoiceMode": False,
            "orderTabs": ["Đặt hàng 1"],
            "invoiceTabs": [],
            "noteAvailable": True,
        }

    def test_prepare_enters_sales_order_route_natively_without_async_page_task(self):
        driver = SimpleNamespace(current_url="https://nongsanlamhong.kiotviet.vn/man/#/Orders")
        navigations: list[str] = []
        ready = self._ready_kiot_state()
        self.automation._ensure_driver = lambda: driver
        self.automation._navigate = lambda url: navigations.append(url) or driver
        self.automation._assert_kiot_login = lambda: None
        self.automation._wait_until = lambda *_args, **_kwargs: ready
        self.automation._run_task = lambda *_args, **_kwargs: self.fail(
            "A ready native sales route must not invoke an async KiotViet page task."
        )

        result = self.automation.prepare_kiot()

        self.assertTrue(result["ok"])
        self.assertEqual(navigations, [KIOT_NEW_ORDER_URL])

    def test_prepare_requires_note_field_before_treating_existing_sales_tab_as_ready(self):
        driver = SimpleNamespace(current_url=KIOT_SALES_URL)
        navigations: list[str] = []
        loading = self._ready_kiot_state()
        loading["noteAvailable"] = False
        ready = self._ready_kiot_state()
        self.automation._ensure_driver = lambda: driver
        self.automation._select_existing_domain_window = lambda *_args: None
        self.automation._kiot_status = Mock(return_value=loading)
        self.automation._navigate = lambda url: navigations.append(url) or driver
        self.automation._assert_kiot_login = lambda: None
        self.automation._wait_until = lambda *_args, **_kwargs: ready
        self.automation._run_task = lambda *_args, **_kwargs: self.fail(
            "A native sales route with a ready note field must not use a page task."
        )

        result = self.automation.prepare_kiot()

        self.assertTrue(result["ok"])
        self.assertEqual(navigations, [KIOT_NEW_ORDER_URL])

    def test_prepare_handles_empty_canvas_with_native_route_not_cross_route_callback(self):
        driver = SimpleNamespace(current_url="https://nongsanlamhong.kiotviet.vn/man/#/Orders")
        navigations: list[str] = []
        initial = self._ready_kiot_state(branch_ok=False)
        final = self._ready_kiot_state(branch_ok=True)
        waits = iter([initial, final])
        self.automation._ensure_driver = lambda: driver
        self.automation._navigate = lambda url: navigations.append(url) or driver
        self.automation._assert_kiot_login = lambda: None
        self.automation._wait_until = lambda *_args, **_kwargs: next(waits)
        self.automation._run_task = Mock(return_value={"ok": True, "needsOrderRoute": True})

        result = self.automation.prepare_kiot()

        self.assertTrue(result["ok"])
        self.assertEqual(navigations, [KIOT_NEW_ORDER_URL, KIOT_NEW_ORDER_URL])
        self.assertEqual(self.automation._run_task.call_count, 1)
        self.assertEqual(self.automation._run_task.call_args.args[0], "kiotPrepareTask")

    def test_prepare_task_signals_native_order_route_instead_of_clicking_cross_route_link(self):
        source = self.automation._task_source("kiotPrepareTask")

        self.assertIn("needsOrderRoute", source)
        self.assertNotIn("safeClick(createOrder", source)

    def test_blank_order_task_uses_only_the_explicit_new_tab_control(self):
        source = self.automation._task_source("kiotCreateBlankOrderTabTask")

        self.assertIn("btn-add-tabs", source)
        self.assertIn("thêm mới", source.lower())
        self.assertIn("addCandidates", source)
        self.assertNotIn("ĐẶT HÀNG", source)

    def test_cleanup_creates_a_blank_tab_before_discarding_the_last_stale_note(self):
        stale = {
            "ok": True,
            "branchOk": True,
            "orderMode": True,
            "invoiceMode": False,
            "orderTabs": ["Đặt hàng 3"],
            "invoiceTabs": [],
            "noteValue": "ĐH_29573072026",
            "activeTab": "Đặt hàng 3",
        }
        expanded = {
            **stale,
            "invoiceTabs": ["Hóa đơn 4"],
            "activeTab": "Hóa đơn 4",
            "orderMode": False,
            "invoiceMode": True,
            "noteValue": "",
        }
        converted = {
            **stale,
            "orderTabs": ["Đặt hàng 3", "Đặt hàng 4"],
            "invoiceTabs": [],
            "noteValue": "",
            "activeTab": "Đặt hàng 4",
        }
        clean = {
            **converted,
            "orderTabs": ["Đặt hàng 4"],
            "activeTab": "Đặt hàng 4",
        }
        states = iter([stale, converted, clean])
        self.automation._kiot_status = lambda: next(states)
        self.automation._dismiss_kiot_discard_dialog = Mock(return_value=False)
        self.automation._find_existing_blank_order_tab = Mock(return_value=None)
        self.automation._create_blank_order_tab_native = Mock(
            return_value={"ok": True, "action": "create_blank_order_tab_native"}
        )
        self.automation._wait_until = Mock(side_effect=[expanded, converted])
        self.automation._wait_for_kiot_tab_close = Mock(return_value=clean)
        self.automation._run_task = Mock(
            side_effect=[
                {"ok": True, "action": "create_blank_order_tab"},
                {"ok": True, "strategy": "blank-invoice-to-order"},
                {"ok": True, "tabLabel": "Đặt hàng 3"},
            ]
        )

        result = self.automation.clean_kiot_tabs()

        self.assertTrue(result["ok"])
        self.assertEqual(result["closed"], 1)
        self.assertEqual(
            [call.args[0] for call in self.automation._run_task.call_args_list],
            [
                "kiotConvertActiveInvoiceTask",
                "kiotRequestCloseTabTask",
            ],
        )
        self.automation._create_blank_order_tab_native.assert_called_once_with()

    def test_cleanup_reuses_an_inactive_blank_order_tab_before_creating_one(self):
        initial = {
            "ok": True,
            "branchOk": True,
            "orderMode": True,
            "invoiceMode": False,
            "orderTabs": ["Đặt hàng 2", "Đặt hàng 3", "Đặt hàng 4"],
            "invoiceTabs": [],
            "activeTab": "Đặt hàng 2",
            "noteValue": "ĐH_30199632026",
            "cartItemCount": 6,
        }
        blank_active = {
            **initial,
            "activeTab": "Đặt hàng 3",
            "noteValue": "",
            "cartItemCount": 0,
        }
        after_first_close = {
            **blank_active,
            "orderTabs": ["Đặt hàng 2", "Đặt hàng 3"],
        }
        clean = {
            **blank_active,
            "orderTabs": ["Đặt hàng 3"],
        }
        states = iter([initial, blank_active, after_first_close, clean])
        self.automation._kiot_status = lambda: next(states)
        self.automation._dismiss_kiot_discard_dialog = Mock(return_value=False)
        self.automation._find_existing_blank_order_tab = Mock(return_value="Đặt hàng 3")
        self.automation._activate_kiot_tab_native = Mock()
        self.automation._wait_until = Mock(return_value=blank_active)
        self.automation._create_blank_order_tab_native = Mock()
        self.automation._wait_for_kiot_tab_close = Mock(side_effect=[after_first_close, clean])
        self.automation._run_task = Mock(
            side_effect=[
                {"ok": True, "tabLabel": "Đặt hàng 4"},
                {"ok": True, "tabLabel": "Đặt hàng 2"},
            ]
        )

        result = self.automation.clean_kiot_tabs()

        self.assertTrue(result["ok"])
        self.assertEqual(result["closed"], 2)
        self.automation._find_existing_blank_order_tab.assert_called_once_with()
        self.automation._activate_kiot_tab_native.assert_called_once_with("Đặt hàng 3")
        self.automation._create_blank_order_tab_native.assert_not_called()
        self.assertEqual(
            [call.args[0] for call in self.automation._run_task.call_args_list],
            ["kiotRequestCloseTabTask", "kiotRequestCloseTabTask"],
        )

    def test_cleanup_treats_order_to_invoice_conversion_as_failure_not_a_closed_tab(self):
        close_source = self.automation._task_source("kiotRequestCloseTabTask")
        self.assertIn('centerX > textRect.right + 2', close_source)
        self.assertIn('closeSide: "right_of_label"', close_source)

        import inspect
        wait_source = inspect.getsource(EdgeAutomation._wait_for_kiot_tab_close)
        self.assertIn("if current_count < before_count:", wait_source)
        self.assertNotIn("or tab_label not in labels", wait_source)

    def test_note_mapping_requires_visible_active_cart_text_not_document_wide_text(self):
        source = self.automation._task_source("kiotFillNotesTask")
        self.assertIn("matchingTabsByCode", source)
        self.assertIn("SKU + customer", source)
        self.assertIn("#cartListProduct", source)
        self.assertIn("inactive carts", source)
        self.assertNotIn("document.body?.innerText", source)

    def test_note_mapping_waits_for_each_cart_before_fallback_scan(self):
        source = self.automation._task_source("kiotFillNotesTask")
        self.assertIn("let labelsAlreadyVerified = false", source)
        self.assertIn("await verifyContent(expectation, tab.label)", source)
        self.assertIn("const matchingTabsByCode = new Map()", source)
        self.assertIn("Scan each cart once and wait for Angular", source)
        self.assertIn("const deadline = Date.now() + 5000", source)
        self.assertIn("const productsOk = (skuOk || nameOk) && countOk;", source)
        self.assertIn("return productsOk;", source)
        self.assertIn("productSignatureCounts", source)
        self.assertIn("requiresCustomerProof", source)
        self.assertIn("same product signature", source)
        self.assertIn("Cannot prove a unique KiotViet tab", source)
        self.assertIn("if (!labelsAlreadyVerified)", source)
        self.assertLess(
            source.index("await verifyContent(expectation, tab.label)"),
            source.index("const matchingTabsByCode = new Map()"),
        )

    def test_note_expectations_reject_duplicate_related_order_codes(self):
        with self.assertRaisesRegex(EdgeAutomationError, "trùng"):
            self.automation._normalize_note_expectations(
                [
                    {"orderCode": "ĐH_1", "customerName": "Khách A"},
                    {"orderCode": "ĐH_1", "customerName": "Khách B"},
                ]
            )

    def test_note_expectations_allow_same_products_for_different_customers_and_codes(self):
        normalized = self.automation._normalize_note_expectations(
            [
                {
                    "orderCode": "ĐH_1",
                    "customerName": "Khách A",
                    "productCodes": ["SP-01"],
                    "productNames": ["Sản phẩm 01"],
                    "lineCount": 1,
                },
                {
                    "orderCode": "ĐH_2",
                    "customerName": "Khách B",
                    "productCodes": ["SP-01"],
                    "productNames": ["Sản phẩm 01"],
                    "lineCount": 1,
                },
            ]
        )
        self.assertEqual([item["orderCode"] for item in normalized], ["ĐH_1", "ĐH_2"])

    def test_note_expectations_preserve_customer_and_cart_fingerprint(self):
        normalized = self.automation._normalize_note_expectations(
            [
                {
                    "orderCode": "ORDER-1",
                    "customerCode": "KH0001",
                    "customerName": "Customer A",
                    "paymentAmount": 1_615_680,
                    "productLines": [
                        {
                            "productCode": "D30W-T",
                            "productName": "Lamp 30W",
                            "quantity": 20,
                            "netAmount": 792_000,
                        }
                    ],
                }
            ]
        )

        self.assertEqual(normalized[0]["customerCode"], "KH0001")
        self.assertEqual(normalized[0]["paymentAmount"], 1_615_680)
        self.assertEqual(normalized[0]["productLines"][0]["quantity"], 20)
        self.assertEqual(normalized[0]["productLines"][0]["netAmount"], 792_000)

    def test_note_mapping_uses_quantity_and_amount_for_same_customer(self):
        source = self.automation._task_source("kiotFillNotesTask")
        self.assertIn("requiresDetailedProof", source)
        self.assertIn("detailedLinesMatch(expectation)", source)
        self.assertIn("paymentAmountMatches(expectation)", source)
        self.assertIn("rowQuantityValues", source)
        self.assertIn("customerCode", source)

    def test_note_mapping_matches_product_codes_as_full_identifiers(self):
        source = self.automation._task_source("kiotFillNotesTask")
        self.assertIn("const productCodeMatches", source)
        self.assertIn("const hasExactProductCode", source)
        self.assertIn("codes.every((code) => hasExactProductCode(code))", source)
        self.assertNotIn("codes.every((code) => hasVisibleText(code))", source)
        self.assertIn("E15``\n    // must not match ``E15N", source)

    def test_note_mapping_proves_a_candidate_shift_before_writing_notes(self):
        source = self.automation._task_source("kiotFillNotesTask")
        fast_path = source[source.index("const candidateShifts"):source.index("// Do not fail solely because one stale blank cart")]
        self.assertIn("const candidateMappings = []", fast_path)
        self.assertIn("if (!candidateOk) continue;", fast_path)
        self.assertLess(fast_path.index("if (!candidateOk) continue;"), fast_path.index("const fastVerified = []"))
        self.assertLess(fast_path.index("const fastVerified = []"), fast_path.index("const readBack = setValue(note, expectation.orderCode);"))

    def test_shell_is_not_considered_ready_for_management_url(self):
        self.assertTrue(self.automation._is_kiot_sales_url(KIOT_SALES_URL))
        self.assertFalse(self.automation._is_kiot_sales_url(KIOT_ORDERS_URL))

    def test_shell_snapshot_matches_visible_tab_and_note_without_encoding_specific_selector(self):
        class Element:
            def __init__(self, text="", **attributes):
                self.text = text
                self.attributes = attributes

            def is_displayed(self):
                return True

            def get_attribute(self, name):
                return self.attributes.get(name)

        class Driver:
            current_url = KIOT_SALES_URL

            def find_element(self, _by, _selector):
                return Element("Đồng Nai 2 - ĐN2 Đặt hàng 1")

            def find_elements(self, by, selector):
                if by == By.CSS_SELECTOR:
                    return [Element(placeholder="Ghi chú đơn hàng")]
                return [Element("Đặt hàng 1")]

        self.automation._ensure_driver = lambda: Driver()
        snapshot = self.automation._kiot_shell_snapshot()

        self.assertTrue(snapshot["ready"])
        self.assertTrue(snapshot["branchVisible"])
        self.assertTrue(snapshot["tabVisible"])
        self.assertTrue(snapshot["noteVisible"])


class EdgeAutomationPurchaseRecoveryTests(unittest.TestCase):
    def test_normalizes_only_audited_purchase_product_metadata(self):
        specs = EdgeAutomation._normalize_purchase_product_specs([
            {"product_code": "A90M2-10W-DM", "product_name": "Đèn A90M2", "source_row": 40},
            {"product_code": "", "product_name": "Bị bỏ qua", "source_row": 41},
        ])
        self.assertEqual(specs, [{"productCode": "A90M2-10W-DM", "productName": "Đèn A90M2", "sourceRow": 40}])

    def test_maps_unavailable_item_banner_to_the_exact_source_row(self):
        specs = EdgeAutomation._normalize_purchase_product_specs([
            {"product_code": "A90M2-10W-DM", "product_name": "Đèn A90M2", "source_row": 40},
            {"product_code": "A90M3-10W-DM", "product_name": "Đèn A90M3", "source_row": 41},
        ])
        missing = EdgeAutomation._missing_purchase_products_from_errors(
            "Dòng 40: Mã hàng không có trên hệ thống hoặc ngừng kinh doanh", specs
        )
        self.assertEqual([item["productCode"] for item in missing], ["A90M2-10W-DM"])

    def test_refuses_ambiguous_row_mapping_and_normalizes_barcode(self):
        specs = EdgeAutomation._normalize_purchase_product_specs([
            {"product_code": "SOURCE-3", "product_name": "Nguồn", "source_row": 3},
            {"product_code": "GENERATED-2", "product_name": "Dòng tạo", "source_row": 99},
        ])
        missing = EdgeAutomation._missing_purchase_products_from_errors(
            "Dòng 3: Mã hàng không có trên hệ thống", specs
        )
        self.assertEqual(missing, [])
        self.assertEqual(EdgeAutomation._purchase_product_barcode("A90M2-10W-DM"), "A90M210WDM")


class EdgeAutomationSopokaExportTests(unittest.TestCase):
    """Regression coverage for Sopoka's fragile page-level Date override."""

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="orderbridge-sopoka-tests-")
        root = Path(self.temp.name)
        self.automation = EdgeAutomation(
            AppConfig(
                data_dir=root / "data",
                outputs_dir=root / "outputs",
                logs_dir=root / "logs",
                temp_dir=root / "temp",
            )
        )

    def tearDown(self) -> None:
        self._close_test_log_handler()
        self.temp.cleanup()

    def _close_test_log_handler(self) -> None:
        root = Path(self.temp.name).resolve()
        for handler in list(self.automation._logger.handlers):
            filename = getattr(handler, "baseFilename", None)
            try:
                is_test_handler = filename is not None and Path(filename).resolve().is_relative_to(root)
            except (OSError, ValueError):
                is_test_handler = False
            if is_test_handler:
                handler.close()
                self.automation._logger.removeHandler(handler)

    def test_export_today_uses_native_sopoka_steps_not_injected_page_task(self):
        calls: list[tuple[str, object]] = []
        downloaded = Path(self.temp.name) / "ReportOrderSaleSummary_test.xls"
        driver = object()

        self.automation._navigate = lambda url: calls.append(("navigate", url)) or driver
        self.automation._assert_sopoka_login = lambda current: calls.append(("login", current))
        self.automation._download_snapshot = lambda: {"before": 1}
        self.automation._select_sopoka_today = lambda expected: calls.append(("today", expected))
        self.automation._click_sopoka_control = (
            lambda selector, label: calls.append(("click", selector))
        )
        self.automation._wait_for_sopoka_results = lambda: calls.append(("wait_results", None))
        self.automation._sopoka_office_export_urls = lambda: set()
        self.automation._wait_for_sopoka_download = (
            lambda before, started, **_kwargs: calls.append(("wait_download", before)) or downloaded
        )
        self.automation._verify_exported_report = (
            lambda path, expected: calls.append(("verify", path)) or {"order_count": 2}
        )
        self.automation._run_task = lambda *_args, **_kwargs: self.fail(
            "Sopoka export must not inject dmsExportPhaseTaskV2 into the page"
        )

        result = self.automation.export_today()

        self.assertTrue(result["ok"])
        self.assertEqual(result["source_path"], str(downloaded))
        self.assertEqual(result["order_count"], 2)
        self.assertEqual(
            [call for call in calls if call[0] == "click"],
            [("click", "button.timK"), ("click", "button.xuatEx")],
        )
        self.assertIn(("wait_results", None), calls)
        progress = self.automation.status()["progress"]
        self.assertFalse(progress["active"])
        self.assertEqual(progress["phase"], "completed")
        self.assertEqual(progress["current"], 2)
        self.assertEqual(progress["total"], 2)

    def test_legacy_sopoka_task_uses_performance_timer_not_date_now(self):
        source = self.automation._task_source("dmsExportPhaseTaskV2")
        self.assertIn("const monotonic = () => performance.now();", source)
        self.assertNotIn("Date.now()", source)

    def test_office_viewer_source_requires_the_expected_origins_and_filename(self):
        source = (
            "http://dms.sopoka.com.vn/exports/example/"
            "ReportOrderSaleSummary_1786182210928.xls"
        )
        viewer = "https://view.officeapps.live.com/op/view.aspx?src=" + quote(source, safe="")

        self.assertEqual(self.automation._office_viewer_source_url(viewer), source)
        self.assertIsNone(
            self.automation._office_viewer_source_url(
                viewer.replace("dms.sopoka.com.vn", "untrusted.example")
            )
        )
        self.assertIsNone(
            self.automation._office_viewer_source_url(
                viewer.replace("ReportOrderSaleSummary", "anything-else")
            )
        )

    def test_direct_office_source_download_writes_only_a_validated_report_file(self):
        source = (
            "http://dms.sopoka.com.vn/exports/example/"
            "ReportOrderSaleSummary_1786182210928.xls"
        )

        class DownloadResponse(io.BytesIO):
            headers = {"Content-Type": "application/vnd.ms-excel"}

            def geturl(self):
                return source

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                self.close()

        self.automation._sopoka_download_cookie_header = lambda _url: "session=opaque"
        payload = b"D0CF" + (b"x" * 1024)
        with patch(
            "automation_app.edge_automation.urlopen",
            return_value=DownloadResponse(payload),
        ) as opened:
            downloaded = self.automation._direct_sopoka_report_download(source)

        self.assertTrue(downloaded.is_file())
        self.assertEqual(downloaded.read_bytes(), payload)
        request = opened.call_args.args[0]
        self.assertEqual(request.full_url, source)
        self.assertEqual(request.get_header("Cookie"), "session=opaque")

    def test_wait_for_download_prefers_verified_office_source_before_native_helper(self):
        source = (
            "http://dms.sopoka.com.vn/exports/example/"
            "ReportOrderSaleSummary_1786182210928.xls"
        )
        downloaded = Path(self.temp.name) / "ReportOrderSaleSummary_smoke.xls"
        downloaded.write_bytes(b"test")
        self.automation._wait_for_report_download = lambda *_args, **_kwargs: (_ for _ in ()).throw(
            EdgeAutomationError("native browser download not ready")
        )
        self.automation._wait_for_sopoka_office_export_url = lambda *_args, **_kwargs: source
        self.automation._direct_sopoka_report_download = lambda actual: (
            self.assertEqual(actual, source) or downloaded.resolve()
        )

        with patch("automation_app.edge_automation.subprocess.run") as native_helper:
            actual = self.automation._wait_for_sopoka_download(
                {}, datetime.now(timezone.utc), existing_office_export_urls=set()
            )

        self.assertEqual(actual, downloaded.resolve())
        native_helper.assert_not_called()

    def test_wait_for_download_allows_slow_office_target_before_native_helper(self):
        """A delayed Office target must not fall through to a missing UIA button."""

        source = (
            "http://dms.sopoka.com.vn/exports/example/"
            "ReportOrderSaleSummary_1786182210928.xls"
        )
        downloaded = Path(self.temp.name) / "ReportOrderSaleSummary_slow.xls"
        downloaded.write_bytes(b"test")
        self.automation._wait_for_report_download = lambda *_args, **_kwargs: (_ for _ in ()).throw(
            EdgeAutomationError("native browser download not ready")
        )
        office_results = iter([None, source])
        self.automation._wait_for_sopoka_office_export_url = lambda *_args, **_kwargs: next(office_results)
        self.automation._direct_sopoka_report_download = lambda actual: (
            self.assertEqual(actual, source) or downloaded.resolve()
        )

        with patch("automation_app.edge_automation.subprocess.run") as native_helper:
            actual = self.automation._wait_for_sopoka_download(
                {}, datetime.now(timezone.utc), existing_office_export_urls=set()
            )

        self.assertEqual(actual, downloaded.resolve())
        native_helper.assert_not_called()

    def test_export_download_uses_one_physical_retry_only_when_no_target_appears(self):
        source = (
            "http://dms.sopoka.com.vn/exports/example/"
            "ReportOrderSaleSummary_1786183787208.xls"
        )
        downloaded = Path(self.temp.name) / "ReportOrderSaleSummary_retry.xls"
        downloaded.write_bytes(b"test")
        button = object()
        target_results = iter([None, source])
        waits: list[float] = []
        pointer_calls: list[tuple[object, str]] = []
        self.automation._wait_for_report_download = lambda *_args, timeout_seconds: (
            waits.append(timeout_seconds)
            or (_ for _ in ()).throw(EdgeAutomationError("browser did not auto-download"))
        )
        self.automation._wait_for_sopoka_office_export_url = (
            lambda *_args, **_kwargs: next(target_results)
        )
        self.automation._dispatch_sopoka_pointer_click = (
            lambda actual, label: pointer_calls.append((actual, label)) or True
        )
        self.automation._direct_sopoka_report_download = lambda actual: (
            self.assertEqual(actual, source) or downloaded.resolve()
        )

        with patch("automation_app.edge_automation.subprocess.run") as native_helper:
            actual = self.automation._wait_for_sopoka_download(
                {},
                datetime.now(timezone.utc),
                existing_office_export_urls=set(),
                export_button=button,
            )

        self.assertEqual(actual, downloaded.resolve())
        self.assertEqual(waits, [4, 4])
        self.assertEqual(pointer_calls, [(button, "nút Xuất Excel")])
        native_helper.assert_not_called()

    def test_result_signature_ignores_date_picker_tables(self):
        class Row:
            text = "order row"

            @staticmethod
            def is_displayed():
                return True

        class Driver:
            def __init__(self):
                self.selector = None

            def find_elements(self, _by, selector):
                self.selector = selector
                return [Row()]

        driver = Driver()
        self.automation._ensure_driver = lambda: driver

        self.assertEqual(self.automation._sopoka_result_signature(), ("order row",))
        self.assertEqual(driver.selector, "table.reportCustomer tbody tr")

    def test_date_picker_reuses_an_already_open_picker_without_toggling_it_closed(self):
        picker = object()
        self.automation._visible_sopoka_date_picker = lambda: picker
        self.automation._wait_for_visible_css = lambda *_args, **_kwargs: self.fail(
            "an open picker must not click #reportrange again"
        )

        self.assertIs(self.automation._open_sopoka_date_picker(), picker)

    def test_today_option_rejects_hidden_picker_entries(self):
        class Option:
            def __init__(self, text, visible):
                self.text = text
                self._visible = visible

            def is_displayed(self):
                return self._visible

        class Picker:
            def find_elements(self, _by, _selector):
                return [Option("Hôm nay", False), Option("Tuần này", True)]

        self.automation._visible_sopoka_date_picker = lambda: Picker()

        self.assertIsNone(self.automation._visible_sopoka_today_option())

    def test_today_option_matches_utf8_and_mojibake_labels(self):
        class Option:
            def __init__(self, text):
                self.text = text

            def is_displayed(self):
                return True

        class Picker:
            def find_elements(self, _by, _selector):
                return [Option("Tuần này"), Option("Hôm nay")]

        self.automation._visible_sopoka_date_picker = lambda: Picker()
        self.assertEqual(self.automation._element_text(self.automation._visible_sopoka_today_option()), "Hôm nay")

    def test_existing_sopoka_page_is_not_switched_back_to_kiot(self):
        """Selector helpers must preserve the page selected by ``_navigate``."""

        driver = SimpleNamespace(current_url=SOPKA_URL)
        self.automation._driver = driver
        self.automation._select_sales_target_direct = Mock()

        self.assertIs(self.automation._ensure_driver(), driver)
        self.automation._select_sales_target_direct.assert_not_called()

    def test_profile_process_fallback_uses_only_the_dedicated_edge_profile(self):
        completed = Mock(returncode=0, stdout="26448\n", stderr="")
        with patch("automation_app.edge_automation.subprocess.run", return_value=completed) as run:
            process_id = self.automation._edge_profile_process_id()

        self.assertEqual(process_id, 26448)
        command = run.call_args.args[0]
        options = run.call_args.kwargs
        self.assertIn("-EncodedCommand", command)
        self.assertEqual(
            options["env"]["ORDERBRIDGE_EDGE_PROFILE"],
            str(self.automation.config.edge_profile_dir.resolve()),
        )

    @staticmethod
    def _devtools_version_response(port: int, *, host: str = "127.0.0.1") -> MagicMock:
        response = MagicMock()
        socket_host = f"[{host}]" if host == "::1" else host
        payload = {
            "Browser": "Edg/151.0.4129.59",
            "webSocketDebuggerUrl": f"ws://{socket_host}:{port}/devtools/browser/test-id",
        }
        response.__enter__.return_value = io.BytesIO(json.dumps(payload).encode("utf-8"))
        response.__exit__.return_value = False
        return response

    def test_live_profile_debugger_prefers_devtools_active_port_before_listener_fallback(self):
        active_port = self.automation.config.edge_profile_dir / "DevToolsActivePort"
        active_port.write_text("51755\n", encoding="utf-8")
        self.automation._profile_loopback_listener_addresses = Mock(return_value=[("127.0.0.1", 60000)])

        with patch(
            "automation_app.edge_automation.urlopen",
            return_value=self._devtools_version_response(51755),
        ) as open_endpoint:
            address = self.automation._live_profile_debugger_address()

        self.assertEqual(address, "127.0.0.1:51755")
        self.automation._profile_loopback_listener_addresses.assert_not_called()
        open_endpoint.assert_called_once_with(
            "http://127.0.0.1:51755/json/version", timeout=0.75
        )

    def test_profile_listener_query_is_scoped_to_exact_profile_pid_and_loopback(self):
        completed = Mock(
            returncode=0,
            stdout="127.0.0.1|51755\n::1|51756\n0.0.0.0|51757\nnot-a-port\n",
            stderr="",
        )
        with patch("automation_app.edge_automation.subprocess.run", return_value=completed) as run:
            listeners = self.automation._profile_loopback_listener_addresses(26448)

        self.assertEqual(listeners, [("127.0.0.1", 51755), ("::1", 51756)])
        command = run.call_args.args[0]
        encoded = command[command.index("-EncodedCommand") + 1]
        script = base64.b64decode(encoded).decode("utf-16-le")
        self.assertIn("Get-NetTCPConnection -State Listen -OwningProcess $edgeProcessId", script)
        self.assertIn("$env:ORDERBRIDGE_EDGE_PID", script)
        self.assertIn("127.0.0.1", script)
        self.assertIn("::1", script)
        self.assertEqual(run.call_args.kwargs["env"]["ORDERBRIDGE_EDGE_PID"], "26448")

    def test_live_profile_debugger_uses_only_a_validated_profile_listener_when_file_is_missing(self):
        self.automation._edge_profile_process_id = lambda: 26448
        self.automation._profile_loopback_listener_addresses = lambda process_id: (
            self.assertEqual(process_id, 26448) or [("127.0.0.1", 49000), ("127.0.0.1", 51755)]
        )
        valid = self._devtools_version_response(51755)
        with patch(
            "automation_app.edge_automation.urlopen",
            side_effect=[OSError("not a DevTools endpoint"), valid],
        ) as open_endpoint:
            address = self.automation._live_profile_debugger_address()

        self.assertEqual(address, "127.0.0.1:51755")
        self.assertEqual(
            [call.args[0] for call in open_endpoint.call_args_list],
            [
                "http://127.0.0.1:49000/json/version",
                "http://127.0.0.1:51755/json/version",
            ],
        )

    def test_session_process_id_falls_back_for_debugger_attached_edge(self):
        self.automation._edge_profile_process_id = lambda: 26448

        self.assertEqual(self.automation._session_browser_process_id({}), 26448)
        self.assertEqual(
            self.automation._session_browser_process_id({"goog:processID": 99}),
            99,
        )

    def test_office_download_uses_profile_pid_when_attached_driver_has_no_capability(self):
        downloaded = Path(self.temp.name) / "ReportOrderSaleSummary_smoke.xls"
        downloaded.write_bytes(b"test")
        self.automation._wait_for_report_download = lambda *_args, **_kwargs: (_ for _ in ()).throw(
            EdgeAutomationError("wait for native download")
        )
        self.automation._ensure_driver = lambda: SimpleNamespace(capabilities={})
        self.automation._edge_profile_process_id = lambda: 26448
        self.automation._wait_for_sopoka_office_export_url = lambda *_args, **_kwargs: None
        completed = SimpleNamespace(
            returncode=0,
            stdout=f'{{"success":true,"path":"{downloaded.as_posix()}","clicked":true}}\n',
            stderr="",
        )

        with patch("automation_app.edge_automation.subprocess.run", return_value=completed) as run:
            actual = self.automation._wait_for_sopoka_download(
                {}, datetime.now(timezone.utc)
            )

        self.assertEqual(actual, downloaded.resolve())
        command = run.call_args.args[0]
        index = command.index("-BrowserProcessId")
        self.assertEqual(command[index + 1], "26448")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
