"""Direct Microsoft Edge automation for the Sopoka → KiotViet workflow.

The browser is controlled through Selenium and a dedicated, persistent Edge
profile.  No browser extension, remote debugger, or browser credentials are
read by this application.  The user signs in once in the visible Edge window;
the session remains in ``data/edge-profile`` for later manual runs.
"""

from __future__ import annotations

import base64
import ctypes
import json
import logging
import os
import re
import subprocess
import tempfile
import threading
import time
import unicodedata
from ctypes import wintypes
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable
from urllib.error import HTTPError
from urllib.parse import parse_qs, unquote, urlsplit
from urllib.request import Request, urlopen

try:  # websocket-client is used only for Edge's local DevTools endpoint.
    import websocket
except ImportError:  # pragma: no cover - the native Escape fallback remains available
    websocket = None

from selenium import webdriver
from selenium.common.exceptions import (
    JavascriptException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from .config import AppConfig
from .report_parser import SUPPORTED_REPORT_EXTENSIONS


# The orders list has a different export and does not contain the related-order
# product rows needed by the KiotViet import.  The sales summary is the one
# report whose column H is ``Mã đặt đơn hàng liên quan``.
SOPKA_URL = "http://dms.sopoka.com.vn/reports/summaryOrderSales"
KIOT_URL = "https://nongsanlamhong.kiotviet.vn/man/#/DashBoard"
KIOT_ORDERS_URL = "https://nongsanlamhong.kiotviet.vn/man/#/Orders"
# Do not enter the sales canvas by clicking an Angular link from the management
# SPA.  That document replacement can discard Selenium's async-script callback
# before it resolves.  A normal WebDriver navigation keeps the transition under
# Python's control instead.
KIOT_SALES_URL = "https://nongsanlamhong.kiotviet.vn/sale/#/"
KIOT_NEW_ORDER_URL = "https://nongsanlamhong.kiotviet.vn/sale/#/?cart=Order"
KIOT_PURCHASE_URL = "https://nongsanlamhong.kiotviet.vn/man/#/PurchaseOrder/new"
KIOT_PURCHASE_LIST_URL = "https://nongsanlamhong.kiotviet.vn/man/#/PurchaseOrder"
KIOT_PRODUCTS_URL = "https://nongsanlamhong.kiotviet.vn/man/#/Products"


class EdgeAutomationError(RuntimeError):
    """A user-facing direct-browser automation failure."""


class EdgeLoginRequired(EdgeAutomationError):
    """Raised when the dedicated automation profile has no active login."""


class EdgeAutomation:
    """Own one visible, dedicated Edge session and execute safe browser actions."""

    def __init__(self, config: AppConfig | None = None) -> None:
        self.config = config or AppConfig()
        self.config.ensure_directories()
        self._lock = threading.RLock()
        # A long browser action owns ``_lock``.  Progress polling must remain
        # responsive while an import/final-order action is in flight, so this
        # tiny snapshot is protected independently and never reads WebDriver.
        self._progress_lock = threading.Lock()
        self._progress: dict[str, Any] = {
            "active": False,
            "phase": "idle",
            "message": "Sẵn sàng điều khiển Edge.",
            "current": 0,
            "total": 0,
            "batch": None,
            "orderCode": None,
            "updatedAt": datetime.now(timezone.utc).isoformat(),
        }
        self._last_known_url: str | None = None
        # KiotViet is a single-page app.  Keep the last shell URL that was
        # observed as interactive so a freshly navigated/loading sales screen
        # gets a short, native-WebDriver settle wait before we inject an
        # inspection task into it.
        self._last_kiot_shell_url: str | None = None
        self._last_kiot_shell_signature: tuple[Any, ...] | None = None
        self._driver: webdriver.Edge | None = None
        self._cancel_requested = False
        self._task_source_cache: dict[str, str] = {}
        self._logger = logging.getLogger("orderbridge.edge")
        if not self._logger.handlers:
            handler = logging.FileHandler(
                self.config.logs_dir / "edge-automation.log", encoding="utf-8"
            )
            handler.setFormatter(
                logging.Formatter("%(asctime)s %(levelname)s %(message)s")
            )
            self._logger.addHandler(handler)
            self._logger.setLevel(logging.INFO)
            self._logger.propagate = False

    # ------------------------------------------------------------------
    # Public API used by Flask
    # ------------------------------------------------------------------
    def status(self) -> dict[str, Any]:
        # Do not acquire ``_lock`` here: the dashboard polls status while a
        # single long Selenium action holds that lock.
        with self._progress_lock:
            progress = dict(self._progress)
        started = self._driver is not None
        return {
            "ok": True,
            "available": self.config.edge_driver_path.is_file()
            and self.config.edge_binary_path.is_file(),
            "started": started,
            "profile_path": str(self.config.edge_profile_dir),
            "download_path": str(self.config.edge_download_dir),
            # A remembered URL keeps this endpoint nonblocking while Selenium
            # is busy with the active page.
            "current_url": self._last_known_url if started else None,
            "progress": progress,
        }

    def execute(self, action: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        """Execute one safe, visible browser action synchronously."""

        normalized = str(action or "").strip().upper()
        data = payload or {}
        with self._lock:
            try:
                if normalized == "PING":
                    driver = self._ensure_driver()
                    return {"ok": True, "mode": "selenium", "current_url": driver.current_url}
                if normalized == "CANCEL_ACTIVE":
                    self._cancel_requested = True
                    self._cancel_sopoka_page()
                    return {"ok": True, "cancelled": True}
                if normalized == "SCHEDULE_PREFLIGHT":
                    return self.scheduled_preflight(
                        data.get("businessDate") or data.get("business_date"),
                        max_clock_skew_seconds=data.get("maxClockSkewSeconds"),
                    )
                if normalized == "CLOSE_SCHEDULED_TABS":
                    if data.get("afterSuccess") is not True:
                        raise EdgeAutomationError("Scheduled browser cleanup requires a completed workflow.")
                    return self.close_scheduled_tabs()
                if normalized == "EXPORT_TODAY":
                    self._cancel_requested = False
                    return self.export_today()
                if normalized == "EXPORT_DATE":
                    self._cancel_requested = False
                    return self.export_date(data.get("business_date") or data.get("businessDate"))
                if normalized == "INSPECT_SOPKA":
                    return self.inspect_sopoka()
                if normalized == "INSPECT_KIOT":
                    return self.inspect_kiot()
                if normalized == "INSPECT_KIOT_CURRENT":
                    self._assert_kiot_login()
                    current_driver = self._ensure_driver()
                    if not self._is_kiot_sales_url(current_driver.current_url):
                        # The branch/order-tab inspection is defined by the
                        # sales shell.  A previous purchase import can leave
                        # Edge on /man/#/PurchaseOrder; return to the sales
                        # shell before running this read-only probe.
                        self._navigate(KIOT_SALES_URL)
                    return {"ok": True, **self._kiot_status()}
                if normalized == "INSPECT_KIOT_CART_TABS":
                    self._assert_kiot_login()
                    current_driver = self._ensure_driver()
                    if not self._is_kiot_sales_url(current_driver.current_url):
                        self._navigate(KIOT_SALES_URL)
                    result = self._run_task("kiotInspectCartTabsTask", timeout_seconds=120)
                    self._task_result_or_raise(result)
                    return result
                if normalized == "INSPECT_KIOT_ORDERS":
                    return self.inspect_kiot_orders()
                if normalized == "PREPARE_KIOT":
                    return self.prepare_kiot()
                if normalized == "CLEAN_KIOT_TABS":
                    if data.get("confirmCloseData") is not True:
                        raise EdgeAutomationError("Cần xác nhận trước khi đóng các tab Đặt hàng cũ.")
                    return self.clean_kiot_tabs()
                if normalized == "IMPORT_BATCH":
                    return self.import_batch(data)
                if normalized == "FILL_NOTES":
                    return self.fill_notes(data)
                if normalized == "SUBMIT_BATCH":
                    return self.submit_batch(data)
                if normalized == "IMPORT_PURCHASE":
                    return self.import_purchase(data)
                raise EdgeAutomationError(f"Lệnh Edge không được hỗ trợ: {normalized or '(trống)'}." )
            except EdgeLoginRequired:
                self._set_progress(
                    active=False,
                    phase="blocked",
                    message="Edge đang chờ người dùng đăng nhập.",
                )
                raise
            except EdgeAutomationError:
                self._set_progress(
                    active=False,
                    phase="error",
                    message="Tác vụ Edge dừng để tránh thao tác không chắc chắn.",
                )
                self._save_failure_screenshot(normalized)
                raise
            except (JavascriptException, TimeoutException, WebDriverException) as exc:
                self._set_progress(
                    active=False,
                    phase="error",
                    message="Edge gặp lỗi khi đang thực hiện tác vụ.",
                )
                self._save_failure_screenshot(normalized)
                raise EdgeAutomationError(f"Edge dừng ở {normalized}: {exc}") from exc

    def export_today(self) -> dict[str, Any]:
        """Export today's sales-summary report through ordinary Edge input."""

        return self._export_for_date(self.config.today())

    def export_date(self, business_date: str | date) -> dict[str, Any]:
        """Export the selected one-day sales-summary report from Sopoka."""

        if isinstance(business_date, date):
            expected_date = business_date
        else:
            try:
                expected_date = datetime.strptime(str(business_date), "%Y-%m-%d").date()
            except (TypeError, ValueError) as exc:
                raise EdgeAutomationError(
                    "Ngày report phải có định dạng YYYY-MM-DD."
                ) from exc
        return self._export_for_date(expected_date)

    def _export_for_date(self, expected_date: date) -> dict[str, Any]:
        """Export one selected business date through ordinary Edge input.

        Sopoka replaces parts of ``Date.prototype`` in ``js/lib/date.js``.
        Executing an injected async JavaScript task in that page can therefore
        call its recursive ``Date.toString`` override and abort before the
        report is exported.  Keep this path deliberately free of page-script
        injection: Selenium only reads the DOM and sends normal mouse input.
        """

        driver = self._navigate(SOPKA_URL)
        self._assert_sopoka_login(driver)
        before = self._download_snapshot()
        # A viewer left open by a previous manual attempt must never be
        # mistaken for the file requested by this run.  Keep only its opaque
        # source URL in memory; it is neither logged nor sent to the UI.
        office_viewers_before = self._sopoka_office_export_urls()
        self._set_progress(
            active=True,
            phase="export_select_date",
            message="Đang chọn ngày nghiệp vụ hôm nay trên Sopoka.",
        )
        if expected_date == self.config.today():
            self._select_sopoka_today(expected_date)
        else:
            self._select_sopoka_date(expected_date)
        self._set_progress(
            active=True,
            phase="export_search",
            message="Đang tìm báo cáo tổng hợp bán hàng trên Sopoka.",
        )
        self._click_sopoka_control("button.timK", "nút Tìm kiếm")
        has_report_rows = self._wait_for_sopoka_results()
        if has_report_rows is False:
            # Sopoka treats an empty day as a valid result and shows a
            # visible "Thông báo: Không có dữ liệu" banner. Do not click
            # Xuất Excel when there is no report file to download.
            self._set_progress(
                active=False,
                phase="completed",
                message="Sopoka không có dữ liệu cho ngày đã chọn; không có đơn mới.",
                current=0,
                total=0,
            )
            return {
                "ok": True,
                "no_new": True,
                "sourcePath": "",
                "source_path": "",
                "business_date": expected_date.isoformat(),
                "order_count": 0,
                "downloaded_at": datetime.now(timezone.utc).isoformat(),
            }
        export_started_at = datetime.now(timezone.utc)
        self._set_progress(
            active=True,
            phase="export_download",
            message="Sopoka đã có kết quả; đang xuất và tải file Excel.",
        )
        export_button = self._click_sopoka_control("button.xuatEx", "nút Xuất Excel")
        downloaded = self._wait_for_sopoka_download(
            before,
            export_started_at,
            existing_office_export_urls=office_viewers_before,
            export_button=export_button,
        )
        report_summary = self._verify_exported_report(downloaded, expected_date)
        self._logger.info(
            "Sopoka sales summary export completed: %s date=%s orders=%s",
            downloaded,
            expected_date.isoformat(),
            report_summary["order_count"],
        )
        self._set_progress(
            active=False,
            phase="completed",
            message="Đã tải và kiểm tra report Sopoka; sẵn sàng đối chiếu SQLite.",
            current=report_summary["order_count"],
            total=report_summary["order_count"],
        )
        return {
            "ok": True,
            "sourcePath": str(downloaded),
            "source_path": str(downloaded),
            "business_date": expected_date.isoformat(),
            "order_count": report_summary["order_count"],
            "downloaded_at": datetime.now(timezone.utc).isoformat(),
        }

    def inspect_sopoka(self) -> dict[str, Any]:
        driver = self._navigate(SOPKA_URL)
        range_element = self._find_visible_css("#reportrange")
        search = self._find_visible_css("button.timK")
        export_button = self._find_visible_css("button.xuatEx")
        return {
            "ok": True,
            "url": driver.current_url,
            "title": driver.title,
            "reportReady": bool(range_element and search and export_button),
            "rangeText": self._element_text(range_element),
            "searchVisible": search is not None,
            "exportVisible": export_button is not None,
        }

    def inspect_kiot(self) -> dict[str, Any]:
        # Inspection needs the sales shell, not the management dashboard:
        # only the former exposes the branch/order-tab semantics we verify.
        self._navigate(KIOT_SALES_URL)
        self._assert_kiot_login()
        return {"ok": True, **self._kiot_status()}

    def scheduled_preflight(
        self,
        business_date: str | date | None,
        *,
        max_clock_skew_seconds: Any = None,
    ) -> dict[str, Any]:
        """Prove the scheduled browser session is safe before an export.

        The remote ``Date`` headers are read with ordinary HTTP HEAD requests,
        not injected page JavaScript.  Sopoka overrides browser Date helpers on
        some deployments, and only the header gives a meaningful comparison
        against the local UTC clock.
        """

        try:
            expected_date = business_date if isinstance(business_date, date) else datetime.strptime(
                str(business_date or ""), "%Y-%m-%d"
            ).date()
        except ValueError as exc:
            raise EdgeAutomationError("Scheduled workflow requires an explicit YYYY-MM-DD business date.") from exc
        if expected_date != self.config.today():
            raise EdgeAutomationError("Scheduled business date does not match Asia/Bangkok today.")
        try:
            skew_limit = int(max_clock_skew_seconds)
        except (TypeError, ValueError):
            skew_limit = self.config.scheduler_max_clock_skew_seconds
        if skew_limit < 1:
            raise EdgeAutomationError("Scheduled clock-skew limit must be positive.")

        sopoka = self.inspect_sopoka()
        if not sopoka.get("reportReady"):
            raise EdgeAutomationError("Sopoka report screen is not ready; scheduled export was not started.")
        kiot = self.inspect_kiot()
        order_tabs = kiot.get("orderTabs") if isinstance(kiot.get("orderTabs"), list) else []
        invoice_tabs = kiot.get("invoiceTabs") if isinstance(kiot.get("invoiceTabs"), list) else []
        blank_kiot = (
            bool(kiot.get("branchOk"))
            and bool(kiot.get("orderMode"))
            and not bool(kiot.get("invoiceMode"))
            and len(order_tabs) == 1
            and not invoice_tabs
            and not str(kiot.get("noteValue") or "").strip()
            and int(kiot.get("cartItemCount") or 0) == 0
        )
        if not blank_kiot:
            raise EdgeAutomationError(
                "KiotViet is not exactly one blank order tab; scheduled run will not alter existing carts."
            )

        now_utc = datetime.now(timezone.utc)
        clocks = {
            "sopoka": self._remote_clock(SOPKA_URL, now_utc, skew_limit),
            "kiotviet": self._remote_clock(KIOT_SALES_URL, now_utc, skew_limit),
        }
        bad = [name for name, clock in clocks.items() if clock["skew_seconds"] > skew_limit]
        if bad:
            details = ", ".join(f"{name}={clocks[name]['skew_seconds']}s" for name in bad)
            raise EdgeAutomationError(
                f"Local clock differs from remote server time beyond {skew_limit}s ({details})."
            )
        return {
            "ok": True,
            "business_date": expected_date.isoformat(),
            "max_clock_skew_seconds": skew_limit,
            "clocks": clocks,
            "sopoka": {"reportReady": True, "rangeText": sopoka.get("rangeText")},
            "kiot": {
                "branch": kiot.get("branch"),
                "orderTabs": order_tabs,
                "cartItemCount": kiot.get("cartItemCount"),
            },
        }

    @staticmethod
    def _remote_clock(url: str, now_utc: datetime, skew_limit: int) -> dict[str, Any]:
        """Read and validate the target server's standard HTTP Date header."""

        request = Request(
            url,
            method="HEAD",
            headers={"User-Agent": "OrderBridge-DN2/1.0", "Cache-Control": "no-cache"},
        )
        headers = None
        try:
            with urlopen(request, timeout=12) as response:
                headers = response.headers
        except HTTPError as exc:
            # A login redirect or 401 still carries the server response Date
            # header and does not change a user's authenticated Edge session.
            headers = exc.headers
        except Exception as exc:
            raise EdgeAutomationError(f"Could not read remote clock for {url}: {exc}") from exc
        raw_date = str(headers.get("Date") or "").strip() if headers is not None else ""
        if not raw_date:
            raise EdgeAutomationError(f"Remote server did not provide an HTTP Date header for {url}.")
        try:
            remote_utc = parsedate_to_datetime(raw_date).astimezone(timezone.utc)
        except (TypeError, ValueError) as exc:
            raise EdgeAutomationError(f"Remote server returned an invalid Date header for {url}.") from exc
        skew_seconds = round(abs((now_utc - remote_utc).total_seconds()), 3)
        return {
            "server_utc": remote_utc.isoformat(),
            "local_utc": now_utc.isoformat(),
            "skew_seconds": skew_seconds,
            "within_limit": skew_seconds <= skew_limit,
        }

    def close_scheduled_tabs(self) -> dict[str, Any]:
        """Close only Sopoka/KiotViet pages owned by a successful schedule.

        The dedicated Edge profile keeps its cookies.  We deliberately leave
        every other page alone and never run this cleanup for an error or an
        ``automation_attention`` run.
        """

        if self._driver is None:
            return {"ok": True, "closed": [], "remaining": [], "message": "No controlled Edge session to close."}
        closed: list[str] = []
        errors: list[str] = []
        try:
            handles = list(self._driver.window_handles)
        except WebDriverException:
            self._driver = None
            return {"ok": True, "closed": [], "remaining": [], "message": "Controlled Edge session was already closed."}
        for handle in handles:
            try:
                self._driver.switch_to.window(handle)
                current_url = str(self._driver.current_url or "")
                hostname = (urlsplit(current_url).hostname or "").casefold()
                if hostname not in {"dms.sopoka.com.vn", "nongsanlamhong.kiotviet.vn"}:
                    continue
                self._driver.close()
                closed.append(current_url)
            except WebDriverException as exc:
                errors.append(str(exc))
        try:
            remaining_handles = list(self._driver.window_handles)
            remaining: list[str] = []
            for handle in remaining_handles:
                self._driver.switch_to.window(handle)
                remaining.append(str(self._driver.current_url or ""))
            if remaining_handles:
                self._driver.switch_to.window(remaining_handles[0])
                self._last_known_url = self._safe_current_url()
            else:
                self._driver = None
                self._last_known_url = None
        except WebDriverException:
            self._driver = None
            self._last_known_url = None
            remaining = []
        return {"ok": not errors, "closed": closed, "remaining": remaining, "errors": errors}

    def inspect_kiot_orders(self) -> dict[str, Any]:
        """Open KiotViet's order list and return a bounded, read-only view."""

        driver = self._navigate(KIOT_ORDERS_URL)
        self._assert_kiot_login()
        body = driver.find_element(By.TAG_NAME, "body").text
        return {
            "ok": True,
            "url": driver.current_url,
            "title": driver.title,
            "body_excerpt": body[:6000],
        }

    def prepare_kiot(self) -> dict[str, Any]:
        existing_driver = self._ensure_driver()
        # Attach sessions may initially point at an unrelated New Tab even
        # though a hydrated KiotViet sales window is already open.  Select the
        # best existing sales window before deciding whether a route change is
        # needed; otherwise we would reload the live carts and lose their
        # in-memory note state.
        self._select_existing_domain_window(existing_driver, KIOT_SALES_URL)
        current_url = existing_driver.current_url
        if self._is_kiot_sales_url(current_url):
            self._assert_kiot_login()
            live_state = self._kiot_status()
            if (
                live_state.get("branchOk")
                and live_state.get("orderMode")
                and not live_state.get("invoiceMode")
                and live_state.get("orderTabs")
                and live_state.get("noteAvailable")
            ):
                return {"ok": True, **live_state}

        # Always request a fresh order canvas before a new batch.  If the
        # visible KiotViet window was left in a stale/blank sales shell, a
        # same-host navigation to the base sales URL keeps that shell alive
        # and the later cleanup never gets a chance to run.  The dedicated
        # ``?cart=Order`` route creates the numbered order tab that both the
        # readiness probe and cleanup command require.
        #
        # This branch is reached only when there are no confirmed live carts;
        # the early return above preserves a genuine in-progress batch.
        target_route = KIOT_NEW_ORDER_URL
        self._logger.info("KiotViet prepare: entering native sales route %s from %s", target_route, current_url)
        driver = self._navigate(target_route)
        self._assert_kiot_login()
        initial = self._wait_until(
            lambda: self._kiot_status(),
            lambda state: bool(state.get("orderTabs") or state.get("invoiceTabs")),
            "t\u1ea3i chi nh\u00e1nh KiotViet",
            timeout_seconds=45,
        )
        if initial.get("invoiceMode"):
            conversion = self._run_task("kiotConvertActiveInvoiceTask")
            self._task_result_or_raise(conversion)
            time.sleep(0.35)
            converted = self._kiot_status()
            if converted.get("invoiceMode"):
                menu_result = self._run_task("kiotChooseConvertInvoiceMenuTask")
                self._task_result_or_raise(menu_result)
            converted = self._wait_until(
                lambda: self._kiot_status(),
                lambda state: bool(state.get("orderMode")) and not bool(state.get("invoiceMode")),
                "chuyển Hóa đơn sang Đặt hàng",
                timeout_seconds=7,
            )
            initial = converted

        if (
            initial.get("branchOk")
            and initial.get("orderMode")
            and initial.get("orderTabs")
            and not initial.get("invoiceMode")
            and initial.get("noteAvailable")
        ):
            self._logger.info("KiotViet prepare: native sales route is ready without an async page task")
            return {"ok": True, **initial}

        branch_verified_on_dashboard = bool(initial.get("branchOk"))
        if not initial.get("orderMode"):
            # A direct retry is safer than returning to the management SPA and
            # asking an injected callback to cross into the sales application.
            self._logger.info("KiotViet prepare: retrying native sales order route after non-order state")
            driver = self._navigate(KIOT_NEW_ORDER_URL)
            self._assert_kiot_login()
            # KiotViet is an Angular screen: document.readyState can be
            # complete before its branch selector is rendered.  Do not run a
            # click task against that transient empty shell.
            sales_state = self._wait_until(
                lambda: self._kiot_status(),
                lambda state: bool(state.get("orderTabs") or state.get("invoiceTabs")),
                "t\u1ea3i giao di\u1ec7n KiotViet",
                timeout_seconds=45,
            )
            if (
                sales_state.get("orderMode")
                and sales_state.get("orderTabs")
                and not sales_state.get("invoiceTabs")
                and sales_state.get("noteAvailable")
            ):
                if branch_verified_on_dashboard and not sales_state.get("branch"):
                    sales_state = {
                        **sales_state,
                        "branch": self.config.expected_branch,
                        "branchOk": True,
                    }
                return {"ok": True, **sales_state}

        self._logger.info(
            "KiotViet prepare: starting in-page task url=%s branch_ok=%s order_mode=%s tabs=%s",
            self._safe_current_url(),
            branch_verified_on_dashboard,
            bool(initial.get("orderMode")),
            len(initial.get("orderTabs") or []),
        )
        prepared = self._run_task(
            "kiotPrepareTask", self.config.expected_branch, branch_verified_on_dashboard
        )
        self._task_result_or_raise(prepared)
        self._logger.info(
            "KiotViet prepare: in-page task completed needs_order_route=%s",
            bool(prepared.get("needsOrderRoute")),
        )
        if prepared.get("needsOrderRoute"):
            # This is the only fallback for an empty sales canvas.  It is a
            # normal Selenium navigation, not an async page-script click.
            driver = self._navigate(KIOT_NEW_ORDER_URL)
            self._assert_kiot_login()
        status = self._wait_until(
            lambda: self._kiot_status(),
            lambda state: (
                (bool(state.get("branchOk")) or (branch_verified_on_dashboard and not state.get("branch")))
                and bool(state.get("orderMode"))
                and not bool(state.get("invoiceMode"))
                and bool(state.get("orderTabs"))
                and bool(state.get("noteAvailable"))
            ),
            "màn hình Đặt hàng tại Đồng Nai 2 - ĐN2",
            timeout_seconds=12,
        )
        if branch_verified_on_dashboard and not status.get("branch"):
            status = {**status, "branch": self.config.expected_branch, "branchOk": True}
        return {"ok": True, **status}

    def clean_kiot_tabs(self) -> dict[str, Any]:
        current = self._kiot_status()
        if not current.get("orderMode") or current.get("invoiceMode") or not current.get("branchOk"):
            current = self.prepare_kiot()
        # Close each cart in one short page task, then accept the discard
        # confirmation in a separate native Selenium command.  Keeping both
        # interactions in one long injected script can prevent this KiotViet
        # build from painting its confirmation dialog until too late.
        self._dismiss_kiot_discard_dialog()
        closed = 0

        # A stopped run can leave exactly one tab with a saved note.  The old
        # cleanup loop treated that as success and later tried to import on top
        # of it. First make a blank cart through KiotViet's explicit
        # ``Thêm mới`` control, then convert that newly-created blank invoice
        # to an order tab. Only after the blank order tab exists do we discard
        # old tabs. This never requires closing the last remaining cart.
        # Always establish a fresh blank foundation before deleting a group of
        # old carts.  Previously this branch only ran when exactly one stale
        # tab existed; with multiple tabs the cleanup preserved the active
        # stale cart (often the last tab) and then failed before import.
        # Keeping a newly-created blank tab active makes the subsequent
        # right-to-left deletion deterministic and prevents notes from being
        # written into an old cart.
        active_label = str(current.get("activeTab") or "").strip()
        active_is_blank = (
            active_label in current.get("orderTabs", [])
            and not str(current.get("noteValue") or "").strip()
            and int(current.get("cartItemCount") or 0) == 0
        )
        # A blank *order* tab is all the cleanup foundation needs.  Invoice
        # tabs are removed by the loop below; creating another cart merely
        # because one exists can hit KiotViet's fragile `Thêm mới` control.
        needs_blank_foundation = not active_is_blank
        # Prefer an already-existing blank order tab at every tab count, not
        # just at the 20-cart ceiling.  An interrupted import commonly leaves
        # one such tab alongside data-filled carts.  Reusing it avoids a
        # needless new-tab click and lets cleanup preserve it as the sole
        # import canvas.
        if needs_blank_foundation:
            blank_label = self._find_existing_blank_order_tab()
            if blank_label:
                self._set_progress(
                    active=True,
                    phase="cleaning_tabs",
                    message="Đang dùng tab Đặt hàng trống có sẵn trước khi dọn tab cũ.",
                )
                self._activate_kiot_tab_native(blank_label)
                current = self._wait_until(
                    lambda: self._kiot_status(),
                    lambda state: (
                        str(state.get("activeTab") or "").strip() == blank_label
                        and not str(state.get("noteValue") or "").strip()
                        and int(state.get("cartItemCount") or 0) == 0
                    ),
                    "kích hoạt tab Đặt hàng trống hiện có",
                    timeout_seconds=12,
                )
                active_is_blank = True
                needs_blank_foundation = False
        if needs_blank_foundation:
            self._set_progress(
                active=True,
                phase="cleaning_tabs",
                message="Đang tạo tab Đặt hàng trống trước khi dọn tab cũ.",
            )
            # KiotViet's Angular handler ignores synthetic DOM clicks on the
            # plus control in some Edge builds.  Use a native WebDriver
            # pointer click for the explicit tab-strip control so the same
            # event path as a user click creates the new cart.  The JS task is
            # retained as a compatibility fallback for older shells where the
            # native selector is temporarily not rendered.
            try:
                created = self._create_blank_order_tab_native()
            except EdgeAutomationError:
                created = self._run_task("kiotCreateBlankOrderTabTask", timeout_seconds=14)
            self._task_result_or_raise(created)
            before_label = str(current["orderTabs"][0])
            before_order_count = len(current.get("orderTabs", []))
            before_invoice_count = len(current.get("invoiceTabs", []))
            current = self._wait_until(
                lambda: self._kiot_status(),
                lambda state: (
                    before_label in state.get("orderTabs", [])
                    and (
                        len(state.get("orderTabs", [])) > before_order_count
                        or len(state.get("invoiceTabs", [])) > before_invoice_count
                    )
                ),
                "tạo tab trống",
                timeout_seconds=10,
            )
            # KiotViet versions differ: in the sales shell, ``newPane()``
            # may create an order tab directly; older builds create an
            # invoice tab that must be converted.  Only run the conversion
            # when the observed state proves that an invoice was created.
            if len(current.get("invoiceTabs", [])) > before_invoice_count:
                converted = self._run_task("kiotConvertActiveInvoiceTask", timeout_seconds=14)
                self._task_result_or_raise(converted)
                current = self._wait_until(
                    lambda: self._kiot_status(),
                    lambda state: len(state.get("orderTabs", [])) > before_order_count and not bool(state.get("invoiceMode")),
                    "chuyển tab trống sang Đặt hàng",
                    timeout_seconds=10,
                )

        for _ in range(30):
            status = self._kiot_status()
            invoices = list(status.get("invoiceTabs", []))
            orders = list(status.get("orderTabs", []))
            if not invoices and len(orders) == 1:
                break
            # Never close the active cart first: it can contain the stopped
            # order's note/data.  Remove inactive carts from right to left;
            # only then leave one active blank foundation tab for import.
            active = str(status.get("activeTab") or "").strip()
            candidates = invoices + orders
            tab_label = next((label for label in reversed(candidates) if label != active), None)
            if tab_label is None:
                tab_label = candidates[-1]
            before_count = len(invoices) + len(orders)
            request = self._run_task("kiotRequestCloseTabTask", tab_label, timeout_seconds=12)
            self._task_result_or_raise(request)
            status = self._wait_for_kiot_tab_close(tab_label, before_count)
            closed += 1
        else:
            raise EdgeAutomationError("KiotViet tab cleanup exceeded the safety limit.")
        if len(status.get("orderTabs", [])) != 1 or status.get("invoiceTabs"):
            raise EdgeAutomationError(
                "KiotViet không còn đúng một tab Đặt hàng trống sau khi dọn."
            )
        if (
            str(status.get("noteValue") or "").strip()
            or int(status.get("cartItemCount") or 0) > 0
        ):
            raise EdgeAutomationError(
                "KiotViet chỉ còn một tab nhưng tab đó vẫn có ghi chú cũ; "
                "tool không import chồng lên đơn đang dở."
            )
        if not status.get("branch") and current.get("branchOk"):
            status = {**status, "branch": self.config.expected_branch, "branchOk": True}
        return {
            "ok": True,
            "closed": closed,
            "orderTabs": status["orderTabs"],
            "invoiceTabs": status["invoiceTabs"],
        }

    def _find_existing_blank_order_tab(self) -> str | None:
        """Read-only scan for an empty order cart that can be a cleanup anchor."""

        result = self._run_task("kiotInspectCartTabsTask", timeout_seconds=120)
        self._task_result_or_raise(result)
        for item in result.get("tabs", []):
            if not isinstance(item, dict):
                continue
            label = str(item.get("tabLabel") or "").strip()
            if (
                label.startswith("Đặt hàng")
                and int(item.get("cartItemCount") or 0) == 0
                and not str(item.get("noteValue") or "").strip()
            ):
                return label
        return None

    def _activate_kiot_tab_native(self, tab_label: str) -> None:
        """Activate one numbered cart with a native WebDriver pointer click."""

        driver = self._ensure_driver()
        escaped = str(tab_label).replace("'", "\\'")
        xpath = (
            "//a[contains(concat(' ', normalize-space(@class), ' '), ' nav-link ')][.//span[normalize-space()="
            f"'{escaped}' or normalize-space(text())='{escaped}']]"
        )
        candidates = driver.find_elements(By.XPATH, xpath)
        for element in candidates:
            try:
                # Selenium reports an offscreen tab as not displayed even
                # though its parent overflow lane contains it.  Dimensions
                # prove it is a real tab; scroll the lane into view before the
                # native pointer click.
                if float(element.rect.get("width", 0)) <= 0 or float(element.rect.get("height", 0)) <= 0:
                    continue
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'nearest',inline:'center'});",
                    element,
                )
                ActionChains(driver).move_to_element(element).click().perform()
                return
            except (WebDriverException, JavascriptException) as error:
                self._logger.warning("Native KiotViet tab activation failed for %s: %s", tab_label, error)
        raise EdgeAutomationError(f"Không thể kích hoạt tab KiotViet {tab_label} bằng thao tác native.")

    def _create_blank_order_tab_native(self) -> dict[str, Any]:
        """Create a blank cart through KiotViet's visible tab-strip plus button."""

        driver = self._ensure_driver()
        selectors = (
            'a.btn-add-tabs[title="Thêm mới"]',
            'button.btn-add-tabs[title="Thêm mới"]',
            'a.btn-add-tabs[aria-label*="Thêm mới"]',
            'button.btn-add-tabs[aria-label*="Thêm mới"]',
        )
        for selector in selectors:
            try:
                candidates = driver.find_elements(By.CSS_SELECTOR, selector)
            except WebDriverException:
                continue
            for element in candidates:
                try:
                    if not element.is_displayed() or not element.is_enabled():
                        continue
                    rect = element.rect
                    if float(rect.get("y", 9999)) >= 180 or float(rect.get("width", 0)) <= 0:
                        continue
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block:'nearest',inline:'center'});",
                        element,
                    )
                    ActionChains(driver).move_to_element(element).click().perform()
                    return {
                        "ok": True,
                        "action": "create_blank_order_tab_native",
                        "selector": selector,
                    }
                except (WebDriverException, JavascriptException) as error:
                    self._logger.warning("Native KiotViet new-tab click failed; trying next candidate: %s", error)
        raise EdgeAutomationError("Không tìm thấy nút Thêm mới hiển thị để tạo tab Đặt hàng trống.")

    def _wait_for_kiot_tab_close(self, tab_label: str, before_count: int) -> dict[str, Any]:
        """Wait for this exact tab's modal, confirm it, and verify tab removal."""

        deadline = time.monotonic() + 18
        last: dict[str, Any] | None = None
        while time.monotonic() < deadline:
            self._dismiss_kiot_discard_dialog(tab_label)
            try:
                last = self._kiot_status()
            except (EdgeAutomationError, TimeoutException, WebDriverException):
                time.sleep(0.3)
                continue
            current_count = len(last.get("orderTabs", [])) + len(last.get("invoiceTabs", []))
            # A cart label can disappear simply because KiotViet converted
            # ``Đặt hàng N`` into ``Hóa đơn N`` after the left conversion
            # icon was hit.  That is not a discard.  Require the total count
            # to decrease before treating cleanup as successful.
            if current_count < before_count:
                return last
            time.sleep(0.3)
        raise EdgeAutomationError(
            f"KiotViet did not close {tab_label}; stopping cleanup before import."
        )

    def _dismiss_kiot_discard_dialog(self, tab_label: str | None = None) -> bool:
        """Click `Đồng ý` only inside the matching KiotViet discard modal."""

        driver = self._ensure_driver()
        title = f"Đóng {tab_label}" if tab_label else "Đóng Đặt hàng"
        # This KiotViet version renders confirmations as Kendo windows rather
        # than ARIA dialogs.  Prefer its stable classes and native button.
        for modal in driver.find_elements(By.CSS_SELECTOR, ".k-widget.k-window"):
            try:
                title_element = modal.find_element(By.CSS_SELECTOR, ".k-window-title")
                window_title = title_element.text or str(
                    title_element.get_attribute("textContent") or ""
                ).strip()
                if title not in window_title:
                    continue
                # KiotViet's Kendo animation can remain at scale(.01) when
                # the long-running WebDriver cleanup loop owns the page.  The
                # confirmation is still the exact matching modal; finish its
                # visual transition before using a native click.
                if not modal.is_displayed():
                    tiny = driver.execute_script(
                        "const r=arguments[0].getBoundingClientRect();"
                        "return r.width > 0 && r.width < 100 && getComputedStyle(arguments[0]).display !== 'none';",
                        modal,
                    )
                    if tiny:
                        driver.execute_script(
                            "arguments[0].style.setProperty('transform','none','important');",
                            modal,
                        )
                if not modal.is_displayed():
                    # If Selenium still rejects the element because Kendo's
                    # animation state is stale, click only the exact danger
                    # button inside the title-matched modal via WebDriver JS.
                    clicked = driver.execute_script(
                        "const b=arguments[0].querySelector('button.btn-danger:not([disabled])');"
                        "if(!b)return false; b.click(); return true;",
                        modal,
                    )
                    if clicked:
                        self._logger.info("Confirmed KiotViet Kendo discard dialog for %s via exact modal fallback.", tab_label or "existing tab")
                        return True
                    continue
                buttons = modal.find_elements(By.CSS_SELECTOR, "button.btn-danger:not([disabled])")
                confirm = next(
                    (
                        button
                        for button in buttons
                        if button.is_displayed()
                        and button.is_enabled()
                        and button.text.strip() == "\u0110\u1ed3ng \u00fd"
                    ),
                    None,
                )
                if confirm is None:
                    clicked = driver.execute_script(
                        "const b=arguments[0].querySelector('button.btn-danger:not([disabled])');"
                        "if(!b)return false; b.click(); return true;",
                        modal,
                    )
                    if clicked:
                        self._logger.info("Confirmed KiotViet Kendo discard dialog for %s via exact modal fallback.", tab_label or "existing tab")
                        return True
                    continue
                ActionChains(driver).move_to_element(confirm).pause(0.05).click(confirm).perform()
                self._logger.info("Confirmed KiotViet Kendo discard dialog for %s.", tab_label or "existing tab")
                return True
            except WebDriverException:
                continue
        modal_xpath = (
            "//*[(@role='dialog' or contains(@class,'modal') or contains(@class,'dialog') or contains(@class,'k-window')) "
            f"and contains(normalize-space(.), \"{title}\")]"
        )
        for modal in driver.find_elements(By.XPATH, modal_xpath):
            try:
                if not modal.is_displayed():
                    tiny = driver.execute_script(
                        "const r=arguments[0].getBoundingClientRect();"
                        "return r.width > 0 && r.width < 100 && getComputedStyle(arguments[0]).display !== 'none';",
                        modal,
                    )
                    if tiny:
                        driver.execute_script(
                            "arguments[0].style.setProperty('transform','none','important');",
                            modal,
                        )
                if not modal.is_displayed():
                    clicked = driver.execute_script(
                        "const b=arguments[0].querySelector('button.btn-danger:not([disabled])');"
                        "if(!b)return false; b.click(); return true;",
                        modal,
                    )
                    if clicked:
                        self._logger.info("Confirmed KiotViet Kendo discard dialog for %s via exact modal fallback.", tab_label or "existing tab")
                        return True
                    continue
                buttons = modal.find_elements(
                    By.XPATH,
                    ".//*[self::button or @role='button' or @ng-click or @onclick]"
                    "[not(@disabled) and @aria-disabled!='true' and normalize-space(.)='Đồng ý']",
                )
                confirm = next(
                    (button for button in buttons if button.is_displayed() and button.is_enabled()),
                    None,
                )
                if confirm is None:
                    clicked = driver.execute_script(
                        "const b=arguments[0].querySelector('button.btn-danger:not([disabled])');"
                        "if(!b)return false; b.click(); return true;",
                        modal,
                    )
                    if clicked:
                        self._logger.info("Confirmed KiotViet Kendo discard dialog for %s via exact modal fallback.", tab_label or "existing tab")
                        return True
                    continue
                ActionChains(driver).move_to_element(confirm).pause(0.05).click(confirm).perform()
                self._logger.info("Confirmed KiotViet discard dialog for %s.", tab_label or "existing tab")
                return True
            except WebDriverException:
                continue
        return False

    def import_batch(self, payload: dict[str, Any]) -> dict[str, Any]:
        file_path = Path(str(payload.get("filePath") or payload.get("file_path") or "")).resolve()
        orders = payload.get("orders")
        if not file_path.is_file():
            raise EdgeAutomationError(f"Không tìm thấy file import: {file_path}")
        if not isinstance(orders, list) or not 1 <= len(orders) <= 19:
            raise EdgeAutomationError("Lô import phải có từ 1 đến 19 đơn.")

        orders = self._normalize_note_expectations(orders)
        batch_label = self._batch_label(payload)
        self._set_progress(
            active=True,
            phase="preparing_kiot",
            message="Đang kiểm tra chi nhánh và tab Đặt hàng trên KiotViet.",
            current=0,
            total=len(orders),
            batch=batch_label,
            order_code=None,
        )
        status = self.prepare_kiot()
        if len(status.get("orderTabs", [])) != 1 or status.get("invoiceTabs"):
            status = self.clean_kiot_tabs()
        if len(status.get("orderTabs", [])) != 1 or status.get("invoiceTabs"):
            raise EdgeAutomationError("Không thể tạo đúng 1 tab Đặt hàng trống trước import.")

        self._set_progress(
            active=True,
            phase="uploading",
            message=f"Đang upload file import gồm {len(orders)} đơn lên KiotViet.",
            current=0,
            total=len(orders),
            batch=batch_label,
            order_code=None,
        )
        # Opening the import panel can briefly outlive the normal page-task
        # timeout while KiotViet repaints the sales shell.  Keep this bounded
        # but long enough to cover that transition; a timeout still fails
        # closed and never uploads a second file.
        opened = self._run_task("kiotOpenImportTask", timeout_seconds=45)
        self._task_result_or_raise(opened)
        upload = WebDriverWait(self._ensure_driver(), 15).until(
            lambda current: current.find_element(By.CSS_SELECTOR, 'input[type="file"]')
        )
        upload.send_keys(str(file_path))
        # Importing 19 sheets is asynchronous in KiotViet.  On a cold SPA it
        # may take well over the old 45-second in-page deadline, which made
        # the dashboard report a script timeout even while carts were still
        # being created.  Use a bounded 3-minute wait and let the browser task
        # keep polling until the exact expected tab count is present.
        import_wait_seconds = max(120, min(240, 12 * len(orders)))
        try:
            imported = self._run_task(
                "kiotWaitImportTask",
                len(orders),
                len(status["orderTabs"]),
                timeout_seconds=import_wait_seconds,
            )
        except TimeoutException:
            # A Selenium script timeout is ambiguous: KiotViet may have
            # finished creating the carts while the transport callback was
            # lost during an Angular repaint.  Re-read the shell before
            # deciding anything.  If the exact tab count is present, accept
            # the import as recovered and never send the workbook again.
            recovered = self._kiot_status()
            expected_count = len(status["orderTabs"]) + len(orders)
            if (
                not recovered.get("invoiceTabs")
                and len(recovered.get("orderTabs", [])) == expected_count
            ):
                self._logger.warning(
                    "KiotViet import callback timed out after carts were created; "
                    "continuing without re-upload (tabs=%s).",
                    expected_count,
                )
                imported = {
                    "ok": True,
                    "recoveredAfterTimeout": True,
                    "tabLabels": list(recovered.get("orderTabs", [])),
                }
            else:
                raise
        self._task_result_or_raise(imported)
        self._set_progress(
            active=True,
            phase="filling_notes",
            message="Đang điền và đọc lại ghi chú Mã đặt đơn hàng cho từng tab.",
            current=0,
            total=len(orders),
            batch=batch_label,
            order_code=None,
        )
        # Verifying/mapping 19 visible carts can take several short Angular
        # renders per tab. The old 85s Selenium deadline expired after import
        # had already succeeded, leaving all tabs present but notes empty.
        notes_wait_seconds = max(180, min(360, 12 * len(orders)))
        try:
            notes = self._run_task("kiotFillNotesTask", orders, timeout_seconds=notes_wait_seconds)
        except TimeoutException:
            # Filling notes is idempotent.  Retry only this read/write mapping
            # task after confirming the imported carts still exist; never
            # retry the upload or a final-order click from this path.
            recovered = self._kiot_status()
            expected_count = len(status["orderTabs"]) + len(orders)
            if recovered.get("invoiceTabs") or len(recovered.get("orderTabs", [])) < expected_count:
                raise
            retry_timeout = max(240, min(480, notes_wait_seconds + 120))
            self._logger.warning(
                "KiotViet note-mapping callback timed out with carts present; "
                "retrying notes only (no upload/order click)."
            )
            notes = self._run_task("kiotFillNotesTask", orders, timeout_seconds=retry_timeout)
        self._task_result_or_raise(notes)
        if not notes.get("ok") or not isinstance(notes.get("verified"), list):
            raise EdgeAutomationError("KiotViet không trả về xác minh ghi chú hợp lệ.")
        if len(notes["verified"]) != len(orders):
            raise EdgeAutomationError("KiotViet did not verify every order note in this batch.")
        # The imported worksheet tab number is not stable when KiotViet keeps
        # an older blank cart.  The page task verifies each SKU/customer and
        # returns the actual label; use that verified mapping for submission.
        verified_by_code = {
            str(item.get("orderCode") or item.get("order_code") or "").strip(): str(
                item.get("tabLabel") or item.get("tab_label") or ""
            ).strip()
            for item in notes["verified"]
            if isinstance(item, dict)
        }
        submit_orders = [
            {
                **order,
                "expectedTabLabel": verified_by_code.get(
                    str(order.get("orderCode") or order.get("order_code") or "").strip(),
                    order.get("expectedTabLabel") or order.get("expected_tab_label"),
                ),
            }
            for order in orders
        ]
        if payload.get("autoSubmit") is True or payload.get("auto_submit") is True:
            submitted = self._submit_verified_orders(submit_orders, batch_label=batch_label)
            if submitted.get("ok") is not True:
                return {
                    "ok": False,
                    "imported": imported,
                    "notes": notes,
                    "submitted": submitted,
                    "finalOrderClicked": False,
                    "message": "Tool đã dừng sau một đơn chưa xác minh được; các đơn còn lại không bị bấm.",
                }
            self._logger.info("KiotViet final order completion: %s", file_path)
            return {
                "ok": True,
                "imported": imported,
                "notes": notes,
                "submitted": submitted,
                "finalOrderClicked": True,
            }
        self._logger.info("KiotViet import/note completion: %s", file_path)
        self._set_progress(
            active=False,
            phase="review_ready",
            message="Đã import và xác minh ghi chú; sẵn sàng đặt hàng.",
            current=len(orders),
            total=len(orders),
            batch=batch_label,
            order_code=None,
        )
        return {"ok": True, "imported": imported, "notes": notes, "finalOrderClicked": False}

    def fill_notes(self, payload: dict[str, Any]) -> dict[str, Any]:
        orders = payload.get("orders")
        if not isinstance(orders, list) or not 1 <= len(orders) <= 19:
            raise EdgeAutomationError("Cần từ 1 đến 19 đơn để điền ghi chú.")
        orders = self._normalize_note_expectations(orders)
        status = self.prepare_kiot()
        if status.get("invoiceTabs") or len(status.get("orderTabs", [])) < len(orders) + 1:
            raise EdgeAutomationError("Các tab KiotViet không khớp lô import an toàn để điền ghi chú.")
        notes_wait_seconds = max(180, min(360, 12 * len(orders)))
        notes = self._run_task("kiotFillNotesTask", orders, timeout_seconds=notes_wait_seconds)
        self._task_result_or_raise(notes)
        return {"ok": True, "notes": notes, "finalOrderClicked": False}

    @staticmethod
    def _normalize_note_expectations(orders: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Normalize DB/UI snake_case and queue camelCase fields for the page task."""

        normalized: list[dict[str, Any]] = []
        seen_order_codes: set[str] = set()
        for order in orders:
            if not isinstance(order, dict):
                raise EdgeAutomationError("Một đơn ghi chú không hợp lệ.")
            item = dict(order)
            item["orderCode"] = str(item.get("orderCode") or item.get("order_code") or "").strip()
            item["customerCode"] = str(item.get("customerCode") or item.get("customer_code") or "").strip()
            item["customerName"] = str(item.get("customerName") or item.get("customer_name") or "").strip()
            item["firstSku"] = str(item.get("firstSku") or item.get("first_sku") or "").strip()
            raw_codes = item.get("productCodes") or item.get("product_codes") or []
            if isinstance(raw_codes, (list, tuple)):
                item["productCodes"] = [str(code or "").strip() for code in raw_codes if str(code or "").strip()]
            else:
                item["productCodes"] = [item["firstSku"]] if item["firstSku"] else []
            raw_names = item.get("productNames") or item.get("product_names") or []
            if isinstance(raw_names, (list, tuple)):
                item["productNames"] = [str(name or "").strip() for name in raw_names if str(name or "").strip()]
            else:
                item["productNames"] = []
            raw_lines = item.get("productLines") or item.get("product_lines") or []
            product_lines: list[dict[str, Any]] = []
            if isinstance(raw_lines, (list, tuple)):
                for raw_line in raw_lines:
                    if not isinstance(raw_line, dict):
                        continue
                    product_code = str(
                        raw_line.get("productCode") or raw_line.get("product_code") or ""
                    ).strip()
                    product_name = str(
                        raw_line.get("productName") or raw_line.get("product_name") or ""
                    ).strip()
                    if not product_code and not product_name:
                        continue
                    product_lines.append(
                        {
                            "productCode": product_code,
                            "productName": product_name,
                            "quantity": raw_line.get("quantity"),
                            "netAmount": raw_line.get("netAmount", raw_line.get("net_amount")),
                        }
                    )
            item["productLines"] = product_lines
            item["paymentAmount"] = item.get("paymentAmount", item.get("payment_amount"))
            item["expectedTabLabel"] = str(
                item.get("expectedTabLabel") or item.get("expected_tab_label") or ""
            ).strip()
            if not item["orderCode"]:
                raise EdgeAutomationError("Mã đặt đơn hàng bị thiếu khi điền ghi chú.")
            if item["orderCode"] in seen_order_codes:
                raise EdgeAutomationError(
                    "Mã đặt hàng liên quan bị trùng trong cùng lô; không thể gán ghi chú an toàn."
                )
            seen_order_codes.add(item["orderCode"])
            normalized.append(item)
        return normalized

    @staticmethod
    def _kiot_text_key(value: Any) -> str:
        """Compare KiotViet labels despite accents or a legacy page encoding."""

        candidates = [str(value or "")]
        try:
            candidates.append(candidates[0].encode("latin1").decode("utf-8"))
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
        keys: list[str] = []
        for candidate in candidates:
            text = unicodedata.normalize("NFKD", candidate)
            text = "".join(char for char in text if not unicodedata.combining(char))
            text = text.replace("Đ", "D").replace("đ", "d")
            keys.append(re.sub(r"\s+", " ", text).strip().casefold())
        return min(keys, key=lambda item: (sum(item.count(mark) for mark in ("Ã", "Ä", "Â", "â")), len(item)))

    @classmethod
    def _normalize_purchase_product_specs(cls, raw_items: Any) -> list[dict[str, Any]]:
        """Normalize product data returned by the audited workbook builder.

        This routine intentionally accepts only the generated builder metadata;
        it never guesses a product code or a name from a KiotViet error banner.
        """

        if not isinstance(raw_items, list):
            return []
        by_code: dict[str, dict[str, Any]] = {}
        for raw in raw_items:
            if not isinstance(raw, dict):
                continue
            code = str(raw.get("product_code") or raw.get("productCode") or "").strip()
            name = str(raw.get("product_name") or raw.get("productName") or "").strip()
            if not code or not name:
                continue
            try:
                source_row = int(raw.get("source_row") or raw.get("sourceRow") or 0)
            except (TypeError, ValueError):
                source_row = 0
            normalized = {
                "productCode": code,
                "productName": name,
                "sourceRow": source_row,
            }
            existing = by_code.get(code.casefold())
            if existing and existing["productName"] != name:
                raise EdgeAutomationError(
                    f"Mã hàng {code} xuất hiện với hai tên khác nhau trong file nguồn; chưa tự tạo hàng."
                )
            by_code[code.casefold()] = normalized
        return list(by_code.values())

    @classmethod
    def _missing_purchase_products_from_errors(
        cls, error_text: str, product_specs: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """Map Kiot's explicit unavailable-item banner back to source rows.

        KiotViet commonly reports ``Dòng 40: Mã hàng không có trên hệ thống
        hoặc ngừng kinh doanh`` without putting the SKU in the banner.  The
        mapping is deliberately conservative: when a row number is ambiguous,
        it returns no product so the import stops rather than creating the
        wrong product.
        """

        if not product_specs:
            return []
        key = cls._kiot_text_key(error_text)
        if "ma hang khong co tren he thong" not in key and "ngung kinh doanh" not in key:
            return []

        selected: dict[str, dict[str, Any]] = {}
        error_folded = str(error_text or "").casefold()
        for item in product_specs:
            if item["productCode"].casefold() in error_folded:
                selected[item["productCode"].casefold()] = item

        source_rows = {
            int(item["sourceRow"]): item
            for item in product_specs
            if isinstance(item.get("sourceRow"), int) and int(item["sourceRow"]) > 0
        }
        for match in re.finditer(r"dong\s*(\d+)\s*:", key):
            reported_row = int(match.group(1))
            from_source = source_rows.get(reported_row)
            # A generated Kiot workbook has its header in row 1, so a report
            # referring to its data-row N can also mean item N-1.  Use that
            # fallback only if the source-row interpretation is unavailable.
            from_generated = product_specs[reported_row - 2] if 2 <= reported_row <= len(product_specs) + 1 else None
            if from_source and from_generated and from_source["productCode"] != from_generated["productCode"]:
                continue
            candidate = from_source or from_generated
            if candidate:
                selected[candidate["productCode"].casefold()] = candidate
        return list(selected.values())

    @staticmethod
    def _purchase_product_barcode(code: str) -> str:
        """KiotViet barcode convention shown in the approved product form."""

        return re.sub(r"[^A-Za-z0-9]", "", str(code or "")).upper()

    @staticmethod
    def parse_purchase_date(value: Any) -> date:
        """Parse the explicit business date for a KiotViet purchase receipt.

        A purchase workbook's filename is not a trustworthy source of the
        receipt date.  The caller must provide an ISO date and we keep that
        date independent from whichever date KiotViet happened to prefill in
        the form.
        """

        raw = str(value or "").strip()
        try:
            parsed = datetime.strptime(raw, "%Y-%m-%d").date()
        except (TypeError, ValueError) as exc:
            raise EdgeAutomationError(
                "Ngày nhập hàng phải có định dạng YYYY-MM-DD, ví dụ 2026-08-31."
            ) from exc
        if parsed.isoformat() != raw:
            raise EdgeAutomationError(
                "Ngày nhập hàng phải có định dạng YYYY-MM-DD, ví dụ 2026-08-31."
            )
        return parsed

    def _set_purchase_date(self, driver: webdriver.Edge, target: date) -> str:
        """Set and verify KiotViet's visible Kendo purchase-date control.

        The date picker is intentionally handled through the normal calendar
        icon and visible input.  Direct Kendo/JavaScript value mutation skips
        Angular's change handling and can leave a receipt dated incorrectly.
        Preserve the operator-visible time component; only replace the date.
        """

        calendar_selector = (
            "span.k-icon.k-i-calendar"
            "[aria-controls='dateTimePickerPurchaseOrder_dateview']"
        )
        input_selector = "input#dateTimePickerPurchaseOrder"
        # KiotViet renders the new receipt's current date/time in the
        # placeholder while ``value`` remains empty.  If that visible date is
        # already the explicitly requested business date, leaving the
        # control untouched is safer than opening its Kendo popup and trying
        # to re-enter the identical value.  Newer Kiot builds keep that popup
        # open after a no-op edit, which blocks the supplier field.
        field = WebDriverWait(driver, 20).until(
            lambda current: self._visible_enabled(
                current.find_elements(By.CSS_SELECTOR, input_selector)
            )
        )
        existing = str(
            field.get_attribute("value")
            or field.get_attribute("placeholder")
            or ""
        ).strip()
        if existing.startswith(target.strftime("%d/%m/%Y")):
            self._logger.info(
                "KiotViet purchase receipt date already matches %s (visible=%s).",
                target.isoformat(),
                existing,
            )
            return existing
        calendar = WebDriverWait(driver, 20).until(
            lambda current: self._visible_enabled(
                current.find_elements(By.CSS_SELECTOR, calendar_selector)
            )
        )
        calendar.click()
        field = WebDriverWait(driver, 20).until(
            lambda current: self._visible_enabled(
                current.find_elements(By.CSS_SELECTOR, input_selector)
            )
        )
        # Kendo renders the current date as a placeholder until the operator
        # edits the control; in that state ``value`` is intentionally empty.
        # Read the placeholder as the fallback so we can preserve the visible
        # time component without guessing or changing the time.
        previous = str(
            field.get_attribute("value")
            or field.get_attribute("placeholder")
            or ""
        ).strip()
        time_match = re.search(r"\b([01]\d|2[0-3]):[0-5]\d\b", previous)
        if time_match is None:
            raise EdgeAutomationError(
                "KiotViet không hiển thị giờ hợp lệ trong ô ngày nhập; chưa thay đổi ngày phiếu."
            )
        expected = f"{target.strftime('%d/%m/%Y')} {time_match.group(0)}"
        field.click()
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys(expected)
        # Tab commits Kendo's normal blur/change flow.  Do not set its value
        # through page script: that can show a date without persisting it.
        field.send_keys(Keys.TAB)

        def committed(current: webdriver.Edge) -> str | bool:
            visible = self._visible_enabled(
                current.find_elements(By.CSS_SELECTOR, input_selector)
            )
            if visible is None:
                return False
            value = str(visible.get_attribute("value") or "").strip()
            return value if value.startswith(target.strftime("%d/%m/%Y")) else False

        committed_value = WebDriverWait(driver, 12).until(committed)
        self._logger.info(
            "Set KiotViet purchase receipt date to %s (visible=%s).",
            target.isoformat(),
            committed_value,
        )
        return str(committed_value)

    def import_purchase(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Import a purchase workbook, enable VAT, select OMEGA and finish.

        This is deliberately separate from the sales-order workflow.  It
        validates the required ĐN2 branch before any upload and never presses
        the sales ``ĐẶT HÀNG`` button.
        """
        file_path = Path(str(payload.get("filePath") or payload.get("file_path") or "")).resolve()
        if not file_path.is_file():
            raise EdgeAutomationError(f"Không tìm thấy file nhập hàng: {file_path}")
        purchase_date = self.parse_purchase_date(
            payload.get("purchaseDate") or payload.get("purchase_date")
        )
        auto_create_missing_products = payload.get("autoCreateMissingProducts") is True or payload.get("auto_create_missing_products") is True
        product_specs = self._normalize_purchase_product_specs(
            payload.get("products") or payload.get("productSpecs") or []
        )
        if auto_create_missing_products and not product_specs:
            raise EdgeAutomationError(
                "Chế độ tự tạo hàng đã bật nhưng file nguồn không trả về danh sách mã/tên hàng đã kiểm tra; chưa upload."
            )
        # Verify the branch while the sales shell exposes its authoritative
        # branch state.  The purchase form hides this label, so checking only
        # its visible text causes a false stop before any upload.
        self._navigate(KIOT_SALES_URL)
        self._assert_kiot_login()
        sales_state = self._kiot_status()
        if not sales_state.get("branchOk"):
            raise EdgeAutomationError(
                "Chỉ được nhập hàng khi KiotViet đang ở chi nhánh Đồng Nai 2 - ĐN2. "
                "Tool không xác minh được chi nhánh hiện tại; chưa upload file."
            )
        branch = str(sales_state.get("branch") or self.config.expected_branch)

        # KiotViet's purchase SPA only creates a fresh input form after the
        # user presses the plus icon on the purchase-list screen.  A missing
        # SKU can be recovered only when the user opted in and Kiot's error
        # maps exactly to a source line.  At most one create/re-upload cycle
        # is allowed; anything uncertain stops before a purchase is finished.
        created_products: list[str] = []
        for attempt in range(2 if auto_create_missing_products else 1):
            driver = self._open_kiot_purchase_form()
            upload_result = self._upload_purchase_workbook(driver, file_path)
            if upload_result["rows"] > 0:
                break
            error_text = str(upload_result.get("errorText") or "")
            missing = self._missing_purchase_products_from_errors(error_text, product_specs)
            if not auto_create_missing_products or attempt or not missing:
                detail = error_text or "KiotViet không trả về dòng hàng sau khi upload."
                raise EdgeAutomationError(
                    "KiotViet chưa nhận được file nhập hàng. "
                    f"Chi tiết: {detail[:900]}"
                )
            self._set_progress(
                active=True,
                phase="create_missing_products",
                message="KiotViet báo thiếu mã hàng; đang tạo các mã đã đối chiếu từ file nguồn.",
                current=0,
                total=len(missing),
            )
            created_products = self._create_missing_purchase_products(driver, missing)
        else:  # pragma: no cover - loop is guaranteed to return or raise above
            raise EdgeAutomationError("Không thể mở lại file nhập hàng sau khi tạo mã thiếu.")

        visible_purchase_date = self._set_purchase_date(driver, purchase_date)
        vat = WebDriverWait(driver, 20).until(
            lambda current: next((e for e in current.find_elements(By.CSS_SELECTOR, 'span.toogle[ng-click*="onToggleChange"]') if e.is_displayed()), None)
        )
        active = "active" in str(vat.get_attribute("class") or "").split()
        if not active:
            vat.click()
        supplier = WebDriverWait(driver, 20).until(
            lambda current: self._visible_enabled(
                current.find_elements(By.ID, "idSupplierSearchTerm")
            )
        )
        supplier.click()
        supplier.send_keys(Keys.CONTROL, "a")
        supplier.send_keys("OMEGA")
        suggestion = WebDriverWait(driver, 15).until(
            lambda current: next(
                (
                    e
                    for e in current.find_elements(
                        By.CSS_SELECTOR,
                        # The supplier autocomplete has ordinary ``li``
                        # children below its input.  Matching its enclosing
                        # ``div.autocomplete`` used to click the container,
                        # which closed the menu but never bound a supplier.
                        "input#idSupplierSearchTerm + div.output-complete li",
                    )
                    if e.is_displayed()
                    and "OMEGA" in self._element_text(e).upper()
                    and "02439366692" in self._element_text(e)
                ),
                None,
            )
        )
        suggestion.click()
        # A precise list-item click must both close the popup and leave the
        # searchable supplier field visible.  Do not press Hoàn thành while
        # a stale/open autocomplete could have discarded the selection.
        WebDriverWait(driver, 10).until(
            lambda current: (
                (selected := self._visible_enabled(
                    current.find_elements(By.ID, "idSupplierSearchTerm")
                )) is not None
                and "OMEGA" in str(selected.get_attribute("value") or "").upper()
                and not any(
                    e.is_displayed()
                    and "OMEGA" in self._element_text(e).upper()
                    and "02439366692" in self._element_text(e)
                    for e in current.find_elements(
                        By.CSS_SELECTOR,
                        "input#idSupplierSearchTerm + div.output-complete li",
                    )
                )
            )
        )
        finish = WebDriverWait(driver, 30).until(
            lambda current: next(
                (
                    e
                    for e in current.find_elements(By.CSS_SELECTOR, "a.kv-btn-primary.kv-btn-xl, button")
                    if e.is_displayed()
                    and e.is_enabled()
                    and self._element_text(e).strip().casefold() in {"hoàn thành", "hoan thanh"}
                ),
                None,
            )
        )
        finish.click()
        # KiotViet can display a confirmation that the supplier was not
        # actually bound (for example after a page-layout change).  Treat
        # that as a hard stop rather than reporting a completed receipt.
        warning_deadline = time.monotonic() + 4.0
        while time.monotonic() < warning_deadline:
            for dialog in driver.find_elements(
                By.CSS_SELECTOR,
                ".k-window-alert, [role='dialog']",
            ):
                try:
                    if dialog.is_displayed() and "chưa nhập nhà cung cấp" in self._kiot_text_key(
                        self._element_text(dialog)
                    ):
                        raise EdgeAutomationError(
                            "KiotViet chưa xác nhận nhà cung cấp OMEGA; chưa hoàn thành phiếu nhập."
                        )
                except StaleElementReferenceException:
                    continue
            time.sleep(0.2)
        return {
            "ok": True,
            "filePath": str(file_path),
            "supplier": "OMEGA",
            "vatEnabled": True,
            "purchaseDate": purchase_date.isoformat(),
            "purchaseDateVisible": visible_purchase_date,
            "branch": branch,
            "createdProducts": created_products,
        }

    @staticmethod
    def _visible_enabled(elements: list[Any]) -> Any | None:
        """Return only an interactable element, excluding Kendo's hidden copies."""

        for element in elements:
            try:
                if element.is_displayed() and element.is_enabled() and element.rect.get("width", 0) > 0:
                    return element
            except WebDriverException:
                continue
        return None

    def _open_kiot_purchase_form(self) -> webdriver.Edge:
        """Open a blank purchase form through the same visible plus button as UI."""

        driver = self._navigate(KIOT_PURCHASE_LIST_URL)
        plus = WebDriverWait(driver, 30).until(
            lambda current: self._visible_enabled(
                current.find_elements(By.CSS_SELECTOR, "i.far.fa-plus.icon-btn")
            )
        )
        plus.click()
        self._wait_document_ready(driver)
        self._wait_until(
            lambda: {
                "url": driver.current_url,
                "file_input": bool(driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]')),
            },
            lambda item: "purchaseorder" in str(item.get("url", "")).lower() and bool(item.get("file_input")),
            "màn hình Nhập hàng KiotViet",
            timeout_seconds=30,
        )
        return driver

    def _purchase_upload_error_text(self, driver: webdriver.Edge) -> str:
        """Return a visible Kiot import error only; do not infer errors from rows."""

        texts: list[str] = []
        selectors = (
            "[role='alert']",
            ".k-notification-error",
            ".k-tooltip-validation",
            ".toast-message",
            ".error-message",
            ".text-danger",
            ".k-upload-status",
        )
        for selector in selectors:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                try:
                    if element.is_displayed():
                        text = self._element_text(element)
                        if text:
                            texts.append(text)
                except WebDriverException:
                    continue
        # Some KiotViet releases place only the one-line validation banner in
        # body text. Read it solely after the known unavailable-item marker is
        # visible, never as a signal to finish a purchase.
        try:
            body = driver.find_element(By.TAG_NAME, "body").text
        except WebDriverException:
            body = ""
        body_key = self._kiot_text_key(body)
        if "ma hang khong co tren he thong" in body_key or "ngung kinh doanh" in body_key:
            texts.append(body)
        return "\n".join(dict.fromkeys(texts))

    def _dismiss_purchase_draft_dialog(
        self, driver: webdriver.Edge, *, timeout_seconds: float = 0.0
    ) -> bool:
        """Dismiss only KiotViet's exact unsaved purchase-draft prompt.

        Opening a new purchase form can surface a modal saying ``Mở bản nháp``
        when the previous browser session left an unsaved local draft.  That
        modal sits above the upload button (``.k-overlay``), so attempting to
        click ``Thực hiện`` first is rejected by WebDriver.  We deliberately
        match the modal title/body and the ``Bỏ qua`` action; no generic
        overlay is removed and no saved purchase is touched.
        """

        deadline = time.monotonic() + max(0.0, float(timeout_seconds))
        while True:
            containers: list[Any] = []
            for selector in (
                "[role='dialog']",
                ".modal-dialog",
                ".k-window",
                ".k-dialog",
                "[class*='modal']",
            ):
                try:
                    containers.extend(driver.find_elements(By.CSS_SELECTOR, selector))
                except WebDriverException:
                    continue
            seen: set[int] = set()
            for modal in containers:
                try:
                    marker = id(modal)
                    if marker in seen or not modal.is_displayed():
                        continue
                    seen.add(marker)
                    modal_key = self._kiot_text_key(self._element_text(modal))
                    if not (
                        "mo ban nhap" in modal_key
                        or ("he thong tim duoc" in modal_key and "ban nhap" in modal_key)
                    ):
                        continue
                    buttons = modal.find_elements(
                        By.XPATH,
                        ".//*[self::button or @role='button' or @ng-click or @onclick]",
                    )
                    skip = next(
                        (
                            button
                            for button in buttons
                            if button.is_displayed()
                            and button.is_enabled()
                            and self._kiot_text_key(self._element_text(button)) == "bo qua"
                        ),
                        None,
                    )
                    if skip is None:
                        continue
                    try:
                        ActionChains(driver).move_to_element(skip).pause(0.05).click(skip).perform()
                    except WebDriverException:
                        skip.click()
                    self._logger.info("Dismissed KiotViet unsaved purchase draft via Bỏ qua.")
                    return True
                except WebDriverException:
                    continue
            if time.monotonic() >= deadline:
                return False
            time.sleep(0.2)

    @staticmethod
    def _visible_purchase_overlay(driver: webdriver.Edge) -> bool:
        """Whether a visible Kendo/Bootstrap overlay still blocks the form."""

        for selector in (".k-overlay", ".modal-backdrop"):
            try:
                for overlay in driver.find_elements(By.CSS_SELECTOR, selector):
                    if not overlay.is_displayed():
                        continue
                    rect = overlay.rect
                    if float(rect.get("width", 0)) > 0 and float(rect.get("height", 0)) > 0:
                        return True
            except WebDriverException:
                continue
        return False

    def _upload_purchase_workbook(self, driver: webdriver.Edge, file_path: Path) -> dict[str, Any]:
        """Upload once and return either parsed rows or a visible validation error."""

        # A draft prompt can appear a moment after the plus button opens the
        # form.  Clear only that exact prompt before choosing a new workbook.
        self._dismiss_purchase_draft_dialog(driver, timeout_seconds=5.0)

        def file_input(current: webdriver.Edge) -> Any | None:
            # KiotViet's blue "Chọn file dữ liệu" button wraps the native
            # ``#files`` input (currently it is visible, but older releases
            # hide it).  The page also embeds KiCi's unrelated hidden chat
            # attachment input after the purchase form.  Selecting the last
            # generic ``input[type=file]`` therefore uploads the workbook to
            # chat and leaves the receipt blank.  Prefer KiotViet's stable
            # upload id, then use only its Kendo/name markers as fallbacks.
            selectors = (
                'input[type="file"]#files',
                'input[type="file"][kendo-upload]',
                'input[type="file"][data-role="upload"]',
                'input[type="file"][name="files"]',
            )
            for selector in selectors:
                for element in current.find_elements(By.CSS_SELECTOR, selector):
                    try:
                        if element.is_enabled() and not element.get_attribute("disabled"):
                            return element
                    except WebDriverException:
                        continue
            # Keep a narrow compatibility fallback for a Kiot page that has
            # changed all known attributes.  Never choose a chat/image input.
            for element in current.find_elements(By.CSS_SELECTOR, 'input[type="file"]'):
                try:
                    classes = str(element.get_attribute("class") or "").casefold()
                    accept = str(element.get_attribute("accept") or "").casefold()
                    if (
                        element.is_enabled()
                        and not element.get_attribute("disabled")
                        and "chat" not in classes
                        and "image/" not in accept
                    ):
                        return element
                except WebDriverException:
                    continue
            return None

        upload = WebDriverWait(driver, 20).until(file_input)
        upload.send_keys(str(file_path))
        # File selection can trigger the same prompt asynchronously.  Resolve
        # it before looking for the upload action so the modal overlay cannot
        # intercept the native click.
        self._dismiss_purchase_draft_dialog(driver, timeout_seconds=5.0)

        def upload_state(current: webdriver.Edge) -> dict[str, Any] | bool:
            execute = None
            if not self._visible_purchase_overlay(current):
                execute = self._visible_enabled(
                    current.find_elements(By.CSS_SELECTOR, "button.k-upload-selected")
                )
            rows = len(current.find_elements(By.CSS_SELECTOR, 'input[ng-model="dataItem.Price"]'))
            error_text = self._purchase_upload_error_text(current)
            return {"execute": execute, "rows": rows, "errorText": error_text} if execute or rows or error_text else False

        state = WebDriverWait(driver, 45).until(upload_state)
        if state["execute"] is not None:
            # Re-check the overlay immediately before the click; Angular may
            # finish closing the draft window between two polling ticks.
            execute = WebDriverWait(driver, 20).until(
                lambda current: (
                    False
                    if self._visible_purchase_overlay(current)
                    else self._visible_enabled(
                        current.find_elements(By.CSS_SELECTOR, "button.k-upload-selected")
                    )
                )
            )
            execute.click()

        deadline = time.monotonic() + 90
        last_error = str(state.get("errorText") or "")
        while time.monotonic() < deadline:
            rows = len(driver.find_elements(By.CSS_SELECTOR, 'input[ng-model="dataItem.Price"]'))
            if rows > 0:
                return {"rows": rows, "errorText": ""}
            error_text = self._purchase_upload_error_text(driver)
            if error_text:
                return {"rows": 0, "errorText": error_text}
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng thao tác Edge.")
            time.sleep(0.3)
        return {"rows": 0, "errorText": last_error}

    def _set_visible_input(self, driver: webdriver.Edge, selector: str, value: str, label: str) -> None:
        element = WebDriverWait(driver, 20).until(
            lambda current: self._visible_enabled(current.find_elements(By.CSS_SELECTOR, selector))
        )
        try:
            element.click()
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(value)
        except WebDriverException as exc:
            raise EdgeAutomationError(f"Không thể điền {label} khi tạo hàng hóa: {exc}") from exc

    def _click_visible_kiot_option(
        self,
        driver: webdriver.Edge,
        expected: str,
        label: str,
        *,
        selectors: tuple[str, ...] | None = None,
    ) -> None:
        expected_key = self._kiot_text_key(expected)

        def visible_match(current: webdriver.Edge) -> Any | None:
            folded_match = None
            for selector in selectors or (".k-animation-container li", "ul.k-list li", ".dropdown-menu li"):
                for element in current.find_elements(By.CSS_SELECTOR, selector):
                    try:
                        element_text = self._element_text(element)
                        if (
                            element.is_displayed()
                            and element.is_enabled()
                            and element.rect.get("width", 0) > 0
                        ):
                            # TBĐ and TBD are distinct categories in this
                            # KiotViet account but fold to the same ASCII key.
                            # Preserve the exact visible value when it exists;
                            # only then use the encoding-tolerant fallback.
                            if element_text.strip() == expected:
                                return element
                            if self._kiot_text_key(element_text) == expected_key and folded_match is None:
                                folded_match = element
                    except WebDriverException:
                        continue
            return folded_match

        option = WebDriverWait(driver, 15).until(visible_match)
        try:
            # Kiot's long category list is rendered below the modal in the
            # page scroll area. Scroll only the visual viewport; the click
            # below remains a normal WebDriver user action.
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", option
            )
            time.sleep(0.08)
            option.click()
        except WebDriverException as exc:
            try:
                # Kendo can initially render a long list below the modal's
                # viewport. A native pointer move scrolls it into view without
                # changing any form value through page JavaScript.
                ActionChains(driver).move_to_element(option).pause(0.08).click(option).perform()
            except WebDriverException as pointer_error:
                raise EdgeAutomationError(
                    f"Không thể chọn {label} '{expected}' khi tạo hàng hóa."
                ) from pointer_error

    def _set_product_kendo_value(
        self, driver: webdriver.Edge, select_css: str, expected: str, label: str
    ) -> None:
        """Choose a Kendo select by its visible wrapper rather than hidden select."""

        def visible_select(current: webdriver.Edge) -> Any | None:
            # Kiot leaves a hidden copy of the prior editor in the DOM. The
            # current modal is appended last, so examine that copy first.
            for candidate in reversed(current.find_elements(By.CSS_SELECTOR, select_css)):
                try:
                    wrapper = candidate.find_element(By.XPATH, "..")
                    if wrapper.is_displayed() and wrapper.rect.get("width", 0) > 0:
                        return candidate
                except WebDriverException:
                    continue
            return None

        select = WebDriverWait(driver, 20).until(visible_select)
        try:
            wrapper = select.find_element(By.XPATH, "..")
            raw_listbox_id = str(wrapper.get_attribute("aria-owns") or "").strip()
            select_id = str(select.get_attribute("id") or "").strip()
            # Kiot's category wrapper sometimes exposes the Angular template
            # literal ``ddlCat{{$id}}_listbox`` in aria-owns. The rendered
            # Kendo popup uses the concrete select id instead.
            listbox_id = raw_listbox_id or f"{select_id}_listbox"
            wrapper.click()
        except WebDriverException as exc:
            raise EdgeAutomationError(f"Không thể mở {label} khi tạo hàng hóa.") from exc
        try:
            # Native keyboard selection is reliable for Kiot's category list,
            # whose long popup is rendered below the modal viewport. It also
            # preserves the distinction between TBĐ and the separate TBD
            # category that has the same accent-folded key.
            ActionChains(driver).send_keys(expected).send_keys(Keys.ENTER).perform()

            def selected_expected(_current: webdriver.Edge) -> bool:
                try:
                    selected = select.find_element(By.CSS_SELECTOR, "option:checked")
                    # The native select is intentionally hidden by Kendo, so
                    # Selenium's visible-text property may be blank even when
                    # the selected option has changed.
                    text = str(selected.get_attribute("textContent") or self._element_text(selected)).strip()
                    return text == expected or self._kiot_text_key(text) == self._kiot_text_key(expected)
                except WebDriverException:
                    return False

            WebDriverWait(driver, 4).until(selected_expected)
            return
        except WebDriverException:
            # Retain the list-item route for Kiot builds where a Kendo widget
            # does not accept keyboard input.
            try:
                wrapper.click()
            except WebDriverException:
                pass
        # An attribute selector supports Kiot's literal Angular id
        # ``ddlCat{{$id}}_listbox`` without CSS escaping. Restricting to this
        # list prevents us from clicking an identically named option in a
        # hidden page filter behind the modal.
        listbox_selectors = (f'[id="{listbox_id}"] li',) if listbox_id else None
        self._click_visible_kiot_option(driver, expected, label, selectors=listbox_selectors)

    def _create_missing_purchase_products(
        self, driver: webdriver.Edge, products: list[dict[str, Any]]
    ) -> list[str]:
        """Create specifically mapped missing SKUs with the approved defaults.

        Defaults replicate the user's form: group ``TBĐ``, brand ``SPK``,
        cost 0, purchase/sales VAT 8%, direct-sale checked. Product availability
        is otherwise left to KiotViet's existing DN2 user session.
        """

        created: list[str] = []
        driver = self._navigate(KIOT_PRODUCTS_URL)
        for index, product in enumerate(products, start=1):
            plus = WebDriverWait(driver, 30).until(
                lambda current: self._visible_enabled(
                    current.find_elements(By.CSS_SELECTOR, "i.far.fa-plus.icon-btn")
                )
            )
            plus.click()
            menu = WebDriverWait(driver, 15).until(
                lambda current: self._visible_enabled(
                    current.find_elements(
                        By.CSS_SELECTOR,
                        'a.kv-dropdown-link[ng-click="AddProduct(pTypeValue.Purchased)"]',
                    )
                )
            )
            menu.click()
            self._set_visible_input(driver, "#idSuggestProductCodeSearchTerm", product["productCode"], "Mã hàng")
            self._set_visible_input(
                driver,
                "#idSuggestProductBarcodeSearchTerm",
                self._purchase_product_barcode(product["productCode"]),
                "Mã vạch",
            )
            self._set_visible_input(driver, "#idSuggestProductNameSearchTerm", product["productName"], "Tên hàng")
            self._set_visible_input(driver, 'input[ng-model="product.Cost"]', "0", "Giá vốn")
            self._set_product_kendo_value(driver, 'select[id^="ddlCat"]', "TBĐ", "Nhóm hàng")
            self._set_product_kendo_value(driver, "#ddlTradeMark", "SPK", "Thương hiệu")
            self._set_product_kendo_value(driver, 'select[k-ng-model="product.PurchaseTaxId"]', "8%", "VAT hàng nhập")
            self._set_product_kendo_value(driver, 'select[k-ng-model="product.TaxId"]', "8%", "VAT hàng bán")
            save = WebDriverWait(driver, 20).until(
                lambda current: self._visible_enabled(
                    current.find_elements(By.CSS_SELECTOR, 'a[ng-click="SaveProduct()"],button[ng-click="SaveProduct()"]')
                )
            )
            save.click()
            WebDriverWait(driver, 30).until(
                lambda current: not self._visible_enabled(
                    current.find_elements(By.CSS_SELECTOR, "#idSuggestProductCodeSearchTerm")
                )
            )
            created.append(product["productCode"])
            self._set_progress(
                active=True,
                phase="create_missing_products",
                message=f"Đã tạo mã hàng thiếu {product['productCode']}.",
                current=index,
                total=len(products),
            )
        return created

    def _assert_kiot_purchase_branch(self, driver: webdriver.Edge) -> str:
        """Fail closed unless the currently visible KiotViet branch is DN2."""

        expected = str(self.config.expected_branch).strip()

        def fold(value: Any) -> str:
            raw = unicodedata.normalize("NFKD", str(value or ""))
            return "".join(char for char in raw if not unicodedata.combining(char)).casefold().strip()

        expected_folded = fold(expected)
        for element in driver.find_elements(By.XPATH, "//*[normalize-space(text())]"):
            try:
                if not element.is_displayed():
                    continue
                text = self._element_text(element).strip()
                if fold(text) == expected_folded:
                    return text
            except WebDriverException:
                continue
        raise EdgeAutomationError(
            "Chỉ được nhập hàng khi KiotViet đang ở chi nhánh Đồng Nai 2 - ĐN2. "
            "Tool chưa tìm thấy nhãn chi nhánh đang chọn; không upload file."
        )

    def submit_batch(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Submit only orders whose visible notes still match the expected code."""

        orders = payload.get("orders")
        if not isinstance(orders, list) or not 1 <= len(orders) <= 19:
            raise EdgeAutomationError("A final-order batch must contain 1 to 19 orders.")
        submitted = self._submit_verified_orders(
            orders,
            batch_label=self._batch_label(payload),
        )
        return {
            "ok": submitted.get("ok") is True,
            # Keep the same nested proof shape as IMPORT_BATCH(autoSubmit),
            # so the workflow can atomically validate all per-order outcomes
            # before it advances SQLite to the next batch.
            "submitted": submitted,
            "allSubmitted": submitted.get("allSubmitted") is True,
            "failed": submitted.get("failed"),
            "message": submitted.get("message"),
            "finalOrderClicked": submitted.get("ok") is True,
        }

    def _submit_verified_orders(
        self,
        orders: list[dict[str, Any]],
        *,
        batch_label: str | None = None,
    ) -> dict[str, Any]:
        """Place a verified KiotViet batch one tab at a time.

        KiotViet can renumber tabs after a completed order.  Submitting from
        the largest tab number down preserves the labels of every remaining
        tab.  Each click is isolated in its own page task and has its own
        print-preview watcher, so a failure produces an exact partial result
        and *never* proceeds to later tabs.
        """

        plan = self._submission_plan(orders)
        self._assert_submission_preflight(plan)
        total = len(plan)
        outcomes: list[dict[str, Any]] = []
        self._set_progress(
            active=True,
            phase="submitting",
            message=f"Đang tự bấm ĐẶT HÀNG: 0/{total}.",
            current=0,
            total=total,
            batch=batch_label,
            order_code=None,
        )

        # Descending labels are deliberate: if KiotViet removes or renumbers
        # a completed cart, lower tab labels remain stable for the next step.
        for position, expectation in enumerate(
            sorted(plan, key=lambda item: item["tabNumber"], reverse=True),
            start=1,
        ):
            code = expectation["orderCode"]
            expected_label = expectation["expectedTabLabel"]
            self._set_progress(
                active=True,
                phase="submitting",
                message=f"Đang đặt đơn {position}/{total}: {code} tại {expected_label}.",
                current=position - 1,
                total=total,
                batch=batch_label,
                order_code=code,
            )

            stop_print_watcher = threading.Event()
            print_state: dict[str, Any] = {
                "cancelled": 0,
                "methods": [],
            }
            print_watcher = threading.Thread(
                target=self._watch_and_cancel_print_preview,
                args=(stop_print_watcher, print_state),
                name=f"kiot-print-canceller-{position}",
                daemon=True,
            )
            print_watcher.start()
            try:
                # The page task checks branch, exact tab, and note again
                # immediately before the click.  Later calls may legitimately
                # see an invoice tab created by an earlier successful order.
                # After a successful order KiotViet can briefly repaint its
                # sales shell; if that safe pre-click check reports the page
                # context only, wait/recover the same sales shell once before
                # quarantining the batch.  No final-order click has happened
                # when this particular assertion fails.
                raw = self._run_task(
                    "kiotSubmitOrdersTask",
                    [expectation],
                    self.config.expected_branch,
                    position > 1,
                    timeout_seconds=65,
                )
                try:
                    self._task_result_or_raise(raw)
                except EdgeAutomationError as context_error:
                    if not self._is_safe_submission_context_error(context_error):
                        raise
                    self._logger.warning(
                        "KiotViet sales shell changed before %s; recovering context once: %s",
                        code,
                        context_error,
                    )
                    submitted_codes = {
                        str(entry.get("orderCode") or entry.get("order_code") or "").strip()
                        for entry in outcomes
                        if isinstance(entry, dict)
                    }
                    remaining_plan = [
                        item for item in plan
                        if str(item.get("orderCode") or "").strip() not in submitted_codes
                    ]
                    self._recover_submission_context(remaining_plan)
                    raw = self._run_task(
                        "kiotSubmitOrdersTask",
                        [expectation],
                        self.config.expected_branch,
                        position > 1,
                        timeout_seconds=65,
                    )
                    self._task_result_or_raise(raw)
                item = self._validate_single_submission(raw, expectation)
            except (EdgeAutomationError, JavascriptException, TimeoutException, WebDriverException) as exc:
                stop_print_watcher.set()
                print_watcher.join(timeout=1.5)
                outcome = {
                    "orderCode": code,
                    "expectedTabLabel": expected_label,
                    "actualTabLabel": None,
                    "status": "failed",
                    "success": False,
                    "error": str(exc),
                    "printPreviewCancelled": bool(print_state["cancelled"]),
                    "printPreviewCancellationCount": int(print_state["cancelled"]),
                    "printPreviewMethods": list(print_state["methods"]),
                }
                outcomes.append(outcome)
                message = (
                    f"Dừng ở {code} ({expected_label}); {len(outcomes) - 1}/{total} "
                    "đơn trước đó đã được xác minh. Các đơn còn lại chưa bị bấm."
                )
                self._logger.error("KiotViet final-order failure: %s", outcome)
                self._set_progress(
                    active=False,
                    phase="failed",
                    message=message,
                    current=len(outcomes) - 1,
                    total=total,
                    batch=batch_label,
                    order_code=code,
                )
                return {
                    "ok": False,
                    "allSubmitted": False,
                    "submitted": outcomes,
                    "failed": outcome,
                    "message": message,
                }
            finally:
                stop_print_watcher.set()
                print_watcher.join(timeout=1.5)

            # Cover the short race where Edge opens the preview just after
            # the page task reports success but before the watcher stops.
            late_method = self._dismiss_print_preview(
                wait_seconds=0.45,
                native_fallback=False,
            )
            if late_method:
                print_state["cancelled"] += 1
                print_state["methods"].append(late_method)
            outcome = {
                **item,
                "status": "submitted",
                "success": True,
                "printPreviewCancelled": bool(print_state["cancelled"]),
                "printPreviewCancellationCount": int(print_state["cancelled"]),
                "printPreviewMethods": list(print_state["methods"]),
            }
            outcomes.append(outcome)
            self._logger.info("KiotViet final-order completed: %s", outcome)
            self._set_progress(
                active=True,
                phase="submitting",
                message=f"Đã đặt {position}/{total}: {code}.",
                current=position,
                total=total,
                batch=batch_label,
                order_code=code,
            )

        message = f"Đã tự đặt hàng thành công {total}/{total} đơn trong lô."
        self._set_progress(
            active=False,
            phase="completed",
            message=message,
            current=total,
            total=total,
            batch=batch_label,
            order_code=None,
        )
        return {
            "ok": True,
            "allSubmitted": True,
            "submitted": outcomes,
            "failed": None,
            "message": message,
        }

    @staticmethod
    def _batch_label(payload: dict[str, Any]) -> str | None:
        raw = (
            payload.get("batchLabel")
            or payload.get("batch_label")
            or payload.get("batchNumber")
            or payload.get("batch_number")
            or payload.get("batchId")
            or payload.get("batch_id")
        )
        value = str(raw or "").strip()
        return value or None

    @staticmethod
    def _submission_plan(orders: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if not isinstance(orders, list) or not 1 <= len(orders) <= 19:
            raise EdgeAutomationError("A final-order batch must contain 1 to 19 orders.")
        seen_codes: set[str] = set()
        seen_tabs: set[str] = set()
        plan: list[dict[str, Any]] = []
        pattern = re.compile(r"^\N{LATIN CAPITAL LETTER D WITH STROKE}ặt\s+hàng\s+([1-9]\d*)$", re.IGNORECASE)
        for source in orders:
            if not isinstance(source, dict):
                raise EdgeAutomationError("Mỗi đơn đặt hàng phải có mã và tab KiotViet dự kiến.")
            code = str(source.get("orderCode") or source.get("order_code") or "").strip()
            label = str(
                source.get("expectedTabLabel") or source.get("expected_tab_label") or ""
            )
            label = re.sub(r"\s+", " ", label).strip()
            match = pattern.fullmatch(label)
            if not code or not match:
                raise EdgeAutomationError(
                    "Không thể tự đặt hàng khi thiếu Mã đặt đơn hàng hoặc tab Đặt hàng dự kiến."
                )
            if code in seen_codes or label in seen_tabs:
                raise EdgeAutomationError("Danh sách tự đặt hàng chứa mã hoặc tab bị lặp.")
            seen_codes.add(code)
            seen_tabs.add(label)
            plan.append(
                {
                    "orderCode": code,
                    "expectedTabLabel": label,
                    "tabNumber": int(match.group(1)),
                }
            )
        return plan

    def _assert_submission_preflight(self, plan: list[dict[str, Any]]) -> None:
        status = self._kiot_status()
        if not status.get("branchOk") or not status.get("orderMode") or status.get("invoiceTabs"):
            raise EdgeAutomationError("ch\u01b0a c\u00f3 tab H\u00f3a \u0111\u01a1n an to\u00e0n; KiotViet ch\u01b0a s\u1eb5n s\u00e0ng.")
        available = {str(label).strip() for label in status.get("orderTabs", [])}
        if not available:
            raise EdgeAutomationError(
                "KiotViet không có tab đặt hàng nào để tiếp tục."
            )

    @staticmethod
    def _validate_single_submission(raw: dict[str, Any], expectation: dict[str, Any]) -> dict[str, Any]:
        completed = raw.get("submitted")
        if not isinstance(completed, list) or len(completed) != 1 or not isinstance(completed[0], dict):
            raise EdgeAutomationError("KiotViet không trả về kết quả xác minh cho đúng một đơn.")
        item = completed[0]
        code = str(item.get("orderCode") or "").strip()
        actual_label = str(item.get("actualTabLabel") or item.get("tabLabel") or "").strip()
        if item.get("success") is not True:
            raise EdgeAutomationError("KiotViet chưa xác nhận đặt hàng thành công.")
        if code != expectation["orderCode"]:
            raise EdgeAutomationError(
                "Kết quả KiotViet không khớp mã đơn/tab đã xác minh; không tiếp tục đơn kế tiếp."
            )
        return {
            "orderCode": code,
            "expectedTabLabel": expectation["expectedTabLabel"],
            "actualTabLabel": actual_label or expectation["expectedTabLabel"],
            "tabLabel": actual_label or expectation["expectedTabLabel"],
            "confirmationClicked": bool(item.get("confirmationClicked")),
            "completionProof": str(item.get("completionProof") or "verified"),
        }

    def _watch_and_cancel_print_preview(
        self,
        stop: threading.Event,
        state: dict[str, Any],
    ) -> None:
        """Close each real Edge print target until one order finishes.

        This runs independently of Selenium's single command channel.  A
        print preview can otherwise block the renderer before the page task
        has a chance to return its completion evidence.
        """

        # Edge can keep the same ``edge://print`` target in DevTools for a
        # moment after Cancel.  Handle each target id once so a quick polling
        # loop never clicks the sales screen after the preview disappears.
        handled_targets: set[str] = set()
        # A print target is often discoverable a few hundred milliseconds
        # before its shadow-DOM Cancel control is hydrated.  Do not mark it
        # handled until cancellation actually succeeds; otherwise the first
        # early miss leaves the preview open forever and the operator must
        # click Cancel manually.
        last_attempt_at: dict[str, float] = {}
        while not stop.is_set():
            driver = self._driver
            capabilities = driver.capabilities if driver is not None else {}
            preview, _reachable = self._find_edge_print_target(capabilities)
            preview_id = str(preview.get("id") or "") if preview else ""
            if not preview_id:
                handled_targets.clear()
                last_attempt_at.clear()
                stop.wait(0.12)
                continue
            now = time.monotonic()
            if (
                preview_id not in handled_targets
                and now - last_attempt_at.get(preview_id, 0.0) >= 0.4
            ):
                last_attempt_at[preview_id] = now
                method = self._dismiss_print_preview(
                    wait_seconds=0.35,
                    native_fallback=True,
                )
                if method:
                    handled_targets.add(preview_id)
                    state["cancelled"] = int(state.get("cancelled", 0)) + 1
                    state.setdefault("methods", []).append(method)
            stop.wait(0.12)

    def _dismiss_print_preview(
        self,
        *,
        wait_seconds: float = 8.0,
        native_fallback: bool = True,
    ) -> str | None:
        """Cancel a real KiotViet print preview without touching sales UI.

        An actual preview is positively identified through its exact
        ``edge://print`` target before any native input is sent.  This makes
        the visible Cancel click safe even when KiotViet is otherwise active.
        """

        if os.name != "nt":  # pragma: no cover - direct Edge runs on Windows
            return None

        driver = self._driver
        capabilities = driver.capabilities if driver is not None else {}
        preview, devtools_reachable = self._find_edge_print_target(capabilities)
        if preview is not None:
            # Edge's print preview is an internal shadow-DOM page.  Clicking
            # its real Cancel control through the target's DevTools context
            # is the same visible Cancel action the operator performs.  It
            # must be preferred over ``Target.closeTarget``: recent Edge
            # builds acknowledge closing ``edge://print`` but can keep that
            # target alive and leave the KiotViet renderer blocked.
            devtools_cancel = self._click_print_cancel_via_devtools(preview)
            if devtools_cancel:
                # Acknowledging the shadow-DOM click is not sufficient on
                # recent Edge builds: the ``edge://print`` target can remain
                # alive for several seconds and block the next Selenium
                # command.  Verify that the exact target disappeared before
                # telling the watcher it has been handled.  If it did not,
                # close that same target through DevTools; never leave a
                # stale preview marked as handled.
                gone_deadline = time.monotonic() + max(0.2, min(wait_seconds, 2.5))
                while time.monotonic() < gone_deadline:
                    current, _ = self._find_edge_print_target(capabilities)
                    if current is None:
                        return devtools_cancel
                    time.sleep(0.08)
                closed, _ = self._close_edge_print_target(
                    capabilities,
                    wait_seconds=max(0.5, min(wait_seconds, 3.0)),
                )
                if closed:
                    return f"{devtools_cancel}+devtools_close"
                # Return no success marker so the watcher retries the exact
                # preview target instead of moving on to another order.
                return None
            # If the close request was unavailable, try the visible Cancel
            # control as a narrow fallback.
            native = self._click_visible_print_cancel(capabilities)
            if native:
                return native
            # If the actual Cancel control cannot be reached, retain a
            # target-close fallback for old Edge builds.  The watcher calls
            # this no more than once per target.
            closed, _ = self._close_edge_print_target(
                capabilities,
                wait_seconds=wait_seconds,
            )
            return "devtools_close" if closed else None
        if devtools_reachable or not native_fallback:
            return None
        return None

    def _click_print_cancel_via_devtools(self, preview: dict[str, Any]) -> str | None:
        """Activate the visible ``Cancel`` button in Edge's print UI.

        ``edge://print`` renders its controls inside nested shadow roots, so
        ordinary Selenium selectors cannot reach them.  The expression walks
        only that internal print target and clicks an exact
        ``fluent-button.cancel`` element; it does not execute in the
        KiotViet page context.
        """

        if websocket is None:
            return None
        socket_url = str(preview.get("webSocketDebuggerUrl") or "").strip()
        if not socket_url:
            return None
        expression = (
            "(() => {"
            "let found = null;"
            "const isCancel = (element) => {"
            "if (!element || !element.matches) return false;"
            "const text = String(element.innerText || element.textContent || '').trim();"
            "const selector = 'fluent-button.cancel,button.cancel,#cancel,[aria-label=\"Cancel\"],[role=\"button\"]';"
            "return element.matches('fluent-button.cancel,button.cancel,#cancel,[aria-label=\"Cancel\"]')"
            " || (element.matches('[role=\"button\"]') && /^Cancel$/i.test(text));"
            "};"
            "function walk(root) {"
            "if (!root || found) return;"
            "if (root.nodeType === 1 && isCancel(root)) { found = root; return; }"
            "for (const element of (root.querySelectorAll ? root.querySelectorAll('*') : [])) {"
            "if (isCancel(element)) { found = element; break; }"
            "if (element.shadowRoot) walk(element.shadowRoot);"
            "}"
            "}"
            "walk(document);"
            "if (!found) return 'not-found';"
            "found.click();"
            "return 'clicked';"
            "})()"
        )
        try:
            socket = websocket.create_connection(socket_url, timeout=2.5, suppress_origin=True)
            try:
                socket.send(
                    json.dumps(
                        {
                            "id": 1,
                            "method": "Runtime.evaluate",
                            "params": {
                                "expression": expression,
                                "returnByValue": True,
                                "userGesture": True,
                            },
                        }
                    )
                )
                reply = json.loads(socket.recv())
            finally:
                socket.close()
            value = str(
                reply.get("result", {})
                .get("result", {})
                .get("value", "")
            )
            if value == "clicked":
                self._logger.info("Clicked Edge print-preview Cancel through its visible shadow control.")
                return "devtools_cancel"
        except Exception as exc:  # pragma: no cover - best effort desktop bridge
            self._logger.debug("DevTools print-preview Cancel click failed: %s", exc)
        return None

    @staticmethod
    def _browser_process_id(capabilities: dict[str, Any]) -> int | None:
        for candidate in (
            capabilities.get("goog:processID"),
            capabilities.get("ms:edgeOptions", {}).get("processId")
            if isinstance(capabilities.get("ms:edgeOptions"), dict)
            else None,
        ):
            try:
                process_id = int(candidate)
            except (TypeError, ValueError):
                continue
            if process_id > 0:
                return process_id
        return None

    def _edge_profile_process_id(self) -> int | None:
        """Return the top-level Edge PID for this tool's dedicated profile.

        A normal WebDriver launch reports ``goog:processID``.  A WebDriver
        session attached to the already-visible Edge profile deliberately does
        not always expose that capability, however.  The native Office
        ``Download file`` control is scoped by PID, so in that one case look
        up only the top-level ``msedge.exe`` whose command line owns this
        application's exact ``--user-data-dir``.  We never select a generic
        Edge process or a renderer/utility child process.
        """

        if os.name != "nt":  # pragma: no cover - this desktop helper is Windows-only
            return None

        profile_path = str(self.config.edge_profile_dir.resolve())
        # ``powershell.exe -Command`` reparses quotation marks inside a script
        # argument.  Encode the fixed script and pass the runtime path only
        # through an environment variable, so unusual Windows path characters
        # cannot alter the process query.
        script = r"""
$needle = [IO.Path]::GetFullPath($env:ORDERBRIDGE_EDGE_PROFILE).TrimEnd('\').ToLowerInvariant()
$candidates = Get-CimInstance Win32_Process -Filter "Name = 'msedge.exe'" |
  Where-Object {
    $command = [string]$_.CommandLine
    $command -and $command.ToLowerInvariant().Contains($needle)
  }
$root = $candidates |
  Where-Object { ([string]$_.CommandLine) -notmatch '(?i)(?:^|\s)--type=' } |
  Select-Object -First 1
if ($null -eq $root) {
  $root = $candidates | Select-Object -First 1
}
if ($null -ne $root) {
  [Console]::Out.WriteLine([int]$root.ProcessId)
}
"""
        encoded_script = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        environment = os.environ.copy()
        environment["ORDERBRIDGE_EDGE_PROFILE"] = profile_path
        try:
            completed = subprocess.run(
                [
                    str(self.config.powershell_path),
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded_script,
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                timeout=8,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:  # pragma: no cover - host specific
            self._logger.warning("Could not locate the dedicated Edge profile PID: %s", exc)
            return None

        if completed.returncode != 0:
            self._logger.warning(
                "Could not locate the dedicated Edge profile PID (PowerShell exit %s): %s",
                completed.returncode,
                (completed.stderr or "").strip(),
            )
            return None
        for line in reversed((completed.stdout or "").splitlines()):
            try:
                process_id = int(line.strip())
            except ValueError:
                continue
            if process_id > 0:
                return process_id
        return None

    def _session_browser_process_id(self, capabilities: dict[str, Any]) -> int | None:
        """Resolve the visible Edge PID for a launched or debugger-attached session."""

        return self._browser_process_id(capabilities) or self._edge_profile_process_id()

    def _click_visible_print_cancel(self, capabilities: dict[str, Any]) -> str | None:
        """Click the visible Edge Preview Cancel control exactly once.

        The current Edge print-preview shadow UI is not exposed through
        UIAutomation or a usable Runtime context.  Its left sidebar layout is
        stable: Cancel is at 17.4% of browser width and 80.4% of height.  The
        caller already verified ``edge://print`` before this method runs.
        """

        browser_pid = self._session_browser_process_id(capabilities)
        user32 = ctypes.windll.user32
        expected_title = f"{self.config.expected_branch} - Bán hàng"
        windows: list[int] = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def collect_window(handle: int, _parameter: int) -> bool:
            if not user32.IsWindowVisible(handle):
                return True
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(handle, ctypes.byref(pid))
            if browser_pid is not None and int(pid.value) != browser_pid:
                return True
            text_length = user32.GetWindowTextLengthW(handle)
            if text_length <= 0:
                return True
            text_buffer = ctypes.create_unicode_buffer(text_length + 1)
            user32.GetWindowTextW(handle, text_buffer, text_length + 1)
            if expected_title in text_buffer.value:
                windows.append(int(handle))
            return True

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        try:
            user32.EnumWindows(collect_window, 0)
            if not windows:
                self._logger.warning("KiotViet Edge window was not found for print cancellation.")
                return None
            handle = windows[0]
            rect = RECT()
            if not user32.GetWindowRect(handle, ctypes.byref(rect)):
                return None
            width = int(rect.right - rect.left)
            height = int(rect.bottom - rect.top)
            if width < 600 or height < 500:
                self._logger.warning(
                    "KiotViet Edge window is too small for a safe print Cancel click: %sx%s.",
                    width,
                    height,
                )
                return None
            cancel_x = int(round(rect.left + width * 0.174))
            cancel_y = int(round(rect.top + height * 0.804))
            user32.ShowWindow(handle, 5)  # SW_SHOW
            user32.SetForegroundWindow(handle)
            time.sleep(0.12)
            if int(user32.GetForegroundWindow()) != int(handle):
                self._logger.warning("Could not foreground KiotViet Edge for print cancellation.")
                return None
            user32.SetCursorPos(cancel_x, cancel_y)
            time.sleep(0.08)
            user32.mouse_event(0x0002, cancel_x, cancel_y, 0, 0)  # left down
            user32.mouse_event(0x0004, cancel_x, cancel_y, 0, 0)  # left up
            self._logger.info(
                "Clicked visible KiotViet print Cancel at %s,%s (pid=%s).",
                cancel_x,
                cancel_y,
                browser_pid,
            )
            return "native_cancel"
        except Exception as exc:  # pragma: no cover - best-effort desktop bridge
            self._logger.warning("Could not click KiotViet print Cancel: %s", exc)
            return None

    def _find_edge_print_target(
        self,
        capabilities: dict[str, Any],
    ) -> tuple[dict[str, Any] | None, bool]:
        """Return a positive print-preview target without mutating it."""

        edge_options = capabilities.get("ms:edgeOptions")
        debugger_address = (
            edge_options.get("debuggerAddress")
            if isinstance(edge_options, dict)
            else None
        )
        if not isinstance(debugger_address, str) or not debugger_address.strip():
            return None, False
        try:
            with urlopen(f"http://{debugger_address.strip()}/json", timeout=1.5) as response:
                targets = json.load(response)
            preview = next(
                (
                    target
                    for target in targets
                    if isinstance(target, dict)
                    and target.get("type") == "page"
                    and str(target.get("url") or "").startswith("edge://print")
                ),
                None,
            )
            return preview, True
        except Exception as exc:  # pragma: no cover - desktop/DevTools availability
            self._logger.debug("Could not inspect Edge print target: %s", exc)
            return None, False

    def _close_edge_print_target(
        self,
        capabilities: dict[str, Any],
        *,
        wait_seconds: float = 8.0,
    ) -> tuple[bool, bool]:
        """Close only this session's ``edge://print/`` DevTools target.

        Edge's print preview is an internal Chromium page.  Sending Escape to
        the window can leave that target alive and block the active WebDriver
        command; closing the exact target is equivalent to pressing Cancel.
        """

        if websocket is None:
            return False, False
        edge_options = capabilities.get("ms:edgeOptions")
        debugger_address = (
            edge_options.get("debuggerAddress")
            if isinstance(edge_options, dict)
            else None
        )
        if not isinstance(debugger_address, str) or not debugger_address.strip():
            return False, False
        address = debugger_address.strip()
        deadline = time.monotonic() + max(0.0, wait_seconds)
        endpoint_reachable = False
        while True:
            try:
                with urlopen(f"http://{address}/json", timeout=1.5) as response:
                    targets = json.load(response)
                endpoint_reachable = True
                preview = next(
                    (
                        target
                        for target in targets
                        if isinstance(target, dict)
                        and target.get("type") == "page"
                        and str(target.get("url") or "").startswith("edge://print")
                    ),
                    None,
                )
                if preview is None:
                    if time.monotonic() >= deadline:
                        return False, endpoint_reachable
                    time.sleep(min(0.12, max(0.01, deadline - time.monotonic())))
                    continue
                with urlopen(f"http://{address}/json/version", timeout=1.5) as response:
                    version = json.load(response)
                socket_url = str(version.get("webSocketDebuggerUrl") or "")
                target_id = str(preview.get("id") or "")
                if not socket_url or not target_id:
                    return False, endpoint_reachable
                socket = websocket.create_connection(
                    socket_url,
                    timeout=2.0,
                    suppress_origin=True,
                )
                try:
                    socket.send(
                        json.dumps(
                            {
                                "id": 1,
                                "method": "Target.closeTarget",
                                "params": {"targetId": target_id},
                            }
                        )
                    )
                    reply = json.loads(socket.recv())
                finally:
                    socket.close()
                if bool(reply.get("result", {}).get("success")):
                    self._logger.info("Closed KiotViet edge://print preview through DevTools.")
                    return True, endpoint_reachable
                return False, endpoint_reachable
            except Exception as exc:  # pragma: no cover - best effort only
                self._logger.debug("DevTools print-preview cancellation retry: %s", exc)
            if time.monotonic() >= deadline:
                # A reachable endpoint with no print target means there is
                # nothing to cancel.  Do not send Escape to the sales screen.
                return False, endpoint_reachable
            time.sleep(0.08)

    # ------------------------------------------------------------------
    # Selenium lifecycle and helpers
    # ------------------------------------------------------------------
    def _ensure_driver(self) -> webdriver.Edge:
        if self._driver is not None:
            try:
                # EdgeDriver may attach to a stale native print-preview target
                # left by a previous order.  Recover only from that internal
                # page.  Selecting KiotViet unconditionally here used to
                # switch a freshly navigated Sopoka tab back to the sales
                # canvas whenever a later selector helper called
                # ``_ensure_driver``; the next Sopoka lookup then failed even
                # though navigation had succeeded.
                current_url = str(self._driver.current_url or "").casefold()
                if current_url.startswith(("edge://print", "chrome-untrusted://print")):
                    self._select_sales_target_direct(self._driver)
                self._last_known_url = self._driver.current_url
                return self._driver
            except WebDriverException:
                self._driver = None
        if not self.config.edge_driver_path.is_file():
            raise EdgeAutomationError(f"Không tìm thấy msedgedriver.exe: {self.config.edge_driver_path}")
        if not self.config.edge_binary_path.is_file():
            raise EdgeAutomationError(f"Không tìm thấy Microsoft Edge: {self.config.edge_binary_path}")

        # When the local server is restarted, Edge itself can remain open with
        # imported KiotViet tabs.  Prefer attaching to that exact persistent
        # profile before attempting a new launch so a restart never discards a
        # pending batch merely because the Python process changed.
        # A final-order click can briefly restart Edge's top-level process
        # while keeping the same profile and KiotViet tabs.  Retry the exact
        # profile DevTools attach for a short window before ever launching a
        # second Edge instance into that locked profile.
        for attempt in range(8):
            attached = self._attach_to_live_profile()
            if attached is not None:
                self._driver = attached
                self._select_sales_target_direct(attached)
                return attached
            if attempt < 7:
                time.sleep(0.5)

        options = EdgeOptions()
        options.binary_location = str(self.config.edge_binary_path)
        # KiotViet is an SPA and clicking the final order button can trigger
        # an in-page route change.  With Selenium's default ``normal`` page
        # load strategy, the *next* DOM command can sit behind that transition
        # for minutes even though the button click has already succeeded.
        # Navigations made by this controller are explicitly waited for in
        # ``_navigate`` below, so it is safe (and much more predictable) to
        # leave route transitions non-blocking here.
        options.page_load_strategy = "none"
        options.add_argument(f"--user-data-dir={self.config.edge_profile_dir}")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--start-maximized")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(self.config.edge_download_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_setting_values.notifications": 2,
            },
        )
        try:
            self._driver = webdriver.Edge(
                service=EdgeService(executable_path=str(self.config.edge_driver_path)),
                options=options,
            )
        except WebDriverException as exc:
            raise EdgeAutomationError(
                "Không thể mở Edge automation. Đóng mọi cửa sổ đang dùng profile "
                f"{self.config.edge_profile_dir} rồi thử lại. Chi tiết: {exc}"
            ) from exc
        self._configure_driver_session(self._driver)
        # A persistent profile can restore many KiotViet carts immediately
        # after Edge starts.  Give that restore a short head start before the
        # first Angular inspection; otherwise EdgeDriver intermittently
        # reports a script timeout even though the visible page soon loads.
        time.sleep(12)
        self._last_known_url = self._safe_current_url()
        self._logger.info("Started direct Edge automation profile=%s", self.config.edge_profile_dir)
        return self._driver

    def _attach_to_live_profile(self) -> webdriver.Edge | None:
        """Attach to the profile's live Edge DevTools endpoint when present."""

        address = self._live_profile_debugger_address()
        if not address:
            return None
        # Edge's built-in Office viewer opens the downloaded Sopoka report in
        # a second Chromium target.  On this Edge build, EdgeDriver can pick
        # that non-interactive viewer rather than the active KiotViet page
        # while creating a debugger-attached session, and then wait forever
        # for Office's runtime to answer.  The report was already saved to
        # the controlled download folder before any import begins, so remove
        # only that completed ReportOrderSaleSummary viewer target.  No normal
        # sales, Sopoka, or user tab is closed here.
        self._close_completed_sopoka_viewer(address)
        self._activate_kiot_target(address)
        options = EdgeOptions()
        options.page_load_strategy = "none"
        options.debugger_address = address
        try:
            driver = webdriver.Edge(
                service=EdgeService(executable_path=str(self.config.edge_driver_path)),
                options=options,
            )
            self._configure_driver_session(driver)
            self._last_known_url = driver.current_url
            self._logger.info("Reattached to visible Edge profile at %s", address)
            return driver
        except WebDriverException as exc:
            # The ActivePort file can be stale after an unclean Edge exit.  A
            # normal launch below remains the safe fallback.
            self._logger.info("Could not reattach to existing Edge profile at %s: %s", address, exc)
            return None

    def _activate_kiot_target(self, address: str) -> None:
        """Activate the live KiotViet target before Selenium enumerates tabs."""

        if websocket is None:
            return
        socket = None
        try:
            with urlopen(f"http://{address}/json/version", timeout=1.5) as response:
                version = json.load(response)
            socket_url = str(version.get("webSocketDebuggerUrl") or "")
            if not socket_url:
                return
            with urlopen(f"http://{address}/json/list", timeout=1.5) as response:
                targets = json.load(response)
            target = next(
                (
                    item for item in targets
                    if isinstance(item, dict)
                    and item.get("type") == "page"
                    and self._is_kiot_sales_url(item.get("url"))
                ),
                None,
            )
            target_id = str(target.get("id") or "") if isinstance(target, dict) else ""
            if not target_id:
                return
            socket = websocket.create_connection(socket_url, timeout=2.0, suppress_origin=True)
            socket.send(json.dumps({"id": 1, "method": "Target.activateTarget", "params": {"targetId": target_id}}))
            socket.recv()
        except Exception as exc:  # pragma: no cover - best effort attach hygiene
            self._logger.debug("Could not activate KiotViet target before Edge attach: %s", exc)
        finally:
            if socket is not None:
                try:
                    socket.close()
                except Exception:
                    pass

    def _select_kiot_sales_window(self, driver: webdriver.Edge) -> None:
        """Focus the existing KiotViet sales page instead of print preview."""

        try:
            handles = list(driver.window_handles)
        except WebDriverException:
            return
        fallback: str | None = None
        for handle in handles:
            try:
                driver.switch_to.window(handle)
                url = str(driver.current_url or "")
            except WebDriverException:
                continue
            if self._is_kiot_sales_url(url):
                self._last_known_url = url
                return
            if fallback is None and not url.casefold().startswith(("edge://print", "chrome-untrusted://print")):
                fallback = handle
        if fallback is not None:
            try:
                driver.switch_to.window(fallback)
            except WebDriverException:
                pass

    def _prepare_driver_window_selection(self, driver: webdriver.Edge) -> None:
        """Make WebDriver window enumeration safe when print preview is stale."""

        try:
            current = str(driver.current_url or "").casefold()
        except WebDriverException:
            current = ""
        if current.startswith("edge://print") or current.startswith("chrome-untrusted://print"):
            # Do this only after EdgeDriver has attached.  Closing the target
            # before attach can cause EdgeDriver itself to wait forever while
            # Chromium recreates the native preview target.
            self._close_edge_print_target(driver.capabilities, wait_seconds=1.5)
        try:
            self._select_kiot_sales_window(driver)
        except WebDriverException:
            # A preview can be recreated during the short transition.  Leave
            # the attached session intact; the next command will retry the
            # same guarded selection instead of launching another Edge.
            self._logger.warning("KiotViet window selection deferred while Edge print preview transitions.")

    def _select_sales_target_direct(self, driver: webdriver.Edge) -> None:
        """Switch to KiotViet by target ID without enumerating window handles."""
        edge_options = driver.capabilities.get("ms:edgeOptions", {})
        address = edge_options.get("debuggerAddress") if isinstance(edge_options, dict) else None
        if not isinstance(address, str) or not address.strip():
            return
        try:
            with urlopen(f"http://{address.strip()}/json/list", timeout=1.5) as response:
                targets = json.load(response)
            target = next((item for item in targets if isinstance(item, dict) and item.get("type") == "page" and self._is_kiot_sales_url(item.get("url"))), None)
            target_id = str(target.get("id") or "") if isinstance(target, dict) else ""
            if not target_id:
                return
            self._close_edge_print_target(driver.capabilities, wait_seconds=0.8)
            driver.switch_to.window(target_id)
            self._last_known_url = str(driver.current_url or "")
        except (OSError, ValueError, WebDriverException) as exc:
            self._logger.debug("Could not select KiotViet target directly: %s", exc)

    def _close_completed_sopoka_viewer(self, address: str) -> None:
        """Close stale in-browser viewers for an already-downloaded report.

        This runs only immediately before attaching a fresh WebDriver after a
        local-server restart.  It solves an EdgeDriver target-selection bug
        without discarding KiotViet carts.  ``suppress_origin`` is required
        by current Edge DevTools security rules for localhost WebSocket
        clients.
        """

        if websocket is None:
            return
        try:
            with urlopen(f"http://{address}/json/version", timeout=1.5) as response:
                version = json.load(response)
            socket_url = str(version.get("webSocketDebuggerUrl") or "")
            if not socket_url:
                return
            socket = websocket.create_connection(
                socket_url,
                timeout=2.0,
                suppress_origin=True,
            )
        except Exception as exc:  # pragma: no cover - best-effort startup hygiene
            self._logger.debug("Could not inspect live Edge targets before attach: %s", exc)
            return

        next_command_id = 0

        def command(method: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
            nonlocal next_command_id
            next_command_id += 1
            command_id = next_command_id
            socket.send(
                json.dumps(
                    {
                        "id": command_id,
                        "method": method,
                        "params": params or {},
                    }
                )
            )
            deadline = time.monotonic() + 2.0
            while time.monotonic() < deadline:
                reply = json.loads(socket.recv())
                if reply.get("id") == command_id:
                    return reply
            return {}

        try:
            initial = command("Target.getTargets")
            targets = initial.get("result", {}).get("targetInfos", [])
            if not isinstance(targets, list):
                return
            viewer_ids = [
                str(target.get("targetId") or "")
                for target in targets
                if isinstance(target, dict)
                and "view.officeapps.live.com/op/view" in str(target.get("url") or "").casefold()
                and "reportordersalesummary" in (
                    str(target.get("title") or "")
                    + " "
                    + str(target.get("url") or "")
                ).casefold()
            ]
            viewer_ids = [target_id for target_id in viewer_ids if target_id]
            if not viewer_ids:
                return
            closed_ids: list[str] = []
            for target_id in viewer_ids:
                reply = command("Target.closeTarget", {"targetId": target_id})
                if reply.get("result", {}).get("success") is True:
                    closed_ids.append(target_id)
            if closed_ids:
                # Let Chromium retire the target before EdgeDriver enumerates
                # tabs for its debugger session.  This is not a browser wait;
                # it is a narrowly bounded DevTools state check.
                deadline = time.monotonic() + 2.0
                while time.monotonic() < deadline:
                    remaining = command("Target.getTargets").get("result", {}).get("targetInfos", [])
                    present = {
                        str(item.get("targetId") or "")
                        for item in remaining
                        if isinstance(item, dict)
                    }
                    if not any(target_id in present for target_id in closed_ids):
                        break
                    time.sleep(0.08)
                self._logger.info(
                    "Closed %s completed Sopoka Office viewer target(s) before Edge attach.",
                    len(closed_ids),
                )
        except Exception as exc:  # pragma: no cover - startup must still try attach
            self._logger.debug("Could not close completed Sopoka Office viewer: %s", exc)
        finally:
            try:
                socket.close()
            except Exception:
                pass

    def _live_profile_debugger_address(self) -> str | None:
        """Return a verified DevTools address for this app's persistent Edge.

        Chromium normally writes ``DevToolsActivePort`` for an Edge process
        launched with ``--remote-debugging-port=0``.  Some current Edge builds
        can remove that file while the process and its local DevTools listener
        remain alive.  Retain the file path as the fast, normal path, then use
        a deliberately narrow Windows fallback: inspect only loopback
        listeners owned by the exact top-level ``msedge.exe`` that has this
        app's ``--user-data-dir``.  Every candidate is still required to
        answer the local ``/json/version`` DevTools endpoint before it can be
        returned.  No generic local port is ever scanned or attached.
        """

        active_port = self.config.edge_profile_dir / "DevToolsActivePort"
        try:
            lines = active_port.read_text(encoding="utf-8").splitlines()
            port = int(lines[0].strip())
        except (OSError, ValueError):
            port = None
        if port is not None:
            address = self._validated_loopback_devtools_address("127.0.0.1", port)
            if address:
                return address

        profile_process_id = self._edge_profile_process_id()
        if profile_process_id is None:
            return None
        for host, candidate_port in self._profile_loopback_listener_addresses(profile_process_id):
            address = self._validated_loopback_devtools_address(host, candidate_port)
            if address:
                self._logger.info(
                    "Recovered live Edge DevTools endpoint at %s through dedicated profile PID %s.",
                    address,
                    profile_process_id,
                )
                return address
        return None

    @staticmethod
    def _loopback_debugger_address(host: str, port: int) -> str | None:
        """Format one allowed loopback DevTools endpoint, or reject it."""

        if host not in {"127.0.0.1", "::1"} or not 1 <= port <= 65535:
            return None
        return f"[{host}]:{port}" if host == "::1" else f"{host}:{port}"

    def _validated_loopback_devtools_address(self, host: str, port: int) -> str | None:
        """Validate a single, already-scoped local DevTools candidate."""

        address = self._loopback_debugger_address(host, port)
        if not address:
            return None
        try:
            with urlopen(f"http://{address}/json/version", timeout=0.75) as response:
                info = json.load(response)
            if not isinstance(info, dict):
                return None
            socket_url = str(info.get("webSocketDebuggerUrl") or "")
            parsed_socket = urlsplit(socket_url)
            if (
                parsed_socket.scheme != "ws"
                or parsed_socket.hostname not in {"127.0.0.1", "::1"}
                or parsed_socket.port != port
                or not parsed_socket.path.startswith("/devtools/browser/")
            ):
                return None
            return address
        except (OSError, ValueError, json.JSONDecodeError):
            return None

    def _profile_loopback_listener_addresses(self, process_id: int) -> list[tuple[str, int]]:
        """List only loopback TCP listeners owned by one exact Edge root PID.

        The caller obtains ``process_id`` via ``_edge_profile_process_id``,
        which matches this app's resolved profile directory and excludes Edge
        renderer/utility children.  PowerShell receives only an integer via an
        environment variable; its output is parsed again defensively here.
        """

        if os.name != "nt" or process_id <= 0:  # pragma: no cover - Windows-only helper
            return []
        script = r"""
$edgeProcessId = [int]$env:ORDERBRIDGE_EDGE_PID
Get-NetTCPConnection -State Listen -OwningProcess $edgeProcessId -ErrorAction SilentlyContinue |
  Where-Object { $_.LocalAddress -eq '127.0.0.1' -or $_.LocalAddress -eq '::1' } |
  Sort-Object @{ Expression = { if ($_.LocalAddress -eq '127.0.0.1') { 0 } else { 1 } } }, LocalPort |
  ForEach-Object { [Console]::Out.WriteLine(("{0}|{1}" -f $_.LocalAddress, $_.LocalPort)) }
"""
        encoded_script = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        environment = os.environ.copy()
        environment["ORDERBRIDGE_EDGE_PID"] = str(process_id)
        try:
            completed = subprocess.run(
                [
                    str(self.config.powershell_path),
                    "-NoLogo",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded_script,
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                timeout=5,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:  # pragma: no cover - host specific
            self._logger.warning(
                "Could not inspect loopback listeners for dedicated Edge PID %s: %s",
                process_id,
                exc,
            )
            return []
        if completed.returncode != 0:
            self._logger.warning(
                "Could not inspect loopback listeners for dedicated Edge PID %s "
                "(PowerShell exit %s): %s",
                process_id,
                completed.returncode,
                (completed.stderr or "").strip(),
            )
            return []

        listeners: list[tuple[str, int]] = []
        seen: set[tuple[str, int]] = set()
        for line in (completed.stdout or "").splitlines():
            host, separator, raw_port = line.strip().partition("|")
            if not separator or host not in {"127.0.0.1", "::1"}:
                continue
            try:
                port = int(raw_port.strip())
            except ValueError:
                continue
            candidate = (host, port)
            if self._loopback_debugger_address(*candidate) and candidate not in seen:
                listeners.append(candidate)
                seen.add(candidate)
        return listeners

    @staticmethod
    def _configure_driver_session(driver: webdriver.Edge) -> None:
        driver.set_page_load_timeout(45)
        driver.set_script_timeout(90)
        # A browser command should never hold the single automation lock for
        # several minutes.  Long imports are handled by explicit async waits;
        # a short transport bound lets the controller recover and audit rather
        # than issuing an unknown duplicate click after a hung page transition.
        driver.command_executor.set_timeout(105)

    def _navigate(self, url: str) -> webdriver.Edge:
        driver = self._ensure_driver()
        # A debugger-attached Edge profile can contain a stale New Tab plus
        # one or more live KiotViet windows.  Selenium's first handle is not
        # guaranteed to be the sales window; navigating it would silently
        # strand the imported carts in the other window.  Prefer the matching
        # existing host and, for KiotViet sales, the window with the most
        # visible order tabs before issuing any navigation.
        self._select_existing_domain_window(driver, url)
        if self._is_kiot_sales_url(url):
            # A same-URL KiotViet SPA navigation can still replace the canvas;
            # do not reuse a readiness sample from the previous canvas.
            self._last_kiot_shell_url = None
            self._last_kiot_shell_signature = None
        target_host = str(urlsplit(url).netloc or "").casefold()
        current_url = str(driver.current_url or "")
        current_host = str(urlsplit(current_url).netloc or "").casefold()
        needs_navigation = current_host != target_host or not current_url.startswith(url)
        self._logger.info(
            "Edge navigation target=%s selected=%s needs_navigation=%s",
            url,
            current_url,
            needs_navigation,
        )
        if needs_navigation:
            driver.get(url)
            # The Edge session uses page_load_strategy="none" so ``get`` may
            # return while the old document is still active.  Waiting only on
            # document.readyState can therefore inspect the previous KiotViet
            # page and falsely report that Sopoka's date picker is missing.
            # Require both the target host and a ready document before any
            # page-specific selector is queried.
            WebDriverWait(driver, 25).until(
                lambda current: (
                    str(urlsplit(str(current.current_url or "")).netloc or "").casefold()
                    == target_host
                    and current.execute_script("return document.readyState") == "complete"
                )
            )
        else:
            self._wait_document_ready(driver)
        actual_url = str(driver.current_url or "")
        actual_host = str(urlsplit(actual_url).netloc or "").casefold()
        if actual_host != target_host:
            raise EdgeAutomationError(
                f"Edge khÃ´ng á»Ÿ Ä‘Ãºng trang Ä‘Ã­ch sau khi Ä‘iá»u hÆ°á»›ng. "
                f"Cáº§n {target_host}, Ä‘ang á»Ÿ {actual_host or actual_url or '(trá»‘ng)'} ."
            )
        self._logger.info("Edge navigation ready target=%s current=%s", url, actual_url)
        self._last_known_url = driver.current_url
        return driver

    def _select_existing_domain_window(self, driver: webdriver.Edge, url: str) -> None:
        """Focus the best already-open window for ``url`` without mutation."""

        try:
            parsed = urlsplit(url)
            host = str(parsed.netloc or "").casefold()
        except Exception:
            return
        if not host:
            return
        # Some lightweight test doubles (and a few transient attach states)
        # do not expose the multi-window properties.  In that case there is
        # nothing to select, so leave the current driver target untouched and
        # let the normal readiness/navigation path continue.
        try:
            handles_attr = getattr(driver, "window_handles", None)
            original = getattr(driver, "current_window_handle", None)
            if handles_attr is None:
                return
            handles = list(handles_attr)
        except (AttributeError, WebDriverException):
            return
        best_handle = None
        best_score = -1
        for handle in handles:
            try:
                driver.switch_to.window(handle)
                current = str(driver.current_url or "").casefold()
                if host not in current:
                    continue
                score = 10
                if self._is_kiot_sales_url(url) and "/sale/" in current:
                    labels = [
                        self._element_text(element)
                        for element in driver.find_elements(By.CSS_SELECTOR, "a.nav-link, li.nav-item")
                    ]
                    order_count = sum(
                        1 for label in labels if re.fullmatch(r"Đặt hàng\s+\d+", label, flags=re.IGNORECASE)
                    )
                    score += order_count * 10
                if url.startswith(KIOT_PRODUCTS_URL):
                    # Prefer an idle Products tab over a manually-open product
                    # editor. Creating a recovery SKU must never click behind
                    # a pre-existing editor (which could contain a user's
                    # unsaved work).
                    product_editors = driver.find_elements(By.CSS_SELECTOR, "#idSuggestProductCodeSearchTerm")
                    has_product_editor = any(element.is_displayed() for element in product_editors)
                    if has_product_editor:
                        score -= 100
                    elif product_editors:
                        # A just-cancelled Kendo editor can remain mounted but
                        # hidden during its animation. A clean list tab is a
                        # better recovery target when one is already open.
                        score -= 1
                if current.startswith(url.casefold()):
                    score += 2
                if score > best_score:
                    best_score = score
                    best_handle = handle
            except WebDriverException:
                continue
        try:
            if best_handle is not None:
                driver.switch_to.window(best_handle)
            elif not self._select_target_direct(driver, host):
                driver.switch_to.window(original)
        except WebDriverException:
            pass

    def _select_target_direct(self, driver: webdriver.Edge, host: str) -> bool:
        """Select a page target which EdgeDriver omitted from window_handles.

        A debugger-attached Edge can expose a live Sopoka page through the
        DevTools target list while EdgeDriver's window enumeration contains
        only the KiotViet page.  Selecting the exact target ID keeps the
        existing authenticated Sopoka tab and avoids blindly navigating the
        KiotViet sales canvas.
        """

        edge_options = driver.capabilities.get("ms:edgeOptions", {})
        address = edge_options.get("debuggerAddress") if isinstance(edge_options, dict) else None
        if not isinstance(address, str) or not address.strip():
            return False
        try:
            with urlopen(f"http://{address.strip()}/json/list", timeout=1.5) as response:
                targets = json.load(response)
            target = next(
                (
                    item
                    for item in targets
                    if isinstance(item, dict)
                    and item.get("type") == "page"
                    and str(urlsplit(str(item.get("url") or "")).netloc or "").casefold() == host
                ),
                None,
            )
            target_id = str(target.get("id") or "") if isinstance(target, dict) else ""
            if not target_id:
                return False
            driver.switch_to.window(target_id)
            self._last_known_url = str(driver.current_url or "")
            return True
        except (OSError, ValueError, WebDriverException):
            return False

    @staticmethod
    def _wait_document_ready(driver: webdriver.Edge) -> None:
        WebDriverWait(driver, 25).until(
            lambda current: current.execute_script("return document.readyState") == "complete"
        )

    def _assert_sopoka_login(self, driver: webdriver.Edge) -> None:
        text = (driver.title + " " + driver.find_element(By.TAG_NAME, "body").text[:2000]).casefold()
        try:
            report_visible = any(
                item.is_displayed()
                for item in driver.find_elements(
                    By.CSS_SELECTOR, "#reportrange, button.timK, button.xuatEx"
                )
            )
            password_visible = any(
                item.is_displayed()
                for item in driver.find_elements(By.CSS_SELECTOR, 'input[type="password"]')
            )
        except WebDriverException:
            report_visible = False
            password_visible = False
        if "login" in driver.current_url.casefold() or (password_visible and not report_visible):
            self._logger.info("Sopoka login required at %s", driver.current_url)
            raise EdgeLoginRequired(
                "Edge automation đã mở Sopoka. Hãy đăng nhập trong cửa sổ Edge vừa mở, "
                "sau đó bấm lại nút tải report. Không cần cung cấp mật khẩu cho tool."
            )
        if "đăng nhập" in text and not report_visible:
            raise EdgeLoginRequired(
                "Sopoka đang yêu cầu đăng nhập trong Edge automation. Hãy đăng nhập rồi chạy lại."
            )

    def _assert_kiot_login(self) -> None:
        driver = self._ensure_driver()
        text = (driver.title + " " + driver.find_element(By.TAG_NAME, "body").text[:2000]).casefold()
        if "login" in driver.current_url.casefold() or "đăng nhập" in text and "kiotviet" not in text:
            self._logger.info("KiotViet login required at %s", driver.current_url)
            raise EdgeLoginRequired(
                "Edge automation đã mở KiotViet. Hãy đăng nhập trong cửa sổ Edge vừa mở, "
                "sau đó bấm lại thao tác."
            )

    @staticmethod
    def _element_text(element: Any | None) -> str:
        """Return visible text without evaluating a page JavaScript snippet."""

        if element is None:
            return ""
        try:
            return re.sub(r"\s+", " ", str(element.text or "")).strip()
        except WebDriverException:
            return ""

    @staticmethod
    def _picker_label_key(value: str) -> str:
        """Return a stable key for Vietnamese date-picker labels.

        Sopoka has been served with both UTF-8 and legacy mojibake text in
        different Edge sessions (``Hôm nay`` vs ``Hôm nay``).  Matching the
        visible label by a folded key keeps the automation independent of the
        page's response encoding while still requiring an exact label.
        """

        candidates = [str(value or "")]
        try:
            candidates.append(candidates[0].encode("latin1").decode("utf-8"))
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
        keys = []
        for candidate in candidates:
            text = unicodedata.normalize("NFKD", candidate)
            text = "".join(char for char in text if not unicodedata.combining(char))
            # Vietnamese Đ/đ do not decompose under NFKD, so fold them
            # explicitly before comparing labels or placeholders.
            text = text.replace("Đ", "D").replace("đ", "d")
            keys.append(re.sub(r"\s+", " ", text).strip().casefold())
        # Prefer the repaired/shortest representation when mojibake is present.
        return min(keys, key=lambda item: (sum(item.count(mark) for mark in ("Ã", "Ä", "Â", "â")), len(item)))

    def _find_visible_css(self, selector: str) -> Any | None:
        try:
            return next(
                (
                    element
                    for element in self._ensure_driver().find_elements(By.CSS_SELECTOR, selector)
                    if element.is_displayed()
                ),
                None,
            )
        except WebDriverException:
            return None

    def _find_visible_xpath(self, xpath: str) -> Any | None:
        try:
            return next(
                (
                    element
                    for element in self._ensure_driver().find_elements(By.XPATH, xpath)
                    if element.is_displayed()
                ),
                None,
            )
        except WebDriverException:
            return None

    def _wait_for_visible_css(self, selector: str, label: str, *, timeout_seconds: int = 20) -> Any:
        try:
            return WebDriverWait(self._ensure_driver(), timeout_seconds).until(
                lambda _current: self._find_visible_css(selector)
            )
        except TimeoutException as exc:
            raise EdgeAutomationError(f"Không tìm thấy {label} trên Sopoka.") from exc

    def _wait_for_visible_xpath(self, xpath: str, label: str, *, timeout_seconds: int = 20) -> Any:
        try:
            return WebDriverWait(self._ensure_driver(), timeout_seconds).until(
                lambda _current: self._find_visible_xpath(xpath)
            )
        except TimeoutException as exc:
            raise EdgeAutomationError(f"Không tìm thấy {label} trên Sopoka.") from exc

    def _click_sopoka_element(self, element: Any, label: str) -> None:
        """Click with WebDriver pointer input, never ``execute_script(click)``.

        The report's date widget is sensitive to JavaScript-triggered click
        handlers.  Prefer Selenium's own element click; it is the normal
        browser click path that opened the live date picker reliably.  The
        W3C pointer route is retained only if Selenium cannot click it.
        """

        driver = self._ensure_driver()
        try:
            element.click()
            return
        except WebDriverException as click_error:
            try:
                ActionChains(driver).move_to_element(element).pause(0.08).click(element).perform()
                return
            except WebDriverException as pointer_error:
                raise EdgeAutomationError(f"Không thể bấm {label} trên Sopoka: {pointer_error}") from click_error

    def _dispatch_sopoka_pointer_click(self, element: Any, label: str) -> bool:
        """Send a real pointer click to the known Sopoka report target.

        Geometry is read from WebDriver only; this never evaluates page
        JavaScript and therefore cannot call Sopoka's patched ``Date`` API.
        The DevTools target is constrained to the sales-summary page, so an
        Office Viewer tab cannot receive the click by mistake.
        """

        if websocket is None:
            return False
        address = self._live_profile_debugger_address()
        if not address:
            return False
        try:
            rectangle = element.rect
            x = float(rectangle["x"]) + float(rectangle["width"]) / 2.0
            y = float(rectangle["y"]) + float(rectangle["height"]) / 2.0
            if x < 0 or y < 0:
                return False
            with urlopen(f"http://{address}/json", timeout=1.5) as response:
                targets = json.load(response)
            target = next(
                (
                    item
                    for item in targets
                    if isinstance(item, dict)
                    and item.get("type") == "page"
                    and str(item.get("url") or "").startswith(SOPKA_URL)
                    and str(item.get("webSocketDebuggerUrl") or "")
                ),
                None,
            )
            if target is None:
                return False
            socket = websocket.create_connection(
                str(target["webSocketDebuggerUrl"]),
                timeout=2.0,
                suppress_origin=True,
            )
        except Exception as exc:  # pragma: no cover - local Edge desktop transport
            self._logger.debug("Could not open Sopoka pointer channel for %s: %s", label, exc)
            return False

        def command(command_id: int, event_type: str) -> bool:
            try:
                socket.send(
                    json.dumps(
                        {
                            "id": command_id,
                            "method": "Input.dispatchMouseEvent",
                            "params": {
                                "type": event_type,
                                "x": x,
                                "y": y,
                                "button": "left",
                                "buttons": 1 if event_type == "mousePressed" else 0,
                                "clickCount": 1,
                            },
                        }
                    )
                )
                deadline = time.monotonic() + 2.0
                while time.monotonic() < deadline:
                    reply = json.loads(socket.recv())
                    if reply.get("id") == command_id:
                        return "error" not in reply
            except Exception as exc:  # pragma: no cover - local Edge desktop transport
                self._logger.debug("Sopoka pointer event failed for %s: %s", label, exc)
            return False

        try:
            return command(1, "mousePressed") and command(2, "mouseReleased")
        finally:
            try:
                socket.close()
            except Exception:
                pass

    def _click_sopoka_control(self, selector: str, label: str) -> Any:
        element = self._wait_for_visible_css(selector, label)
        if not element.is_enabled():
            raise EdgeAutomationError(f"{label} đang bị khóa trên Sopoka.")
        self._click_sopoka_element(element, label)
        return element

    def _visible_sopoka_date_picker(self) -> Any | None:
        """Return the actually open date-range picker, if one is open."""

        try:
            return next(
                (
                    picker
                    for picker in self._ensure_driver().find_elements(
                        By.CSS_SELECTOR, ".daterangepicker"
                    )
                    if picker.is_displayed()
                ),
                None,
            )
        except WebDriverException:
            return None

    def _visible_sopoka_today_option(self) -> Any | None:
        """Find the visible ``Hôm nay`` option inside the open picker."""

        picker = self._visible_sopoka_date_picker()
        if picker is None:
            return None
        try:
            for option in picker.find_elements(By.CSS_SELECTOR, "li,button,a,span"):
                if option.is_displayed() and self._picker_label_key(self._element_text(option)) == "hom nay":
                    return option
        except WebDriverException:
            return None
        return None

    def _visible_sopoka_apply_option(self) -> Any | None:
        picker = self._visible_sopoka_date_picker()
        if picker is None:
            return None
        try:
            return next(
                (
                    option
                    for option in picker.find_elements(
                        By.CSS_SELECTOR, ".applyBtn,button.applyBtn,input.applyBtn"
                    )
                    if option.is_displayed() and option.is_enabled()
                ),
                None,
            )
        except WebDriverException:
            return None

    def _wait_for_open_sopoka_date_picker(self, timeout_seconds: float) -> Any | None:
        try:
            return WebDriverWait(self._ensure_driver(), timeout_seconds).until(
                lambda _current: self._visible_sopoka_date_picker()
            )
        except TimeoutException:
            return None

    def _open_sopoka_date_picker(self) -> Any:
        """Ensure the report date picker is open before selecting a preset.

        A user can leave this toggle open between runs.  Clicking it blindly
        would close the list and make the subsequent ``Hôm nay`` lookup time
        out.  We therefore inspect normal Selenium visibility first, then use
        bounded pointer retries only when it is genuinely closed.
        """

        already_open = self._visible_sopoka_date_picker()
        if already_open is not None:
            return already_open

        range_element = self._wait_for_visible_css(
            "#reportrange", "bộ chọn ngày của báo cáo tổng hợp bán hàng"
        )
        self._click_sopoka_element(range_element, "bộ chọn ngày của báo cáo tổng hợp bán hàng")
        picker = self._wait_for_open_sopoka_date_picker(3.0)
        if picker is not None:
            return picker

        # A second Selenium pointer click is useful when an overlay has just
        # finished animating.  It is still regular browser input.
        try:
            ActionChains(self._ensure_driver()).move_to_element(range_element).pause(0.1).click(
                range_element
            ).perform()
        except WebDriverException:
            pass
        picker = self._wait_for_open_sopoka_date_picker(3.0)
        if picker is not None:
            return picker

        # Last resort for an EdgeDriver action that is acknowledged but does
        # not reach the visible page: CDP sends a physical pointer event to
        # the already-known Sopoka tab.  No page script is evaluated.
        if self._dispatch_sopoka_pointer_click(
            range_element, "bộ chọn ngày của báo cáo tổng hợp bán hàng"
        ):
            picker = self._wait_for_open_sopoka_date_picker(3.0)
            if picker is not None:
                return picker

        raise EdgeAutomationError(
            "Sopoka chưa mở được bộ chọn ngày bằng thao tác chuột an toàn."
        )

    def _select_sopoka_today(self, expected_date: Any) -> None:
        """Select and verify the *Hôm nay* date range using only DOM/UI calls."""

        self._open_sopoka_date_picker()
        try:
            today = WebDriverWait(self._ensure_driver(), 8).until(
                lambda _current: self._visible_sopoka_today_option()
            )
        except TimeoutException as exc:
            raise EdgeAutomationError(
                "Không tìm thấy mục Hôm nay trong bộ chọn ngày đang mở trên Sopoka."
            ) from exc
        self._click_sopoka_element(today, "mục Hôm nay")

        # Some versions of the picker apply the preset immediately.  When an
        # Apply control remains, use it; otherwise verification below is the
        # source of truth.
        apply = self._visible_sopoka_apply_option()
        if apply is not None:
            self._click_sopoka_element(apply, "nút Áp dụng")

        expected = expected_date.strftime("%d/%m/%Y")
        range_element = self._wait_for_visible_css("#reportrange", "ô khoảng ngày")
        try:
            WebDriverWait(self._ensure_driver(), 20).until(
                lambda _current: self._sopoka_range_matches_today(range_element, expected)
            )
        except TimeoutException as exc:
            actual = self._element_text(range_element) or "(trống)"
            raise EdgeAutomationError(
                "Sopoka chưa hiển thị đúng khoảng ngày hôm nay. "
                f"Đang thấy: {actual}."
            ) from exc

    def _select_sopoka_date(self, expected_date: date) -> None:
        """Select an arbitrary one-day range through the visible picker."""

        picker = self._open_sopoka_date_picker()
        custom = None
        try:
            custom = next(
                (
                    option
                    for option in picker.find_elements(By.CSS_SELECTOR, "li,button,a")
                    if option.is_displayed() and self._element_text(option) == "Tùy chỉnh"
                ),
                None,
            )
        except WebDriverException:
            custom = None
        if custom is not None:
            self._click_sopoka_element(custom, "mục Tùy chỉnh")

        try:
            inputs = [
                element
                for element in picker.find_elements(By.CSS_SELECTOR, "input")
                if element.is_displayed() and element.is_enabled()
            ]
        except WebDriverException as exc:
            raise EdgeAutomationError("Không đọc được ô ngày tùy chỉnh trên Sopoka.") from exc
        if len(inputs) < 2:
            raise EdgeAutomationError("Sopoka không hiển thị đủ ô Từ/Đến cho ngày đã chọn.")

        def input_key(element: Any) -> str:
            return " ".join(
                str(element.get_attribute(name) or "")
                for name in ("name", "id", "placeholder", "class")
            ).casefold()

        start = next(
            (item for item in inputs if any(token in input_key(item) for token in ("start", "from", "từ"))),
            inputs[0],
        )
        end = next(
            (
                item
                for item in inputs
                if item is not start and any(token in input_key(item) for token in ("end", "to", "đến"))
            ),
            inputs[1],
        )
        value = expected_date.strftime("%d/%m/%Y")
        for element in (start, end):
            element.click()
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(value)
        apply = self._visible_sopoka_apply_option()
        if apply is not None:
            self._click_sopoka_element(apply, "nút Áp dụng")

        expected = expected_date.strftime("%d/%m/%Y")
        range_element = self._wait_for_visible_css("#reportrange", "ô khoảng ngày")
        try:
            WebDriverWait(self._ensure_driver(), 20).until(
                lambda _current: self._sopoka_range_matches_today(range_element, expected)
            )
        except TimeoutException as exc:
            actual = self._element_text(range_element) or "(trống)"
            raise EdgeAutomationError(
                f"Sopoka chưa hiển thị đúng ngày đã chọn. Đang thấy: {actual}."
            ) from exc

    def _sopoka_range_matches_today(self, element: Any, expected: str) -> bool:
        dates = re.findall(r"\d{2}/\d{2}/\d{4}", self._element_text(element))
        return len(dates) >= 2 and dates[0] == expected and dates[1] == expected

    def _sopoka_result_signature(self) -> tuple[str, ...]:
        """A bounded, read-only signature of visible report rows."""

        driver = self._ensure_driver()
        try:
            rows = [
                row
                for row in driver.find_elements(
                    By.CSS_SELECTOR, "table.reportCustomer tbody tr"
                )
                if row.is_displayed()
            ]
        except WebDriverException:
            return ()
        return tuple(
            re.sub(r"\s+", " ", self._element_text(row)).casefold()[:250]
            for row in rows[:12]
        )

    def _sopoka_is_busy(self) -> bool:
        # jQuery's internal counters require injected JavaScript.  A visible
        # page overlay is sufficient here and keeps the fragile Sopoka page
        # script untouched.
        selectors = ".loading,.loader,.spinner,.fa-spinner,[aria-busy='true']"
        return self._find_visible_css(selectors) is not None

    def _sopoka_has_no_data_message(self) -> bool:
        """Return whether Sopoka visibly reports an empty result set."""

        try:
            body = self._ensure_driver().find_element(By.TAG_NAME, "body")
            text = self._element_text(body).casefold()
        except WebDriverException:
            return False
        return any(
            marker in text
            for marker in (
                "không có dữ liệu",
                "không tìm thấy dữ liệu",
                "không có kết quả",
                "trên tổng số 0 bản ghi",
                "tổng số 0 bản ghi",
            )
        )

    def _wait_for_sopoka_results(self, *, timeout_seconds: float = 75) -> bool:
        """Wait until Sopoka has finished rendering a searched report.

        An empty report is valid, so we require a short stable interval and
        explicitly detect Sopoka's visible no-data banner instead of trying
        to export a file that does not exist.
        """

        started = time.monotonic()
        deadline = started + timeout_seconds
        last = self._sopoka_result_signature()
        stable_passes = 0
        while time.monotonic() < deadline:
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng chờ kết quả Sopoka.")
            current = self._sopoka_result_signature()
            export_button = self._find_visible_css("button.xuatEx")
            try:
                table_text = self._element_text(
                    self._find_visible_css("table.reportCustomer")
                ).casefold()
            except WebDriverException:
                table_text = ""
            still_prompting = "bấm tìm kiếm để xem báo cáo" in table_text
            stable_passes = stable_passes + 1 if current == last and not self._sopoka_is_busy() else 0
            last = current
            if time.monotonic() - started >= 2.5 and stable_passes >= 5:
                if self._sopoka_has_no_data_message():
                    return False
                if (
                    export_button is not None
                    and export_button.is_enabled()
                    and not still_prompting
                ):
                    return True
            time.sleep(0.35)
        raise EdgeAutomationError("Sopoka chưa trả xong kết quả tìm kiếm nên chưa thể xuất Excel an toàn.")

    def _wait_for_report_download(
        self, before: dict[str, int], *, timeout_seconds: float = 150
    ) -> Path:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng chờ file report.")
            candidates = [
                item
                for item in self.config.edge_download_dir.glob("ReportOrderSaleSummary*")
                if item.is_file()
                and item.suffix.casefold() in SUPPORTED_REPORT_EXTENSIONS
                and item.stat().st_size > 0
                and before.get(str(item), -1) != item.stat().st_mtime_ns
            ]
            pending = list(self.config.edge_download_dir.glob("*.crdownload"))
            if candidates and not pending:
                newest = max(candidates, key=lambda item: item.stat().st_mtime_ns)
                first_size = newest.stat().st_size
                time.sleep(0.8)
                if newest.exists() and newest.stat().st_size == first_size:
                    return newest.resolve()
            time.sleep(0.35)
        raise EdgeAutomationError(
            f"Edge không tải xong report sau {int(timeout_seconds)} giây. "
            f"Kiểm tra Downloads: {self.config.edge_download_dir}"
        )

    @staticmethod
    def _validated_sopoka_export_url(source_url: str) -> str | None:
        """Return a safe Sopoka report endpoint or ``None``.

        The URL originates from browser state and is never assembled from a
        filename or logged.  Keep the accepted host and filename narrow so a
        compromised/non-report Office tab cannot turn this into a generic
        authenticated downloader.
        """

        try:
            source_url = str(source_url).strip()
            source = urlsplit(source_url)
            filename = unquote(source.path.rsplit("/", 1)[-1])
        except (TypeError, ValueError):
            return None
        if (
            source.scheme.casefold() not in {"http", "https"}
            or (source.hostname or "").casefold() != "dms.sopoka.com.vn"
            or source.username is not None
            or source.password is not None
            or not re.fullmatch(
                r"ReportOrderSaleSummary_[A-Za-z0-9._-]+\.(?:xls|xlsx|xlsm|xltx|xltm)",
                filename,
                flags=re.IGNORECASE,
            )
        ):
            return None
        return source_url

    @classmethod
    def _office_viewer_source_url(cls, viewer_url: str) -> str | None:
        """Return one strictly validated Sopoka ``src`` URL from Office Viewer.

        Office Viewer receives the real export endpoint as a percent-encoded
        ``src`` query parameter.  Treat it as untrusted browser state until
        its viewer origin, Sopoka origin, and filename all match this exact
        workflow.  The caller deliberately never logs the returned URL.
        """

        try:
            viewer = urlsplit(viewer_url)
            if (
                viewer.scheme.casefold() != "https"
                or (viewer.hostname or "").casefold() != "view.officeapps.live.com"
                or not viewer.path.casefold().startswith("/op/view")
            ):
                return None
            values = parse_qs(viewer.query, keep_blank_values=False).get("src", [])
            if len(values) != 1:
                return None
            source_url = str(values[0]).strip()
            # A viewer URL can contain one extra encoded layer depending on
            # whether Edge has normalized the address bar.  Decode at most
            # twice, rather than turning arbitrary input into a new URL.
            for _unused in range(2):
                decoded = unquote(source_url)
                if decoded == source_url:
                    break
                source_url = decoded
        except (TypeError, ValueError):
            return None
        return cls._validated_sopoka_export_url(source_url)

    def _sopoka_office_export_urls(self) -> set[str]:
        """Read validated report sources from live Office Viewer targets only."""

        address = self._live_profile_debugger_address()
        if not address:
            return set()
        try:
            with urlopen(f"http://{address}/json", timeout=1.5) as response:
                targets = json.load(response)
        except Exception as exc:  # pragma: no cover - local Edge availability
            self._logger.debug("Could not list Office Viewer targets: %s", exc)
            return set()
        if not isinstance(targets, list):
            return set()
        sources: set[str] = set()
        for target in targets:
            if not isinstance(target, dict) or target.get("type") != "page":
                continue
            source = self._office_viewer_source_url(str(target.get("url") or ""))
            if source:
                sources.add(source)
        return sources

    def _wait_for_sopoka_office_export_url(
        self,
        existing_urls: set[str],
        *,
        timeout_seconds: float = 20,
    ) -> str | None:
        """Wait briefly for the Office Viewer target created by this export."""

        deadline = time.monotonic() + max(0.1, timeout_seconds)
        while time.monotonic() < deadline:
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng chờ file report.")
            fresh = self._sopoka_office_export_urls().difference(existing_urls)
            if fresh:
                # Sopoka's generated filename contains a millisecond marker.
                # Sort on that filename part, not the opaque export folder.
                def source_marker(value: str) -> tuple[int, str]:
                    match = re.search(
                        r"_(\d+)\.(?:xls|xlsx|xlsm|xltx|xltm)$",
                        urlsplit(value).path,
                        flags=re.IGNORECASE,
                    )
                    return (int(match.group(1)) if match else -1, value)

                return max(fresh, key=source_marker)
            time.sleep(0.25)
        return None

    @staticmethod
    def _cookie_header_for_sopoka_url(
        cookies: list[dict[str, Any]], source_url: str
    ) -> str:
        """Build a request Cookie header without ever persisting cookie values."""

        source = urlsplit(source_url)
        host = (source.hostname or "").casefold()
        secure_request = source.scheme.casefold() == "https"
        now = time.time()
        pairs: list[str] = []
        seen: set[str] = set()
        for cookie in cookies:
            if not isinstance(cookie, dict):
                continue
            name = str(cookie.get("name") or "")
            value = str(cookie.get("value") or "")
            domain = str(cookie.get("domain") or "").lstrip(".").casefold()
            if (
                not name
                or any(character in name for character in "=;\r\n\t ")
                or any(character in value for character in "\r\n")
                or not domain
                or not (host == domain or host.endswith(f".{domain}"))
                or (bool(cookie.get("secure")) and not secure_request)
            ):
                continue
            expiry = cookie.get("expires", cookie.get("expiry"))
            try:
                if expiry is not None and float(expiry) >= 0 and float(expiry) < now:
                    continue
            except (TypeError, ValueError):
                continue
            pair = f"{name}={value}"
            if pair not in seen:
                pairs.append(pair)
                seen.add(pair)
        return "; ".join(pairs)

    def _sopoka_download_cookie_header(self, source_url: str) -> str:
        """Read only cookies applicable to the verified Sopoka export URL."""

        driver = self._ensure_driver()
        cookies: list[dict[str, Any]] = []
        try:
            response = driver.execute_cdp_cmd("Network.getCookies", {"urls": [source_url]})
            if isinstance(response, dict) and isinstance(response.get("cookies"), list):
                cookies = response["cookies"]
        except (AttributeError, WebDriverException):
            pass
        if not cookies:
            try:
                response = driver.execute_cdp_cmd("Network.getAllCookies", {})
                if isinstance(response, dict) and isinstance(response.get("cookies"), list):
                    cookies = response["cookies"]
            except (AttributeError, WebDriverException):
                pass
        if not cookies:
            try:
                current = driver.get_cookies()
                if isinstance(current, list):
                    cookies = current
            except WebDriverException:
                pass
        return self._cookie_header_for_sopoka_url(cookies, source_url)

    def _direct_sopoka_report_download(self, source_url: str) -> Path:
        """Download the validated Office ``src`` with in-memory session cookies."""

        validated = self._validated_sopoka_export_url(source_url)
        if validated is None:
            raise EdgeAutomationError("Office Viewer không trả về đúng file report Sopoka cần tải.")
        source_url = validated
        source = urlsplit(source_url)
        filename = unquote(source.path.rsplit("/", 1)[-1])
        target = self.config.edge_download_dir / filename
        if target.exists():
            target = self.config.edge_download_dir / (
                f"{target.stem}_direct_{int(time.time() * 1000)}{target.suffix}"
            )
        while target.exists():
            target = self.config.edge_download_dir / (
                f"{target.stem}_1{target.suffix}"
            )

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Edg/136.0 Safari/537.36",
            "Referer": SOPKA_URL,
        }
        cookie_header = self._sopoka_download_cookie_header(source_url)
        if cookie_header:
            headers["Cookie"] = cookie_header
        temporary_path: Path | None = None
        try:
            request = Request(source_url, headers=headers, method="GET")
            with urlopen(request, timeout=25) as response:
                final = urlsplit(str(response.geturl() or source_url))
                if (final.hostname or "").casefold() != "dms.sopoka.com.vn":
                    raise EdgeAutomationError("Sopoka chuyển file report sang máy chủ không được xác thực.")
                content_type = str(response.headers.get("Content-Type") or "").casefold()
                if "text/html" in content_type:
                    raise EdgeAutomationError("Sopoka trả về trang HTML thay vì file Excel report.")
                with tempfile.NamedTemporaryFile(
                    mode="wb",
                    prefix=".orderbridge-",
                    suffix=".part",
                    dir=self.config.edge_download_dir,
                    delete=False,
                ) as temporary:
                    temporary_path = Path(temporary.name)
                    copied = 0
                    while True:
                        chunk = response.read(64 * 1024)
                        if not chunk:
                            break
                        copied += len(chunk)
                        if copied > 100 * 1024 * 1024:
                            raise EdgeAutomationError("File report Sopoka vượt giới hạn an toàn 100 MB.")
                        temporary.write(chunk)
            if copied <= 0:
                raise EdgeAutomationError("Sopoka trả về file report rỗng.")
            os.replace(temporary_path, target)
            temporary_path = None
            self._logger.info("Downloaded verified Sopoka report through Office Viewer (%s bytes).", copied)
            return target.resolve()
        except EdgeAutomationError:
            raise
        except Exception as exc:
            raise EdgeAutomationError(
                f"Không thể tải trực tiếp file report từ Office Viewer ({type(exc).__name__})."
            ) from exc
        finally:
            if temporary_path is not None:
                try:
                    temporary_path.unlink(missing_ok=True)
                except OSError:
                    pass

    def _wait_for_sopoka_download(
        self,
        before: dict[str, int],
        started_at: datetime,
        *,
        existing_office_export_urls: set[str] | None = None,
        export_button: Any | None = None,
    ) -> Path:
        """Download the Office viewer export, including Edge's browser-chrome prompt.

        Sopoka opens the Excel report in Edge's Office viewer.  The visible ``Download
        file`` control belongs to Edge's native chrome, so Selenium cannot see
        it as page DOM.  A short direct-download wait keeps automatic-download
        Edge installations fast; otherwise a small UI Automation helper clicks
        only that named control in this dedicated Edge process.
        """

        try:
            return self._wait_for_report_download(before, timeout_seconds=4)
        except EdgeAutomationError:
            pass

        existing_urls = existing_office_export_urls or set()
        # Office Viewer is a separate Edge target.  On a cold profile (or
        # while Edge is restoring several tabs) that target can take longer
        # than the old eight-second probe to appear.  Falling through to the
        # native UIA helper too early produced a misleading "Download file
        # button" error even though the viewer was still being created.  Keep
        # polling the validated target before touching browser chrome.
        source_url = self._wait_for_sopoka_office_export_url(
            existing_urls, timeout_seconds=25
        )
        if source_url:
            try:
                return self._direct_sopoka_report_download(source_url)
            except EdgeAutomationError as exc:
                # Do not include the source URL or any session data in logs.
                # The native prompt remains a narrowly-scoped final fallback.
                self._logger.warning("Verified Office Viewer download fallback failed: %s", exc)

        # A Selenium click can occasionally be acknowledged without firing
        # Sopoka's ``ket_xuat`` handler.  We only retry when eight seconds
        # produced neither an automatic file nor a *new* validated Office
        # Viewer target.  That makes one physical pointer retry safer than a
        # blind second WebDriver click, which could otherwise create two
        # exports.
        if source_url is None and export_button is not None:
            if self._dispatch_sopoka_pointer_click(export_button, "nút Xuất Excel"):
                self._logger.info(
                    "Retried Sopoka export with a verified physical pointer after no download target."
                )
                try:
                    return self._wait_for_report_download(before, timeout_seconds=4)
                except EdgeAutomationError:
                    pass
                source_url = self._wait_for_sopoka_office_export_url(
                    existing_urls, timeout_seconds=30
                )
                if source_url:
                    try:
                        return self._direct_sopoka_report_download(source_url)
                    except EdgeAutomationError as exc:
                        self._logger.warning(
                            "Verified Office Viewer download after pointer retry failed: %s", exc
                        )

        # Give a slow Office target one final short grace period before the
        # native helper.  This is intentionally after the single verified
        # pointer retry, so no duplicate export click can be generated.
        if source_url is None:
            source_url = self._wait_for_sopoka_office_export_url(
                existing_urls, timeout_seconds=15
            )
            if source_url:
                try:
                    return self._direct_sopoka_report_download(source_url)
                except EdgeAutomationError as exc:
                    self._logger.warning(
                        "Verified Office Viewer final direct download failed: %s", exc
                    )

        helper = self.config.scripts_dir / "download_edge_file.ps1"
        if not helper.is_file():
            raise EdgeAutomationError(
                f"Không tìm thấy helper bấm Download file của Edge: {helper}"
            )
        capabilities = self._ensure_driver().capabilities
        browser_process_id = self._session_browser_process_id(capabilities)
        if browser_process_id is None:
            raise EdgeAutomationError(
                "Không đọc được tiến trình Edge của profile automation để bấm Download file."
            )

        command = [
            str(self.config.powershell_path),
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(helper),
            "-DownloadDirectory",
            str(self.config.edge_download_dir),
            "-ExpectedPrefix",
            "ReportOrderSaleSummary",
            "-NotBeforeUtc",
            started_at.astimezone(timezone.utc).isoformat(),
            "-BrowserProcessId",
            str(browser_process_id),
            "-TimeoutSeconds",
            "35",
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=50,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise EdgeAutomationError("Quá thời gian chờ Edge bấm Download file của report Sopoka.") from exc

        payload: dict[str, Any] | None = None
        for line in reversed(completed.stdout.splitlines()):
            try:
                value = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(value, dict):
                payload = value
                break
        if completed.returncode != 0 or not payload or payload.get("success") is not True:
            detail = (completed.stderr or completed.stdout or "không có chi tiết").strip()
            raise EdgeAutomationError(
                "Edge không tải được file report sau khi mở Office viewer. "
                f"Chi tiết: {detail[-800:]}"
            )
        downloaded = Path(str(payload.get("path") or "")).resolve()
        if (
            not downloaded.is_file()
            or downloaded.suffix.casefold() not in SUPPORTED_REPORT_EXTENSIONS
            or not downloaded.name.casefold().startswith("reportordersalesummary")
        ):
            raise EdgeAutomationError("Edge trả về file tải xuống không đúng report tổng hợp bán hàng.")
        return downloaded

    def _verify_exported_report(self, path: Path, expected_date: Any) -> dict[str, Any]:
        """Reject a stale/wrong export before it can reach SQLite or KiotViet."""

        from .report_parser import ReportParseError, parse_report

        try:
            report = parse_report(
                path,
                expected_date=expected_date,
                max_orders_per_batch=self.config.max_orders_per_batch,
            )
        except (ReportParseError, OSError) as exc:
            raise EdgeAutomationError(
                "File Sopoka vừa tải không phải báo cáo tổng hợp bán hàng đúng ngày hôm nay. "
                f"Tool đã dừng trước khi ghi SQLite hoặc upload KiotViet. Chi tiết: {exc}"
            ) from exc
        return {"order_count": len(report.orders)}

    def _download_snapshot(self) -> dict[str, int]:
        return {
            str(item): item.stat().st_mtime_ns
            for item in self.config.edge_download_dir.glob("ReportOrderSaleSummary*")
            if item.suffix.casefold() in SUPPORTED_REPORT_EXTENSIONS
            if item.is_file()
        }

    def _kiot_status(self) -> dict[str, Any]:
        # ``kiotInspectTask`` is read-only. Unlike actions that click/import
        # data, it is synchronous, so use WebDriver's synchronous script
        # endpoint. This avoids an async callback being lost while KiotViet
        # replaces its initial loading shell.
        self._ensure_kiot_shell_ready()
        try:
            result = self._run_sync_task("kiotInspectTask", self.config.expected_branch)
        except (JavascriptException, TimeoutException) as first_error:
            # Retrying this one read-only inspection cannot create an order,
            # edit a note, or close a tab. A route transition may keep the
            # same URL while recreating the SPA DOM, so force a new settle.
            self._logger.warning(
                "KiotViet inspection was interrupted while the sales shell loaded; "
                "waiting and retrying once: %s",
                first_error,
            )
            self._ensure_kiot_shell_ready(force=True)
            try:
                result = self._run_sync_task("kiotInspectTask", self.config.expected_branch)
            except (JavascriptException, TimeoutException) as retry_error:
                raise EdgeAutomationError(
                    "KiotViet chưa nạp xong giao diện bán hàng nên không thể kiểm tra an toàn. "
                    "Không có tab hoặc đơn hàng nào bị thay đổi; hãy thử lại khi màn hình KiotViet hiện đầy đủ."
                ) from retry_error
        self._task_result_or_raise(result)
        return result

    def _ensure_kiot_shell_ready(self, *, force: bool = False) -> None:
        """Wait for an interactive KiotViet shell without executing page code.

        ``document.readyState`` becomes ``complete`` before KiotViet's sales
        SPA has hydrated. Calling an inspection in that blank spinner state
        can race the route transition. This probe uses normal WebDriver
        element/text access only and waits for stable sales semantics.
        """

        snapshot = self._kiot_shell_snapshot()
        current_url = snapshot.get("url")
        signature = self._kiot_shell_signature(snapshot)
        if (
            not force
            and snapshot.get("ready")
            and isinstance(current_url, str)
            and current_url == self._last_kiot_shell_url
            and signature == self._last_kiot_shell_signature
        ):
            return

        # KiotViet can spend 30–40 seconds hydrating the first sales canvas
        # after a browser/profile restart.  This is still read-only; waiting
        # here is safer than returning a false "not ready" and leaving stale
        # tabs uncleared before import.
        deadline = time.monotonic() + 55.0
        consecutive_ready = 0
        last_snapshot = snapshot
        while time.monotonic() < deadline:
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng chờ KiotViet tải giao diện.")
            snapshot = self._kiot_shell_snapshot()
            last_snapshot = snapshot
            if snapshot.get("ready"):
                consecutive_ready += 1
                if consecutive_ready >= 2:
                    current_url = snapshot.get("url")
                    self._last_kiot_shell_url = (
                        str(current_url) if isinstance(current_url, str) else None
                    )
                    self._last_kiot_shell_signature = self._kiot_shell_signature(snapshot)
                    return
            else:
                consecutive_ready = 0
            time.sleep(0.25)

        raise EdgeAutomationError(
            "KiotViet chưa nạp xong giao diện bán hàng sau 55 giây. "
            "Tool đã dừng trước khi thay đổi tab hoặc đơn hàng. "
            f"Trạng thái tải: branch={bool(last_snapshot.get('branchVisible'))}, "
            f"tab={bool(last_snapshot.get('tabVisible'))}, "
            f"ghi_chú={bool(last_snapshot.get('noteVisible'))}."
        )

    @staticmethod
    def _is_kiot_sales_url(url: Any) -> bool:
        """Return true only for the interactive KiotViet sales application."""

        try:
            parts = urlsplit(str(url or ""))
        except (TypeError, ValueError):
            return False
        return (
            parts.hostname == "nongsanlamhong.kiotviet.vn"
            and parts.path.casefold().startswith("/sale/")
        )

    @staticmethod
    def _kiot_shell_signature(snapshot: dict[str, Any]) -> tuple[Any, ...]:
        """Small DOM-state marker so a replaced SPA shell is never reused."""

        return (
            snapshot.get("url"),
            bool(snapshot.get("salesUrl")),
            bool(snapshot.get("tabVisible")),
            bool(snapshot.get("noteVisible")),
        )

    def _kiot_shell_snapshot(self) -> dict[str, Any]:
        """Return a small read-only signal for KiotViet SPA hydration."""

        driver = self._ensure_driver()
        current_url: str | None
        try:
            current_url = str(driver.current_url or "")
        except WebDriverException:
            current_url = None

        text = ""
        note_visible = False
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            text = self._element_text(body)
            note_visible = any(
                element.is_displayed()
                and "ghi chu don hang" in self._picker_label_key(
                    " ".join(
                        str(element.get_attribute(attribute) or "")
                        for attribute in ("placeholder", "data-placeholder", "aria-label", "name", "id")
                    )
                )
                for element in driver.find_elements(
                    By.CSS_SELECTOR, 'input, textarea, [contenteditable="true"]'
                )
            )
        except WebDriverException:
            pass

        normalized = self._picker_label_key(text)
        expected_branch = self._picker_label_key(self.config.expected_branch)
        branch_visible = bool(expected_branch and expected_branch in normalized)
        body_tab_visible = bool(
            re.search(r"(?:dat\s+hang|hoa\s+don)(?:\s+\d+)?", normalized)
        )
        # A management menu also contains the text "Đặt hàng".  Treat the
        # shell as ready only when a real, numbered sales tab or a note field
        # is visible on the sales application itself.
        sales_url = self._is_kiot_sales_url(current_url)
        tab_visible = False
        try:
            tab_visible = any(
                element.is_displayed()
                and bool(
                    re.fullmatch(
                        r"(?:dat hang|hoa don)\s*\d+",
                        self._picker_label_key(self._element_text(element)),
                        flags=re.IGNORECASE,
                    )
                )
                for element in driver.find_elements(
                    By.XPATH,
                    "//*[self::button or self::a or self::li or self::div or self::span "
                    "or @role='tab' or @role='button']",
                )
            )
        except WebDriverException:
            tab_visible = False
        return {
            # A numbered tab can appear before its Angular order canvas has
            # rendered the note field.  Treat it as loading until both are
            # present, otherwise the next import can race a blank cart.
            "ready": bool(sales_url and tab_visible and note_visible),
            "url": current_url,
            "salesUrl": sales_url,
            "branchVisible": branch_visible,
            "tabVisible": tab_visible,
            "bodyTabVisible": body_tab_visible,
            "noteVisible": note_visible,
        }

    def _run_task(self, function_name: str, *args: Any, timeout_seconds: int = 30) -> dict[str, Any]:
        driver = self._ensure_driver()
        source = self._task_source(function_name)
        driver.set_script_timeout(timeout_seconds)
        # Selenium's HTTP client has its own read deadline independent of the
        # browser JavaScript deadline. The former fixed 105s timeout expired
        # while KiotViet was legitimately verifying 19 imported carts.
        try:
            driver.command_executor.set_timeout(max(105, int(timeout_seconds) + 30))
        except Exception:
            pass
        wrapper = f"""
{source}
const done = arguments[arguments.length - 1];
Promise.resolve({function_name}(...Array.prototype.slice.call(arguments, 0, -1)))
  .then((result) => done({{ok: true, result: result || {{}}}}))
  .catch((error) => done({{ok: false, error: String(error && (error.stack || error.message) || error)}}));
"""
        response = driver.execute_async_script(wrapper, *args)
        if not isinstance(response, dict) or response.get("ok") is not True:
            detail = response.get("error") if isinstance(response, dict) else repr(response)
            raise EdgeAutomationError(f"JavaScript Edge task {function_name} thất bại: {detail}")
        result = response.get("result")
        if not isinstance(result, dict):
            raise EdgeAutomationError(f"JavaScript Edge task {function_name} trả về dữ liệu không hợp lệ.")
        return result

    def _run_sync_task(self, function_name: str, *args: Any) -> dict[str, Any]:
        """Execute a synchronous read-only browser task without a callback."""

        driver = self._ensure_driver()
        source = self._task_source(function_name)
        wrapper = f"""
{source}
return {function_name}(...arguments);
"""
        result = driver.execute_script(wrapper, *args)
        if not isinstance(result, dict):
            raise EdgeAutomationError(
                f"JavaScript Edge task {function_name} returned invalid data."
            )
        return result

    def _task_source(self, function_name: str) -> str:
        cached = self._task_source_cache.get(function_name)
        if cached:
            return cached
        source_path = self.config.app_root / "chrome-extension" / "background.js"
        source = source_path.read_text(encoding="utf-8")
        pattern = re.compile(rf"(?:async\s+)?function\s+{re.escape(function_name)}\s*\(")
        match = pattern.search(source)
        if not match:
            raise EdgeAutomationError(f"Không tìm thấy browser task {function_name}.")
        brace = source.find("{", match.end())
        if brace < 0:
            raise EdgeAutomationError(f"Browser task {function_name} không có thân hàm.")
        depth = 0
        quote: str | None = None
        line_comment = False
        block_comment = False
        escaped = False
        index = brace
        while index < len(source):
            char = source[index]
            following = source[index + 1] if index + 1 < len(source) else ""
            if line_comment:
                if char in "\r\n":
                    line_comment = False
            elif block_comment:
                if char == "*" and following == "/":
                    block_comment = False
                    index += 1
            elif quote:
                if escaped:
                    escaped = False
                elif char == "\\":
                    escaped = True
                elif char == quote:
                    quote = None
            elif char in "'\"`":
                quote = char
            elif char == "/" and following == "/":
                line_comment = True
                index += 1
            elif char == "/" and following == "*":
                block_comment = True
                index += 1
            elif char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    extracted = source[match.start() : index + 1]
                    self._task_source_cache[function_name] = extracted
                    return extracted
            index += 1
        raise EdgeAutomationError(f"Không thể đọc trọn browser task {function_name}.")

    @staticmethod
    def _task_result_or_raise(result: dict[str, Any]) -> None:
        if result.get("error"):
            raise EdgeAutomationError(str(result["error"]))

    @staticmethod
    def _is_safe_submission_context_error(error: EdgeAutomationError) -> bool:
        """Identify a pre-click shell/branch mismatch safe to recover once."""

        text = str(error or "").casefold()
        return any(
            marker in text
            for marker in (
                "not on the required branch",
                "not in order mode",
                "still has an invoice tab",
                "missing đặt hàng",
                "missing đặthàng",
            )
        )

    def _recover_submission_context(self, plan: list[dict[str, Any]]) -> dict[str, Any]:
        """Recover a repainting KiotViet shell without clicking an order."""

        try:
            status = self._kiot_status()
        except EdgeAutomationError:
            status = self.prepare_kiot()
        if not status.get("branchOk") or not status.get("orderMode") or status.get("invoiceTabs"):
            status = self.prepare_kiot()
        if not status.get("branchOk") or not status.get("orderMode") or status.get("invoiceTabs"):
            raise EdgeAutomationError(
                "KiotViet chưa trở lại đúng chi nhánh/Đặt hàng sau khi chuyển tab; chưa bấm đơn."
            )
        available = {str(label).strip() for label in status.get("orderTabs", [])}
        if not available:
            raise EdgeAutomationError("KiotViet has no order tabs to continue.")
        return status

    def _set_progress(
        self,
        *,
        active: bool,
        phase: str,
        message: str,
        current: int | None = None,
        total: int | None = None,
        batch: str | None = None,
        order_code: str | None = None,
    ) -> None:
        """Store a nonblocking snapshot consumed by ``/api/browser/status``."""

        with self._progress_lock:
            next_progress = dict(self._progress)
            next_progress.update(
                {
                    "active": bool(active),
                    "phase": str(phase),
                    "message": str(message),
                    "batch": batch,
                    "orderCode": order_code,
                    "updatedAt": datetime.now(timezone.utc).isoformat(),
                }
            )
            if current is not None:
                next_progress["current"] = max(0, int(current))
            if total is not None:
                next_progress["total"] = max(0, int(total))
            self._progress = next_progress

    def _wait_until(
        self,
        producer: Callable[[], dict[str, Any]],
        predicate: Callable[[dict[str, Any]], bool],
        label: str,
        *,
        timeout_seconds: float,
    ) -> dict[str, Any]:
        deadline = time.monotonic() + timeout_seconds
        last: dict[str, Any] | None = None
        while time.monotonic() < deadline:
            if self._cancel_requested:
                raise EdgeAutomationError("Người dùng đã dừng thao tác Edge.")
            last = producer()
            if predicate(last):
                return last
            time.sleep(0.25)
        raise EdgeAutomationError(f"Quá thời gian chờ {label}. Trạng thái cuối: {last}")

    def _cancel_sopoka_page(self) -> None:
        if self._driver is None:
            return
        try:
            self._driver.execute_script("window.__orderBridgeExportCancelled = true;")
        except WebDriverException:
            return

    def _save_failure_screenshot(self, action: str) -> None:
        if self._driver is None:
            return
        try:
            safe_action = re.sub(r"[^a-z0-9_-]+", "-", action.casefold()) or "error"
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            path = self.config.logs_dir / f"edge-{safe_action}-{stamp}.png"
            self._driver.save_screenshot(str(path))
            self._logger.error("Browser action failed: %s screenshot=%s", action, path)
        except WebDriverException:
            return

    def _safe_current_url(self) -> str | None:
        try:
            return self._driver.current_url if self._driver else None
        except WebDriverException:
            return None


__all__ = ["EdgeAutomation", "EdgeAutomationError", "EdgeLoginRequired"]
