const DMS_REPORT_URL = "http://dms.sopoka.com.vn/reports/summaryOrderSales";
const KIOT_URL = "https://nongsanlamhong.kiotviet.vn/man/#/DashBoard";
const KIOT_ORDERS_URL = "https://nongsanlamhong.kiotviet.vn/man/#/Orders";
const EXPECTED_BRANCH = "Đồng Nai 2 - ĐN2";
const DASHBOARD_PREFIXES = ["http://127.0.0.1:8765/", "http://localhost:8765/"];
// Every action is initiated from the dashboard.  There is deliberately no
// background polling, scheduled report download, or automatic ordering.

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function pushEvent(level, message, details = {}) {
  chrome.tabs.query({}).then((tabs) => {
    for (const tab of tabs) {
      if (DASHBOARD_PREFIXES.some((prefix) => tab.url?.startsWith(prefix))) {
        chrome.tabs.sendMessage(tab.id, {
          type: "PUSH_EVENT",
          event: { level, message, details, at: new Date().toISOString() },
        }).catch(() => {});
      }
    }
  });
}

function isHostTab(tab, hostname) {
  try {
    return new URL(tab.url).hostname === hostname;
  } catch {
    return false;
  }
}

async function findTab(hostname, preferredWindowId = null) {
  const tabs = await chrome.tabs.query({});
  const choose = (candidates) => candidates.find((tab) => tab.active) || candidates[0] || null;

  // A user can have more than one Chrome window open.  The dashboard bridge
  // supplies its own window id, which is the only safe place to drive the
  // visible workflow.  Never silently take a same-domain tab from another
  // window when that id is available.
  if (Number.isInteger(preferredWindowId)) {
    return choose(tabs.filter((tab) => tab.windowId === preferredWindowId && isHostTab(tab, hostname)));
  }

  try {
    const lastFocused = await chrome.windows.getLastFocused({ populate: true });
    const inLastFocused = choose((lastFocused.tabs || []).filter((tab) => isHostTab(tab, hostname)));
    if (inLastFocused) return inLastFocused;
  } catch {
    // Fall through to the complete tab list.
  }
  return choose(tabs.filter((tab) => isHostTab(tab, hostname)));
}

async function getOrCreateTab(hostname, url, preferredWindowId = null) {
  let tab = await findTab(hostname, preferredWindowId);
  if (!tab) {
    const createProperties = { url, active: true };
    if (Number.isInteger(preferredWindowId)) createProperties.windowId = preferredWindowId;
    tab = await chrome.tabs.create(createProperties);
  } else {
    tab = await chrome.tabs.update(tab.id, { active: true });
    await chrome.windows.update(tab.windowId, { focused: true });
  }
  return tab;
}

async function waitForComplete(tabId, timeoutMs = 30000, shouldCancel = null) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (typeof shouldCancel === "function" && shouldCancel()) {
      throw new Error("Người dùng đã dừng thao tác Edge.");
    }
    const tab = await chrome.tabs.get(tabId);
    if (tab.status === "complete") return tab;
    await sleep(250);
  }
  throw new Error("Trang tải quá lâu. Hãy kiểm tra mạng rồi thử lại.");
}

async function runInTab(tabId, func, args = []) {
  const [result] = await chrome.scripting.executeScript({
    target: { tabId },
    world: "MAIN",
    func,
    args,
  });
  if (!result) throw new Error("Không nhận được phản hồi từ trang.");
  if (result.result?.error) throw new Error(result.result.error);
  return result.result;
}

async function sendMessageWithReload(tabId, message) {
  try {
    return await chrome.tabs.sendMessage(tabId, message);
  } catch (error) {
    const text = String(error?.message || error);
    if (!/receiving end does not exist|context invalidated|could not establish connection/i.test(text)) {
      throw error;
    }
    await chrome.tabs.reload(tabId);
    await waitForComplete(tabId);
    await sleep(700);
    return chrome.tabs.sendMessage(tabId, message);
  }
}

function dmsInspectTask() {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    if (!element || !element.isConnected || element.hidden || element.offsetParent === null) return false;
    const style = getComputedStyle(element);
    return style.visibility !== "hidden" && style.display !== "none";
  };
  const values = [...document.querySelectorAll('button.timK,button.xuatEx,td._GATo,input,[class*="date" i],[class*="range" i],li')]
    .filter(visible)
    .map((element) => norm(element.value || element.textContent || element.getAttribute("aria-label")))
    .filter(Boolean);
  const dateTexts = values.filter((value) => /\d{2}\/\d{2}\/\d{4}\s*(-|đến)\s*\d{2}\/\d{2}\/\d{4}/i.test(value));
  const labels = [...new Set(values.filter((value) => ["Hôm nay", "Áp dụng", "Tìm kiếm", "Xuất Excel"].includes(value)))];
  return { ok: true, url: location.href, title: document.title, labels, dateTexts: dateTexts.slice(0, 5) };
}

// Runs in Sopoka's MAIN world. The page exposes On_SeachData() and
// ket_xuat() directly there, which is more reliable than relying on a
// framework synthetic click handler from an isolated content script.
// Selectors are intentionally the exact elements supplied from the live page.
async function dmsExportTodayTask() {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const assertNotCancelled = () => {
    if (window.__orderBridgeExportCancelled === true) {
      throw new Error("Người dùng đã dừng thao tác xuất báo cáo.");
    }
  };
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
  const visible = (element) => {
    if (!element) return false;
    if (!element.isConnected || element.hidden || element.offsetParent === null) return false;
    const style = getComputedStyle(element);
    return style.visibility !== "hidden" && style.display !== "none";
  };
  const actionable = (element) => visible(element) && !element.disabled && element.getAttribute("aria-disabled") !== "true";
  const waitFor = async (finder, label, timeoutMs = 15000) => {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      assertNotCancelled();
      const result = finder();
      if (result) return result;
      await sleepOnPage(180);
    }
    throw new Error(`Không tìm thấy ${label} trên Sopoka.`);
  };
  // One click only: Sopoka's date picker closes itself after a duplicated
  // synthetic click (for example dispatchEvent plus click()).
  const realClick = (element) => element.click();
  const now = new Date();
  const todayText = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
  let stage = "mở bộ chọn ngày";
  try {
    assertNotCancelled();
    const picker = await waitFor(
      () => {
        const cell = document.querySelector("td._GATo");
        return actionable(cell) ? cell : null;
      },
      "bộ chọn ngày td._GATo",
    );
    realClick(picker);

    stage = "chọn Hôm nay";
    const today = await waitFor(
      () => [...document.querySelectorAll("li")].find((element) => actionable(element) && norm(element.textContent) === "hôm nay"),
      "mục Hôm nay",
    );
    realClick(today);

    stage = "áp dụng khoảng ngày";
    const apply = [...document.querySelectorAll("button,input[type=button]")]
      .find((element) => actionable(element) && norm(element.value || element.textContent) === "áp dụng");
    if (apply) realClick(apply);
    await sleepOnPage(600);
    assertNotCancelled();

    stage = "bấm Tìm kiếm";
    const search = await waitFor(
      () => {
        const button = document.querySelector("button.timK");
        return actionable(button) ? button : null;
      },
      "nút Tìm kiếm button.timK",
    );
    const before = resultSignature();
    if (typeof window.On_SeachData === "function") window.On_SeachData();
    else realClick(search);

    // Sopoka starts its request asynchronously.  Waiting for a minimum time
    // and several stable passes prevents exporting the pre-search table.
    stage = "chờ kết quả tìm kiếm";
    const searchStartedAt = Date.now();
    await sleepOnPage(1100);
    assertNotCancelled();
    let stablePasses = 0;
    let prior = resultSignature();
    await waitFor(
      () => {
        const busy = [...document.querySelectorAll(".loading,.loader,.spinner,.fa-spinner,[aria-busy=\"true\"]")].some(visible);
        const current = resultSignature();
        stablePasses = !busy && current === prior ? stablePasses + 1 : 0;
        prior = current;
        const exportButton = document.querySelector("button.xuatEx");
        const minimumWaitMet = Date.now() - searchStartedAt >= 2200;
        return minimumWaitMet && !busy && stablePasses >= 5 && actionable(exportButton) ? exportButton : null;
      },
      "kết quả tìm kiếm ổn định",
      45000,
    );

    stage = "bấm Xuất Excel";
    const exportButton = await waitFor(
      () => {
        const button = document.querySelector("button.xuatEx");
        return actionable(button) ? button : null;
      },
      "nút Xuất Excel button.xuatEx",
    );
    if (typeof window.ket_xuat === "function") window.ket_xuat();
    else realClick(exportButton);
    return { ok: true, stage, todayText, before, after: resultSignature(), usedDirectHandlers: true };
  } catch (error) {
    return {
      ok: false,
      stage,
      error: `Sopoka dừng ở bước ${stage}: ${String(error?.message || error)}`,
    };
  }

  function resultSignature() {
    const rows = [...document.querySelectorAll("table tbody tr")].filter(visible);
    return `${rows.length}:${rows.slice(0, 5).map((row) => String(row.textContent || "").length).join(",")}`;
  }
}

// Sopoka's date control recursively handles HTMLElement.click() when called
// from an injected function.  These phases only locate/tag elements in the
// page; the actual date-picker clicks are sent as genuine Chrome mouse input
// by debuggerClick below.
async function dmsExportPhaseTask(phase) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
  const assertNotCancelled = () => {
    if (window.__orderBridgeExportCancelled === true) throw new Error("Người dùng đã dừng thao tác xuất báo cáo.");
  };
  const waitFor = async (finder, label, timeoutMs = 15000) => {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      assertNotCancelled();
      const result = finder();
      if (result) return result;
      await sleepOnPage(180);
    }
    throw new Error(`Không tìm thấy ${label} trên Sopoka.`);
  };
  const tag = (element, target) => {
    document.querySelectorAll(`[data-orderbridge-target]`).forEach((node) => node.removeAttribute("data-orderbridge-target"));
    element.setAttribute("data-orderbridge-target", target);
    return { ok: true, target };
  };
  const resultSignature = () => {
    const rows = [...document.querySelectorAll("table tbody tr")];
    return `${rows.length}:${rows.slice(0, 5).map((row) => String(row.textContent || "").length).join(",")}`;
  };

  try {
    if (phase === "reset") {
      window.__orderBridgeExportCancelled = false;
      return { ok: true, phase };
    }
    if (phase === "find-today") {
      const today = await waitFor(
        () => [...document.querySelectorAll("li")].find((element) => norm(element.textContent) === "hôm nay"),
        "mục Hôm nay",
      );
      return { ...tag(today, "today"), phase };
    }
    if (phase === "find-apply") {
      const apply = [...document.querySelectorAll("button,input[type=button]")]
        .find((element) => norm(element.value || element.textContent) === "áp dụng");
      return apply ? { ...tag(apply, "apply"), phase, present: true } : { ok: true, phase, present: false };
    }
    if (phase === "search") {
      assertNotCancelled();
      if (typeof window.On_SeachData !== "function") throw new Error("Không tìm thấy hàm On_SeachData() của Sopoka.");
      window.On_SeachData();
      return { ok: true, phase, before: resultSignature() };
    }
    if (phase === "wait-results") {
      // Sopoka recursively errors when extension code asks the date/report
      // widgets for layout information.  A short fixed wait keeps this phase
      // DOM-read-only and relies on its documented export handler afterward.
      await sleepOnPage(2800);
      assertNotCancelled();
      const exportButton = document.querySelector("button.xuatEx");
      if (!exportButton) throw new Error("Không tìm thấy nút Xuất Excel button.xuatEx trên Sopoka.");
      return { ...tag(exportButton, "export"), phase, after: resultSignature() };
    }
    if (phase === "export") {
      assertNotCancelled();
      if (typeof window.ket_xuat !== "function") throw new Error("Không tìm thấy hàm ket_xuat() của Sopoka.");
      window.ket_xuat();
      return { ok: true, phase };
    }
    throw new Error(`Pha Sopoka không hợp lệ: ${phase}.`);
  } catch (error) {
    return { ok: false, phase, error: `Sopoka dừng ở bước ${phase}: ${String(error?.message || error)}` };
  }
}

// Direct Edge automation loads only the two functions below from this file.
// They target the sales-summary report, not the order-list export.  Keeping
// the selectors here also lets the browser task wait for Sopoka's AJAX result
// before the Python layer clicks Export.
function dmsSummaryInspectTask() {
  const visible = (element) => {
    if (!element || !element.isConnected || element.hidden || element.offsetParent === null) return false;
    const style = getComputedStyle(element);
    return style.display !== "none" && style.visibility !== "hidden";
  };
  const range = [...document.querySelectorAll("#reportrange")].find(visible);
  const search = [...document.querySelectorAll("button.timK")].find(visible);
  const exportButton = [...document.querySelectorAll("button.xuatEx")].find(visible);
  return {
    ok: true,
    url: location.href,
    title: document.title,
    reportReady: Boolean(range && search && exportButton),
    rangeText: String(range?.textContent || "").replace(/\s+/g, " ").trim(),
    searchVisible: Boolean(search),
    exportVisible: Boolean(exportButton),
  };
}

async function dmsExportPhaseTaskV2(phase) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  // Sopoka's date.js replaces its Date-now API with a Date object. Arithmetic on
  // that object can recurse through Date.prototype.toString(), so injected
  // report helpers must use the browser's monotonic timer instead.
  const monotonic = () => performance.now();
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
  const visible = (element) => {
    if (!element || !element.isConnected || element.hidden || element.offsetParent === null) return false;
    const style = getComputedStyle(element);
    return style.display !== "none" && style.visibility !== "hidden";
  };
  const actionable = (element) => visible(element) && !element.disabled && element.getAttribute("aria-disabled") !== "true";
  const assertNotCancelled = () => {
    if (window.__orderBridgeExportCancelled === true) {
      throw new Error("Người dùng đã dừng thao tác xuất báo cáo.");
    }
  };
  const waitFor = async (finder, label, timeoutMs = 15000) => {
    const deadline = monotonic() + timeoutMs;
    while (monotonic() < deadline) {
      assertNotCancelled();
      const found = finder();
      if (found) return found;
      await sleepOnPage(180);
    }
    throw new Error(`Không tìm thấy ${label} trên báo cáo tổng hợp bán hàng.`);
  };
  const tag = (element, target) => {
    document.querySelectorAll("[data-orderbridge-target]").forEach((node) => node.removeAttribute("data-orderbridge-target"));
    element.setAttribute("data-orderbridge-target", target);
    return { ok: true, target };
  };
  const resultRows = () => [...document.querySelectorAll("table tbody tr")].filter(visible);
  const resultSignature = () => resultRows()
    .slice(0, 12)
    .map((row) => norm(row.textContent).slice(0, 250))
    .join("|");
  const isBusy = () => {
    const jqueryBusy = typeof window.$ === "function" && Number(window.$.active || 0) > 0;
    const overlayBusy = [...document.querySelectorAll(".loading,.loader,.spinner,.fa-spinner,[aria-busy=\"true\"]")]
      .some(visible);
    return jqueryBusy || overlayBusy;
  };
  const expectedToday = () => {
    const now = new Date();
    return `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
  };

  if (phase === "reset") {
    window.__orderBridgeExportCancelled = false;
    return { ok: true, phase };
  }
  if (phase === "find-today") {
    const today = await waitFor(
      () => [...document.querySelectorAll(".daterangepicker li")]
        .find((element) => actionable(element) && norm(element.textContent) === "hôm nay"),
      "mục Hôm nay",
    );
    return { ...tag(today, "today"), phase };
  }
  if (phase === "find-apply") {
    const apply = [...document.querySelectorAll(".daterangepicker button.applyBtn, .daterangepicker input.applyBtn")]
      .find((element) => actionable(element) && norm(element.value || element.textContent).includes("áp dụng"));
    return apply ? { ...tag(apply, "apply"), phase, present: true } : { ok: true, phase, present: false };
  }
  if (phase === "verify-today") {
    const range = await waitFor(
      () => [...document.querySelectorAll("#reportrange")].find(visible),
      "ô khoảng ngày #reportrange",
    );
    const rangeText = String(range.textContent || "").replace(/\s+/g, " ").trim();
    const dates = rangeText.match(/\d{2}\/\d{2}\/\d{4}/g) || [];
    const expected = expectedToday();
    return {
      ok: true,
      phase,
      expectedDate: expected,
      rangeText,
      dates,
      matches: dates.length >= 2 && dates[0] === expected && dates[1] === expected,
    };
  }
  if (phase === "search") {
    const search = await waitFor(
      () => [...document.querySelectorAll("button.timK")].find(actionable),
      "nút Tìm kiếm",
    );
    const before = resultSignature();
    window.__orderBridgeSearchBefore = before;
    window.__orderBridgeSearchStartedAt = monotonic();
    if (typeof window.On_SeachData === "function") window.On_SeachData();
    else search.click();
    return { ok: true, phase, before };
  }
  if (phase === "wait-results") {
    const before = String(window.__orderBridgeSearchBefore || "");
    const storedStartedAt = Number(window.__orderBridgeSearchStartedAt);
    const startedAt = Number.isFinite(storedStartedAt) ? storedStartedAt : monotonic();
    let last = resultSignature();
    let stablePasses = 0;
    let rows = resultRows().length;
    const exportButton = await waitFor(() => {
      const current = resultSignature();
      rows = resultRows().length;
      stablePasses = !isBusy() && current === last ? stablePasses + 1 : 0;
      last = current;
      const exportControl = [...document.querySelectorAll("button.xuatEx")].find(actionable);
      const stillPromptingToSearch = norm(document.querySelector("table")?.textContent)
        .includes("bấm tìm kiếm để xem báo cáo");
      return monotonic() - startedAt >= 2500 && !isBusy() && stablePasses >= 5 && exportControl
        && !stillPromptingToSearch
        ? exportControl
        : null;
    }, "kết quả Tìm kiếm ổn định", 60000);
    return {
      ...tag(exportButton, "export"),
      phase,
      stable: true,
      rows,
      before,
      after: last,
      changed: last !== before,
    };
  }
  throw new Error(`Pha Sopoka không hợp lệ: ${phase}.`);
}

async function debuggerClick(tabId, selector, label, timeoutMs = 15000, shouldCancel = null) {
  const deadline = Date.now() + timeoutMs;
  let lastBoxError = null;
  while (Date.now() < deadline) {
    if (typeof shouldCancel === "function" && shouldCancel()) throw new Error("Người dùng đã dừng thao tác Edge.");
    const debuggee = { tabId };
    let attached = false;
    try {
      await chrome.debugger.attach(debuggee, "1.3");
      attached = true;
      const { root } = await chrome.debugger.sendCommand(debuggee, "DOM.getDocument", { depth: -1, pierce: true });
      const { nodeIds = [] } = await chrome.debugger.sendCommand(debuggee, "DOM.querySelectorAll", { nodeId: root.nodeId, selector });
      for (const nodeId of nodeIds) {
        try {
          await chrome.debugger.sendCommand(debuggee, "DOM.scrollIntoViewIfNeeded", { nodeId });
          const { model } = await chrome.debugger.sendCommand(debuggee, "DOM.getBoxModel", { nodeId });
          const quad = model.content || model.border;
          if (!Array.isArray(quad) || quad.length < 8) continue;
          const points = [[quad[0], quad[1]], [quad[2], quad[3]], [quad[4], quad[5]], [quad[6], quad[7]]];
          const x = points.reduce((sum, point) => sum + point[0], 0) / points.length;
          const y = points.reduce((sum, point) => sum + point[1], 0) / points.length;
          await chrome.debugger.sendCommand(debuggee, "Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
          await chrome.debugger.sendCommand(debuggee, "Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
          return { ok: true, selector, x: Math.round(x), y: Math.round(y) };
        } catch (error) {
          lastBoxError = error;
        }
      }
    } finally {
      if (attached) await chrome.debugger.detach(debuggee).catch(() => {});
    }
    await sleep(180);
  }
  const reason = String(lastBoxError?.message || lastBoxError || "không có phần tử hiển thị");
  throw new Error(`Không tìm thấy ${label} đang hiển thị (${selector}): ${reason}`);
}

async function debuggerClickByExactText(tabId, selector, expectedText, label, timeoutMs = 15000, shouldCancel = null) {
  const normalizeHtmlText = (html) => String(html || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
    .normalize("NFC")
    .toLowerCase();
  const target = String(expectedText || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (typeof shouldCancel === "function" && shouldCancel()) throw new Error("Người dùng đã dừng thao tác Edge.");
    const debuggee = { tabId };
    let attached = false;
    try {
      await chrome.debugger.attach(debuggee, "1.3");
      attached = true;
      const { root } = await chrome.debugger.sendCommand(debuggee, "DOM.getDocument", { depth: -1, pierce: true });
      const { nodeIds = [] } = await chrome.debugger.sendCommand(debuggee, "DOM.querySelectorAll", { nodeId: root.nodeId, selector });
      for (const nodeId of nodeIds) {
        const { outerHTML } = await chrome.debugger.sendCommand(debuggee, "DOM.getOuterHTML", { nodeId });
        if (normalizeHtmlText(outerHTML) !== target) continue;
        try {
          await chrome.debugger.sendCommand(debuggee, "DOM.scrollIntoViewIfNeeded", { nodeId });
          const { model } = await chrome.debugger.sendCommand(debuggee, "DOM.getBoxModel", { nodeId });
          const quad = model.content || model.border;
          if (!Array.isArray(quad) || quad.length < 8) continue;
          const points = [[quad[0], quad[1]], [quad[2], quad[3]], [quad[4], quad[5]], [quad[6], quad[7]]];
          const x = points.reduce((sum, point) => sum + point[0], 0) / points.length;
          const y = points.reduce((sum, point) => sum + point[1], 0) / points.length;
          await chrome.debugger.sendCommand(debuggee, "Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
          await chrome.debugger.sendCommand(debuggee, "Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
          return { ok: true, selector, text: expectedText, x: Math.round(x), y: Math.round(y) };
        } catch {
          // A hidden copy of the same list item may exist.  Continue to the
          // next matching node and only click one with a real box model.
        }
      }
    } finally {
      if (attached) await chrome.debugger.detach(debuggee).catch(() => {});
    }
    await sleep(180);
  }
  throw new Error(`Không tìm thấy ${label} với nội dung chính xác "${expectedText}".`);
}

function waitForReportDownload(startedAt, timeoutMs = 150000) {
  let cleanup = () => {};
  let cancel = () => {};
  const promise = new Promise((resolve, reject) => {
    let timeout = null;
    let settled = false;
    let observedItem = null;
    let warningShown = false;

    const matchesReport = (item) => {
      const filename = String(item?.filename || "").split(/[\\/]/).pop();
      return /reportordersalesummary/i.test(filename) && /\.xls$/i.test(filename);
    };
    cleanup = () => {
      if (timeout) clearTimeout(timeout);
      chrome.downloads.onChanged.removeListener(onChanged);
      chrome.downloads.onCreated.removeListener(onCreated);
    };
    const settle = (value, isError = false) => {
      if (settled) return;
      settled = true;
      cleanup();
      if (isError) reject(value instanceof Error ? value : new Error(String(value)));
      else resolve(value);
    };
    const explainBlockedDownload = (item, reason = "") => {
      const filename = String(item?.filename || "file báo cáo").split(/[\\/]/).pop();
      const suffix = reason ? ` (${reason})` : "";
      return `Edge đang chặn ${filename}${suffix}. Hãy bấm Giữ lại trong danh sách tải xuống; tool sẽ tự tiếp tục ngay khi file hoàn tất.`;
    };
    const inspect = async (id) => {
      const [item] = await chrome.downloads.search({ id });
      return item || null;
    };
    const warnBlocked = (item, reason) => {
      if (warningShown) return;
      warningShown = true;
      pushEvent("warning", explainBlockedDownload(item, reason), {
        filename: item?.filename || null,
        downloadId: item?.id || null,
        reason: reason || null,
      });
    };
    const onCreated = (item) => {
      if (!matchesReport(item) || new Date(item.startTime || 0).getTime() < startedAt - 2000) return;
      observedItem = item;
      pushEvent("info", "Sopoka đã bắt đầu tải file report; đang chờ Edge hoàn tất.", {
        filename: item.filename,
        downloadId: item.id,
      });
      if (item.danger && item.danger !== "safe" && item.danger !== "accepted") warnBlocked(item, item.danger);
    };
    const onChanged = async (delta) => {
      if (settled) return;
      const item = await inspect(delta.id);
      if (!item || !matchesReport(item) || new Date(item.startTime || 0).getTime() < startedAt - 2000) return;
      observedItem = item;
      const danger = delta.danger?.current || item.danger;
      if (danger && danger !== "safe" && danger !== "accepted") warnBlocked(item, danger);
      if (delta.state?.current === "complete") {
        settle(item);
        return;
      }
      if (delta.state?.current === "interrupted" || item.state === "interrupted") {
        const reason = item.error || delta.error?.current || "interrupted";
        if (/user_canceled|dangerous|blocked/i.test(String(reason))) warnBlocked(item, reason);
        settle(new Error(`Edge không hoàn tất tải report (${reason}). ${explainBlockedDownload(item, reason)}`), true);
      }
    };
    cancel = (reason = "Người dùng đã dừng chờ file report.") => settle(new Error(reason), true);
    chrome.downloads.onCreated.addListener(onCreated);
    chrome.downloads.onChanged.addListener(onChanged);
    timeout = setTimeout(async () => {
      const items = await chrome.downloads.search({ orderBy: ["-startTime"], limit: 20 });
      const item = items.find((entry) => matchesReport(entry) && new Date(entry.startTime || 0).getTime() >= startedAt - 2000) || observedItem;
      if (item?.state === "complete") {
        settle(item);
      } else if (item) {
        const reason = item.error || item.danger || item.state || "đang chờ Edge";
        if (item.danger && item.danger !== "safe" && item.danger !== "accepted") warnBlocked(item, reason);
        settle(new Error(`Chưa có file report hoàn tất sau ${Math.ceil(timeoutMs / 1000)} giây. ${explainBlockedDownload(item, reason)}`), true);
      } else {
        settle(new Error(`Sopoka chưa tạo file ReportOrderSaleSummary trong ${Math.ceil(timeoutMs / 1000)} giây. Không có dữ liệu nào được ghi vào SQLite.`), true);
      }
    }, timeoutMs);
  });
  return { promise, cancel };
}

async function exportToday(preferredWindowId = null) {
  let cancelled = false;
  let downloadWatcher = null;
  let pageResult = null;
  const cancelThisExport = () => {
    cancelled = true;
    downloadWatcher?.cancel("Người dùng đã dừng tác vụ tải report trước khi có file hoàn tất.");
  };
  cancelActiveBrowserAction = cancelThisExport;
  try {
    pushEvent("info", "Đang mở báo cáo tổng hợp bán hàng trên Sopoka", { stage: "open_sopoka" });
    const tab = await getOrCreateTab("dms.sopoka.com.vn", DMS_REPORT_URL, preferredWindowId);
    if (tab.url !== DMS_REPORT_URL) {
      await chrome.tabs.update(tab.id, { url: DMS_REPORT_URL, active: true });
    } else {
      await chrome.tabs.update(tab.id, { active: true });
    }
    await waitForComplete(tab.id, 30000, () => cancelled);
    if (cancelled) throw new Error("Người dùng đã dừng tác vụ tải report.");
    pushEvent("info", "Đang chọn Hôm nay và yêu cầu Sopoka xuất Excel.", { stage: "sopoka_export" });
    const startedAt = Date.now();
    downloadWatcher = waitForReportDownload(startedAt);
    // The page action may fail before we await the watcher.  Attach a passive
    // handler so a cancellation never becomes an unhandled service-worker
    // rejection; the awaited original promise still carries the real error.
    downloadWatcher.promise.catch(() => {});
    await debuggerClick(tab.id, "td._GATo", "bộ chọn ngày Sopoka", 15000, () => cancelled);
    pushEvent("info", "Đã mở bộ chọn ngày Sopoka.", { stage: "date_picker" });
    if (cancelled) throw new Error("Người dùng đã dừng tác vụ tải report.");
    await debuggerClickByExactText(tab.id, "li", "Hôm nay", "mục Hôm nay", 15000, () => cancelled);
    pushEvent("info", "Đã chọn Hôm nay trên Sopoka.", { stage: "today_selected" });
    // SOPoka's current report page needs exactly the three user-provided
    // actions: open date picker -> Hôm nay -> Xuất Excel.  Do not insert
    // Search or Apply, because they are not part of this live workflow.
    await sleep(300);
    if (cancelled) throw new Error("Người dùng đã dừng tác vụ tải report.");
    pageResult = await debuggerClick(tab.id, "button.xuatEx", "nút Xuất Excel Sopoka", 15000, () => cancelled);
    if (cancelled) throw new Error("Người dùng đã dừng tác vụ tải report.");
    pushEvent("info", "Đã bấm Xuất Excel. Đang theo dõi file report trong Edge.", { stage: "waiting_download" });
    const download = await downloadWatcher.promise;
    if (cancelled) throw new Error("Người dùng đã dừng tác vụ tải report.");
    pushEvent("success", "Đã tải báo cáo Sopoka", { filename: download.filename, stage: "download_complete" });
    return { ok: true, pageResult, sourcePath: download.filename };
  } catch (error) {
    downloadWatcher?.cancel(String(error?.message || error));
    throw error;
  } finally {
    if (cancelActiveBrowserAction === cancelThisExport) cancelActiveBrowserAction = null;
  }
}

async function inspectSopoka(preferredWindowId = null) {
  const tab = await getOrCreateTab("dms.sopoka.com.vn", DMS_REPORT_URL, preferredWindowId);
  if (tab.url !== DMS_REPORT_URL) await chrome.tabs.update(tab.id, { url: DMS_REPORT_URL, active: true });
  await waitForComplete(tab.id);
  await sleep(1200);
  return { tabId: tab.id, ...(await runInTab(tab.id, dmsInspectTask)) };
}

function kiotInspectTask(expectedBranch) {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const branchText = (value) => norm(value).replace(/\s*select$/i, "").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const elements = [...document.querySelectorAll("button,a,li,div,span")].filter(visible);
  const branchElements = elements.filter((element) => {
    const text = branchText(element.textContent);
    const rect = element.getBoundingClientRect();
    const atApplicationEdge = rect.top < 130 || rect.bottom > innerHeight - 120;
    return atApplicationEdge && text.length < 80 && /Đồng Nai\s*\d+\s*-\s*ĐN\d+/i.test(text);
  });
  const branchCandidates = [...new Set(branchElements.map((element) => branchText(element.textContent)))];
  const tabElements = elements
    .filter((element) => /^(Đặt hàng|Hóa đơn)\s*\d+$/i.test(norm(element.textContent)) && element.getBoundingClientRect().top < 110)
    .filter((element, index, items) => items.findIndex((item) => norm(item.textContent) === norm(element.textContent)) === index)
    .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
  const tabLabels = tabElements.map((element) => norm(element.textContent));
  const isActiveTab = (element) => {
    let current = element;
    for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
      if (current.getAttribute?.("aria-selected") === "true") return true;
      if (/(^|\s)(active|selected|current)(\s|$)/i.test(String(current.className?.baseVal || current.className || ""))) return true;
    }
    return false;
  };
  const activeTab = tabElements.find(isActiveTab) || null;
  const activeTabLabel = activeTab ? norm(activeTab.textContent) : null;
  const note = [...document.querySelectorAll('input,textarea,[contenteditable="true"]')]
    .find((element) => visible(element) && /ghi\s*chu\s*don\s*hang/i.test(
      norm(`${element.getAttribute("placeholder") || ""} ${element.getAttribute("data-placeholder") || ""} ${element.getAttribute("aria-label") || ""}`)
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D")
    ));
  // A stopped/partially imported cart can have no note yet while still
  // containing product rows. Note text alone is not a safe blank-cart proof.
  const cartRoot = document.querySelector("#cartListProduct");
  const cartItemCount = cartRoot && visible(cartRoot)
    ? cartRoot.querySelectorAll(".carts-item").length
    : 0;
  const dangerousButtons = [...document.querySelectorAll("button,[role=button]")].filter(visible).filter((element) => norm(element.textContent || element.getAttribute("aria-label")).toLowerCase() === "đặt hàng");
  const createOrderCandidates = dangerousButtons
    .filter((element) => {
      const rect = element.getBoundingClientRect();
      return rect.top > 55 && rect.top < 320 && rect.width < 260 && rect.height < 100;
    })
    .map((element) => {
      const rect = element.getBoundingClientRect();
      return { label: norm(element.textContent || element.getAttribute("aria-label")), top: Math.round(rect.top), width: Math.round(rect.width), height: Math.round(rect.height) };
    });
  const orderTabs = tabLabels.filter((text) => /^Đặt hàng/i.test(text));
  const invoiceTabs = tabLabels.filter((text) => /^Hóa đơn/i.test(text));
  const invoiceMode = /^Hóa đơn/i.test(activeTabLabel || "") || (!activeTabLabel && invoiceTabs.length > 0 && orderTabs.length === 0);
  const orderMode = /^Đặt hàng/i.test(activeTabLabel || "") || (Boolean(note) && !invoiceMode && orderTabs.length > 0);
  return {
    ok: true,
    url: location.href,
    branch: branchCandidates.sort((a, b) => a.length - b.length)[0] || null,
    branchOk: branchCandidates.some((text) => branchText(text).startsWith(branchText(expectedBranch))),
    orderMode,
    invoiceMode,
    activeTab: activeTabLabel,
    tabLabels,
    orderTabs,
    invoiceTabs,
    noteAvailable: Boolean(note),
    noteValue: note ? norm(note.isContentEditable ? note.textContent : note.value) : null,
    cartItemCount,
    finalOrderButtonCount: dangerousButtons.length,
    createOrderCandidates,
  };
}

function kiotCreateBlankOrderTabTask() {
  const visible = (element) => {
    if (!element || !element.isConnected) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const tabLabels = () => [...document.querySelectorAll("button,a,li,div,span")]
    .filter(visible)
    .map((element) => norm(element.textContent))
    .filter((label) => /^Đặt hàng\s*\d+$/i.test(label))
    .filter((label, index, values) => values.indexOf(label) === index);
  const clickLikeUser = (element) => {
    element.scrollIntoView?.({ block: "nearest", inline: "center" });
    element.focus?.({ preventScroll: true });
    const rect = element.getBoundingClientRect();
    const hit = document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2);
    const target = hit && (element === hit || element.contains(hit)) ? hit : element;
    for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
      const EventType = type.startsWith("pointer") && typeof PointerEvent !== "undefined" ? PointerEvent : MouseEvent;
      target.dispatchEvent(new EventType(type, {
        bubbles: true,
        cancelable: true,
        view: window,
        button: 0,
        buttons: type.endsWith("down") ? 1 : 0,
      }));
    }
    target.click?.();
  };
  const controlText = (element) => norm([
    element.getAttribute?.("title"),
    element.getAttribute?.("aria-label"),
    element.textContent,
  ].filter(Boolean).join(" "));
  const addCandidates = () => {
    const all = [...document.querySelectorAll("a.btn-add-tabs,button.btn-add-tabs,[role=button].btn-add-tabs,[title], [aria-label]")]
      .filter(visible)
      .filter((element, index, values) => values.indexOf(element) === index)
      .filter((element) => {
        const rect = element.getBoundingClientRect();
        return rect.top < 180 && rect.bottom > 0;
      });
    const exactClass = all.filter((element) => /(^|\s)btn-add-tabs(\s|$)/i.test(String(element.className || "")));
    const explicit = exactClass.filter((element) => /^(thêm mới|them moi)$/i.test(controlText(element)));
    const titled = exactClass.filter((element) => /thêm mới|them moi/i.test(controlText(element)));
    const anyExplicit = all.filter((element) => /^(thêm mới|them moi)$/i.test(controlText(element)));
    return [...new Set([...explicit, ...titled, ...anyExplicit])];
  };

  try {
    const before = tabLabels();
    const add = addCandidates()[0];
    if (!add) throw new Error("Không tìm thấy nút Thêm mới để tạo tab Đặt hàng trống.");

    clickLikeUser(add);
    return {
      ok: true,
      beforeOrderTabs: before,
      action: "create_blank_order_tab",
    };
  } catch (error) {
    return { error: String(error?.message || error) };
  }
}

function kiotConvertActiveInvoiceTask() {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    if (!element || !element.isConnected) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const classText = (element) => String(element?.className?.baseVal || element?.className || "");
  const isActive = (element) => {
    let current = element;
    for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
      if (current.getAttribute?.("aria-selected") === "true") return true;
      if (/(^|\s)(active|selected|current)(\s|$)/i.test(classText(current))) return true;
    }
    return false;
  };
  const click = (element) => {
    element.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window }));
    element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, view: window }));
    element.click();
  };
  const markAndClick = (element, strategy) => {
    if (!element || !visible(element)) return null;
    const text = norm(`${element.textContent || ""} ${element.getAttribute?.("aria-label") || ""} ${element.getAttribute?.("title") || ""}`);
    if (/(?:đóng|close|times|xóa|delete)/i.test(text) || /(?:close|times|remove|delete)/i.test(classText(element))) return null;
    element.setAttribute("data-orderbridge-convert-invoice", "1");
    click(element);
    return { ok: true, strategy, target: text || element.tagName };
  };

  try {
    const tabCandidates = [...document.querySelectorAll("button,a,[role=tab],[role=button],li,div,span")]
      .filter(visible)
      .filter((element) => /^Hóa đơn\s*\d+$/i.test(norm(element.textContent)) && element.getBoundingClientRect().top < 120);
    const invoiceTab = tabCandidates.find(isActive) || tabCandidates[0];
    if (!invoiceTab) throw new Error("Không tìm thấy tab Hóa đơn đang mở để chuyển sang Đặt hàng.");

    const textNode = tabCandidates
      .filter((element) => norm(element.textContent) === norm(invoiceTab.textContent))
      .sort((left, right) => (left.children.length - right.children.length) || (left.getBoundingClientRect().width - right.getBoundingClientRect().width))[0] || invoiceTab;
    const textRect = textNode.getBoundingClientRect();
    const tabRoot = invoiceTab.closest("[role=tab],button,a,li") || invoiceTab.parentElement || invoiceTab;
    const tabRect = tabRoot.getBoundingClientRect();

    // Prefer an explicitly labelled conversion control.  KiotViet versions
    // differ, so title/aria/data-tooltip are checked before visual geometry.
    const labelled = [...document.querySelectorAll("[title],[aria-label],[data-tooltip],[data-original-title]")]
      .filter(visible)
      .filter((element) => {
        const label = norm(`${element.getAttribute("title") || ""} ${element.getAttribute("aria-label") || ""} ${element.getAttribute("data-tooltip") || ""} ${element.getAttribute("data-original-title") || ""}`);
        return /chuyển\s+sang\s+đặt\s+hàng/i.test(label);
      })
      .sort((left, right) => {
        const l = left.getBoundingClientRect(); const r = right.getBoundingClientRect();
        return Math.abs(l.left - textRect.left) - Math.abs(r.left - textRect.left);
      })[0];
    const labelledResult = markAndClick(labelled, "labelled-control");
    if (labelledResult) return labelledResult;

    // In the observed KiotViet UI the conversion icon is a small control just
    // to the left of “Hóa đơn N”; the close control is on the right.  Scope
    // the fallback to that small left-hand area so it cannot touch payment or
    // final-order controls elsewhere on the page.
    const iconCandidates = [...document.querySelectorAll("button,a,[role=button],i,svg,span,div")]
      .filter(visible)
      .map((element) => ({ element, rect: element.getBoundingClientRect(), className: classText(element) }))
      .filter(({ element, rect, className }) => {
        const nearInvoice = rect.top < 120
          && rect.bottom >= textRect.top - 8
          && rect.top <= textRect.bottom + 8
          && rect.right >= textRect.left - 56
          && rect.left <= textRect.left + 8;
        const compact = rect.width <= 56 && rect.height <= 56;
        const notClose = !/(close|times|remove|delete|fa-xmark|fa-times)/i.test(className)
          && !/(đóng|close)/i.test(norm(element.getAttribute("aria-label") || element.getAttribute("title") || element.textContent));
        return nearInvoice && compact && notClose && element !== textNode;
      })
      .sort((left, right) => {
        const leftScore = /exchange|switch|transfer|convert|swap|arrow|sync|repeat|rotate/i.test(left.className) ? -1000 : 0;
        const rightScore = /exchange|switch|transfer|convert|swap|arrow|sync|repeat|rotate/i.test(right.className) ? -1000 : 0;
        if (leftScore !== rightScore) return leftScore - rightScore;
        const leftDistance = Math.abs(left.rect.right - textRect.left);
        const rightDistance = Math.abs(right.rect.right - textRect.left);
        if (leftDistance !== rightDistance) return leftDistance - rightDistance;
        return (left.rect.width * left.rect.height) - (right.rect.width * right.rect.height);
      });
    for (const candidate of iconCandidates) {
      const result = markAndClick(candidate.element, "left-tab-icon");
      if (result) return result;
    }

    // Last resort: activate the invoice tab itself.  Some KiotViet builds
    // expose the conversion menu from the tab click rather than a separate
    // icon.  The next task will only select the exact conversion menu item.
    if (tabRect.top < 120) {
      const result = markAndClick(invoiceTab, "invoice-tab");
      if (result) return result;
    }
    throw new Error("Không xác định được nút Chuyển sang đặt hàng của tab Hóa đơn.");
  } catch (error) {
    return { error: String(error?.message || error) };
  }
}

function kiotChooseConvertInvoiceMenuTask() {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    if (!element || !element.isConnected) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  try {
    const item = [...document.querySelectorAll("button,a,li,[role=button],[role=menuitem],[role=option]")]
      .filter(visible)
      .filter((element) => norm(element.textContent || element.getAttribute("aria-label")) === "Chuyển sang đặt hàng")
      .sort((left, right) => (left.getBoundingClientRect().width * left.getBoundingClientRect().height) - (right.getBoundingClientRect().width * right.getBoundingClientRect().height))[0];
    if (!item) return { ok: true, selected: false };
    item.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    item.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window }));
    item.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, view: window }));
    item.click();
    return { ok: true, selected: true };
  } catch (error) {
    return { error: String(error?.message || error) };
  }
}

function kiotPrepareTask(expectedBranch, branchAlreadyVerified = false) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const branchText = (value) => norm(value).replace(/\s*select$/i, "").trim();
  const visible = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const all = () => [...document.querySelectorAll("button,a,li,div,span")].filter(visible);
  const branchControls = () => all().filter((element) => {
    const text = branchText(element.textContent);
    const rect = element.getBoundingClientRect();
    return text.length < 80
      && /Đồng Nai\s*\d+\s*-\s*ĐN\d+/i.test(text)
      && (rect.top < 130 || rect.bottom > innerHeight - 120);
  }).sort((a, b) => {
    const ar = a.getBoundingClientRect(); const br = b.getBoundingClientRect();
    return ar.width * ar.height - br.width * br.height;
  });
  const safeClick = (element, purpose, { allowCreateOrder = false } = {}) => {
    if (!element) throw new Error(`Không tìm thấy ${purpose}.`);
    const label = norm(element.textContent || element.getAttribute("aria-label") || element.value).replace(/^\+\s*/, "").toLowerCase();
    if (element.matches("button,[role=button]") && label === "đặt hàng" && !allowCreateOrder) {
      throw new Error("Chặn an toàn: tool không được phép bấm nút ĐẶT HÀNG cuối cùng.");
    }
    element.click();
  };
  const tabs = () => all()
    .filter((element) => /^(Đặt hàng|Hóa đơn)\s*\d+$/i.test(norm(element.textContent)) && element.getBoundingClientRect().top < 110)
    .filter((element, index, items) => items.findIndex((item) => norm(item.textContent) === norm(element.textContent)) === index)
    .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
  const noteField = () => document.querySelector('input[placeholder*="Ghi chú đơn hàng" i],textarea[placeholder*="Ghi chú đơn hàng" i],[contenteditable="true"][data-placeholder*="Ghi chú đơn hàng" i]');

  return (async () => {
    let expected = branchControls().find((element) => branchText(element.textContent).startsWith(branchText(expectedBranch)));
    if (!expected) {
      const currentBranch = branchControls()[0];
      if (!currentBranch && branchAlreadyVerified) {
        // The sales canvas does not render the branch name after it has been
        // checked on the dashboard.  Continue with the verified session,
        // rather than clicking an unrelated control.
      } else {
      safeClick(currentBranch, "bộ chọn chi nhánh");
      const currentRect = currentBranch.getBoundingClientRect();
      await sleepOnPage(300);
      const exactMatches = all()
        .filter((element) => branchText(element.textContent || element.getAttribute("aria-label")) === branchText(expectedBranch))
        .map((element) => element.closest('button,a,li,[role="option"],[role="menuitem"],[role="button"]') || element)
        .filter((element, index, items) => visible(element) && items.indexOf(element) === index)
        .sort((a, b) => {
          const semanticA = a.matches('button,a,li,[role="option"],[role="menuitem"],[role="button"]') ? 0 : 1;
          const semanticB = b.matches('button,a,li,[role="option"],[role="menuitem"],[role="button"]') ? 0 : 1;
          if (semanticA !== semanticB) return semanticA - semanticB;
          return Math.abs(a.getBoundingClientRect().top - currentRect.top)
            - Math.abs(b.getBoundingClientRect().top - currentRect.top);
        });
      expected = exactMatches[0];
      safeClick(expected, `chi nhánh ${expectedBranch}`);
      await sleepOnPage(600);
      if (!branchControls().some((element) => branchText(element.textContent).startsWith(branchText(expectedBranch)))) {
        throw new Error(`Chưa chuyển đúng chi nhánh ${expectedBranch}.`);
      }
      }
    }

    const existingOrderTab = tabs().find((element) => /^Đặt hàng/i.test(norm(element.textContent)));
    if (existingOrderTab) {
      safeClick(existingOrderTab, "tab Đặt hàng hiện có");
      await sleepOnPage(500);
    }

    if (!noteField() || !tabs().some((element) => /^Đặt hàng/i.test(norm(element.textContent)))) {
      // Do not use an injected async callback to cross from KiotViet's
      // management SPA into `/sale/`: that navigation destroys this document
      // before Selenium receives the callback.  Python performs the same
      // route through native WebDriver navigation after receiving this signal.
      return { ok: true, needsOrderRoute: true };
    }

    const deadline = Date.now() + 12000;
    while (Date.now() < deadline) {
      if (noteField() && tabs().some((element) => /^Đặt hàng/i.test(norm(element.textContent)))) return { ok: true };
      await sleepOnPage(300);
    }
    throw new Error("Không mở được tab Đặt hàng để nhập file.");
  })().catch((error) => ({ error: String(error?.message || error) }));
}

function kiotOpenOrdersMenuTask() {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const clickableFor = (element) => element?.closest('a,button,li,[role="button"],[role="menuitem"]') || element;
  return (async () => {
    const orderMenuLabel = [...document.querySelectorAll("span.ng-binding")]
      .find((element) => visible(element) && norm(element.textContent) === "Đơn hàng" && element.getBoundingClientRect().top < 120);
    const orderMenu = clickableFor(orderMenuLabel);
    if (!orderMenu || !visible(orderMenu)) throw new Error("Không tìm thấy tab Đơn hàng trên thanh điều hướng KiotViet.");
    orderMenu.click();
    await sleepOnPage(250);
    const orderEntryLabel = [...document.querySelectorAll("a,button,li,span")]
      .find((element) => visible(element) && norm(element.textContent) === "Đặt hàng" && element.getBoundingClientRect().top < 320);
    const orderEntry = clickableFor(orderEntryLabel);
    if (!orderEntry || !visible(orderEntry)) throw new Error("Không tìm thấy mục Đặt hàng trong menu Đơn hàng.");
    orderEntry.click();
    return { ok: true };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

function kiotCleanTabsTask(confirmCloseData) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  return (async () => {
    if (!confirmCloseData) throw new Error("Explicit confirmation is required before closing old KiotViet tabs.");
    const labelFor = (element) => {
      const match = norm(element?.textContent).match(/^(\u0110\u1eb7t h\u00e0ng|H\u00f3a \u0111\u01a1n)\s*(\d+)(?:\s*[\u00d7\u2715x])?$/i);
      return match ? `${match[1]} ${Number(match[2])}` : null;
    };
    const getTabs = () => {
      const byLabel = new Map();
      for (const element of [...document.querySelectorAll("button,a,li,div,span")]) {
        const label = labelFor(element);
        if (!label || !visible(element) || element.getBoundingClientRect().top >= 110) continue;
        const existing = byLabel.get(label);
        if (!existing || element.getBoundingClientRect().width < existing.getBoundingClientRect().width) byLabel.set(label, element);
      }
      return [...byLabel.entries()].map(([label, element]) => ({ label, element }))
        .sort((a, b) => a.element.getBoundingClientRect().left - b.element.getBoundingClientRect().left);
    };
    const closeFor = (tab) => {
      const tabRect = tab.element.getBoundingClientRect();
      const candidates = [...document.querySelectorAll('button,[role="button"],[aria-label],[title],i,span,svg')]
        .filter(visible)
        .filter((element) => {
          const ariaTitle = norm(`${element.getAttribute("aria-label") || ""} ${element.getAttribute("title") || ""}`);
          const text = norm(element.textContent);
          const classes = String(element.className?.baseVal || element.className || "");
          return /(?:^|\s)(?:close|times|fa-times|icon-close|kv-close)(?:\s|$)/i.test(classes)
            || /(?:\u0111\u00f3ng|close)/i.test(ariaTitle) || /^[\u00d7\u2715x]$/i.test(text);
        })
        .filter((element) => {
          const rect = element.getBoundingClientRect();
          const verticallyAligned = rect.bottom >= tabRect.top - 8 && rect.top <= tabRect.bottom + 8;
          return verticallyAligned && rect.left >= tabRect.left - 12 && rect.left <= tabRect.right + 100;
        })
        .sort((left, right) => Math.abs(left.getBoundingClientRect().left - tabRect.right) - Math.abs(right.getBoundingClientRect().left - tabRect.right));
      const close = candidates[0];
      if (!close) return null;
      return close.closest('button,a,[role="button"],[ng-click],[onclick]') || close;
    };
    const confirmDiscard = async (timeout = 7500) => {
      // KiotViet renders its discard dialog asynchronously.  The former
      // one-shot check could run before the modal (and its "Đồng ý" button)
      // existed, leaving cleanup stopped at the first tab with unsaved data.
      // On a busy KiotViet session this modal can be rendered several
      // seconds after the tab close click, so keep polling long enough to
      // cover that delayed Angular transition.
      const deadline = Date.now() + timeout;
      while (Date.now() < deadline) {
        // KiotViet has used both native <button> controls and clickable
        // Angular wrapper elements for this modal.  Match the visible text
        // first, then walk to the actual interactive ancestor.
        const visibleButtons = [...document.querySelectorAll("button,a,div,span,[role=button],[ng-click],[onclick]")]
          .filter(visible)
          .filter((element) => /^(?:\u0110\u1ed3ng \u00fd|C\u00f3|X\u00e1c nh\u1eadn|B\u1ecf qua)$/i.test(norm(element.textContent)))
          .sort((left, right) => {
            const leftRect = left.getBoundingClientRect();
            const rightRect = right.getBoundingClientRect();
            return (leftRect.width * leftRect.height) - (rightRect.width * rightRect.height);
          });
        const dialog = [...document.querySelectorAll('[role="dialog"],.modal,.modal-dialog,[class*="dialog" i]')]
          .filter(visible)
          .find((element) => /(?:d\u1eef li\u1ec7u|ch\u01b0a l\u01b0u|kh\u00f4ng \u0111\u01b0\u1ee3c l\u01b0u|thay \u0111\u1ed5i|tab|\u0111\u00f3ng)/i.test(norm(element.textContent)));
        const confirm = dialog
          ? visibleButtons.find((element) => dialog.contains(element)) || visibleButtons[0]
          : visibleButtons[0] || null;
        if (confirm) {
          const clickTarget = confirm.closest('button,a,[role="button"],[ng-click],[onclick]') || confirm;
          clickTarget.focus?.({ preventScroll: true });
          for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
            const EventType = type.startsWith("pointer") && typeof PointerEvent !== "undefined" ? PointerEvent : MouseEvent;
            clickTarget.dispatchEvent(new EventType(type, { bubbles: true, cancelable: true, view: window, button: 0 }));
          }
          clickTarget.click();
          return true;
        }
        await sleepOnPage(100);
      }
      return false;
    };
    const waitForStableTabs = async (timeout = 3000) => {
      const deadline = Date.now() + timeout;
      let prior = null;
      let stableReads = 0;
      let latest = [];
      while (Date.now() < deadline) {
        latest = getTabs();
        const key = latest.map((tab) => tab.label).join("|");
        stableReads = key === prior ? stableReads + 1 : 0;
        if (stableReads >= 3) return latest;
        prior = key;
        await sleepOnPage(120);
      }
      return latest;
    };
    // A prior interrupted cleanup can leave this confirmation on screen.
    // Dismiss it before trying to interact with the underlying tab strip.
    // This is only for a dialog left by an earlier interrupted attempt; do
    // not spend the full modal timeout before the first tab close.
    await confirmDiscard(600);
    let closed = 0;
    while (true) {
      const currentTabs = await waitForStableTabs();
      const invoice = currentTabs.filter((tab) => /^H\u00f3a \u0111\u01a1n/i.test(tab.label)).at(-1);
      const orders = currentTabs.filter((tab) => /^\u0110\u1eb7t h\u00e0ng/i.test(tab.label));
      const tab = invoice || (orders.length > 1 ? orders.at(-1) : null);
      if (!tab) break;
      const close = closeFor(tab);
      if (!close) throw new Error(`Could not find the close control for ${tab.label}.`);
      const beforeCount = currentTabs.length;
      close.click();
      await sleepOnPage(180);
      // The confirmation belongs to the close we just issued.  Wait here,
      // rather than before that click, so a slow Angular modal cannot leave
      // the tab strip blocked.
      await confirmDiscard(9500);
      let afterTabs = await waitForStableTabs(1200);
      // Some sessions render the dialog just after the first timeout.  Give
      // it one bounded follow-up chance before declaring cleanup unsafe.
      if (afterTabs.length >= beforeCount) {
        await confirmDiscard(5000);
        afterTabs = await waitForStableTabs(1200);
      }
      if (afterTabs.length >= beforeCount) throw new Error(`KiotViet did not close ${tab.label}; stopping cleanup.`);
      closed += 1;
    }
    const remaining = await waitForStableTabs();
    const invoices = remaining.filter((tab) => /^H\u00f3a \u0111\u01a1n/i.test(tab.label));
    const orders = remaining.filter((tab) => /^\u0110\u1eb7t h\u00e0ng/i.test(tab.label));
    if (invoices.length || orders.length !== 1) throw new Error("Cleanup did not leave exactly one order tab and zero invoice tabs.");
    return { ok: true, closed, remainingTabs: remaining.map((tab) => tab.label) };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

function kiotRequestCloseTabTask(tabLabel) {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const labelFor = (element) => {
    const match = norm(element?.textContent).match(/^(\u0110\u1eb7t h\u00e0ng|H\u00f3a \u0111\u01a1n)\s*(\d+)(?:\s*[\u00d7\u2715x])?$/i);
    return match ? `${match[1]} ${Number(match[2])}` : null;
  };
  const closeFor = (tab) => {
    const textRect = tab.getBoundingClientRect();
    // The small arrow at the *left* of a cart changes Đặt hàng to Hóa đơn.
    // It is not a close control.  Only a close candidate to the right of the
    // text label can discard a cart.
    const candidates = [...document.querySelectorAll('button,[role="button"],[aria-label],[title],i,span,svg')]
      .filter(visible)
      .filter((element) => {
        const ariaTitle = norm(`${element.getAttribute("aria-label") || ""} ${element.getAttribute("title") || ""}`);
        const text = norm(element.textContent);
        const classes = String(element.className?.baseVal || element.className || "");
        return /(?:^|\s)(?:close|times|fa-times|icon-close|kv-close)(?:\s|$)/i.test(classes)
          || /(?:\u0111\u00f3ng|close)/i.test(ariaTitle) || /^[\u00d7\u2715x]$/i.test(text);
      })
      .filter((element) => {
        const rect = element.getBoundingClientRect();
        const centerX = rect.left + rect.width / 2;
        return rect.bottom >= textRect.top - 8 && rect.top <= textRect.bottom + 8
          && centerX > textRect.right + 2 && rect.left <= textRect.right + 110;
      })
      .sort((left, right) => Math.abs(left.getBoundingClientRect().left - textRect.right) - Math.abs(right.getBoundingClientRect().left - textRect.right));
    const close = candidates[0];
    return close?.closest('button,a,[role="button"],[ng-click],[onclick]') || close || null;
  };
  try {
    const label = norm(tabLabel);
    const tab = [...document.querySelectorAll("button,a,li,div,span")]
      .filter((element) => labelFor(element) === label && visible(element) && element.getBoundingClientRect().top < 110)
      .sort((left, right) => left.getBoundingClientRect().width - right.getBoundingClientRect().width)[0];
    if (!tab) return { error: `Không tìm thấy ${label} để đóng.` };
    const close = closeFor(tab);
    if (!close) return { error: `Không tìm thấy nút đóng của ${label}.` };
    close.click();
    return { ok: true, tabLabel: label, closeSide: "right_of_label" };
  } catch (error) {
    return { error: String(error?.message || error) };
  }
}

function kiotInspectDiscardDialogTask() {
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const roots = [document];
  const elements = [];
  while (roots.length) {
    const root = roots.pop();
    for (const element of root.querySelectorAll ? root.querySelectorAll("*") : []) {
      elements.push(element);
      if (element.shadowRoot) roots.push(element.shadowRoot);
    }
  }
  const describe = (element) => ({
    tag: element.tagName?.toLowerCase(),
    className: String(element.className?.baseVal || element.className || "").slice(0, 240),
    role: element.getAttribute?.("role") || "",
    text: norm(element.textContent).slice(0, 360),
    html: String(element.outerHTML || "").slice(0, 700),
  });
  const textMatches = elements
    .filter(visible)
    .filter((element) => /^(?:\u0110\u1ed3ng \u00fd|B\u1ecf qua)$/i.test(norm(element.textContent)))
    .slice(0, 12)
    .map(describe);
  const dialogs = elements
    .filter(visible)
    .filter((element) => /\u0110\u00f3ng\s+\u0110\u1eb7t h\u00e0ng\s*8/i.test(norm(element.textContent)))
    .slice(0, 12)
    .map(describe);
  return {
    ok: true,
    bodyIncludesConfirm: norm(document.body?.innerText).includes("Đồng ý"),
    bodyIncludesTitle: /Đóng\s+Đặt hàng\s*8/i.test(norm(document.body?.innerText)),
    textMatches,
    dialogs,
  };
}

function kiotOpenImportTask() {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  return (async () => {
    const menuButtons = [...document.querySelectorAll("button,[role=button],i,svg")].filter(visible).filter((element) => /menu|bars|hamburger/i.test(element.getAttribute("aria-label") || element.getAttribute("title") || element.className?.baseVal || element.className || "") || norm(element.textContent) === "☰");
    const menu = menuButtons.sort((a, b) => b.getBoundingClientRect().right - a.getBoundingClientRect().right)[0] || [...document.querySelectorAll("button,div,span,i,svg")].filter(visible).filter((element) => element.getBoundingClientRect().top < 70 && element.getBoundingClientRect().right > innerWidth - 80 && /☰|≡/.test(norm(element.textContent)))[0];
    if (!menu) throw new Error("Không tìm thấy menu góc phải của màn hình đặt hàng.");
    menu.click();
    await sleepOnPage(250);
    const importItem = [...document.querySelectorAll("button,a,li,div,span")].filter(visible).filter((element) => norm(element.textContent).toLowerCase() === "import file").sort((a, b) => a.getBoundingClientRect().width - b.getBoundingClientRect().width)[0];
    if (!importItem) throw new Error("Không tìm thấy mục Import file.");
    importItem.click();
    await sleepOnPage(350);
    return { ok: true, fileInputPresent: Boolean(document.querySelector('input[type="file"]')) };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

async function setFileInput(tabId, filePath) {
  const debuggee = { tabId };
  await chrome.debugger.attach(debuggee, "1.3");
  try {
    const { root } = await chrome.debugger.sendCommand(debuggee, "DOM.getDocument", { depth: -1, pierce: true });
    const { nodeId } = await chrome.debugger.sendCommand(debuggee, "DOM.querySelector", { nodeId: root.nodeId, selector: 'input[type="file"]' });
    if (!nodeId) throw new Error("Hộp chọn file chưa xuất hiện.");
    await chrome.debugger.sendCommand(debuggee, "DOM.setFileInputFiles", { nodeId, files: [filePath] });
  } finally {
    await chrome.debugger.detach(debuggee).catch(() => {});
  }
}

function kiotWaitImportTask(expectedNewTabs, previousTabCount) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const visible = (element) => {
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const tabLabels = () => [...new Set([...document.querySelectorAll("button,a,li,div,span")].filter(visible).filter((element) => /^(Đặt hàng|Hóa đơn)\s*\d+$/i.test(norm(element.textContent)) && element.getBoundingClientRect().top < 100).map((element) => norm(element.textContent)))];
  return (async () => {
    await sleepOnPage(350);
    const importDialog = [...document.querySelectorAll('[role="dialog"],.modal,.modal-dialog,[class*="dialog" i]')]
      .filter(visible)
      .find((element) => /(?:nhận danh sách đặt hàng|file excel|tối đa 19 trang tính|xử lý dữ liệu)/i.test(norm(element.textContent)));
    if (importDialog) {
      const importConfirm = [...importDialog.querySelectorAll("button,[role=button]")]
        .filter(visible)
        .find((element) => /^Nhập dữ liệu$|^Import$|^Thực hiện$|^Đồng ý$|^Xác nhận$/i.test(norm(element.textContent)));
      if (importConfirm) {
        importConfirm.click();
        await sleepOnPage(350);
      }
    }
    // A 19-sheet workbook can take over a minute while the sales SPA creates
    // each cart.  Do not let the page task time out during that legitimate
    // work; the Python caller supplies its own bounded 120--240s timeout.
    const deadline = Date.now() + Math.max(120000, Math.min(240000, Number(expectedNewTabs || 1) * 12000));
    while (Date.now() < deadline) {
      const labels = tabLabels();
      const orderLabels = labels.filter((label) => /^Đặt hàng\s*\d+$/i.test(label));
      const invoiceLabels = labels.filter((label) => /^Hóa đơn\s*\d+$/i.test(label));
      // Importing while KiotViet is in invoice mode can still create the
      // expected number of tabs, but they are the wrong objects. Reject that
      // state before notes or orders are touched; the caller will stop safely.
      if (invoiceLabels.length === 0 && orderLabels.length === previousTabCount + expectedNewTabs) {
        return { ok: true, tabLabels: labels };
      }
      await sleepOnPage(500);
    }
    const labels = tabLabels();
    throw new Error(`Import chưa rõ trạng thái: có ${labels.length} tab, dự kiến ${previousTabCount + expectedNewTabs}. Tool sẽ không tự upload lại.`);
  })().catch((error) => ({ error: String(error?.message || error) }));
}

function kiotFillNotesTask(orderExpectations) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const tabLabel = (element) => {
    const match = norm(element?.textContent).match(/^(Đặt hàng)\s*(\d+)(?:\s*[×✕x])?$/i);
    return match ? `${match[1]} ${Number(match[2])}` : null;
  };
  const rendered = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const inViewport = (element) => {
    if (!rendered(element)) return false;
    const rect = element.getBoundingClientRect();
    return rect.bottom > 0 && rect.top < innerHeight && rect.right > 0 && rect.left < innerWidth;
  };
  const classText = (element) => String(element?.className?.baseVal || element?.className || "");
  const isInteractiveTab = (element, label) => {
    if (!element || !rendered(element) || tabLabel(element) !== label) return false;
    const role = String(element.getAttribute("role") || "").toLowerCase();
    const tag = element.tagName?.toLowerCase();
    return role === "tab"
      || /\b(active|selected|tab|cart|order)\b/i.test(classText(element))
      || ["a", "button", "li"].includes(tag)
      || element.hasAttribute("ng-click")
      || element.hasAttribute("onclick");
  };
  const scoreTarget = (element, label, depth) => {
    if (!isInteractiveTab(element, label)) return -1;
    const role = String(element.getAttribute("role") || "").toLowerCase();
    const tag = element.tagName?.toLowerCase();
    let score = 0;
    if (role === "tab") score += 1000;
    if (element.hasAttribute("ng-click") || element.hasAttribute("onclick")) score += 700;
    if (["a", "button", "li"].includes(tag)) score += 500;
    if (/\b(tab|cart|order)\b/i.test(classText(element))) score += 300;
    return score - depth;
  };
  const getTabs = () => {
    const found = [...document.querySelectorAll("button,a,li,div,span")]
      .filter((element) => rendered(element) && element.getBoundingClientRect().top < 125)
      .map((element) => ({ element, label: tabLabel(element) }))
      .filter((entry) => entry.label);
    const byLabel = new Map();
    for (const entry of found) {
      const candidates = [];
      let current = entry.element;
      for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
        const score = scoreTarget(current, entry.label, depth);
        if (score >= 0) candidates.push({ element: current, score, depth });
      }
      // A label span normally lives inside the clickable tab.  Choose the
      // closest semantic or Angular click target, not the decorative span.
      const best = candidates.sort((left, right) => right.score - left.score || left.depth - right.depth)[0]
        || { element: entry.element, score: 0, depth: 0 };
      const existing = byLabel.get(entry.label);
      if (!existing || best.score > existing.score || (best.score === existing.score && best.depth < existing.depth)) {
        byLabel.set(entry.label, { label: entry.label, target: best.element, score: best.score, depth: best.depth });
      }
    }
    return [...byLabel.values()].sort((left, right) => Number(left.label.match(/\d+$/)[0]) - Number(right.label.match(/\d+$/)[0]));
  };
  const noteField = () => [...document.querySelectorAll('input[placeholder*="Ghi chú đơn hàng" i],textarea[placeholder*="Ghi chú đơn hàng" i],[contenteditable="true"][data-placeholder*="Ghi chú đơn hàng" i]')]
    .find(inViewport);
  const activeMatches = (tab) => {
    const candidates = [tab.target, ...(tab.target.querySelectorAll?.('[aria-selected="true"],.active,.selected,.current') || [])];
    for (const candidate of candidates) {
      let current = candidate;
      for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
        if (String(current.getAttribute("aria-selected")) === "true" || /(^|\s)(active|selected|current)(\s|$)/i.test(classText(current))) return true;
      }
    }
    return false;
  };
  const revealTarget = (target) => {
    let current = target.parentElement;
    for (let depth = 0; current && depth < 10; depth += 1, current = current.parentElement) {
      if (current.scrollWidth <= current.clientWidth + 2) continue;
      const targetRect = target.getBoundingClientRect();
      const containerRect = current.getBoundingClientRect();
      if (targetRect.left < containerRect.left + 8) current.scrollLeft += targetRect.left - containerRect.left - 36;
      else if (targetRect.right > containerRect.right - 8) current.scrollLeft += targetRect.right - containerRect.right + 36;
    }
  };
  const activate = async (tab) => {
    const activeNow = getTabs().find(activeMatches);
    if (activeNow?.label === tab.label) return activeNow;
    for (let attempt = 0; attempt < 12; attempt += 1) {
      const fresh = getTabs().find((entry) => entry.label === tab.label);
      if (!fresh) throw new Error(`Tab ${tab.label} biến mất trong khi đang kiểm tra.`);
      const target = fresh.target;
      revealTarget(target);
      const rect = target.getBoundingClientRect();
      const centerIsReachable = inViewport(target) && rect.width > 2 && rect.height > 2;
      const hit = centerIsReachable ? document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2) : null;
      const clickTarget = hit && (target === hit || target.contains(hit)) ? hit : target;
      clickTarget.focus?.({ preventScroll: true });
      for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
        const EventType = type.startsWith("pointer") && typeof PointerEvent !== "undefined" ? PointerEvent : MouseEvent;
        clickTarget.dispatchEvent(new EventType(type, { bubbles: true, cancelable: true, view: window, button: 0, buttons: type.endsWith("down") ? 1 : 0 }));
      }
      clickTarget.click();
      await sleepOnPage(150);
      const active = getTabs().find(activeMatches);
      if (active?.label === tab.label) return active;
    }
    const active = getTabs().find(activeMatches);
    throw new Error(`Không kích hoạt được ${tab.label}; tab hiện tại là ${active?.label || "không xác định"}.`);
  };
  const hasVisibleText = (needle) => {
    const expected = norm(needle).toLowerCase();
    if (!expected) return true;
    // Product rows can be inside the active cart's own scroll area.  Their
    // individual cells may be below the browser viewport, but #cartListProduct
    // is the single live cart component bound to the selected tab.  Reading
    // that container is safe; unlike document.body it does not include stale
    // inactive carts from other KiotViet tabs.
    const activeCart = document.querySelector("#cartListProduct");
    if (activeCart && rendered(activeCart) && norm(activeCart.textContent).toLowerCase().includes(expected)) return true;
    const elementMatch = [...document.querySelectorAll("td,span,a,div,input")].some((element) => {
      if (!(element.matches("td,input") || element.children.length === 0)) return false;
      const value = norm(element.value || element.textContent).toLowerCase();
      return value.length <= 260 && value.includes(expected) && inViewport(element);
    });
    // Do not fall back to document.body.innerText here. KiotViet keeps the
    // text of some inactive carts in the document, so that fallback can make
    // two different tabs appear to belong to the same order. A visible leaf
    // in the active cart is the required proof before writing a note.
    return elementMatch;
  };
  const escapeRegExp = (value) => String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const productCodeMatches = (text, code) => {
    const expected = fold(code);
    if (!expected) return false;
    // A SKU is an identifier, not free-form prose.  In particular, ``E15``
    // must not match ``E15N`` (or ``E15-NEW``), otherwise two unrelated
    // imported carts can appear to have the same product signature.
    const codeCharacter = "[\\p{L}\\p{N}_+\\-./()]";
    const pattern = new RegExp(`(^|[^${codeCharacter.slice(1, -1)}])${escapeRegExp(expected)}(?=$|[^${codeCharacter.slice(1, -1)}])`, "u");
    return pattern.test(fold(text));
  };
  const activeCartItemCount = () => {
    const activeCart = document.querySelector("#cartListProduct");
    if (!activeCart || !rendered(activeCart)) return null;
    // KiotViet renders a hidden serial/IMEI companion row immediately after
    // every product row, both carrying the `.carts-item` class.  Counting all
    // of them doubles the worksheet line count and makes every otherwise
    // correct tab look like a mismatch.  Only non-companion rows represent
    // imported products.
    const rows = activeCart.querySelectorAll(".carts-item:not(.row-serial-child)");
    return rows.length || null;
  };
  const fold = (value) => norm(value).normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "").replace(/\u0111/g, "d").toLowerCase();
  const numericValue = (value) => {
    const raw = norm(value).replace(/[^0-9,.-]/g, "");
    if (!/\d/.test(raw)) return null;
    const negative = raw.startsWith("-") ? -1 : 1;
    const text = raw.replace(/-/g, "");
    const separators = [...text.matchAll(/[.,]/g)].map((match) => match.index);
    let normalized = text;
    if (separators.length) {
      const last = separators[separators.length - 1];
      const trailing = text.length - last - 1;
      const repeated = separators.length > 1;
      // Amounts use Vietnamese thousands separators (1.234.567 or
      // 1,234,567); a final one/two-digit part is a decimal quantity.
      const decimal = trailing > 0 && trailing <= 2 && !(repeated && trailing === 2);
      normalized = decimal
        ? `${text.slice(0, last).replace(/[.,]/g, "")}.${text.slice(last + 1)}`
        : text.replace(/[.,]/g, "");
    }
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed * negative : null;
  };
  const numberMatches = (actual, expected) => {
    const left = numericValue(actual);
    const right = numericValue(expected);
    return left !== null && right !== null && Math.abs(left - right) < 0.0001;
  };
  const numericTokens = (value) => norm(value).match(/-?\d+(?:[.,]\d+)*/g) || [];
  const activeCartRows = () => {
    const cart = document.querySelector("#cartListProduct");
    if (!cart || !rendered(cart)) return [];
    return [...cart.querySelectorAll(".carts-item:not(.row-serial-child)")];
  };
  const rowText = (row) => norm([
    row.textContent,
    ...[...row.querySelectorAll("input,textarea,[contenteditable=true]")]
      .map((element) => element.isContentEditable ? element.textContent : element.value),
  ].join(" "));
  const hasExactProductCode = (code) => activeCartRows()
    .some((row) => productCodeMatches(rowText(row), code));
  const rowQuantityValues = (row) => {
    const fields = [...row.querySelectorAll("input,textarea,[contenteditable=true]")];
    const marked = fields.filter((field) => /quantity|quant|so.?luong|sl/i.test([
      field.name,
      field.id,
      field.className,
      field.getAttribute("ng-model"),
      field.getAttribute("data-field"),
      field.getAttribute("aria-label"),
      field.getAttribute("placeholder"),
    ].join(" ")));
    const source = marked.length ? marked : fields;
    return source.flatMap((field) => numericTokens(field.isContentEditable ? field.textContent : field.value));
  };
  const expectedLines = (expectation) => {
    const lines = Array.isArray(expectation.productLines) ? expectation.productLines : [];
    return lines.map((line) => ({
      productCode: norm(line?.productCode || line?.product_code),
      productName: norm(line?.productName || line?.product_name),
      quantity: line?.quantity,
      netAmount: line?.netAmount ?? line?.net_amount,
    })).filter((line) => line.productCode || line.productName);
  };
  const rowMatchesLine = (row, line) => {
    const text = fold(rowText(row));
    const codeOk = !line.productCode || productCodeMatches(rowText(row), line.productCode);
    const nameOk = !line.productName || text.includes(fold(line.productName));
    return codeOk && nameOk;
  };
  const detailedLinesMatch = (expectation) => {
    const lines = expectedLines(expectation);
    const rows = activeCartRows();
    if (!lines.length || !rows.length) return false;
    return lines.every((line) => {
      const candidates = rows.filter((row) => rowMatchesLine(row, line));
      if (candidates.length !== 1) return false;
      const row = candidates[0];
      const quantityOk = rowQuantityValues(row).some((value) => numberMatches(value, line.quantity));
      const amountOk = line.netAmount === null || line.netAmount === undefined
        || numericTokens(rowText(row)).some((value) => numberMatches(value, line.netAmount));
      return quantityOk && amountOk;
    });
  };
  const paymentAmountMatches = (expectation) => {
    if (expectation.paymentAmount === null || expectation.paymentAmount === undefined) return false;
    const labels = [...document.querySelectorAll("span,div,td,label,strong,b,p")]
      .filter((element) => rendered(element) && element.children.length === 0)
      .filter((element) => /^(khach can tra|tong thanh toan)$/.test(fold(element.textContent)));
    const values = [...document.querySelectorAll("span,div,td,strong,b,input")]
      .filter(rendered)
      .filter((element) => element.children.length === 0 || element.matches("input"))
      .filter((element) => numberMatches(element.value || element.textContent, expectation.paymentAmount));
    return labels.some((label) => {
      const labelRect = label.getBoundingClientRect();
      return values.some((value) => {
        const valueRect = value.getBoundingClientRect();
        return valueRect.left >= labelRect.left
          && Math.abs(valueRect.top - labelRect.top) < 90
          && valueRect.right > innerWidth * 0.55;
      });
    });
  };
  const productSignature = (expectation) => {
    const codes = [...new Set([
      ...(Array.isArray(expectation.productCodes) ? expectation.productCodes : []),
      expectation.firstSku,
    ].map((value) => norm(value).toLowerCase()).filter(Boolean))].sort();
    const names = [...new Set((Array.isArray(expectation.productNames) ? expectation.productNames : [])
      .map((value) => norm(value).toLowerCase()).filter(Boolean))].sort();
    const lineCount = Number(expectation.lineCount) || 0;
    return JSON.stringify({ codes, names, lineCount });
  };
  // ``orderCode`` is the report's "Mã đặt hàng liên quan".  It is unique in
  // the incoming batch, but it is not present on KiotViet until this task has
  // safely written it into Ghi chú.  For two distinct report orders that have
  // the same products and quantities, require the visible customer name to
  // prove which imported cart belongs to which related-order code.
  const productSignatureCounts = new Map();
  const customerCounts = new Map();
  const customerKey = (expectation) => fold(expectation.customerCode || expectation.customerName);
  for (const expectation of orderExpectations) {
    const signature = productSignature(expectation);
    productSignatureCounts.set(signature, (productSignatureCounts.get(signature) || 0) + 1);
    const customer = customerKey(expectation);
    if (customer) customerCounts.set(customer, (customerCounts.get(customer) || 0) + 1);
  }
  const requiresCustomerProof = (expectation) => (
    (productSignatureCounts.get(productSignature(expectation)) || 0) > 1
  );
  const requiresDetailedProof = (expectation) => (
    (customerCounts.get(customerKey(expectation)) || 0) > 1
  );
  const cartMatchesExpectation = (expectation) => {
    const codes = [...new Set([
      ...(Array.isArray(expectation.productCodes) ? expectation.productCodes : []),
      expectation.firstSku,
    ].map((value) => norm(value)).filter(Boolean))];
    const names = [...new Set((Array.isArray(expectation.productNames) ? expectation.productNames : [])
      .map((value) => norm(value)).filter(Boolean))];
    // Product codes are the stable identity of an imported worksheet.  Require
    // every code when available; this handles carts where the customer panel
    // is outside the viewport, while still rejecting two carts with the same
    // partial signature.
    const skuOk = codes.length > 0 && codes.every((code) => hasExactProductCode(code));
    // Some KiotViet builds render only the product name in the cart row and
    // keep the SKU in a hidden binding.  The names came from the same Excel
    // rows as the codes, so they are a safe second identity proof.  If both
    // signatures are present, either complete signature is sufficient; a
    // duplicate name remains rejected later when candidate count is checked.
    const nameOk = names.length > 0 && names.every((name) => hasVisibleText(name));
    const customerProofs = [expectation.customerCode, expectation.customerName]
      .map((value) => norm(value)).filter(Boolean);
    const customerOk = customerProofs.some((value) => hasVisibleText(value));
    const itemCount = activeCartItemCount();
    const countOk = !Number(expectation.lineCount) || itemCount === Number(expectation.lineCount);
    const productsOk = (skuOk || nameOk) && countOk;
    // Customer text is normally supplemental because some KiotViet layouts
    // hide that panel while Angular activates a cart.  It becomes mandatory
    // only when another worksheet has the exact same product signature.  In
    // that case product code/name/quantity alone cannot distinguish two
    // different related-order codes, so customer evidence prevents a note
    // from being written to the wrong cart.
    if (requiresCustomerProof(expectation) && !customerOk) return false;
    // Same-customer carts can share product names and the number of rows.  In
    // that case match every product line's quantity and the amount due before
    // the tab is eligible to receive a related-order code.
    if (requiresDetailedProof(expectation)) {
      return customerOk && productsOk && detailedLinesMatch(expectation) && paymentAmountMatches(expectation);
    }
    if (codes.length > 0 || names.length > 0) {
      return productsOk;
    }
    return countOk;
  };
  const verifyContent = async (expectation, label) => {
    const deadline = Date.now() + 5000;
    while (Date.now() < deadline) {
      if (cartMatchesExpectation(expectation)) return;
      await sleepOnPage(150);
    }
    const missing = [
      expectation.firstSku && !hasExactProductCode(expectation.firstSku) ? `SKU ${expectation.firstSku}` : null,
      Array.isArray(expectation.productNames) && expectation.productNames.length > 0
        && !expectation.productNames.every((name) => hasVisibleText(name))
        ? `tên sản phẩm trong Excel không khớp`
        : null,
      expectation.customerName && !hasVisibleText(expectation.customerName) ? `khách ${expectation.customerName}` : null,
    ].filter(Boolean).join("; ");
    throw new Error(`Tab ${label} không khớp ${missing}; dừng để tránh ghi chú nhầm.`);
  };
  const setValue = (element, value) => {
    if (element.isContentEditable) {
      element.focus();
      element.textContent = value;
      element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
      element.dispatchEvent(new Event("change", { bubbles: true }));
      return norm(element.textContent);
    }
    element.focus();
    const prototype = element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(prototype, "value")?.set;
    setter.call(element, value);
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
    return norm(element.value);
  };
  const readValue = (element) => norm(element?.isContentEditable ? element.textContent : element?.value);

  return (async () => {
    const tabs = getTabs();
    if (tabs.length < orderExpectations.length + 1) throw new Error(`Thiếu tab sau import: có ${tabs.length}, cần ít nhất ${orderExpectations.length + 1}.`);
    // Fast, deterministic path for the normal import layout.  The previous
    // exhaustive remapping retried every tab up to 12 times and could exceed
    // several minutes on KiotViet's horizontal tab strip even when labels
    // were already the persisted expected labels.  Never fall back to a
    // positional guess: if every exact label is present, prove its visible
    // SKU/customer and write the note once.
    const tabNumber = (label) => Number(String(label || "").match(/(\d+)$/)?.[1] || 0);
    const candidateShifts = [0, 1, -1];
    for (const shift of candidateShifts) {
      const candidateTabs = orderExpectations.map((expectation) => {
        const number = tabNumber(expectation.expectedTabLabel) + shift;
        return tabs.find((tab) => tabNumber(tab.label) === number) || null;
      });
      if (!candidateTabs.every(Boolean)) continue;
      const candidateMappings = [];
      let candidateOk = true;
      for (let index = 0; index < orderExpectations.length; index += 1) {
        const expectation = orderExpectations[index];
        const tab = candidateTabs[index];
        revealTarget(tab.target);
        tab.target.click();
        await sleepOnPage(180);
        const contentDeadline = Date.now() + 3500;
        while (Date.now() < contentDeadline) {
          if (cartMatchesExpectation(expectation)) break;
          await sleepOnPage(120);
        }
        if (!cartMatchesExpectation(expectation)) { candidateOk = false; break; }
        candidateMappings.push({ expectation, tab, sheetIndex: index + 1 });
      }
      // Prove the whole candidate shift before changing a note.  A failed
      // shift used to leave a prefix of correctly-looking notes behind, then
      // fall through to remapping.  That made a stopped import unnecessarily
      // ambiguous even though no final order had been clicked.
      if (!candidateOk) continue;
      const fastVerified = [];
      for (const mapping of candidateMappings) {
        const { expectation, tab, sheetIndex } = mapping;
        await activate(tab);
        const note = noteField();
        if (!note) throw new Error(`Không tìm thấy ô Ghi chú đơn hàng ở tab ${tab.label}.`);
        const readBack = setValue(note, expectation.orderCode);
        if (readBack !== norm(expectation.orderCode)) throw new Error(`Không xác minh được ghi chú ${expectation.orderCode}.`);
        note.blur();
        await sleepOnPage(80);
        fastVerified.push({ sheetIndex, tabLabel: tab.label, orderCode: expectation.orderCode });
      }
      return { ok: true, verified: fastVerified, finalOrderClicked: false };
    }
    // Do not fail solely because one stale blank cart is present. The
    // SKU/customer remapping below proves each imported cart before writing
    // its note and safely handles that extra tab.
    let importedTabs = orderExpectations.map((expectation, index) => {
      const expectedLabel = norm(expectation.expectedTabLabel);
      if (expectedLabel) {
        const exactTab = tabs.find((tab) => tab.label === expectedLabel);
        if (!exactTab) throw new Error(`Không tìm thấy tab dự kiến ${expectedLabel}.`);
        return exactTab;
      }
      return tabs[tabs.length - orderExpectations.length + index];
    });

    // Prefer the persisted sheet-to-tab labels when their visible customer and
    // SKU can be proved.  This avoids a redundant full re-identification pass
    // when KiotViet retained a blank cart but did not shift imported labels.
    // If labels have shifted, the proof fails before any note is written and
    // the safe SKU + customer fallback below is used instead.
    let labelsAlreadyVerified = false;
    try {
      for (let index = 0; index < orderExpectations.length; index += 1) {
        const expectation = orderExpectations[index];
        const tab = importedTabs[index];
        await activate(tab);
        await verifyContent(expectation, tab.label);
      }
      labelsAlreadyVerified = true;
    } catch (_labelVerificationError) {
      labelsAlreadyVerified = false;
    }

    if (!labelsAlreadyVerified) {
      // A retained blank cart can shift the numeric labels (for example,
      // imported carts may be 2, 4..21 while the workbook expects 2..20).
      // Scan each cart once and wait for Angular to finish rendering before
      // evaluating its signature.  The previous implementation checked a
      // shifted cart only once after 160ms; a still-loading cart then looked
      // like a non-match and produced the misleading "none" error.
      const matchingTabsByCode = new Map();
      for (const tab of tabs) {
        await activate(tab);
        let matches = [];
        const deadline = Date.now() + 5000;
        do {
          matches = orderExpectations.filter((expectation) => cartMatchesExpectation(expectation));
          if (matches.length) break;
          await sleepOnPage(140);
        } while (Date.now() < deadline);
        for (const expectation of matches) {
          const code = norm(expectation.orderCode);
          if (!code) throw new Error("Má»™t worksheet import thiáº¿u MÃ£ Ä‘áº·t Ä‘Æ¡n hÃ ng liÃªn quan.");
          const current = matchingTabsByCode.get(code) || [];
          current.push(tab);
          matchingTabsByCode.set(code, current);
        }
      }
      const usedLabels = new Set();
      importedTabs = orderExpectations.map((expectation) => {
        const code = norm(expectation.orderCode);
        const candidates = matchingTabsByCode.get(code) || [];
        if (candidates.length !== 1 || usedLabels.has(candidates[0]?.label)) {
          const labels = candidates.map((entry) => entry.label).join(", ") || "none";
          const active = getTabs().find(activeMatches);
          const note = noteField();
          const itemCount = activeCartItemCount();
          throw new Error(
            `Cannot prove a unique KiotViet tab for order ${code}: ${labels}. `
            + `active=${active?.label || "unknown"}; noteEmpty=${note ? !readValue(note) : "unknown"}; `
            + `cartItems=${itemCount ?? "unknown"}.`
          );
        }
        usedLabels.add(candidates[0].label);
        return candidates[0];
      });
      labelsAlreadyVerified = true;
    }

    if (!labelsAlreadyVerified) {
    // KiotViet can insert the retained blank cart into the middle of the
    // numeric sequence (for example, leave an empty "Đặt hàng 3" while the
    // first imported sheet becomes "Đặt hàng 2").  Numeric labels and
    // "last N tabs" are therefore not a safe mapping.  Inspect every visible
    // cart once and pair a worksheet only with its unique SKU + customer.
    // No note is written until this whole mapping has been proved.
    const matchingTabsByCode = new Map();
    for (const tab of tabs) {
      await activate(tab);
      let matches = [];
      const deadline = Date.now() + 3500;
      do {
        matches = orderExpectations.filter((expectation) => {
          return cartMatchesExpectation(expectation);
        });
        if (matches.length) break;
        await sleepOnPage(140);
      } while (Date.now() < deadline);
      for (const expectation of matches) {
        const code = norm(expectation.orderCode);
        if (!code) throw new Error("Một worksheet import thiếu Mã đặt đơn hàng liên quan.");
        const current = matchingTabsByCode.get(code) || [];
        current.push(tab);
        matchingTabsByCode.set(code, current);
      }
    }
    const usedLabels = new Set();
    importedTabs = orderExpectations.map((expectation) => {
      const code = norm(expectation.orderCode);
      const candidates = matchingTabsByCode.get(code) || [];
      if (candidates.length !== 1 || usedLabels.has(candidates[0]?.label)) {
        const labels = candidates.map((entry) => entry.label).join(", ") || "không có";
        throw new Error(`Không xác định duy nhất tab theo SKU/khách cho đơn ${code}: ${labels}.`);
      }
      usedLabels.add(candidates[0].label);
      return candidates[0];
    });
    }

    // Validate all mapping before touching any note.  This makes a selector
    // mismatch fail closed with zero partial writes.
    for (let index = 0; index < orderExpectations.length; index += 1) {
      const expectation = orderExpectations[index];
      const tab = importedTabs[index];
      await activate(tab);
      const actualLabel = tab.label;
      if (!/^\u0110\u1eb7t\s+h\u00e0ng\s+[1-9]\d*$/i.test(actualLabel || "")) {
        throw new Error(`Sai liên kết tab: cần ${expectation.expectedTabLabel}, thực tế ${actualLabel}.`);
      }
      await verifyContent(expectation, actualLabel);
    }

    const verified = [];
    for (let index = 0; index < orderExpectations.length; index += 1) {
      const expectation = orderExpectations[index];
      const tab = importedTabs[index];
      await activate(tab);
      await verifyContent(expectation, tab.label);
      const note = noteField();
      if (!note) throw new Error(`Không tìm thấy ô Ghi chú đơn hàng ở tab ${index + 1}.`);
      const readBack = setValue(note, expectation.orderCode);
      if (readBack !== norm(expectation.orderCode)) throw new Error(`Không xác minh được ghi chú ${expectation.orderCode}.`);
      note.blur();
      await sleepOnPage(90);
      verified.push({ sheetIndex: index + 1, tabLabel: tab.label, orderCode: expectation.orderCode });
    }
    for (let index = 0; index < orderExpectations.length; index += 1) {
      await activate(importedTabs[index]);
      await sleepOnPage(90);
      if (readValue(noteField()) !== norm(orderExpectations[index].orderCode)) {
        throw new Error(`Ghi chú ${orderExpectations[index].orderCode} không được lưu khi kiểm tra lại tab.`);
      }
    }
    return { ok: true, verified, finalOrderClicked: false };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

// Read-only diagnostic used when a stopped batch cannot prove its tab map.
// It activates each visible cart but never edits a note, imports a file, or
// clicks the final-order button.  Returning the cart signature makes a
// mapping failure actionable instead of reporting only "none".
function kiotInspectCartTabsTask() {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const rendered = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const labelFor = (element) => {
    const match = norm(element?.textContent).match(/^(Đặt hàng|Hóa đơn)\s*(\d+)(?:\s*[×✕x])?$/i);
    return match ? `${match[1]} ${Number(match[2])}` : null;
  };
  const getTabs = () => {
    const byLabel = new Map();
    for (const element of [...document.querySelectorAll("button,a,li,div,span")]) {
      const label = labelFor(element);
      if (!label || !rendered(element) || element.getBoundingClientRect().top >= 125) continue;
      const target = element.closest('button,a,li,[role="tab"],[role="button"],[ng-click]') || element;
      const old = byLabel.get(label);
      if (!old || target.getBoundingClientRect().width < old.target.getBoundingClientRect().width) byLabel.set(label, { label, target });
    }
    return [...byLabel.values()].sort((a, b) => Number(a.label.match(/\d+$/)[0]) - Number(b.label.match(/\d+$/)[0]));
  };
  const isActive = (tab) => {
    let current = tab.target;
    for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
      if (current.getAttribute?.("aria-selected") === "true" || /(^|\s)(active|selected|current)(\s|$)/i.test(String(current.className || ""))) return true;
    }
    return false;
  };
  const activate = async (tab) => {
    const current = getTabs().find(isActive);
    if (current?.label === tab.label) return;
    tab.target.scrollIntoView?.({ block: "nearest", inline: "center" });
    const rect = tab.target.getBoundingClientRect();
    const hit = document.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2);
    const clickTarget = hit && (tab.target === hit || tab.target.contains(hit)) ? hit : tab.target;
    clickTarget.focus?.({ preventScroll: true });
    for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
      const EventType = type.startsWith("pointer") && typeof PointerEvent !== "undefined" ? PointerEvent : MouseEvent;
      clickTarget.dispatchEvent(new EventType(type, { bubbles: true, cancelable: true, view: window, button: 0 }));
    }
    clickTarget.click();
    await sleepOnPage(450);
  };
  const read = () => {
    const cart = document.querySelector("#cartListProduct");
    const note = [...document.querySelectorAll("input,textarea,[contenteditable=true]")]
      .find((element) => rendered(element) && /ghi\s*chu\s*don\s*hang/i.test(norm(`${element.getAttribute("placeholder") || ""} ${element.getAttribute("data-placeholder") || ""}`)));
    const leaves = cart ? [...cart.querySelectorAll("td,input,span,a")].filter(rendered) : [];
    return {
      cartItemCount: cart ? cart.querySelectorAll(".carts-item").length : 0,
      cartText: norm(cart?.textContent).slice(0, 1200),
      leafText: leaves.map((element) => norm(element.value || element.textContent)).filter(Boolean).slice(0, 80),
      noteValue: note ? norm(note.isContentEditable ? note.textContent : note.value) : null,
    };
  };
  const addControls = () => {
    const controls = [];
    const seen = new Set();
    for (const element of [...document.querySelectorAll("a,button,[role=button],i,span,div")]) {
      if (!rendered(element) || seen.has(element)) continue;
      const rect = element.getBoundingClientRect();
      const className = String(element.className?.baseVal || element.className || "");
      const title = norm(element.getAttribute("title"));
      const aria = norm(element.getAttribute("aria-label"));
      const text = norm(element.textContent);
      const marker = `${className} ${title} ${aria} ${text}`.toLowerCase();
      const likely = /btn-add-tabs|thêm mới|them moi|fa-plus|icon-btn|add-tabs/.test(marker);
      if (!likely || rect.top >= 180 || rect.bottom <= 0) continue;
      seen.add(element);
      controls.push({
        tag: element.tagName?.toLowerCase(),
        className: className.slice(0, 220),
        title: title.slice(0, 120),
        ariaLabel: aria.slice(0, 120),
        text: text.slice(0, 160),
        top: Math.round(rect.top),
        left: Math.round(rect.left),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
        outerHTML: element.outerHTML.slice(0, 600),
      });
    }
    return controls.slice(0, 80);
  };
  return (async () => {
    const tabs = getTabs();
    const results = [];
    for (const tab of tabs) {
      await activate(tab);
      const active = getTabs().find(isActive);
      results.push({
        tabLabel: tab.label,
        activeAfterClick: active?.label || null,
        targetHtml: tab.target.outerHTML.slice(0, 500),
        ...read(),
      });
    }
    return { ok: true, activeTab: getTabs().find(isActive)?.label || null, tabs: results, addControls: addControls() };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

// Used by the direct Edge controller only.  The note is re-read immediately
// before each final click, so a tab mismatch fails closed without submitting.
function kiotSubmitOrdersTask(orderExpectations, expectedBranch, allowExistingInvoiceTabs = false) {
  const sleepOnPage = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim();
  const classText = (element) => String(element?.className?.baseVal || element?.className || "");
  const rendered = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  };
  const inViewport = (element) => {
    if (!rendered(element)) return false;
    const rect = element.getBoundingClientRect();
    return rect.bottom > 0 && rect.top < innerHeight && rect.right > 0 && rect.left < innerWidth;
  };
  const tabLabel = (element) => {
    const match = norm(element?.textContent).match(/^(\u0110\u1eb7t h\u00e0ng)\s*(\d+)(?:\s*[\u00d7\u2715x])?$/i);
    return match ? `${match[1]} ${Number(match[2])}` : null;
  };
  const isInteractiveTab = (element, label) => {
    if (!element || !rendered(element) || tabLabel(element) !== label) return false;
    const role = String(element.getAttribute("role") || "").toLowerCase();
    const tag = element.tagName?.toLowerCase();
    return role === "tab" || /\b(active|selected|tab|cart|order)\b/i.test(classText(element))
      || ["a", "button", "li"].includes(tag) || element.hasAttribute("ng-click") || element.hasAttribute("onclick");
  };
  const getTabs = () => {
    const found = [...document.querySelectorAll("button,a,li,div,span")]
      .filter((element) => rendered(element) && element.getBoundingClientRect().top < 125)
      .map((element) => ({ element, label: tabLabel(element) }))
      .filter((entry) => entry.label);
    const byLabel = new Map();
    for (const entry of found) {
      const candidates = [];
      let current = entry.element;
      for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
        if (!isInteractiveTab(current, entry.label)) continue;
        let score = 0;
        if (String(current.getAttribute("role") || "").toLowerCase() === "tab") score += 1000;
        if (current.hasAttribute("ng-click") || current.hasAttribute("onclick")) score += 700;
        if (["a", "button", "li"].includes(current.tagName?.toLowerCase())) score += 500;
        if (/\b(tab|cart|order)\b/i.test(classText(current))) score += 300;
        candidates.push({ target: current, label: entry.label, score, depth });
      }
      const best = candidates.sort((left, right) => right.score - left.score || left.depth - right.depth)[0]
        || { target: entry.element, label: entry.label, score: 0, depth: 0 };
      const existing = byLabel.get(entry.label);
      if (!existing || best.score > existing.score || (best.score === existing.score && best.depth < existing.depth)) byLabel.set(entry.label, best);
    }
    return [...byLabel.values()].sort((left, right) => Number(left.label.match(/\d+$/)[0]) - Number(right.label.match(/\d+$/)[0]));
  };
  const assertSafeContext = () => {
    const branchText = (value) => norm(value).replace(/\s*select$/i, "").trim();
    const wanted = branchText(expectedBranch);
    const branchOk = [...document.querySelectorAll("button,a,li,div,span")]
      .filter(rendered)
      .some((element) => {
        const rect = element.getBoundingClientRect();
        const text = branchText(element.textContent);
        return text.length < 80 && (rect.top < 130 || rect.bottom > innerHeight - 120)
          && /^Đồng Nai\s*\d+\s*-\s*ĐN\d+/i.test(text)
          && text.startsWith(wanted);
      });
    if (!wanted || !branchOk) throw new Error(`KiotViet is not on the required branch ${expectedBranch}; final order was not clicked.`);
    const tabs = getTabs();
    if (!allowExistingInvoiceTabs && tabs.some((tab) => /^Hóa đơn/i.test(tab.label))) {
      throw new Error("KiotViet still has an Invoice tab; final order was not clicked.");
    }
    if (!tabs.some((tab) => /^Đặt hàng/i.test(tab.label))) throw new Error("KiotViet is not in Order mode; final order was not clicked.");
  };
  const activeMatches = (tab) => {
    const candidates = [tab.target, ...(tab.target.querySelectorAll?.('[aria-selected="true"],.active,.selected,.current') || [])];
    for (const candidate of candidates) {
      let current = candidate;
      for (let depth = 0; current && depth < 7; depth += 1, current = current.parentElement) {
        if (String(current.getAttribute("aria-selected")) === "true" || /(^|\s)(active|selected|current)(\s|$)/i.test(classText(current))) return true;
      }
    }
    return false;
  };
  const activate = async (label, code) => {
    const waitForExpectedNote = async (expectedCode, timeoutMs = 3000) => {
      const deadline = Date.now() + timeoutMs;
      while (Date.now() < deadline) {
        const current = readValue(noteField());
        if (current === expectedCode) return true;
        await sleepOnPage(120);
      }
      return false;
    };
    for (let attempt = 0; attempt < 5; attempt += 1) {
      const tab = getTabs().find((item) => item.label === label);
      // Numeric labels are not durable: KiotViet renumbers carts after each
      // successful order.  Even when the requested label still exists, the
      // durable identity is the exact note code, so verify it after every
      // activation instead of trusting the label alone.
      const candidates = tab
        ? [tab, ...getTabs().filter((item) => item !== tab).reverse()]
        : [...getTabs()].reverse();
      for (const candidate of candidates) {
        if (!activeMatches(candidate)) {
          candidate.target.scrollIntoView?.({ block: "nearest", inline: "nearest" });
          candidate.target.focus?.({ preventScroll: true });
          for (const type of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
            const EventType = type.startsWith("pointer") && typeof PointerEvent !== "undefined" ? PointerEvent : MouseEvent;
            candidate.target.dispatchEvent(new EventType(type, { bubbles: true, cancelable: true, view: window, button: 0, buttons: type.endsWith("down") ? 1 : 0 }));
          }
          candidate.target.click();
        }
        if (await waitForExpectedNote(code, 3000)) return candidate;
      }
      await sleepOnPage(250);
    }
    throw new Error(`Missing ${label} before final order.`);
  };
  const noteField = () => [...document.querySelectorAll('input[placeholder*="Ghi ch\u00fa \u0111\u01a1n h\u00e0ng" i],textarea[placeholder*="Ghi ch\u00fa \u0111\u01a1n h\u00e0ng" i],[contenteditable="true"][data-placeholder*="Ghi ch\u00fa \u0111\u01a1n h\u00e0ng" i]')].find(inViewport);
  const readValue = (element) => norm(element?.isContentEditable ? element.textContent : element?.value);
  const exactText = (element, expected) => norm(element?.textContent).toLocaleLowerCase("vi") === expected;
  const clickableParent = (element) => {
    let current = element;
    for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
      const tag = current.tagName?.toLowerCase();
      if (["button", "a"].includes(tag) || current.getAttribute("role") === "button" || current.hasAttribute("ng-click") || current.hasAttribute("onclick")) return current;
    }
    return element;
  };
  const finalOrderButton = () => {
    const expected = "\u0111\u1eb7t h\u00e0ng";
    const exact = document.querySelector("button#saveOrder.btn.btn-primary.btn-xl");
    if (exact && inViewport(exact) && exactText(exact, expected)
      && !exact.hasAttribute("disabled") && String(exact.getAttribute("aria-disabled")) !== "true") {
      return exact;
    }
    const candidates = [...document.querySelectorAll("button,a,div,span,[role=button]")]
      .filter(inViewport)
      .filter((element) => exactText(element, expected))
      .map(clickableParent)
      .filter((element, index, all) => all.indexOf(element) === index)
      .filter((element) => !element.hasAttribute("disabled") && String(element.getAttribute("aria-disabled")) !== "true")
      .filter((element) => {
        const rect = element.getBoundingClientRect();
        return rect.top > 110 && rect.width >= 42 && rect.height >= 22;
      })
      .sort((left, right) => {
        const a = left.getBoundingClientRect();
        const b = right.getBoundingClientRect();
        return (b.top + b.right / 10000) - (a.top + a.right / 10000);
      });
    return candidates[0] || null;
  };
  const clickLikeUser = (element) => {
    element.scrollIntoView?.({ block: "center", inline: "center" });
    element.focus?.({ preventScroll: true });
    element.click();
  };
  const suppressPrintPreview = () => {
    // KiotViet calls window.print() immediately after a successful final
    // order.  Printing is not part of this import workflow; replacing only
    // that browser-side side effect is equivalent to pressing Cancel in the
    // native preview, while keeping the completed order intact.
    try {
      window.print = () => { window.__orderBridgePrintSuppressed = true; };
      if (window.top && window.top !== window) {
        window.top.print = () => { window.__orderBridgePrintSuppressed = true; };
      }

      // KiotViet's `loadAndPrint` also creates a hidden /sale/ iframe.  On
      // Edge 151 that iframe can replace the sales canvas and keep the page
      // task pending even though the order was already saved.  Canceling the
      // native preview is safe, but preventing this internal print document
      // is safer and avoids a transport timeout followed by a duplicate
      // retry.  This guard is page-local and never touches the chat iframe.
      const isPrintFrame = (node) => {
        if (!node || String(node.tagName || '').toUpperCase() !== 'IFRAME') return false;
        const src = String(node.getAttribute?.('src') || node.src || '');
        return /\/sale\/(?:$|[#?])/i.test(src) || /\/sale\/?$/i.test(src);
      };
      const removePrintFrames = () => {
        for (const frame of [...document.querySelectorAll('iframe')]) {
          if (isPrintFrame(frame)) {
            try { frame.remove(); } catch (_) {}
            window.__orderBridgePrintFramesBlocked = (window.__orderBridgePrintFramesBlocked || 0) + 1;
          }
        }
      };
      removePrintFrames();
      if (window.__orderBridgePrintGuardVersion !== 2) {
        const oldAppend = Node.prototype.appendChild;
        const oldInsert = Node.prototype.insertBefore;
        Node.prototype.appendChild = function(node) {
          if (isPrintFrame(node)) {
            window.__orderBridgePrintFramesBlocked = (window.__orderBridgePrintFramesBlocked || 0) + 1;
            return node;
          }
          return oldAppend.call(this, node);
        };
        Node.prototype.insertBefore = function(node, reference) {
          if (isPrintFrame(node)) {
            window.__orderBridgePrintFramesBlocked = (window.__orderBridgePrintFramesBlocked || 0) + 1;
            return node;
          }
          return oldInsert.call(this, node, reference);
        };
        const observer = new MutationObserver(() => removePrintFrames());
        observer.observe(document.documentElement || document, { childList: true, subtree: true, attributes: true, attributeFilter: ['src'] });
        window.__orderBridgePrintGuardObserver = observer;
        window.__orderBridgePrintGuardVersion = 2;
      }
    } catch (_) {
      // The controller retains an OS-level Escape fallback after the click.
    }
  };
  const successPattern = /(?:\u0111\u1eb7t h\u00e0ng th\u00e0nh c\u00f4ng|t\u1ea1o \u0111\u01a1n th\u00e0nh c\u00f4ng|l\u01b0u th\u00e0nh c\u00f4ng|\u0111\u01a1n h\u00e0ng \u0111\u00e3 \u0111\u01b0\u1ee3c t\u1ea1o)/i;
  const successNoticeNodes = () => [...document.querySelectorAll('[role="alert"],.toast,.notification,.k-notification,.alert,[class*="toast" i],[class*="notify" i]')]
    .filter(rendered)
    .filter((element) => successPattern.test(norm(element.textContent)));
  const confirmationButton = () => {
    const dialog = [...document.querySelectorAll('[role="dialog"],.modal,.modal-dialog,[class*="dialog" i]')]
      .filter(rendered)
      .find((element) => /\u0111\u1eb7t h\u00e0ng/i.test(norm(element.textContent)));
    if (!dialog) return null;
    return [...dialog.querySelectorAll("button,[role=button],a")]
      .filter(inViewport)
      .map(clickableParent)
      .find((element) => /^(?:\u0111\u1eb7t h\u00e0ng|x\u00e1c nh\u1eadn|\u0111\u1ed3ng \u00fd)$/i.test(norm(element.textContent))) || null;
  };
  const waitForCompletion = async (label, code, before) => {
    const deadline = Date.now() + 45000;
    let clickedConfirmation = false;
    let sawTargetNote = false;
    while (Date.now() < deadline) {
      const current = getTabs().find((tab) => tab.label === label);
      if (!current) {
        return { success: true, confirmation: clickedConfirmation, proof: "tab_closed" };
      }
      const active = getTabs().find(activeMatches);
      const note = readValue(noteField());
      if (note === norm(code)) sawTargetNote = true;
      // KiotViet may renumber the surviving cart immediately after saving.
      // Do not require the old numeric label to remain present: once this
      // exact note was active, a changed/cleared note is completion proof.
      if (sawTargetNote && note !== norm(code)) {
        return { success: true, confirmation: clickedConfirmation, proof: "tab_reset" };
      }
      if (successNoticeNodes().some((node) => !before.successNotices.has(node))) {
        return { success: true, confirmation: clickedConfirmation, proof: "success_notice" };
      }
      const confirm = confirmationButton();
      if (confirm && !clickedConfirmation) {
        clickLikeUser(confirm);
        clickedConfirmation = true;
      }
      await sleepOnPage(220);
    }
    throw new Error(`No KiotViet confirmation after final order for ${code}; stopping without retry.`);
  };

  return (async () => {
    if (!Array.isArray(orderExpectations) || orderExpectations.length < 1 || orderExpectations.length > 19) throw new Error("Final-order batch must contain 1 to 19 orders.");
    assertSafeContext();
    const submitted = [];
    for (const expectation of [...orderExpectations].reverse()) {
      const label = norm(expectation.expectedTabLabel);
      const code = norm(expectation.orderCode);
      if (!label || !code) throw new Error("Missing tab label or order code before final order.");
      const activeTab = await activate(label, code);
      const note = noteField();
      if (!note || readValue(note) !== code) throw new Error(`Note in ${label} does not match ${code}; final order was not clicked.`);
      const button = finalOrderButton();
      if (!button) throw new Error(`Could not find a safe final-order button for ${code}.`);
      const before = { successNotices: new Set(successNoticeNodes()) };
      suppressPrintPreview();
      clickLikeUser(button);
      // Python starts a native Escape canceller in parallel before this page
      // task is invoked.  Once printing is dismissed (or suppressed), wait
      // until this exact tab is gone before allowing the next order.
      const completion = await waitForCompletion(label, code, before);
      submitted.push({
        orderCode: code,
        expectedTabLabel: label,
        actualTabLabel: String(activeTab?.label || label),
        tabLabel: String(activeTab?.label || label),
        status: "submitted",
        success: true,
        completionProof: String(completion?.proof || "verified"),
        confirmationClicked: Boolean(completion?.confirmation),
      });
    }
    return { ok: true, submitted };
  })().catch((error) => ({ error: String(error?.message || error) }));
}

async function inspectKiot(preferredWindowId = null) {
  const tab = await getOrCreateTab("nongsanlamhong.kiotviet.vn", KIOT_URL, preferredWindowId);
  await waitForComplete(tab.id);
  return { tabId: tab.id, ...(await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH])) };
}

async function prepareKiot(preferredWindowId = null) {
  const tab = await getOrCreateTab("nongsanlamhong.kiotviet.vn", KIOT_URL, preferredWindowId);
  await waitForComplete(tab.id);
  let initial = await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH]);
  if (initial.invoiceMode) {
    pushEvent("info", "Phát hiện tab Hóa đơn; đang chuyển sang Đặt hàng trước khi import.", initial);
    const conversion = await runInTab(tab.id, kiotConvertActiveInvoiceTask);
    if (conversion?.error) throw new Error(conversion.error);
    await sleep(260);

    // The icon can either convert immediately or open a short contextual
    // menu.  Select only the exact “Chuyển sang đặt hàng” action if shown.
    let converted = await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH]);
    if (converted.invoiceMode) {
      const menuSelection = await runInTab(tab.id, kiotChooseConvertInvoiceMenuTask);
      if (menuSelection?.error) throw new Error(menuSelection.error);
    }

    const conversionDeadline = Date.now() + 5000;
    while (Date.now() < conversionDeadline) {
      await sleep(250);
      converted = await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH]);
      if (converted.orderMode && !converted.invoiceMode) break;
    }
    if (!converted.orderMode || converted.invoiceMode) {
      throw new Error("Không chuyển được tab Hóa đơn sang Đặt hàng. Tool đã dừng an toàn, không điều hướng hoặc thay đổi dữ liệu khác.");
    }
    initial = converted;
    pushEvent("success", "Đã chuyển tab Hóa đơn sang Đặt hàng.", conversion);
  }
  if (!initial.orderMode) {
    await chrome.tabs.update(tab.id, { url: KIOT_URL, active: true });
    await waitForComplete(tab.id);
    await sleep(900);
    await runInTab(tab.id, kiotOpenOrdersMenuTask);
    const deadline = Date.now() + 3500;
    let reachedOrders = false;
    while (Date.now() < deadline) {
      const current = await chrome.tabs.get(tab.id);
      if (/\/man\/#\/Orders(?:$|[/?])/i.test(current.url || "")) {
        reachedOrders = true;
        break;
      }
      await sleep(250);
    }
    if (!reachedOrders) {
      // Some KiotViet Angular menu instances ignore programmatic clicks even
      // though the menu is visibly open.  Route directly to the same screen
      // as a safe, deterministic fallback, then use the verified order link.
      await chrome.tabs.update(tab.id, { url: KIOT_ORDERS_URL, active: true });
    }
    await waitForComplete(tab.id);
    await sleep(700);
  }
  await runInTab(tab.id, kiotPrepareTask, [EXPECTED_BRANCH]);
  await sleep(900);
  let status = await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH]);
  if (!status.branchOk) {
    await sleep(1400);
    status = await runInTab(tab.id, kiotInspectTask, [EXPECTED_BRANCH]);
  }
  if (!status.branchOk) throw new Error(`Chi nhánh hiện tại không phải ${EXPECTED_BRANCH}.`);
  if (!status.orderMode || status.invoiceMode) throw new Error("Chưa vào đúng màn hình/tab Đặt hàng.");
  pushEvent("success", "KiotViet đã đúng chi nhánh và chế độ Đặt hàng", status);
  return { ok: true, tabId: tab.id, ...status };
}

async function cleanKiotTabs(confirmCloseData, preferredWindowId = null) {
  const status = await prepareKiot(preferredWindowId);
  const result = await runInTab(status.tabId, kiotCleanTabsTask, [confirmCloseData]);
  await sleep(500);
  const prepared = await prepareKiot(preferredWindowId);
  if (prepared.orderTabs.length !== 1 || prepared.invoiceTabs.length !== 0) {
    throw new Error(`Sau khi dọn cần đúng 1 tab Đặt hàng trắng; hiện có ${prepared.orderTabs.length} Đặt hàng và ${prepared.invoiceTabs.length} Hóa đơn.`);
  }
  const response = { ...result, orderTabs: prepared.orderTabs, invoiceTabs: prepared.invoiceTabs };
  pushEvent("success", "Đã dọn các tab cũ và giữ 1 tab Đặt hàng", response);
  return response;
}

async function importBatch(payload, preferredWindowId = null) {
  if (!payload?.filePath || !Array.isArray(payload.orders) || payload.orders.length < 1 || payload.orders.length > 19) {
    throw new Error("Lô import phải có từ 1 đến 19 đơn và đường dẫn file hợp lệ.");
  }
  const status = await prepareKiot(preferredWindowId);
  if (status.orderTabs.length !== 1 || status.invoiceTabs.length !== 0) throw new Error(`Cần đúng 1 tab Đặt hàng trắng trước import; hiện có ${status.orderTabs.length} Đặt hàng và ${status.invoiceTabs.length} Hóa đơn.`);
  await runInTab(status.tabId, kiotOpenImportTask);
  await setFileInput(status.tabId, payload.filePath);
  const imported = await runInTab(status.tabId, kiotWaitImportTask, [payload.orders.length, status.orderTabs.length]);
  const notes = await runInTab(status.tabId, kiotFillNotesTask, [payload.orders]);
  pushEvent("success", "Đã import lô và điền ghi chú; chưa bấm ĐẶT HÀNG", { imported, notes });
  return { ok: true, imported, notes, finalOrderClicked: false };
}

async function fillNotes(payload, preferredWindowId = null) {
  if (!Array.isArray(payload?.orders) || payload.orders.length < 1 || payload.orders.length > 19) {
    throw new Error("Cần từ 1 đến 19 đơn để điền ghi chú.");
  }
  const status = await prepareKiot(preferredWindowId);
  if (status.invoiceTabs.length !== 0 || status.orderTabs.length < payload.orders.length + 1) {
    throw new Error("Số tab Đặt hàng hiện tại không phải lô import an toàn để điền ghi chú.");
  }
  const notes = await runInTab(status.tabId, kiotFillNotesTask, [payload.orders]);
  return { ok: true, notes, finalOrderClicked: false };
}

async function executeAction(action, payload, preferredWindowId = null) {
  switch (action) {
    case "PING":
      return { ok: true, version: chrome.runtime.getManifest().version };
    case "EXPORT_TODAY":
      return exportToday(preferredWindowId);
    case "INSPECT_SOPKA":
      return inspectSopoka(preferredWindowId);
    case "INSPECT_KIOT":
      return { ok: true, ...(await inspectKiot(preferredWindowId)) };
    case "PREPARE_KIOT":
      return prepareKiot(preferredWindowId);
    case "CLEAN_KIOT_TABS":
      return cleanKiotTabs(Boolean(payload?.confirmCloseData), preferredWindowId);
    case "IMPORT_BATCH":
      return importBatch(payload, preferredWindowId);
    case "FILL_NOTES":
      return fillNotes(payload, preferredWindowId);
    default:
      throw new Error(`Lệnh extension không được hỗ trợ: ${action}`);
  }
}

let activeBrowserAction = null;
let cancelActiveBrowserAction = null;

function cancelActiveAction() {
  if (!activeBrowserAction) {
    return { ok: true, cancelled: false, message: "Không có tác vụ Edge nào đang chạy." };
  }
  if (typeof cancelActiveBrowserAction !== "function") {
    return {
      ok: true,
      cancelled: false,
      action: activeBrowserAction,
      message: `Lệnh ${activeBrowserAction} không có điểm dừng an toàn ở thời điểm này.`,
    };
  }
  const action = activeBrowserAction;
  cancelActiveBrowserAction();
  pushEvent("warning", `Đã nhận yêu cầu dừng ${action}. Không có file nào sẽ được phân tích cho đến khi bạn chạy lại.`, { action });
  return { ok: true, cancelled: true, action };
}

async function handleAction(action, payload, preferredWindowId = null) {
  const normalizedAction = String(action || "").toUpperCase();
  if (normalizedAction === "PING") return executeAction(normalizedAction, payload, preferredWindowId);
  if (normalizedAction === "CANCEL_ACTIVE") return cancelActiveAction();
  if (activeBrowserAction) {
    throw new Error(`Edge đang chạy lệnh ${activeBrowserAction}; hãy chờ lệnh đó hoàn tất.`);
  }
  activeBrowserAction = normalizedAction;
  try {
    return await executeAction(normalizedAction, payload, preferredWindowId);
  } finally {
    cancelActiveBrowserAction = null;
    activeBrowserAction = null;
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  const senderUrl = sender.tab?.url || sender.url || "";
  if (!DASHBOARD_PREFIXES.some((prefix) => senderUrl.startsWith(prefix))) {
    sendResponse({ ok: false, error: "Chỉ dashboard cục bộ được phép điều khiển extension." });
    return false;
  }
  handleAction(message?.action, message?.payload || {}, sender.tab?.windowId ?? null)
    .then((result) => sendResponse(result))
    .catch((error) => {
      const messageText = String(error?.message || error);
      pushEvent("error", messageText);
      sendResponse({ ok: false, error: messageText });
    });
  return true;
});
