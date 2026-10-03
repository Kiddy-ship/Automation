const ui = {
  extensionStatus: document.querySelector("#extensionStatus"),
  automationModeBadge: document.querySelector("#automationModeBadge"),
  checkExtension: document.querySelector("#checkExtension"),
  checkBranch: document.querySelector("#checkBranch"),
  checkMode: document.querySelector("#checkMode"),
  checkAutomation: document.querySelector("#checkAutomation"),
  businessDate: document.querySelector("#businessDate"),
  businessDateHint: document.querySelector("#businessDateHint"),
  workflowDate: document.querySelector("#workflowDate"),
  primaryRunButton: document.querySelector("#primaryRunButton"),
  stopRunButton: document.querySelector("#stopRunButton"),
  runStatus: document.querySelector("#runStatus"),
  downloadButton: document.querySelector("#downloadButton"),
  sourceForm: document.querySelector("#sourceForm"),
  sourcePath: document.querySelector("#sourcePath"),
  purchaseForm: document.querySelector("#purchaseForm"),
  purchaseSourcePath: document.querySelector("#purchaseSourcePath"),
  purchaseBaselinePath: document.querySelector("#purchaseBaselinePath"),
  purchaseDate: document.querySelector("#purchaseDate"),
  purchaseImportButton: document.querySelector("#purchaseImportButton"),
  autoCreateMissingProducts: document.querySelector("#autoCreateMissingProducts"),
  purchaseSummary: document.querySelector("#purchaseSummary"),
  sourceBadge: document.querySelector("#sourceBadge"),
  sourceSummary: document.querySelector("#sourceSummary"),
  scanSummary: document.querySelector("#scanSummary"),
  sourceMode: document.querySelector("#sourceMode"),
  batchList: document.querySelector("#batchList"),
  batchCount: document.querySelector("#batchCount"),
  metricSource: document.querySelector("#metricSource"),
  metricSourceHint: document.querySelector("#metricSourceHint"),
  metricOrders: document.querySelector("#metricOrders"),
  metricLines: document.querySelector("#metricLines"),
  metricBatches: document.querySelector("#metricBatches"),
  metricDuplicates: document.querySelector("#metricDuplicates"),
  metricDuplicateHint: document.querySelector("#metricDuplicateHint"),
  metricCompletedOrders: document.querySelector("#metricCompletedOrders"),
  metricCompletedHint: document.querySelector("#metricCompletedHint"),
  inspectKiotButton: document.querySelector("#inspectKiotButton"),
  prepareKiotButton: document.querySelector("#prepareKiotButton"),
  cleanTabsButton: document.querySelector("#cleanTabsButton"),
  autoSubmitToggle: document.querySelector("#autoSubmitToggle"),
  autoSubmitLabel: document.querySelector("#autoSubmitLabel"),
  automationMode: document.querySelector("#automationMode"),
  clearLogButton: document.querySelector("#clearLogButton"),
  eventLog: document.querySelector("#eventLog"),
  workflowProgress: document.querySelector("#workflowProgress"),
  workflowSteps: [...document.querySelectorAll("#workflowSteps li")],
  runProgressPercent: document.querySelector("#runProgressPercent"),
  runProgressLabel: document.querySelector("#runProgressLabel"),
  runProgressDetail: document.querySelector("#runProgressDetail"),
  runProgressBar: document.querySelector("#runProgressBar"),
  runBatchProgress: document.querySelector("#runBatchProgress"),
  runBatchProgressLabel: document.querySelector("#runBatchProgressLabel"),
  runBatchProgressCounter: document.querySelector("#runBatchProgressCounter"),
  runBatchSegments: document.querySelector("#runBatchSegments"),
  runBatchProgressDetail: document.querySelector("#runBatchProgressDetail"),
  runResultSummary: document.querySelector("#runResultSummary"),
  liveAutomationProgress: document.querySelector("#liveAutomationProgress"),
  liveProgressPhase: document.querySelector("#liveProgressPhase"),
  liveProgressDetail: document.querySelector("#liveProgressDetail"),
  liveProgressCounter: document.querySelector("#liveProgressCounter"),
  liveProgressBar: document.querySelector("#liveProgressBar"),
  batchProgressSummary: document.querySelector("#batchProgressSummary"),
  duplicateModal: document.querySelector("#duplicateModal"),
  duplicateList: document.querySelector("#duplicateList"),
  stopDuplicateButton: document.querySelector("#stopDuplicateButton"),
  continueDuplicateButton: document.querySelector("#continueDuplicateButton"),
  cleanupModal: document.querySelector("#cleanupModal"),
  cancelCleanupButton: document.querySelector("#cancelCleanupButton"),
  confirmCleanupButton: document.querySelector("#confirmCleanupButton"),
  toastRegion: document.querySelector("#toastRegion"),
  workflowTabButton: document.querySelector("#workflowTabButton"),
  ledgerTabButton: document.querySelector("#ledgerTabButton"),
  workflowView: document.querySelector("#workflowView"),
  ledgerView: document.querySelector("#ledgerView"),
  ledgerDate: document.querySelector("#ledgerDate"),
  refreshLedgerButton: document.querySelector("#refreshLedgerButton"),
  ledgerSelectAll: document.querySelector("#ledgerSelectAll"),
  deleteLedgerButton: document.querySelector("#deleteLedgerButton"),
  ledgerCountLabel: document.querySelector("#ledgerCountLabel"),
  ledgerStatus: document.querySelector("#ledgerStatus"),
  ledgerTable: document.querySelector("#ledgerTable"),
};

const SUPPORTED_REPORT_FILE_RE = /\.(?:xls|xlsx|xlsm|xltx|xltm)$/i;
const SUPPORTED_REPORT_FILE_LABEL = ".xls, .xlsx, .xlsm, .xltx hoặc .xltm";

const appState = {
  extensionReady: false,
  extensionVersion: null,
  browserStatus: null,
  browserProgress: null,
  server: null,
  activeRun: null,
  pendingDuplicateRunId: null,
  pendingDuplicateScanId: null,
  pendingRequests: new Map(),
  localEvents: [],
  busy: false,
  busyButton: null,
  activeChromeAction: null,
  // A fetch timeout does not cancel the Flask/Edge task.  Keep the primary
  // button locked until a later status poll proves that the browser action
  // ended, so a user cannot accidentally queue a second export.
  timedOutExport: null,
  autoRunPending: false,
  workflowSession: null,
  // Progress is held per durable SQLite run. It prevents an old polling
  // snapshot from making the *whole workflow* bar appear to jump backward
  // while the next import file is being unlocked.
  workflowProgressByRun: new Map(),
  lastBatchTransitionByRun: new Map(),
  ledgerOrders: [],
};

const BUSINESS_TIME_ZONE = "Asia/Bangkok";
// The Selenium export has bounded waits for page navigation, report search,
// and Edge's native Office download control.  Its worst case is longer than
// the old 220-second UI request timeout.  Do not let the browser dashboard
// give up early and invite a second click while the first Edge action still
// owns the automation lock.
const EXPORT_TODAY_TIMEOUT_MS = 5 * 60 * 1000;

function bangkokBusinessDate(now = new Date()) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: BUSINESS_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(now);
  const value = Object.fromEntries(parts.filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  return `${value.year}-${value.month}-${value.day}`;
}

function formatDate(dateString) {
  const normalized = String(dateString || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!normalized) return "--/--/----";
  const [, year, month, day] = normalized;
  return `${day}/${month}/${year}`;
}

function isIsoBusinessDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
}

function operationalBusinessDate() {
  const browserToday = bangkokBusinessDate();
  const serverToday = appState.server?.today;
  // ``/api/state`` is polled, so its date can be up to one poll behind at
  // midnight.  Both dates use Asia/Bangkok; prefer the live local rendering
  // whenever the previous server snapshot belongs to yesterday.
  return isIsoBusinessDate(serverToday) && serverToday === browserToday
    ? serverToday
    : browserToday;
}

function selectedWorkflowDate() {
  return ui.workflowDate?.value || operationalBusinessDate();
}

function refreshBusinessDateChip(run = null, session = null) {
  const today = operationalBusinessDate();
  const runDate = run?.business_date || run?.businessDate;
  // A completed historical run must not pin the dashboard to yesterday.
  // Only an in-flight run may deliberately display its audited report date.
  const activeRunDate = isIsoBusinessDate(runDate) && isAutoRunActive(run);
  const displayDate = activeRunDate ? runDate : today;
  if (ui.businessDate) ui.businessDate.textContent = formatDate(displayDate);
  if (ui.businessDateHint) {
    ui.businessDateHint.textContent = activeRunDate && displayDate !== today
      ? `Lượt đang chạy · ${formatDate(displayDate)} · Asia/Bangkok`
      : "Asia/Bangkok · tự đổi lúc 00:00";
  }
  return { today, displayDate };
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
  }[character]));
}

function basename(path) {
  return String(path || "").split(/[\\/]/).pop() || "Chưa có";
}

function toast(message, level = "info") {
  const element = document.createElement("div");
  element.className = `toast ${level}`;
  element.textContent = message;
  ui.toastRegion.append(element);
  setTimeout(() => element.remove(), 5200);
}

function appendLog(event, persist = false) {
  const normalized = {
    level: event.level || "info",
    message: event.message || "Hoạt động mới",
    at: event.at || new Date().toISOString(),
    details: event.details || {},
  };
  appState.localEvents.unshift(normalized);
  appState.localEvents = appState.localEvents.slice(0, 80);
  renderEvents();
  if (persist) {
    api("/api/events", { method: "POST", body: normalized }).catch(() => {});
  }
}

function setRunStatus(message, tone = "info") {
  if (!ui.runStatus) return;
  ui.runStatus.textContent = message;
  ui.runStatus.className = `run-status ${tone}`;
}

function isCurrentSession(token) {
  return Boolean(token) && appState.workflowSession?.token === token && !appState.workflowSession?.cancelled;
}

function beginFreshWorkflow({ sourcePath = "", mode = "manual" } = {}) {
  const isManual = mode === "manual";
  const session = {
    token: crypto.randomUUID(),
    stage: 0,
    state: isManual ? "parsing" : "exporting",
    sourcePath,
    run: null,
    expectedRunId: null,
    scan: null,
    cancelled: false,
  };
  appState.workflowSession = session;
  appState.pendingDuplicateRunId = null;
  appendLog({ level: "info", message: isManual ? "Bắt đầu lượt xử lý file mới — tiến trình đã reset từ bước 1." : "Bắt đầu lượt tải Sopoka mới — tiến trình đã reset từ bước 1." });
  setRunStatus(isManual ? "Đang kiểm tra file report mới với SQLite." : "Đang điều khiển Edge mở Sopoka và tải report hôm nay.");
  renderState();
  return session;
}

function updateWorkflowSession(token, stage, message, state = null) {
  if (!isCurrentSession(token)) return false;
  appState.workflowSession.stage = stage;
  if (state) appState.workflowSession.state = state;
  if (message) setRunStatus(message, state === "error" ? "error" : state === "blocked" ? "warning" : "info");
  renderState();
  return true;
}

function finishWorkflowError(token, error, context = "Tác vụ") {
  if (!isCurrentSession(token)) return;
  const message = `${context} đã dừng. ${error.message}`;
  appState.workflowSession.stage = 0;
  appState.workflowSession.state = "error";
  setRunStatus(message, "error");
  appendLog({ level: "error", message });
  renderState();
}

function workflowErrorText(value) {
  const text = String(value ?? "");
  // Older failed runs can contain Vietnamese text which was accidentally
  // persisted as UTF-8 bytes interpreted by Windows-1252. Decode only when
  // that unmistakable signature is present; normal Unicode stays untouched.
  if (!/[ÃÂÆÄáº»]/.test(text) || typeof TextDecoder === "undefined") return text;
  try {
    const cp1252Byte = (character) => {
      const code = character.codePointAt(0);
      const specials = {
        0x20ac: 0x80, 0x201a: 0x82, 0x192: 0x83, 0x201e: 0x84,
        0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x2c6: 0x88,
        0x2030: 0x89, 0x160: 0x8a, 0x2039: 0x8b, 0x152: 0x8c,
        0x17d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93,
        0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
        0x2dc: 0x98, 0x2122: 0x99, 0x161: 0x9a, 0x203a: 0x9b,
        0x153: 0x9c, 0x17e: 0x9e, 0x178: 0x9f,
      };
      return code <= 0xff ? code : specials[code];
    };
    const bytes = Uint8Array.from([...text], cp1252Byte);
    if ([...bytes].some((byte) => !Number.isInteger(byte))) return text;
    const repaired = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return /\p{L}/u.test(repaired) ? repaired : text;
  } catch {
    return text;
  }
}

// Repair legacy UTF-8-as-CP1252 labels at the DOM boundary.  This keeps old
// event messages readable without changing SQLite data or user input values.
function repairUiText() {
  if (typeof TextDecoder === "undefined" || !document.body) return;
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    const repaired = workflowErrorText(node.nodeValue);
    if (repaired !== node.nodeValue) node.nodeValue = repaired;
  }
}

function renderEvents() {
  const events = [...appState.localEvents];
  if (!events.length) {
    ui.eventLog.innerHTML = '<div class="log-empty">Các bước chạy sẽ xuất hiện tại đây.</div>';
    return;
  }
  ui.eventLog.innerHTML = events.map((event) => {
    const time = new Date(event.at).toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" });
    return `<div class="log-item ${escapeHtml(event.level)}"><i></i><span>${escapeHtml(workflowErrorText(event.message))}</span><time>${time}</time></div>`;
  }).join("");
}

async function api(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
    body: options.body && typeof options.body !== "string" ? JSON.stringify(options.body) : options.body,
  });
  let payload;
  try { payload = await response.json(); } catch { payload = { ok: false, error: `HTTP ${response.status}` }; }
  const requiresDuplicateConfirmation = Boolean(
    payload.requiresDuplicateConfirmation || payload.requires_duplicate_confirmation,
  );
  // Direct Edge actions can return an explicit {ok:false} while HTTP itself
  // is technically successful.  Treat that as an error at the UI boundary;
  // otherwise EXPORT_TODAY would continue with an undefined source path and
  // misleadingly look like a SQLite/parsing failure.
  if ((!response.ok || payload.ok === false) && !requiresDuplicateConfirmation) {
    throw new Error(payload.message || payload.error || `HTTP ${response.status}`);
  }
  return payload;
}

async function chromeRequest(action, payload = {}, timeoutMs = 90000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await api("/api/browser/commands", {
      method: "POST",
      body: { action, payload },
      signal: controller.signal,
    });
  } catch (error) {
    if (error.name === "AbortError") throw new Error(`${action}: Edge WebDriver không phản hồi sau ${Math.ceil(timeoutMs / 1000)} giây.`);
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function isEdgeRequestTimeout(error) {
  return /kh[oô]ng ph[aả]n h[oồ]i sau \d+ gi[aâ]y/i.test(String(error?.message || ""));
}

function watchExportUntilBrowserStops(token, reason = "timeout") {
  appState.timedOutExport = { token, reason, startedAt: Date.now() };
  appState.activeChromeAction = reason === "cancelling" ? "CANCELLING" : "EXPORT_TODAY";
  // Keep the dashboard in its guarded state.  This prevents a second export
  // from being queued while the original synchronous server request may still
  // own Edge even though this page no longer awaits that response.
  setBusy(true, ui.primaryRunButton);
}

function reconcileWatchedExport(browserStatus) {
  const watch = appState.timedOutExport;
  if (!watch || !browserStatus) return;
  const progress = browserStatus.progress || browserStatus.browser_progress || null;
  if (progress?.active) return;

  // The UI response was lost, so it must never guess a source path or submit
  // anything.  Release the controls only after Edge has actually stopped and
  // tell the user that no SQLite/KiotViet work was automatically continued.
  appState.timedOutExport = null;
  appState.activeChromeAction = null;
  const session = appState.workflowSession;
  if (session?.token === watch.token) {
    session.stage = 0;
    session.state = watch.reason === "cancelling" ? "stopped" : "error";
  }
  const phase = statusKey(progress?.phase || progress?.status);
  const cancelled = watch.reason === "cancelling";
  const message = cancelled
    ? "Edge đã dừng tải report. Không có file nào được phân tích trong lượt này."
    : phase === "COMPLETED"
      ? "Edge đã tải xong sau khi UI mất kết nối, nhưng UI không nhận được đường dẫn file. Tool chưa phân tích hay import file này."
      : "Edge đã kết thúc sau khi UI mất kết nối. Tool chưa phân tích hay import dữ liệu nào trong lượt này.";
  setRunStatus(message, cancelled ? "warning" : "error");
  appendLog({ level: cancelled ? "warning" : "error", message });
  toast(message, cancelled ? "info" : "error");
  setBusy(false);
  renderState();
}

async function refreshBrowserStatus() {
  try {
    const status = await api("/api/browser/status");
    appState.browserStatus = status;
    appState.browserProgress = status.progress || status.browser_progress || null;
    appState.extensionReady = Boolean(status.available);
    appState.extensionVersion = status.started ? "đang mở" : "sẵn sàng";
  } catch {
    appState.browserStatus = null;
    appState.browserProgress = null;
    appState.extensionReady = false;
    appState.extensionVersion = null;
  }
  reconcileWatchedExport(appState.browserStatus);
  renderExtensionStatus();
  renderLiveAutomationProgress();
}

function renderExtensionStatus() {
  const online = appState.extensionReady;
  ui.extensionStatus.className = `status-pill ${online ? "status-online" : "status-offline"}`;
  ui.extensionStatus.querySelector("span").textContent = online ? `Edge WebDriver · ${appState.extensionVersion}` : "Chưa sẵn sàng Edge WebDriver";
  setCheck(ui.checkExtension, online ? "ok" : "pending", online ? "Điều khiển trực tiếp qua msedgedriver.exe" : "Không tìm thấy Edge hoặc msedgedriver.exe");
}

function setCheck(element, status, message) {
  element.className = `check-row ${status}`;
  element.querySelector(".check-dot").textContent = status === "ok" ? "✓" : status === "error" ? "×" : "";
  element.querySelector("small").textContent = message;
}

function setBusy(busy, button = null) {
  appState.busy = busy;
  const buttons = [ui.primaryRunButton, ui.downloadButton, ui.inspectKiotButton, ui.prepareKiotButton, ui.cleanTabsButton, ui.purchaseImportButton];
  buttons.filter(Boolean).forEach((item) => { item.disabled = busy; });
  document.querySelectorAll("[data-auto-run-batch],[data-retry-batch],[data-abort-run]").forEach((item) => { item.disabled = busy; });
  if (busy && button) appState.busyButton = button;
  const iconButton = appState.busyButton;
  if (iconButton) {
    const icon = iconButton.querySelector(".button-icon");
    if (icon) icon.classList.toggle("spin", busy);
  }
  if (!busy) appState.busyButton = null;
  if (ui.stopRunButton) {
    const canStop = Boolean(busy && ["EXPORT_TODAY", "EXPORT_DATE"].includes(appState.activeChromeAction));
    ui.stopRunButton.hidden = !canStop;
    ui.stopRunButton.disabled = !canStop;
  }
}

function normalizeRun(statePayload) {
  return statePayload.active_run || statePayload.activeRun || statePayload.run || null;
}

async function refreshState() {
  try {
    const state = await api("/api/state");
    appState.server = state;
    appState.activeRun = normalizeRun(state);
    if (appState.autoRunPending && !isAutoRunActive(appState.activeRun)) {
      appState.autoRunPending = false;
      setBusy(false);
    }
    const session = appState.workflowSession;
    if (session?.expectedRunId && String(appState.activeRun?.id) === String(session.expectedRunId)) {
      session.run = appState.activeRun;
      session.stage = workflowIndex(appState.activeRun);
    }
    const serverEvents = state.events || appState.activeRun?.events || [];
    if (serverEvents.length) {
      const known = new Set(appState.localEvents.map((event) => `${event.at}|${event.message}`));
      for (const event of serverEvents) {
        const key = `${event.at || event.created_at}|${event.message}`;
        if (!known.has(key)) appState.localEvents.push({ ...event, at: event.at || event.created_at });
      }
      appState.localEvents.sort((a, b) => new Date(b.at) - new Date(a.at));
    }
    renderState();
  } catch (error) {
    toast(`Không đọc được trạng thái: ${error.message}`, "error");
  }
}

function workflowIndex(run) {
  if (!run) return 0;
  const status = String(run.status || "").toUpperCase();
  if (["DOWNLOADING", "PARSING", "IDLE"].includes(status)) return 0;
  if (["DUPLICATE_REVIEW", "REVIEW_REQUIRED"].includes(status)) return 1;
  if (["PARSED", "PREPARING_BATCHES", "BATCHES_READY", "GENERATING", "PREPARED", "READY"].includes(status)) return 2;
  if (["BATCH_IN_PROGRESS", "KIOT_PRECHECK", "IMPORTING_BATCH", "FILLING_NOTES"].includes(status)) return 3;
  if (needsAutomationAttention(status)) return 3;
  if (["WAITING_USER_CONFIRM", "WAITING_USER_CONFIRMATION", "COMPLETE", "COMPLETED"].includes(status)) return 4;
  return 1;
}

function renderWorkflow(run) {
  const active = appState.workflowSession ? appState.workflowSession.stage : workflowIndex(run);
  ui.workflowProgress.style.width = `${active * 25}%`;
  ui.workflowSteps.forEach((step, index) => {
    step.classList.toggle("active", index === active);
    step.classList.toggle("done", index < active);
  });
}

function getBatches(run) {
  return run?.batches || [];
}

function getCount(run, keys, fallback = 0) {
  for (const key of keys) if (run?.[key] !== undefined && run?.[key] !== null) return run[key];
  return fallback;
}

function statusKey(value) {
  return String(value || "").trim().toUpperCase().replace(/[\s-]+/g, "_");
}

function toCount(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function isCompletedStatus(value) {
  return ["USER_CONFIRMED", "CONFIRMED", "DONE", "COMPLETE", "COMPLETED", "SUBMITTED", "ORDERED", "FINALIZED"].includes(statusKey(value));
}

function isWaitingCompletionStatus(value) {
  return ["WAITING_USER_CONFIRM", "WAITING_USER_CONFIRMATION", "NOTES_FILLED", "READY_TO_SUBMIT", "SUBMITTING"].includes(statusKey(value));
}

function needsAutomationAttention(value) {
  return ["AUTOMATION_ATTENTION", "AUTOMATION_FAILED", "PARTIAL_SUBMISSION"].includes(statusKey(value));
}

function isAutoRunActive(run) {
  const status = statusKey(run?.status);
  if (!status) return false;
  return !["COMPLETED", "FAILED", "CANCELLED", "EXPIRED", "AUTOMATION_ATTENTION"].includes(status);
}

function batchOrderStats(batch) {
  const orders = Array.isArray(batch?.orders) ? batch.orders : [];
  const total = toCount(batch?.order_count ?? batch?.sheet_count ?? orders.length);
  const completed = isCompletedStatus(batch?.status)
    ? total
    : orders.filter((order) => isCompletedStatus(order?.status)).length;
  const verified = orders.filter((order) => Boolean(order?.note_verified || order?.noteVerified)).length;
  return { total, completed, verified, pending: Math.max(0, total - completed) };
}

function batchIdentityValues(batch) {
  return [batch?.id, batch?.batch_id, batch?.batch_number, batch?.batch_no]
    .map((value) => String(value ?? "").trim())
    .filter(Boolean);
}

function isBrowserProgressForBatch(batch, progress = appState.browserProgress) {
  if (!progress?.active) return false;
  const phase = statusKey(progress.phase || progress.status);
  if (![
    "PREPARING_KIOT", "UPLOADING", "IMPORTING", "FILLING_NOTES", "NOTES", "SUBMITTING",
  ].includes(phase)) return false;
  const reported = String(progress.batch ?? progress.batchNumber ?? progress.batch_number ?? "").trim();
  if (!reported) return Boolean(batch?.is_current);
  const identities = batchIdentityValues(batch);
  if (identities.includes(reported)) return true;
  const reportedNumber = reported.match(/\d+/)?.[0];
  return Boolean(reportedNumber && identities.some((identity) => identity === reportedNumber));
}

function liveBatchProgressInfo(batch) {
  const progress = appState.browserProgress;
  if (!isBrowserProgressForBatch(batch, progress)) return null;
  const phase = statusKey(progress.phase || progress.status);
  const total = toCount(progress.total ?? progress.count);
  const current = Math.min(total || Number.MAX_SAFE_INTEGER, toCount(progress.current ?? progress.index ?? progress.completed));
  const ratio = total > 0 ? Math.min(1, current / total) : 0;
  const phaseInfo = {
    PREPARING_KIOT: [16, 0, "Đang kiểm tra KiotViet trước import"],
    UPLOADING: [36, 1, "Đang upload file vào KiotViet"],
    IMPORTING: [42, 1, "KiotViet đang tạo các tab Đặt hàng"],
    FILLING_NOTES: [Math.round(52 + ratio * 19), 2, `Đang điền ghi chú ${current}/${total || "?"} tab`],
    NOTES: [Math.round(58 + ratio * 16), 2, `Đang xác minh ghi chú ${current}/${total || "?"} tab`],
    SUBMITTING: [Math.round(76 + ratio * 22), 3, `Đang tự bấm ĐẶT HÀNG ${current}/${total || "?"}`],
  };
  const [percent, stage, fallback] = phaseInfo[phase] || [42, 1, "Edge đang xử lý file import"];
  return {
    percent: Math.min(98, Math.max(1, percent)),
    stage,
    label: progress.message || fallback,
    tone: "waiting",
    stats: batchOrderStats(batch),
    isLive: true,
  };
}

function currentBatchContext(batches) {
  const entries = batches.map((batch, index) => ({ batch, index, info: batchProgressInfo(batch) }));
  const live = entries.find(({ batch }) => isBrowserProgressForBatch(batch));
  const explicit = entries.find(({ batch }) => batch?.is_current === true);
  const active = entries.find(({ batch }) => {
    const status = statusKey(batch?.status);
    return !isCompletedStatus(status) && status !== "BLOCKED";
  });
  const selected = live || explicit || active || null;
  return {
    ...selected,
    total: entries.length,
    completed: entries.filter(({ batch }) => isCompletedStatus(batch?.status)).length,
    entries,
  };
}

function overallProgressInfo(run, batches) {
  const context = currentBatchContext(batches);
  const totalWeight = Math.max(1, context.entries.reduce((sum, { batch }) => sum + Math.max(1, batchOrderStats(batch).total), 0));
  const weightedProgress = context.entries.reduce(
    (sum, { batch, info }) => sum + info.percent * Math.max(1, batchOrderStats(batch).total),
    0,
  );
  const calculated = statusKey(run?.status) === "COMPLETED" ? 100 : Math.round(weightedProgress / totalWeight);
  const runKey = run?.id ?? run?.run_id;
  const previous = runKey ? appState.workflowProgressByRun.get(String(runKey)) : null;
  // The durable state can be observed between two SQLite transitions. Do not
  // make the total bar move backwards just because the next batch starts at 0.
  const percent = statusKey(run?.status) === "COMPLETED"
    ? 100
    : Math.max(previous?.percent ?? 0, calculated);
  if (runKey) appState.workflowProgressByRun.set(String(runKey), { percent, updatedAt: Date.now() });
  return { percent, calculated, context };
}

function batchProgressInfo(batch) {
  const live = liveBatchProgressInfo(batch);
  if (live) return live;
  const status = statusKey(batch?.status || batch?.display_status || "READY");
  const stats = batchOrderStats(batch);
  const noteRatio = stats.total ? stats.verified / stats.total : 0;
  let percent = 9;
  let stage = 0;
  let label = "File import đã sẵn sàng";
  let tone = "ready";

  if (isCompletedStatus(status)) {
    percent = 100;
    stage = 4;
    label = "Đã hoàn tất và đối soát";
    tone = "done";
  } else if (needsAutomationAttention(status)) {
    percent = Math.max(12, Math.round(42 + noteRatio * 38));
    stage = stats.verified ? 3 : 2;
    label = "Cần kiểm tra: kết quả đặt hàng chưa đầy đủ";
    tone = "error";
  } else if (status === "FAILED") {
    percent = Math.max(12, Math.round(28 + noteRatio * 38));
    stage = stats.verified ? 2 : 1;
    label = "Có lỗi — cần xử lý lại lô";
    tone = "error";
  } else if (status === "BLOCKED" || batch?.is_unlocked === false) {
    percent = 0;
    stage = 0;
    label = "Chờ lô trước hoàn tất";
    tone = "locked";
  } else if (isWaitingCompletionStatus(status)) {
    percent = Math.max(78, Math.round(72 + noteRatio * 18));
    stage = 3;
    label = "Đã xác minh ghi chú — chờ hoàn tất";
    tone = "waiting";
  } else if (["NOTES_IN_PROGRESS", "FILLING_NOTES"].includes(status)) {
    percent = Math.round(52 + noteRatio * 25);
    stage = 2;
    label = `Đang xác minh ghi chú ${stats.verified}/${stats.total} tab`;
    tone = "waiting";
  } else if (["IMPORTED"].includes(status)) {
    percent = Math.max(48, Math.round(48 + noteRatio * 18));
    stage = 2;
    label = "Đã import — đang đối chiếu ghi chú";
    tone = "waiting";
  } else if (["IMPORTING", "IMPORT_STARTED", "KIOT_PRECHECK"].includes(status)) {
    percent = 31;
    stage = 1;
    label = "Đang import sang KiotViet";
    tone = "waiting";
  } else if (["PENDING", "BUILDING"].includes(status)) {
    percent = 5;
    stage = 0;
    label = "Đang tạo file import";
    tone = "waiting";
  }

  return { percent: Math.min(100, Math.max(0, percent)), stage, label, tone, stats };
}

function orderProgressInfo(order, batch) {
  const batchStatus = statusKey(batch?.status);
  const orderStatus = statusKey(order?.status);
  const verified = Boolean(order?.note_verified || order?.noteVerified);
  if (isCompletedStatus(batchStatus) || isCompletedStatus(orderStatus)) {
    return { label: "Hoàn tất", className: "is-done", detail: "Đã đối soát" };
  }
  if (needsAutomationAttention(batchStatus) || needsAutomationAttention(orderStatus)) {
    return { label: "Cần kiểm tra", className: "is-error", detail: "Kết quả đặt hàng chưa được xác minh" };
  }
  if (batchStatus === "FAILED") {
    return { label: "Cần xử lý lại", className: "is-error", detail: "Lô đang có lỗi" };
  }
  if (isWaitingCompletionStatus(batchStatus)) {
    return { label: "Chờ hoàn tất", className: "is-warning", detail: verified ? "Ghi chú đã xác minh" : "Đang đối soát" };
  }
  if (verified || ["NOTE_FILLED", "NOTES_FILLED"].includes(orderStatus)) {
    return { label: "Đã ghi chú", className: "is-active", detail: "Đang hoàn tất lô" };
  }
  if (["NOTES_IN_PROGRESS", "FILLING_NOTES", "IMPORTED"].includes(batchStatus)) {
    return { label: "Đang ghi chú", className: "is-active", detail: "Đang cập nhật tab" };
  }
  if (["IMPORTING", "IMPORT_STARTED", "KIOT_PRECHECK"].includes(batchStatus)) {
    return { label: "Đang import", className: "is-active", detail: "Đang mở tab KiotViet" };
  }
  if (batchStatus === "BLOCKED" || batch?.is_unlocked === false) {
    return { label: "Chờ lô trước", className: "", detail: "Chưa được mở khóa" };
  }
  return { label: "Sẵn sàng", className: "", detail: "Chờ import" };
}

function renderAutomationMode(run) {
  const browser = appState.browserStatus || {};
  const values = [
    run?.auto_submit_enabled,
    run?.autoSubmitEnabled,
    appState.server?.auto_submit_enabled,
    appState.server?.autoSubmitEnabled,
    browser.auto_submit_enabled,
    browser.autoSubmitEnabled,
  ];
  const automatic = values.some((value) => value === true || value === 1 || value === "true");
  const modeText = automatic ? "Tự động toàn trình" : "Kiểm tra cuối";
  const detail = automatic
    ? "Tool sẽ tự hoàn tất theo cấu hình workflow."
    : "Workflow hiện vẫn dừng ở bước hoàn tất để đối soát.";

  if (ui.automationModeBadge) {
    ui.automationModeBadge.className = `mode-badge ${automatic ? "mode-auto" : "mode-review"}`;
    const label = ui.automationModeBadge.querySelector("span:last-child");
    if (label) label.textContent = automatic ? "Tự động toàn trình" : "Chờ kiểm tra cuối";
    ui.automationModeBadge.title = detail;
  }
  if (ui.automationMode) ui.automationMode.textContent = modeText;
  if (ui.autoSubmitToggle) ui.autoSubmitToggle.checked = automatic;
  if (ui.autoSubmitLabel) ui.autoSubmitLabel.textContent = automatic ? "Bật · tự hoàn tất theo workflow" : "Tắt · chờ hoàn tất lô";
  if (ui.checkAutomation) setCheck(ui.checkAutomation, automatic ? "ok" : "pending", automatic ? "Tool sẽ tự hoàn tất theo cấu hình" : "Workflow đang chờ cấu hình hoàn tất tự động");
}

function renderRunMonitor(run, session, batches) {
  let percent = 0;
  let label = "Chờ bắt đầu lượt xử lý mới";
  let detail = "Kết quả sẽ được cập nhật theo từng file báo cáo và từng lô import.";
  let result = "Chưa có kết quả workflow để tổng hợp.";
  let tone = "neutral";

  if (!run && session) {
    const state = statusKey(session.state);
    const sessionProgress = {
      EXPORTING: [8, "Đang tải Báo cáo tổng hợp bán hàng", "Edge đang mở Sopoka và lấy file Excel hôm nay."],
      PARSING: [25, "Đang đối chiếu dữ liệu với SQLite", "Chỉ các mã đơn hợp lệ và chưa hoàn tất trong ngày mới được giữ lại."],
      PREPARED: [44, "Đã tạo file import", "Đang chờ trạng thái lô được đồng bộ từ SQLite."],
      IMPORTING: [58, "Đang import KiotViet", "Tool đang chuẩn bị tab và ghi chú tương ứng cho từng mã đơn."],
      BLOCKED: [25, "Workflow đang dừng an toàn", "Có dữ liệu cần đối soát trước khi tiếp tục."],
      ERROR: [0, "Workflow đã dừng do lỗi", "Xem nhật ký thời gian thực để biết bước cần xử lý."],
      STOPPED: [0, "Workflow đã được dừng", "Không có thao tác mới nào được gửi sang KiotViet."],
      COMPLETE: [100, "Không còn đơn mới cần xử lý", "Lượt quét đã hoàn tất an toàn."],
    };
    const current = sessionProgress[state] || [Math.min(95, Math.max(0, toCount(session.stage) * 25)), "Đang chuẩn bị workflow", "Đang đồng bộ trạng thái tác vụ."];
    [percent, label, detail] = current;
    result = state === "ERROR" ? "Lượt xử lý chưa hoàn thành. Không ghi nhận kết quả mới." : "Lượt xử lý mới đang hoạt động.";
    tone = state === "ERROR" ? "error" : state === "BLOCKED" || state === "STOPPED" ? "warning" : "neutral";
  } else if (run) {
    const totalOrders = toCount(getCount(run, ["order_count", "orders_count", "orderCount"], 0));
    const totalBatches = batches.length;
    const overall = overallProgressInfo(run, batches);
    const { context } = overall;
    const completedBatches = context.completed;
    const completedOrders = batches.reduce((sum, batch) => sum + batchOrderStats(batch).completed, 0);
    const status = statusKey(run.status);
    percent = totalBatches ? overall.percent : Math.max(24, workflowIndex(run) * 25);
    const currentBatch = context.batch;
    const currentNo = currentBatch?.batch_number ?? currentBatch?.batch_no ?? (context.index !== undefined ? context.index + 1 : null);
    const currentPercent = context.info?.percent;
    label = status === "COMPLETED"
      ? "Workflow đã hoàn tất"
      : currentNo
        ? `Đang xử lý lô ${currentNo}/${totalBatches}`
        : "Đã phân tích file, đang chuẩn bị lô";
    detail = currentNo
      ? `${totalOrders.toLocaleString("vi-VN")} mã đơn · ${completedBatches}/${totalBatches} lô hoàn tất · lô ${currentNo}/${totalBatches} hiện ${currentPercent ?? 0}%. Thanh tổng giữ tiến độ xuyên suốt, không reset khi chuyển lô.`
      : `${totalOrders.toLocaleString("vi-VN")} mã đơn · ${totalBatches} lô · ${completedBatches}/${totalBatches} lô đã hoàn tất.`;
    if (needsAutomationAttention(status)) {
      label = "Cần kiểm tra kết quả tự động";
      detail = workflowErrorText(run.error_message) || "Có thao tác đặt hàng chưa được xác minh đầy đủ. Tool đã dừng, không tự thử lại và chưa ghi mã vào SQLite.";
      tone = "error";
      result = "Cần kiểm tra thủ công lô được đánh dấu đỏ; các lô sau vẫn bị khóa.";
    } else if (status === "FAILED") {
      label = "Workflow cần xử lý lỗi";
      detail = run.error_message || "Một lô chưa hoàn tất. Xem trạng thái chi tiết bên dưới.";
      tone = "error";
    } else if (["DUPLICATE_REVIEW", "WAITING_USER_CONFIRMATION"].includes(status)) {
      tone = "warning";
    } else if (status === "COMPLETED") {
      tone = "success";
    }
    if (!needsAutomationAttention(status)) {
      result = status === "COMPLETED"
        ? `Hoàn tất ${completedOrders}/${totalOrders} mã đơn và ${completedBatches}/${totalBatches} lô.`
        : `Đã hoàn tất ${completedOrders}/${totalOrders} mã đơn · còn ${Math.max(0, totalOrders - completedOrders)} mã đang chờ.`;
    }
  }

  if (ui.runProgressPercent) ui.runProgressPercent.textContent = `${percent}%`;
  if (ui.runProgressLabel) ui.runProgressLabel.textContent = label;
  if (ui.runProgressDetail) ui.runProgressDetail.textContent = detail;
  if (ui.runProgressBar) ui.runProgressBar.style.width = `${percent}%`;
  if (ui.runResultSummary) ui.runResultSummary.innerHTML = `<span class="result-dot ${tone}"></span><span>${escapeHtml(result)}</span>`;
}

function renderRunBatchProgress(run, batches) {
  if (!ui.runBatchProgress || !ui.runBatchSegments) return;
  if (!run || !batches.length) {
    ui.runBatchProgress.classList.add("hidden");
    return;
  }

  const overall = overallProgressInfo(run, batches);
  const { context } = overall;
  const current = context.batch;
  const currentNumber = current?.batch_number ?? current?.batch_no ?? (context.index !== undefined ? context.index + 1 : null);
  const runKey = String(run.id ?? run.run_id ?? "");
  const currentKey = current ? String(current.id ?? current.batch_id ?? currentNumber ?? "") : "complete";
  const previousKey = runKey ? appState.lastBatchTransitionByRun.get(runKey) : null;
  let transitionText = "";

  if (runKey && !previousKey) {
    appState.lastBatchTransitionByRun.set(runKey, currentKey);
  } else if (runKey && previousKey !== currentKey) {
    appState.lastBatchTransitionByRun.set(runKey, currentKey);
    if (currentNumber) {
      transitionText = `Đã chuyển sang lô ${currentNumber}/${context.total}. `;
      appendLog({ level: "info", message: `${transitionText}Thanh tổng workflow tiếp tục, không reset.` });
    }
  }

  const status = statusKey(run.status);
  const currentInfo = context.info;
  const caption = status === "COMPLETED"
    ? `Đã hoàn tất ${context.completed}/${context.total} lô`
    : currentNumber
      ? `Đang xử lý lô ${currentNumber}/${context.total}`
      : `${context.completed}/${context.total} lô đã hoàn tất`;
  const detail = needsAutomationAttention(status)
    ? "Workflow cần kiểm tra ở lô màu đỏ; các lô sau vẫn bị khóa và thanh tổng giữ nguyên tiến độ đã đạt."
    : status === "COMPLETED"
      ? `Đã hoàn tất tất cả ${context.total} file import. Thanh tổng đạt 100%.`
      : currentNumber
        ? `${transitionText}Lô ${currentNumber}/${context.total} hiện ${currentInfo?.percent ?? 0}% · ${context.completed}/${context.total} lô đã đối soát. Thanh tổng ${overall.percent}% không reset khi đổi file.`
        : `${context.completed}/${context.total} lô đã đối soát. Thanh tổng ${overall.percent}% vẫn giữ kết quả các lô trước.`;

  ui.runBatchProgress.classList.remove("hidden");
  if (ui.runBatchProgressLabel) ui.runBatchProgressLabel.textContent = `${context.total} file import · tiến độ theo lô`;
  if (ui.runBatchProgressCounter) ui.runBatchProgressCounter.textContent = status === "COMPLETED" ? `${context.total}/${context.total} lô` : currentNumber ? `Lô ${currentNumber}/${context.total}` : `${context.completed}/${context.total} lô`;
  if (ui.runBatchProgressDetail) ui.runBatchProgressDetail.textContent = detail;
  ui.runBatchSegments.style.setProperty("--batch-count", String(Math.max(1, context.total)));
  ui.runBatchSegments.innerHTML = context.entries.map(({ batch, index, info }) => {
    const batchNumber = batch?.batch_number ?? batch?.batch_no ?? index + 1;
    const isCurrent = current === batch;
    const isDone = isCompletedStatus(batch?.status);
    const className = [
      "run-batch-segment",
      isDone ? "is-done" : "",
      isCurrent ? "is-current" : "",
      info.tone === "error" ? "is-error" : "",
    ].filter(Boolean).join(" ");
    const text = isDone ? `Lô ${batchNumber} · xong` : isCurrent ? `Lô ${batchNumber} · ${info.percent}%` : `Lô ${batchNumber} · chờ`;
    return `<span class="${className}" title="${escapeHtml(text)}"><i style="width:${info.percent}%"></i><span>${escapeHtml(text)}</span></span>`;
  }).join("");
}

function renderLiveAutomationProgress() {
  const progress = appState.browserProgress;
  if (!ui.liveAutomationProgress) return;
  const phase = statusKey(progress?.phase || progress?.status || "");
  const inactive = !progress || ["", "IDLE", "READY", "NONE", "COMPLETE", "COMPLETED", "DONE"].includes(phase);
  if (inactive) {
    ui.liveAutomationProgress.classList.add("hidden");
    return;
  }
  const current = toCount(progress.current ?? progress.index ?? progress.completed);
  const total = toCount(progress.total ?? progress.count);
  const rawPercent = toCount(progress.percent ?? progress.progress);
  const percent = total > 0 ? Math.round((current / total) * 100) : Math.min(100, rawPercent);
  const phaseNames = {
    EXPORTING: "Edge đang tải báo cáo Sopoka",
    DOWNLOADING: "Edge đang tải file Excel",
    PARSING: "Đang phân tích file Excel",
    IMPORTING: "Đang import file vào KiotViet",
    FILLING_NOTES: "Đang điền ghi chú đơn hàng",
    NOTES: "Đang xác minh ghi chú",
    SUBMITTING: "Đang hoàn tất đơn hàng",
    PREPARING: "Đang chuẩn bị KiotViet",
  };
  const label = progress.label || phaseNames[phase] || `Edge đang chạy: ${phase.toLowerCase().replace(/_/g, " ")}`;
  const orderCode = progress.orderCode || progress.order_code;
  const detail = progress.message || (orderCode ? `Đang xử lý mã ${orderCode}` : "Đang đồng bộ thao tác trình duyệt.");
  ui.liveAutomationProgress.classList.remove("hidden");
  ui.liveAutomationProgress.dataset.phase = phase.toLowerCase();
  if (ui.liveProgressPhase) ui.liveProgressPhase.textContent = label;
  if (ui.liveProgressDetail) ui.liveProgressDetail.textContent = detail;
  if (ui.liveProgressCounter) ui.liveProgressCounter.textContent = total ? `${Math.min(current, total)}/${total}` : `${percent}%`;
  if (ui.liveProgressBar) ui.liveProgressBar.style.width = `${percent}%`;
}

function renderState() {
  const session = appState.workflowSession;
  // While a fresh cycle is running, deliberately do not repaint the prior
  // completed run on top of the new progress display.
  const run = session ? session.run : appState.activeRun;
  const runDate = run?.business_date || run?.businessDate;
  const { today: operationalDate } = refreshBusinessDateChip(run, session);
  // A source/run keeps its own report date for audit text below. The top chip
  // intentionally stays on the *operational* Bangkok date after midnight,
  // even when the latest completed run belongs to the previous day.
  const date = isIsoBusinessDate(runDate) ? runDate : operationalDate;
  renderWorkflow(run);

  const sourcePath = run?.source_path || run?.sourcePath || session?.sourcePath || "";
  const orderCount = getCount(run, ["order_count", "orders_count", "orderCount"], run?.orders?.length || 0);
  const lineCount = getCount(run, ["line_count", "lines_count", "lineCount"], run?.orders?.reduce((sum, order) => sum + (order.items?.length || order.line_count || 0), 0) || 0);
  const batches = getBatches(run);
  const duplicates = run?.duplicates || run?.duplicate_orders || [];
  const completedOrders = batches.reduce((sum, batch) => sum + batchOrderStats(batch).completed, 0);
  const completedBatches = batches.filter((batch) => isCompletedStatus(batch.status)).length;

  ui.metricSource.textContent = sourcePath ? basename(sourcePath) : "Chưa có";
  ui.metricSource.title = sourcePath;
  ui.metricSourceHint.textContent = sourcePath ? `Ngày ${formatDate(date)}` : "Chờ tải báo cáo";
  ui.metricOrders.textContent = orderCount.toLocaleString("vi-VN");
  ui.metricLines.textContent = `${lineCount.toLocaleString("vi-VN")} dòng sản phẩm`;
  ui.metricBatches.textContent = batches.length.toLocaleString("vi-VN");
  ui.metricDuplicates.textContent = duplicates.length.toLocaleString("vi-VN");
  ui.metricDuplicateHint.textContent = duplicates.length ? "Cần xác nhận thủ công" : "Chưa phát hiện";
  if (ui.metricCompletedOrders) ui.metricCompletedOrders.textContent = completedOrders.toLocaleString("vi-VN");
  if (ui.metricCompletedHint) ui.metricCompletedHint.textContent = `${completedOrders.toLocaleString("vi-VN")}/${toCount(orderCount).toLocaleString("vi-VN")} đơn · ${completedBatches}/${batches.length} lô`;
  ui.batchCount.textContent = batches.length;

  if (run) {
    const runId = run.id ?? run.run_id;
    const runStatus = String(run.status || "đang chuẩn bị").replace(/_/g, " ");
    const attention = needsAutomationAttention(run.status);
    ui.sourceBadge.className = `tag ${attention ? "error" : duplicates.length ? "warning" : "success"}`;
    ui.sourceBadge.textContent = attention ? "Cần kiểm tra" : duplicates.length ? "Có dữ liệu trùng" : "Đã kiểm tra";
    ui.sourceSummary.classList.remove("hidden");
    ui.sourceSummary.innerHTML = `<b>${runId ? `Run #${escapeHtml(runId)} · ` : ""}${orderCount} đơn · ${lineCount} dòng sản phẩm</b><br>File: ${escapeHtml(sourcePath)}<br>Trạng thái: ${escapeHtml(runStatus)} · SQLite kiểm tra theo ngày nghiệp vụ ${formatDate(date)}.`;
    ui.sourcePath.value = sourcePath;
  } else if (session) {
    const sessionState = String(session.state || "");
    const isError = sessionState === "error";
    const isBlocked = sessionState === "blocked";
    const isStopped = sessionState === "stopped";
    ui.sourceBadge.className = `tag ${isError || isBlocked || isStopped ? "warning" : "neutral"}`;
    ui.sourceBadge.textContent = isError || isStopped ? "Đã dừng" : isBlocked ? "Cần đối soát" : sessionState === "parsing" ? "Đang kiểm tra" : "Đang tải";
    ui.sourceSummary.classList.remove("hidden");
    ui.sourceSummary.innerHTML = sourcePath
      ? `<b>Đang xử lý lượt mới</b><br>File: ${escapeHtml(sourcePath)}<br>Chưa dùng dữ liệu/lô của lượt cũ.`
      : "<b>Đang chờ file report mới từ Sopoka.</b><br>Chưa có dữ liệu nào được gửi sang SQLite.";
    ui.sourcePath.value = sourcePath;
  } else {
    ui.sourceBadge.className = "tag neutral";
    ui.sourceBadge.textContent = "Chưa xử lý";
    ui.sourceSummary.classList.add("hidden");
  }
  const scan = session ? (session.scan || null) : (appState.server?.scan || appState.server?.last_scan || null);
  if (scan) {
    const status = String(scan.status || "");
    const blocked = Boolean(scan.blocked) || ["unresolved", "changed_confirmed", "active_workflow"].includes(status);
    const noNew = status === "no_new";
    ui.scanSummary.className = `scan-summary ${blocked ? "error" : noNew ? "warning" : "ok"}`;
    ui.scanSummary.classList.remove("hidden");
    const kept = Number(scan.new_order_count || 0);
    const skipped = Number(scan.skipped_confirmed_count || 0);
    const unresolved = Number(scan.unresolved_count || 0);
    const changed = Number(scan.changed_confirmed_count || 0);
    const wording = blocked
      ? `<b>Đang dừng an toàn.</b> ${unresolved} mã chưa được xác nhận${changed ? ` · ${changed} mã đã xác nhận nhưng dữ liệu thay đổi` : ""}. Không tạo file/import mới.`
      : noNew
        ? `<b>Không có đơn mới.</b> Đã bỏ qua an toàn ${skipped} mã đã xác nhận.`
        : `<b>Bộ dữ liệu làm việc: ${kept} mã mới.</b> Đã bỏ qua an toàn ${skipped} mã đã xác nhận.`;
    ui.scanSummary.innerHTML = `${wording}<br><span>Report: ${Number(scan.source_order_count || 0)} mã · audit: ${escapeHtml(scan.audit_path || "đang lưu")}</span>`;
    if (blocked) {
      ui.sourceBadge.className = "tag warning";
      ui.sourceBadge.textContent = "Cần đối soát";
    }
  } else {
    ui.scanSummary.classList.add("hidden");
  }
  if (ui.sourceMode) {
    ui.sourceMode.title = "Luồng mặc định mở Edge, tải Báo cáo tổng hợp bán hàng hôm nay rồi kiểm tra file trước khi xử lý.";
    ui.sourceMode.textContent = "Tự động · Edge / Sopoka hôm nay";
  }
  renderAutomationMode(run);
  renderRunMonitor(run, session, batches);
  renderRunBatchProgress(run, batches);
  renderLiveAutomationProgress();
  renderBatches(batches, run);
  renderEvents();
  repairUiText();
}

function batchStatusInfo(statusValue) {
  const status = statusKey(statusValue || "READY");
  if (isCompletedStatus(status)) return { label: "Đã hoàn tất", className: "done" };
  if (needsAutomationAttention(status)) return { label: "Cần kiểm tra thủ công", className: "error" };
  if (["WAITING_USER_CONFIRM", "WAITING_USER_CONFIRMATION", "NOTES_FILLED", "READY_TO_SUBMIT"].includes(status)) return { label: "Chờ hoàn tất", className: "waiting" };
  if (["IMPORTING", "IMPORTED", "NOTES_IN_PROGRESS", "FILLING_NOTES", "SUBMITTING"].includes(status)) return { label: "Đang xử lý", className: "waiting" };
  if (["FAILED"].includes(status)) return { label: "Cần xử lý lỗi", className: "error" };
  if (["BLOCKED"].includes(status)) return { label: "Chờ xác nhận lô trước", className: "locked" };
  if (["PENDING", "BUILDING"].includes(status)) return { label: "Đang tạo file", className: "waiting" };
  return { label: "Sẵn sàng import", className: "ready" };
}

function renderBatchStages(info) {
  const stages = ["Tệp", "Import", "Ghi chú", "Hoàn tất"];
  return stages.map((label, index) => {
    let className = "";
    if (info.tone === "error" && index === Math.min(3, info.stage)) className = "error";
    else if (info.stage >= 4 || index < info.stage) className = "done";
    else if (index === info.stage && info.tone !== "locked") className = "active";
    return `<span class="batch-stage ${className}"><i></i>${label}</span>`;
  }).join("");
}

function renderBatchOrders(batch, deletionEnabled = false) {
  const orders = Array.isArray(batch.orders) ? batch.orders : [];
  if (!orders.length) return "";
  const verified = batchOrderStats(batch).verified;
  const rows = orders.map((order, index) => {
    const progress = orderProgressInfo(order, batch);
    const position = order.sheet_index ?? index + 1;
    const orderCode = order.order_code || order.orderCode || `Đơn ${position}`;
    const customer = order.customer_name || order.customerName || "Chưa có tên khách";
    const sku = order.first_sku || order.firstSku;
    const tab = order.actual_tab_label || order.actualTabLabel || order.expected_tab_label || order.expectedTabLabel || `Đặt hàng ${position}`;
    const detail = `${customer}${sku ? ` · ${sku}` : ""}`;
    const selector = deletionEnabled
      ? `<label class="order-delete-cell" title="Tích chọn ${escapeHtml(orderCode)} để xóa"><input type="checkbox" data-delete-order value="${escapeHtml(orderCode)}" aria-label="Xóa ${escapeHtml(orderCode)}"><span class="order-sequence">${escapeHtml(position)}</span></label>`
      : `<span class="order-sequence">${escapeHtml(position)}</span>`;
    return `<div class="order-progress ${progress.className}" data-order-code="${escapeHtml(orderCode)}" data-order-status="${escapeHtml(statusKey(order.status))}">
      ${selector}
      <div class="order-main"><b>${escapeHtml(orderCode)}</b><small title="${escapeHtml(detail)}">${escapeHtml(detail)}</small></div>
      <div class="order-tab"><b title="${escapeHtml(tab)}">${escapeHtml(tab)}</b><small>${order.note_verified || order.noteVerified ? "Ghi chú khớp" : progress.detail}</small></div>
      <span class="order-state" title="${escapeHtml(progress.detail)}"><i></i>${escapeHtml(progress.label)}</span>
    </div>`;
  }).join("");
  return `<details class="batch-orders" ${batch.is_current ? "open" : ""}>
    <summary><span>Chi tiết ${orders.length} mã đơn · ${verified}/${orders.length} ghi chú đã xác minh</span></summary>
    <div class="order-progress-list">${rows}</div>
  </details>`;
}

function renderBatches(batches, run = null) {
  if (!batches.length) {
    ui.batchList.className = "batch-list empty-state";
    ui.batchList.innerHTML = '<div class="empty-illustration"><span></span><span></span><span></span></div><b>Chưa có lô import</b><p>Hãy tải hoặc phân tích báo cáo để bắt đầu.</p>';
    ui.batchProgressSummary?.classList.add("hidden");
    return;
  }
  ui.batchList.className = "batch-list";
  const runStatus = statusKey(run?.status);
  const abortEnabled = ["AUTOMATION_ATTENTION", "FAILED"].includes(runStatus)
    || batches.some((batch) => ["AUTOMATION_ATTENTION", "FAILED"].includes(statusKey(batch.status)));
  // Aborting is the normal recovery path: retain orders, workbooks, and
  // events for audit while allowing a fresh report scan. Destructive order
  // deletion remains an API-only maintenance action and is never the default
  // response to an interrupted KiotViet workflow.
  const deletionEnabled = false;
  const abortToolbar = abortEnabled
    ? `<div class="delete-orders-toolbar" data-delete-toolbar>
        <b>Đóng workflow lỗi</b>
        <span>Giữ nguyên order, workbook và nhật ký audit; chỉ dùng sau khi đã dọn tab KiotViet và xác minh không có đơn đặt ngoài ledger.</span>
        <button class="button button-danger" type="button" data-abort-run="${escapeHtml(run?.id ?? "")}">Hủy workflow, giữ dữ liệu audit</button>
      </div>`
    : "";
  ui.batchList.innerHTML = abortToolbar + batches.map((batch, index) => {
    const effectiveStatus = batch.display_status || batch.status;
    const status = batchStatusInfo(effectiveStatus);
    const progress = batchProgressInfo(batch);
    const batchId = batch.id ?? batch.batch_id ?? index + 1;
    const orderCount = batch.order_count ?? batch.sheet_count ?? batch.orders?.length ?? 0;
    const lineCount = batch.line_count ?? batch.orders?.reduce((sum, order) => sum + (order.items?.length || order.line_count || 0), 0) ?? 0;
    const canConfirm = batch.can_confirm ?? String(batch.status || "").toLowerCase() === "waiting_user_confirmation";
    const canImport = batch.can_import ?? (index === 0 && String(batch.status || "ready").toLowerCase() === "ready");
    const canRetry = batch.can_retry ?? ["failed", "automation_attention"].includes(String(batch.status || "").toLowerCase());
    const canResumePartial = batch.can_resume_partial === true;
    const batchStatus = statusKey(batch.status);
    const canAutoRun = !needsAutomationAttention(batchStatus)
      && batch.is_unlocked !== false
      && (canImport || canConfirm || ["READY", "WAITING_USER_CONFIRMATION", "WAITING_USER_CONFIRM", "NOTES_FILLED"].includes(batchStatus));
    const action = canResumePartial
      ? `<button class="button button-primary" data-auto-run-batch="${batchId}">Tiếp tục tab còn lại</button>`
      : canAutoRun
      ? `<button class="button button-primary" data-auto-run-batch="${batchId}">${isWaitingCompletionStatus(batchStatus) ? "Tiếp tục đặt hàng tự động" : "Tự động xử lý lô này"}</button>`
      : canRetry
        ? `<button class="button button-ghost" data-retry-batch="${batchId}">Dọn tab tự động & chạy lại</button>`
        : "";
    const batchNumber = batch.batch_number ?? batch.batch_no ?? index + 1;
    const cardClasses = [
      "batch-card",
      batch.is_current ? "is-current" : "",
      progress.tone === "done" ? "is-complete" : "",
      progress.tone === "error" ? "is-error" : "",
    ].filter(Boolean).join(" ");
    return `<article class="${cardClasses}" data-batch-id="${batchId}" data-batch-status="${escapeHtml(statusKey(batch.status))}" data-batch-progress="${progress.percent}">
      <div class="batch-card__top">
        <div class="batch-no">${String(batchNumber).padStart(2, "0")}</div>
        <div><b>Lô ${escapeHtml(batchNumber)} · ${escapeHtml(orderCount)} đơn · ${escapeHtml(lineCount)} dòng</b><small title="${escapeHtml(batch.workbook_path || batch.file_path || "")}">${escapeHtml(batch.workbook_path || batch.file_path || "Đang chuẩn bị đường dẫn file")}</small></div>
        <div class="batch-actions"><span class="batch-status ${status.className}">${status.label}</span>${action}</div>
      </div>
      <div class="batch-progress">
        <div class="batch-progress__top"><b>${escapeHtml(progress.label)}</b><strong>${progress.percent}%</strong></div>
        <div class="progress-track" aria-label="Tiến độ lô ${escapeHtml(batchNumber)}"><i style="width: ${progress.percent}%"></i></div>
        <div class="batch-stages">${renderBatchStages(progress)}</div>
      </div>
      ${renderBatchOrders(batch, deletionEnabled)}
    </article>`;
  }).join("");
  const totalOrders = batches.reduce((sum, batch) => sum + batchOrderStats(batch).total, 0);
  const completedOrders = batches.reduce((sum, batch) => sum + batchOrderStats(batch).completed, 0);
  const totalWeight = batches.reduce((sum, batch) => sum + Math.max(1, batchOrderStats(batch).total), 0);
  const averageProgress = Math.round(batches.reduce((sum, batch) => sum + batchProgressInfo(batch).percent * Math.max(1, batchOrderStats(batch).total), 0) / totalWeight);
  if (ui.batchProgressSummary) {
    ui.batchProgressSummary.classList.remove("hidden");
    ui.batchProgressSummary.innerHTML = `<strong>${completedOrders}/${totalOrders} đơn hoàn tất</strong><div class="progress-track"><i style="width: ${averageProgress}%"></i></div><span>${averageProgress}% toàn workflow</span>`;
  }
  ui.batchList.querySelectorAll("[data-auto-run-batch]").forEach((button) => button.addEventListener("click", () => autoRunBatch(button.dataset.autoRunBatch)));
  ui.batchList.querySelectorAll("[data-retry-batch]").forEach((button) => button.addEventListener("click", () => retryBatch(button.dataset.retryBatch)));
  ui.batchList.querySelectorAll("[data-abort-run]").forEach((button) => button.addEventListener("click", () => abortRun(button.dataset.abortRun)));
  const selectAll = ui.batchList.querySelector("[data-delete-select-all]");
  if (selectAll) {
    selectAll.addEventListener("change", () => {
      ui.batchList.querySelectorAll("[data-delete-order]").forEach((item) => { item.checked = selectAll.checked; });
      updateFailedRunSelectionState();
    });
  }
  ui.batchList.querySelectorAll("[data-delete-order]").forEach((item) => {
    item.addEventListener("change", updateFailedRunSelectionState);
  });
  const deleteButton = ui.batchList.querySelector("[data-delete-run]");
  if (deleteButton) {
    deleteButton.addEventListener("click", () => deleteSelectedOrders(deleteButton.dataset.deleteRun));
    updateFailedRunSelectionState();
  }
}

async function processSource(sourcePath, allowDuplicates = false, runId = null, options = {}) {
  const { manageBusy = true, token = appState.workflowSession?.token || null, scanId = null, businessDate = selectedWorkflowDate() } = options;
  let createdRunId = null;
  let shouldAutoImport = false;
  if (token && !isCurrentSession(token)) return;
  if (manageBusy) setBusy(true, ui.primaryRunButton);
  if (token) updateWorkflowSession(token, 1, "Đang đối chiếu Mã đặt đơn hàng liên quan với SQLite.", "parsing");
  appendLog({
    level: "info",
    message: runId ? "Đang tiếp tục workflow tự động đang chờ." : `Đang phân tích ${basename(sourcePath)}`,
  });
  try {
    const payload = scanId
      ? await api(`/api/scans/${scanId}/duplicate-decision`, { method: "POST", body: { continue: true } })
      : runId
        ? await api(`/api/runs/${runId}/duplicate-decision`, { method: "POST", body: { continue: true } })
        : await api("/api/runs", { method: "POST", body: { source_path: sourcePath, allow_duplicates: allowDuplicates, business_date: businessDate } });
    if (token && !isCurrentSession(token)) return;
    const scan = payload.scan || payload.last_scan;
    if (appState.workflowSession && (!token || isCurrentSession(token))) appState.workflowSession.scan = scan || null;
    if (scan) {
      if (payload.requiresDuplicateConfirmation || payload.requires_duplicate_confirmation || scan.status === "duplicate_review") {
        const duplicates = (scan.decisions || []).filter((item) => item.decision === "skipped_confirmed");
        appState.pendingDuplicateScanId = scan.id;
        if (token) updateWorkflowSession(token, 1, "Phát hiện mã đã xử lý trong ngày. Đang chờ bạn quyết định.", "blocked");
        showDuplicateModal(duplicates);
        toast("Phát hiện dữ liệu trùng. Workflow đã tạm dừng.", "error");
        appendLog({ level: "warning", message: `Phát hiện ${duplicates.length} mã đã xác nhận trong ngày; chờ quyết định của bạn.` });
        await refreshState();
        return;
      }
      const blocked = Boolean(scan.blocked) || ["unresolved", "changed_confirmed", "active_workflow"].includes(String(scan.status || ""));
      if (blocked) {
        const codes = (scan.decisions || []).filter((item) => item.decision === "unresolved" || item.decision === "changed_confirmed").map((item) => item.order_code).join(", ");
        const message = `Dừng an toàn: ${scan.unresolved_count || 0} mã chưa xác nhận${codes ? ` (${codes})` : ""}.`;
        if (token) updateWorkflowSession(token, 1, `${message} Không tạo file/import mới.`, "blocked");
        toast(message, "error");
        appendLog({ level: "warning", message });
      } else if (scan.status === "no_new") {
        const message = `Không có đơn mới; đã bỏ qua ${scan.skipped_confirmed_count || 0} mã đã xác nhận.`;
        if (token) updateWorkflowSession(token, 1, message, "complete");
        toast(message, "success");
        appendLog({ level: "success", message });
      } else {
        const createdRun = normalizeRun(payload);
        if (token && createdRun?.id) {
          appState.workflowSession.run = createdRun;
          appState.workflowSession.expectedRunId = createdRun.id;
          updateWorkflowSession(token, 2, "Đã tạo file import chỉ chứa các mã mới.", "prepared");
        }
        createdRunId = createdRun?.id || null;
        shouldAutoImport = Boolean(createdRunId && scan.status === "run_created");
        const message = `Đã tạo file import chỉ cho ${scan.new_order_count || 0} mã mới; bỏ qua ${scan.skipped_confirmed_count || 0} mã đã xác nhận.`;
        toast(message, "success");
        appendLog({ level: "success", message });
      }
    } else if (payload.requiresDuplicateConfirmation || payload.requires_duplicate_confirmation) {
      const run = payload.run || payload.active_run || payload;
      appState.pendingDuplicateRunId = run.id || payload.run_id;
      showDuplicateModal(payload.duplicates || run.duplicates || []);
      toast("Phát hiện dữ liệu trùng. Workflow đã tạm dừng.", "error");
    } else {
        const createdRun = normalizeRun(payload);
        if (token && createdRun?.id) {
          appState.workflowSession.run = createdRun;
          appState.workflowSession.expectedRunId = createdRun.id;
          updateWorkflowSession(token, 2, "Đã tạo các lô import mới.", "prepared");
        }
        createdRunId = createdRun?.id || null;
        shouldAutoImport = Boolean(createdRunId && !runId);
        toast("Đã phân tích và tạo file import theo lô.", "success");
      appendLog({ level: "success", message: "Đã tạo các lô import, chưa gửi sang KiotViet" });
    }
    await refreshState();
    if (shouldAutoImport && (!token || isCurrentSession(token))) {
      await startUnlockedBatch({ token, expectedRunId: createdRunId, reason: "report mới" });
    }
  } catch (error) {
    if (token) finishWorkflowError(token, error, "Đối chiếu SQLite");
    else appendLog({ level: "error", message: error.message });
    toast(error.message, "error");
  } finally {
    if (manageBusy && !appState.autoRunPending) setBusy(false, ui.primaryRunButton);
  }
}

function showDuplicateModal(duplicates) {
  ui.duplicateList.innerHTML = duplicates.map((duplicate) => `<div class="duplicate-item"><b>${escapeHtml(duplicate.order_code || duplicate.orderCode)}</b><span>${escapeHtml(duplicate.customer_name || duplicate.customerName || duplicate.reason || "Đã xử lý trong ngày")}</span></div>`).join("") || '<div class="duplicate-item"><b>File đã xử lý</b><span>Trùng dấu vân tay dữ liệu</span></div>';
  ui.duplicateModal.classList.remove("hidden");
}

async function exportAndProcess() {
  if (appState.busy) return;
  const session = beginFreshWorkflow({ mode: "export" });
  const token = session.token;
  const businessDate = selectedWorkflowDate();
  // The selected date is authoritative, including when it happens to be
  // today.  Always use the explicit-date command so a stale operational-day
  // clock can never redirect the workflow to an unintended report range.
  const exportAction = "EXPORT_DATE";
  appState.activeChromeAction = exportAction;
  setBusy(true, ui.primaryRunButton);
  try {
    appendLog({ level: "info", message: "Yêu cầu Edge tải report hôm nay từ Sopoka." });
    const result = await chromeRequest(exportAction, { business_date: businessDate }, EXPORT_TODAY_TIMEOUT_MS);
    if (!isCurrentSession(token)) return;
    if (result?.no_new === true) {
      appState.activeChromeAction = null;
      updateWorkflowSession(token, 5, `Sopoka không có dữ liệu ngày ${formatDate(businessDate)}; không có đơn mới để upload.`, "completed");
      setRunStatus(`Ngày ${formatDate(businessDate)} không có đơn mới.`, "success");
      appendLog({ level: "success", message: `Đã kiểm tra Sopoka ngày ${formatDate(businessDate)}: không có dữ liệu.` });
      toast("Ngày đã chọn không có đơn mới.", "success");
      renderState();
      return;
    }
    const sourcePath = String(result?.sourcePath || result?.source_path || "").trim();
    if (!sourcePath || !SUPPORTED_REPORT_FILE_RE.test(sourcePath)) {
      throw new Error(`Edge không trả về đường dẫn file report hợp lệ (${SUPPORTED_REPORT_FILE_LABEL}). Không phân tích hay import dữ liệu nào.`);
    }
    appState.activeChromeAction = null;
    setBusy(true, ui.primaryRunButton);
    appState.workflowSession.sourcePath = sourcePath;
    updateWorkflowSession(token, 1, "Đã có file report. Đang kiểm tra các mã đã xác nhận trong SQLite.", "parsing");
    appendLog({ level: "success", message: `Đã nhận file ${basename(sourcePath)} từ Edge.` });
    toast("Đã tải báo cáo. Bắt đầu kiểm tra dữ liệu.", "success");
    await processSource(sourcePath, false, null, { manageBusy: false, token, businessDate });
  } catch (error) {
    if (isCurrentSession(token)) {
      if (isEdgeRequestTimeout(error)) {
        watchExportUntilBrowserStops(token, "timeout");
        const message = "UI đã quá thời gian chờ phúc đáp, nhưng Edge có thể vẫn đang tải report. Tool đang giám sát Edge; không chạy chồng thêm lượt mới.";
        setRunStatus(message, "warning");
        appendLog({ level: "warning", message });
        toast(message, "info");
        await refreshBrowserStatus();
      } else {
        finishWorkflowError(token, error, "Tải báo cáo — chưa có file nào được phân tích hay ghi SQLite");
        toast(error.message, "error");
      }
    }
  } finally {
    if (appState.workflowSession?.token === token) {
      const stillWatchingExport = appState.timedOutExport?.token === token;
      if (!stillWatchingExport) {
        appState.activeChromeAction = null;
        if (!appState.autoRunPending) setBusy(false, ui.primaryRunButton);
      }
    }
  }
}

function startManualProcess() {
  const path = ui.sourcePath?.value.trim() || "";
  if (!path) {
    ui.sourcePath?.scrollIntoView({ behavior: "smooth", block: "center" });
    ui.sourcePath?.focus();
    toast(`Dán đường dẫn file report vừa tải (${SUPPORTED_REPORT_FILE_LABEL}) trước khi chạy.`, "error");
    return;
  }
  if (appState.busy) return;
  const session = beginFreshWorkflow({ sourcePath: path, mode: "manual" });
  processSource(path, false, null, { token: session.token, businessDate: selectedWorkflowDate() });
}

async function stopActiveExport() {
  const session = appState.workflowSession;
  if (!session || !["EXPORT_TODAY", "EXPORT_DATE"].includes(appState.activeChromeAction)) return;
  session.cancelled = true;
  session.stage = 0;
  session.state = "stopped";
  watchExportUntilBrowserStops(session.token, "cancelling");
  setRunStatus("Đang gửi lệnh dừng tới Edge. Không có file nào sẽ được phân tích trong lượt này.", "warning");
  appendLog({ level: "warning", message: "Người dùng đã yêu cầu dừng tác vụ tải report." });
  renderState();
  setBusy(true, ui.primaryRunButton);
  try {
    const result = await chromeRequest("CANCEL_ACTIVE", {}, 10000);
    appendLog({ level: "warning", message: result.cancelled ? "Edge đã nhận lệnh dừng tác vụ tải report." : (result.message || "Edge không còn tác vụ cần dừng.") });
  } catch (error) {
    appendLog({ level: "error", message: `Không gửi được lệnh dừng: ${error.message}` });
    toast(error.message, "error");
  } finally {
    if (appState.workflowSession?.token === session.token) {
      await refreshBrowserStatus();
      renderState();
    }
  }
}

async function inspectKiot() {
  ui.inspectKiotButton.classList.add("spin");
  try {
    const result = await chromeRequest("INSPECT_KIOT");
    setCheck(ui.checkBranch, result.branchOk ? "ok" : "error", result.branchOk ? "Đúng Đồng Nai 2 - ĐN2" : `Đang ở ${result.branch || "chi nhánh khác"}`);
    setCheck(ui.checkMode, result.orderMode && !result.invoiceMode ? "ok" : "error", result.orderMode ? `${result.orderTabs?.length || 0} tab Đặt hàng` : "Chưa vào màn hình Đặt hàng");
    appendLog({ level: result.branchOk && result.orderMode ? "success" : "warning", message: "Đã kiểm tra trạng thái KiotViet", details: result });
    return result;
  } catch (error) {
    setCheck(ui.checkBranch, "error", error.message);
    setCheck(ui.checkMode, "error", "Không kiểm tra được");
    toast(error.message, "error");
    throw error;
  } finally {
    ui.inspectKiotButton.classList.remove("spin");
  }
}

async function prepareKiot() {
  setBusy(true);
  try {
    const result = await chromeRequest("PREPARE_KIOT", {}, 60000);
    setCheck(ui.checkBranch, "ok", "Đúng Đồng Nai 2 - ĐN2");
    setCheck(ui.checkMode, "ok", `${result.orderTabs?.length || 0} tab Đặt hàng`);
    toast("KiotViet đã ở đúng chi nhánh và tab Đặt hàng.", "success");
    appendLog({ level: "success", message: "KiotViet pre-check đạt" });
  } catch (error) {
    toast(error.message, "error");
    appendLog({ level: "error", message: error.message });
  } finally { setBusy(false); }
}

async function requestAutoRun(runId, { token = null, expectedRunId = null, reason = "workflow", manageBusy = true } = {}) {
  const resolvedRunId = runId ?? appState.activeRun?.id ?? appState.activeRun?.run_id;
  if (!resolvedRunId) {
    toast("Không tìm thấy workflow để tự động chạy.", "error");
    return false;
  }
  if (token && !isCurrentSession(token)) return false;
  if (expectedRunId && String(resolvedRunId) !== String(expectedRunId)) return false;
  const canResumePartial = appState.activeRun?.batches?.some((batch) => batch.can_resume_partial === true);
  if (needsAutomationAttention(appState.activeRun?.status) && !canResumePartial) {
    toast("Workflow đang cần kiểm tra thủ công; tool sẽ không tự chạy lại lô này.", "error");
    return false;
  }
  if (appState.autoRunPending) return true;

  appState.autoRunPending = true;
  if (manageBusy) setBusy(true);
  if (token) updateWorkflowSession(token, 3, "Máy chủ đang tự động import, điền ghi chú và đặt hàng tuần tự.", "importing");
  setRunStatus("Đã giao workflow cho máy chủ. Edge sẽ xử lý tuần tự từng lô và tự bấm ĐẶT HÀNG sau khi xác minh ghi chú.");
  appendLog({ level: "info", message: `Đã yêu cầu tự động chạy workflow từ ${reason}.` });
  try {
    const result = await api(`/api/runs/${encodeURIComponent(resolvedRunId)}/auto-run`, { method: "POST", body: {} });
    const returnedRun = normalizeRun(result);
    if (returnedRun?.id) {
      appState.activeRun = returnedRun;
      if (token && isCurrentSession(token)) {
        appState.workflowSession.run = returnedRun;
        appState.workflowSession.expectedRunId = returnedRun.id;
      }
    }
    await refreshState();
    const attention = needsAutomationAttention(appState.activeRun?.status) || needsAutomationAttention(result?.run?.status);
    if (attention) {
      appState.autoRunPending = false;
      const message = workflowErrorText(appState.activeRun?.error_message) || result?.message || "Kết quả đặt hàng chưa được xác minh đầy đủ; tool đã dừng an toàn.";
      toast(message, "error");
      appendLog({ level: "error", message: "Workflow chuyển sang trạng thái cần kiểm tra thủ công." });
      return false;
    }
    appState.autoRunPending = isAutoRunActive(appState.activeRun);
    if (appState.autoRunPending) {
      toast("Đang tự động xử lý các lô theo thứ tự. Theo dõi thanh tiến độ bên trên.", "success");
      appendLog({ level: "success", message: "Máy chủ đã nhận workflow tự động; Edge đang xử lý tuần tự." });
    } else {
      toast("Workflow tự động đã hoàn tất.", "success");
      appendLog({ level: "success", message: "Workflow tự động đã hoàn tất và đồng bộ SQLite." });
    }
    return true;
  } catch (error) {
    appState.autoRunPending = false;
    await refreshState();
    const attention = needsAutomationAttention(appState.activeRun?.status);
    if (token) finishWorkflowError(token, error, "Tự động đặt hàng");
    const message = attention
      ? (workflowErrorText(appState.activeRun?.error_message) || "Có kết quả đặt hàng chưa được xác minh đầy đủ. Tool đã dừng, chưa ghi mã vào SQLite và không tự thử lại.")
      : error.message;
    toast(message, "error");
    appendLog({ level: "error", message: attention ? "Workflow chuyển sang trạng thái cần kiểm tra thủ công." : `Không thể khởi chạy workflow tự động: ${error.message}` });
    return false;
  } finally {
    if (manageBusy && !appState.autoRunPending) setBusy(false);
  }
}

async function startUnlockedBatch({ token = null, expectedRunId = null, reason = "workflow" } = {}) {
  const runId = expectedRunId ?? appState.activeRun?.id ?? appState.activeRun?.run_id;
  return requestAutoRun(runId, { token, expectedRunId, reason, manageBusy: false });
}

async function autoRunBatch(batchId) {
  const batch = getBatches(appState.activeRun).find((item, index) => String(item.id ?? item.batch_id ?? index + 1) === String(batchId));
  if (!batch) return toast("Không tìm thấy lô import.", "error");
  if (needsAutomationAttention(batch.status)) {
    return retryBatch(batchId);
  }
  const label = batch.batch_number || batch.batch_no || batchId;
  return requestAutoRun(appState.activeRun?.id ?? appState.activeRun?.run_id, { reason: `lô ${label}` });
}

// Compatibility wrappers: legacy buttons or external callers now use the
// server-side sequential runner and never send a browser command from the UI.
async function importBatch(batchId) { return autoRunBatch(batchId); }
async function confirmBatch(batchId) { return autoRunBatch(batchId); }

async function retryBatch(batchId) {
  const batch = getBatches(appState.activeRun).find((item, index) => String(item.id ?? item.batch_id ?? index + 1) === String(batchId));
  const count = Number(batch?.order_count ?? batch?.sheet_count ?? batch?.orders?.length ?? 0);
  const confirmed = window.confirm(`Tool sẽ dọn tab KiotViet dở trước khi import lại lô này. Xác nhận thử lại ${count || "các"} mã chưa được ghi SQLite?`);
  if (!confirmed) return;
  try {
    setBusy(true);
    setRunStatus("Đang dọn tab KiotViet dở trước khi thử lại lô lỗi.");
    const cleaned = await chromeRequest("CLEAN_KIOT_TABS", { confirmCloseData: true }, 90000);
    appendLog({ level: "warning", message: `Đã dọn ${cleaned.closed || 0} tab cũ; đang mở lại lô ${batchId}.` });
    await api(`/api/batches/${batchId}/retry`, { method: "POST", body: { user_confirmed_cleanup: true } });
    toast("Lô đã được mở lại; đang khởi chạy tự động.", "success");
    appendLog({ level: "warning", message: `Đã mở lại lô ${batchId} sau xác nhận dọn tab, chuẩn bị chạy tự động.` });
    await refreshState();
    await autoRunBatch(batchId);
  } catch (error) { toast(error.message, "error"); }
  finally {
    // requestAutoRun keeps the dashboard locked while the server owns Edge.
    if (!appState.autoRunPending) setBusy(false);
  }
}

async function abortRun(runId) {
  const confirmed = window.confirm(
    "Bạn xác nhận đã dọn các tab KiotViet chưa hoàn thành và không có đơn đặt ngoài ledger? Workflow sẽ bị hủy vĩnh viễn, nhưng order, workbook và nhật ký vẫn được giữ để audit.",
  );
  if (!confirmed) return;
  try {
    setBusy(true);
    setRunStatus("Đang đóng workflow lỗi và giữ nguyên dữ liệu audit.", "warning");
    await api(`/api/runs/${runId}/abort`, {
      method: "POST",
      body: {
        confirm: true,
        kiot_tabs_cleared: true,
        acknowledge_external_order_risk: true,
      },
    });
    toast("Workflow đã được hủy; dữ liệu audit được giữ nguyên.", "success");
    appendLog({ level: "warning", message: `Đã hủy workflow ${runId}; không xóa order, workbook hoặc ledger.` });
    await refreshState();
  } catch (error) {
    toast(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function deleteSelectedOrders(runId) {
  const selected = [...document.querySelectorAll("[data-delete-order]:checked")]
    .map((item) => String(item.value || "").trim())
    .filter(Boolean);
  if (!selected.length) {
    toast("Hãy tích chọn mã cần xóa trước.", "warning");
    return;
  }
  const total = document.querySelectorAll("[data-delete-order]").length;
  if (selected.length !== total) {
    toast(`Hãy tích đủ ${total} mã của workflow trước khi hủy lô.`, "warning");
    return;
  }
  const confirmed = window.confirm(
    `Bạn sắp hủy workflow lỗi và giải phóng ${selected.length} mã chưa xác nhận. Không có đơn nào trong danh sách này đã được ghi vào ledger SQLite. Tool chỉ thực hiện khi toàn bộ 39 mã của workflow được tích chọn. Tiếp tục?`,
  );
  if (!confirmed) return;
  try {
    setBusy(true);
    await api(`/api/runs/${runId}/delete-selected`, {
      method: "POST",
      body: { order_codes: selected, confirm: true },
    });
    toast(`Đã hủy lô lỗi và giải phóng ${selected.length} mã. Bạn có thể quét report lại.`, "success");
    appendLog({ level: "warning", message: `Đã hủy lô lỗi và giải phóng ${selected.length} mã chưa xác nhận.` });
    await refreshState();
  } catch (error) {
    toast(error.message, "error");
  } finally {
    setBusy(false);
  }
}

function updateFailedRunSelectionState() {
  const boxes = [...document.querySelectorAll("[data-delete-order]")];
  const selected = boxes.filter((item) => item.checked).length;
  const selectAll = document.querySelector("[data-delete-select-all]");
  if (selectAll) {
    selectAll.checked = Boolean(boxes.length && selected === boxes.length);
    selectAll.indeterminate = Boolean(selected && selected < boxes.length);
  }
  const button = document.querySelector("[data-delete-run]");
  if (button) {
    button.disabled = !boxes.length || selected !== boxes.length;
    button.textContent = `Hủy lô lỗi & giải phóng mã (${selected}/${boxes.length})`;
  }
}

function switchView(view) {
  const ledger = view === "ledger";
  ui.workflowView?.classList.toggle("hidden", ledger);
  ui.ledgerView?.classList.toggle("hidden", !ledger);
  ui.workflowTabButton?.classList.toggle("active", !ledger);
  ui.ledgerTabButton?.classList.toggle("active", ledger);
  ui.workflowTabButton?.setAttribute("aria-selected", String(!ledger));
  ui.ledgerTabButton?.setAttribute("aria-selected", String(ledger));
  if (ledger) loadLedgerOrders();
}

function selectedLedgerCodes() {
  return [...document.querySelectorAll("[data-ledger-order]:checked")]
    .map((item) => String(item.value || "").trim())
    .filter(Boolean);
}

function updateLedgerSelectionState() {
  const boxes = [...document.querySelectorAll("[data-ledger-order]")];
  const selected = selectedLedgerCodes();
  if (ui.ledgerSelectAll) {
    ui.ledgerSelectAll.checked = Boolean(boxes.length && selected.length === boxes.length);
    ui.ledgerSelectAll.indeterminate = Boolean(selected.length && selected.length < boxes.length);
  }
  if (ui.deleteLedgerButton) ui.deleteLedgerButton.disabled = selected.length === 0;
}

function renderLedgerOrders() {
  const rows = Array.isArray(appState.ledgerOrders) ? appState.ledgerOrders : [];
  if (ui.ledgerCountLabel) ui.ledgerCountLabel.textContent = `${rows.length} mã đã lưu`;
  if (!ui.ledgerTable) return;
  if (!rows.length) {
    ui.ledgerTable.innerHTML = '<div class="ledger-status">Không có mã đã xác nhận trong ngày này.</div>';
    updateLedgerSelectionState();
    repairUiText();
    return;
  }
  ui.ledgerTable.innerHTML = `<table class="ledger-table"><thead><tr>
    <th>Chọn</th><th>Mã đặt đơn liên quan</th><th>Khách hàng</th><th>Ngày đơn</th>
    <th>Lô / run</th><th>Trạng thái</th><th>Đã ghi lúc</th>
  </tr></thead><tbody>${rows.map((row) => {
    const code = row.order_code || "";
    const customer = row.customer_name || row.customer_code || "—";
    const run = row.run_id ? `#${row.run_id} · Lô ${row.batch_number ?? "—"}` : "—";
    const status = row.confirmation_method === "automation_verified_submission" ? "Tự động đã xác nhận" : "Đã xác nhận";
    return `<tr><td><input type="checkbox" data-ledger-order value="${escapeHtml(code)}" aria-label="Chọn ${escapeHtml(code)}"></td>
      <td><b>${escapeHtml(code)}</b></td><td>${escapeHtml(customer)}</td>
      <td>${escapeHtml(formatDate(row.order_date || row.business_date))}</td>
      <td>${escapeHtml(run)}</td><td>${escapeHtml(status)}</td>
      <td>${escapeHtml(row.confirmed_at || "—")}</td></tr>`;
  }).join("")}</tbody></table>`;
  updateLedgerSelectionState();
  repairUiText();
}

async function loadLedgerOrders({ silent = false } = {}) {
  if (!ui.ledgerDate) return;
  if (!ui.ledgerDate.value) ui.ledgerDate.value = operationalBusinessDate();
  if (!silent && ui.ledgerStatus) {
    ui.ledgerStatus.className = "ledger-status";
    ui.ledgerStatus.textContent = "Đang đọc SQLite…";
  }
  try {
    const date = ui.ledgerDate.value || operationalBusinessDate();
    const result = await api(`/api/ledger/orders?business_date=${encodeURIComponent(date)}`);
    appState.ledgerOrders = Array.isArray(result.orders) ? result.orders : [];
    if (ui.ledgerStatus) {
      ui.ledgerStatus.className = "ledger-status success";
      ui.ledgerStatus.textContent = `Đã đọc ${appState.ledgerOrders.length} mã từ SQLite ngày ${formatDate(date)}.`;
    }
    renderLedgerOrders();
  } catch (error) {
    appState.ledgerOrders = [];
    if (ui.ledgerStatus) {
      ui.ledgerStatus.className = "ledger-status error";
      ui.ledgerStatus.textContent = error.message;
    }
    renderLedgerOrders();
  }
}

async function deleteSelectedLedgerOrders() {
  const selected = selectedLedgerCodes();
  if (!selected.length) {
    toast("Hãy tích chọn mã cần xóa trước.", "warning");
    return;
  }
  const confirmed = window.confirm(
    `Bạn sắp xóa ${selected.length} mã khỏi SQLite. Thao tác này không hoàn tác đơn đã đặt trên KiotViet và có thể tạo đơn trùng khi chạy lại. Bạn chắc chắn muốn tiếp tục?`,
  );
  if (!confirmed) return;
  try {
    if (ui.deleteLedgerButton) ui.deleteLedgerButton.disabled = true;
    await api("/api/ledger/orders/delete", {
      method: "POST",
      body: {
        business_date: ui.ledgerDate?.value || operationalBusinessDate(),
        order_codes: selected,
        confirm: true,
        acknowledge_external_order_risk: true,
      },
    });
    toast(`Đã xóa ${selected.length} mã khỏi SQLite.`, "success");
    appendLog({ level: "warning", message: `User đã xóa ${selected.length} mã đã chọn khỏi ledger SQLite.` });
    await loadLedgerOrders();
    await refreshState();
  } catch (error) {
    toast(error.message, "error");
    updateLedgerSelectionState();
  }
}

ui.primaryRunButton.addEventListener("click", exportAndProcess);
ui.downloadButton.addEventListener("click", exportAndProcess);
ui.stopRunButton?.addEventListener("click", stopActiveExport);
ui.sourceForm.addEventListener("submit", (event) => {
  event.preventDefault();
  startManualProcess();
});
ui.purchaseForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const sourcePath = String(ui.purchaseSourcePath?.value || "").trim();
  const baselinePath = String(ui.purchaseBaselinePath?.value || "").trim();
  const purchaseDate = String(ui.purchaseDate?.value || "").trim();
  if (!sourcePath) {
    toast("Vui lòng dán đường dẫn file nguồn nhập hàng.", "error");
    ui.purchaseSourcePath?.focus();
    return;
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(purchaseDate)) {
    toast("Vui lòng chọn ngày phiếu nhập hợp lệ.", "error");
    ui.purchaseDate?.focus();
    return;
  }
  if (appState.busy) {
    toast("Một tác vụ Edge khác đang chạy. Vui lòng chờ hoàn tất.", "info");
    return;
  }
  setBusy(true, ui.purchaseImportButton);
  if (ui.purchaseSummary) {
    ui.purchaseSummary.classList.remove("hidden");
    ui.purchaseSummary.textContent = "Đang đọc file nguồn, tạo workbook và kiểm tra chi nhánh Đồng Nai 2…";
  }
  try {
    const result = await api("/api/purchase/import", {
      method: "POST",
      body: {
        source_path: sourcePath,
        baseline_source_path: baselinePath || undefined,
        purchase_date: purchaseDate,
        auto_create_missing_products: ui.autoCreateMissingProducts?.checked === true,
      },
      timeoutMs: 5 * 60 * 1000,
    });
    const build = result?.build || {};
    const edge = result?.edge || {};
    const comparison = build.comparison || {};
    const createdProducts = Array.isArray(edge.createdProducts) && edge.createdProducts.length
      ? ` Đã tự tạo ${edge.createdProducts.length} mã thiếu: ${edge.createdProducts.join(", ")}.`
      : "";
    const compared = comparison.mode === "positive_quantity_delta"
      ? ` Đối chiếu: ${comparison.incremental_sku_count || 0} mã tăng thêm (${(comparison.new_skus || []).length} mã mới).`
      : "";
    const skipped = edge.skipped === true
      ? " Không có số lượng tăng thêm nên không tạo phiếu trên KiotViet."
      : "";
    const message = `Đã tạo ${build.item_count || 0} dòng hợp lệ, bỏ qua ${build.skipped_non_delivery || 0} dòng không giao. Ngày phiếu: ${edge.purchaseDateVisible || purchaseDate}. File: ${build.output_path || "—"}. ${edge.supplier ? `Đã chọn NCC ${edge.supplier}.` : ""}${compared}${createdProducts}${skipped}`;
    if (ui.purchaseSummary) ui.purchaseSummary.textContent = message;
    appendLog({ level: "success", message });
    toast("Nhập hàng OMEGA đã hoàn tất trên KiotViet.", "success");
  } catch (error) {
    const message = error?.message || "Không thể nhập hàng.";
    if (ui.purchaseSummary) ui.purchaseSummary.textContent = `Đã dừng an toàn: ${message}`;
    appendLog({ level: "error", message: `Nhập hàng: ${message}` });
    toast(message, "error");
  } finally {
    setBusy(false, ui.purchaseImportButton);
  }
});
ui.inspectKiotButton.addEventListener("click", inspectKiot);
ui.prepareKiotButton.addEventListener("click", prepareKiot);
ui.cleanTabsButton.addEventListener("click", () => ui.cleanupModal.classList.remove("hidden"));
ui.cancelCleanupButton.addEventListener("click", () => ui.cleanupModal.classList.add("hidden"));
ui.confirmCleanupButton.addEventListener("click", async () => {
  ui.cleanupModal.classList.add("hidden");
  setBusy(true);
  try {
    const result = await chromeRequest("CLEAN_KIOT_TABS", { confirmCloseData: true }, 60000);
    toast(`Đã đóng ${result.closed || 0} tab cũ.`, "success");
    appendLog({ level: "success", message: `Đã đóng ${result.closed || 0} tab đặt hàng cũ` });
    await inspectKiot();
  } catch (error) { toast(error.message, "error"); } finally { setBusy(false); }
});
ui.stopDuplicateButton.addEventListener("click", async () => {
  const runId = appState.pendingDuplicateRunId;
  const scanId = appState.pendingDuplicateScanId;
  ui.duplicateModal.classList.add("hidden");
  if (scanId) await api(`/api/scans/${scanId}/duplicate-decision`, { method: "POST", body: { continue: false } }).catch(() => {});
  else if (runId) await api(`/api/runs/${runId}/duplicate-decision`, { method: "POST", body: { continue: false } }).catch(() => {});
  appState.pendingDuplicateRunId = null;
  appState.pendingDuplicateScanId = null;
  toast("Workflow đã dừng, không import dữ liệu trùng.");
  refreshState();
});
ui.continueDuplicateButton.addEventListener("click", async () => {
  const runId = appState.pendingDuplicateRunId;
  const scanId = appState.pendingDuplicateScanId;
  ui.duplicateModal.classList.add("hidden");
  appState.pendingDuplicateRunId = null;
  appState.pendingDuplicateScanId = null;
  if (scanId) return processSource(null, false, null, { scanId });
  if (!runId) return toast("Không tìm thấy workflow đang chờ.", "error");
  await processSource(null, true, runId);
});
ui.clearLogButton.addEventListener("click", () => { appState.localEvents = []; renderEvents(); });
ui.workflowTabButton?.addEventListener("click", () => switchView("workflow"));
ui.ledgerTabButton?.addEventListener("click", () => switchView("ledger"));
ui.refreshLedgerButton?.addEventListener("click", () => loadLedgerOrders());
ui.ledgerDate?.addEventListener("change", () => loadLedgerOrders());
ui.ledgerSelectAll?.addEventListener("change", () => {
  document.querySelectorAll("[data-ledger-order]").forEach((item) => { item.checked = ui.ledgerSelectAll.checked; });
  updateLedgerSelectionState();
});
ui.ledgerTable?.addEventListener("change", (event) => {
  if (event.target?.matches?.("[data-ledger-order]")) updateLedgerSelectionState();
});
ui.deleteLedgerButton?.addEventListener("click", deleteSelectedLedgerOrders);

renderExtensionStatus();
refreshBusinessDateChip();
refreshBrowserStatus();
refreshState();
if (ui.ledgerDate) ui.ledgerDate.value = operationalBusinessDate();
if (ui.workflowDate) ui.workflowDate.value = operationalBusinessDate();
setInterval(refreshState, 15000);
// Keep the operational date correct if this dashboard remains open across
// midnight; the server snapshot is refreshed separately every 15 seconds.
setInterval(() => refreshBusinessDateChip(appState.workflowSession?.run || appState.activeRun, appState.workflowSession), 30000);
setInterval(() => {
  refreshBrowserStatus();
}, 2000);
