/*
 * Lightweight browser-free regression checks for the dashboard's critical
 * primary-button guard.  This deliberately runs the real static/app.js in a
 * tiny DOM shim: no browser, Edge, Sopoka, or KiotViet is contacted.
 */
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { webcrypto } = require("crypto");
const { TextDecoder } = require("util");

function element() {
  return {
    className: "",
    textContent: "",
    title: "",
    innerHTML: "",
    value: "",
    hidden: false,
    disabled: false,
    dataset: {},
    style: { setProperty() {} },
    classList: { add() {}, remove() {}, toggle() {} },
    append() {},
    remove() {},
    focus() {},
    scrollIntoView() {},
    addEventListener() {},
    querySelector() { return element(); },
    querySelectorAll() { return []; },
  };
}

const nodes = new Map();
const document = {
  querySelector(selector) {
    if (!nodes.has(selector)) nodes.set(selector, element());
    return nodes.get(selector);
  },
  querySelectorAll(selector) {
    return selector === "#workflowSteps li"
      ? [element(), element(), element(), element(), element()]
      : [];
  },
  createElement: element,
};

let responsePayload = { ok: true, available: true, started: false, today: "2026-08-08", run: null };
const context = vm.createContext({
  console,
  document,
  window: { confirm: () => false },
  crypto: webcrypto,
  AbortController,
  TextDecoder,
  Intl,
  Date,
  Math,
  Object,
  String,
  Number,
  Boolean,
  RegExp,
  Error,
  JSON,
  Promise,
  setInterval: () => 0,
  clearInterval() {},
  setTimeout: () => 0,
  clearTimeout() {},
  fetch: async () => ({
    ok: true,
    status: 200,
    json: async () => responsePayload,
  }),
});

const source = fs.readFileSync(path.join(__dirname, "..", "static", "app.js"), "utf8");
vm.runInContext(source, context, { filename: "static/app.js" });

async function main() {
  // A stale /api/state snapshot must not keep yesterday's date after midnight.
  const expectedToday = vm.runInContext("bangkokBusinessDate()", context);
  const effectiveToday = vm.runInContext("appState.server = { today: '2000-01-01' }; operationalBusinessDate()", context);
  assert.strictEqual(effectiveToday, expectedToday);
  assert.ok(vm.runInContext("EXPORT_TODAY_TIMEOUT_MS", context) >= 300000);

  // Historic workflow errors persisted before the UTF-8 repair must render
  // as Vietnamese rather than raw Windows-1252 mojibake in the dashboard.
  const repaired = vm.runInContext(
    "workflowErrorText('KiotViet chÆ°a náº¡p xong giao diá»‡n bÃ¡n hÃ ng')",
    context,
  );
  assert.strictEqual(repaired, "KiotViet chưa nạp xong giao diện bán hàng");

  // The retry route always sends a clean-tab command before reopening the
  // quarantined batch, so an explicit retry never imports on top of a stale
  // KiotViet cart.
  const retrySource = vm.runInContext("retryBatch.toString()", context);
  assert.ok(retrySource.indexOf("CLEAN_KIOT_TABS") < retrySource.indexOf("/retry"));
  const deleteSource = vm.runInContext("deleteSelectedOrders.toString()", context);
  assert.ok(deleteSource.includes("selected.length !== total"));
  assert.ok(vm.runInContext("updateFailedRunSelectionState.toString()", context).includes("selected !== boxes.length"));

  // An explicit API {ok:false} must stop the primary flow before an undefined
  // sourcePath reaches SQLite parsing.
  responsePayload = { ok: false, message: "browser said no" };
  await assert.rejects(
    () => vm.runInContext("api('/api/browser/commands')", context),
    /browser said no/,
  );

  // When a request response is lost, the UI keeps the primary button locked
  // until Edge status reports that the original export has ended.
  responsePayload = { ok: true, available: true, started: true, progress: { active: true, phase: "export_download" } };
  vm.runInContext("watchExportUntilBrowserStops('smoke-token', 'timeout')", context);
  assert.strictEqual(vm.runInContext("appState.busy", context), true);
  assert.strictEqual(vm.runInContext("appState.activeChromeAction", context), "EXPORT_TODAY");
  assert.strictEqual(nodes.get("#stopRunButton").hidden, false);
  assert.strictEqual(
    vm.runInContext("isEdgeRequestTimeout(new Error('EXPORT_TODAY: Edge WebDriver không phản hồi sau 300 giây.'))", context),
    true,
  );
  vm.runInContext("reconcileWatchedExport({ progress: { active: true, phase: 'export_download' } })", context);
  assert.strictEqual(vm.runInContext("appState.busy", context), true);
  vm.runInContext("reconcileWatchedExport({ progress: { active: false, phase: 'error' } })", context);
  assert.strictEqual(vm.runInContext("appState.busy", context), false);
  assert.strictEqual(vm.runInContext("appState.activeChromeAction", context), null);

  console.log("ui_flow_smoke: passed");
}

main().catch((error) => {
  console.error(error.stack || error);
  process.exitCode = 1;
});
