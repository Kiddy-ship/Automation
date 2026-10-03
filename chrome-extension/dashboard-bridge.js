(() => {
  const ALLOWED = new Set(["http://127.0.0.1:8765", "http://localhost:8765"]);
  const post = (payload) => window.postMessage({ channel: "sopoka-automation-extension", ...payload }, window.origin);

  post({ type: "READY", version: chrome.runtime.getManifest().version });

  // The visible dashboard tab acts as the localhost bridge.  This lets the
  // local controller request only the same safe actions available in the UI;
  // there is deliberately no final-order-submit action.
  const runRemoteBridge = async () => {
    try {
      const poll = await fetch(`${location.origin}/api/bridge/poll`, {
        cache: "no-store",
        credentials: "same-origin",
      });
      const envelope = poll.ok ? await poll.json() : null;
      const command = envelope?.command;
      if (command?.id && command?.action) {
        let response;
        try {
          response = await chrome.runtime.sendMessage({
            action: command.action,
            payload: command.payload || {},
          });
        } catch (error) {
          response = { ok: false, error: String(error?.message || error) };
        }
        await fetch(`${location.origin}/api/bridge/results/${encodeURIComponent(command.id)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          credentials: "same-origin",
          body: JSON.stringify({ response }),
        });
        post({ type: "EVENT", event: { level: response?.ok === false ? "error" : "info", message: `Edge: ${command.action}`, details: response || {} } });
      }
    } catch {
      // A short retry keeps the bridge resilient while Flask/Chrome reloads.
    }
    window.setTimeout(runRemoteBridge, 350);
  };
  runRemoteBridge();

  window.addEventListener("message", async (event) => {
    if (event.source !== window || !ALLOWED.has(event.origin)) return;
    const message = event.data;
    if (!message || message.channel !== "sopoka-automation-ui") return;
    if (message.type === "HELLO") {
      post({ type: "READY", version: chrome.runtime.getManifest().version });
      return;
    }
    if (message.type !== "REQUEST") return;
    try {
      const response = await chrome.runtime.sendMessage({
        action: message.action,
        payload: message.payload || {},
      });
      post({ type: "RESPONSE", requestId: message.requestId, response });
    } catch (error) {
      post({
        type: "RESPONSE",
        requestId: message.requestId,
        response: { ok: false, error: String(error?.message || error) },
      });
    }
  });

  chrome.runtime.onMessage.addListener((message) => {
    if (message?.type === "PUSH_EVENT") post({ type: "EVENT", event: message.event });
  });
})();
