(() => {
  if (globalThis.__orderBridgeSopokaLoaded) return;
  globalThis.__orderBridgeSopokaLoaded = true;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const norm = (value) => String(value || "")
    .normalize("NFC")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

  const visible = (element) => {
    if (!element) return false;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return rect.width > 0
      && rect.height > 0
      && style.visibility !== "hidden"
      && style.display !== "none";
  };

  const actionable = (element) => visible(element)
    && !element.disabled
    && element.getAttribute("aria-disabled") !== "true"
    && getComputedStyle(element).pointerEvents !== "none";

  const waitFor = async (finder, label, timeoutMs = 12000) => {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const element = finder();
      if (element) return element;
      await sleep(200);
    }
    throw new Error(`Không tìm thấy ${label} trên Sopoka.`);
  };

  const click = (element) => {
    // Do not call scrollIntoView here.  Sopoka's large report grid can recurse
    // through layout handlers when it is programmatically scrolled.
    element.click();
  };

  const expectedToday = () => {
    const now = new Date();
    return [
      String(now.getDate()).padStart(2, "0"),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getFullYear()),
    ].join("/");
  };

  const relevantDateElements = (anchor) => {
    const anchorRect = anchor.getBoundingClientRect();
    return [...document.querySelectorAll('input,button,td,span,div')].filter((element) => {
      if (!visible(element)) return false;
      const rect = element.getBoundingClientRect();
      const inTopRightControl = rect.top < 210
        && rect.right > innerWidth * 0.52
        && rect.width < 720
        && rect.height < 240;
      const inPickerInput = element.matches("input")
        && rect.right > anchorRect.left - 460
        && rect.left < anchorRect.right + 100
        && rect.top < Math.min(innerHeight, 700);
      return inTopRightControl || inPickerInput;
    });
  };

  const readVisibleDateValues = (anchor) => {
    const candidates = relevantDateElements(anchor);
    const values = [];
    for (const element of candidates) {
      const text = String(element.value || element.textContent || "").replace(/\s+/g, " ").trim();
      const dates = text.match(/\d{1,2}\/\d{1,2}\/\d{4}/g) || [];
      values.push(...dates);
    }
    return values;
  };

  const visibleRangeMatches = (expectedDate, anchor) => {
    const candidates = relevantDateElements(anchor);
    let matchingInputs = 0;
    for (const element of candidates) {
      const text = String(element.value || element.textContent || "").replace(/\s+/g, " ").trim();
      const dates = text.match(/\d{1,2}\/\d{1,2}\/\d{4}/g) || [];
      if (dates.length >= 2 && dates[0] === expectedDate && dates[1] === expectedDate) return true;
      if (element.matches("input") && dates.length === 1 && dates[0] === expectedDate) matchingInputs += 1;
    }
    return matchingInputs >= 2;
  };

  const exportToday = async () => {
    let stage = "open-date";
    try {
      const expectedDate = expectedToday();
      const dateButton = await waitFor(
        () => document.querySelector("td._GATo"),
        "bộ chọn ngày td._GATo",
      );
      click(dateButton);

      stage = "select-today";
      const today = await waitFor(
        () => [...document.querySelectorAll("li")]
          .find((element) => actionable(element) && norm(element.textContent) === "hôm nay"),
        "mục Hôm nay",
      );
      click(today);

      stage = "apply-date";
      const apply = await waitFor(
        () => [...document.querySelectorAll("button,input[type=button]")]
          .find((element) => actionable(element) && norm(element.value || element.textContent) === "áp dụng"),
        "nút Áp dụng",
        3500,
      ).catch(() => null);
      if (apply) click(apply);

      stage = "verify-date";
      await waitFor(
        () => (visibleRangeMatches(expectedDate, dateButton) ? dateButton : null),
        `bộ lọc ${expectedDate} - ${expectedDate}`,
        8000,
      );

      stage = "search";
      const search = await waitFor(
        () => {
          const element = document.querySelector("button.timK");
          return actionable(element) ? element : null;
        },
        "nút Tìm kiếm button.timK",
      );
      click(search);

      stage = "wait-results";
      await sleep(1600);
      let stablePasses = 0;
      let priorSignature = null;
      await waitFor(
        () => {
          const busy = [...document.querySelectorAll('.loading,.loader,.spinner,.fa-spinner,[aria-busy="true"]')]
            .some((element) => visible(element));
          const element = document.querySelector("button.xuatEx");
          const rows = [...document.querySelectorAll("table tbody tr")];
          const signature = `${rows.length}:${rows.slice(0, 3).map((row) => row.textContent?.length || 0).join(",")}`;
          stablePasses = !busy && signature === priorSignature ? stablePasses + 1 : 0;
          priorSignature = signature;
          return !busy && stablePasses >= 2 && actionable(element) ? element : null;
        },
        "dữ liệu báo cáo tải xong",
        30000,
      );

      stage = "export";
      const exportButton = await waitFor(
        () => {
          const element = document.querySelector("button.xuatEx");
          return actionable(element) ? element : null;
        },
        "nút Xuất Excel button.xuatEx",
      );
      click(exportButton);
      return {
        ok: true,
        selectedPreset: "Hôm nay",
        expectedDate,
        observedDates: readVisibleDateValues(dateButton).slice(0, 12),
      };
    } catch (error) {
      throw new Error(`Sopoka/${stage}: ${String(error?.message || error)}`);
    }
  };

  let inFlight = null;
  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type !== "ORDERBRIDGE_SOPKA_EXPORT_TODAY") return false;
    if (inFlight) {
      sendResponse({ ok: false, error: "Sopoka/export: một lần xuất báo cáo khác đang chạy." });
      return false;
    }
    inFlight = exportToday();
    inFlight
      .then((result) => sendResponse(result))
      .catch((error) => sendResponse({ ok: false, error: String(error?.message || error) }))
      .finally(() => { inFlight = null; });
    return true;
  });
})();
