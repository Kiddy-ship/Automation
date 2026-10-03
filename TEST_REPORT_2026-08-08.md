# OrderBridge smoke/E2E report — 2026-08-08

## Result

**PASS** for the dashboard download and SQLite duplicate-review path.

The test deliberately stopped at the duplicate-review dialog. It did not create
a new workflow run, import a workbook, open KiotViet, fill a note, or submit an
order.

## Real dashboard E2E

Executed through the real `Tải report & chạy tự động` button on
`http://127.0.0.1:8765`:

1. Dashboard rendered the Bangkok business date as `08/08/2026`.
2. Edge opened Sopoka's sales-summary report, selected `Hôm nay`, searched,
   exported, and downloaded `ReportOrderSaleSummary_1786184012516.xls`.
3. The downloaded report was 51,200 bytes and parsed successfully for
   `08/08/2026`, with 32 related-order codes.
4. SQLite found 2 already-confirmed codes and 30 new codes. The dashboard
   displayed the duplicate-review dialog for `ĐH_29524392026` and
   `ĐH_29488912026`.
5. The test chose **Dừng workflow** in the real UI. The scan is now
   `cancelled`; SQLite contains no new run or confirmed-order ledger entry from
   this smoke test.

Evidence screenshot: `logs/smoke-ui-duplicate-review.png`.

## Regression fixes covered

- Sopoka's patched `Date` implementation is no longer touched by a page-level
  export script. The export uses normal browser input only.
- The date picker is checked before clicking. An already-open picker is reused
  instead of being toggled closed.
- The export button is verified by observing a new validated Office Viewer
  target or downloaded file. It gets one physical-pointer retry only when
  neither exists.
- Office Viewer downloads the validated Sopoka report URL using only applicable
  in-memory browser session cookies. Cookies and the generated URL are never
  written to the UI or logs. The native Edge `Download file` action remains a
  bounded fallback.
- UI timeout is five minutes and does not re-enable the Run button while Edge
  might still be working. It also rejects a semantic `{ok:false}` response and
  a missing `.xls` path before SQLite/KiotViet.
- The business-date chip uses `Asia/Bangkok` and changes at midnight. For
  multiple import files, the global progress bar stays monotonic while each
  batch/file gets an individual segment and a `Lô X/Y` label.

## Automated checks

Run from the workspace root:

```powershell
.\.venv\Scripts\python.exe -m py_compile app.py automation_app\edge_automation.py automation_app\service.py
& 'C:\Program Files\nodejs\node.exe' --check static\app.js
& 'C:\Program Files\nodejs\node.exe' tests\ui_flow_smoke.js
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Final result for the original smoke: **40/40 Python tests passed**, JavaScript
syntax check passed, and `ui_flow_smoke.js` passed.

## Follow-up verification — 2026-08-10

The SQLite tracking tab and print-preview cancellation fix were verified after
the follow-up changes:

- `py_compile`, both JavaScript syntax checks, and `ui_flow_smoke.js` passed.
- Full Python regression suite: **56/56 tests passed**.
- `GET /api/ledger/orders?business_date=2026-08-10` returned 6 confirmed
  records; the live dashboard rendered the same 6 checkbox rows.
- The print watcher now retries until Edge's Cancel control is hydrated and
  also handles the short post-submit race. Live logs recorded
  `devtools_cancel`/`devtools_close` and continued to the next order.
- Deleting ledger rows is explicit and guarded by both a UI confirmation and a
  risk acknowledgement; it removes only SQLite source/ledger rows and does
  not undo an order already created in KiotViet.

## Historical-date recovery run — 2026-08-12

- Processed `ReportOrderSaleSummary_1786507937622.xls` with selected business
  date `2026-08-11` (the report contained 37 related-order codes).
- SQLite had no confirmed codes for that date, so the run was split into two
  batches of 19 and 18 orders.
- Batch 1 and batch 2 were both submitted and confirmed; the 11/08 ledger now
  contains **37/37** codes and the run is `completed`.
- Batch 2 first timed out after upload. The visible KiotViet page was checked,
  all imported tabs were present, and all 18 notes were verified. The workbook
  was not uploaded a second time; the run was reconciled and continued safely.
- A DevTools print watcher automatically cancelled each visible Edge print
  preview during the final submissions.

The selected-date backend path now accepts `business_date` for both report
export and manual `.xls` processing, while normal runs still default to the
current Bangkok business date.
