# Daily Bought vs Sold Review

The office-facing [user guide](bought-vs-sold-user-guide.html) is available from
**How this works** on the review page, with Print / Save as PDF.

Open **Purchase Orders → Bought vs Sold Review**. This feature uses Office PO
Review permissions: viewers can inspect/export; editors can upload, confirm,
reverse confirmations, and update investigations.

The working queue starts **October 27, 2026**. Earlier purchases are excluded.
Undated invoices uploaded from that date remain in the queue for date review;
upload date never replaces the purchase date for aging. Future invoice dates
are flagged for review. Unresolved quantities stay in the queue across uploads.

Each day, upload the complete last seven days of Field Wise ticket lines and
select the export's actual through date. Submitted, Approved, and Invoiced count
as sold. Rolling imports retain history, replace revised tickets, and avoid
counting overlapping exports twice. Missing tickets are inactivated only within
the stated complete window. Older historical backfills require a separate review.

Purchase invoices are read from the app database and attachments when the page
opens or Refresh is selected. Content hashes cache parsing and detect document
replacement. Parsing is a preview and does not insert accounting cost records.
Parsing failures, unreconciled totals, duplicate documents, credits, and missing
dates remain review exceptions.

The allowance is **0–3 calendar days** from purchase date. Over 3 days requires
review only if every day from purchase through the latest loaded sales day is
covered. Timing is assessed through that loaded day, not an unsupported current
day. An unconfirmed quantity is not a conclusion that a sale was missed.

Exact item, job and customer links are suggestions. Confirm the unit conversion
and quantity explicitly. A sales quantity cannot be allocated twice. Reversed
confirmations remain in history. Corrected source lines, disallowed ticket
statuses, or ticket returns invalidate affected confirmations for rechecking.

Flag an issue with a group, description, status, and update notes. Issue status
and sold quantity are independent: resolving an issue does not mark it sold.
Issue edits use revision checks to prevent overwriting another reviewer's work.
Changed or removed source lines preserve the original issue and history.

All review/import/cache/issue tables use the `material_` prefix in the main
dashboard database, so the normal database backup includes review history.
Original uploaded exports are retained under `data/material-review-exports`;
include that folder and `data/uploads` in file backups.

## Isolated offline dashboard

Run `run_offline_dashboard.py` from the repository with the application's Python
environment, then open `http://127.0.0.1:8774/?tab=boughtVsSold` and sign in with
the usual dashboard account from the downloaded database.

This launcher creates `outputs/offline-dashboard/data` once using the downloaded
October 27 snapshot and PO invoice files. Re-running preserves local edits. It
uses loopback only, disables reminder jobs and external service credentials,
and does not contact or change the live server. Purchases added in this local
dashboard become available to its review; new live-server uploads are not fetched.

The earlier standalone preview on port 8773 remains a separate snapshot review.
Use port 8774 for the integrated offline dashboard.
