# PM Project Cost Tracker

A local project-cost tracking tool for industrial electrical jobs with master projects, subprojects, change orders, Field Wise imports, vendor invoices, and profitability dashboards.

## Start

Install Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

Run:

```powershell
python app.py
```

Then open:

```text
http://127.0.0.1:8765
```

The daily **Purchase Orders → Bought vs Sold Review** workflow and the isolated
offline launcher are documented in [Daily Bought vs Sold Review](docs/bought-vs-sold-review.md).

## T&M NTE forecast

Estimated Final Sales equals recorded sales to date plus the current NTE budget
multiplied by the percentage of work remaining. Missing sales values remain
flagged as provisional; internal costs are not substituted. Existing full-budget
overrides at 100% completion still apply. Dashboard, HTML and PDF reports use the
same calculation. Sales-forecast history is stored separately from the former
cost-based projection history.

## Server Bind Address

By default the app binds to `127.0.0.1` for local development.

For a Lightsail/Tailscale server, bind to all interfaces:

```powershell
$env:PM_TRACKER_HOST="0.0.0.0"
$env:PM_TRACKER_PORT="8765"
python app.py
```

On Linux/systemd, set:

```text
PM_TRACKER_HOST=0.0.0.0
PM_TRACKER_PORT=8765
```

## OCR Invoice Imports

Scanned/image-only vendor invoices require OCR support. Install the dependencies in `requirements.txt`; the app uses `rapidocr_onnxruntime` as an automatic OCR fallback when a PDF has no embedded text.

## First Workflow

1. Create a master project.
2. Add subprojects such as AC, FC, and E-Rack.
3. Add change orders under the correct subproject.
4. Import a Field Wise job summary `.xlsx`.
5. Review uncoded cost records and assign them to the correct subproject/change order.
6. Enter vendor invoices or upload invoice PDFs for manual coding.
7. Use the dashboard to view total job profitability and drill into subprojects/change orders.
