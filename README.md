# PM Project Cost Tracker

A local project-cost tracking tool for industrial electrical jobs with master projects, subprojects, change orders, Field Wise imports, vendor invoices, and profitability dashboards.

## Start

Run:

```powershell
python app.py
```

Then open:

```text
http://127.0.0.1:8765
```

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

## First Workflow

1. Create a master project.
2. Add subprojects such as AC, FC, and E-Rack.
3. Add change orders under the correct subproject.
4. Import a Field Wise job summary `.xlsx`.
5. Review uncoded cost records and assign them to the correct subproject/change order.
6. Enter vendor invoices or upload invoice PDFs for manual coding.
7. Use the dashboard to view total job profitability and drill into subprojects/change orders.
