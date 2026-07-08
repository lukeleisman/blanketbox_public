# Blanket Box Vending — Public

Public-facing code for [blanketboxvending.com](https://blanketboxvending.com).

## Live Prices Page

`blanketboxvending.com/prices` — shows live product photos and prices for each machine location, updated every 30 minutes via GitHub Actions.

### How it works

1. `ops_scraper.py` logs into the Sandstar OPS API and fetches current products and prices for all machines
2. GitHub Actions runs the scraper every 30 minutes and commits `docs/products.json`
3. GitHub Pages serves `docs/` — the prices app JS and JSON data
4. The WordPress page at `/prices` loads `docs/prices/app.js` which fetches the JSON and renders the UI

### WordPress embed

Add a Custom HTML block to the `/prices` page:
```html
<div id="bb-prices-app" style="min-height:70vh"></div>
<script src="https://lukeleisman.github.io/blanketbox_public/prices/app.js"></script>
```

### Updating locations

When a machine moves, edit `LOCATION_GROUPS` in `ops_scraper.py`.  
The URL slugs (used in QR codes) should never be renamed after QR codes are printed.

### Local test run

```bash
export SANDSTAR_USERNAME="info@blanketboxvending.com"
export SANDSTAR_PASSWORD_HASH="<hash from 1Password>"
python ops_scraper.py
```

## mHUB Sales Report

`mhub_sales_report.py` — daily email for the mHUB Prototyping Shop machine. Runs at 8 AM Central via GitHub Actions. Sends a 7-day summary on Mondays; other days it alerts only if the previous day's sales exceed the threshold.

Env vars: `SANDSTAR_USERNAME`, `SANDSTAR_PASSWORD_HASH`, `GMAIL_ADDRESS`, `GMAIL_APP_PASSWORD`, `MHUB_REPORT_EMAIL_TO` (comma-separated for multiple recipients, default `luke@blanketboxvending.com`).

```bash
python mhub_sales_report.py --email
```

## Restock Report

`restock_report.py` — restock report for all machines, emailed Mondays and Thursdays at 7 AM Eastern via GitHub Actions. Pulls live order and inventory data from the Sandstar API and computes sales rates.

Env vars: `SANDSTAR_USERNAME`, `SANDSTAR_PASSWORD_HASH`, `GMAIL_ADDRESS`, `GMAIL_APP_PASSWORD`, `RESTOCK_EMAIL_TO`.

```bash
python restock_report.py --email
```

## Sales Data Verification

`verify_sales_data.py` — cross-checks Sandstar order XLSX exports against `docs/inventory_history.csv` to validate sales data accuracy. Run manually, not scheduled.

```bash
python verify_sales_data.py --start 2026-04-07 --end 2026-04-30
```

## Data Files (`docs/`)

- `products.json` — current products/prices per machine, feeds the live prices page
- `inventory.json` / `inventory.csv` — current stock snapshot
- `inventory_history.csv` — 30-min stock snapshots over time, used to derive sales rates
- `sales_rates.json` — precomputed sales rates, fallback for the restock report if the live order export times out
