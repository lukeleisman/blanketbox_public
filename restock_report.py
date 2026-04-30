#!/usr/bin/env python3
"""
restock_report.py — Generate and email restock reports.

Fetches live order data and inventory directly from the Sandstar API,
computes sales rates, and generates the report — fully self-contained
for GitHub Actions with no local files needed.

Falls back to precomputed docs/sales_rates.json if the order data
export times out.

Usage:
    python restock_report.py              # print text report
    python restock_report.py --email      # generate and email HTML report

Env vars:
    SANDSTAR_USERNAME / SANDSTAR_PASSWORD_HASH  (for API access)
    GMAIL_ADDRESS / GMAIL_APP_PASSWORD          (for email)
    RESTOCK_EMAIL_TO                            (recipient, optional)
"""

import csv
import io
import json
import math
import os
import re
import smtplib
import sys
import time
from calendar import monthrange
from datetime import date, datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

from ops_scraper import (
    login, scrape_all, build_inventory_json,
    LOGIN_URL, BASE_HEADERS,
)

# ============================================================
# CONFIGURATION
# ============================================================

MACHINE_DISPLAY_NAMES = {
    "333 North Broad Ambient": "333 North Broad (Ambient)",
    "333 North Broad Cooler":  "333 North Broad (Cooler)",
    "PhillyTrueCooler":        "Philly True Cooler",
    "TheHaven#2-cooler":       "Haven Cooler",
    "ChicagoDryCab2-Warren":   "Warren Park",
    "mHUBPrototypingShop":     "mHUB Prototyping Shop",
    "BorellisBox":             "Borelli's Box",
}

MACHINE_TO_LOCATION = {
    "333 North Broad Ambient": "doylestown",
    "333 North Broad Cooler":  "doylestown",
    "PhillyTrueCooler":        "philly",
    "TheHaven#2-cooler":       "haven",
    "ChicagoDryCab2-Warren":   "warren",
    "mHUBPrototypingShop":     "mhub",
    "BorellisBox":             "borellis",
}

STOCKING_ROUTES = {
    "philly_west": {
        "name": "West Philly",
        "machines": ["PhillyTrueCooler"],
        "frequency_days": 7,
    },
    "philly_suburbs": {
        "name": "Philly Suburbs",
        "machines": ["333 North Broad Ambient", "333 North Broad Cooler",
                     "TheHaven#2-cooler"],
        "frequency_days": 7,
    },
    "chicago": {
        "name": "Chicago",
        "machines": ["ChicagoDryCab2-Warren", "BorellisBox"],
        "frequency_days": 14,
    },
    "chicago_mhub": {
        "name": "Chicago - mHUB",
        "machines": ["mHUBPrototypingShop"],
        "frequency_days": 14,
    },
}

# Machine name normalization (order data → current inventory name)
MACHINE_NAME_MAP = {
    "333 North Broad Ambient": "333 North Broad Ambient",
    "333 North Broad Cooler":  "333 North Broad Cooler",
    "PhillyTrueCooler":        "PhillyTrueCooler",
    "TheHaven#2-cooler":       "TheHaven#2-cooler",
    "ChicagoDryCab2-Warren":   "ChicagoDryCab2-Warren",
    "ChicagoDryCab1":          "ChicagoDryCab1",
    "BorellisBox":             "BorellisBox",
    "TheHaven#1-ambient":      "333 North Broad Ambient",
    "ChicagoDryCab2":          "ChicagoDryCab2-Warren",
    "SpringGardenDryCab":      None,  # decommissioned
}

TEST_MACHINE_SERIALS = {"2405131108007687", "2405131108006876"}

# Sales rate parameters
SHORT_WINDOW_DAYS = 30
LONG_WINDOW_DAYS = 90
MIN_SALES_HIGH_CONFIDENCE = 10
MIN_SALES_MEDIUM_CONFIDENCE = 5

# Sandstar order export API
SALES_URL    = "https://webapi-us.sandstar.com/order/v2/findSaleInfo"
MESSAGE_URL  = "https://webapi-us.sandstar.com/message/findWebUserMessage"
DOWNLOAD_URL = "https://webapi-us.sandstar.com/homePage/download"
ORGAN_SN     = "000332"

POLL_INTERVAL = 3
POLL_TIMEOUT  = 180


# ============================================================
# ORDER DATA FETCHING (from Sandstar API)
# ============================================================

def fetch_order_data(token: str, days_back: int = 90) -> list[dict]:
    """
    Fetch recent order data from Sandstar's async export API.

    Returns a list of dicts with keys:
        machine, product_name, barcode, quantity, order_date
    """
    headers = {**BASE_HEADERS, "x-token": token}
    end = datetime.now()
    start = end - timedelta(days=days_back)
    start_str = start.strftime("%Y-%m-%d 00:00:00")
    end_str = end.strftime("%Y-%m-%d 23:59:59")

    print(f"  Triggering order export ({start_str[:10]} to {end_str[:10]})...")
    payload = {
        "organSns": ORGAN_SN,
        "startTime": start_str,
        "endTime": end_str,
        "phaseList": "3,4,5",
        "reportAsynchronous": True,
    }
    r = requests.post(SALES_URL + "?exportType=3", json=payload,
                      headers=headers, timeout=30)
    r.raise_for_status()

    # Poll for the download link
    t0 = time.time()
    deadline = t0 + POLL_TIMEOUT
    filename = None

    while time.time() < deadline:
        try:
            msg_payload = {"page": 1, "pageSize": 50, "readState": 0}
            mr = requests.post(MESSAGE_URL, json=msg_payload,
                               headers=headers, timeout=10)
            mr.raise_for_status()
            body = mr.json()
            messages = (body.get("data") or {}).get("resultList") or []
            print(f"  DEBUG: poll returned {len(messages)} message(s)", file=sys.stderr)

            for msg in messages:
                subject = msg.get("subject", "")
                content = msg.get("content", "")
                send_time = msg.get("sendTime", "")
                print(f"  DEBUG: msg subject={subject!r} sendTime={send_time!r} content={content[:80]!r}", file=sys.stderr)
                try:
                    send_ts = datetime.strptime(send_time, "%Y-%m-%d %H:%M:%S").timestamp()
                except ValueError:
                    send_ts = 0
                if send_ts < t0:
                    print(f"  DEBUG: skipping (too old: {send_ts:.0f} < {t0:.0f})", file=sys.stderr)
                    continue
                if subject == "OrderDetailsReport":
                    m = re.search(r'fileName=([^\s"\'<>&]+\.xlsx)', content)
                    if m:
                        filename = m.group(1)
                        print(f"  Export ready: {filename}")
                        break
                    else:
                        print(f"  DEBUG: subject matched but no fileName in content: {content!r}", file=sys.stderr)
        except Exception as e:
            print(f"  Warning: message poll failed ({e})", file=sys.stderr)

        if filename:
            break
        remaining = int(deadline - time.time())
        if remaining > 0:
            print(f"  Waiting for export... ({remaining}s left)")
            time.sleep(POLL_INTERVAL)

    if not filename:
        return None  # Signal caller to use fallback

    # Download the XLSX
    print("  Downloading order data...")
    r = requests.get(DOWNLOAD_URL, params={"fileName": filename},
                     headers=headers, timeout=120, stream=True)
    r.raise_for_status()
    xlsx_bytes = io.BytesIO(r.content)

    # Parse the XLSX — it has a 2-row header
    try:
        import openpyxl
    except ImportError:
        # Fallback: try pandas
        import pandas as pd
        df = pd.read_excel(xlsx_bytes, header=2)
        return _dataframe_to_orders(df)

    wb = openpyxl.load_workbook(xlsx_bytes, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 3:
        return []

    # Row 3 (index 2) has the actual column headers
    col_headers = [str(c).strip() if c else "" for c in rows[2]]
    col_map = {name: i for i, name in enumerate(col_headers)}

    orders = []
    for row in rows[3:]:
        machine_raw = str(row[col_map.get("Machine", 3)] or "").strip()
        payment_status = str(row[col_map.get("Payment status", 11)] or "").strip()

        if payment_status != "Success":
            continue
        if machine_raw in TEST_MACHINE_SERIALS or re.match(r"^\d+$", machine_raw):
            continue

        machine = MACHINE_NAME_MAP.get(machine_raw)
        if not machine:
            continue

        order_time_raw = row[col_map.get("Order time", 8)]
        if isinstance(order_time_raw, datetime):
            order_date = order_time_raw
        else:
            try:
                order_date = datetime.strptime(str(order_time_raw), "%Y-%m-%d %H:%M:%S")
            except (ValueError, TypeError):
                continue

        orders.append({
            "machine": machine,
            "product_name": str(row[col_map.get("Product name", 16)] or "").strip(),
            "barcode": str(row[col_map.get("Product barcode", 15)] or "").strip(),
            "quantity": int(row[col_map.get("Quantity", 17)] or 0),
            "order_date": order_date,
        })

    print(f"  Parsed {len(orders)} orders")
    return orders


def _dataframe_to_orders(df) -> list[dict]:
    """Convert a pandas DataFrame from the XLSX into our order dict format."""
    import pandas as pd
    orders = []
    for _, row in df.iterrows():
        machine_raw = str(row.get("Machine", "")).strip()
        payment = str(row.get("Payment status", "")).strip()
        if payment != "Success":
            continue
        if machine_raw in TEST_MACHINE_SERIALS or re.match(r"^\d+$", machine_raw):
            continue
        machine = MACHINE_NAME_MAP.get(machine_raw)
        if not machine:
            continue

        order_date = pd.to_datetime(row.get("Order time"), errors="coerce")
        if pd.isna(order_date):
            continue

        orders.append({
            "machine": machine,
            "product_name": str(row.get("Product name", "")).strip(),
            "barcode": str(row.get("Product barcode", "")).strip(),
            "quantity": int(row.get("Quantity", 0)),
            "order_date": order_date.to_pydatetime(),
        })
    return orders


# ============================================================
# SALES RATE COMPUTATION
# ============================================================

def compute_sales_rates(orders: list[dict], as_of: date = None) -> tuple[dict, dict]:
    """
    Compute sales rates from order data.

    Returns (rate_lookup, global_rates):
        rate_lookup: {(machine, product_name): {daily_rate, confidence, ...}}
        global_rates: {product_name: avg_daily_rate_per_machine}
    """
    if as_of is None:
        as_of = date.today()

    as_of_dt = datetime(as_of.year, as_of.month, as_of.day)
    short_start = as_of_dt - timedelta(days=SHORT_WINDOW_DAYS)
    long_start = as_of_dt - timedelta(days=LONG_WINDOW_DAYS)

    # Aggregate by (machine, product_name)
    agg = {}  # (machine, product) → {short_qty, long_qty, barcodes}
    global_short = {}  # product → total qty in short window
    global_machines = {}  # product → set of machines

    for o in orders:
        key = (o["machine"], o["product_name"])
        if key not in agg:
            agg[key] = {"short": 0, "long": 0, "barcode": o["barcode"]}

        if o["order_date"] >= short_start:
            agg[key]["short"] += o["quantity"]
            global_short.setdefault(o["product_name"], 0)
            global_short[o["product_name"]] += o["quantity"]
            global_machines.setdefault(o["product_name"], set())
            global_machines[o["product_name"]].add(o["machine"])

        if o["order_date"] >= long_start:
            agg[key]["long"] += o["quantity"]

        agg[key]["barcode"] = o["barcode"]  # keep latest

    # Global rate = total across machines / days / num_machines
    global_rates = {}
    for pname, qty in global_short.items():
        n_machines = len(global_machines.get(pname, {1}))
        global_rates[pname] = qty / SHORT_WINDOW_DAYS / max(n_machines, 1)

    # Per-machine rates with confidence blending
    rate_lookup = {}
    for (machine, product_name), data in agg.items():
        rate_short = data["short"] / SHORT_WINDOW_DAYS
        rate_long = data["long"] / LONG_WINDOW_DAYS
        rate_global = global_rates.get(product_name, 0)

        if data["short"] >= MIN_SALES_HIGH_CONFIDENCE:
            confidence = "high"
            rate_blended = 0.80 * rate_short + 0.20 * rate_long
        elif data["short"] >= MIN_SALES_MEDIUM_CONFIDENCE:
            confidence = "medium"
            rate_blended = 0.50 * rate_short + 0.30 * rate_long + 0.20 * rate_global
        elif data["short"] > 0:
            confidence = "low"
            rate_blended = 0.30 * rate_short + 0.30 * rate_long + 0.40 * rate_global
        else:
            confidence = "no_recent"
            rate_blended = (0.40 * rate_long + 0.60 * rate_global
                            if data["long"] > 0 else rate_global)

        rate_lookup[(machine, product_name)] = {
            "daily_rate": round(rate_blended, 4),
            "confidence": confidence,
            "total_sold_30d": data["short"],
            "total_sold_90d": data["long"],
        }

    return rate_lookup, global_rates


def load_sales_rates_json(path: str = "docs/sales_rates.json") -> tuple[dict, dict]:
    """Fallback: load precomputed sales rates from JSON."""
    with open(path) as f:
        data = json.load(f)

    lookup = {}
    global_rates_raw = {}
    for r in data["rates"]:
        lookup[(r["machine"], r["product_name"])] = r
        global_rates_raw.setdefault(r["product_name"], []).append(r["daily_rate"])

    global_rates = {n: sum(v) / len(v) for n, v in global_rates_raw.items()}
    return lookup, global_rates


# ============================================================
# REPORT BUILDING
# ============================================================

def fetch_live_inventory(token: str) -> dict:
    """Fetch current inventory from Sandstar API, with fallback to inventory.json."""
    machines = scrape_all(token)
    updated = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    total_products = sum(len(m["products"]) for m in machines)
    if total_products == 0 and machines:
        fallback = "docs/inventory.json"
        if os.path.exists(fallback):
            print(f"  WARNING: Live API returned 0 products across all machines — "
                  f"falling back to {fallback}", file=sys.stderr)
            with open(fallback) as f:
                return json.load(f)
        print("  ERROR: Live API returned 0 products and no fallback inventory.json found.",
              file=sys.stderr)

    return build_inventory_json(machines, updated)


def build_report_data(inventory: dict, rate_lookup: dict, global_rates: dict) -> list:
    """Combine inventory with sales rates into restock report data."""
    machine_route = {}
    machine_freq = {}
    for route_id, route in STOCKING_ROUTES.items():
        for m in route["machines"]:
            machine_route[m] = route_id
            machine_freq[m] = route["frequency_days"]

    today = date.today()
    machines_out = []

    for m in inventory["machines"]:
        mname = m["machine_name"]
        if mname not in MACHINE_DISPLAY_NAMES:
            print(f"  WARNING: Unknown machine name '{mname}' (freezer {m.get('freezer_id', '?')}) — "
                  f"not in MACHINE_DISPLAY_NAMES. Update the config maps.", file=sys.stderr)
        freq = machine_freq.get(mname, 7)
        route_id = machine_route.get(mname, "unassigned")

        products_out = []
        for p in m["products"]:
            pname = p["name"]
            stock = p["stock"]
            capacity = p["capacity"] or 0

            rate_info = rate_lookup.get((mname, pname))
            if rate_info:
                daily_rate = rate_info["daily_rate"]
                confidence = rate_info["confidence"]
            elif pname in global_rates:
                daily_rate = global_rates[pname]
                confidence = "global_only"
            else:
                daily_rate = 0
                confidence = "no_data"

            if daily_rate > 0 and stock > 0:
                est_days = stock / daily_rate
                est_sellout = today + timedelta(days=int(est_days))
            elif stock == 0:
                est_days = 0
                est_sellout = today
            else:
                est_days = None
                est_sellout = None

            is_oos = stock == 0
            needs_restock = is_oos or (est_days is not None and est_days < freq)

            if capacity > 0:
                qty_to_fill = max(0, capacity - stock)
            elif daily_rate > 0:
                qty_to_fill = max(0, int(math.ceil(daily_rate * freq)) - stock)
            else:
                qty_to_fill = 0

            products_out.append({
                "name": pname,
                "barcode": p["barcode"],
                "stock": stock,
                "capacity": capacity,
                "daily_rate": daily_rate,
                "confidence": confidence,
                "est_days_remaining": round(est_days, 1) if est_days is not None else None,
                "est_sellout_date": est_sellout.isoformat() if est_sellout else None,
                "needs_restock": needs_restock,
                "is_oos": is_oos,
                "qty_to_fill": qty_to_fill,
            })

        products_out.sort(key=lambda p: (
            not p["is_oos"],
            p["est_days_remaining"] if p["est_days_remaining"] is not None else 9999,
        ))

        machines_out.append({
            "machine_name": mname,
            "display_name": MACHINE_DISPLAY_NAMES.get(mname, mname),
            "location_id": MACHINE_TO_LOCATION.get(mname, "unknown"),
            "route": route_id,
            "frequency_days": freq,
            "updated": inventory["updated"],
            "total_skus": len(products_out),
            "oos_count": sum(1 for p in products_out if p["is_oos"]),
            "needs_restock_count": sum(1 for p in products_out if p["needs_restock"]),
            "products": products_out,
        })

    return machines_out


# ============================================================
# REPORT FORMATTING
# ============================================================

def format_html_report(machines: list) -> str:
    """Generate HTML email report."""
    today_str = date.today().strftime("%B %d, %Y")
    updated = machines[0]["updated"] if machines else "N/A"

    css = """<style>
      body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
             color: #1a1a1a; max-width: 900px; margin: 0 auto; padding: 16px; }
      h1 { color: #1C3D5A; font-size: 22px; margin-bottom: 4px; }
      .meta { color: #666; font-size: 13px; margin-bottom: 20px; }
      .route-header { background: #1C3D5A; color: white; padding: 8px 14px;
                      font-size: 16px; font-weight: 600; margin-top: 24px;
                      border-radius: 4px 4px 0 0; }
      .machine { border: 1px solid #ddd; padding: 12px 14px; margin-bottom: 16px; }
      .machine-name { font-size: 16px; font-weight: 600; color: #1C3D5A; }
      .machine-stats { font-size: 13px; color: #666; margin: 4px 0 10px 0; }
      .section-label { font-weight: 600; font-size: 13px; margin: 10px 0 4px 0; }
      .oos { color: #c0392b; }
      .low { color: #d97706; }
      .ok-summary { color: #1a7a4a; font-size: 13px; margin: 8px 0; }
      table { border-collapse: collapse; width: 100%; font-size: 13px; margin-bottom: 10px; }
      th { text-align: left; padding: 4px 8px; border-bottom: 2px solid #ddd;
           font-size: 12px; color: #666; }
      td { padding: 4px 8px; border-bottom: 1px solid #eee; }
      tr.oos-row td { background: #fef2f2; }
      tr.low-row td { background: #fffbeb; }
      .confidence { font-size: 11px; color: #999; }
      .summary-table th { background: #f0f4fa; }
    </style>"""

    html = [f"<html><head>{css}</head><body>"]
    html.append(f"<h1>Blanket Box Restock Report</h1>")
    html.append(f'<div class="meta">{today_str} &mdash; Inventory as of {updated}</div>')

    total_oos = sum(m["oos_count"] for m in machines)
    total_restock = sum(m["needs_restock_count"] for m in machines)
    html.append(f'<div style="font-size:15px; margin-bottom:16px;">'
                f'<span class="oos">{total_oos} items out of stock</span> &middot; '
                f'<span class="low">{total_restock} items need restocking</span></div>')

    routes = {}
    for m in machines:
        routes.setdefault(m["route"], []).append(m)

    for route_id, route_machines in routes.items():
        route_cfg = STOCKING_ROUTES.get(route_id, {"name": route_id})
        html.append(f'<div class="route-header">{route_cfg.get("name", route_id)}</div>')

        for m in route_machines:
            html.append('<div class="machine">')
            html.append(f'<div class="machine-name">{m["display_name"]}</div>')
            html.append(
                f'<div class="machine-stats">'
                f'Every {m["frequency_days"]}d &middot; '
                f'{m["total_skus"]} items &middot; '
                f'<span class="oos">{m["oos_count"]} OOS</span> &middot; '
                f'<span class="low">{m["needs_restock_count"]} need restock</span>'
                f'</div>'
            )

            oos = [p for p in m["products"] if p["is_oos"]]
            low = [p for p in m["products"] if p["needs_restock"] and not p["is_oos"]]
            ok = [p for p in m["products"] if not p["needs_restock"]]

            if oos:
                html.append(f'<div class="section-label oos">Out of Stock ({len(oos)})</div>')
                html.append(_product_table(oos))
            if low:
                html.append(f'<div class="section-label low">Low Stock ({len(low)})</div>')
                html.append(_product_table(low))
            if ok:
                html.append(f'<div class="ok-summary">{len(ok)} items OK</div>')

            html.append("</div>")

    totals = {}
    for m in machines:
        for p in m["products"]:
            if p["needs_restock"] and p["qty_to_fill"] > 0:
                totals.setdefault(p["name"], 0)
                totals[p["name"]] += p["qty_to_fill"]

    if totals:
        html.append('<h2 style="color:#1C3D5A; font-size:18px; margin-top:28px;">'
                    'Total Restock Quantities</h2>')
        html.append('<table class="summary-table"><tr><th>Product</th><th>Qty</th></tr>')
        for name, qty in sorted(totals.items(), key=lambda x: -x[1]):
            html.append(f"<tr><td>{name}</td><td>{qty}</td></tr>")
        html.append("</table>")

    html.append("</body></html>")
    return "\n".join(html)


def _product_table(products: list) -> str:
    rows = ["<table>",
            "<tr><th>Product</th><th>Stock</th><th>Rate</th>"
            "<th>Days Left</th><th>Sellout</th><th>Fill</th><th></th></tr>"]
    for p in products:
        cls = "oos-row" if p["is_oos"] else "low-row"
        rate = f'{p["daily_rate"]:.2f}/d' if p["daily_rate"] > 0 else "&mdash;"
        days = f'{p["est_days_remaining"]:.0f}d' if p["est_days_remaining"] is not None else "&mdash;"
        sellout = p["est_sellout_date"] or "&mdash;"
        conf = f'<span class="confidence">{p["confidence"]}</span>'
        rows.append(
            f'<tr class="{cls}"><td>{p["name"]}</td><td>{p["stock"]}</td>'
            f'<td>{rate}</td><td>{days}</td><td>{sellout}</td>'
            f'<td>{p["qty_to_fill"]}</td><td>{conf}</td></tr>'
        )
    rows.append("</table>")
    return "\n".join(rows)


def format_text_report(machines: list) -> str:
    """Simple text summary for stdout."""
    lines = [f"BLANKET BOX RESTOCK REPORT — {date.today().strftime('%B %d, %Y')}",
             f"Inventory snapshot: {machines[0]['updated'] if machines else 'N/A'}",
             "=" * 60]

    routes = {}
    for m in machines:
        routes.setdefault(m["route"], []).append(m)

    for route_id, route_machines in routes.items():
        route_cfg = STOCKING_ROUTES.get(route_id, {"name": route_id})
        lines.append(f"\n--- {route_cfg.get('name', route_id).upper()} ---")

        for m in route_machines:
            lines.append(f"\n  {m['display_name']}  "
                         f"({m['oos_count']} OOS, {m['needs_restock_count']} need restock)")
            for p in m["products"]:
                if p["is_oos"]:
                    lines.append(f"    [OOS]  {p['name']}")
                elif p["needs_restock"]:
                    days = f"{p['est_days_remaining']:.0f}d" if p["est_days_remaining"] else "?"
                    lines.append(f"    [LOW]  {p['name']}  stock={p['stock']}  {days} left")

    lines.append("")
    return "\n".join(lines)


# ============================================================
# EMAIL
# ============================================================

def send_email(html_body: str):
    sender = os.environ.get("GMAIL_ADDRESS", "luke.blanketboxvending@gmail.com")
    password = os.environ["GMAIL_APP_PASSWORD"]
    recipient = os.environ.get("RESTOCK_EMAIL_TO", sender)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Restock Report — {date.today().strftime('%b %d, %Y')}"
    msg["From"] = sender
    msg["To"] = recipient
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, [recipient], msg.as_string())

    print(f"Email sent to {recipient}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("Logging in to Sandstar OPS...")
    token = login()

    # Fetch inventory first while the token is fresh (before the long order export poll)
    print("Fetching inventory...")
    inventory = fetch_live_inventory(token)
    print(f"  {len(inventory['machines'])} machines")

    # Re-login before the order export poll in case the token expires during the wait
    print("Logging in to Sandstar OPS (order export)...")
    token = login()

    # Try fetching live order data (last 90 days)
    print("Fetching order data...")
    orders = fetch_order_data(token, days_back=LONG_WINDOW_DAYS)

    if orders is not None:
        print(f"Computing sales rates from {len(orders)} orders...")
        rate_lookup, global_rates = compute_sales_rates(orders)
        print(f"  {len(rate_lookup)} product×machine rates")
    else:
        # Fallback to precomputed rates
        rates_path = "docs/sales_rates.json"
        if os.path.exists(rates_path):
            print(f"  Order export timed out — falling back to {rates_path}")
            rate_lookup, global_rates = load_sales_rates_json(rates_path)
            print(f"  Loaded {len(rate_lookup)} rates from file")
        else:
            print("  ERROR: No order data and no fallback sales_rates.json")
            sys.exit(1)

    # Build report
    machines = build_report_data(inventory, rate_lookup, global_rates)

    if "--email" in sys.argv:
        html = format_html_report(machines)
        send_email(html)
    else:
        print()
        print(format_text_report(machines))


if __name__ == "__main__":
    main()
