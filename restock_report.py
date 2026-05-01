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

# Restocking totals groupings
REGION_ROUTES = {
    "Philly": {"philly_west", "philly_suburbs"},
    "Chicago": {"chicago", "chicago_mhub"},
}

# Sandstar order export API
SALES_URL    = "https://webapi-us.sandstar.com/order/v2/findSaleInfo"
MESSAGE_URL  = "https://webapi-us.sandstar.com/message/findWebUserMessage"
DOWNLOAD_URL = "https://webapi-us.sandstar.com/homePage/download"
ORGAN_SN     = "000332"

POLL_INTERVAL = 3
POLL_TIMEOUT  = 90


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

    # Snapshot ALL existing OrderDetailsReport message sendTimes (read + unread)
    # before triggering. Snapshotting only unread would miss previously-read exports
    # (e.g. from monthly report runs) and falsely match them as "new" during polling.
    pre_trigger_times = set()
    try:
        sr = requests.post(MESSAGE_URL, json={"page": 1, "pageSize": 50},
                           headers=headers, timeout=10)
        sr.raise_for_status()
        for msg in (sr.json().get("data") or {}).get("resultList") or []:
            if msg.get("subject") == "OrderDetailsReport":
                pre_trigger_times.add(msg.get("sendTime"))
    except Exception as e:
        print(f"  Warning: pre-trigger snapshot failed ({e})", file=sys.stderr)

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
    trigger_body = r.json()
    print(f"  Trigger response: status={trigger_body.get('status')}, "
          f"message={trigger_body.get('message')}")
    if trigger_body.get("data"):
        print(f"  Trigger data: {trigger_body['data']}")

    # Poll for a new OrderDetailsReport message (not in pre-trigger snapshot).
    # Try unread messages first, then fall back to all messages (in case the
    # notification was marked read by another session or the web portal).
    deadline = time.time() + POLL_TIMEOUT
    filename = None

    while time.time() < deadline:
        for read_state in (0, None):
            if filename:
                break
            try:
                poll_payload = {"page": 1, "pageSize": 50}
                if read_state is not None:
                    poll_payload["readState"] = read_state
                mr = requests.post(MESSAGE_URL, json=poll_payload,
                                   headers=headers, timeout=10)
                mr.raise_for_status()
                messages = (mr.json().get("data") or {}).get("resultList") or []

                for msg in messages:
                    if msg.get("subject") != "OrderDetailsReport":
                        continue
                    if msg.get("sendTime") in pre_trigger_times:
                        continue
                    m = re.search(r'fileName=([^\s"\'<>&]+\.xlsx)', msg.get("content", ""))
                    if m:
                        filename = m.group(1)
                        print(f"  Export ready: {filename} (read_state={read_state})")
                        break
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

    # Parse the XLSX — it has a 2-row header (rows 0–1 are title/subtitle, row 2 is columns).
    # Use pandas when available: it handles sparse worksheet XML correctly.
    # openpyxl read_only=True streaming mode can stop at the title row on some Sandstar exports.
    try:
        import pandas as pd
        df = pd.read_excel(xlsx_bytes, header=2, engine="openpyxl")
        print(f"  XLSX rows (pandas): {len(df)}", file=sys.stderr)
        if len(df) == 0:
            return []
        return _dataframe_to_orders(df)
    except ImportError:
        pass

    import openpyxl
    wb = openpyxl.load_workbook(xlsx_bytes, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    print(f"  XLSX raw row count (openpyxl): {len(rows)}", file=sys.stderr)
    if rows:
        print(f"  XLSX row[0]: {rows[0]}", file=sys.stderr)
    if len(rows) >= 3:
        print(f"  XLSX row[2] (expected headers): {rows[2]}", file=sys.stderr)

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

def _poisson_sigma(count: int, window_days: int, n_machines: int = 1) -> float:
    """Poisson counting uncertainty: σ_rate = √N / T / n_machines."""
    return math.sqrt(max(count, 0)) / window_days / max(n_machines, 1)


def _blend_sigma(weights: list[float], sigmas: list[float]) -> float:
    """Error propagation for weighted sum: σ = √(Σ wi² σi²)."""
    return math.sqrt(sum(w * w * s * s for w, s in zip(weights, sigmas)))


def compute_sales_rates(orders: list[dict], as_of: date = None) -> tuple[dict, dict, dict]:
    """
    Compute sales rates from order data.

    Returns (rate_lookup, global_rates_short, global_rates_long):
        rate_lookup:        {(machine, product_name): {daily_rate, rate_sigma, confidence, ...}}
        global_rates_short: {product_name: avg_daily_rate_per_machine over 30d}
        global_rates_long:  {product_name: avg_daily_rate_per_machine over 90d}
    """
    if as_of is None:
        as_of = date.today()

    as_of_dt = datetime(as_of.year, as_of.month, as_of.day)
    short_start = as_of_dt - timedelta(days=SHORT_WINDOW_DAYS)
    long_start = as_of_dt - timedelta(days=LONG_WINDOW_DAYS)

    # Aggregate by (machine, product_name)
    agg = {}  # (machine, product) → {short_qty, long_qty, barcode, last_order_date}
    global_short_qty = {}   # product → total qty in 30d window across machines
    global_short_mach = {}  # product → set of machines (30d)
    global_long_qty = {}    # product → total qty in 90d window across machines
    global_long_mach = {}   # product → set of machines (90d)

    for o in orders:
        key = (o["machine"], o["product_name"])
        if key not in agg:
            agg[key] = {"short": 0, "long": 0, "barcode": o["barcode"], "last_order_date": None}

        if o["order_date"] >= short_start:
            agg[key]["short"] += o["quantity"]
            global_short_qty.setdefault(o["product_name"], 0)
            global_short_qty[o["product_name"]] += o["quantity"]
            global_short_mach.setdefault(o["product_name"], set())
            global_short_mach[o["product_name"]].add(o["machine"])

        if o["order_date"] >= long_start:
            agg[key]["long"] += o["quantity"]
            global_long_qty.setdefault(o["product_name"], 0)
            global_long_qty[o["product_name"]] += o["quantity"]
            global_long_mach.setdefault(o["product_name"], set())
            global_long_mach[o["product_name"]].add(o["machine"])

        agg[key]["barcode"] = o["barcode"]  # keep latest
        if agg[key]["last_order_date"] is None or o["order_date"] > agg[key]["last_order_date"]:
            agg[key]["last_order_date"] = o["order_date"]

    # Global rates (per-machine average) and their Poisson uncertainties
    global_rates_short = {}
    global_sigma_short = {}
    for pname, qty in global_short_qty.items():
        n = max(len(global_short_mach.get(pname, set())), 1)
        global_rates_short[pname] = qty / SHORT_WINDOW_DAYS / n
        global_sigma_short[pname] = _poisson_sigma(qty, SHORT_WINDOW_DAYS, n)

    global_rates_long = {}
    global_sigma_long = {}
    for pname, qty in global_long_qty.items():
        n = max(len(global_long_mach.get(pname, set())), 1)
        global_rates_long[pname] = qty / LONG_WINDOW_DAYS / n
        global_sigma_long[pname] = _poisson_sigma(qty, LONG_WINDOW_DAYS, n)

    # Per-machine rates with confidence blending, including both global windows.
    # Recent global (30d) is weighted more heavily than long-term global (90d).
    rate_lookup = {}
    for (machine, product_name), data in agg.items():
        rate_short = data["short"] / SHORT_WINDOW_DAYS
        rate_long = data["long"] / LONG_WINDOW_DAYS
        rgs = global_rates_short.get(product_name, 0)
        rgl = global_rates_long.get(product_name, 0)

        sig_short = _poisson_sigma(data["short"], SHORT_WINDOW_DAYS)
        sig_long  = _poisson_sigma(data["long"], LONG_WINDOW_DAYS)
        sig_gs    = global_sigma_short.get(product_name, 0)
        sig_gl    = global_sigma_long.get(product_name, 0)

        if data["short"] >= MIN_SALES_HIGH_CONFIDENCE:
            confidence = "high"
            w = [0.80, 0.20]
            rate_blended = w[0] * rate_short + w[1] * rate_long
            rate_sigma   = _blend_sigma(w, [sig_short, sig_long])
        elif data["short"] >= MIN_SALES_MEDIUM_CONFIDENCE:
            confidence = "medium"
            w = [0.50, 0.20, 0.20, 0.10]
            rate_blended = w[0]*rate_short + w[1]*rate_long + w[2]*rgs + w[3]*rgl
            rate_sigma   = _blend_sigma(w, [sig_short, sig_long, sig_gs, sig_gl])
        elif data["short"] > 0:
            confidence = "low"
            w = [0.25, 0.20, 0.35, 0.20]
            rate_blended = w[0]*rate_short + w[1]*rate_long + w[2]*rgs + w[3]*rgl
            rate_sigma   = _blend_sigma(w, [sig_short, sig_long, sig_gs, sig_gl])
        else:
            confidence = "no_recent"
            if data["long"] > 0:
                w = [0.25, 0.45, 0.30]
                rate_blended = w[0]*rate_long + w[1]*rgs + w[2]*rgl
                rate_sigma   = _blend_sigma(w, [sig_long, sig_gs, sig_gl])
            else:
                w = [0.60, 0.40]
                rate_blended = w[0]*rgs + w[1]*rgl
                rate_sigma   = _blend_sigma(w, [sig_gs, sig_gl])

        rate_lookup[(machine, product_name)] = {
            "daily_rate": round(rate_blended, 4),
            "rate_sigma": round(rate_sigma, 4),
            "confidence": confidence,
            "total_sold_30d": data["short"],
            "total_sold_90d": data["long"],
            "last_order_date": data.get("last_order_date"),
        }

    return rate_lookup, global_rates_short, global_sigma_short


def load_sales_rates_json(path: str = "docs/sales_rates.json") -> tuple[dict, dict, dict]:
    """Fallback: load precomputed sales rates from JSON."""
    with open(path) as f:
        data = json.load(f)

    lookup = {}
    global_rates_raw = {}
    for r in data["rates"]:
        lookup[(r["machine"], r["product_name"])] = r
        global_rates_raw.setdefault(r["product_name"], []).append(r["daily_rate"])

    global_rates = {n: sum(v) / len(v) for n, v in global_rates_raw.items()}
    return lookup, global_rates, {}


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


def build_report_data(inventory: dict, rate_lookup: dict, global_rates: dict,
                      global_sigmas: dict = None) -> list:
    """Combine inventory with sales rates into restock report data."""
    if global_sigmas is None:
        global_sigmas = {}

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
                rate_sigma  = rate_info.get("rate_sigma", 0)
                confidence  = rate_info["confidence"]
                # Range: compare blended rate vs. global rate (captures systematic bias)
                global_r = global_rates.get(pname, 0)
                global_s = global_sigmas.get(pname, 0)
                if global_r > 0:
                    if global_r >= daily_rate:
                        rate_high, sig_high = global_r, global_s
                        rate_low,  sig_low  = daily_rate, rate_sigma
                    else:
                        rate_high, sig_high = daily_rate, rate_sigma
                        rate_low,  sig_low  = global_r, global_s
                else:
                    rate_high, sig_high = daily_rate, rate_sigma
                    rate_low,  sig_low  = daily_rate, rate_sigma
            elif pname in global_rates:
                daily_rate  = global_rates[pname]
                rate_sigma  = global_sigmas.get(pname, 0)
                confidence  = "global_only"
                rate_high, sig_high = daily_rate, rate_sigma
                rate_low,  sig_low  = daily_rate, rate_sigma
            else:
                daily_rate = rate_sigma = 0
                confidence = "no_data"
                rate_high = sig_high = rate_low = sig_low = 0

            if stock == 0:
                est_days    = 0
                last_dt     = rate_info.get("last_order_date") if rate_info else None
                est_sellout = last_dt.date() if isinstance(last_dt, datetime) else last_dt
                sellout_early = est_sellout
                sellout_late  = est_sellout
            elif daily_rate > 0:
                est_days    = stock / daily_rate
                est_sellout = today + timedelta(days=math.ceil(est_days))
                # Pessimistic: highest plausible rate → earliest sellout
                r_pess = rate_high + sig_high
                sellout_early = today + timedelta(days=math.ceil(stock / r_pess))
                # Optimistic: lowest plausible rate → latest sellout (floor to avoid ÷0)
                r_opt = max(rate_low - sig_low, 1e-9)
                days_late = stock / r_opt
                sellout_late = (today + timedelta(days=math.ceil(days_late))
                                if days_late <= 365 else None)
            else:
                est_days = est_sellout = sellout_early = sellout_late = None

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
                "pct_remaining": round(stock / capacity * 100) if capacity > 0 else None,
                "daily_rate": daily_rate,
                "rate_sigma": rate_sigma,
                "confidence": confidence,
                "est_days_remaining": round(est_days, 1) if est_days is not None else None,
                "est_sellout_date": est_sellout.isoformat() if est_sellout else None,
                "sellout_early": sellout_early.isoformat() if sellout_early else None,
                "sellout_late": sellout_late.isoformat() if sellout_late else None,
                "needs_restock": needs_restock,
                "is_oos": is_oos,
                "qty_to_fill": qty_to_fill,
            })

        products_out.sort(key=lambda p: (
            not p["is_oos"],
            p["sellout_early"] or "9999-99-99",
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
      .ok  { color: #1a7a4a; }
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
            ok  = [p for p in m["products"] if not p["needs_restock"]]

            if oos:
                html.append(f'<div class="section-label oos">Out of Stock ({len(oos)})</div>')
                html.append(_product_table(oos))
            if low:
                html.append(f'<div class="section-label low">Low Stock ({len(low)})</div>')
                html.append(_product_table(low))
            if ok:
                html.append(f'<div class="section-label ok">Stocked ({len(ok)})</div>')
                html.append(_product_table(ok, row_class=""))

            html.append("</div>")

    any_totals = False
    for region_name, route_ids in REGION_ROUTES.items():
        region_totals = {}
        for m in machines:
            if m["route"] in route_ids:
                for p in m["products"]:
                    if p["needs_restock"] and p["qty_to_fill"] > 0:
                        region_totals[p["name"]] = region_totals.get(p["name"], 0) + p["qty_to_fill"]
        if region_totals:
            if not any_totals:
                html.append('<h2 style="color:#1C3D5A; font-size:18px; margin-top:28px;">'
                            'Restock Quantities</h2>')
                any_totals = True
            html.append(f'<h3 style="color:#1C3D5A; font-size:15px; margin: 14px 0 4px 0;">'
                        f'{region_name}</h3>')
            html.append('<table class="summary-table"><tr><th>Product</th><th>Qty</th></tr>')
            for name, qty in sorted(region_totals.items(), key=lambda x: -x[1]):
                html.append(f"<tr><td>{name}</td><td>{qty}</td></tr>")
            html.append("</table>")

    html.append("</body></html>")
    return "\n".join(html)


def _product_table(products: list, row_class: str = None) -> str:
    rows = ["<table>",
            "<tr><th>Product</th><th>Stock</th><th>%</th><th>Rate</th>"
            "<th>Days Left</th><th>Sellout Range</th><th>Fill</th><th></th></tr>"]
    for p in products:
        if row_class is not None:
            cls = row_class
        else:
            cls = "oos-row" if p["is_oos"] else "low-row"
        pct  = f'{p["pct_remaining"]}%' if p["pct_remaining"] is not None else "&mdash;"
        if p["daily_rate"] > 0:
            sigma = p.get("rate_sigma", 0)
            rate = (f'{p["daily_rate"]:.2f}&plusmn;{sigma:.2f}/d' if sigma > 0
                    else f'{p["daily_rate"]:.2f}/d')
        else:
            rate = "&mdash;"
        days = f'{p["est_days_remaining"]:.0f}d' if p["est_days_remaining"] is not None else "&mdash;"
        early = p.get("sellout_early")
        late  = p.get("sellout_late")
        if early and late and early != late:
            sellout = f'{early} &ndash; {late}'
        elif early:
            sellout = early
        elif p["est_sellout_date"]:
            sellout = p["est_sellout_date"]
        else:
            sellout = "&mdash;"
        conf = f'<span class="confidence">{p["confidence"]}</span>'
        rows.append(
            f'<tr class="{cls}"><td>{p["name"]}</td><td>{p["stock"]}</td>'
            f'<td>{pct}</td><td>{rate}</td><td>{days}</td><td>{sellout}</td>'
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

    if orders:
        print(f"Computing sales rates from {len(orders)} orders...")
        rate_lookup, global_rates, global_sigmas = compute_sales_rates(orders)
        print(f"  {len(rate_lookup)} product×machine rates")
    else:
        # Fallback to precomputed rates
        rates_path = "docs/sales_rates.json"
        if os.path.exists(rates_path):
            print(f"  Order export timed out — falling back to {rates_path}")
            rate_lookup, global_rates, global_sigmas = load_sales_rates_json(rates_path)
            print(f"  Loaded {len(rate_lookup)} rates from file")
        else:
            print("  ERROR: No order data and no fallback sales_rates.json")
            sys.exit(1)

    # Build report
    machines = build_report_data(inventory, rate_lookup, global_rates, global_sigmas)

    if "--email" in sys.argv:
        html = format_html_report(machines)
        send_email(html)
    else:
        print()
        print(format_text_report(machines))


if __name__ == "__main__":
    main()
