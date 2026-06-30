#!/usr/bin/env python3
"""
mhub_sales_report.py — Daily/weekly sales email for mHUB Prototyping Shop.

Runs daily at 8 AM. On the weekly summary day (default: Monday) it always sends
a 7-day summary. On other days it checks the previous day's sales and sends an
alert if item count exceeds DAILY_ALERT_THRESHOLD.

Usage:
    python mhub_sales_report.py                   # dry-run (print summary, no email)
    python mhub_sales_report.py --email           # auto mode + send email
    python mhub_sales_report.py --mode weekly --email
    python mhub_sales_report.py --mode daily  --email
    python mhub_sales_report.py --days 30         # override: look back N days, always send

Env vars:
    SANDSTAR_USERNAME / SANDSTAR_PASSWORD_HASH  (API access)
    GMAIL_ADDRESS / GMAIL_APP_PASSWORD          (email sending)
    MHUB_REPORT_EMAIL_TO                        (recipient, default luke@blanketboxvending.com)
"""

import argparse
import io
import os
import re
import smtplib
import sys
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests

from ops_scraper import login, fetch_machine_products, BASE_HEADERS

# ============================================================
# CONFIGURATION
# ============================================================

MHUB_FREEZER_ID = "15520"
MHUB_MACHINE_KEY = "mHUBPrototypingShop"  # how it appears in order exports

SALES_URL    = "https://webapi-us.sandstar.com/order/v2/findSaleInfo"
MESSAGE_URL  = "https://webapi-us.sandstar.com/message/findWebUserMessage"
DOWNLOAD_URL = "https://webapi-us.sandstar.com/homePage/download"
ORGAN_SN     = "000332"

POLL_INTERVAL = 3
POLL_TIMEOUT  = 60

TEST_MACHINE_SERIALS = {"2405131108007687", "2405131108006876"}

# Revenue column: "Sales volume" = per-product-line total (Quantity × unit price).
# "Order amount" is the per-order deduplicated total — not usable at the product level.
AMOUNT_COLUMN_CANDIDATES = ["Sales volume", "Total amount", "Actual payment", "Amount", "Subtotal"]

# Daily alert fires when item count for the previous day exceeds this threshold.
DAILY_ALERT_THRESHOLD = 10

# Weekday for the weekly summary email (0=Monday … 6=Sunday).
WEEKLY_SUMMARY_WEEKDAY = 0

# ── Mock data (used with --mock flag) ────────────────────────
# Each row: (day_offset, hour, minute, order_suffix, product_name, qty, unit_price)
# day_offset is relative to start_date; for a 1-day window all offsets clamp to 0,
# which puts 13 items on a single day → triggers the daily alert threshold.
_MOCK_ORDERS_TEMPLATE = [
    (0, 9, 15, "001", "robot remote kit",                           2, 20.00),
    (0, 9, 15, "001", "gorilla super glue 2 tubes",                 1,  6.60),
    (0, 10, 5, "002", "Loctite Threadlocker Red 271 0.2oz",         4,  1.65),
    (0, 10, 5, "002", "Robot Brown Bag",                            2,  5.27),
    (1, 11, 0, "003", "loctite epoxy clear",                        1,  7.69),
    (2, 9, 45, "004", "OLFA heavy duty 18mm snap off knife boxcutter", 1, 16.47),
    (2, 9, 45, "004", "sandable primer",                            1, 16.47),
    (3, 14, 20, "005", "orignal gorilla glue 2fl",                  1, 12.08),
    (4, 8,  5, "006", "GB electrical tape 5 pack",                  2,  6.60),
    (5, 10, 10, "007", "Dremel Kit",                                1, 24.00),
    (5, 10, 10, "007", "locktite vynil fabric and plastic",         2,  9.87),
    (6, 15, 30, "008", "mix measure bucket",                        1,  1.00),
    (6, 15, 30, "008", "3M pro grade sandpaper",                    1, 12.00),
]

_MOCK_INVENTORY = [
    {"name": "robot remote kit",                           "stock":  6, "price": 20.00},
    {"name": "gorilla super glue 2 tubes",                 "stock":  0, "price":  6.60},
    {"name": "Loctite Threadlocker Red 271 0.2oz",         "stock":  2, "price":  1.65},
    {"name": "Robot Brown Bag",                            "stock": 14, "price":  5.27},
    {"name": "loctite epoxy clear",                        "stock":  3, "price":  7.69},
    {"name": "OLFA heavy duty 18mm snap off knife boxcutter","stock": 8, "price": 16.47},
    {"name": "sandable primer",                            "stock":  5, "price": 16.47},
    {"name": "orignal gorilla glue 2fl",                   "stock":  1, "price": 12.08},
    {"name": "GB electrical tape 5 pack",                  "stock": 11, "price":  6.60},
    {"name": "Dremel Kit",                                 "stock":  0, "price": 24.00},
    {"name": "locktite vynil fabric and plastic",          "stock":  4, "price":  9.87},
    {"name": "mix measure bucket",                         "stock": 20, "price":  1.00},
    {"name": "3M pro grade sandpaper",                     "stock":  7, "price": 12.00},
]


# ============================================================
# MOCK DATA
# ============================================================

def generate_mock_data(start_date: date, end_date: date) -> tuple[list[dict], list[dict]]:
    """
    Return (orders, inventory) using hardcoded fake data.
    day_offset in _MOCK_ORDERS_TEMPLATE is clamped to the range length so a
    1-day window gets all orders on the same day (testing the daily alert path)
    and a 7-day window spreads them across the week.
    """
    num_days = (end_date - start_date).days  # 0 for single day, 6 for 7-day window

    IL_TAX_RATE = 0.0975  # IL 7.25% + Chicago 1.25% + Cook 1.25%

    # First pass: build orders and compute per-order pre-tax totals
    orders = []
    order_pretax: dict[str, float] = {}
    for day_off, hour, minute, suffix, name, qty, price in _MOCK_ORDERS_TEMPLATE:
        actual_day   = min(day_off, num_days)
        order_num    = f"MOCK-{suffix}"
        line_pretax  = round(qty * price, 2)
        order_pretax[order_num] = round(order_pretax.get(order_num, 0) + line_pretax, 2)
        order_dt = datetime.combine(start_date + timedelta(days=actual_day),
                                    datetime.min.time().replace(hour=hour, minute=minute))
        orders.append({
            "product_name": name,
            "order_number": order_num,
            "quantity":     qty,
            "amount":       line_pretax,
            "order_date":   order_dt,
            "_suffix":      suffix,  # temp; used below to fill order-level fields
        })

    # Second pass: attach order-level pre-tax and payment totals (with tax)
    for o in orders:
        pretax = order_pretax[o["order_number"]]
        o["order_amount"]   = pretax
        o["payment_amount"] = round(pretax * (1 + IL_TAX_RATE), 2)
        del o["_suffix"]

    return orders, list(_MOCK_INVENTORY)


# ============================================================
# ORDER FETCHING
# ============================================================

def fetch_mhub_orders(token: str,
                      start_date: date = None, end_date: date = None) -> list[dict] | None:
    """
    Fetch orders from the Sandstar async export, filtered to the mHUB machine.

    Returns a list of dicts: {product_name, order_number, quantity, amount, order_date}
    Returns None if the export times out.
    Returns [] if the export succeeded but found no mHUB orders in range.
    """
    headers = {**BASE_HEADERS, "x-token": token}

    if end_date is None:
        end_date = date.today()
    if start_date is None:
        start_date = end_date - timedelta(days=6)

    start_str = f"{start_date} 00:00:00"
    end_str   = f"{end_date} 23:59:59"

    # Snapshot the most recent message sendTime before triggering so we can
    # ignore stale cached exports (same approach as restock_report.py).
    max_pre_time = ""
    try:
        sr = requests.post(MESSAGE_URL, json={"page": 1, "pageSize": 1},
                           headers=headers, timeout=10)
        sr.raise_for_status()
        pre_msgs = (sr.json().get("data") or {}).get("resultList") or []
        if pre_msgs:
            max_pre_time = pre_msgs[0].get("sendTime", "")
            print(f"  Pre-trigger message time: {max_pre_time}", file=sys.stderr)
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
    r = requests.post(SALES_URL + "?exportType=3", json=payload, headers=headers, timeout=30)
    r.raise_for_status()

    # Poll for the export message
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
                mr = requests.post(MESSAGE_URL, json=poll_payload, headers=headers, timeout=10)
                mr.raise_for_status()
                messages = (mr.json().get("data") or {}).get("resultList") or []
                for msg in messages:
                    if msg.get("subject") != "OrderDetailsReport":
                        continue
                    send_time_str = msg.get("sendTime") or ""
                    if send_time_str <= max_pre_time:
                        continue
                    m = re.search(r'fileName=([^\s"\'<>&]+\.xlsx)', msg.get("content", ""))
                    if m:
                        filename = m.group(1)
                        print(f"  Export ready: {filename}", file=sys.stderr)
                        break
            except Exception as e:
                print(f"  Warning: poll failed ({e})", file=sys.stderr)
        if filename:
            break
        remaining = int(deadline - time.time())
        if remaining > 0:
            print(f"  Waiting for export... ({remaining}s left)", file=sys.stderr)
            time.sleep(POLL_INTERVAL)

    if not filename:
        print("  Timed out waiting for export.", file=sys.stderr)
        return None

    print("  Downloading order data...", file=sys.stderr)
    r = requests.get(DOWNLOAD_URL, params={"fileName": filename},
                     headers=headers, timeout=120, stream=True)
    r.raise_for_status()
    xlsx_bytes = io.BytesIO(r.content)

    return _parse_mhub_orders(xlsx_bytes, start_date, end_date)


def _parse_mhub_orders(xlsx_bytes: io.BytesIO,
                       start_date: date, end_date: date) -> list[dict]:
    """Parse the XLSX export and return mHUB orders within the date range."""
    try:
        import pandas as pd
        df = pd.read_excel(xlsx_bytes, header=2, engine="openpyxl")
        print(f"  XLSX rows: {len(df)}", file=sys.stderr)
        print(f"  Columns: {list(df.columns)}", file=sys.stderr)
        return _parse_from_dataframe(df, start_date, end_date)
    except ImportError:
        pass

    import openpyxl
    wb = openpyxl.load_workbook(xlsx_bytes, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return _parse_from_rows(rows, start_date, end_date)


def _parse_from_dataframe(df, start_date: date, end_date: date) -> list[dict]:
    import pandas as pd
    orders = []

    amount_col = None
    for candidate in AMOUNT_COLUMN_CANDIDATES:
        if candidate in df.columns:
            amount_col = candidate
            print(f"  Using amount column: '{amount_col}'", file=sys.stderr)
            break
    if amount_col is None:
        print("  Warning: no amount column found; revenue will be $0.00", file=sys.stderr)

    for _, row in df.iterrows():
        machine_raw = str(row.get("Machine", "")).strip()
        if machine_raw != MHUB_MACHINE_KEY:
            continue
        # Mirror FullMonthlyReport.py clean_sales_df: drop offline channel and zero-amount rows
        if str(row.get("Payment channel", "")).strip() == "offline":
            continue
        order_amount = row.get("Order amount")
        if order_amount is None or order_amount == 0:
            continue

        order_time = pd.to_datetime(row.get("Order time"), errors="coerce")
        if pd.isna(order_time):
            continue
        d = order_time.date()
        if d < start_date or d > end_date:
            continue

        qty = int(row.get("Quantity", 0) or 0)
        if qty <= 0:
            continue

        amount = 0.0
        if amount_col:
            try:
                amount = float(row.get(amount_col) or 0)
            except (TypeError, ValueError):
                pass

        orders.append({
            "product_name":   str(row.get("Product name", "")).strip(),
            "order_number":   str(row.get("Order number", "")).strip(),
            "quantity":       qty,
            "amount":         amount,   # Sales volume: pre-tax per product line
            "order_amount":   float(row.get("Order amount") or 0),    # pre-tax per order
            "payment_amount": float(row.get("Payment Amount") or 0),  # incl. tax per order
            "order_date":     order_time.to_pydatetime(),
        })

    return orders


def _parse_from_rows(rows: list, start_date: date, end_date: date) -> list[dict]:
    if len(rows) < 3:
        return []

    col_headers = [str(c).strip() if c else "" for c in rows[2]]
    col_map = {name: i for i, name in enumerate(col_headers)}
    print(f"  openpyxl columns: {col_headers}", file=sys.stderr)

    amount_col_idx = None
    for candidate in AMOUNT_COLUMN_CANDIDATES:
        if candidate in col_map:
            amount_col_idx = col_map[candidate]
            print(f"  Using amount column: '{candidate}' (idx {amount_col_idx})", file=sys.stderr)
            break
    if amount_col_idx is None:
        print("  Warning: no amount column found; revenue will be $0.00", file=sys.stderr)

    order_amount_idx = col_map.get("Order amount")
    payment_chan_idx = col_map.get("Payment channel")
    order_num_idx    = col_map.get("Order number", 7)

    orders = []
    for row in rows[3:]:
        machine_raw = str(row[col_map.get("Machine", 3)] or "").strip()
        if machine_raw != MHUB_MACHINE_KEY:
            continue
        if machine_raw in TEST_MACHINE_SERIALS:
            continue
        # Mirror FullMonthlyReport.py clean_sales_df: drop offline channel and zero-amount rows
        if payment_chan_idx is not None:
            if str(row[payment_chan_idx] or "").strip() == "offline":
                continue
        if order_amount_idx is not None:
            order_amt = row[order_amount_idx]
            if order_amt is None or order_amt == 0:
                continue

        order_time_raw = row[col_map.get("Order time", 8)]
        if isinstance(order_time_raw, datetime):
            order_dt = order_time_raw
        else:
            try:
                order_dt = datetime.strptime(str(order_time_raw), "%Y-%m-%d %H:%M:%S")
            except (ValueError, TypeError):
                continue
        d = order_dt.date()
        if d < start_date or d > end_date:
            continue

        qty = int(row[col_map.get("Quantity", 17)] or 0)
        if qty <= 0:
            continue

        amount = 0.0
        if amount_col_idx is not None:
            try:
                amount = float(row[amount_col_idx] or 0)
            except (TypeError, ValueError):
                pass

        orders.append({
            "product_name":   str(row[col_map.get("Product name", 16)] or "").strip(),
            "order_number":   str(row[order_num_idx] or "").strip(),
            "quantity":       qty,
            "amount":         amount,   # Sales volume: pre-tax per product line
            "order_amount":   float(row[col_map.get("Order amount",   9)] or 0),
            "payment_amount": float(row[col_map.get("Payment Amount", 10)] or 0),
            "order_date":     order_dt,
        })

    return orders


# ============================================================
# INVENTORY FETCHING
# ============================================================

def fetch_mhub_inventory(token: str) -> list[dict]:
    """Fetch live inventory for the mHUB freezer."""
    try:
        items = fetch_machine_products(token, MHUB_FREEZER_ID)
        products = []
        for item in items:
            if item.get("isSale") != 1:
                continue
            products.append({
                "name":  item["goodsName"],
                "stock": item.get("stockRealtime") or 0,
                "price": item.get("price") or 0.0,
            })
        return sorted(products, key=lambda p: p["name"])
    except Exception as e:
        print(f"  Warning: inventory fetch failed: {e}", file=sys.stderr)
        return []


# ============================================================
# AGGREGATION
# ============================================================

def aggregate_sales(orders: list[dict]) -> list[dict]:
    """Sum units and revenue per product, sorted by revenue descending."""
    by_product: dict[str, dict] = defaultdict(lambda: {"qty": 0, "revenue": 0.0})
    for o in orders:
        by_product[o["product_name"]]["qty"]     += o["quantity"]
        by_product[o["product_name"]]["revenue"] += o["amount"]
    return sorted(
        [{"name": name, **vals} for name, vals in by_product.items()],
        key=lambda x: (-x["revenue"], -x["qty"]),
    )


# ============================================================
# TAX / ORDER TOTALS
# ============================================================

def compute_order_totals(orders: list[dict]) -> dict:
    """
    Deduplicate on order_number and sum order_amount / payment_amount.
    Returns {pretax, tax, total}. Tax is 0 if payment_amount data is absent.
    """
    seen: set[str] = set()
    pretax = 0.0
    paid   = 0.0
    for o in orders:
        if o["order_number"] not in seen:
            seen.add(o["order_number"])
            pretax += o.get("order_amount", 0)
            paid   += o.get("payment_amount", 0)
    tax = round(paid - pretax, 2)
    return {"pretax": round(pretax, 2), "tax": tax, "total": round(paid, 2)}


# ============================================================
# HTML EMAIL
# ============================================================

def format_html_email(
    sales: list[dict],
    orders: list[dict],
    inventory: list[dict],
    start_date: date,
    end_date: date,
    has_revenue: bool = True,
    email_type: str = "weekly",
) -> str:
    single_day  = (start_date == end_date)
    date_range  = _date_range_str(start_date, end_date)
    banner_sub  = ("mHUB Prototyping Shop · Activity Alert"
                   if email_type == "daily"
                   else "mHUB Prototyping Shop · Weekly Sales Report")

    # ── CSS ──────────────────────────────────────────────────
    css = """
    <style>
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
        color: #1a1a1a;
        max-width: 680px;
        margin: 0 auto;
        padding: 16px;
        background: #f4f6f8;
      }
      .card {
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(0,0,0,.08);
        margin-bottom: 16px;
      }
      .banner { background: #1C3D5A; padding: 20px 24px; }
      .banner-title { color: white; font-size: 20px; font-weight: 700;
                      letter-spacing: 0.5px; margin: 0; }
      .banner-sub   { color: #8eb8d8; font-size: 13px; margin-top: 4px; }
      .body  { padding: 20px 24px; }
      .intro { font-size: 14px; color: #444; margin-bottom: 16px; line-height: 1.5; }
      h2 {
        font-size: 13px; font-weight: 600; color: #1C3D5A;
        margin: 24px 0 8px 0; text-transform: uppercase; letter-spacing: 0.5px;
        border-bottom: 2px solid #e8eef4; padding-bottom: 4px;
      }
      table { border-collapse: collapse; width: 100%; font-size: 13px; }
      th {
        text-align: left; padding: 7px 10px; background: #f0f4fa;
        border-bottom: 2px solid #dde4ed; color: #555;
        font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px;
      }
      th.num { text-align: right; }
      td { padding: 7px 10px; border-bottom: 1px solid #eee; vertical-align: middle; }
      td.num { text-align: right; font-variant-numeric: tabular-nums; }
      tr:last-child td { border-bottom: none; }
      .totals-row td {
        font-weight: 600; border-top: 2px solid #dde4ed;
        border-bottom: none; background: #f9fbfd;
      }
      .oos { color: #c0392b; }
      .low { color: #d97706; }
      .ts  { color: #888; white-space: nowrap; }
      .footer { font-size: 11px; color: #aaa; text-align: center; padding: 12px 0 4px; }
    </style>
    """

    # ── Sales summary table ───────────────────────────────────
    total_units = sum(s["qty"] for s in sales)
    order_tots  = compute_order_totals(orders)
    has_tax     = has_revenue and order_tots["tax"] > 0.005

    rev_hdr = '<th class="num">Pre-tax</th>' if has_revenue else ""

    sales_rows = []
    for s in sales:
        rev_cell = f'<td class="num">${s["revenue"]:.2f}</td>' if has_revenue else ""
        sales_rows.append(
            f'<tr><td>{_esc(s["name"])}</td>'
            f'<td class="num">{s["qty"]}</td>'
            f'{rev_cell}</tr>'
        )

    # Footer: pre-tax total, then tax row, then total-paid row (if tax data present)
    if has_revenue:
        footer_rows = (
            f'<tr class="totals-row">'
            f'<td>Total</td><td class="num">{total_units}</td>'
            f'<td class="num">${order_tots["pretax"]:.2f}</td>'
            f'</tr>'
        )
        if has_tax:
            footer_rows += (
                f'<tr><td colspan="2" style="text-align:right; color:#888; font-size:12px; '
                f'padding:4px 10px;">Sales tax</td>'
                f'<td class="num" style="color:#888; font-size:12px;">'
                f'+ ${order_tots["tax"]:.2f}</td></tr>'
                f'<tr><td colspan="2" style="text-align:right; font-weight:600; '
                f'padding:6px 10px;">Total collected</td>'
                f'<td class="num" style="font-weight:600;">${order_tots["total"]:.2f}</td></tr>'
            )
    else:
        footer_rows = (
            f'<tr class="totals-row">'
            f'<td>Total</td><td class="num">{total_units}</td>'
            f'</tr>'
        )

    sales_table = f"""
    <table>
      <thead><tr><th>Product</th><th class="num">Units&nbsp;Sold</th>{rev_hdr}</tr></thead>
      <tbody>{"".join(sales_rows)}</tbody>
      <tfoot>{footer_rows}</tfoot>
    </table>"""

    # ── Transaction list ──────────────────────────────────────
    txn_table = _format_transactions(orders, has_revenue, single_day)

    # ── Inventory ─────────────────────────────────────────────
    # Suppress when all stock values are implausibly high (mHUB initializes at 9999
    # and just decrements — values like 9927 still mean "not meaningfully tracked").
    if inventory and all(p["stock"] > 100 for p in inventory):
        inventory = []

    inv_section = ""
    if inventory:
        inv_rows = []
        for p in inventory:
            stock = p["stock"]
            cls   = 'class="oos"' if stock == 0 else ('class="low"' if stock <= 3 else "")
            inv_rows.append(
                f'<tr><td>{_esc(p["name"])}</td>'
                f'<td class="num" {cls}>{stock}</td>'
                f'<td class="num">${p["price"]:.2f}</td></tr>'
            )
        inv_section = f"""
        <h2>Current Inventory</h2>
        <p class="intro">
          Here is the current inventory in the machine, if it was set by the stockers.
        </p>
        <table>
          <thead><tr><th>Product</th><th class="num">In&nbsp;Stock</th>
          <th class="num">Price</th></tr></thead>
          <tbody>{"".join(inv_rows)}</tbody>
        </table>"""

    # ── Assemble ──────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8">{css}</head>
<body>
  <div class="card">
    <div class="banner">
      <div class="banner-title">BLANKET BOX</div>
      <div class="banner-sub">{banner_sub}</div>
    </div>
    <div class="body">
      <p class="intro">
        Here is the Blanket Box sales report for {date_range}.
      </p>
      <h2>Sales Summary</h2>
      {sales_table}
      <h2>Transactions</h2>
      {txn_table}
      {inv_section}
    </div>
  </div>
  <div class="footer">Blanket Box Vending &middot; blanketboxvending.com</div>
</body>
</html>"""


def _format_transactions(orders: list[dict], has_revenue: bool, single_day: bool) -> str:
    """
    Transaction list sorted by time, shaded by order group.
    When tax data is present, a compact subtotal row is appended after each order.
    """
    sorted_orders = sorted(orders, key=lambda o: o["order_date"])

    # Determine whether any order actually has tax (payment > order amount)
    show_tax = has_revenue and any(
        o.get("payment_amount", 0) > o.get("order_amount", 0) + 0.005
        for o in sorted_orders
    )

    rev_hdr  = '<th class="num">Pre-tax</th>' if has_revenue else ""
    time_hdr = "Time" if single_day else "Date &amp; Time"
    n_cols   = 3 + (1 if has_revenue else 0)  # time + product + qty [+ pre-tax]

    BG = ("#ffffff", "#f4f7fb")
    current_order = None
    shade_idx = 0
    # Accumulate lines per order so we can emit a subtotal after each group
    pending_rows: list[str] = []
    current_order_amount   = 0.0
    current_payment_amount = 0.0
    rows: list[str] = []

    def _flush_order():
        """Emit buffered product rows + subtotal for the completed order."""
        rows.extend(pending_rows)
        if show_tax and current_order is not None:
            tax   = round(current_payment_amount - current_order_amount, 2)
            total = current_payment_amount
            subtotal_style = (
                "font-size:11px; color:#666; padding:3px 10px; "
                "border-bottom:2px solid #dde4ed;"
            )
            rows.append(
                f'<tr>'
                f'<td colspan="{n_cols - 1}" style="text-align:right; {subtotal_style}">'
                f'Order subtotal: <strong>${current_order_amount:.2f}</strong> pre-tax'
                f' + ${tax:.2f} tax</td>'
                f'<td class="num" style="{subtotal_style}">'
                f'<strong>${total:.2f}</strong></td>'
                f'</tr>'
            )
        pending_rows.clear()

    for o in sorted_orders:
        if o["order_number"] != current_order:
            if current_order is not None:
                _flush_order()
            current_order          = o["order_number"]
            current_order_amount   = o.get("order_amount", 0)
            current_payment_amount = o.get("payment_amount", 0)
            shade_idx = 1 - shade_idx

        bg = BG[shade_idx]
        dt = o["order_date"]
        time_part = dt.strftime("%I:%M %p").lstrip("0")
        ts = time_part if single_day else f"{dt:%a %b} {dt.day} &middot; {time_part}"

        rev_cell = f'<td class="num" style="background:{bg}">${o["amount"]:.2f}</td>' if has_revenue else ""
        pending_rows.append(
            f'<tr style="background:{bg}">'
            f'<td class="ts" style="background:{bg}">{ts}</td>'
            f'<td style="background:{bg}">{_esc(o["product_name"])}</td>'
            f'<td class="num" style="background:{bg}">{o["quantity"]}</td>'
            f'{rev_cell}'
            f'</tr>'
        )

    _flush_order()  # emit the last order

    return f"""
    <table>
      <thead><tr><th>{time_hdr}</th><th>Product</th>
      <th class="num">Qty</th>{rev_hdr}</tr></thead>
      <tbody>{"".join(rows)}</tbody>
    </table>"""


def _date_range_str(start: date, end: date) -> str:
    if start == end:
        return start.strftime("%B %d, %Y")
    if start.year == end.year and start.month == end.month:
        return f"{start.strftime('%B %d')}–{end.strftime('%d, %Y')}"
    if start.year == end.year:
        return f"{start.strftime('%B %d')}–{end.strftime('%B %d, %Y')}"
    return f"{start.strftime('%B %d, %Y')}–{end.strftime('%B %d, %Y')}"


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ============================================================
# EMAIL SENDING
# ============================================================

def send_email(html_body: str, subject: str):
    sender     = os.environ.get("GMAIL_ADDRESS", "luke.blanketboxvending@gmail.com")
    password   = os.environ["GMAIL_APP_PASSWORD"]
    to_env     = os.environ.get("MHUB_REPORT_EMAIL_TO", "luke@blanketboxvending.com")
    recipients = [e.strip() for e in to_env.split(",") if e.strip()]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, recipients, msg.as_string())

    print(f"Email sent to {', '.join(recipients)}")


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="mHUB sales report")
    parser.add_argument("--email", action="store_true", help="send HTML email")
    parser.add_argument(
        "--mode", choices=["auto", "weekly", "daily"], default="auto",
        help=(
            "auto: weekly on WEEKLY_SUMMARY_WEEKDAY, daily alert on other days; "
            "weekly: always send 7-day summary; "
            "daily: always send yesterday's report (ignores threshold)"
        ),
    )
    parser.add_argument(
        "--days", type=int, default=None,
        help="override lookback window in days (always sends, ignores threshold)",
    )
    parser.add_argument(
        "--mock", action="store_true",
        help="use fake data instead of live API (no credentials needed, always sends)",
    )
    args = parser.parse_args()

    today     = date.today()
    yesterday = today - timedelta(days=1)

    # ── Determine date range and email type ───────────────────
    if args.days is not None:
        # Manual override for testing
        start_date = today - timedelta(days=args.days - 1)
        end_date   = today
        email_type = "weekly" if args.days > 1 else "daily"
        apply_threshold = False
    elif args.mode == "weekly" or (args.mode == "auto" and today.weekday() == WEEKLY_SUMMARY_WEEKDAY):
        start_date = yesterday - timedelta(days=6)
        end_date   = yesterday
        email_type = "weekly"
        apply_threshold = False
    else:
        # Daily check: only alert if above threshold
        start_date = yesterday
        end_date   = yesterday
        email_type = "daily"
        apply_threshold = True

    date_range = _date_range_str(start_date, end_date)
    print(f"mHUB sales report [{email_type}{'  MOCK' if args.mock else ''}]: {start_date} to {end_date}")

    # ── Fetch or generate orders ──────────────────────────────
    if args.mock:
        print("Using mock data (no API calls).")
        orders, inventory = generate_mock_data(start_date, end_date)
    else:
        print("Logging in (orders)...")
        token = login()
        print("Fetching order export...")
        orders = fetch_mhub_orders(token, start_date=start_date, end_date=end_date)

        if orders is None:
            print("Order export timed out.", file=sys.stderr)
            sys.exit(1)

        if not orders:
            print(f"No sales from {start_date} to {end_date}. Nothing to report.")
            sys.exit(0)

    total_items = sum(o["quantity"] for o in orders)
    print(f"  {len(orders)} order lines · {total_items} items total")

    if not args.mock and apply_threshold:
        if total_items <= DAILY_ALERT_THRESHOLD:
            print(
                f"  {total_items} items sold yesterday — at or below threshold "
                f"({DAILY_ALERT_THRESHOLD}). No alert sent."
            )
            sys.exit(0)
        print(f"  {total_items} items sold yesterday — above threshold. Sending alert.")

    has_revenue = any(o["amount"] > 0 for o in orders)
    if not has_revenue:
        print("  Note: no revenue data in export (amount column absent)", file=sys.stderr)

    # ── Fetch inventory (skipped in mock mode) ────────────────
    if not args.mock:
        print("Logging in (inventory)...")
        token = login()
        print("Fetching inventory...")
        inventory = fetch_mhub_inventory(token)
        print(f"  {len(inventory)} products in machine")

    # ── Aggregate & print summary ─────────────────────────────
    sales         = aggregate_sales(orders)
    total_revenue = sum(s["revenue"] for s in sales)

    print(f"\nmHUB — {date_range}")
    print(f"  {len(sales)} products · {total_items} units · ${total_revenue:.2f}")
    for s in sales:
        rev = f"  ${s['revenue']:.2f}" if has_revenue else ""
        print(f"    {s['name']}: {s['qty']} sold{rev}")

    meaningful_inv = [p for p in inventory if p["stock"] <= 100]
    if meaningful_inv:
        print(f"\nInventory ({len(inventory)} products):")
        for p in inventory:
            print(f"    {p['name']}: {p['stock']} in stock @ ${p['price']:.2f}")
    else:
        print("\nInventory: all values implausibly high (9999-mode) — omitted from report")

    print(f"\nTransactions ({len(orders)} lines):")
    for o in sorted(orders, key=lambda x: x["order_date"]):
        dt  = o["order_date"]
        rev = f"  ${o['amount']:.2f}" if has_revenue else ""
        print(f"    {dt:%Y-%m-%d %H:%M}  {o['product_name']}  qty={o['quantity']}{rev}")

    # ── Send email ────────────────────────────────────────────
    if args.email:
        html = format_html_email(
            sales, orders, inventory, start_date, end_date,
            has_revenue=has_revenue, email_type=email_type,
        )
        if email_type == "weekly":
            subject = f"mHUB Sales Report — {date_range}"
        else:
            subject = f"High activity day at Blanket Box — {yesterday.strftime('%B %d, %Y')}"
        send_email(html, subject)


if __name__ == "__main__":
    main()
