#!/usr/bin/env python3
"""
verify_sales_data.py — Cross-check order export XLSXs against inventory_history.csv.

Two sources:
  XLSX  — individual transaction records from Sandstar (ground truth)
  History — 30-min stock snapshots; negative deltas ≈ sales

Known limitation: history undercounts when a product sells out between scrapes.
If stock hits 0 partway through a 30-min window, sales after that point are
invisible to the scraper (delta stays at 0).

Usage:
    # Compare Apr + May XLSXs against history (auto date range from history coverage):
    python verify_sales_data.py

    # Explicit date range:
    python verify_sales_data.py --start 2026-04-07 --end 2026-04-30

    # Custom XLSX paths:
    python verify_sales_data.py --xlsx path/to/file.xlsx path/to/other.xlsx
"""

import argparse
import csv
import io
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ── Machine name normalisation (same as restock_report.py) ──────────────────

MACHINE_NAME_MAP: dict[str, str | None] = {
    "333 North Broad Ambient": "333 North Broad Ambient",
    "333 North Broad Cooler":  "333 North Broad Cooler",
    "Philly TrueCooler II":    "Philly TrueCooler II",
    "PhillyTrueCooler":        "Philly TrueCooler II",
    "TheHaven#2-cooler":       "TheHaven#2-cooler",
    "ChicagoDryCab2-Warren":   "ChicagoDryCab2-Warren",
    "ChicagoDryCab1":          "ChicagoDryCab1",
    "BorellisBox":             "BorellisBox",
    "TheHaven#1-ambient":      "333 North Broad Ambient",
    "ChicagoDryCab2":          "ChicagoDryCab2-Warren",
    "SpringGardenDryCab":      None,
}

TEST_MACHINE_SERIALS = {"2405131108007687", "2405131108006876"}

DEFAULT_HISTORY_PATH = Path(__file__).parent / "docs/inventory_history.csv"
DEFAULT_ORDERDATA_DIR = Path.home() / "Luke/business/ComfortVending/orderdata"


# ── Source parsers ───────────────────────────────────────────────────────────

def parse_history_sales(
    path: Path | str,
    start_date: date | None = None,
    end_date:   date | None = None,
) -> dict[tuple[str, str], int]:
    """
    Sum negative stock deltas in inventory_history.csv → units sold per (machine, product).

    Timestamps in the CSV are UTC ISO-8601. Date comparisons use UTC date.
    """
    sales: dict[tuple[str, str], int] = {}
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            delta = int(row["delta"])
            if delta >= 0:
                continue
            machine = MACHINE_NAME_MAP.get(row["machine_name"], row["machine_name"])
            if machine is None:
                continue
            ts = datetime.fromisoformat(row["timestamp"].replace("Z", "+00:00"))
            d = ts.date()
            if start_date and d < start_date:
                continue
            if end_date and d > end_date:
                continue
            key = (machine, row["name"])
            sales[key] = sales.get(key, 0) + abs(delta)
    return sales


def _agg_from_df(
    df: pd.DataFrame,
    start_date: date | None,
    end_date:   date | None,
) -> dict[tuple[str, str], int]:
    """
    Aggregate a parsed XLSX DataFrame → units sold per (machine, product).

    Accepts DataFrames produced by pd.read_excel(..., header=2).
    """
    sales: dict[tuple[str, str], int] = {}
    for _, row in df.iterrows():
        if str(row.get("Payment status", "")).strip() != "Success":
            continue
        machine_raw = str(row.get("Machine", "")).strip()
        if machine_raw in TEST_MACHINE_SERIALS or re.match(r"^\d+$", machine_raw):
            continue
        machine = MACHINE_NAME_MAP.get(machine_raw)
        if not machine:
            continue
        order_time = pd.to_datetime(row.get("Order time"), errors="coerce")
        if pd.isna(order_time):
            continue
        d = order_time.date()
        if start_date and d < start_date:
            continue
        if end_date and d > end_date:
            continue
        product = str(row.get("Product name", "")).strip()
        qty = int(row.get("Quantity", 0))
        key = (machine, product)
        sales[key] = sales.get(key, 0) + qty
    return sales


def parse_xlsx_sales(
    paths: list[Path | str],
    start_date: date | None = None,
    end_date:   date | None = None,
) -> dict[tuple[str, str], int]:
    """Load one or more OrderDetailsReport XLSXs and aggregate sales."""
    combined: dict[tuple[str, str], int] = {}
    for p in paths:
        df = pd.read_excel(p, header=2, engine="openpyxl")
        for (machine, product), qty in _agg_from_df(df, start_date, end_date).items():
            combined[(machine, product)] = combined.get((machine, product), 0) + qty
    return combined


# ── Comparison output ────────────────────────────────────────────────────────

def compare_sources(
    xlsx: dict[tuple[str, str], int],
    hist: dict[tuple[str, str], int],
    start_date: date | None,
    end_date:   date | None,
) -> None:
    all_keys = sorted(xlsx.keys() | hist.keys())
    date_range = (
        f"{start_date} → {end_date}" if start_date or end_date else "all dates"
    )
    print(f"\nSales comparison: XLSX vs inventory_history  [{date_range}]")
    print(f"{'Machine':<32} {'Product':<40} {'XLSX':>6} {'Hist':>6} {'Diff':>6} {'Note'}")
    print("-" * 110)

    only_xlsx = only_hist = both_match = both_differ = 0
    for machine, product in all_keys:
        x = xlsx.get((machine, product), 0)
        h = hist.get((machine, product), 0)
        diff = x - h
        if x == 0:
            note = "history only"
            only_hist += 1
        elif h == 0:
            note = "XLSX only (OOS gap?)"
            only_xlsx += 1
        elif diff == 0:
            note = ""
            both_match += 1
        else:
            pct = abs(diff) / x * 100
            note = f"{pct:.0f}% off"
            both_differ += 1

        flag = "  " if diff == 0 else ">>"
        print(f"{flag} {machine:<30} {product:<40} {x:>6} {h:>6} {diff:>+6}  {note}")

    print("-" * 110)
    print(f"Matched: {both_match}  |  Differ: {both_differ}  |  "
          f"XLSX only: {only_xlsx}  |  History only: {only_hist}")
    if both_differ > 0:
        over = sum(1 for (m, p) in all_keys
                   if xlsx.get((m, p), 0) > 0 and hist.get((m, p), 0) > 0
                   and xlsx[(m, p)] != hist[(m, p)] and xlsx[(m, p)] > hist[(m, p)])
        under = both_differ - over
        print(f"  History undercounts (OOS gap): {over}  |  overcounts: {under}")
    print()


def history_date_range(path: Path | str) -> tuple[date, date]:
    """Return (min_date, max_date) from inventory_history timestamps."""
    dates = []
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            ts = datetime.fromisoformat(row["timestamp"].replace("Z", "+00:00"))
            dates.append(ts.date())
    return min(dates), max(dates)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", nargs="+", metavar="PATH",
                    help="OrderDetailsReport XLSX file(s) to compare against history. "
                         "Defaults to Apr + May 2026 from the standard orderdata directory.")
    ap.add_argument("--start", metavar="YYYY-MM-DD",
                    help="Start date (inclusive). Defaults to first date in history.")
    ap.add_argument("--end", metavar="YYYY-MM-DD",
                    help="End date (inclusive). Defaults to last date in history.")
    ap.add_argument("--history", default=str(DEFAULT_HISTORY_PATH),
                    metavar="PATH", help="Path to inventory_history.csv.")
    args = ap.parse_args()

    history_path = Path(args.history)
    if not history_path.exists():
        sys.exit(f"History file not found: {history_path}")

    hist_start, hist_end = history_date_range(history_path)
    print(f"History coverage: {hist_start} → {hist_end}")

    start_date = date.fromisoformat(args.start) if args.start else hist_start
    end_date   = date.fromisoformat(args.end)   if args.end   else hist_end

    if args.xlsx:
        xlsx_paths = [Path(p) for p in args.xlsx]
    else:
        # Default: any monthly XLSX that overlaps the date window
        xlsx_paths = []
        for p in sorted(DEFAULT_ORDERDATA_DIR.glob("*_OrderDetailsReport.xlsx")):
            xlsx_paths.append(p)
        if not xlsx_paths:
            sys.exit(f"No OrderDetailsReport XLSXs found in {DEFAULT_ORDERDATA_DIR}")
        print(f"Using XLSXs (filtered to {start_date} → {end_date}):")
        for p in xlsx_paths:
            print(f"  {p.name}")

    missing = [p for p in xlsx_paths if not p.exists()]
    if missing:
        sys.exit("Missing XLSX files:\n" + "\n".join(f"  {p}" for p in missing))

    print("\nParsing XLSX sales...")
    xlsx_sales = parse_xlsx_sales(xlsx_paths, start_date, end_date)
    total_xlsx = sum(xlsx_sales.values())
    print(f"  {len(xlsx_sales)} (machine, product) pairs, {total_xlsx} units total")

    print("Parsing inventory history...")
    hist_sales = parse_history_sales(history_path, start_date, end_date)
    total_hist = sum(hist_sales.values())
    print(f"  {len(hist_sales)} (machine, product) pairs, {total_hist} units total")

    compare_sources(xlsx_sales, hist_sales, start_date, end_date)


if __name__ == "__main__":
    main()
