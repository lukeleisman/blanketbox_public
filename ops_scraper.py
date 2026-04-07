"""
ops_scraper.py — fetch live data from Sandstar OPS API and write:

  docs/products.json   — prices page (blanketboxvending.com/prices)
  docs/inventory.json  — per-machine stock levels (feeds Inventory.ipynb)
  docs/inventory.csv   — same data, flat CSV for easy download/Excel

Credentials via environment variables (GitHub Actions secrets or local export):
  SANDSTAR_USERNAME      e.g. info@blanketboxvending.com
  SANDSTAR_PASSWORD_HASH the pre-hashed password sent to the API

--- Updating locations ---
When a machine moves or is added, update LOCATION_GROUPS below:
  key   = stable URL slug used in QR codes (never change once QRs are printed)
  value = list of freezer IDs at that location

Machine names and addresses are pulled from the API automatically.
"""

import csv, io, os, sys, json, requests
from datetime import datetime, timezone

LOGIN_URL   = "https://webapi-us.sandstar.com/user/login"
GOODS_URL   = "https://webapi-us.sandstar.com/goods/v2/getGoodsAtShelvesByFreezerIdV2"
DETAIL_URL  = "https://webapi-us.sandstar.com/freezer/findFreezerDetail/v1"
IMAGE_BASE  = "https://image-us.sandstar.com/images/"

BASE_HEADERS = {
    "Content-Type": "application/json",
    "app-scope": "12",
    "origin": "https://prod-ops-us.sandstar.com",
}

# --- EDIT THIS when machines move or are added ----------------------------
# key   = stable URL slug (used in QR codes — don't rename after printing)
# value = list of freezer IDs at that location
LOCATION_GROUPS = {
    "doylestown": ["15519", "16214"],   # 333 N Broad St, Doylestown PA
    "philly":     ["15521"],            # 4200 Ludlow St, Philadelphia PA
    "haven":      ["16213"],            # 1000 Terrain St, Malvern PA
    "warren":     ["15518"],            # Warren Park Fieldhouse, Chicago
    "mckinley":   ["15520"],            # McKinley Park Fieldhouse, Chicago
    "borellis":   ["15452"],            # 2124 W Lawrence Ave, Chicago
}

LOCATION_DISPLAY_NAMES = {
    "doylestown": "333 North Broad",
    "philly":     "Philadelphia",
    "haven":      "The Haven",
    "warren":     "Warren Park",
    "mckinley":   "McKinley Park",
    "borellis":   "Borelli's Box",
}
# --------------------------------------------------------------------------


def login() -> str:
    username = os.environ["SANDSTAR_USERNAME"]
    pw_hash  = os.environ["SANDSTAR_PASSWORD_HASH"]
    r = requests.post(LOGIN_URL, json={"userSn": username, "password": pw_hash},
                      headers=BASE_HEADERS, timeout=30)
    r.raise_for_status()
    body = r.json()
    if body["status"] != 200:
        raise SystemExit(f"Login failed: {body['message']}")
    return body["data"]["token"]


def fetch_freezer_detail(token: str, freezer_id: str) -> dict:
    headers = {**BASE_HEADERS, "x-token": token}
    r = requests.post(DETAIL_URL, json={"id": freezer_id, "temperatureUnit": 1},
                      headers=headers, timeout=15)
    r.raise_for_status()
    body = r.json()
    if body["status"] != 200:
        raise RuntimeError(f"Detail API error for freezer {freezer_id}: {body['message']}")
    return body["data"]


def fetch_machine_products(token: str, freezer_id: str) -> list:
    headers = {**BASE_HEADERS, "x-token": token}
    r = requests.post(GOODS_URL,
                      json={"freezerId": freezer_id, "page": 1, "pageSize": 10000},
                      headers=headers, timeout=30)
    r.raise_for_status()
    body = r.json()
    if body["status"] != 200:
        raise RuntimeError(f"Goods API error for freezer {freezer_id}: {body['message']}")
    return body["data"]["resultList"]


def scrape_all(token: str) -> list:
    """
    Fetch details + products for every machine.
    Returns a list of machine dicts, each with full per-product inventory rows.
    This is the single source of truth — products.json and inventory files
    are both derived from this.
    """
    machines = []
    for slug, freezer_ids in LOCATION_GROUPS.items():
        display_name = LOCATION_DISPLAY_NAMES.get(slug, slug)
        for fid in freezer_ids:
            print(f"  Fetching machine {fid}…")

            # Machine metadata
            try:
                detail = fetch_freezer_detail(token, fid)
                machine_name = detail.get("freezerName", fid)
                address      = detail.get("address", "")
            except Exception as e:
                print(f"    WARNING: detail fetch failed: {e}", file=sys.stderr)
                machine_name = fid
                address      = ""

            # Products
            try:
                items = fetch_machine_products(token, fid)
            except Exception as e:
                print(f"    WARNING: products fetch failed: {e}", file=sys.stderr)
                items = []

            products = []
            for item in items:
                if item.get("isSale") != 1:
                    continue
                pid = item.get("skuid") or item.get("barcode")
                if not pid:
                    continue
                stock    = item.get("stockRealtime") or 0
                capacity = item.get("capacity") or 0
                products.append({
                    "skuid":         pid,
                    "barcode":       item.get("barcode", ""),
                    "name":          item["goodsName"],
                    "price":         item["price"],
                    "image":         IMAGE_BASE + item["picture"] if item.get("picture") else None,
                    "stock":         stock,
                    "capacity":      capacity,
                    "fill_pct":      round(stock / capacity * 100) if capacity else None,
                    "stock_warning": item.get("stockWarning") or 0,
                    "stock_updated": item.get("stockTime", ""),
                })

            machines.append({
                "freezer_id":   fid,
                "location_id":  slug,
                "location_name": display_name,
                "machine_name": machine_name,
                "address":      address,
                "products":     products,
            })

    return machines


# ── Output builders ────────────────────────────────────────────────────────

def build_prices_json(machines: list, updated: str) -> dict:
    """Merge machines at same location, keep only fields the prices page needs."""
    locations = {}
    for m in machines:
        slug = m["location_id"]
        if slug not in locations:
            locations[slug] = {
                "id":           slug,
                "name":         m["location_name"],
                "address":      m["address"],
                "products":     {},   # skuid → merged product
            }
        for p in m["products"]:
            pid = p["skuid"]
            if pid not in locations[slug]["products"]:
                locations[slug]["products"][pid] = {
                    "id":      pid,
                    "name":    p["name"],
                    "price":   p["price"],
                    "image":   p["image"],
                    "stock":   0,
                }
            locations[slug]["products"][pid]["stock"] += p["stock"]

    out_locations = []
    for loc in locations.values():
        products = sorted(loc["products"].values(), key=lambda p: p["name"].lower())
        for p in products:
            p["inStock"] = p["stock"] > 0
        in_stock = sum(1 for p in products if p["inStock"])
        out_locations.append({
            "id":           loc["id"],
            "name":         loc["name"],
            "address":      loc["address"],
            "productCount": len(products),
            "inStockCount": in_stock,
            "products":     products,
        })

    return {"updated": updated, "locations": out_locations}


def build_inventory_json(machines: list, updated: str) -> dict:
    """Full per-machine inventory — used by Inventory.ipynb."""
    out_machines = []
    for m in machines:
        products = sorted(m["products"], key=lambda p: p["name"].lower())
        total    = len(products)
        in_stock = sum(1 for p in products if p["stock"] > 0)
        low      = sum(1 for p in products
                       if 0 < p["stock"] <= p["stock_warning"] and p["stock_warning"] > 0)
        out_machines.append({
            "freezer_id":    m["freezer_id"],
            "location_id":   m["location_id"],
            "location_name": m["location_name"],
            "machine_name":  m["machine_name"],
            "address":       m["address"],
            "total_skus":    total,
            "in_stock":      in_stock,
            "out_of_stock":  total - in_stock,
            "low_stock":     low,
            "products":      products,
        })
    return {"updated": updated, "machines": out_machines}


def build_inventory_csv(machines: list, updated: str) -> str:
    """Flat CSV — one row per machine × product. Easy to open in Excel."""
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([
        "updated", "location_id", "location_name", "machine_name", "freezer_id",
        "address", "skuid", "barcode", "name", "price",
        "stock", "capacity", "fill_pct", "stock_warning", "stock_updated",
    ])
    for m in machines:
        for p in sorted(m["products"], key=lambda x: x["name"].lower()):
            writer.writerow([
                updated,
                m["location_id"], m["location_name"], m["machine_name"], m["freezer_id"],
                m["address"],
                p["skuid"], p["barcode"], p["name"], p["price"],
                p["stock"], p["capacity"],
                p["fill_pct"] if p["fill_pct"] is not None else "",
                p["stock_warning"], p["stock_updated"],
            ])
    return buf.getvalue()


# ── Helpers ────────────────────────────────────────────────────────────────

def atomic_write(path: str, content: str, binary: bool = False):
    tmp = path + ".tmp"
    mode = "wb" if binary else "w"
    with open(tmp, mode) as f:
        f.write(content if binary else content)
    os.replace(tmp, path)


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    print("Logging in to Sandstar OPS…")
    token = login()
    print("Login OK.\n")

    updated = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    print("Scraping machines…")
    machines = scrape_all(token)
    print()

    os.makedirs("docs", exist_ok=True)

    # 1. prices page
    prices = build_prices_json(machines, updated)
    atomic_write("docs/products.json", json.dumps(prices, separators=(",", ":")))
    total_products = sum(l["productCount"] for l in prices["locations"])
    print(f"✓ docs/products.json  — {total_products} products across {len(prices['locations'])} locations")

    # 2. inventory JSON
    inv = build_inventory_json(machines, updated)
    atomic_write("docs/inventory.json", json.dumps(inv, indent=2))
    total_skus = sum(m["total_skus"] for m in inv["machines"])
    low_total  = sum(m["low_stock"]  for m in inv["machines"])
    print(f"✓ docs/inventory.json — {total_skus} SKU-slots across {len(inv['machines'])} machines"
          f"  ({low_total} low-stock)")

    # 3. inventory CSV
    csv_data = build_inventory_csv(machines, updated)
    atomic_write("docs/inventory.csv", csv_data)
    rows = csv_data.count("\n") - 1
    print(f"✓ docs/inventory.csv  — {rows} rows")

    # Summary by machine
    print()
    for m in inv["machines"]:
        bar = "█" * (m["in_stock"] * 20 // max(m["total_skus"], 1))
        bar = bar.ljust(20, "░")
        print(f"  {m['machine_name']:<35s} {bar}  {m['in_stock']:>3}/{m['total_skus']} in stock"
              + (f"  ⚠ {m['low_stock']} low" if m["low_stock"] else ""))


if __name__ == "__main__":
    main()
