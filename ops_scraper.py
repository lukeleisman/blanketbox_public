"""
ops_scraper.py — fetch live product/price data from Sandstar OPS API
and write docs/products.json for the blanketboxvending.com/prices page.

Credentials come from environment variables (set as GitHub Actions secrets
for automated runs, or export locally for manual runs):
  SANDSTAR_USERNAME      e.g. info@blanketboxvending.com
  SANDSTAR_PASSWORD_HASH the pre-hashed password sent to the API

--- Updating locations ---
When a machine moves or is added, update LOCATION_GROUPS below:
  key   = stable URL slug used in QR codes (never change once QRs are printed)
  value = list of freezer IDs that share that location page

Machine names and addresses are pulled from the API automatically —
you never need to hardcode them here.
"""

import os, sys, json, requests
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

# Display names for each slug (override the auto-generated one if desired)
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


def build_location(token: str, slug: str, freezer_ids: list) -> dict:
    # Fetch machine details for address and display name
    details = []
    for fid in freezer_ids:
        try:
            details.append(fetch_freezer_detail(token, fid))
        except Exception as e:
            print(f"  WARNING: could not fetch detail for {fid}: {e}", file=sys.stderr)

    # Use address from first machine that has one
    address = next((d["address"] for d in details if d.get("address")), "")
    display_name = LOCATION_DISPLAY_NAMES.get(slug) or (details[0]["freezerName"] if details else slug)

    # Fetch and merge products across all machines at this location
    merged = {}  # pid → product dict
    for fid in freezer_ids:
        try:
            items = fetch_machine_products(token, fid)
        except Exception as e:
            print(f"  WARNING: could not fetch products for {fid}: {e}", file=sys.stderr)
            continue
        for item in items:
            if item.get("isSale") != 1:
                continue
            pid = item.get("skuid") or item.get("barcode")
            if not pid:
                continue
            if pid not in merged:
                merged[pid] = {
                    "id": pid,
                    "name": item["goodsName"],
                    "price": item["price"],
                    "image": IMAGE_BASE + item["picture"] if item.get("picture") else None,
                    "stock": 0,
                }
            merged[pid]["stock"] += item.get("stockRealtime") or 0

    products = sorted(merged.values(), key=lambda p: p["name"].lower())
    for p in products:
        p["inStock"] = p["stock"] > 0

    in_stock = sum(1 for p in products if p["inStock"])
    return {
        "id": slug,
        "name": display_name,
        "address": address,
        "productCount": len(products),
        "inStockCount": in_stock,
        "products": products,
    }


def main():
    print("Logging in to Sandstar OPS…")
    token = login()
    print("Login OK.")

    output = {
        "updated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "locations": [],
    }

    for slug, freezer_ids in LOCATION_GROUPS.items():
        print(f"Fetching {slug} (freezers: {', '.join(freezer_ids)})…")
        loc_data = build_location(token, slug, freezer_ids)
        print(f"  → {loc_data['name']}  |  {loc_data['address']}")
        print(f"     {loc_data['productCount']} products, {loc_data['inStockCount']} in stock")
        output["locations"].append(loc_data)

    os.makedirs("docs", exist_ok=True)
    tmp_path = "docs/products.json.tmp"
    out_path = "docs/products.json"
    with open(tmp_path, "w") as f:
        json.dump(output, f, separators=(",", ":"))
    os.replace(tmp_path, out_path)

    total = sum(l["productCount"] for l in output["locations"])
    print(f"\n✓ Wrote {total} products across {len(output['locations'])} locations → {out_path}")


if __name__ == "__main__":
    main()
