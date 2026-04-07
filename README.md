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
