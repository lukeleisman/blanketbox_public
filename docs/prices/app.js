/* Blanket Box Vending — Live Prices App
 * Loaded by WordPress page at blanketboxvending.com/prices
 * Fetches products.json from GitHub Pages, renders a location-selector
 * then a searchable product grid.
 *
 * WordPress embed (Custom HTML block):
 *   <div id="bb-prices-app"></div>
 *   <script src="https://lukeleisman.github.io/blanketbox_public/prices/app.js"></script>
 */
(function () {
  'use strict';

  const DATA_URL = 'https://lukeleisman.github.io/blanketbox_public/products.json';

  const ICONS = {
    doylestown: '🏛️',
    philly:     '🏙️',
    haven:      '🏠',
    warren:     '🌳',
    mckinley:   '🌳',
    borellis:   '📦',
  };

  /* ── CSS (injected into <head> so WordPress styles don't conflict) ────── */
  const CSS = `
#bb-app {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
  background: #f5f5f2;
  min-height: 70vh;
  color: #1a1a1a;
  -webkit-font-smoothing: antialiased;
}
#bb-app *, #bb-app *::before, #bb-app *::after {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

/* Header */
.bb-header {
  background: #1C3D5A;
  color: #fff;
  padding: 14px 18px;
  display: flex;
  align-items: center;
  gap: 10px;
}
.bb-back-btn {
  background: none;
  border: none;
  color: #fff;
  font-size: 26px;
  line-height: 1;
  cursor: pointer;
  padding: 0 6px 0 0;
  opacity: 0.85;
}
.bb-back-btn:hover { opacity: 1; }
.bb-header-title {
  font-size: 17px;
  font-weight: 700;
  letter-spacing: -0.2px;
}
.bb-header-sub {
  font-size: 12px;
  opacity: 0.65;
  margin-top: 2px;
}

/* Location selector */
.bb-selector-wrap {
  padding: 24px 16px 32px;
  max-width: 500px;
  margin: 0 auto;
}
.bb-selector-heading {
  font-size: 21px;
  font-weight: 800;
  color: #1C3D5A;
  margin-bottom: 4px;
}
.bb-selector-sub {
  font-size: 13px;
  color: #888;
  margin-bottom: 20px;
}
.bb-loc-card {
  background: #fff;
  border-radius: 16px;
  padding: 16px 18px;
  margin-bottom: 10px;
  display: flex;
  align-items: center;
  gap: 14px;
  cursor: pointer;
  border: 2px solid transparent;
  box-shadow: 0 2px 10px rgba(0,0,0,0.07);
  transition: border-color .15s, box-shadow .15s, transform .1s;
  user-select: none;
}
.bb-loc-card:hover {
  border-color: #1C3D5A;
  box-shadow: 0 4px 18px rgba(28,61,90,0.13);
  transform: translateY(-1px);
}
.bb-loc-icon { font-size: 28px; flex-shrink: 0; }
.bb-loc-info { flex: 1; }
.bb-loc-name { font-size: 16px; font-weight: 700; }
.bb-loc-addr  { font-size: 12px; color: #aaa; margin-top: 2px; }
.bb-loc-count { font-size: 12px; color: #4a9; margin-top: 2px; font-weight: 500; }
.bb-loc-arrow { font-size: 22px; color: #ccc; flex-shrink: 0; }

/* Search bar */
.bb-search-wrap {
  background: #fff;
  padding: 10px 14px;
  border-bottom: 1px solid #ebebeb;
  position: sticky;
  top: 0;
  z-index: 10;
}
.bb-search {
  width: 100%;
  padding: 9px 14px 9px 36px;
  border: 1.5px solid #e0e0e0;
  border-radius: 10px;
  font-size: 15px;
  font-family: inherit;
  outline: none;
  background: #f9f9f9 url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%23aaa' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'/%3E%3C/svg%3E") 10px center no-repeat;
  transition: border-color .15s, background .15s;
}
.bb-search:focus { border-color: #1C3D5A; background-color: #fff; }

/* Updated timestamp */
.bb-updated {
  text-align: center;
  font-size: 11px;
  color: #bbb;
  padding: 7px 0 2px;
}

/* Product grid */
.bb-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 10px;
  padding: 10px;
}
@media (min-width: 480px) { .bb-grid { grid-template-columns: repeat(3, 1fr); } }
@media (min-width: 720px) { .bb-grid { grid-template-columns: repeat(4, 1fr); gap: 14px; padding: 14px; } }

/* Product card */
.bb-card {
  background: #fff;
  border-radius: 14px;
  overflow: hidden;
  box-shadow: 0 2px 8px rgba(0,0,0,0.07);
}
.bb-card.bb-oos { opacity: 0.45; }
.bb-card-img {
  width: 100%;
  aspect-ratio: 1 / 1;
  background: #f0f0ed;
  overflow: hidden;
  display: flex;
  align-items: center;
  justify-content: center;
}
.bb-card-img img {
  width: 100%;
  height: 100%;
  object-fit: cover;
  display: block;
}
.bb-card-img .bb-img-fallback {
  font-size: 38px;
  opacity: 0.4;
}
.bb-card-body { padding: 8px 10px 12px; }
.bb-card-name {
  font-size: 12px;
  font-weight: 500;
  color: #333;
  line-height: 1.35;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
  min-height: 2.7em;
  margin-bottom: 5px;
}
.bb-card-price { font-size: 17px; font-weight: 800; color: #1C3D5A; }
.bb-card-oos-label { font-size: 11px; color: #bbb; font-weight: 500; }

/* Empty / loading states */
.bb-state {
  text-align: center;
  padding: 60px 20px;
  color: #aaa;
  font-size: 15px;
  line-height: 1.6;
}
.bb-no-results {
  grid-column: 1 / -1;
  text-align: center;
  padding: 48px 16px;
  color: #aaa;
  font-size: 14px;
}

/* Footer */
.bb-footer {
  text-align: center;
  padding: 28px 16px 16px;
  font-size: 11px;
  color: #ccc;
}
`;

  /* ── State ────────────────────────────────────────────────────────────── */
  let appData = null;
  let container = null;

  /* ── Helpers ──────────────────────────────────────────────────────────── */
  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function timeAgo(iso) {
    const secs = Math.floor((Date.now() - new Date(iso)) / 1000);
    if (secs < 90)    return 'just now';
    if (secs < 3600)  return `${Math.floor(secs / 60)} min ago`;
    if (secs < 86400) return `${Math.floor(secs / 3600)}h ago`;
    return `${Math.floor(secs / 86400)}d ago`;
  }

  /* ── URL routing ──────────────────────────────────────────────────────── */
  function getLocId() {
    return new URLSearchParams(window.location.search).get('loc');
  }

  function navigate(locId) {
    const u = new URL(window.location.href);
    if (locId) u.searchParams.set('loc', locId);
    else u.searchParams.delete('loc');
    history.pushState({}, '', u);
    route();
  }

  /* ── Render: location selector ────────────────────────────────────────── */
  function renderSelector() {
    container.innerHTML = '';

    // Header
    const hdr = el('div', 'bb-header');
    const htitle = el('div');
    htitle.appendChild(el('div', 'bb-header-title', '🧺 Blanket Box Vending'));
    htitle.appendChild(el('div', 'bb-header-sub', "Scan to see what's inside"));
    hdr.appendChild(htitle);
    container.appendChild(hdr);

    // Location cards
    const wrap = el('div', 'bb-selector-wrap');
    wrap.appendChild(el('div', 'bb-selector-heading', 'Select your location'));
    wrap.appendChild(el('div', 'bb-selector-sub', 'Tap to see live prices'));

    for (const loc of appData.locations) {
      const card = el('div', 'bb-loc-card');
      card.appendChild(el('span', 'bb-loc-icon', ICONS[loc.id] || '📍'));

      const info = el('div', 'bb-loc-info');
      info.appendChild(el('div', 'bb-loc-name', loc.name));
      if (loc.address) info.appendChild(el('div', 'bb-loc-addr', loc.address));
      info.appendChild(el('div', 'bb-loc-count',
        `${loc.inStockCount} item${loc.inStockCount !== 1 ? 's' : ''} in stock`));
      card.appendChild(info);
      card.appendChild(el('span', 'bb-loc-arrow', '›'));

      card.addEventListener('click', () => navigate(loc.id));
      wrap.appendChild(card);
    }

    wrap.appendChild(el('div', 'bb-updated', `Prices updated ${timeAgo(appData.updated)}`));
    container.appendChild(wrap);
  }

  /* ── Render: product grid ─────────────────────────────────────────────── */
  function renderLocation(loc) {
    container.innerHTML = '';

    // Header with back button
    const hdr = el('div', 'bb-header');
    const backBtn = el('button', 'bb-back-btn', '‹');
    backBtn.title = 'All locations';
    backBtn.addEventListener('click', () => navigate(null));
    hdr.appendChild(backBtn);

    const htitle = el('div');
    htitle.appendChild(el('div', 'bb-header-title', loc.name));
    htitle.appendChild(el('div', 'bb-header-sub',
      loc.address || `${loc.inStockCount} item${loc.inStockCount !== 1 ? 's' : ''} in stock`));
    hdr.appendChild(htitle);
    container.appendChild(hdr);

    // Search
    const searchWrap = el('div', 'bb-search-wrap');
    const searchInput = el('input', 'bb-search');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search products…';
    searchInput.setAttribute('autocomplete', 'off');
    searchWrap.appendChild(searchInput);
    container.appendChild(searchWrap);

    container.appendChild(el('div', 'bb-updated', `Prices updated ${timeAgo(appData.updated)}`));

    // Grid
    const grid = el('div', 'bb-grid');
    container.appendChild(grid);

    function updateGrid() {
      const q = searchInput.value.trim().toLowerCase();
      grid.innerHTML = '';

      const filtered = q
        ? loc.products.filter(p => p.name.toLowerCase().includes(q))
        : loc.products;

      // In-stock items first
      const ordered = [
        ...filtered.filter(p => p.inStock),
        ...filtered.filter(p => !p.inStock),
      ];

      if (ordered.length === 0) {
        grid.appendChild(el('div', 'bb-no-results', 'No products match your search.'));
        return;
      }

      for (const p of ordered) {
        grid.appendChild(productCard(p));
      }
    }

    searchInput.addEventListener('input', updateGrid);
    updateGrid();

    container.appendChild(el('div', 'bb-footer', 'blanketboxvending.com'));
  }

  /* ── Render: single product card ──────────────────────────────────────── */
  function productCard(p) {
    const card = el('div', 'bb-card' + (p.inStock ? '' : ' bb-oos'));

    // Image
    const imgWrap = el('div', 'bb-card-img');
    if (p.image) {
      const img = document.createElement('img');
      img.alt = p.name;
      img.loading = 'lazy';
      img.onerror = () => {
        imgWrap.innerHTML = '';
        imgWrap.appendChild(el('span', 'bb-img-fallback', '📦'));
      };
      img.src = p.image;
      imgWrap.appendChild(img);
    } else {
      imgWrap.appendChild(el('span', 'bb-img-fallback', '📦'));
    }
    card.appendChild(imgWrap);

    // Body
    const body = el('div', 'bb-card-body');
    body.appendChild(el('div', 'bb-card-name', p.name));
    if (p.inStock) {
      body.appendChild(el('div', 'bb-card-price', `$${p.price.toFixed(2)}`));
    } else {
      body.appendChild(el('div', 'bb-card-oos-label', 'Out of stock'));
    }
    card.appendChild(body);

    return card;
  }

  /* ── Router ───────────────────────────────────────────────────────────── */
  function route() {
    const locId = getLocId();
    if (locId) {
      const loc = appData.locations.find(l => l.id === locId);
      if (loc) return renderLocation(loc);
    }
    renderSelector();
  }

  /* ── Bootstrap ────────────────────────────────────────────────────────── */
  function init() {
    // Mount point: use #bb-prices-app if present, otherwise append to body
    container = document.getElementById('bb-prices-app') || document.body;
    container.id = 'bb-app';

    // Inject styles
    const style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);

    // Loading state
    const loading = el('div', 'bb-state', 'Loading prices…');
    container.appendChild(loading);

    fetch(DATA_URL)
      .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.json(); })
      .then(data => {
        appData = data;
        container.innerHTML = '';
        route();
        window.addEventListener('popstate', route);
      })
      .catch(err => {
        container.innerHTML = '';
        const msg = el('div', 'bb-state');
        msg.innerHTML = 'Could not load prices — please try again.<br><small style="font-size:11px;opacity:.6">' + err.message + '</small>';
        container.appendChild(msg);
      });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
