# =============================================================================
# modules/zoho.py (FINAL WITH FULL LOGGING + 429 FIX)
# =============================================================================

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import asyncio
import aiohttp
from PIL import Image as PILImage
import config

# ---------------------------------------------------------------------------
# In-memory cache
# ---------------------------------------------------------------------------
_cache_by_sku  = {}
_cache_by_name = {}
_units_cache   = {}
_cache_loaded  = False

PER_PAGE       = 200
MAX_CONCURRENT = 3
RATE_LIMIT     = asyncio.Semaphore(3)  # prevent 429


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

async def load_cache(log_fn=None):
    global _cache_loaded

    _cache_by_sku.clear()
    _cache_by_name.clear()
    _units_cache.clear()
    _cache_loaded = False

    await _ensure_token(log_fn)
    await _log(log_fn, "[Zoho] 🔄 Starting full cache load...")

    page = 1
    total = 0

    while True:
        pages = list(range(page, page + MAX_CONCURRENT))
        await _log(log_fn, f"[Zoho] ⚡ Fetching pages {pages}")

        results = await asyncio.gather(
            *[_fetch_page(p, log_fn) for p in pages],
            return_exceptions=True
        )

        finished = False

        for p_num, result in zip(pages, results):

            if isinstance(result, Exception):
                await _log(log_fn, f"[Zoho] ❌ Page {p_num} crashed: {result}")
                continue

            items, more = result

            if items is None:
                await _log(log_fn, f"[Zoho] ⚠️ Page {p_num} returned None")
                continue

            _index_items(items)
            total += len(items)

            await _log(log_fn, f"[Zoho] 📦 Page {p_num} → {len(items)} items")

            if not more:
                finished = True

        await _log(log_fn, f"[Zoho] 📊 Total loaded so far: {total}")

        if finished:
            break

        page += MAX_CONCURRENT

    _cache_loaded = True
    await _log(log_fn, f"[Zoho] ✅ Done — {len(_cache_by_sku)} by SKU, {len(_cache_by_name)} by name")


# ---------------------------------------------------------------------------
# Lookup APIs
# ---------------------------------------------------------------------------

async def get_item_photo(sku: str, item_name: str = ""):
    clean_sku  = _clean(sku)
    clean_name = _clean(item_name)

    item = _cache_by_sku.get(clean_sku) if clean_sku else None

    if not item and clean_name:
        item = _cache_by_name.get(clean_name)
        if item:
            print(f"  [Zoho] ✅ Found by name: '{clean_name}'")
    elif item:
        print(f"  [Zoho] ✅ Found by SKU: '{clean_sku}'")

    if not item:
        print(f"  [Zoho] ⚪️ Not found: sku='{clean_sku}' name='{clean_name}'")
        return None

    if not item.get("image_name"):
        print(f"  [Zoho] ⚪️ No image for: {item.get('name')}")
        return None

    return await _fetch_image(item["item_id"], item.get("name"))


async def get_units_per_box(sku: str, item_name: str = "") -> int | None:
    clean_sku  = _clean(sku)
    clean_name = _clean(item_name)

    key = clean_sku or clean_name

    if key in _units_cache:
        print(f"  [Zoho] ⚡ Cache hit (Units/Box): {key}")
        return _units_cache[key]

    item = _cache_by_sku.get(clean_sku) if clean_sku else None
    if not item and clean_name:
        item = _cache_by_name.get(clean_name)

    if not item:
        print(f"  [Zoho] ⚪️ units_per_box: item not found sku='{clean_sku}' name='{clean_name}'")
        return None

    full_item = await _fetch_item_detail(item["item_id"])
    if not full_item:
        return None

    for cf in full_item.get("custom_fields", []):
        label = str(cf.get("label", "")).lower()
        if "units" in label and "box" in label:
            raw = cf.get("value", "")
            print(f"  [Zoho] 📦 Units per Box for '{item.get('name')}': {raw}")
            try:
                val = int(float(raw))
                _units_cache[key] = val
                return val
            except:
                return None

    print(f"  [Zoho] ⚪️ No Units per Box for '{item.get('name')}'")
    return None


# ---------------------------------------------------------------------------
# Fetch Page (with logging + retry)
# ---------------------------------------------------------------------------

async def _fetch_page(page, log_fn=None):
    url = f"{config.ZOHO_BASE_URL}/items"
    params = {
        "organization_id": config.ZOHO_ORG_ID,
        "page": page,
        "per_page": PER_PAGE,
    }

    for attempt in range(5):
        async with RATE_LIMIT:
            try:
                await asyncio.sleep(0.15)

                async with aiohttp.ClientSession(headers=_headers()) as session:
                    async with session.get(url, params=params) as r:

                        if r.status == 401:
                            await _log(log_fn, "[Zoho] 🔄 Token expired — refreshing...")
                            await _refresh_token()
                            continue

                        if r.status == 429:
                            wait = 2 ** attempt
                            await _log(log_fn, f"[Zoho] ⏳ 429 on page {page}, retry in {wait}s")
                            await asyncio.sleep(wait)
                            continue

                        if r.status != 200:
                            body = await r.text()
                            await _log(log_fn, f"[Zoho] ❌ HTTP {r.status} page {page}: {body[:100]}")
                            return None, False

                        data = await r.json()
                        items = data.get("items", [])

                        return items, len(items) == PER_PAGE

            except Exception as e:
                await _log(log_fn, f"[Zoho] ❌ Page {page} error: {e}")

    return None, False


# ---------------------------------------------------------------------------
# Fetch Item Detail (with logs + retry)
# ---------------------------------------------------------------------------

async def _fetch_item_detail(item_id):
    url = f"{config.ZOHO_BASE_URL}/items/{item_id}"
    params = {"organization_id": config.ZOHO_ORG_ID}

    for attempt in range(5):
        async with RATE_LIMIT:
            try:
                await asyncio.sleep(0.15)

                async with aiohttp.ClientSession(headers=_headers()) as session:
                    async with session.get(url, params=params) as r:

                        if r.status == 401:
                            await _refresh_token()
                            continue

                        if r.status == 429:
                            wait = 2 ** attempt
                            print(f"  [Zoho] ⏳ 429 detail retry in {wait}s (id={item_id})")
                            await asyncio.sleep(wait)
                            continue

                        if r.status != 200:
                            print(f"  [Zoho] ❌ Item detail HTTP {r.status} for id={item_id}")
                            return None

                        data = await r.json()
                        return data.get("item")

            except Exception as e:
                print(f"  [Zoho] ❌ Item detail error id={item_id}: {e}")

    return None


# ---------------------------------------------------------------------------
# Fetch Image (with logs + retry)
# ---------------------------------------------------------------------------

async def _fetch_image(item_id, label):
    url = f"{config.ZOHO_BASE_URL}/items/{item_id}/image"
    params = {"organization_id": config.ZOHO_ORG_ID}

    for attempt in range(5):
        async with RATE_LIMIT:
            try:
                await asyncio.sleep(0.15)

                async with aiohttp.ClientSession(headers=_headers()) as session:
                    async with session.get(url, params=params) as r:

                        if r.status == 401:
                            await _refresh_token()
                            continue

                        if r.status == 429:
                            wait = 2 ** attempt
                            print(f"  [Zoho] ⏳ 429 image retry in {wait}s for '{label}'")
                            await asyncio.sleep(wait)
                            continue

                        if r.status != 200:
                            print(f"  [Zoho] ❌ Image HTTP {r.status} for '{label}'")
                            return None

                        data = await r.read()
                        return _to_pil(data, label)

            except Exception as e:
                print(f"  [Zoho] ❌ Image error '{label}': {e}")

    return None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _index_items(items):
    for item in items:
        sku  = _clean(item.get("sku"))
        name = _clean(item.get("name"))
        if sku:
            _cache_by_sku[sku] = item
        if name:
            _cache_by_name[name] = item


async def _ensure_token(log_fn=None):
    if not getattr(config, "ZOHO_ACCESS_TOKEN", ""):
        await _log(log_fn, "[Zoho] 🔄 No token, refreshing...")
        await _refresh_token()


async def _refresh_token():
    url = "https://accounts.zoho.com/oauth/v2/token"
    params = {
        "refresh_token": config.ZOHO_REFRESH_TOKEN,
        "client_id": config.ZOHO_CLIENT_ID,
        "client_secret": config.ZOHO_CLIENT_SECRET,
        "grant_type": "refresh_token",
    }

    async with aiohttp.ClientSession() as session:
        async with session.post(url, params=params) as r:
            data = await r.json()

            if "access_token" in data:
                config.ZOHO_ACCESS_TOKEN = data["access_token"]
                print("[Zoho] ✅ Token refreshed")
                return True

            print("[Zoho] ❌ Token refresh failed", data)
            return False


def _headers():
    return {"Authorization": f"Zoho-oauthtoken {config.ZOHO_ACCESS_TOKEN}"}


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip().lower()
    return "" if s in ("", "nan", "none") else s


async def _log(log_fn, msg):
    print(msg)
    if log_fn:
        await log_fn(msg)


def _to_pil(data, label):
    try:
        img = PILImage.open(io.BytesIO(data)).convert("RGB")
        img = img.resize(
            (config.PHOTO_WIDTH_PX, config.PHOTO_HEIGHT_PX),
            PILImage.LANCZOS
        )
        return img
    except Exception as e:
        print(f"  [Zoho] ❌ Could not decode image '{label}': {e}")
        return None