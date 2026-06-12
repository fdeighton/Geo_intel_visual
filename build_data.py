"""Geocode an executive prospect list and bake it into data.js for the map viewer.

Usage:
    python build_data.py [path-to-xlsx-or-csv]

With no argument it picks the most-recently-modified .xlsx/.csv in this folder
(ignoring data.js and Excel ~$ lock files). It then auto-selects the richest
sheet and resolves columns by flexible header matching, so dropping in a new
export with slightly different headers / sheet names still works.

Idempotent: geocode results are cached in geocode_cache.json next to this
script, so reruns only look up addresses they haven't seen before.
"""
import json, os, sys, glob, time, urllib.request, urllib.parse, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "geocode_cache.json")
OUT = os.path.join(HERE, "data.js")
UA = "geo-intel-visual/1.0 (fdeighton@fitzrovia.ca)"

# Michael Garron Hospital — the "Distance to MGH" reference; biases geocoding
# toward the right Toronto when an address is ambiguous.
GEO_SUFFIX = "Toronto, Ontario, Canada"

# Canonical field -> accepted header spellings (compared case/space/punct-insensitive).
ALIASES = {
    "name":        ["name", "fullname", "prospectname", "donorname", "contactname"],
    "role":        ["role", "title", "jobtitle", "position"],
    "company":     ["company", "organization", "organisation", "employer", "firm", "org"],
    "address":     ["propertyaddress", "address", "streetaddress", "fulladdress", "mailingaddress"],
    "salePrice":   ["saleprice", "price", "soldprice", "salepriceusd", "purchaseprice", "lastsaleprice"],
    "saleDate":    ["saledate", "date", "solddate", "purchasedate", "lastsaledate"],
    "distanceMGH": ["distancetomghkm", "distancetomgh", "distancekm", "distance", "mghdistance"],
    "prestige":    ["prestigescore0100", "prestigescore", "prestige"],
    "tier":        ["senioritytier", "tier", "seniority"],
    "wealth":      ["wealthsignal", "wealth", "wealthtier"],
    "category":    ["donorcategory", "category", "segment"],
    "notes":       ["notes", "note", "comments", "comment", "remarks"],
}
# Fields a "good" sheet should have; used to score which sheet to read.
KEY_FIELDS = ["name", "address"]


def norm(s):
    return "".join(c for c in str(s).lower() if c.isalnum())


def resolve_columns(headers):
    """Map canonical field -> column index, by alias match on normalized headers."""
    normed = [(i, norm(h)) for i, h in enumerate(headers)]
    out = {}
    for field, names in ALIASES.items():
        for want in names:
            hit = next((i for i, n in normed if n == want), None)
            if hit is not None:
                out[field] = hit
                break
    return out


def pick_source():
    if len(sys.argv) > 1:
        p = sys.argv[1]
        if not os.path.isabs(p):
            p = os.path.join(HERE, p)
        if not os.path.exists(p):
            sys.exit("File not found: " + p)
        return p
    candidates = []
    for pat in ("*.xlsx", "*.xls", "*.csv"):
        for f in glob.glob(os.path.join(HERE, pat)):
            base = os.path.basename(f)
            if base.startswith("~$") or base == "data.js":
                continue
            candidates.append(f)
    if not candidates:
        sys.exit("No .xlsx/.xls/.csv found in " + HERE + " (pass a path as an argument).")
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def read_rows(path):
    """Return (headers, list-of-row-tuples) from the richest sheet of an Excel
    file, or from a CSV."""
    if path.lower().endswith(".csv"):
        import csv
        with open(path, newline="", encoding="utf-8-sig") as f:
            r = list(csv.reader(f))
        print("  source: CSV")
        return r[0], r[1:]

    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    best = None  # (score, ncols, name, headers, data)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [("" if h is None else str(h).strip()) for h in rows[0]]
        cols = resolve_columns(headers)
        if not all(k in cols for k in KEY_FIELDS):
            continue  # can't even find name + address
        score = len(cols)
        cand = (score, len(headers), ws.title, headers, rows[1:])
        if best is None or cand[:2] > best[:2]:
            best = cand
    if best is None:
        sys.exit("No sheet with usable Name + Address columns found in " + os.path.basename(path))
    print(f"  sheet: '{best[2]}'  ({best[0]} recognized columns)")
    return best[3], best[4]


# ---- geocoding -------------------------------------------------------------
cache = {}
if os.path.exists(CACHE):
    with open(CACHE, encoding="utf-8") as f:
        cache = json.load(f)


def geocode(addr):
    a = addr.strip()
    if a in cache:
        return cache[a]
    q = a if GEO_SUFFIX.split(",")[0].lower() in a.lower() else a + ", " + GEO_SUFFIX
    if "ontario" not in q.lower():
        q = a + ", " + GEO_SUFFIX
    params = urllib.parse.urlencode({"q": q, "format": "json", "limit": 1, "countrycodes": "ca"})
    req = urllib.request.Request("https://nominatim.openstreetmap.org/search?" + params,
                                 headers={"User-Agent": UA})
    result = None
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            d = json.load(r)
        if d:
            result = [round(float(d[0]["lat"]), 6), round(float(d[0]["lon"]), 6)]
    except Exception as e:
        print("  ! geocode error on", addr, ":", e, flush=True)
    cache[a] = result
    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    time.sleep(1.1)  # Nominatim usage policy: <= 1 request / second
    return result


# ---- value helpers ---------------------------------------------------------
def cell(row, i):
    if i is None or i >= len(row):
        return None
    v = row[i]
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return v


def as_text(row, i):
    v = cell(row, i)
    return str(v).strip() if v is not None else ""


def as_num(row, i):
    v = cell(row, i)
    if v is None or v == "" or (isinstance(v, str) and v.strip().upper() in ("N/A", "NA", "-")):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = "".join(c for c in str(v) if c.isdigit() or c in ".-")
    try:
        return float(s) if s not in ("", "-", ".", "-.") else None
    except ValueError:
        return None


# ---- wealth signal: blend of sale-price tier and prestige tier -------------
# The source "Wealth Signal" column is just a relabeling of the prestige score
# (role-based), so it ignores the home's value. We recompute it as a blend so
# both the property and the person's seniority count.
WEALTH_LABELS = {1: "Low", 2: "Medium", 3: "High", 4: "Very High"}


def _price_tier(p):
    if p is None:
        return None
    if p >= 5_000_000:
        return 4
    if p >= 3_500_000:
        return 3
    if p >= 2_000_000:
        return 2
    return 1


def _prestige_tier(p):
    if p is None:
        return None
    if p >= 90:
        return 4
    if p >= 80:
        return 3
    if p >= 60:
        return 2
    return 1


def wealth_signal(price, prestige):
    tiers = [t for t in (_price_tier(price), _prestige_tier(prestige)) if t is not None]
    if not tiers:
        return ""  # neither price nor prestige available
    avg = sum(tiers) / len(tiers)
    # Round to nearest tier; a .5 tie rounds DOWN (e.g. Very-High home + junior
    # role = (4+1)/2 = 2.5 -> Medium), avoiding banker's-rounding surprises.
    tier = int(avg - 0.5) if avg % 1 == 0.5 else round(avg)
    return WEALTH_LABELS[max(1, min(4, tier))]


# ---- main ------------------------------------------------------------------
src = pick_source()
print("Reading", os.path.basename(src))
headers, data = read_rows(src)
cols = resolve_columns(headers)
missing = [f for f in ALIASES if f not in cols]
if missing:
    print("  (optional fields not found, left blank:", ", ".join(missing) + ")")

# Geocode unique addresses first.
uniq = sorted({as_text(r, cols.get("address")) for r in data if as_text(r, cols.get("address"))})
print("Geocoding", len(uniq), "unique addresses...", flush=True)
for n, a in enumerate(uniq, 1):
    geocode(a)
    if n % 20 == 0:
        located = sum(1 for x in uniq[:n] if cache.get(x))
        print(f"  {n}/{len(uniq)} ({located} located)", flush=True)

records = []
for r in data:
    name = as_text(r, cols.get("name"))
    if not name:
        continue
    addr = as_text(r, cols.get("address"))
    coord = cache.get(addr) if addr else None
    sale_price = as_num(r, cols.get("salePrice"))
    prestige = as_num(r, cols.get("prestige"))
    records.append({
        "name": name,
        "role": as_text(r, cols.get("role")),
        "company": as_text(r, cols.get("company")),
        "address": addr,
        "salePrice": sale_price,
        "saleDate": as_text(r, cols.get("saleDate")),
        "distanceMGH": as_num(r, cols.get("distanceMGH")),
        "prestige": prestige,
        "tier": as_text(r, cols.get("tier")),
        "wealth": wealth_signal(sale_price, prestige),
        "category": as_text(r, cols.get("category")),
        "notes": as_text(r, cols.get("notes")),
        "lat": coord[0] if coord else None,
        "lng": coord[1] if coord else None,
    })

located = sum(1 for r in records if r["lat"] is not None)
print(f"\n{len(records)} records, {located} geocoded, {len(records) - located} without coordinates", flush=True)
if located < len(records):
    print("  unmapped (stay in the table, flagged 'no map'):")
    for r in records:
        if r["lat"] is None:
            print("   -", r["name"], "|", r["address"] or "(no address)")

meta = {"source": os.path.basename(src), "count": len(records), "located": located}
with open(OUT, "w", encoding="utf-8") as f:
    f.write("/* Auto-generated by build_data.py — do not edit by hand.\n")
    f.write(f"   {meta['count']} prospects from {meta['source']}; {meta['located']} geocoded\n")
    f.write("   via OpenStreetMap/Nominatim. Re-run build_data.py to regenerate. */\n")
    f.write("window.PROSPECT_META = ")
    json.dump(meta, f, ensure_ascii=False)
    f.write(";\n")
    f.write("window.PROSPECT_DATA = ")
    json.dump(records, f, ensure_ascii=False)
    f.write(";\n")
print("Wrote", OUT, flush=True)
