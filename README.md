# Executive Prospect Map

A standalone, local-first dashboard that plots the executive prospect list
(`executive_list_handoff.xlsx`) on an interactive map alongside a synchronized
table, with a few filters. The coordinates are **pre-geocoded** from each
property address and baked into a small `data.js`, so the viewer itself is just
two static files and opens instantly — no backend, no build step at view time,
no upload, no API keys.

---

## Quick start

### Just open it
Double-click **`index.html`**. It opens in your default browser and shows the
map populated with every prospect.

> Map tiles and Leaflet load from public CDNs, so an internet connection is
> needed for the map to draw. No keys are used or required.

### Or serve the folder (avoids `file://` edge cases)
```bash
# from this folder
python -m http.server 8000
# then open http://localhost:8000
```

---

## What it shows

- **Map** — every prospect whose address geocoded successfully appears as a
  pin, colored by **prestige score** (cooler = lower, brass = higher). The red
  marker is Michael Garron Hospital, the reference point for the "Distance to
  MGH" figures. Click any pin for a profile card.
- **Profile card** — Name, Role, Company, Sale Price, Prestige, Wealth Signal,
  Distance to MGH, Seniority Tier, Donor Category, Property Address (with sale
  date), and Notes.
- **Synchronized table** — mirrors the current filtered view. Click a row to
  fly the map to that prospect and open their card.
- **Filters** — name/company search, donor category, seniority tier, wealth
  signal, minimum sale price, and minimum prestige score.
- **Sorting** — by prestige, sale price, distance to MGH, or name (dropdown +
  direction toggle, or click the numeric column headers).
- **Summary stats** — prospects shown, median sale price, top donor category,
  and the count of very-high-wealth prospects. Always reflect the filtered set.

---

## Files

```
index.html      The app shell, layout, and styles
app.js          Mapping, filtering, sorting, stats (reads window.PROSPECT_DATA)
data.js         Auto-generated: the prospects with baked-in lat/lng
build_data.py   Regenerates data.js from the Excel (geocodes addresses)
geocode_cache.json  Cached address -> lat/lng, so reruns are fast
executive_list_handoff.xlsx  Source data
README.md       This file
```

---

## Dropping in a new Excel file

When the data changes, just **drop the new file into this folder and re-run the
build** — no code edits needed:

```bash
pip install openpyxl          # one time
python build_data.py          # uses the newest .xlsx/.csv in this folder
# or point it at a specific file:
python build_data.py "C:\path\to\new_list.xlsx"
```

The script is tolerant of how the new export is shaped:

- **File** — with no argument it picks the most recently modified
  `.xlsx`/`.xls`/`.csv` in this folder (ignoring `data.js` and Excel `~$` lock
  files). Pass a path to override.
- **Sheet** — for multi-sheet workbooks it auto-selects the sheet with the most
  recognized columns (here, *Role Sorted*).
- **Columns** — headers are matched case/space/punctuation-insensitively with
  aliases, so `Sale Price`, `sale_price`, `Price`, and `Sold Price` all resolve.
  To teach it a new spelling, add it to the `ALIASES` table at the top of
  `build_data.py`. A column it can't find is simply left blank.

It geocodes each unique property address via OpenStreetMap's free Nominatim
service (≤1 request/second, per their usage policy) and writes `data.js`.
Results are cached in `geocode_cache.json` next to the script, so reruns only
look up addresses they haven't seen — a refreshed list with mostly familiar
addresses rebuilds in seconds. The build prints which file/sheet it used and
lists any addresses that didn't resolve.

---

## Notes

- **Geocoding is approximate.** Addresses are resolved to a single best match
  with `, Toronto, Ontario, Canada` appended when not already present. A handful
  of addresses may not resolve; those prospects stay in the table flagged
  `no map` and are excluded from the map only.
- **Data is read entirely in the browser.** Nothing is uploaded anywhere when
  you view the dashboard. (Geocoding happens once, offline, in `build_data.py`.)
- The map auto-fits to the located prospects (and the hospital) on load.
```
