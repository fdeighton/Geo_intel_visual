/* Executive Prospect Map — local-first, no backend, no build at view time.
   Reads window.PROSPECT_DATA (baked from the Excel by build_data.py, with
   coordinates already geocoded), maps every located prospect, keeps a
   synchronized table, a few filters, sorting and summary stats. */
(function () {
  "use strict";

  // Michael Garron Hospital, 825 Coxwell Ave — the "Distance to MGH" reference.
  var MGH = { lat: 43.689654, lng: -79.325653, name: "Michael Garron Hospital" };

  // ---- State ------------------------------------------------------------
  var records = [];
  var maxPrestige = 0, maxPrice = 0;
  var selectedId = null;
  var sortDir = "desc";

  var map, markerLayer, mghMarker;
  var markersById = {};

  var els = {};
  ["sourceName","fName","fCategory","fTier","fWealth","fPrice","fPriceOut",
   "fPrestige","fPrestigeOut","resetBtn","sortField","sortDir","tbody","tblCount",
   "stTotal","stTotalSub","stMedian","stMedianSub","stCat","stCatSub","stWealth","stWealthSub",
   "toast","empty","emptyTitle","emptyMsg","exportBtn","exportLabel"
  ].forEach(function (id) { els[id] = document.getElementById(id); });

  // ---- Helpers ----------------------------------------------------------
  function str(v) { return v == null ? "" : String(v).trim(); }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[c]; }); }
  function fmtMoney(v) { return v == null ? "—" : "$" + Math.round(v).toLocaleString(); }
  function fmtMoneyShort(v) {
    if (v == null) return "—";
    if (v >= 1e6) return "$" + (v / 1e6).toFixed(v >= 1e7 ? 1 : 2).replace(/\.?0+$/, "") + "M";
    if (v >= 1e3) return "$" + Math.round(v / 1e3) + "k";
    return "$" + Math.round(v);
  }
  function fmtScore(v) { return v == null ? "—" : (Math.round(v * 10) / 10).toString(); }
  function fmtKm(v) { return v == null ? "—" : (Math.round(v * 100) / 100) + " km"; }

  function wealthClass(w) {
    var n = w.toLowerCase();
    if (/very high/.test(n)) return "w-vhigh";
    if (/high/.test(n)) return "w-high";
    if (/med/.test(n)) return "w-med";
    return "";
  }

  // ---- Map init ---------------------------------------------------------
  function initMap() {
    map = L.map("map", { zoomControl: true, preferCanvas: true }).setView([43.69, -79.37], 12);
    L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
      attribution: '&copy; OpenStreetMap &copy; CARTO', maxZoom: 19, subdomains: "abcd"
    }).addTo(map);
    markerLayer = L.layerGroup().addTo(map);

    // When a prospect's card closes, return its marker to the normal style.
    // The deferred check tells "closed the card" apart from "switched to
    // another prospect" — switching opens a new popup in the same tick, so
    // popupOpen is true again by the time the timeout runs and we leave it.
    var popupOpen = false;
    map.on("popupopen", function () { popupOpen = true; });
    map.on("popupclose", function () {
      popupOpen = false;
      setTimeout(function () { if (!popupOpen) deselect(); }, 0);
    });

    mghMarker = L.circleMarker([MGH.lat, MGH.lng], {
      radius: 9, color: "#fff", weight: 2.5, fillColor: "#9a3b3b", fillOpacity: 1
    }).bindTooltip(MGH.name, { direction: "top", offset: [0, -6] }).addTo(map);
  }

  // ---- Prestige color ramp ---------------------------------------------
  var RAMP = ["#9aa6b4", "#7d97a0", "#5f8a7e", "#9a8240", "#b0853f"];
  function scoreColor(prestige) {
    if (prestige == null || maxPrestige <= 0) return "#9aa6b4";
    var f = Math.max(0, Math.min(1, prestige / maxPrestige));
    return RAMP[Math.min(4, Math.floor(f * 4.0001))];
  }

  // ---- Ingest baked data ------------------------------------------------
  function ingest(data) {
    records = data.map(function (raw, i) {
      var lat = raw.lat, lng = raw.lng;
      var hasGeo = typeof lat === "number" && typeof lng === "number" &&
                   lat >= -90 && lat <= 90 && lng >= -180 && lng <= 180;
      return {
        _id: i,
        name: str(raw.name) || "(unnamed)",
        role: str(raw.role),
        company: str(raw.company),
        address: str(raw.address),
        salePrice: typeof raw.salePrice === "number" ? raw.salePrice : null,
        saleDate: str(raw.saleDate),
        distanceMGH: typeof raw.distanceMGH === "number" ? raw.distanceMGH : null,
        prestige: typeof raw.prestige === "number" ? raw.prestige : null,
        tier: str(raw.tier),
        wealth: str(raw.wealth),
        category: str(raw.category),
        notes: str(raw.notes),
        lat: hasGeo ? lat : null,
        lng: hasGeo ? lng : null,
        hasGeo: hasGeo
      };
    });

    maxPrestige = records.reduce(function (m, r) { return r.prestige != null ? Math.max(m, r.prestige) : m; }, 0);
    maxPrice    = records.reduce(function (m, r) { return r.salePrice != null ? Math.max(m, r.salePrice) : m; }, 0);

    configureRanges();
    populateDropdowns();

    var geoN = records.filter(function (r) { return r.hasGeo; }).length;
    var srcName = (window.PROSPECT_META && window.PROSPECT_META.source) || "executive_list_handoff.xlsx";
    els.sourceName.textContent = srcName + " · " + records.length + " prospects";
    els.empty.classList.add("hidden");
    render(true);
    toast(records.length + " prospects · " + geoN + " mapped");
  }

  function configureRanges() {
    var pMax = Math.max(1, Math.ceil(maxPrice / 1e5) * 1e5);
    els.fPrice.max = pMax; els.fPrice.step = 1e5; els.fPrice.value = 0; els.fPriceOut.textContent = "$0";
    var prMax = Math.max(1, Math.ceil(maxPrestige));
    els.fPrestige.max = prMax; els.fPrestige.value = 0; els.fPrestigeOut.textContent = "0";
  }

  function uniqueSorted(arr) {
    var seen = {}, out = [];
    arr.forEach(function (v) { v = str(v); if (v && !seen[v]) { seen[v] = 1; out.push(v); } });
    return out.sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); });
  }
  function fillSelect(sel, items) {
    var cur = sel.value, all = sel.firstChild;
    sel.innerHTML = ""; sel.appendChild(all);
    items.forEach(function (v) { var o = document.createElement("option"); o.value = v; o.textContent = v; sel.appendChild(o); });
    if (items.indexOf(cur) >= 0) sel.value = cur;
  }
  function populateDropdowns() {
    fillSelect(els.fCategory, uniqueSorted(records.map(function (r) { return r.category; })));
    fillSelect(els.fTier, uniqueSorted(records.map(function (r) { return r.tier; })));
    fillSelect(els.fWealth, uniqueSorted(records.map(function (r) { return r.wealth; })));
  }

  // ---- Filtering & sorting ---------------------------------------------
  function currentFilters() {
    return {
      q: els.fName.value.trim().toLowerCase(),
      category: els.fCategory.value,
      tier: els.fTier.value,
      wealth: els.fWealth.value,
      minPrice: parseFloat(els.fPrice.value) || 0,
      minPrestige: parseFloat(els.fPrestige.value) || 0
    };
  }
  function applyFilters() {
    var f = currentFilters();
    return records.filter(function (r) {
      if (f.q && (r.name.toLowerCase().indexOf(f.q) < 0 && r.company.toLowerCase().indexOf(f.q) < 0)) return false;
      if (f.category && r.category !== f.category) return false;
      if (f.tier && r.tier !== f.tier) return false;
      if (f.wealth && r.wealth !== f.wealth) return false;
      if (f.minPrice > 0 && (r.salePrice == null || r.salePrice < f.minPrice)) return false;
      if (f.minPrestige > 0 && (r.prestige == null || r.prestige < f.minPrestige)) return false;
      return true;
    });
  }
  function sortRecords(list) {
    var field = els.sortField.value;
    var mult = sortDir === "asc" ? 1 : -1;
    return list.slice().sort(function (a, b) {
      if (field === "name") return a.name.localeCompare(b.name) * mult;
      var av = a[field], bv = b[field];
      var an = av == null, bn = bv == null;
      if (an && bn) return 0;
      if (an) return 1;   // nulls always last
      if (bn) return -1;
      return (av - bv) * mult;
    });
  }

  // ---- Rendering --------------------------------------------------------
  function render(fitBounds) {
    var filtered = applyFilters();
    renderTable(sortRecords(filtered));
    renderMarkers(filtered, fitBounds);
    renderStats(filtered);
    els.tblCount.textContent = "(" + filtered.length + ")";
  }

  function renderTable(list) {
    var tb = els.tbody;
    tb.innerHTML = "";
    if (list.length === 0) {
      var tr = document.createElement("tr");
      tr.className = "nocoord";
      tr.innerHTML = '<td colspan="4" style="padding:26px 12px;color:var(--muted);text-align:center">No prospects match the current filters.</td>';
      tb.appendChild(tr);
      return;
    }
    var frag = document.createDocumentFragment();
    list.forEach(function (r) {
      var tr = document.createElement("tr");
      tr.dataset.id = r._id;
      if (!r.hasGeo) tr.className = "nocoord";
      if (r._id === selectedId) tr.classList.add("sel");
      var rc = [r.role, r.company].filter(Boolean).join(" · ");
      tr.innerHTML =
        '<td><div class="nm">' + esc(r.name) + (r.hasGeo ? "" : ' <span class="nogeo">no map</span>') + '</div>' +
          (rc ? '<div class="sub2">' + esc(rc) + '</div>' : '') + '</td>' +
        '<td>' + (r.wealth ? '<span class="pill ' + wealthClass(r.wealth) + '">' + esc(r.wealth) + '</span>' : '<span style="color:var(--muted)">—</span>') + '</td>' +
        '<td class="r num">' + fmtMoneyShort(r.salePrice) + '</td>' +
        '<td class="r"><span class="badge num" style="background:' + scoreColor(r.prestige) + '">' + fmtScore(r.prestige) + '</span></td>';
      tr.addEventListener("click", function () { selectRecord(r._id, true); });
      frag.appendChild(tr);
    });
    tb.appendChild(frag);
  }

  function profileHTML(r) {
    var rc = [r.role ? "<b>" + esc(r.role) + "</b>" : "", r.company ? esc(r.company) : ""].filter(Boolean).join(" at ");
    function cell(k, v) { return '<div class="cell"><div class="ck">' + k + '</div><div class="cv">' + v + '</div></div>'; }
    return '<div class="card">' +
      '<div class="top"><div class="nm-lg">' + esc(r.name) + '</div>' +
        (rc ? '<div class="rc">' + rc + '</div>' : '') + '</div>' +
      '<div class="grid">' +
        cell("Sale price", fmtMoney(r.salePrice)) +
        cell("Prestige", '<span class="badge num" style="background:' + scoreColor(r.prestige) + '">' + fmtScore(r.prestige) + '</span>') +
        cell("Wealth signal", r.wealth ? esc(r.wealth) : "—") +
        cell("Distance to MGH", fmtKm(r.distanceMGH)) +
        cell("Seniority tier", r.tier ? esc(r.tier) : "—") +
        cell("Donor category", r.category ? esc(r.category) : "—") +
        '<div class="cell full"><div class="ck">Property address</div><div class="cv" style="font-weight:500;font-size:13px">' + (r.address ? esc(r.address) : "—") + (r.saleDate ? ' <span style="color:var(--muted);font-weight:400">· sold ' + esc(r.saleDate) + '</span>' : '') + '</div></div>' +
      '</div>' +
      '</div>';
  }

  function renderMarkers(filtered, fitBounds) {
    markerLayer.clearLayers();
    markersById = {};
    var pts = [];
    filtered.forEach(function (r) {
      if (!r.hasGeo) return;
      var selected = r._id === selectedId;
      var m = L.circleMarker([r.lat, r.lng], {
        radius: selected ? 10 : 6.5,
        color: selected ? "#1b2433" : "#ffffff",
        weight: selected ? 2.5 : 1.5,
        fillColor: scoreColor(r.prestige),
        fillOpacity: 0.92
      });
      m.bindPopup(profileHTML(r), { closeButton: true, autoPanPadding: [40, 40] });
      m.on("click", function () { selectRecord(r._id, false); });
      m.addTo(markerLayer);
      markersById[r._id] = m;
      pts.push([r.lat, r.lng]);
    });
    if (fitBounds && pts.length) {
      pts.push([MGH.lat, MGH.lng]);
      map.fitBounds(pts, { padding: [40, 40], maxZoom: 14 });
    }
  }

  function median(nums) {
    if (!nums.length) return null;
    var s = nums.slice().sort(function (a, b) { return a - b; });
    var mid = Math.floor(s.length / 2);
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  }

  function renderStats(filtered) {
    els.stTotal.textContent = filtered.length.toLocaleString();
    var geoN = filtered.filter(function (r) { return r.hasGeo; }).length;
    els.stTotalSub.textContent = filtered.length ? geoN + " on map · " + (filtered.length - geoN) + " unmapped" : "";

    var prices = filtered.map(function (r) { return r.salePrice; }).filter(function (v) { return v != null; });
    var med = median(prices);
    els.stMedian.textContent = med == null ? "—" : fmtMoneyShort(med);
    els.stMedianSub.textContent = prices.length ? "across " + prices.length + " sales" : "no price data";

    // Top donor category, skipping the catch-all "Other" so the stat names a
    // real industry. Falls through to the next-leading category.
    var counts = {};
    filtered.forEach(function (r) { if (r.category) counts[r.category] = (counts[r.category] || 0) + 1; });
    var top = Object.keys(counts)
      .filter(function (c) { return c.trim().toLowerCase() !== "other"; })
      .sort(function (a, b) { return counts[b] - counts[a]; })[0];
    if (top) { els.stCat.textContent = top; els.stCatSub.textContent = counts[top] + " prospect" + (counts[top] === 1 ? "" : "s"); }
    else { els.stCat.textContent = "—"; els.stCatSub.textContent = ""; }

    var vh = filtered.filter(function (r) { return /very high/i.test(r.wealth); }).length;
    els.stWealth.textContent = vh.toLocaleString();
    els.stWealthSub.textContent = filtered.length ? Math.round(vh / filtered.length * 100) + "% of shown" : "";
  }

  // ---- Selection sync ---------------------------------------------------
  function selectRecord(id, fromTable) {
    selectedId = id;
    Array.prototype.forEach.call(els.tbody.querySelectorAll("tr"), function (tr) {
      tr.classList.toggle("sel", tr.dataset.id === String(id));
    });
    var r = records[id];
    Object.keys(markersById).forEach(function (mid) {
      var sel = mid === String(id);
      markersById[mid].setStyle({ radius: sel ? 10 : 6.5, color: sel ? "#1b2433" : "#ffffff", weight: sel ? 2.5 : 1.5 });
      if (sel) markersById[mid].bringToFront();
    });
    if (r && r.hasGeo && markersById[id]) {
      map.flyTo([r.lat, r.lng], Math.max(map.getZoom(), 14), { duration: 0.6 });
      markersById[id].openPopup();
    }
    var tr = els.tbody.querySelector('tr[data-id="' + id + '"]');
    if (tr) tr.scrollIntoView({ block: "nearest", behavior: fromTable ? "smooth" : "auto" });
  }

  // Clear the current selection: shrink every marker back to normal and drop
  // the highlighted table row. Called when a prospect's card is dismissed.
  function deselect() {
    selectedId = null;
    Object.keys(markersById).forEach(function (mid) {
      markersById[mid].setStyle({ radius: 6.5, color: "#ffffff", weight: 1.5 });
    });
    Array.prototype.forEach.call(els.tbody.querySelectorAll("tr.sel"), function (tr) {
      tr.classList.remove("sel");
    });
  }

  // ---- Excel export -----------------------------------------------------
  // Columns mirror the source workbook (Role Sorted sheet), in the same order
  // and with the same header text, so the export drops in like the original.
  var EXPORT_COLUMNS = [
    ["Name", "name"],
    ["Role", "role"],
    ["Company", "company"],
    ["Property Address", "address"],
    ["Sale Price", "salePrice"],
    ["Sale Date", "saleDate"],
    ["Distance to MGH (km)", "distanceMGH"],
    ["Prestige Score (0-100)", "prestige"],
    ["Seniority Tier", "tier"],
    ["Wealth Signal", "wealth"],
    ["Donor Category", "category"],
    ["Notes", "notes"]
  ];

  function exportExcel() {
    if (typeof XLSX === "undefined") {
      toast("Excel library didn't load — check your connection and retry.", true);
      return;
    }
    var list = sortRecords(applyFilters()); // exactly what the table shows
    if (!list.length) { toast("No prospects match the filters — nothing to export.", true); return; }

    var aoa = [EXPORT_COLUMNS.map(function (c) { return c[0]; })];
    list.forEach(function (r) {
      aoa.push(EXPORT_COLUMNS.map(function (c) {
        var v = r[c[1]];
        return (v === null || v === undefined) ? "" : v;
      }));
    });

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [22, 24, 26, 30, 12, 12, 18, 20, 18, 14, 16, 32].map(function (w) { return { wch: w }; });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Prospects");

    var d = new Date();
    var stamp = d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
    var base = ((window.PROSPECT_META && window.PROSPECT_META.source) || "prospects").replace(/\.[^.]+$/, "");
    XLSX.writeFile(wb, base + "_filtered_" + list.length + "_" + stamp + ".xlsx");
    toast("Exported " + list.length + " prospect" + (list.length === 1 ? "" : "s") + " to Excel.");
  }

  // ---- Toast ------------------------------------------------------------
  var toastTimer;
  function toast(msg, isErr) {
    els.toast.textContent = msg;
    els.toast.className = "toast show" + (isErr ? " err" : "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { els.toast.className = "toast"; }, 3600);
  }

  // ---- Wiring -----------------------------------------------------------
  function debounce(fn, ms) { var t; return function () { clearTimeout(t); t = setTimeout(fn, ms); }; }
  var rerender = function () { render(false); };
  var rerenderDeb = debounce(rerender, 160);

  // Frame-throttled render for sliders. Unlike a debounce, a continuous drag
  // can't starve it — at most one render is queued per animation frame, so the
  // map/list update live and smoothly while you drag.
  var rafPending = false;
  function rerenderRAF() {
    if (rafPending) return;
    rafPending = true;
    requestAnimationFrame(function () { rafPending = false; render(false); });
  }

  function wire() {
    els.exportBtn.addEventListener("click", exportExcel);
    els.fName.addEventListener("input", rerenderDeb);
    els.fCategory.addEventListener("change", rerender);
    els.fTier.addEventListener("change", rerender);
    els.fWealth.addEventListener("change", rerender);
    // Sliders: instant label on every move, live frame-throttled filtering
    // while dragging, and a guaranteed final render on release (change).
    els.fPrice.addEventListener("input", function () { els.fPriceOut.textContent = fmtMoneyShort(parseFloat(els.fPrice.value) || 0); rerenderRAF(); });
    els.fPrice.addEventListener("change", rerender);
    els.fPrestige.addEventListener("input", function () { els.fPrestigeOut.textContent = els.fPrestige.value; rerenderRAF(); });
    els.fPrestige.addEventListener("change", rerender);

    els.resetBtn.addEventListener("click", function () {
      els.fName.value = ""; els.fCategory.value = ""; els.fTier.value = ""; els.fWealth.value = "";
      els.fPrice.value = 0; els.fPriceOut.textContent = "$0";
      els.fPrestige.value = 0; els.fPrestigeOut.textContent = "0";
      render(true);
    });

    els.sortField.addEventListener("change", rerender);
    els.sortDir.addEventListener("click", function () {
      sortDir = sortDir === "desc" ? "asc" : "desc";
      els.sortDir.style.transform = sortDir === "asc" ? "rotate(180deg)" : "";
      rerender();
    });
    Array.prototype.forEach.call(document.querySelectorAll("th.sortable"), function (th) {
      th.addEventListener("click", function () {
        var f = th.dataset.sort;
        if (els.sortField.value === f) { sortDir = sortDir === "desc" ? "asc" : "desc"; }
        else { els.sortField.value = f; sortDir = "desc"; }
        els.sortDir.style.transform = sortDir === "asc" ? "rotate(180deg)" : "";
        rerender();
      });
    });
  }

  // ---- Go ---------------------------------------------------------------
  initMap();
  wire();
  if (Array.isArray(window.PROSPECT_DATA)) {
    ingest(window.PROSPECT_DATA);
  } else {
    els.emptyTitle.textContent = "No data found";
    els.emptyMsg.textContent = "data.js is missing or empty. Run build_data.py to regenerate it from the Excel file.";
  }
})();
