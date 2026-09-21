/* Dividend Compass — Office Add-in dashboard logic
   -------------------------------------------------------------------------
   Rendering (charts, tables, KPIs) is a faithful port of the approved
   design template. Data can come from two places:

     1. LIVE  — read out of the open workbook's "Main Dashboard" tab via
        Office.js, using the RANGES config below. Main Dashboard already
        computes every figure this page needs (Sections 01-08 built with
        plain formulas), so this page just reads finished values — it does
        not re-derive anything from Portfolio Overview.
     2. SAMPLE — the fallback data baked in below, used whenever Office.js
        isn't available (e.g. previewing this file directly in a browser)
        or a live read fails. The status badge in the header reflects
        which one is on screen.

   NOTE: the exact cell addresses in RANGES are placeholders and need to be
   confirmed against the current "Main Dashboard" tab layout (cells were
   recently realigned) before this will pull real numbers. Until then the
   page runs correctly on sample data. Update RANGES, then everything else
   below works unchanged.
   ------------------------------------------------------------------------- */

(function () {
  "use strict";

  // ===========================================================================
  // 1. CONFIG — cell/range addresses on the "Main Dashboard" sheet.
  //    Fill these in against the live workbook, then set USE_LIVE_DATA = true.
  // ===========================================================================
  var USE_LIVE_DATA = true; // flip to false to force sample data for a demo
  var SHEET_NAME = "Main Dashboard";

  var RANGES = {
    // Hero KPIs
    portfolioValue:      "C6",
    portfolioHoldingsN:  "C14",
    cashOnSide:          "C7",
    unrealisedReturn:    "F6",
    unrealisedReturnPct: "F7",
    divIncomeAnnual:     "C9",
    divIncomeGrowthPct:  "C10",
    yieldCurrent:        "F9",
    yieldOnCost:         "G9",

    // 01 Holdings — top 10 table, columns: #, Holding, Ticker, MarketValue, %Portfolio, GLAmt, GLPct
    holdingsTable:        "B17:H26",
    statTotalPositions:   "C14",
    statLargestPct:       "D14",
    statLargestName:      "D15",
    statTop5Pct:          "F14",

    // 07/08 Asset mix + geographic mix (Sections 07-08 on Main Dashboard)
    assetMixTable: "B53:D56",   // Asset Type | Allocation % | Market Value
    geoMixTable:   "B60:D68",   // Region | Mix % | Value

    // 02 Capital gain/loss
    capGainTotal:    "C31",
    capGainPct:      "C32",
    positionsUp:     "C33",
    positionsDown:   "C34",
    capitalTable:    "B36:D45", // Name | Ticker | G/L amount, top movers

    // 03 Unrealised P&L with dividend
    unrealisedTable: "B48:F58", // Name | Ticker | Capital G/L | Dividends | Total

    // 04 Realised P&L with dividend
    realisedCap:     "C41",
    realisedDiv:     "F42",
    realisedTotal:   "H41",
    realisedTable:   "B44:F54",

    // 05 Dividend growth YoY (dynamic array, last 5 years + YTD)
    growthTable:     "B47:C53", // Year | Total Dividend | (growth % computed)

    // 06 Yield on cost / current yield
    yocPortfolio:    "F47",
    curYieldPortfolio: "F48",
    yieldTable:      "B60:D69" // Holding | Yield on cost | Current yield
  };

  // ===========================================================================
  // 2. SAMPLE DATA — fallback, identical to the approved design template
  // ===========================================================================
  var SAMPLE = buildSampleData();

  // ===========================================================================
  // 3. BOOT
  // ===========================================================================
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () {
      init();
    });
  } else {
    // Standalone preview (no Office host) — just render sample data.
    document.addEventListener("DOMContentLoaded", function () {
      setStatus(false);
      renderAll(SAMPLE);
    });
  }

  function init() {
    if (!USE_LIVE_DATA || typeof Excel === "undefined") {
      setStatus(false);
      renderAll(SAMPLE);
      return;
    }
    loadLiveData()
      .then(function (data) {
        setStatus(true);
        renderAll(data);
      })
      .catch(function (err) {
        console.error("Dividend Compass: live read failed, falling back to sample data.", err);
        setStatus(false);
        renderAll(SAMPLE);
      });
  }

  function setStatus(isLive) {
    var badge = document.getElementById("status-badge");
    var updated = document.getElementById("updated-label");
    if (badge) {
      badge.textContent = isLive ? "Live · from your workbook" : "Sample data · preview";
      badge.classList.toggle("live", isLive);
    }
    if (updated) {
      updated.textContent = "Pulled " + new Date().toLocaleDateString("en-SG", { day: "numeric", month: "short", year: "numeric" }) +
        ", " + new Date().toLocaleTimeString("en-SG", { hour: "2-digit", minute: "2-digit" });
    }
  }

  // ===========================================================================
  // 4. LIVE DATA — Office.js reads from Main Dashboard
  // ===========================================================================
  function loadLiveData() {
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getItem(SHEET_NAME);
      var loaded = {};
      Object.keys(RANGES).forEach(function (key) {
        var r = sheet.getRange(RANGES[key]);
        r.load("values");
        loaded[key] = r;
      });
      return context.sync().then(function () {
        return mapLiveRangesToData(loaded);
      });
    });
  }

  function mapLiveRangesToData(r) {
    function v(key) { return r[key] && r[key].values ? r[key].values[0][0] : null; }
    function num(x) { return typeof x === "number" ? x : 0; }

    var data = JSON.parse(JSON.stringify(SAMPLE)); // start from sample shape, overwrite with live values

    data.kpi.portfolioValue = num(v("portfolioValue"));
    data.kpi.holdingsCount = num(v("statTotalPositions"));
    data.kpi.cashOnSide = num(v("cashOnSide"));
    data.kpi.unrealisedReturn = num(v("unrealisedReturn"));
    data.kpi.unrealisedReturnPct = num(v("unrealisedReturnPct"));
    data.kpi.divIncomeAnnual = num(v("divIncomeAnnual"));
    data.kpi.divIncomeGrowthPct = num(v("divIncomeGrowthPct"));
    data.kpi.yieldCurrent = num(v("yieldCurrent"));
    data.kpi.yieldOnCost = num(v("yieldOnCost"));

    data.stats.totalPositions = num(v("statTotalPositions"));
    data.stats.largestPct = num(v("statLargestPct"));
    data.stats.largestName = v("statLargestName") || data.stats.largestName;
    data.stats.top5Pct = num(v("statTop5Pct"));

    var holdingsRows = (r.holdingsTable && r.holdingsTable.values) || [];
    if (holdingsRows.length) {
      data.holdings = holdingsRows
        .filter(function (row) { return row[1]; })
        .map(function (row) {
          return { name: row[1], tkr: row[2], mv: num(row[3]), glp: num(row[4]), gl: num(row[5]), div: num(row[6]) };
        });
    }

    var assetRows = (r.assetMixTable && r.assetMixTable.values) || [];
    if (assetRows.length) {
      var catColors = ["var(--cat-1)", "var(--cat-2)", "var(--cat-3)", "var(--ink-3)"];
      data.assetMix = assetRows
        .filter(function (row) { return row[0]; })
        .map(function (row, i) {
          return { name: row[0], value: num(row[2]), color: catColors[i % catColors.length] };
        });
    }

    var geoRows = (r.geoMixTable && r.geoMixTable.values) || [];
    if (geoRows.length) {
      data.geoMix = geoRows
        .filter(function (row) { return row[0]; })
        .map(function (row) {
          return { name: row[0], code: "", value: num(row[2]) };
        });
    }

    var capRows = (r.capitalTable && r.capitalTable.values) || [];
    if (capRows.length) {
      data.capitalMovers = capRows
        .filter(function (row) { return row[0]; })
        .map(function (row) { return { name: row[0], tkr: row[1], gl: num(row[2]) }; });
    }

    var unrealRows = (r.unrealisedTable && r.unrealisedTable.values) || [];
    if (unrealRows.length) {
      data.unrealisedTotals = unrealRows
        .filter(function (row) { return row[0]; })
        .map(function (row) { return { name: row[0], tkr: row[1], cap: num(row[2]), div: num(row[3]), total: num(row[4]) }; });
    }

    var realRows = (r.realisedTable && r.realisedTable.values) || [];
    if (realRows.length) {
      data.realisedTotals = realRows
        .filter(function (row) { return row[0]; })
        .map(function (row) { return { name: row[0], date: "Closed position", cap: num(row[2]), div: num(row[3]), total: num(row[4]) }; });
    }
    data.realisedSummary.cap = num(v("realisedCap"));
    data.realisedSummary.div = num(v("realisedDiv"));
    data.realisedSummary.total = num(v("realisedTotal"));

    var yearRows = (r.growthTable && r.growthTable.values) || [];
    if (yearRows.length) {
      var years = yearRows.filter(function (row) { return row[0]; }).map(function (row) { return { y: String(row[0]), v: num(row[1]) }; });
      for (var i = 0; i < years.length; i++) {
        years[i].g = i > 0 && years[i - 1].v ? +(((years[i].v - years[i - 1].v) / years[i - 1].v) * 100).toFixed(2) : null;
      }
      if (years.length) years[years.length - 1].ytd = new Date().getMonth() < 11;
      data.years = years;
    }

    data.yieldSummary.yoc = num(v("yocPortfolio"));
    data.yieldSummary.cur = num(v("curYieldPortfolio"));

    var yieldRows = (r.yieldTable && r.yieldTable.values) || [];
    if (yieldRows.length) {
      data.yields = yieldRows
        .filter(function (row) { return row[0]; })
        .map(function (row) { return { name: row[0], yoc: num(row[1]), cur: num(row[2]) }; });
    }

    return data;
  }

  // ===========================================================================
  // 5. SAMPLE DATA BUILDER (fallback — mirrors the approved design template)
  // ===========================================================================
  function buildSampleData() {
    var holdings = [
      { name: "Vanguard Australian Shares High Yield ETF", tkr: "VHY.AX", mv: 967429, gl: 210864, glp: 26.95, div: 236929 },
      { name: "iShares UK Dividend UCITS ETF GBP (Dist)", tkr: "IUKD.L", mv: 749809, gl: 216399, glp: 41.30, div: 140107 },
      { name: "Tracker Fund Of Hong Kong", tkr: "2800.HK", mv: 372311, gl: 8776, glp: 2.29, div: 104621 },
      { name: "State Street SPDR Straits Times Index ETF", tkr: "ES3.SI", mv: 309960, gl: 163186, glp: 111.18, div: 58168 },
      { name: "Stoneweg Europe Stapled Trust", tkr: "SET.SI", mv: 207710, gl: -40247, glp: -16.00, div: 55293 },
      { name: "CapitaLand Ascendas REIT", tkr: "A17U.SI", mv: 223184, gl: -39012, glp: -14.88, div: 35176 },
      { name: "Mapletree Industrial Trust", tkr: "ME8U.SI", mv: 213963, gl: -44290, glp: -17.15, div: 39050 },
      { name: "Microsoft Corporation", tkr: "MSFT", mv: 199293, gl: 39971, glp: 24.98, div: 348 },
      { name: "CapitaLand Integrated Commercial Trust", tkr: "C38U.SI", mv: 200405, gl: 10420, glp: 5.48, div: 12727 },
      { name: "Frasers Logistics & Commercial Trust", tkr: "BUOU.SI", mv: 176580, gl: -42738, glp: -19.49, div: 31905 },
      { name: "Other 16 holdings", tkr: "—", mv: 1650781, gl: -120034, glp: -6.78, div: 138028 }
    ];

    var assetMix = [
      { name: "Reit/Trust", value: 2066082, color: "var(--cat-1)" },
      { name: "ETF", value: 2754158, color: "var(--cat-2)" },
      { name: "Stock", value: 451185, color: "var(--cat-3)" },
      { name: "Cash", value: 599856, color: "var(--ink-3)" }
    ];

    var geoMix = [
      { name: "Singapore", code: "SG", value: 1855626 },
      { name: "Australia", code: "AU", value: 963124 },
      { name: "Hong Kong / China", code: "HK · CN", value: 804866 },
      { name: "United Kingdom", code: "UK", value: 744335 },
      { name: "United States", code: "US", value: 551207 },
      { name: "Europe", code: "EU", value: 205839 },
      { name: "India", code: "IN", value: 146427 },
      { name: "Japan", code: "JP", value: 0 },
      { name: "Cash", code: "Multi-ccy", value: 599856 }
    ];

    var realisedTotals = [
      { name: "Broadcom Inc.", date: "Closed position", cap: 511816, div: 2709 },
      { name: "Manulife US REIT", date: "Closed position", cap: -77445, div: 0 },
      { name: "Bank of China Limited", date: "Closed position", cap: 44136, div: 26843 },
      { name: "China Construction Bank Corp.", date: "Closed position", cap: 31131, div: 21225 },
      { name: "BP p.l.c.", date: "Closed position", cap: 36907, div: 13347 },
      { name: "Prime US REIT", date: "Closed position", cap: -46321, div: 0 },
      { name: "Yangzijiang Shipbuilding (Holdings)", date: "Closed position", cap: 21461, div: 5600 },
      { name: "CapitaLand Ascendas REIT (past lots)", date: "Closed position", cap: 18746, div: 9111 },
      { name: "CK Asset Holdings Limited", date: "Closed position", cap: -39720, div: 21197 },
      { name: "Westpac Banking Corporation", date: "Closed position", cap: 2169, div: 7068 }
    ].map(function (r) { return Object.assign({}, r, { total: r.cap + r.div }); });

    var years = [
      { y: "2021", v: 160467, g: 42.92 },
      { y: "2022", v: 168399, g: 3.29 },
      { y: "2023", v: 172962, g: 7.80 },
      { y: "2024", v: 171419, g: 3.30 },
      { y: "2025", v: 211031, g: 26.18 },
      { y: "2026", v: 145656, g: null, ytd: true }
    ];

    var yields = [
      { name: "Stoneweg Europe Stapled Trust", yoc: 9.3, cur: 11.1 },
      { name: "Mapletree Pan Asia Commercial Trust", yoc: 6.4, cur: 6.6 },
      { name: "Digital Core REIT", yoc: 0, cur: 0, note: "Newly acquired — no distribution received yet" },
      { name: "iShares UK Dividend UCITS ETF", yoc: 6.5, cur: 4.6 },
      { name: "CapitaLand Ascott Trust", yoc: 6.3, cur: 7.4 },
      { name: "Link Real Estate Investment Trust", yoc: 6.5, cur: 6.8 },
      { name: "Vanguard Australian Shares High Yield ETF", yoc: 4.5, cur: 3.5 },
      { name: "CapitaLand Ascendas REIT", yoc: 5.3, cur: 6.2 },
      { name: "State Street SPDR STI ETF", yoc: 6.6, cur: 3.1 },
      { name: "Portfolio overall", yoc: 4.7, cur: 4.4, hl: true }
    ];

    var totalMv = holdings.reduce(function (a, h) { return a + h.mv; }, 0);
    var totalCapGl = holdings.reduce(function (a, h) { return a + h.gl; }, 0);
    var totalDivRcvd = holdings.reduce(function (a, h) { return a + h.div; }, 0);

    return {
      kpi: {
        portfolioValue: 5271425, holdingsCount: 26, cashOnSide: 599856,
        unrealisedReturn: 1215647, unrealisedReturnPct: 24.8,
        divIncomeAnnual: 248661, divIncomeGrowthPct: 17.8,
        yieldCurrent: 4.4, yieldOnCost: 4.7
      },
      stats: { totalPositions: 26, largestPct: 18.3, largestName: "Vanguard Australian Shares High Yield ETF", top5Pct: 49.5 },
      holdings: holdings,
      totalMv: totalMv, totalCapGl: totalCapGl, totalDivRcvd: totalDivRcvd,
      assetMix: assetMix,
      geoMix: geoMix,
      capitalMovers: holdings.filter(function (h) { return h.tkr !== "—"; })
        .slice().sort(function (a, b) { return Math.abs(b.gl) - Math.abs(a.gl); }).slice(0, 9)
        .map(function (h) { return { name: h.name, tkr: h.tkr, gl: h.gl }; }),
      capitalSummary: { gain: 363295, gainPct: 7.4, up: 9, down: 17, of: 26 },
      unrealisedTotals: holdings.map(function (h) { return { name: h.name, tkr: h.tkr, cap: h.gl, div: h.div, total: h.gl + h.div }; }),
      realisedTotals: realisedTotals,
      realisedSummary: { cap: 933311, div: 540550, total: 1473860 },
      years: years,
      yieldSummary: { yoc: 4.7, cur: 4.4, spreadPp: 0.3 },
      yields: yields
    };
  }

  // ===========================================================================
  // 6. RENDERING — ported 1:1 from the approved design; takes any `data` object
  //    matching the shape produced above (live or sample).
  // ===========================================================================
  var fmt = function (n) { return Math.round(Math.abs(n)).toLocaleString("en-SG"); };
  var signed = function (n) { return (n >= 0 ? "+" : "\u2212") + fmt(n); };
  var signedSgd = function (n) { return (n >= 0 ? "+S$" : "\u2212S$") + fmt(n); };

  var tip = null;
  function showTip(e, html) {
    if (!tip) return;
    tip.innerHTML = html;
    tip.classList.add("show");
    moveTip(e);
  }
  function moveTip(e) {
    if (!tip) return;
    tip.style.left = e.clientX + "px";
    tip.style.top = (e.clientY - 10) + "px";
  }
  function hideTip() { if (tip) tip.classList.remove("show"); }

  function clear(el) { while (el && el.firstChild) el.removeChild(el.firstChild); }

  function renderAll(data) {
    tip = document.getElementById("tip");

    renderKpis(data);
    renderHoldingsList(data);
    renderHoldingsTable(data);
    renderAssetMix(data);
    renderGeoMix(data);
    renderCapital(data);
