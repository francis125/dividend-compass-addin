/* Dividend Compass — Office Add-in dashboard logic
   -------------------------------------------------------------------------
   Rendering (charts, tables, KPIs) is a faithful port of the approved
   design template. Data can come from two places:

     1. LIVE  — read directly out of the workbook's own source-of-truth
        sheets via Office.js (SHEETS/RANGES config below), not from any
        pre-computed dashboard tab:
          - Portfolio Overview        — holdings, weightings, asset mix,
                                         geographic mix, cash, portfolio/
                                         cash/grand totals
          - Unrealized P&L Dashboard  — per-holding capital G/L, dividends
                                         received on open positions, and
                                         the portfolio-wide totals row
          - Realised P&L Dashboard    — per-company closed-position P/L
                                         and the grand-total closed P/L
          - Dividend Dashboard        — year-by-year dividend income
                                         (also used to derive lifetime
                                         dividends, cross-checked against
                                         the Dividend Ledger)
          - YoC Calc (Do Not Edit)    — per-holding and portfolio-level
                                         yield on cost / current yield
                                         (TTM), which this workbook already
                                         maintains as the refined source
        All five layouts were confirmed directly against an uploaded copy
        of the workbook (FrancisAcc v97), not guessed.
     2. SAMPLE — the fallback data baked in below, used whenever Office.js
        isn't available (e.g. previewing this file directly in a browser)
        or a live read fails. The status badge in the header reflects
        which one is on screen.

   If you rename a sheet, or reorder/insert columns in one of the ranges
   below, update SHEETS/RANGES to match — everything else keeps working
   unchanged since rendering is driven entirely off the data shape, not
   cell addresses.
   ------------------------------------------------------------------------- */

(function () {
  "use strict";

  // ===========================================================================
  // 1. CONFIG — sheet names + ranges on the real source sheets.
  // ===========================================================================
  var USE_LIVE_DATA = true;

  var SHEETS = {
    portfolioOverview: "Portfolio Overview",
    unrealised:        "Unrealized P&L Dashboard",
    realised:           "Realised P&L Dashboard",
    dividends:          "Dividend Dashboard",
    yocCalc:            "YoC Calc (Do Not Edit)",
    dividendLedger:      "Dividend Ledger"
  };

  var RANGES = {
    // Portfolio Overview — 26 current holdings (row 5-30)
    //   B Ticker | C Company | D Last price | E Units | F Avg cost | G Mkt value SGD |
    //   H Country | I Type | J Weightage
    poHoldings:  { sheet: "portfolioOverview", addr: "B5:J30" },
    // Portfolio Segment summary — row50 Equity, row51 Cash, row53 Grand Total
    //   G label | H Value SGD | I Weightage % | J Div % | K Div SGD
    poSummary:   { sheet: "portfolioOverview", addr: "G50:K53" },
    // Asset mix — K:Type, L:Allocation %, M:Allocation SGD (rows 57-60)
    poAssetMix:  { sheet: "portfolioOverview", addr: "K57:M60" },
    // Geographic mix (equity only) — D:Region, E:Value SGD (rows 57-64)
    poGeoMix:    { sheet: "portfolioOverview", addr: "D57:E64" },

    // Unrealized P&L Dashboard — per-holding (row 7-32)
    //   B Company | C Brokerage | D Units | E AvgCostLC | F TotalCostLC | G TotalCostSGD |
    //   H UnitCostLC | I MktValueSGD | J P/L SGD | K Return % | L DivsRcvdLC | M DivsRcvdSGD
    uHoldings:   { sheet: "unrealised", addr: "B7:M32" },
    // Grand-total row (row 160): G TotalCostSGD | I MktValueSGD | J P/L SGD | K Return % |
    //   M DivsRcvdSGD | Q TotalReturnSGD | R TotalReturn %
    uTotals:     { sheet: "unrealised", addr: "G160:R160" },

    // Realised P&L Dashboard — per-company closed positions (row 6-106)
    //   B Company | C Txns | D UnitsSold | E TotalCostSGD | F ExitCostSGD |
    //   G DivsRcvdLC | H DivsRcvdSGD | I ClosedP/LLC | J ClosedP/LSGD | K TotalReturnSGD | L Return%
    rHoldings:   { sheet: "realised", addr: "B6:L106" },
    rTotal:      { sheet: "realised", addr: "J137" }, // grand-total closed P/L SGD

    // Dividend Ledger — per-payment rows (row 4-738), used to attribute
    // dividends to CLOSED lots per company, the same way Unrealized P&L
    // Dashboard's own "Divs Rcvd" column sources its per-company total from
    // this same ledger's "Div Attributable (SGD)" column, matched by
    // Company Name (col B). Column W here is "Div Attributable Closed
    // Lots (SGD)" — the ledger already splits each payout between the
    // units still held at that date and the units since sold, so this is
    // the exact per-holding dividend figure for the Realised P&L section.
    //   B Company Name | ... (cols C-V unused here) ... | W Div Attributable Closed Lots (SGD)
    divLedger:   { sheet: "dividendLedger", addr: "B4:W738" },

    // Dividend Dashboard — year-by-year (row 7-26, 2007-2026)
    //   B Year | K Total Div SGD
    divYears:    { sheet: "dividends", addr: "B7:K26" },

    // YoC Calc (Do Not Edit) — per-holding TTM yields (row 4-29, same 26 holdings/order as Portfolio Overview)
    //   B Ticker | C Company | I Yield on Cost TTM | J Current Yield TTM
    yocHoldings: { sheet: "yocCalc", addr: "B4:J29" },
    // Portfolio-level totals block (rows 32-36)
    //   D32 Portfolio YoC | D33 Portfolio Current Yield | D34 Implied Annual Div SGD |
    //   D35 Total Mkt Value SGD | D36 Total Cost Basis SGD
    yocTotals:   { sheet: "yocCalc", addr: "D32:D36" }
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
    // Guard against the script running after DOMContentLoaded already fired
    // (e.g. a script tag at the end of the body) — in that case the event
    // listener below would never fire and the page would sit empty.
    var bootStandalone = function () {
      setStatus(false);
      renderAll(SAMPLE);
    };
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", bootStandalone);
    } else {
      bootStandalone();
    }
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
  // 4. LIVE DATA — Office.js reads straight from the source sheets
  // ===========================================================================
  function loadLiveData() {
    return Excel.run(function (context) {
      var sheetCache = {};
      function getSheet(key) {
        if (!sheetCache[key]) sheetCache[key] = context.workbook.worksheets.getItem(SHEETS[key]);
        return sheetCache[key];
      }
      var loaded = {};
      Object.keys(RANGES).forEach(function (key) {
        var cfg = RANGES[key];
        var rng = getSheet(cfg.sheet).getRange(cfg.addr);
        rng.load("values");
        loaded[key] = rng;
      });
      return context.sync().then(function () {
        return mapLiveRangesToData(loaded);
      });
    });
  }

  function num(x) { return typeof x === "number" ? x : 0; }
  function str(x) { return x === null || x === undefined ? "" : String(x).trim(); }
  function rows(r, key) { return (r[key] && r[key].values) || []; }

  function mapLiveRangesToData(r) {
    var data = {};

    // ---- Portfolio Overview: 26 current holdings ------------------------
    // B Ticker | C Company | D LastPrice | E Units | F AvgCost | G MktValueSGD | H Country | I Type | J Weightage
    var poRows = rows(r, "poHoldings")
      .filter(function (row) { return str(row[1]); })
      .map(function (row) {
        return {
          tkr: str(row[0]), name: str(row[1]), mv: num(row[5]),
          country: str(row[6]), type: str(row[7]), weight: num(row[8])
        };
      });
    var totalMv = poRows.reduce(function (a, h) { return a + h.mv; }, 0);

    // ---- Unrealized P&L Dashboard: per-holding cap G/L + dividends ------
    // B Company | ... | G TotalCostSGD | ... | I MktValueSGD | J P/L SGD | K Return% | L DivsRcvdLC | M DivsRcvdSGD
    var uRows = rows(r, "uHoldings").filter(function (row) { return str(row[0]); });
    var uByName = {};
    uRows.forEach(function (row) {
      uByName[str(row[0])] = { cost: num(row[5]), mv: num(row[7]), gl: num(row[8]), glp: num(row[9]) * 100, div: num(row[11]) };
    });
    // uTotals spans the full contiguous range G160:R160 (12 columns: G..R),
    // so every column in between is present in the row array even though we
    // only care about a few of them — index by real column offset from G.
    var uTotalsRow = rows(r, "uTotals")[0] || [];
    // G=0 TotalCostSGD | H=1 | I=2 MktValueSGD | J=3 P/L SGD | K=4 Return% | L=5 |
    // M=6 DivsRcvdSGD | N=7 | O=8 | P=9 | Q=10 TotalReturnSGD | R=11 TotalReturn%
    var uTotalCost = num(uTotalsRow[0]), uTotalPL = num(uTotalsRow[3]), uTotalPLPct = num(uTotalsRow[4]) * 100,
        uTotalDiv = num(uTotalsRow[6]), uTotalReturn = num(uTotalsRow[10]), uTotalReturnPct = num(uTotalsRow[11]) * 100;

    // ---- YoC Calc: per-holding TTM yields, row-aligned with Portfolio Overview
    var yocRows = rows(r, "yocHoldings");
    var yocByTicker = {};
    yocRows.forEach(function (row) {
      if (!str(row[0])) return;
      yocByTicker[str(row[0])] = { yoc: num(row[7]) * 100, cur: num(row[8]) * 100 };
    });
    var yocTotalsCol = rows(r, "yocTotals").map(function (row) { return num(row[0]); });
    var portfolioYoc = (yocTotalsCol[0] || 0) * 100, portfolioCurYield = (yocTotalsCol[1] || 0) * 100,
        impliedAnnualDiv = yocTotalsCol[2] || 0, yocTotalMv = yocTotalsCol[3] || totalMv;

    // ---- Merge holdings: name/ticker/weight from Portfolio Overview,
    //      G/L + dividends from Unrealized P&L Dashboard (matched by company name),
    //      yields from YoC Calc (matched by ticker) ------------------------
    var holdings = poRows.map(function (h) {
      var u = uByName[h.name] || { gl: 0, glp: 0, div: 0 };
      return { name: h.name, tkr: h.tkr, mv: h.mv, gl: u.gl, glp: u.glp, div: u.div };
    });
    var totalCapGl = uTotalPL, totalDivRcvd = uTotalDiv;

    // ---- Portfolio Overview summary block (Equity / Cash / Grand Total) --
    // Range starts at row50 itself (the Equity row), not the header above it.
    var poSummaryRows = rows(r, "poSummary"); // row0=Equity(50), row1=Cash(51), row2=blank(52), row3=GrandTotal(53)
    var equityRow = poSummaryRows[0] || [], cashRow = poSummaryRows[1] || [];
    var portfolioValue = num(equityRow[1]) || totalMv; // G=0(label) H=1(Value SGD)
    var cashOnSide = num(cashRow[1]);

    // ---- Asset mix / Geographic mix --------------------------------------
    var catColors = ["var(--cat-1)", "var(--cat-2)", "var(--cat-3)", "var(--ink-3)"];
    var assetMix = rows(r, "poAssetMix")
      .filter(function (row) { return str(row[0]); })
      .map(function (row, i) { return { name: str(row[0]), value: num(row[2]), color: catColors[i % catColors.length] }; });
    var geoMix = rows(r, "poGeoMix")
      .filter(function (row) { return str(row[0]); })
      .map(function (row) { return { name: regionName(str(row[0])), code: str(row[0]), value: num(row[1]) }; });
    // Cash is intentionally excluded from Geographic mix — that chart is
    // equity exposure by region only.

    // ---- Dividend Ledger: dividends attributable to CLOSED lots, per company
    // Realised P&L Dashboard's own G/H "Divs Rcvd" columns are never
    // populated in the workbook — but the ledger already computes this per
    // payment (col W = "Div Attributable Closed Lots (SGD)"), so sum it by
    // company name, the same SUMIFS pattern the workbook itself uses to
    // populate Unrealized P&L Dashboard's dividend column from this sheet.
    var divLedgerByCompany = {};
    rows(r, "divLedger").forEach(function (row) {
      var name = str(row[0]);
      if (!name) return;
      // addr starts at col B, so col offset 0=B(name) ... 21=W(Div Attributable Closed Lots SGD)
      divLedgerByCompany[name] = (divLedgerByCompany[name] || 0) + num(row[21]);
    });

    // ---- Realised P&L Dashboard: closed positions ------------------------
    var rRows = rows(r, "rHoldings")
      .filter(function (row) { return str(row[0]); })
      .map(function (row) {
        var name = str(row[0]);
        var cap = num(row[8]);
        var div = divLedgerByCompany[name] || 0;
        return { name: name, date: "Closed position", cap: cap, div: div, total: cap + div };
      });
    var realisedCap = num((rows(r, "rTotal")[0] || [])[0]) || rRows.reduce(function (a, x) { return a + x.cap; }, 0);

    // ---- Dividend Dashboard: year-by-year, also gives lifetime total -----
    var divYearRows = rows(r, "divYears").filter(function (row) { return str(row[0]) && row[9] !== null && row[9] !== ""; });
    var nowYear = new Date().getFullYear();
    var years = divYearRows.slice(-6).map(function (row) {
      return { y: str(row[0]), v: num(row[9]) };
    });
    for (var i = 0; i < years.length; i++) {
      years[i].g = i > 0 && years[i - 1].v ? ((years[i].v - years[i - 1].v) / years[i - 1].v) * 100 : null;
      if (parseInt(years[i].y, 10) === nowYear) { years[i].ytd = true; years[i].g = null; }
    }

    // Realised dividends collected = sum of each closed position's own
    // "Div Attributable Closed Lots" figure from the Dividend Ledger (see
    // divLedgerByCompany above) — precise per-holding, not an estimate.
    var realisedDiv = rRows.reduce(function (a, x) { return a + x.div; }, 0);
    var realisedTotal = realisedCap + realisedDiv;

    // ---- Stats: largest holding, top-5 concentration, up/down counts -----
    var byWeight = holdings.slice().sort(function (a, b) { return b.mv - a.mv; });
    var largest = byWeight[0] || { name: "", mv: 0 };
    var top5Mv = byWeight.slice(0, 5).reduce(function (a, h) { return a + h.mv; }, 0);
    var posUp = holdings.filter(function (h) { return h.gl >= 0; }).length;
    var posDown = holdings.length - posUp;

    // Growth KPI: today's annualised (TTM) dividend run-rate vs the most
    // recently COMPLETED full calendar year's total (skip the YTD year).
    var lastFullYear = years.length && years[years.length - 1].ytd ? years[years.length - 2] : years[years.length - 1];
    var divIncomeAnnual = impliedAnnualDiv || totalDivRcvd;
    data.kpi = {
      portfolioValue: portfolioValue, holdingsCount: holdings.length, cashOnSide: cashOnSide,
      unrealisedReturn: uTotalReturn, unrealisedReturnPct: uTotalReturnPct,
      divIncomeAnnual: divIncomeAnnual,
      divIncomeGrowthPct: lastFullYear ? pctChange(divIncomeAnnual, lastFullYear.v) : 0,
      yieldCurrent: portfolioCurYield, yieldOnCost: portfolioYoc
    };
    data.stats = {
      totalPositions: holdings.length,
      largestPct: portfolioValue ? largest.mv / portfolioValue * 100 : 0,
      largestName: largest.name,
      top5Pct: portfolioValue ? top5Mv / portfolioValue * 100 : 0
    };
    // The "01 Holdings" bar list/table is titled "top 10" and the render
    // code shows every entry it's given (no sorting/limiting of its own),
    // so build that top-10-plus-rollup shape here — same pattern as the
    // approved design (10 named holdings + one "Other N holdings" row).
    var byMvDesc = holdings.slice().sort(function (a, b) { return b.mv - a.mv; });
    var top10 = byMvDesc.slice(0, 10);
    var rest = byMvDesc.slice(10);
    if (rest.length) {
      var restTotals = rest.reduce(function (a, h) { a.mv += h.mv; a.gl += h.gl; a.div += h.div; return a; }, { mv: 0, gl: 0, div: 0 });
      var restCost = restTotals.mv - restTotals.gl;
      top10.push({
        name: "Other " + rest.length + " holdings", tkr: "—",
        mv: restTotals.mv, gl: restTotals.gl,
        glp: restCost ? restTotals.gl / restCost * 100 : 0,
        div: restTotals.div
      });
    }
    data.holdings = top10;
    data.totalMv = totalMv; data.totalCapGl = totalCapGl; data.totalDivRcvd = totalDivRcvd;
    data.assetMix = assetMix;
    data.geoMix = geoMix;
    data.capitalMovers = holdings.slice().sort(function (a, b) { return Math.abs(b.gl) - Math.abs(a.gl); }).slice(0, 9)
      .map(function (h) { return { name: h.name, tkr: h.tkr, gl: h.gl }; });
    data.capitalSummary = { gain: totalCapGl, gainPct: uTotalPLPct, up: posUp, down: posDown, of: holdings.length };
    data.unrealisedTotals = holdings.map(function (h) { return { name: h.name, tkr: h.tkr, cap: h.gl, div: h.div, total: h.gl + h.div }; });
    data.realisedTotals = rRows.slice().sort(function (a, b) { return Math.abs(b.total) - Math.abs(a.total); }).slice(0, 10);
    data.realisedSummary = { cap: realisedCap, div: realisedDiv, total: realisedTotal, count: rRows.length };
    data.years = years;
    data.yieldSummary = { yoc: portfolioYoc, cur: portfolioCurYield };
    data.yields = holdings.map(function (h) {
      var y = yocByTicker[h.tkr] || { yoc: 0, cur: 0 };
      return { name: h.name, yoc: y.yoc, cur: y.cur };
    }).filter(function (y) { return y.yoc || y.cur; })
      .sort(function (a, b) { return b.yoc - a.yoc; }).slice(0, 9);
    data.yields.push({ name: "Portfolio overall", yoc: portfolioYoc, cur: portfolioCurYield, hl: true });

    return data;
  }

  function pctChange(current, previous) {
    if (!previous) return 0;
    return +(((current - previous) / previous) * 100).toFixed(2);
  }

  var REGION_NAMES = { SG: "Singapore", AU: "Australia", HK: "Hong Kong / China", EU: "Europe", UK: "United Kingdom", US: "United States", IN: "India", JP: "Japan" };
  function regionName(code) { return REGION_NAMES[code] || code; }

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
      { name: "Japan", code: "JP", value: 0 }
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
  var signed = function (n) { return (n >= 0 ? "+" : "−") + fmt(n); };
  var signedSgd = function (n) { return (n >= 0 ? "+$" : "−$") + fmt(n); };
  // Every percentage on the dashboard is shown to exactly 2 decimal places.
  var pct2 = function (n) { return (n || 0).toFixed(2); };
  var signedPct2 = function (n) { return (n >= 0 ? "+" : "−") + pct2(Math.abs(n)); };

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
    renderUnrealised(data);
    renderRealised(data);
    renderGrowth(data);
    renderYield(data);

    // Some embedded WebViews — notably the one macOS desktop Excel uses to
    // host taskpanes — don't reliably reflow after this many DOM inserts
    // land in one burst (Office.js resolves the whole live-data read at
    // once, so all ten render* calls above fire back-to-back). Left alone,
    // that shows up as sections visually overlapping their neighbours
    // until something else forces a relayout (e.g. resizing the pane).
    // Forcing a synchronous reflow now, and once more on the next frame,
    // reliably clears it without any visible flash.
    forceReflow();
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(forceReflow);
  }

  function forceReflow() {
    if (!document.body) return;
    var prevDisplay = document.body.style.display;
    document.body.style.display = "none";
    void document.body.offsetHeight; // eslint-disable-line no-unused-expressions
    document.body.style.display = prevDisplay;
  }

  function renderKpis(data) {
    setText("kpi-portfolio-value", "$" + fmt(data.kpi.portfolioValue));
    setText("kpi-portfolio-sub", data.kpi.holdingsCount + " holdings · +$" + fmt(data.kpi.cashOnSide) + " cash on the side");

    var uEl = document.getElementById("kpi-unrealised");
    if (uEl) { uEl.textContent = signedSgd(data.kpi.unrealisedReturn); uEl.classList.toggle("up", data.kpi.unrealisedReturn >= 0); uEl.classList.toggle("down", data.kpi.unrealisedReturn < 0); }
    setText("kpi-unrealised-sub", signedPct2(data.kpi.unrealisedReturnPct) + "% on cost · capital + dividends since purchase");

    setText("kpi-div-income", "$" + fmt(data.kpi.divIncomeAnnual));
    setText("kpi-div-sub", signedPct2(data.kpi.divIncomeGrowthPct) + "% vs prior full year");

    setText("kpi-yield", pct2(data.kpi.yieldCurrent) + "% / " + pct2(data.kpi.yieldOnCost) + "%");

    setText("stat-total-positions", data.stats.totalPositions);
    setText("stat-largest-pct", pct2(data.stats.largestPct) + "%");
    setText("stat-largest-name", data.stats.largestName);
    setText("stat-top5", pct2(data.stats.top5Pct) + "%");
  }

  function setText(id, text) {
    var el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function renderHoldingsList(data) {
    var el = document.getElementById("chart-holdings");
    if (!el) return;
    clear(el);
    var max = Math.max.apply(null, data.holdings.map(function (h) { return h.mv; }));
    data.holdings.forEach(function (h) {
      var pct = pct2(h.mv / data.totalMv * 100);
      var row = document.createElement("div");
      row.className = "row";
      row.innerHTML =
        '<div class="name">' + h.name + "<small>" + h.tkr + "</small></div>" +
        '<div class="track"><div class="fill num" style="width:' + (h.mv / max * 100) + '%"></div></div>' +
        '<div class="val num">$' + fmt(h.mv) + "</div>";
      var fillEl = row.querySelector(".fill");
      fillEl.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + h.name + "</b>" + pct + "% of portfolio &middot; $" + fmt(h.mv)); });
      fillEl.addEventListener("mouseleave", hideTip);
      el.appendChild(row);
    });
  }

  function renderHoldingsTable(data) {
    var t = document.getElementById("table-holdings");
    if (!t) return;
    var thead = "<thead><tr><th>Holding</th><th>Ticker</th><th>Market value</th><th>% Portfolio</th><th>Unrealised G/L</th></tr></thead>";
    var rows = data.holdings.map(function (h) {
      var pct = pct2(h.mv / data.totalMv * 100);
      var cls = h.gl >= 0 ? "up" : "down";
      return "<tr><td>" + h.name + '</td><td class="tkr">' + h.tkr + '</td><td class="num">$' + fmt(h.mv) + '</td><td class="num">' + pct + '%</td>' +
        '<td class="num ' + cls + '">' + signedSgd(h.gl) + " (" + signedPct2(h.glp) + "%)</td></tr>";
    }).join("");
    var totalRow = '<tr class="total"><td>Total</td><td></td><td class="num">$' + fmt(data.totalMv) + '</td><td class="num">100%</td><td class="num up">' + signedSgd(data.totalCapGl) + "</td></tr>";
    t.innerHTML = thead + "<tbody>" + rows + totalRow + "</tbody>";
  }

  function renderAssetMix(data) {
    var bar = document.getElementById("chart-assetmix");
    var legend = document.getElementById("legend-assetmix");
    if (!bar || !legend) return;
    clear(bar); clear(legend);
    var total = data.assetMix.reduce(function (a, d) { return a + d.value; }, 0);
    data.assetMix.forEach(function (d) {
      var pct = (d.value / total * 100);
      var seg = document.createElement("div");
      seg.className = "seg";
      seg.style.width = pct + "%";
      seg.style.background = d.color;
      seg.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + d.name + "</b>" + pct2(pct) + "% &middot; $" + fmt(d.value)); });
      seg.addEventListener("mouseleave", hideTip);
      bar.appendChild(seg);

      var li = document.createElement("li");
      li.innerHTML = '<i style="background:' + d.color + '"></i><span class="lname">' + d.name + '</span><span class="lval num">' + pct2(pct) + "% &middot; $" + fmt(d.value) + "</span>";
      legend.appendChild(li);
    });
  }

  function renderGeoMix(data) {
    var el = document.getElementById("chart-geomix");
    if (!el) return;
    clear(el);
    var total = data.geoMix.reduce(function (a, d) { return a + d.value; }, 0);
    var max = Math.max.apply(null, data.geoMix.map(function (d) { return d.value; }));
    data.geoMix.slice().sort(function (a, b) { return b.value - a.value; }).forEach(function (d) {
      var pct = (d.value / total * 100);
      var row = document.createElement("div");
      row.className = "row";
      row.innerHTML =
        '<div class="name">' + d.name + "<small>" + d.code + "</small></div>" +
        '<div class="track"><div class="fill num" style="width:' + (max ? d.value / max * 100 : 0) + '%; background:var(--accent);"></div></div>' +
        '<div class="val num">' + (d.value ? pct2(pct) + "%" : "—") + "</div>";
      var fillEl = row.querySelector(".fill");
      fillEl.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + d.name + "</b>" + (d.value ? pct2(pct) + "% · $" + fmt(d.value) : "No current exposure")); });
      fillEl.addEventListener("mouseleave", hideTip);
      el.appendChild(row);
    });
  }

  function renderDiverging(elId, list, valueFn, subFn) {
    var el = document.getElementById(elId);
    if (!el) return;
    clear(el);
    if (!list.length) return;
    var max = Math.max.apply(null, list.map(function (d) { return Math.abs(valueFn(d)); }));
    list.forEach(function (d) {
      var v = valueFn(d);
      var isGain = v >= 0;
      // Half-width (0-50%) on either side of the centre axis at 50%.
      var halfW = (max ? Math.abs(v) / max * 50 : 0);
      var fillLeft = isGain ? "50%" : (50 - halfW) + "%";
      var fillWidth = halfW + "%";
      var row = document.createElement("div");
      row.className = "row";
      row.innerHTML =
        '<div class="name">' + d.name + (subFn ? "<small>" + subFn(d) + "</small>" : "") + "</div>" +
        '<div class="dtrack">' +
          '<div class="axis"></div>' +
          '<div class="dfill ' + (isGain ? "gain" : "loss") + ' num" style="left:' + fillLeft + '; width:' + fillWidth + ';"></div>' +
        "</div>" +
        '<div class="val num ' + (isGain ? "up" : "down") + '">' + signedSgd(v) + "</div>";
      var fillEl = row.querySelector(".dfill");
      if (fillEl) {
        fillEl.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + d.name + "</b>" + (isGain ? "Gain: " : "Loss: ") + signedSgd(v)); });
        fillEl.addEventListener("mouseleave", hideTip);
      }
      el.appendChild(row);
    });
  }

  function renderCapital(data) {
    setText("stat-cap-gain", signedSgd(data.capitalSummary.gain));
    var el = document.getElementById("stat-cap-gain"); if (el) el.classList.add(data.capitalSummary.gain >= 0 ? "up" : "down");
    setText("stat-cap-gain-pct", signedPct2(data.capitalSummary.gainPct) + "% on cost");
    setText("stat-pos-up", data.capitalSummary.up);
    setText("stat-pos-up-of", "of " + data.capitalSummary.of + " holdings");
    setText("stat-pos-down", data.capitalSummary.down);
    setText("stat-pos-down-of", "of " + data.capitalSummary.of + " holdings");

    renderDiverging("chart-capital", data.capitalMovers, function (h) { return h.gl; }, function (h) { return h.tkr; });
  }

  function renderUnrealised(data) {
    setText("stat-u-cap", signedSgd(data.totalCapGl));
    setText("stat-u-div", "+$" + fmt(data.totalDivRcvd));
    var total = data.totalCapGl + data.totalDivRcvd;
    setText("stat-u-total", signedSgd(total));
    setText("stat-u-total-pct", signedPct2(data.kpi.unrealisedReturnPct) + "% on cost · since each position was purchased");

    var top10 = data.unrealisedTotals.slice().sort(function (a, b) { return Math.abs(b.total) - Math.abs(a.total); }).slice(0, 10);
    renderDiverging("chart-unrealised", top10, function (h) { return h.total; }, function (h) { return h.tkr; });

    var t = document.getElementById("table-unrealised");
    if (t) {
      var thead = "<thead><tr><th>Holding</th><th>Ticker</th><th>Capital G/L</th><th>Dividends</th><th>Total return</th></tr></thead>";
      var rows = data.unrealisedTotals.map(function (h) {
        return "<tr><td>" + h.name + '</td><td class="tkr">' + h.tkr + '</td>' +
          '<td class="num ' + (h.cap >= 0 ? "up" : "down") + '">' + signedSgd(h.cap) + "</td>" +
          '<td class="num" style="color:var(--div);">+$' + fmt(h.div) + "</td>" +
          '<td class="num ' + (h.total >= 0 ? "up" : "down") + '">' + signedSgd(h.total) + "</td></tr>";
      }).join("");
      var totalRow = '<tr class="total"><td>Total</td><td></td><td class="num up">' + signedSgd(data.totalCapGl) + '</td><td class="num" style="color:var(--div);">+$' + fmt(data.totalDivRcvd) + '</td><td class="num up">' + signedSgd(total) + "</td></tr>";
      t.innerHTML = thead + "<tbody>" + rows + totalRow + "</tbody>";
    }
  }

  function renderRealised(data) {
    setText("stat-r-cap", signedSgd(data.realisedSummary.cap));
    setText("stat-r-div", "+$" + fmt(data.realisedSummary.div));
    setText("stat-r-total", signedSgd(data.realisedSummary.total));

    renderDiverging("chart-realised", data.realisedTotals, function (r) { return r.total; }, function (r) { return r.date; });

    var t = document.getElementById("table-realised");
    if (t) {
      var thead = "<thead><tr><th>Holding</th><th>Closed</th><th>Capital G/L</th><th>Dividends</th><th>Total return</th></tr></thead>";
      var rows = data.realisedTotals.map(function (r) {
        return "<tr><td>" + r.name + '</td><td class="tkr">' + r.date.replace("Closed ", "") + '</td>' +
          '<td class="num ' + (r.cap >= 0 ? "up" : "down") + '">' + signedSgd(r.cap) + "</td>" +
          '<td class="num" style="color:var(--div);">+$' + fmt(r.div) + "</td>" +
          '<td class="num ' + (r.total >= 0 ? "up" : "down") + '">' + signedSgd(r.total) + "</td></tr>";
      }).join("");
      var totalRow = '<tr class="total"><td>All closed positions</td><td></td><td class="num up">' + signedSgd(data.realisedSummary.cap) + '</td><td class="num" style="color:var(--div);">+$' + fmt(data.realisedSummary.div) + '</td><td class="num up">' + signedSgd(data.realisedSummary.total) + "</td></tr>";
      t.innerHTML = thead + "<tbody>" + rows + totalRow + "</tbody>";
    }
  }

  function renderGrowth(data) {
    var years = data.years;
    if (years.length >= 2) {
      var first = years[0], last = years[years.length - (years[years.length - 1].ytd ? 2 : 1)];
      var n = years.filter(function (y) { return !y.ytd; }).length - 1;
      var cagr = n > 0 && first.v ? (Math.pow(last.v / first.v, 1 / n) - 1) * 100 : 0;
      setText("cagr-label", n + "-year CAGR (" + first.y + "–" + last.y + ")");
      setText("stat-cagr", signedPct2(cagr) + "%");
    }
    var ytdYear = years[years.length - 1];
    if (ytdYear && ytdYear.ytd) {
      setText("ytd-label", ytdYear.y + " so far");
      setText("stat-ytd", "$" + fmt(ytdYear.v));
      var prior = years[years.length - 2];
      setText("stat-ytd-sub", prior ? "through this year · full-year " + prior.y + " was $" + fmt(prior.v) : "year to date");
    } else {
      setText("ytd-label", "Latest year");
      setText("stat-ytd", ytdYear ? "$" + fmt(ytdYear.v) : "—");
    }

    var el = document.getElementById("chart-growth");
    if (!el) return;
    clear(el);
    var max = Math.max.apply(null, years.map(function (y) { return y.v; }));
    years.forEach(function (y) {
      var h = (y.v / max * 160);
      var col = document.createElement("div");
      col.className = "yr";
      col.innerHTML =
        (y.g != null ? '<div class="yr-growth">' + signedPct2(y.g) + "%</div>" : (y.ytd ? '<div class="yr-growth" style="color:var(--ink-3);">YTD</div>' : '<div class="yr-growth">&nbsp;</div>')) +
        '<div class="yr-val num">$' + fmt(y.v) + "</div>" +
        '<div class="yr-bar num' + (y.ytd ? " ytd" : "") + '" style="height:' + h + 'px"></div>' +
        '<div class="yr-label">' + y.y + (y.ytd ? "<small>year to date</small>" : "") + "</div>";
      var bar = col.querySelector(".yr-bar");
      bar.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + y.y + (y.ytd ? " (year to date)" : "") + "</b>$" + fmt(y.v) + (y.g != null ? " &middot; " + signedPct2(y.g) + "% vs prior year" : "")); });
      bar.addEventListener("mouseleave", hideTip);
      el.appendChild(col);
    });
  }

  function renderYield(data) {
    setText("stat-yoc", pct2(data.yieldSummary.yoc) + "%");
    setText("stat-cur-yield", pct2(data.yieldSummary.cur) + "%");
    var spread = data.yieldSummary.yoc - data.yieldSummary.cur;
    var spreadEl = document.getElementById("stat-spread");
    if (spreadEl) { spreadEl.textContent = (spread >= 0 ? "+" : "−") + pct2(Math.abs(spread)) + "pp"; spreadEl.classList.toggle("up", spread >= 0); spreadEl.classList.toggle("down", spread < 0); }
    setText("stat-spread-sub", spread >= 0 ? "on cost still beats current — a positive spread across holdings" : "current yield now beats yield on cost");

    var el = document.getElementById("chart-yield");
    if (!el) return;
    clear(el);
    var max = 12; // scale 0-12%
    data.yields.forEach(function (y) {
      var row = document.createElement("div");
      row.className = "row" + (y.hl ? " highlight" : "");
      var p1 = (y.yoc / max * 100), p2 = (y.cur / max * 100);
      var lo = Math.min(p1, p2), hi = Math.max(p1, p2);
      row.innerHTML =
        '<div class="name">' + y.name + (y.note ? ' <span title="' + y.note + '" style="color:var(--ink-3); font-weight:400; font-size:11px; cursor:help;">&nbsp;(' + y.note + ')</span>' : "") + "</div>" +
        '<div class="dtrack2">' +
        '<div class="dline" style="left:' + lo + '%; width:' + (hi - lo) + '%;"></div>' +
        '<div class="dot cost" style="left:calc(' + p1 + '% - 6px)"></div>' +
        '<div class="dot cur" style="left:calc(' + p2 + '% - 6px)"></div>' +
        "</div>" +
        '<div class="dumb-vals num">' + pct2(y.yoc) + "% <b>/</b> " + pct2(y.cur) + "%</div>";
      var costDot = row.querySelector(".dot.cost"), curDot = row.querySelector(".dot.cur");
      costDot.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + y.name + "</b>Yield on cost: " + pct2(y.yoc) + "%"); });
      costDot.addEventListener("mouseleave", hideTip);
      curDot.addEventListener("mousemove", function (e) { moveTip(e); showTip(e, "<b>" + y.name + "</b>Current yield: " + pct2(y.cur) + "%"); });
      curDot.addEventListener("mouseleave", hideTip);
      el.appendChild(row);
    });
  }
})();
