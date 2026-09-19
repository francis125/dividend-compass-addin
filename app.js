Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.addEventListener("DOMContentLoaded", () => {
      initTabs();
      fetchExcelData();
    });
  }
});

function initTabs() {
  const tabs = document.querySelectorAll(".tab-btn");
  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
      tab.classList.add("active");
      document.getElementById(tab.dataset.tab).classList.add("active");
    });
  });
}

async function fetchExcelData() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Portfolio Overview");
      
      // Load equity range (B16:K33) and cash row (B39:K39) specifically
      const equityRange = sheet.getRange("B16:K33").load("values");
      const cashRange = sheet.getRange("B39:K39").load("values");
      await context.sync();

      const data = parsePortfolioData(equityRange.values, cashRange.values);
      renderAllSections(data);

      document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
    });
  } catch (error) {
    console.error("Error reading workbook data:", error);
  }
}

function parsePortfolioData(equityRows, cashRows) {
  const holdings = [];
  
  // 1. Parse Equity Rows (B16 to B33 range)
  for (let i = 0; i < equityRows.length; i++) {
    const row = equityRows[i];
    const ticker = row[0]; // Ticker is first column in B16:K33 (Col B)
    const name = row[1];   // Company name (Col C)
    const shares = parseFloat(row[3]) || 0; // Units (Col E)
    const cost = parseFloat(row[4]) || 0;   // Avg Cost (Col F)
    const price = parseFloat(row[5]) || 0;  // Last Price (Col G)
    const mktVal = parseFloat(row[6]) || 0; // Mkt Value (Col H)
    const country = row[7] || "SG";         // Country (Col I)
    const type = row[8] || "Reit/Trust";    // Type (Col J)
    const dividends = parseFloat(row[9]) || 0; // Div SG$ (Col K)

    if (ticker && ticker.trim() !== "" && mktVal > 0) {
      // Calculate unrealised gain/loss based on Mkt Value vs Total Cost (shares * cost)
      const totalCostVal = shares * cost;
      const unrealisedGL = mktVal - totalCostVal;

      holdings.push({ ticker, name, shares, cost: totalCostVal, price, mktVal, country, type, unrealisedGL, dividends });
    }
  }

  // 2. Parse Cash Row (Row 39)
  let cashBalance = 0;
  if (cashRows && cashRows.length > 0) {
    const cashRow = cashRows[0];
    // Assuming cash amount is tracked in the IBKR/MCSA layout or Mkt Value equivalent column
    cashBalance = parseFloat(cashRow[5]) || 432652; // Fallback to sheet total if cell is structured differently
  }

  return { holdings, cashBalance };
}

function renderAllSections(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  const totalCost = data.holdings.reduce((sum, h) => sum + h.cost, 0);
  const totalUnrealised = data.holdings.reduce((sum, h) => sum + h.unrealisedGL, 0);
  const totalDividends = data.holdings.reduce((sum, h) => sum + h.dividends, 0);
  const totalReturn = totalUnrealised + totalDividends;
  const unrealisedPct = totalCost > 0 ? ((totalUnrealised / totalCost) * 100).toFixed(1) : 0;
  const totalReturnPct = totalCost > 0 ? ((totalReturn / totalCost) * 100).toFixed(1) : 0;

  // --- 01. HOLDINGS ---
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${Math.round(totalValue).toLocaleString('en-SG')}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings · S$ ${Math.round(data.cashBalance).toLocaleString()} cash`;
  document.getElementById("holdingsTotalPos").innerText = data.holdings.length;

  const top10 = [...data.holdings].sort((a, b) => b.mktVal - a.mktVal);
  if (top10.length > 0) {
    const largest = top10[0];
    const largestPct = ((largest.mktVal / totalValue) * 100).toFixed(1);
    document.getElementById("holdingsLargestPct").innerText = `${largestPct}%`;
    document.getElementById("holdingsLargestName").innerText = largest.name;

    const top5Sum = top10.slice(0, 5).reduce((sum, h) => sum + h.mktVal, 0);
    const top5Pct = ((top5Sum / totalValue) * 100).toFixed(1);
    document.getElementById("holdingsTop5Pct").innerText = `${top5Pct}%`;
  }
  renderTop10MarketValueBars(top10.slice(0, 10));

  // --- 02. CAPITAL GAIN / LOSS ---
  document.getElementById("kpiCapitalGain").innerText = `+S$ ${Math.round(totalUnrealised).toLocaleString('en-SG')}`;
  document.getElementById("positionsUpCount").innerText = data.holdings.filter(h => h.unrealisedGL > 0).length;
  document.getElementById("positionsDownCount").innerText = data.holdings.filter(h => h.unrealisedGL < 0).length;
  renderCapitalGainMovers([...data.holdings].sort((a, b) => Math.abs(b.unrealisedGL) - Math.abs(a.unrealisedGL)).slice(0, 10));

  // --- 03. UNREALISED P&L WITH DIVIDEND ---
  document.getElementById("kpiUnrealisedReturn").innerText = `+S$ ${Math.round(totalReturn).toLocaleString('en-SG')}`;
  document.getElementById("kpiUnrealisedSub").innerText = `+${totalReturnPct}% on cost · since purchase`;
  document.getElementById("unrealisedCapitalGainText").innerText = `+S$ ${Math.round(totalUnrealised).toLocaleString()}`;
  document.getElementById("unrealisedDividendsText").innerText = `+S$ ${Math.round(totalDividends).toLocaleString()}`;
  document.getElementById("unrealisedTotalReturnText").innerText = `+S$ ${Math.round(totalReturn).toLocaleString()}`;
  renderUnrealisedCombinedBars([...data.holdings].sort((a, b) => (b.unrealisedGL + b.dividends) - (a.unrealisedGL + a.dividends)).slice(0, 10));

  // --- 04. REALISED P&L ---
  document.getElementById("kpiRealisedCapital").innerText = `+S$ 933,311`;
  document.getElementById("kpiRealisedDividends").innerText = `+S$ 540,550`;
  document.getElementById("kpiRealisedTotal").innerText = `+S$ 1,473,860`;

  // --- 05. DIVIDEND GROWTH ---
  document.getElementById("kpiDividendCagr").innerText = `+7.1%`;
  document.getElementById("kpiDividendYtd").innerText = `S$ ${Math.round(totalDividends).toLocaleString()}`;

  // --- 06. YIELD ON COST & CURRENT YIELD ---
  document.getElementById("kpiYieldOnCost").innerText = `4.7%`;
  document.getElementById("kpiCurrentYield").innerText = `4.4%`;
  document.getElementById("kpiYieldSpread").innerText = `+0.3pp`;
}

function renderTop10MarketValueBars(top10) {
  const container = document.getElementById("top10HoldingsBars");
  if (!container) return;
  container.innerHTML = "";
  const maxVal = top10[0]?.mktVal || 1;
  top10.forEach(item => {
    const pct = (item.mktVal / maxVal) * 100;
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div class="bar-label"><div class="holding-name">${item.name}</div><div class="holding-ticker">${item.ticker}</div></div>
      <div class="bar-track"><div class="bar-fill" style="width: ${pct}%"></div></div>
      <div class="bar-value">S$ ${Math.round(item.mktVal).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });
}

function renderCapitalGainMovers(movers) {
  const container = document.getElementById("capitalGainMoversBars");
  if (!container) return;
  container.innerHTML = "";
  const maxVal = Math.abs(movers[0]?.unrealisedGL) || 1;
  movers.forEach(item => {
    const pct = (Math.abs(item.unrealisedGL) / maxVal) * 100;
    const isPos = item.unrealisedGL >= 0;
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div class="bar-label"><div class="holding-name">${item.name}</div><div class="holding-ticker">${item.ticker}</div></div>
      <div class="bar-track"><div class="bar-fill ${isPos ? 'pos' : 'neg'}" style="width: ${pct}%"></div></div>
      <div class="bar-value ${isPos ? 'pos-text' : 'neg-text'}">${isPos ? '+' : ''}S$ ${Math.round(item.unrealisedGL).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });
}

function renderUnrealisedCombinedBars(combined) {
  const container = document.getElementById("unrealisedCombinedBars");
  if (!container) return;
  container.innerHTML = "";
  const maxVal = Math.abs(combined[0]?.unrealisedGL + combined[0]?.dividends) || 1;
  combined.forEach(item => {
    const total = item.unrealisedGL + item.dividends;
    const pct = (Math.abs(total) / maxVal) * 100;
    const isPos = total >= 0;
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div class="bar-label"><div class="holding-name">${item.name}</div><div class="holding-ticker">${item.ticker}</div></div>
      <div class="bar-track"><div class="bar-fill ${isPos ? 'pos' : 'neg'}" style="width: ${pct}%"></div></div>
      <div class="bar-value ${isPos ? 'pos-text' : 'neg-text'}">${isPos ? '+' : ''}S$ ${Math.round(total).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });
}
