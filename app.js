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
  await Excel.run(async (context) => {
    const poSheet = context.workbook.worksheets.getItem("Portfolio Overview");
    const unrlSheet = context.workbook.worksheets.getItem("Unrealized P&L Dashboard");
    const divSheet = context.workbook.worksheets.getItem("Dividend Dashboard");

    const poRange = poSheet.getUsedRange().load("values");
    const unrlRange = unrlSheet.getUsedRange().load("values");
    const divRange = divSheet.getUsedRange().load("values");
    
    await context.sync();

    const poData = parsePortfolioOverview(poRange.values);
    const metrics = parseDashboardMetrics(unrlRange.values, divRange.values, poData);
    
    renderKPIs(poData, metrics);
    renderHoldingsCharts(poData);

    document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
  }).catch((error) => {
    console.error("Error reading workbook data:", error);
  });
}

function parsePortfolioOverview(rows) {
  const holdings = [];
  let cashBalance = 0;
  
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    const ticker = row[1];
    const name = row[2];
    const mktVal = parseFloat(row[6]) || 0;
    const country = row[7];
    const type = row[8];

    if (ticker && mktVal > 0) {
      holdings.push({ ticker, name, mktVal, country, type });
    }
  }
  return { holdings, cashBalance };
}

function parseDashboardMetrics(unrlRows, divRows, poData) {
  let unrealisedGain = 0;
  let costBasis = 0;
  let annualDiv = 0;

  // Scan Unrealized P&L sheet for totals if available
  try {
    for (let r = 0; r < unrlRows.length; r++) {
      for (let c = 0; c < unrlRows[r].length; c++) {
        if (unrlRows[r][c] === "Total Unrealised P&L" || unrlRows[r][c] === "Total Gain / Loss") {
          unrealisedGain = parseFloat(unrlRows[r][c + 2]) || 0;
        }
      }
    }
  } catch (e) {
    console.warn("Could not auto-parse unrealised summary:", e);
  }

  const totalValue = poData.holdings.reduce((sum, h) => sum + h.mktVal, 0) + poData.cashBalance;
  const returnPct = costBasis > 0 ? (unrealisedGain / costBasis) * 100 : 0;
  
  return {
    unrealisedGain,
    returnPct,
    annualDiv,
    currentYield: totalValue > 0 ? (annualDiv / totalValue) * 100 : 0,
    costYield: costBasis > 0 ? (annualDiv / costBasis) * 100 : 0
  };
}

function renderKPIs(data, metrics) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  
  // Portfolio Value
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${totalValue.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings · S$ ${data.cashBalance.toLocaleString()} cash on the side`;
  
  // Unrealised Return
  const sign = metrics.unrealisedGain >= 0 ? "+" : "";
  document.getElementById("kpiUnrealisedReturn").innerText = `${sign}S$ ${Math.round(metrics.unrealisedGain).toLocaleString()}`;
  document.getElementById("kpiUnrealisedReturnSub").innerText = `${metrics.returnPct.toFixed(1)}% on cost · capital + dividends`;

  // Dividend Income
  document.getElementById("kpiAnnualisedDiv").innerText = `S$ ${Math.round(metrics.annualDiv).toLocaleString()}`;

  // Yields
  document.getElementById("kpiYieldPair").innerText = `${metrics.currentYield.toFixed(2)}% / ${metrics.costYield.toFixed(2)}%`;
}

function renderHoldingsCharts(data) {
  const top10 = [...data.holdings].sort((a, b) => b.mktVal - a.mktVal).slice(0, 10);
  const container = document.getElementById("top10HoldingsBars");
  container.innerHTML = "";
  const maxVal = top10[0]?.mktVal || 1;

  top10.forEach(item => {
    const pct = (item.mktVal / maxVal) * 100;
    const row = document.createElement("div");
    row.className = "bar-row";
    row.innerHTML = `
      <div class="bar-label">
        <div class="holding-name">${item.name}</div>
        <div class="holding-ticker">${item.ticker}</div>
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width: ${pct}%"></div>
      </div>
      <div class="bar-value">S$ ${Math.round(item.mktVal).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });
}
