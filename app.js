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
      const range = sheet.getUsedRange().load("values");
      await context.sync();

      const data = parsePortfolioData(range.values);
      renderDashboard(data);

      document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
    });
  } catch (error) {
    console.error("Error reading workbook data:", error);
  }
}

function parsePortfolioData(rows) {
  const holdings = [];
  let cashBalance = 599856; // Fallback or parsed from summary row if dynamic

  // Iterate past headers (assuming rows start around index 4 based on your sheet setup)
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    const ticker = row[1];
    const name = row[2];
    const shares = parseFloat(row[3]) || 0;
    const cost = parseFloat(row[4]) || 0;
    const price = parseFloat(row[5]) || 0;
    const mktVal = parseFloat(row[6]) || 0;
    const country = row[7];
    const type = row[8];
    const unrealisedGL = parseFloat(row[9]) || 0;
    const dividends = parseFloat(row[10]) || 0;

    if (ticker && mktVal > 0) {
      holdings.push({ ticker, name, shares, cost, price, mktVal, country, type, unrealisedGL, dividends });
    }
  }

  return { holdings, cashBalance };
}

function renderDashboard(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  const totalUnrealised = data.holdings.reduce((sum, h) => sum + h.unrealisedGL, 0);
  const totalCost = data.holdings.reduce((sum, h) => sum + h.cost, 0);
  const unrealisedPct = totalCost > 0 ? ((totalUnrealised / totalCost) * 100).toFixed(1) : 0;

  // 1. KPI Top Cards
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${Math.round(totalValue).toLocaleString('en-SG')}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings · S$ ${data.cashBalance.toLocaleString()} cash on the side`;
  
  document.getElementById("kpiUnrealisedReturn").innerText = `+S$ ${Math.round(totalUnrealised).toLocaleString('en-SG')}`;
  document.getElementById("kpiUnrealisedSub").innerText = `+${unrealisedPct}% on cost · capital + dividends since purchase`;

  // 2. Holdings Section KPIs
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

  renderTop10Bars(top10.slice(0, 10), totalValue);
}

function renderTop10Bars(top10, totalValue) {
  const container = document.getElementById("top10HoldingsBars");
  if (!container) return;
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
