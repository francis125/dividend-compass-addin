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
    // Target the correct worksheet
    const poSheet = context.workbook.worksheets.getItem("Portfolio Overview");
    const poRange = poSheet.getUsedRange().load("values");
    await context.sync();

    // Parse the portfolio data rows dynamically
    const poData = parsePortfolioOverview(poRange.values);

    // Render components
    renderKPIs(poData);
    renderHoldingsCharts(poData);

    document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
  }).catch((error) => {
    console.error("Error reading workbook data:", error);
  });
}

function parsePortfolioOverview(rows) {
  let holdings = [];
  let cashBalance = 0;
  let headerRowIndex = -1;

  // Scan dynamically to find the header row containing 'ticker' or 'symbol'
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row.some(cell => typeof cell === 'string' && (cell.toLowerCase().includes('ticker') || cell.toLowerCase().includes('symbol')))) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) return { holdings, cashBalance };

  // Parse data rows below the header
  for (let i = headerRowIndex + 1; i < rows.length; i++) {
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

function renderKPIs(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${totalValue.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings · S$ ${data.cashBalance.toLocaleString()} cash on the side`;
  document.getElementById("holdingsTotalPos").innerText = data.holdings.length;
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
