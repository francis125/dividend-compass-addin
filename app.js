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
    const poRange = poSheet.getUsedRange().load("values");
    await context.sync();

    const poData = parsePortfolioOverview(poRange.values);
    renderKPIs(poData);
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
        <span class="holding-name">${item.name}</span>
        <span class="holding-ticker">${item.ticker}</span>
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width: ${pct}%"></div>
      </div>
      <div class="bar-value">S$ ${Math.round(item.mktVal).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });

  if (top10.length > 0) {
    const largest = top10[0];
    const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
    const largestPct = totalValue > 0 ? ((largest.mktVal / totalValue) * 100).toFixed(1) : 0;
    
    document.getElementById("holdingsLargestPct").innerText = `${largestPct}%`;
    document.getElementById("holdingsLargestName").innerText = largest.name;

    const top5Sum = top10.slice(0, 5).reduce((sum, h) => sum + h.mktVal, 0);
    const top5Pct = totalValue > 0 ? ((top5Sum / totalValue) * 100).toFixed(1) : 0;
    document.getElementById("holdingsTop5Pct").innerText = `${top5Pct}%`;
  }
}
