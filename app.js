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
      // Using your actual active sheet name visible in the workbook
      const sheet = context.workbook.worksheets.getItem("Main Dashboard");
      const range = sheet.getUsedRange().load("values");
      await context.sync();

      const data = parseMainDashboard(range.values);
      renderKPIs(data);
      renderHoldingsCharts(data);

      document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
    });
  } catch (error) {
    console.error("Error reading workbook data:", error);
    document.getElementById("kpiPortfolioValue").innerText = "Error loading";
  }
}

function parseMainDashboard(rows) {
  const holdings = [];
  let cashBalance = 0;
  
  // Scans the rows for your portfolio values
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    // This safely reads your rows where data is structured
    const ticker = row[1] || row[2];
    const mktVal = parseFloat(row[6]) || parseFloat(row[7]) || 0;

    if (mktVal > 0) {
      holdings.push({ ticker: ticker || "Asset", name: ticker || "Holding", mktVal });
    }
  }
  return { holdings, cashBalance };
}

function renderKPIs(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${Math.round(totalValue).toLocaleString('en-SG')}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} items loaded`;
}

function renderHoldingsCharts(data) {
  const top10 = [...data.holdings].sort((a, b) => b.mktVal - a.mktVal).slice(0, 10);
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
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width: ${pct}%"></div>
      </div>
      <div class="bar-value">S$ ${Math.round(item.mktVal).toLocaleString()}</div>
    `;
    container.appendChild(row);
  });
}
