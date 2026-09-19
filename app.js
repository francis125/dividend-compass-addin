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
    const sheet = context.workbook.worksheets.getItem("Portfolio Overview");
    const range = sheet.getUsedRange().load("values, rowCount, columnCount");
    await context.sync();

    const rows = range.values;
    console.log(`Loaded used range: ${range.rowCount} rows, ${range.columnCount} cols`);

    let holdings = [];
    let cashBalance = 0;
    let headerIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      const rowText = rows[i].join(" ").toLowerCase();
      if (rowText.includes("ticker") || rowText.includes("symbol")) {
        headerIndex = i;
        break;
      }
    }

    const startRow = headerIndex !== -1 ? headerIndex + 1 : 4;

    for (let i = startRow; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const ticker = row[1] ? String(row[1]).trim() : "";
      const name = row[2] ? String(row[2]).trim() : ticker;
      const mktVal = parseFloat(row[6]) || 0;
      const country = row[7] ? String(row[7]).trim() : "SG";
      const type = row[8] ? String(row[8]).trim() : "Stock";

      if (ticker && mktVal > 0 && !ticker.toLowerCase().includes("total")) {
        holdings.push({ ticker, name, mktVal, country, type });
      }
    }

    renderKPIs({ holdings, cashBalance });
    renderHoldingsCharts({ holdings, cashBalance });

    document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleTimeString()}`;
  }).catch((error) => {
    console.error("Excel data read error:", error);
  });
}

function renderKPIs(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${Math.round(totalValue).toLocaleString()}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings loaded`;
  document.getElementById("holdingsTotalPos").innerText = data.holdings.length;
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
