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
    // Explicitly grab A1:N35 to guarantee we capture all tickers, values, and columns
    const poRange = poSheet.getRange("A1:N35").load("values");
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
  let cashBalance = 599856; 

  // Loop dynamically through every row returned in the A1:N35 range
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    const ticker = row[1] ? String(row[1]).trim() : ""; // Column B
    const name = row[2] ? String(row[2]).trim() : "";   // Column C
    
    // Stop if we hit empty rows, headers, or the footer summary row
    if (!ticker || ticker === "Ticker" || ticker.startsWith("SGD")) continue;

    const price = formalismNumber(row[3]); // Column D
    const shares = formalismNumber(row[4]); // Column E
    const cost = formalismNumber(row[5]);   // Column F
    const mktVal = formalismNumber(row[6]); // Column G (Mkt Value SGD)
    const country = row[7] || "SG";         // Column H
    const type = row[8] || "Stock";         // Column I
    const dividends = formalismNumber(row[12]); // Column M (Dividend SGD)

    if (mktVal > 0) {
      const totalCostVal = shares * cost;
      const unrealisedGL = mktVal - totalCostVal;
      holdings.push({ ticker, name, shares, cost: totalCostVal, price, mktVal, country, type, unrealisedGL, dividends });
    }
  }

  return { holdings, cashBalance };
}

function formalismNumber(val) {
  return parseFloat(String(val).replace(/,/g, '')) || 0;
}

function renderKPIs(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${totalValue.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.holdings.length} holdings · S$ ${data.cashBalance.toLocaleString()} cash on the side`;
}

function renderHoldingsCharts(data) {
  // Sort descending by market value and strictly limit to top 10 items
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
      <div class="bar-label-group">
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
