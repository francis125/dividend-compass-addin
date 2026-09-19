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
      const poSheet = context.workbook.worksheets.getItem("Portfolio Overview");
      const poRange = poSheet.getRange("A1:N35").load("values");
      await context.sync();

      const poData = parsePortfolioOverview(poRange.values);
      renderKPIs(poData);
      renderHoldingsCharts(poData);
      
      const pulledElem = document.getElementById("lastPulled");
      if (pulledElem) {
        pulledElem.innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
      }
    });
  } catch (error) {
    console.error("Error reading workbook data:", error);
  }
}

function parsePortfolioOverview(rows) {
  const holdings = [];
  let cashBalance = 599856; 

  // Loop through rows 5 onwards (array index 4 is row 5)
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    const ticker = row[1] ? String(row[1]).trim() : "";
    const name = row[2] ? String(row[2]).trim() : "";
    
    // Stop if we hit empty rows, headers, or totals
    if (!ticker || ticker === "Ticker" || ticker.toUpperCase().startsWith("SGD")) break;

    const price = formalismNumber(row[3]);      
    const shares = formalismNumber(row[4]);     
    const cost = formalismNumber(row[5]);       
    const mktVal = formalismNumber(row[6]);     
    const country = row[7] || "SG";             
    const type = row[8] || "Stock";             
    const dividends = formalismNumber(row[12]); 

    if (mktVal > 0) {
      const totalCostVal = shares * cost;
      const unrealisedGL = mktVal - totalCostVal;
      holdings.push({ 
        ticker, 
        name, 
        shares, 
        cost: totalCostVal, 
        price, 
        mktVal, 
        country, 
        type, 
        unrealisedGL, 
        dividends 
      });
    }
  }

  return { holdings, cashBalance };
}

function formalismNumber(val) {
  if (val === undefined || val === null) return 0;
  return parseFloat(String(val).replace(/,/g, '')) || 0;
}

function renderKPIs(data) {
  const totalValue = data.holdings.reduce((sum, h) => sum + h.mktVal, 0) + data.cashBalance;
  const valElem = document.getElementById("kpiPortfolioValue");
  const countElem = document.getElementById("kpiHoldingsCount");
  if (valElem) valElem.innerText = `S$ ${Math.round(totalValue).toLocaleString('en-SG')}`;
  if (countElem) countElem.innerText = `${data.holdings.length} holdings · S$ ${data.cashBalance.toLocaleString()} cash on the side`;
}

function renderHoldingsCharts(data) {
  const container = document.getElementById("top10HoldingsBars");
  if (!container) return;

  container.innerHTML = ""; 

  // Sort descending by market value and take top 10
  const top10 = [...data.holdings].sort((a, b) => b.mktVal - a.mktVal).slice(0, 10);
  if (top10.length === 0) return;

  const maxVal = top10[0].mktVal || 1;

  top10.forEach(item => {
    const pct = Math.max((item.mktVal / maxVal) * 100, 2); // ensures tiny bars are still visible
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
