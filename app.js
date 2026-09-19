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
      const poRange = poSheet.getRange("B5:G30").load("values"); // Directly target exact columns B through G for rows 5 to 30
      await context.sync();

      console.log("Raw Excel Data Retrieved:", poRange.values);
      const holdings = [];

      poRange.values.forEach((row, index) => {
        const ticker = row[0] ? String(row[0]).trim() : ""; // Column B
        const name = row[1] ? String(row[1]).trim() : "";   // Column C
        const price = formalismNumber(row[2]);              // Column D
        const shares = formalismNumber(row[3]);             // Column E
        const cost = formalismNumber(row[4]);               // Column F
        const mktVal = formalismNumber(row[5]);             // Column G

        if (ticker && mktVal > 0) {
          holdings.push({ ticker, name, mktVal });
        }
      });

      console.log("Parsed Holdings Count:", holdings.length);
      renderHoldingsCharts(holdings);

      const pulledElem = document.getElementById("lastPulled");
      if (pulledElem) {
        pulledElem.innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}`;
      }
    });
  } catch (error) {
    console.error("Error reading workbook data:", error);
  }
}

function formalismNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  return parseFloat(String(val).replace(/[$SGD,\s]/g, '')) || 0;
}

function renderHoldingsCharts(holdings) {
  const container = document.getElementById("top10HoldingsBars");
  if (!container) {
    console.error("Error: #top10HoldingsBars container element not found in DOM!");
    return;
  }

  container.innerHTML = ""; 

  if (holdings.length === 0) {
    container.innerHTML = `<div style="color: #f85149; padding: 10px;">No holdings parsed! Check console logs.</div>`;
    return;
  }

  const top10 = [...holdings].sort((a, b) => b.mktVal - a.mktVal).slice(0, 10);
  const maxVal = top10[0].mktVal || 1;

  top10.forEach(item => {
    const pct = Math.max((item.mktVal / maxVal) * 100, 5);
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
