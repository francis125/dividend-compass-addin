async function fetchExcelData() {
  await Excel.run(async (context) => {
    // 1. Target Main Dashboard sheet where your data resides
    const mainSheet = context.workbook.worksheets.getItem("Main Dashboard");
    
    // Read the used range of Main Dashboard
    const mainRange = mainSheet.getUsedRange().load("values");
    await context.sync();

    // 2. Parse Main Dashboard Data (Matching your workbook layout)
    const rows = mainRange.values;
    const poData = parseMainDashboard(rows);

    // 3. Render Dashboard Components
    renderKPIs(poData);
    renderHoldingsCharts(poData);
    
    document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric'})}`;
  }).catch((error) => {
    console.error("Error reading workbook data:", error);
  });
}

function parseMainDashboard(rows) {
  // Extract summary metrics from Main Dashboard top section
  const portfolioValue = parseFloat(rows[4]?.[2]) || 0; // Row 5, Col C
  const unrealisedReturn = parseFloat(rows[4]?.[5]) || 0; // Row 5, Col F
  const annualisedDiv = parseFloat(rows[7]?.[2]) || 0; // Row 8, Col C
  
  const totalPositions = parseInt(rows[12]?.[2]) || 0; // Row 13, Col C
  const largestHoldingPct = parseFloat(rows[12]?.[5]) || 0; // Row 13, Col F
  const top5Concentration = parseFloat(rows[12]?.[8]) || 0; // Row 13, Col I

  const holdings = [];
  
  // Parse holdings table starting around row 16 onwards
  for (let i = 15; i < rows.length; i++) {
    const row = rows[i];
    const name = row[2]; // Col C
    const ticker = row[5]; // Col F
    const mktVal = parseFloat(row[7]) || 0; // Col H
    
    if (ticker && mktVal > 0) {
      holdings.push({ ticker, name, mktVal });
    }
  }

  return {
    portfolioValue,
    unrealisedReturn,
    annualisedDiv,
    totalPositions,
    largestHoldingPct,
    top5Concentration,
    holdings,
    cashBalance: 0
  };
}

function renderKPIs(data) {
  document.getElementById("kpiPortfolioValue").innerText = `S$ ${data.portfolioValue.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  document.getElementById("kpiHoldingsCount").innerText = `${data.totalPositions} holdings · S$0 cash on the side`;
  document.getElementById("kpiUnrealisedReturn").innerText = `+S$ ${data.unrealisedReturn.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  document.getElementById("kpiAnnualisedDiv").innerText = `S$ ${data.annualisedDiv.toLocaleString('en-SG', { maximumFractionDigits: 0 })}`;
  
  document.getElementById("holdingsTotalPos").innerText = data.totalPositions;
  document.getElementById("holdingsLargestPct").innerText = `${(data.largestHoldingPct * 100).toFixed(1)}%`;
  document.getElementById("holdingsTop5Pct").innerText = `${(data.top5Concentration * 100).toFixed(1)}%`;
}
