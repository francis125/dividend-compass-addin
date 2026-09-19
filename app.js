function parsePortfolioData(rows) {
  const holdings = [];
  let cashBalance = 599856;

  // Loop through your equity rows
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    
    const ticker = row[1]; // Col C
    const name = row[2];   // Col D
    const shares = parseFloat(row[4]) || 0; // Col F
    const cost = parseFloat(row[5]) || 0;   // Col G
    const price = parseFloat(row[6]) || 0;  // Col H
    const mktVal = parseFloat(row[7]) || 0; // Col I
    const country = row[8] || "SG";         // Col J
    const type = row[9] || "Stock";         // Col K
    const dividends = parseFloat(row[11]) || 0; // Col M

    if (ticker && mktVal > 0) {
      const totalCostVal = shares * cost;
      const unrealisedGL = mktVal - totalCostVal;
      holdings.push({ ticker, name, shares, cost: totalCostVal, price, mktVal, country, type, unrealisedGL, dividends });
    }
  }

  return { holdings, cashBalance };
}

function renderHoldingsCharts(data) {
    // Sort descending by market value and take the exact top 10
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
            <div class="bar-value">S$${Math.round(item.mktVal).toLocaleString()}</div>
        `;
        container.appendChild(row);
    });
}
