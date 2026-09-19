function parsePortfolioData(rows) {
  const holdings = [];
  let cashBalance = 599856;

  // Equity rows run from index 3 to 28 (Excel rows 5 to 30)
  for (let i = 3; i <= 28; i++) {
    const row = rows[i];
    if (!row) continue;
    
    const ticker = row[1]; // Col C (Ticker / type indicator e.g. .L, .G)
    const name = row[2];   // Col D (Company name)
    const shares = parseFloat(row[4]) || 0; // Col F (Units)
    const cost = parseFloat(row[5]) || 0;   // Col G (Avg Cost)
    const price = parseFloat(row[6]) || 0;  // Col H (Last Price)
    const mktVal = parseFloat(row[7]) || 0; // Col I (Mkt Value SGD)
    const country = row[8] || "SG";         // Col J (Country)
    const type = row[9] || "Stock";         // Col K (Type)
    const dividends = parseFloat(row[11]) || 0; // Col M (Div SG$)

    if (ticker && mktVal > 0) {
      const totalCostVal = shares * cost;
      const unrealisedGL = mktVal - totalCostVal;
      holdings.push({ ticker, name, shares, cost: totalCostVal, price, mktVal, country, type, unrealisedGL, dividends });
    }
  }

  return { holdings, cashBalance };
}
