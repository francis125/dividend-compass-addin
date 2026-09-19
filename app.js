async function fetchExcelData() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Portfolio Overview");
    // Pull a broad range to guarantee all rows are captured
    const range = sheet.getRange("A1:Q200").load("values");
    await context.sync();

    const rows = range.values;
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

    console.log(`Successfully parsed ${holdings.length} holdings out of ${rows.length} rows.`);

    renderKPIs({ holdings, cashBalance });
    renderHoldingsCharts({ holdings, cashBalance });

    document.getElementById("lastPulled").innerText = `Pulled ${new Date().toLocaleTimeString()}`;
  }).catch((error) => {
    console.error("Excel data read error:", error);
  });
}
