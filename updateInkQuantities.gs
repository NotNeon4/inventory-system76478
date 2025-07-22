function updateInksQuantities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Inks Log");
  const trackerSheet = ss.getSheetByName("Inks Tracker");

  if (!logSheet || !trackerSheet) {
    Logger.log("❌ Missing required sheet(s).");
    return;
  }

  const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues(); // Skip header
  const trackerData = trackerSheet.getRange(2, 1, trackerSheet.getLastRow() - 1, 1).getValues(); // Just ID

  const quantityMap = {};

  logData.forEach(([timestamp, id, action, qty]) => {
    if (!id || !qty || !action) return;

    const quantity = Number(qty);
    if (!quantityMap[id]) quantityMap[id] = 0;

    if (action.toLowerCase().includes("add (to stock)")) {
      quantityMap[id] += quantity;
    } else if (action.toLowerCase().includes("take (from stock)")) {
      quantityMap[id] -= quantity;
    }
  });

  // Update the Inks Tracker quantities
  trackerData.forEach((row, i) => {
    const id = row[0];
    const updatedQty = quantityMap[id] ?? 0;
    trackerSheet.getRange(i + 2, 3).setValue(updatedQty); // Column C = Quantity
  });

  Logger.log("✅ Ink quantities updated.");
}
