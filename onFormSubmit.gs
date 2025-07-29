function onFormSubmit(e) {
    Logger.log("Form submission detected. Running stock checks.");
  runStockChecks();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Stock Tracker");
  const formSheet = ss.getSheetByName("Materials Log");

  if (!stockSheet || !formSheet) {
    Logger.log("❌ Missing required sheet(s).");
    return;
  }

  const stockData = stockSheet.getDataRange().getValues();
const stockItemIds = stockData.slice(1).map(row => row[0]);
  const formData = formSheet.getDataRange().getValues();
  const formItemIds = formData.slice(1).map(row => row[1].trim()).filter(id => id);
  const uniqueFormItemIds = [...new Set(formItemIds)];

  let addedCount = 0;
  uniqueFormItemIds.forEach(id => {
    if (!stockItemIds.includes(id)) {
      stockSheet.appendRow([id]);
      addedCount++;
      Logger.log(`✅ Added: ${id}`);
    }
  });

  if (addedCount > 0) {
    const lastRow = stockSheet.getLastRow();
    const itemId = stockSheet.getRange(lastRow, 1).getValue();
    if (itemId) {
      generateLabelViaSlides(lastRow, true);  // ⬅️ pass 'true' for triggered context
    }
  } else {
    Logger.log("ℹ️ No new items to add.");
  }
}
