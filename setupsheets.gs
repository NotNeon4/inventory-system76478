function setupUrgentStockSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup "Config" sheet
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
  } else {
    configSheet.clear(); // Clear existing content
  }

  configSheet.getRange("A1:F1").setValues([[
    "Item ID",         // A
    "Urgent Threshold",// B
    "Daily Threshold", // C
    "Notify Type",     // D: urgent / daily / both
    "Last Urgent Sent",// E
    "Notes"            // F: optional
  ]]);

  configSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#d9e1f2");

  // Setup "Urgent Queue" sheet
  let queueSheet = ss.getSheetByName("Urgent Queue");
  if (!queueSheet) {
    queueSheet = ss.insertSheet("Urgent Queue");
  } else {
    queueSheet.clear();
  }

  queueSheet.getRange("A1:D1").setValues([[
    "Item ID",     // A
    "Quantity",    // B
    "Threshold",   // C
    "Timestamp"    // D
  ]]);

  queueSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#fce4d6");

  SpreadsheetApp.getUi().alert("✅ Setup complete! Sheets 'Config' and 'Urgent Queue' are ready.");
}
