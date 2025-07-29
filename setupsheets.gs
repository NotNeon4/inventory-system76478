function setupUrgentStockSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup "Config" sheet
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
  } else {
    configSheet.clear(); // Clear existing content
  }

  // Define headers based on your final screenshot for Config sheet (A-J, then N for Default Order Emails)
  const headers = [
    "Item ID",         // A (0)
    "Urgent Threshold",// B (1)
    "Urgent Comparison",// C (2)
    "Daily Threshold", // D (3)
    "Daily Comparison",// E (4)
    "Notify Type",     // F (5)
    "Emails",          // G (6)
    "On Order Flag",   // H (7)
    "Last Urgent Sent",// I (8)
    "Notes",           // J (9)
    "",                // K (10) - empty
    "",                // L (11) - empty
    "",                // M (12) - empty
    "Default Order Emails" // N (13)
  ];

  // Set values for the header row
  configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  configSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9e1f2");

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
