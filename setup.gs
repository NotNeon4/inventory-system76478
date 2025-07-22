function setupOrderSheets() {
  const ss = SpreadsheetApp.getActive();

  // ====== ORDERS SHEET ======
  let ordersSheet = ss.getSheetByName("Orders");
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet("Orders");
  } else {
    ordersSheet.clear(); // Optional: clear existing data
  }

  ordersSheet.getRange("A1:D1").setValues([[
    "Order Number", "Timestamp", "Ordered By", "Status"
  ]]);

  // ====== ORDER ITEMS SHEET ======
  let orderItemsSheet = ss.getSheetByName("Order Items");
  if (!orderItemsSheet) {
    orderItemsSheet = ss.insertSheet("Order Items");
  } else {
    orderItemsSheet.clear(); // Optional: clear existing data
  }

  orderItemsSheet.getRange("A1:C1").setValues([[
    "Order Number", "Item ID", "Quantity"
  ]]);

  SpreadsheetApp.flush();
  Logger.log("Orders and Order Items sheets initialized.");
}
