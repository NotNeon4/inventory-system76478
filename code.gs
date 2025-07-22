/**
 * @OnlyCurrentDoc
 */

function doGet(e) {
  // Check if a 'page' parameter is present in the URL
  if (e.parameter.page === 'config') {
    return HtmlService.createHtmlOutputFromFile('config')
      .setTitle('Jigsaw Inventory Config')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (e.parameter.page === 'orderHistory') {
    return HtmlService.createHtmlOutputFromFile('orderHistory')
      .setTitle('Jigsaw Order History')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    // Default to index.html if no page parameter or a different one
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Jigsaw Inventory Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Returns the URL for the Order History page.
 */
function getOrderHistoryFileUrl() {
  const url = ScriptApp.getService().getUrl();
  return url + '?page=orderHistory';
}

/**
 * Returns the URL for the main Index page.
 */
function getIndexFileUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Returns the URL for the Config page.
 */
function getConfigFileUrl() {
  const url = ScriptApp.getService().getUrl();
  return url + '?page=config';
}

/**
 * Simple test function for pinging the server.
 */
function pingTest() {
  return "pong";
}

/**
 * Retrieves all Item IDs and Item Names from the "Product info" sheet.
 * Assumes "Item ID" is in Column A and "Item Name" is in Column H.
 * This data is used to populate the datalist in index.html.
 * @returns {Array<Object>} An array of objects, each with 'id' and 'name' properties.
 */
function getAllItemDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productInfoSheet = ss.getSheetByName("Product info");

  if (!productInfoSheet) {
    Logger.log("Error: 'Product info' sheet not found in getAllItemDetails.");
    throw new Error("Required sheet 'Product info' not found.");
  }

  const lastRow = productInfoSheet.getLastRow();
  if (lastRow < 2) { // Only header row or empty
    Logger.log("No product data found in 'Product info' sheet.");
    return [];
  }

  // Get data from Column A (Item ID) and Column H (Item Name)
  const itemIdsRange = productInfoSheet.getRange(2, 1, lastRow - 1, 1); // Column A, from row 2
  const itemNamesRange = productInfoSheet.getRange(2, 8, lastRow - 1, 1); // Column H, from row 2
  
  const itemIds = itemIdsRange.getValues();
  const itemNames = itemNamesRange.getValues();

  const itemDetails = [];
  for (let i = 0; i < itemIds.length; i++) {
    const id = itemIds[i][0] ? itemIds[i][0].toString().trim() : '';
    const name = itemNames[i][0] ? itemNames[i][0].toString().trim() : '';

    if (id) { // Only add if Item ID exists
      itemDetails.push({ id: id, name: name });
    }
  }

  Logger.log(`Found ${itemDetails.length} item details from 'Product info' sheet.`);
  return itemDetails;
}

// Modify getAllItemIDs to simply use the new getAllItemDetails (for compatibility if still used)
function getAllItemIDs() {
  const itemDetails = getAllItemDetails();
  return itemDetails.map(item => item.id); // Still returns just IDs for compatibility if needed elsewhere
}

/**
 * Helper function to get an item's name by its ID.
 * This is primarily for server-side lookup when submitting the order.
 * @param {string} itemId The ID of the item to look up.
 * @returns {string|null} The item name, or null if not found.
 */
function getItemNameById(itemId) { // This function is generally not used by updated submitOrder, but kept for clarity
  const itemDetails = getAllItemDetails();
  const foundItem = itemDetails.find(item => item.id === itemId);
  return foundItem ? foundItem.name : null;
}

/**
 * Submits a new order to the 'Orders' and 'Order items' sheets,
 * and sends an email notification.
 * @param {Object} orderData - An object containing all order details:
 * - {Array<Object>} items: Array of item objects {itemId: string, itemName: string, qty: number, itemNotes: string}.
 * - {Array<string>} emails: Array of recipient email addresses.
 * - {string} orderNotes: Optional notes for the entire order.
 * @returns {Object} An object with success status and order number.
 */
function submitOrder(orderData) { // IMPORTANT: Now accepts a single 'orderData' object
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  const orderItemsSheet = ss.getSheetByName("Order items");
  // Stock Tracker sheet is no longer used for updates in this function
  const configSheet = ss.getSheetByName("Config"); // Used for config values (thresholds, default emails etc)
  const productInfoSheet = ss.getSheetByName("Product info"); // Used to get item names for email template

  // Defensive check for sheet existence
  // Removed stockTrackerSheet from this check as it's not directly modified here
  if (!orderSheet || !orderItemsSheet || !configSheet || !productInfoSheet) {
    Logger.log("Error: One or more required sheets (Orders, Order items, Config, Product info) not found in submitOrder.");
    throw new Error("One or more required sheets not found.");
  }

  // Extract data from the single orderData object
  const items = orderData.items;
  const recipientEmails = orderData.emails;
  const orderNotes = orderData.orderNotes || ""; // Ensure orderNotes is a string, default to empty

  Logger.log("submitOrder: Function started.");
  Logger.log("submitOrder: Recipient Emails received: " + JSON.stringify(recipientEmails));
  Logger.log("submitOrder: Items received: " + JSON.stringify(items));
  Logger.log("submitOrder: Order Notes received: '" + orderNotes + "'");

  // Ensure recipientEmails is an array and clean it
  if (!Array.isArray(recipientEmails)) {
    Logger.log("Error: submitOrder: recipientEmails is not an array. Value: " + JSON.stringify(recipientEmails));
    throw new Error("Invalid email recipients format. Expected an array of emails.");
  }
  const validRecipientEmails = recipientEmails.filter(email => email && typeof email === 'string' && email.trim() !== '');
  if (validRecipientEmails.length === 0) {
      Logger.log("Warning: submitOrder: No valid recipient emails after server-side filter. Cannot send order email.");
  }

  // --- Generate sequential Order Number ---
  let nextOrderNumVal = 1;
  const lastRow = orderSheet.getLastRow();
  if (lastRow > 1) {
    const lastOrderNumString = orderSheet.getRange(lastRow, 1).getValue(); // Assuming Order # is in Column A
    const match = lastOrderNumString.toString().match(/ORD-(\d+)/);
    if (match) {
      nextOrderNumVal = parseInt(match[1]) + 1;
    }
  }
  const newOrderNum = "ORD-" + Utilities.formatString("%04d", nextOrderNumVal); // Format as ORD-0001
  Logger.log(`Generated new order number: ${newOrderNum}`);
  // --- END ORDER NUMBER GENERATION ---

  const timestamp = new Date();
  const status = "Pending";

  // --- Prepare Product Info Map for Item Name lookup (for email template) ---
  const productInfoMap = new Map(); // Map: lowercased_item_id -> item_name
  const productInfoData = productInfoSheet.getDataRange().getValues();
  if (productInfoData.length > 1) { // Assuming headers are in row 1
      const pHeaders = productInfoData[0];
      const pIdCol = pHeaders.indexOf("Item ID"); // Column A
      const pNameCol = pHeaders.indexOf("Item Name"); // Column H (index 7 if A=0)

      if (pIdCol === -1 || pNameCol === -1) {
          Logger.log("Warning: 'Item ID' or 'Item Name' header not found in 'Product info' for lookup.");
      } else {
          for (let r = 1; r < productInfoData.length; r++) { // Start from row 2
              const id = productInfoData[r][pIdCol];
              const name = productInfoData[r][pNameCol];
              if (id) productInfoMap.set(id.toString().trim().toLowerCase(), name ? name.toString().trim() : 'Unknown Item');
          }
      }
  } else {
    Logger.log("Warning: 'Product info' sheet is empty or only contains headers. Item names in email may be 'Unknown Item'.");
  }
  // --- End Product Info Map ---


  // --- Prepare Item Rows for 'Order items' sheet ---
  const itemRowsForOrderItemsSheet = [];
  const itemsForEmailTemplate = []; // For sending to email template

  if (!Array.isArray(items) || items.length === 0) {
      Logger.log("No items provided in the order (items array is empty).");
      throw new Error("No items provided in the order.");
  }

  items.forEach(item => {
    const itemId = item.itemId ? item.itemId.toString().trim() : '';
    const qty = parseInt(item.qty);
    const itemNotes = item.itemNotes ? item.itemNotes.toString().trim() : '';

    if (!itemId) {
      Logger.log("Error: Item ID is empty for one of the items.");
      throw new Error("Item ID cannot be empty.");
    }
    if (isNaN(qty) || qty <= 0) {
      Logger.log(`Error: Invalid quantity (${item.qty}) for item ${itemId}.`);
      throw new Error(`Invalid quantity for item ${itemId}.`);
    }

    // Get item name from client-provided data or lookup
    const itemName = item.itemName || productInfoMap.get(itemId.toLowerCase()) || 'Unknown Item';

    // Store in "Order items" sheet (assuming Column D for Item Notes)
    // Columns: Order #, Item ID, Qty, Item Notes (A, B, C, D)
    itemRowsForOrderItemsSheet.push([newOrderNum, itemId, qty, itemNotes]); // Item Notes go in Column D

    // Prepare item for email template
    itemsForEmailTemplate.push({
      itemId: itemId,
      itemName: itemName,
      qty: qty,
      itemNotes: itemNotes // Include item notes for email
    });
  });

  // --- Update Sheets ---

  // Append new order to 'Orders' sheet
  // Columns: Order #, Timestamp, Email, Status, Order Notes (A, B, C, D, E)
  orderSheet.appendRow([newOrderNum, timestamp, validRecipientEmails.join(','), status, orderNotes]); // Order Notes go in Column E
  Logger.log(`Order #${newOrderNum} added to Orders sheet.`);

  // Append new order items to 'Order items' sheet
  if (itemRowsForOrderItemsSheet.length > 0) {
      // getRange(startRow, startColumn, numRows, numColumns)
      orderItemsSheet.getRange(orderItemsSheet.getLastRow() + 1, 1, itemRowsForOrderItemsSheet.length, itemRowsForOrderItemsSheet[0].length).setValues(itemRowsForOrderItemsSheet);
      Logger.log(`Items for order #${newOrderNum} added to Order items sheet.`);
  }

  // --- REMOVED STOCK TRACKER UPDATE LOGIC ---
  // The logic for updating `stockTrackerSheet` based on `stockChanges` has been removed.
  // The `stockChanges` object is still populated for completeness but no longer used for sheet updates.
  // The `stockTrackerSheet` reference is kept in the initial sheet checks because `getStockStatus` uses it.
  // If you later decide to re-enable stock deduction, this is where it would go.
  Logger.log("Stock deduction feature is disabled. Stock Tracker sheet was NOT updated.");

  // --- Send Confirmation Email ---
  sendOrderConfirmationEmail(validRecipientEmails, newOrderNum, itemsForEmailTemplate, orderNotes); // Pass all required data

  return { success: true, orderNum: newOrderNum }; // Return object for client-side success handler
}

/**
 * Retrieves stock status for urgent and low stock items.
 * @returns {Object} An object containing urgent and low stock items, and their thresholds.
 */
function getStockStatus() { // Modified to return thresholds for client-side display
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker");
  const configSheet = ss.getSheetByName("Config");

  if (!stockTrackerSheet || !configSheet) {
    Logger.log("Error: Required sheets 'Stock Tracker' or 'Config' not found in getStockStatus.");
    throw new Error("Required sheets 'Stock Tracker' or 'Config' not found.");
  }

  const { urgentThreshold, lowThreshold } = getStockThresholds(); // Use helper to get thresholds

  const lastRow = stockTrackerSheet.getLastRow();
  const urgentItems = [];
  const lowItems = [];

  if (lastRow > 1) { // If there's data beyond headers
    const stockValues = stockTrackerSheet.getDataRange().getValues();
    const headers = stockValues[0];
    const dataRows = stockValues.slice(1);

    const itemIdCol = headers.indexOf("Item ID");
    const qtyCol = headers.indexOf("Qty");

    if (itemIdCol === -1 || qtyCol === -1) {
      Logger.log("Error: Missing 'Item ID' or 'Qty' header in 'Stock Tracker' sheet for getStockStatus.");
      throw new Error("Missing 'Item ID' or 'Qty' header in 'Stock Tracker' sheet.");
    }

    dataRows.forEach(row => {
      const itemId = row[itemIdCol];
      const qty = row[qtyCol];
      if (typeof itemId === 'string' && itemId.trim() !== '' && typeof qty === 'number') {
        if (qty <= urgentThreshold) {
          urgentItems.push({ ItemID: itemId, Qty: qty });
        } else if (qty <= lowThreshold) {
          lowItems.push({ ItemID: itemId, Qty: qty });
        }
      }
    });
  }

  return { urgent: urgentItems, low: lowItems, urgentThreshold: urgentThreshold, lowThreshold: lowThreshold };
}

/**
 * Retrieves the urgent and low stock thresholds from the Config sheet.
 * @returns {Object} An object containing urgentThreshold and lowThreshold.
 */
function getStockThresholds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found in getStockThresholds.");
    throw new Error("Required sheet 'Config' not found.");
  }

  const configValues = configSheet.getDataRange().getValues();
  let urgentThreshold = 0; // Default
  let lowThreshold = 5;    // Default

  const configHeaders = configValues[0];
  const configData = configValues.slice(1);

  const keyCol = configHeaders.indexOf("Setting"); // Assuming a 'Setting' column for config values
  const valueCol = configHeaders.indexOf("Value"); // Assuming a 'Value' column for config values

  if (keyCol > -1 && valueCol > -1) {
    configData.forEach(row => {
      const setting = row[keyCol];
      const value = row[valueCol];
      if (setting === "Urgent Threshold" && !isNaN(parseInt(value))) {
        urgentThreshold = parseInt(value);
      } else if (setting === "Low Threshold" && !isNaN(parseInt(value))) {
        lowThreshold = parseInt(value);
      }
    });
  } else {
      Logger.log("Config sheet does not have 'Setting' and 'Value' columns for thresholds in getStockThresholds. Using defaults.");
  }

  return {
    urgentThreshold: urgentThreshold,
    lowThreshold: lowThreshold
  };
}

/**
 * Retrieves a list of default recipient emails from the 'Config' sheet.
 * Assumes 'Default Emails' are stored as a comma-separated string in cell K2.
 * @returns {Array<string>} An array of default email addresses.
 */
function getDefaultEmails() {
    Logger.log("Attempting to load default emails from Config sheet.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
        Logger.log("Error: Config sheet not found in getDefaultEmails.");
        return [];
    }

    const defaultEmailsRaw = configSheet.getRange("K2").getValue(); // Assuming K2 for default emails
    Logger.log("Raw default emails from K2: " + defaultEmailsRaw);

    let defaultEmails = [];
    if (typeof defaultEmailsRaw === 'string' && defaultEmailsRaw.trim() !== '') {
        defaultEmails = defaultEmailsRaw.split(',').map(e => e.trim()).filter(Boolean);
    } else {
        Logger.log("K2 is empty or not a string. No default emails to load.");
    }

    Logger.log("Default emails retrieved: " + JSON.stringify(defaultEmails));
    return defaultEmails;
}

/**
 * Sends an order confirmation email using a dynamic HTML template.
 * @param {string[]} recipientEmails - An array of email addresses to send the email to.
 * @param {string} orderNumber - The unique order number.
 * @param {Object[]} items - An array of item objects, each with itemId, itemName, qty, and itemNotes.
 * @param {string} orderNotes - The general notes for the order.
 */
function sendOrderConfirmationEmail(recipientEmails, orderNumber, items, orderNotes) {
  if (!recipientEmails || recipientEmails.length === 0) {
    Logger.log("No recipient emails provided for order confirmation. Skipping email.");
    return;
  }

  const template = HtmlService.createTemplateFromFile('EmailTemplate');
  template.orderNumber = orderNumber;
  template.items = items;
  template.orderNotes = orderNotes; // Pass orderNotes to template

  const htmlBody = template.evaluate().getContent();

  try {
    MailApp.sendEmail({
      to: recipientEmails.join(','),
      subject: `Jigsaw Inventory Portal - New Order #${orderNumber} Confirmation`,
      htmlBody: htmlBody,
    });
    Logger.log(`HTML order confirmation email for #${orderNumber} sent successfully to: ${recipientEmails.join(', ')}`);
  } catch (e) {
    Logger.log(`Error sending HTML order confirmation email for #${orderNumber}: ${e.message}`);
    throw new Error(`Failed to send order confirmation email: ${e.message}`);
  }
}

/**
 * Retrieves all orders and their associated items from the 'Orders' and 'Order items' sheets.
 * Includes Item Names where available from the 'Product info' sheet.
 * @returns {Array<Object>} An array of order objects, each containing order details and an 'items' array.
 */
function getOrdersWithItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  const orderItemsSheet = ss.getSheetByName("Order items");
  const productInfoSheet = ss.getSheetByName("Product info");

  if (!orderSheet || !orderItemsSheet || !productInfoSheet) {
    Logger.log("Error: Required sheets (Orders, Order items, Product info) not found in getOrdersWithItems.");
    throw new Error("Required sheets 'Orders', 'Order items', or 'Product info' not found.");
  }

  Logger.log("Starting getOrdersWithItems function.");

  // 1. Get Product Info details (ID and Name) for quick lookup
  const lastProductInfoRow = productInfoSheet.getLastRow();
  const itemDetailsMap = new Map(); // Map: lowercased_item_id -> item_name

  if (lastProductInfoRow > 1) {
    const productInfoValues = productInfoSheet.getDataRange().getValues();
    const productInfoHeaders = productInfoValues[0];

    const pIdCol = productInfoHeaders.indexOf("Item ID");
    const pNameCol = productInfoHeaders.indexOf("Item Name");

    if (pIdCol === -1 || pNameCol === -1) {
      Logger.log("Error: 'Item ID' or 'Item Name' columns not found in 'Product info' sheet for order history item lookup.");
      throw new Error("Missing 'Item ID' or 'Item Name' columns in 'Product info' sheet.");
    }

    for (let r = 1; r < productInfoValues.length; r++) { // Start from row 2
      const id = productInfoValues[r][pIdCol];
      const name = productInfoValues[r][pNameCol];
      if (id) {
        itemDetailsMap.set(id.toString().trim().toLowerCase(), name ? name.toString().trim() : 'Unknown Item Name');
      }
    };
  }
  Logger.log(`Loaded ${itemDetailsMap.size} item details from 'Product info' for order history display.`);

  // 2. Get all Order data
  const lastOrderRow = orderSheet.getLastRow();
  if (lastOrderRow < 2) {
    Logger.log("No order data rows found in 'Orders' sheet.");
    return [];
  }
  const orderValues = orderSheet.getDataRange().getValues();
  const orderHeaders = orderValues[0];
  const orderData = orderValues.slice(1);

  Logger.log("Found 'Orders' sheet.");
  Logger.log("Orders sheet headers: " + orderHeaders.join(', '));
  Logger.log("Orders sheet data rows count: " + orderData.length);

  const ordersMap = new Map();
  orderData.forEach(row => {
    const order = {};
    orderHeaders.forEach((header, i) => {
      // Clean header for property name (e.g., "Order #" -> "Order")
      let propName = header.replace(/[^a-zA-Z0-9\s]/g, '').trim(); // Allow spaces for matching
      
      if (header === "Order #") {
          order.orderNum = row[i];
      } else if (header === "Timestamp") {
          order.timestamp = (row[i] instanceof Date) ? row[i].toISOString() : (row[i] ? new Date(row[i]).toISOString() : null);
      } else if (header === "Email") { // Assuming "Email" (singular) is the header for requester email
          order.email = row[i] ? row[i].toString().trim() : ''; // Changed to 'email' to match frontend in orderhistory.html
      } else if (header === "Status") {
          order.status = row[i] ? row[i].toString().trim() : '';
      } else if (header === "Order Notes") { // NEW: Handle Order Notes
          order.orderNotes = row[i] ? row[i].toString().trim() : '';
      } else {
          // General catch-all for other headers, lowercased and stripped
          order[propName.replace(/\s+/g, '').toLowerCase()] = row[i]; // Remove spaces for propName consistency
      }
    });
    order.items = []; // Initialize an empty array for items
    if (order.orderNum) {
        ordersMap.set(order.orderNum, order);
    } else {
        Logger.log("Warning: Order row skipped due to missing OrderNum: " + JSON.stringify(row));
    }
  });
  Logger.log(`Orders map populated with ${ordersMap.size} entries.`);

  // 3. Get all Order Items data
  const lastOrderItemRow = orderItemsSheet.getLastRow();
  if (lastOrderItemRow > 1) {
    const orderItemValues = orderItemsSheet.getDataRange().getValues();
    const orderItemHeaders = orderItemValues[0];
    const orderItemData = orderItemValues.slice(1);

    Logger.log("Found 'Order items' sheet.");
    Logger.log("Order items sheet headers: " + orderItemHeaders.join(', '));
    Logger.log("Order items sheet data rows count: " + orderItemData.length);

    const oiOrderNumCol = orderItemHeaders.indexOf("Order #");
    const oiItemIdCol = orderItemHeaders.indexOf("Item ID");
    const oiQtyCol = orderItemHeaders.indexOf("Qty");
    const oiItemNotesCol = orderItemHeaders.indexOf("Item Notes"); // NEW: Get Item Notes Column

    if (oiOrderNumCol === -1 || oiItemIdCol === -1 || oiQtyCol === -1) {
        Logger.log("Error: Missing required columns (Order #, Item ID, or Qty) in 'Order items' sheet.");
        throw new Error("Missing required columns in 'Order items' sheet.");
    }

    orderItemData.forEach(row => {
      const orderNum = row[oiOrderNumCol];
      const itemId = row[oiItemIdCol];
      const qty = row[oiQtyCol];
      const itemNotes = oiItemNotesCol !== -1 ? (row[oiItemNotesCol] ? row[oiItemNotesCol].toString().trim() : '') : ''; // NEW: Get item notes

      if (ordersMap.has(orderNum)) {
        const order = ordersMap.get(orderNum);
        const itemName = itemDetailsMap.has(itemId ? itemId.toString().trim().toLowerCase() : '')
                         ? itemDetailsMap.get(itemId.toString().trim().toLowerCase())
                         : 'Unknown Item Name';
        order.items.push({ itemId: itemId, itemName: itemName, qty: qty, itemNotes: itemNotes }); // NEW: Include itemNotes
      } else {
          Logger.log(`Warning: Order item (Order #${orderNum}, Item ${itemId}) found without a matching main order.`);
      }
    });
  } else {
      Logger.log("No order item data rows found.");
  }

  const ordersArray = Array.from(ordersMap.values());
  Logger.log(`Finished getOrdersWithItems function. Returning ${ordersArray.length} orders.`);
  return ordersArray;
}

/**
 * Updates the status of an order in the 'Orders' sheet.
 * @param {string} orderNum - The order number to update.
 * @param {string} newStatus - The new status to set.
 * @returns {boolean} True if update was successful, false otherwise.
 */
function updateOrderStatusByOrderNum(orderNum, newStatus) {
  Logger.log("Updating status for order: " + orderNum + " to " + newStatus);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    Logger.log("Error: Orders sheet not found in updateOrderStatusByOrderNum.");
    throw new Error("Orders sheet not found.");
  }

  const values = orderSheet.getDataRange().getValues();
  const headers = values[0];
  const orderNumCol = headers.indexOf("Order #");
  const statusCol = headers.indexOf("Status");

  if (orderNumCol === -1 || statusCol === -1) {
    Logger.log("Error: Required columns (Order # or Status) not found in Orders sheet for updateOrderStatusByOrderNum.");
    throw new Error("Required columns (Order # or Status) not found in Orders sheet.");
  }

  for (let i = 1; i < values.length; i++) { // Start from 1 to skip headers
    if (values[i][orderNumCol].toString() === orderNum.toString()) {
      orderSheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      Logger.log("Order " + orderNum + " status updated successfully.");
      return true;
    }
  }
  Logger.log("Order " + orderNum + " not found for status update.");
  return false; // Order not found
}

/**
 * Deletes specific items from an order in the 'Order items' sheet.
 * @param {string} orderNum - The order number.
 * @param {Array<string>} itemIdsToDelete - An array of item IDs to delete from that order.
 * @returns {boolean} True if items were deleted, false otherwise.
 */
function deleteOrderItems(orderNum, itemIdsToDelete) {
  Logger.log("Deleting items " + itemIdsToDelete.join(', ') + " from order " + orderNum);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderItemsSheet = ss.getSheetByName("Order items");
  if (!orderItemsSheet) {
    Logger.log("Error: Order items sheet not found in deleteOrderItems.");
    throw new Error("Order items sheet not found.");
  }

  const values = orderItemsSheet.getDataRange().getValues();
  const headers = values[0];
  const data = values.slice(1);

  const orderNumCol = headers.indexOf("Order #");
  const itemIdCol = headers.indexOf("Item ID");

  if (orderNumCol === -1 || itemIdCol === -1) {
    Logger.log("Error: Required columns (Order # or Item ID) not found in Order items sheet for deleteOrderItems.");
    throw new Error("Required columns (Order # or Item ID) not found in Order items sheet.");
  }

  let rowsToDelete = [];
  data.forEach((row, index) => {
    if (row[orderNumCol].toString() === orderNum.toString() && itemIdsToDelete.includes(row[itemIdCol].toString())) {
      rowsToDelete.push(index + 2); // Get actual row number in sheet (1-based, +1 for header, +1 for 0-indexed array)
    }
  });

  // Sort in descending order to delete from bottom up, preventing row index shifts
  rowsToDelete.sort((a, b) => b - a);

  for (let i = 0; i < rowsToDelete.length; i++) {
    orderItemsSheet.deleteRow(rowsToDelete[i]);
  }
  Logger.log("Deleted " + rowsToDelete.length + " items for order " + orderNum + ".");
  return rowsToDelete.length > 0;
}

/**
 * Removes a specific email from an order's email list.
 * Ensures at least one email remains.
 * @param {string} orderNum - The order number.
 * @param {string} emailToRemove - The email address to remove.
 * @returns {boolean} True if removal was successful, false if not found or last email.
 */
function removeEmailFromOrder(orderNum, emailToRemove) {
  Logger.log(`Attempting to remove email "${emailToRemove}" from order ${orderNum}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    Logger.log("Error: 'Orders' sheet not found in removeEmailFromOrder.");
    throw new Error("'Orders' sheet not found.");
  }

  const values = orderSheet.getDataRange().getValues();
  const headers = values[0];
  const orderNumCol = headers.indexOf("Order #");
  const emailCol = headers.indexOf("Email"); // Use 'Email' as per your sheet

  if (orderNumCol === -1 || emailCol === -1) {
    Logger.log("Error: Missing 'Order #' or 'Email' column in 'Orders' sheet for removeEmailFromOrder.");
    throw new Error("Missing required columns in 'Orders' sheet.");
  }

  for (let i = 1; i < values.length; i++) { // Start from 1 to skip headers
    if (values[i][orderNumCol].toString() === orderNum.toString()) {
      let currentEmails = values[i][emailCol].toString().split(',').map(e => e.trim()).filter(Boolean);
      const initialLength = currentEmails.length;

      const newEmails = currentEmails.filter(email => email !== emailToRemove);

      if (newEmails.length === initialLength) {
        Logger.log(`Email "${emailToRemove}" not found for order ${orderNum}. No change made.`);
        return false; // Email not found
      }
      if (newEmails.length === 0) {
        Logger.log(`Cannot remove last email "${emailToRemove}" from order ${orderNum}. An order must have at least one recipient.`);
        return false; // Cannot remove the last email
      }

      orderSheet.getRange(i + 1, emailCol + 1).setValue(newEmails.join(','));
      Logger.log(`Email "${emailToRemove}" removed successfully from order ${orderNum}. New emails: ${newEmails.join(',')}`);
      return true;
    }
  }
  Logger.log(`Order ${orderNum} not found for email removal.`);
  return false; // Order not found
}

/**
 * Cancels an entire order by updating its status to 'Cancelled'.
 * @param {string} orderNum - The order number to cancel.
 * @returns {boolean} True if the order was found and status updated, false otherwise.
 */
function cancelOrder(orderNum) {
    Logger.log(`Attempting to cancel order: ${orderNum}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName("Orders");
    if (!orderSheet) {
        Logger.log("Error: 'Orders' sheet not found in cancelOrder.");
        throw new Error("'Orders' sheet not found.");
    }

    const values = orderSheet.getDataRange().getValues();
    const headers = values[0];
    const orderNumCol = headers.indexOf("Order #");
    const statusCol = headers.indexOf("Status");

    if (orderNumCol === -1 || statusCol === -1) {
        Logger.log("Error: Required columns ('Order #' or 'Status') not found in 'Orders' sheet for cancelOrder.");
        throw new Error("Missing required columns in 'Orders' sheet.");
    }

    for (let i = 1; i < values.length; i++) { // Start from 1 to skip headers
        if (values[i][orderNumCol].toString() === orderNum.toString()) {
            orderSheet.getRange(i + 1, statusCol + 1).setValue("Cancelled");
            Logger.log(`Order ${orderNum} cancelled successfully.`);
            return true;
        }
    }
    Logger.log(`Order ${orderNum} not found for cancellation.`);
    return false;
}

/**
 * Retrieves application settings from the 'Config' sheet.
 * Assumes:
 * - 'Urgent Threshold' is in cell B2
 * - 'Low Threshold' is in cell C2
 * - 'Default Emails' are in cell K2 (comma-separated string)
 *
 * @returns {Object} An object containing the configuration settings.
 */
function getAppSettings() {
  Logger.log("Fetching app settings from Config sheet.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found in getAppSettings.");
    throw new Error("'Config' sheet not found.");
  }

  // Define the range where your settings are located
  const urgentThreshold = configSheet.getRange("B2").getValue();
  const lowThreshold = configSheet.getRange("C2").getValue();
  const defaultEmailsRaw = configSheet.getRange("K2").getValue();

  const defaultEmailsArray = typeof defaultEmailsRaw === 'string' && defaultEmailsRaw.trim() !== ''
    ? defaultEmailsRaw.split(',').map(email => email.trim()).filter(Boolean)
    : [];

  const settings = {
    urgentThreshold: urgentThreshold,
    lowThreshold: lowThreshold,
    defaultEmails: defaultEmailsArray
  };

  Logger.log("App settings retrieved: " + JSON.stringify(settings));
  return settings;
}

/**
 * Saves application settings to the 'Config' sheet.
 * Assumes:
 * - 'Urgent Threshold' will be written to cell B2
 * - 'Low Threshold' will be written to cell C2
 * - 'Default Emails' will be written to cell K2 (as a comma-separated string)
 *
 * @param {Object} settings - An object containing settings to save:
 * {number} urgentThreshold, {number} lowThreshold, {Array<string>} defaultEmails
 */
function saveAppSettings(settings) {
  Logger.log("Saving app settings to Config sheet: " + JSON.stringify(settings));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found in saveAppSettings.");
    throw new Error("'Config' sheet not found.");
  }

  // Validate inputs
  if (typeof settings.urgentThreshold !== 'number' && !isNaN(parseInt(settings.urgentThreshold))) {
      settings.urgentThreshold = parseInt(settings.urgentThreshold);
  }
  if (typeof settings.lowThreshold !== 'number' && !isNaN(parseInt(settings.lowThreshold))) {
      settings.lowThreshold = parseInt(settings.lowThreshold);
  }
  const emailsToSave = Array.isArray(settings.defaultEmails) ? settings.defaultEmails.join(',') : '';

  try {
    configSheet.getRange("B2").setValue(settings.urgentThreshold);
    configSheet.getRange("C2").setValue(settings.lowThreshold);
    configSheet.getRange("K2").setValue(emailsToSave);
    Logger.log("App settings saved successfully.");
  } catch (e) {
    Logger.log("Error saving settings to Config sheet: " + e.message);
    throw new Error("Failed to save settings: " + e.message);
  }
}
