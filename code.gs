/**
 * @OnlyCurrentDoc
 */

/**
 * Handles GET requests to the web app.
 * Routes to different HTML pages or handles specific actions based on URL parameters.
 * @param {Object} e The event object, containing URL parameters.
 * @returns {HtmlOutput} The HTML content to serve.
 */
function doGet(e) {
  // Check if a 'page' parameter is present in the URL
  if (e.parameter.page === 'config') {
    return HtmlService.createHtmlOutputFromFile('config')
      .setTitle('Jigsaw Inventory Config')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } 
  // Route to the order history page
  else if (e.parameter.page === 'orderHistory') {
    return HtmlService.createHtmlOutputFromFile('orderHistory')
      .setTitle('Jigsaw Order History')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } 
  // Handle action to mark an order as processing (from email link)
  else if (e.parameter.action === 'processOrder') {
    const orderNum = e.parameter.orderNum;
    if (orderNum) {
      try {
        const success = updateOrderStatusByOrderNum(orderNum, 'Processing');
        if (success) {
          return HtmlService.createHtmlOutput(`<p style="font-family: sans-serif; text-align: center; margin-top: 50px; font-size: 18px; color: green;">Order <strong>${orderNum}</strong> has been set to <strong>Processing</strong>.</p><p style="font-family: sans-serif; text-align: center; font-size: 14px;">You can close this window.</p>`)
            .setTitle('Order Status Updated')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
            // If updateOrderStatusByOrderNum returns false (order not found or already processed)
            return HtmlService.createHtmlOutput(`<p style="font-family: sans-serif; text-align: center; margin-top: 50px; font-size: 18px; color: orange;">Order <strong>${orderNum}</strong> not found or already processed.</p><p style="font-family: sans-serif; text-align: center; font-size: 14px;">Please check the order history.</p>`)
            .setTitle('Order Not Found/Already Processed')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
      } catch (error) {
        Logger.log(`Error processing order ${orderNum}: ${error.message}`);
        return HtmlService.createHtmlOutput(`<p style="font-family: sans-serif; text-align: center; margin-top: 50px; font-size: 18px; color: red;">Error updating order <strong>${orderNum}</strong>: ${error.message}</p><p style="font-family: sans-serif; text-align: center; font-size: 14px;">Please try again or contact support.</p>`)
          .setTitle('Error Updating Order')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } else {
      return HtmlService.createHtmlOutput('<p style="font-family: sans-serif; text-align: center; margin-top: 50px; font-size: 18px; color: red;">Invalid request: Order number missing.</p>')
        .setTitle('Invalid Request')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  // Default to the main inventory portal page
  else {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Jigsaw Inventory Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Returns the URL for the Order History page.
 * @returns {string} The URL of the order history web app.
 */
function getOrderHistoryFileUrl() {
  const url = ScriptApp.getService().getUrl();
  return url + '?page=orderHistory';
}

/**
 * Returns the URL for the main Index page (Inventory Portal).
 * @returns {string} The URL of the main inventory portal web app.
 */
function getIndexFileUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Returns the URL for the Config page.
 * @returns {string} The URL of the configuration web app.
 */
function getConfigFileUrl() {
  const url = ScriptApp.getService().getUrl();
  return url + '?page=config';
}

/**
 * Simple test function for pinging the server.
 * @returns {string} A "pong" response.
 */
function pingTest() {
  return "pong";
}

/**
 * Retrieves all Item IDs and Item Names from the "Product info" sheet.
 * This is used to populate the datalist for adding new items in config.html and index.html.
 * Assumes "Item ID" is in Column A and "Item Name" is in Column H.
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
  if (lastRow < 2) {
    Logger.log("No product data found in 'Product info' sheet.");
    return [];
  }

  // Get data from Column A (Item ID) and Column H (Item Name)
  const itemIdsRange = productInfoSheet.getRange(2, 1, lastRow - 1, 1);
  const itemNamesRange = productInfoSheet.getRange(2, 8, lastRow - 1, 1);
  
  const itemIds = itemIdsRange.getValues();
  const itemNames = itemNamesRange.getValues();

  const itemDetails = [];
  for (let i = 0; i < itemIds.length; i++) {
    const id = itemIds[i][0] ? itemIds[i][0].toString().trim() : '';
    const name = itemNames[i][0] ? itemNames[i][0].toString().trim() : '';

    if (id) {
      itemDetails.push({ id: id, name: name });
    }
  }

  Logger.log(`Found ${itemDetails.length} item details from 'Product info' sheet.`);
  return itemDetails;
}

/**
 * Retrieves all Item IDs from the "Product info" sheet.
 * (Kept for compatibility if still used elsewhere, otherwise redundant with getAllItemDetails).
 * @returns {Array<string>} An array of item IDs.
 */
function getAllItemIDs() {
  const itemDetails = getAllItemDetails();
  return itemDetails.map(item => item.id);
}

/**
 * Helper function to get an item's name by its ID.
 * @param {string} itemId The ID of the item to look up.
 * @returns {string|null} The item name, or null if not found.
 */
function getItemNameById(itemId) {
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
function submitOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  const orderItemsSheet = ss.getSheetByName("Order items");
  const configSheet = ss.getSheetByName("Config"); 
  const productInfoSheet = ss.getSheetByName("Product info"); 

  if (!orderSheet || !orderItemsSheet || !configSheet || !productInfoSheet) {
    Logger.log("Error: One or more required sheets (Orders, Order items, Config, Product info) not found in submitOrder.");
    throw new Error("One or more required sheets not found.");
  }

  const items = orderData.items;
  const recipientEmails = orderData.emails;
  const orderNotes = orderData.orderNotes || "";

  Logger.log("submitOrder: Function started.");
  Logger.log("submitOrder: Recipient Emails received: " + JSON.stringify(recipientEmails));
  Logger.log("submitOrder: Items received: " + JSON.stringify(items));
  Logger.log("submitOrder: Order Notes received: '" + orderNotes + "'");

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
    const lastOrderNumString = orderSheet.getRange(lastRow, 1).getValue(); 
    const match = lastOrderNumString.toString().match(/ORD-(\d+)/);
    if (match) {
      nextOrderNumVal = parseInt(match[1]) + 1;
    }
  }
  const newOrderNum = "ORD-" + Utilities.formatString("%04d", nextOrderNumVal);
  Logger.log(`Generated new order number: ${newOrderNum}`);
  // --- END ORDER NUMBER GENERATION ---

  const timestamp = new Date();
  const status = "Pending";

  // --- Prepare Product Info Map for Item Name lookup (for email template) ---
  const productInfoMap = new Map();
  const productInfoData = productInfoSheet.getDataRange().getValues();
  if (productInfoData.length > 1) {
      const pHeaders = productInfoData[0];
      const pIdCol = pHeaders.indexOf("Item ID");
      const pNameCol = pHeaders.indexOf("Item Name");

      if (pIdCol === -1 || pNameCol === -1) {
          Logger.log("Warning: 'Item ID' or 'Item Name' header not found in 'Product info' for lookup.");
      } else {
          for (let r = 1; r < productInfoData.length; r++) {
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
  const itemsForEmailTemplate = [];

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

    const itemName = item.itemName || productInfoMap.get(itemId.toLowerCase()) || 'Unknown Item';

    itemRowsForOrderItemsSheet.push([newOrderNum, itemId, qty, itemNotes]);

    itemsForEmailTemplate.push({
      itemId: itemId,
      itemName: itemName,
      qty: qty,
      itemNotes: itemNotes
    });
  });

  // --- Update Sheets ---
  orderSheet.appendRow([newOrderNum, timestamp, validRecipientEmails.join(','), status, orderNotes]);
  Logger.log(`Order #${newOrderNum} added to Orders sheet.`);

  if (itemRowsForOrderItemsSheet.length > 0) {
      orderItemsSheet.getRange(orderItemsSheet.getLastRow() + 1, 1, itemRowsForOrderItemsSheet.length, itemRowsForOrderItemsSheet[0].length).setValues(itemRowsForOrderItemsSheet);
      Logger.log(`Items for order #${newOrderNum} added to Order items sheet.`);
  }

  Logger.log("Stock deduction feature is disabled. Stock Tracker sheet was NOT updated.");

  // --- Send Confirmation Email ---
  sendOrderConfirmationEmail(validRecipientEmails, newOrderNum, itemsForEmailTemplate, orderNotes);

  // NEW: Update "On Order Flag" for items in this new order
  updateOnOrderFlag(itemRowsForOrderItemsSheet.map(row => row[1]), true);

  return { success: true, orderNum: newOrderNum };
}

/**
 * Retrieves stock status for urgent and low stock items from both stock sheets.
 * Now reads thresholds and comparison types from the "Config" sheet.
 * @returns {Object} An object containing urgent and low stock items.
 */
function getStockStatus() {
  Logger.log("getStockStatus: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker");
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker");
  const configSheet = ss.getSheetByName("Config");

  if (!stockTrackerSheet || !inksTrackerSheet || !configSheet) {
    Logger.log("Error: Missing required sheet(s) for getStockStatus (Stock Tracker, Inks Tracker, or Config).");
    throw new Error("Required sheet(s) not found for stock status calculation.");
  }

  const urgentItems = [];
  const lowItems = [];

  const configData = configSheet.getDataRange().getValues();
  Logger.log("getStockStatus: Raw configData: " + JSON.stringify(configData)); // DEBUG
  if (configData.length < 1) { // Check if sheet is completely empty
      Logger.log("Warning: Config sheet is empty. No item configurations to load.");
      return { urgent: [], low: [], urgentThreshold: 0, lowThreshold: 5 }; // Return empty results
  }

  const configHeaders = configData[0].map(h => h.toString().trim()); // Get headers and trim
  const configValues = configData.slice(1); // Actual data rows
  Logger.log("getStockStatus: Trimmed configHeaders: " + JSON.stringify(configHeaders)); // DEBUG

  // Dynamically find column indices based on trimmed headers
  const configItemIdCol = configHeaders.indexOf("Item ID");
  const urgentThresholdCol = configHeaders.indexOf("Urgent Threshold");
  const urgentComparisonCol = configHeaders.indexOf("Urgent Comparison");
  const dailyThresholdCol = configHeaders.indexOf("Daily Threshold");
  const dailyComparisonCol = configHeaders.indexOf("Daily Comparison");
  const notifyTypeCol = configHeaders.indexOf("Notify Type");
  const emailsCol = configHeaders.indexOf("Emails");
  const onOrderFlagCol = configHeaders.indexOf("On Order Flag");

  // Validate headers are found
  if (configItemIdCol === -1 || urgentThresholdCol === -1 || urgentComparisonCol === -1 ||
      dailyThresholdCol === -1 || dailyComparisonCol === -1 || notifyTypeCol === -1 ||
      emailsCol === -1 || onOrderFlagCol === -1) {
    Logger.log("Error: Missing one or more required headers in 'Config' sheet for getStockStatus.");
    const missingHeaders = [];
    if (configItemIdCol === -1) missingHeaders.push("Item ID");
    if (urgentThresholdCol === -1) missingHeaders.push("Urgent Threshold");
    if (urgentComparisonCol === -1) missingHeaders.push("Urgent Comparison");
    if (dailyThresholdCol === -1) missingHeaders.push("Daily Threshold");
    if (dailyComparisonCol === -1) missingHeaders.push("Daily Comparison");
    if (notifyTypeCol === -1) missingHeaders.push("Notify Type");
    if (emailsCol === -1) missingHeaders.push("Emails");
    if (onOrderFlagCol === -1) missingHeaders.push("On Order Flag");
    Logger.log("Missing headers detected by getStockStatus: " + missingHeaders.join(", "));
    throw new Error("Missing required headers in 'Config' sheet. Please check the sheet headers for exact spelling and no extra characters.");
  }
  Logger.log("getStockStatus: All required header indices found."); // DEBUG

  const itemThresholdsMap = new Map();

  // Populate itemThresholdsMap from Config sheet
  configValues.forEach((row, rowIndex) => { // Added rowIndex for debugging
    // Defensive check for row length before accessing index
    const maxHeaderIndex = Math.max(configItemIdCol, urgentThresholdCol, urgentComparisonCol, dailyThresholdCol, dailyComparisonCol,
                               notifyTypeCol, emailsCol, onOrderFlagCol); // Max index for item config fields
    if (row.length <= maxHeaderIndex) {
      Logger.log(`Warning: getStockStatus: Skipping config row ${rowIndex + 2} due to insufficient columns to read all expected headers: ${JSON.stringify(row)}`);
      return;
    }

    const itemId = row[configItemIdCol] ? row[configItemIdCol].toString().trim() : '';
    if (itemId) {
      const urgent = !isNaN(parseInt(row[urgentThresholdCol])) ? parseInt(row[urgentThresholdCol]) : 0;
      const urgentComp = row[urgentComparisonCol] ? row[urgentComparisonCol].toString().trim() : 'less_than_or_equal';
      const daily = !isNaN(parseInt(row[dailyThresholdCol])) ? parseInt(row[dailyThresholdCol]) : 5;
      const dailyComp = row[dailyComparisonCol] ? row[dailyComparisonCol].toString().trim() : 'less_than_or_equal';
      const notifyType = row[notifyTypeCol] ? row[notifyTypeCol].toString().trim().toLowerCase() : 'both';
      const emails = row[emailsCol] ? row[emailsCol].toString().trim() : '';
      const onOrderFlag = row[onOrderFlagCol] ? row[onOrderFlagCol].toString().trim() : '';

      itemThresholdsMap.set(itemId, { 
        urgent: urgent, 
        urgentComparison: urgentComp,
        daily: daily, 
        dailyComparison: dailyComp,
        notifyType: notifyType, 
        emails: emails,
        onOrderFlag: onOrderFlag
      });
      Logger.log(`getStockStatus: Mapped item config for ${itemId}: ${JSON.stringify(itemThresholdsMap.get(itemId))}`); // DEBUG
    } else {
        Logger.log(`Warning: getStockStatus: Skipping config row ${rowIndex + 2} due to empty Item ID.`); // DEBUG
    }
  });
  Logger.log(`getStockStatus: itemThresholdsMap size: ${itemThresholdsMap.size}`); // DEBUG


  // Helper function to process a single stock sheet
  const processStockSheet = (sheet, qtyColumnIndex) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log(`No data in ${sheet.getName()} to check stock status.`); // DEBUG
      return;
    }
    
    const stockValues = sheet.getDataRange().getValues();
    const stockHeaders = stockValues[0].map(h => h.toString().trim()); // Get headers and trim
    const stockDataRows = stockValues.slice(1);
    Logger.log(`processStockSheet: Processing ${sheet.getName()}. Trimmed headers: ${JSON.stringify(stockHeaders)}`); // DEBUG

    const stockItemIdCol = stockHeaders.indexOf("Item ID");
    const stockQtyCol = qtyColumnIndex; // Assuming pre-determined column index by parameter

    if (stockItemIdCol === -1 || stockQtyCol === -1) { // Check dynamic indexOf for safety
      Logger.log(`Warning: Missing 'Item ID' or 'Quantity' header in ${sheet.getName()} for stock status. (Or quantity column index is wrong).`);
      return;
    }

    stockDataRows.forEach((row, rowIndex) => {
      // Defensive check for row length before accessing index
      if (row.length <= Math.max(stockItemIdCol, stockQtyCol)) {
          Logger.log(`Warning: Skipping row ${rowIndex + 2} in ${sheet.getName()} due to insufficient columns: ${JSON.stringify(row)}`);
          return;
      }

      const itemId = row[stockItemIdCol] ? row[stockItemIdCol].toString().trim() : '';
      const qty = row[stockQtyCol];
      Logger.log(`processStockSheet: Checking ${sheet.getName()} row ${rowIndex + 2}: ItemID=${itemId}, Qty=${qty}`); // DEBUG

      if (typeof itemId === 'string' && itemId.trim() !== '' && typeof qty === 'number') {
        const config = itemThresholdsMap.get(itemId.trim());

        if (config) {
          // Evaluate Urgent Threshold
          let isUrgent = false;
          if (config.urgentComparison === 'less_than') {
              isUrgent = (qty < config.urgent);
          } else { // default to less_than_or_equal
              isUrgent = (qty <= config.urgent);
          }

          // Evaluate Daily Threshold
          let isDailyLow = false;
          if (config.dailyComparison === 'less_than') {
              isDailyLow = (qty < config.daily);
          } else { // default to less_than_or_equal
              isDailyLow = (qty <= config.daily);
          }

          // Suppress urgent if On Order Flag is true
          const suppressUrgent = (config.onOrderFlag === 'TRUE');
          Logger.log(`  - Config: ${JSON.stringify(config)}, isUrgent=${isUrgent}, isDailyLow=${isDailyLow}, suppressUrgent=${suppressUrgent}`); // DEBUG

          if (isUrgent && (config.notifyType === 'urgent' || config.notifyType === 'both') && !suppressUrgent) {
            urgentItems.push({ ItemID: itemId, Qty: qty, Threshold: config.urgent, NotifyType: config.notifyType, Emails: config.emails });
            Logger.log(`  - ADDED to Urgent: ${itemId}`); // DEBUG
          } else if (isDailyLow && (config.notifyType === 'daily' || config.notifyType === 'both')) {
            lowItems.push({ ItemID: itemId, Qty: qty, Threshold: config.daily, NotifyType: config.notifyType, Emails: config.emails });
            Logger.log(`  - ADDED to Low: ${itemId}`); // DEBUG
          }
        } else {
          Logger.log(`Warning: Item ID '${itemId}' found in ${sheet.getName()} but not in Config (no thresholds configured for it).`);
        }
      } else {
          Logger.log(`Warning: Skipping row ${rowIndex + 2} in ${sheet.getName()} due to invalid ItemID or Quantity (ItemID: '${itemId}', Qty: '${qty}').`); // DEBUG
      }
    });
  };

  // Process both stock sheets
  processStockSheet(stockTrackerSheet, 5); // Assuming Qty is in Column F (index 5) for Stock Tracker
  processStockSheet(inksTrackerSheet, 2); // Assuming Qty is in Column C (index 2) for Inks Tracker

  // These are now just placeholders for the UI titles in index.html, 
  // as the actual logic uses per-item thresholds.
  const defaultUrgentThreshold = 0; 
  const defaultLowThreshold = 5;     

  Logger.log(`getStockStatus: Finished. Urgent items count: ${urgentItems.length}, Low items count: ${lowItems.length}`); // DEBUG
  return { 
    urgent: urgentItems, 
    low: lowItems, 
    urgentThreshold: defaultUrgentThreshold, 
    lowThreshold: defaultLowThreshold        
  };
}

/**
 * DEPRECATED: This function is no longer used for per-item thresholds.
 * Global thresholds are replaced by item-specific thresholds in the Config sheet.
 * @returns {Object} An object containing urgentThreshold and lowThreshold.
 */
function getStockThresholds() {
  Logger.log("DEPRECATED: getStockThresholds() called. Use per-item thresholds from Config sheet directly.");
  return { urgentThreshold: 0, lowThreshold: 5 }; // Return defaults for compatibility if accidentally called
}

/**
 * Retrieves a list of global default recipient emails from the 'Config' sheet.
 * This function now looks for the column header "Default Order Emails" and gets the value from row 2.
 * @returns {Array<string>} An array of default email addresses.
 */
function getGlobalDefaultEmails() { 
    Logger.log("Attempting to load global default emails from Config sheet (column 'Default Order Emails').");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
        Logger.log("Error: Config sheet not found in getGlobalDefaultEmails.");
        return [];
    }

    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim()); // Trim headers
    const defaultEmailsCol = headers.indexOf("Default Order Emails"); // Find column index by header name

    if (defaultEmailsCol === -1) {
        Logger.log("Error: 'Default Order Emails' column not found in Config sheet. Check spelling/existence.");
        return [];
    }

    // Get value from row 2 (index 1) of the found column
    const defaultEmailsRaw = configSheet.getRange(2, defaultEmailsCol + 1).getValue(); 
    Logger.log("Raw global default emails from 'Default Order Emails' column: " + defaultEmailsRaw);

    let defaultEmails = [];
    if (typeof defaultEmailsRaw === 'string' && defaultEmailsRaw.trim() !== '') {
        defaultEmails = defaultEmailsRaw.split(',').map(e => e.trim()).filter(Boolean);
    } else {
        Logger.log("Column 'Default Order Emails' is empty or not a string. No global default emails to load.");
    }

    Logger.log("Global default emails retrieved: " + JSON.stringify(defaultEmails));
    return defaultEmails;
}

/**
 * Saves a list of global default recipient emails to the 'Config' sheet.
 * This function now looks for the column header "Default Order Emails" and saves to row 2.
 * @param {Array<string>} emails - An array of email addresses to save.
 */
function saveGlobalDefaultEmails(emails) {
    Logger.log("Saving global default emails to Config sheet (column 'Default Order Emails').");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
        Logger.log("Error: Config sheet not found in saveGlobalDefaultEmails.");
        throw new Error("Required sheet 'Config' not found.");
    }

    const headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim()); // Trim headers
    const defaultEmailsCol = headers.indexOf("Default Order Emails"); // Find column index by header name

    if (defaultEmailsCol === -1) {
        Logger.log("Error: 'Default Order Emails' column not found in Config sheet for saving. Check spelling/existence.");
        throw new Error("Missing 'Default Order Emails' column in Config sheet.");
    }

    const emailsToSave = Array.isArray(emails) ? emails.join(',') : '';
    try {
        // Save to row 2 (index 1) of the found column
        configSheet.getRange(2, defaultEmailsCol + 1).setValue(emailsToSave); 
        Logger.log(`Global default emails saved: ${emailsToSave}`);
    } catch (e) {
        Logger.log(`Error saving global default emails: ${e.message}`);
        throw new Error(`Failed to save global default emails: ${e.message}`);
    }
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
  template.orderNotes = orderNotes;
  template.appUrl = ScriptApp.getService().getUrl();

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
  Logger.log("getOrdersWithItems: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("Orders");
  const orderItemsSheet = ss.getSheetByName("Order items");
  const productInfoSheet = ss.getSheetByName("Product info");

  if (!orderSheet || !orderItemsSheet || !productInfoSheet) {
    Logger.log("Error: Required sheets (Orders, Order items, Product info) not found in getOrdersWithItems.");
    throw new Error("Required sheets 'Orders', 'Order items', or 'Product info' not found.");
  }

  // 1. Get Product Info details (ID and Name) for quick lookup
  const lastProductInfoRow = productInfoSheet.getLastRow();
  const itemDetailsMap = new Map();

  if (lastProductInfoRow > 1) {
    const productInfoValues = productInfoSheet.getDataRange().getValues();
    const productInfoHeaders = productInfoValues[0].map(h => h.toString().trim()); // Trim headers
    Logger.log("getOrdersWithItems: ProductInfo Headers: " + JSON.stringify(productInfoHeaders)); // DEBUG

    const pIdCol = productInfoHeaders.indexOf("Item ID");
    const pNameCol = productInfoHeaders.indexOf("Item Name");

    if (pIdCol === -1 || pNameCol === -1) {
      Logger.log("Error: 'Item ID' or 'Item Name' columns not found in 'Product info' sheet for order history item lookup. Check spelling/existence.");
      throw new Error("Missing 'Item ID' or 'Item Name' columns in 'Product info' sheet.");
    }

    for (let r = 1; r < productInfoValues.length; r++) {
      // Defensive check for row length
      if (productInfoValues[r].length <= Math.max(pIdCol, pNameCol)) {
          Logger.log(`Warning: getOrdersWithItems: Skipping Product Info row ${r + 1} due to insufficient columns: ${JSON.stringify(productInfoValues[r])}`);
          continue;
      }
      const id = productInfoValues[r][pIdCol] ? productInfoValues[r][pIdCol].toString().trim() : '';
      const name = productInfoValues[r][pNameCol] ? productInfoValues[r][pNameCol].toString().trim() : '';
      if (id) {
        itemDetailsMap.set(id.toLowerCase(), name);
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
  const orderHeaders = orderValues[0].map(h => h.toString().trim()); // Trim headers
  const orderData = orderValues.slice(1);
  Logger.log("getOrdersWithItems: Order Headers: " + JSON.stringify(orderHeaders)); // DEBUG

  const ordersMap = new Map();
  orderData.forEach((row, rowIndex) => {
    // Dynamically find indices for this row as well
    const orderNumCol = orderHeaders.indexOf("Order #");
    const timestampCol = orderHeaders.indexOf("Timestamp");
    const emailCol = orderHeaders.indexOf("Email");
    const statusCol = orderHeaders.indexOf("Status");
    const orderNotesCol = orderHeaders.indexOf("Order Notes");

    // Defensive check for row length
    const maxOrderHeaderIndex = Math.max(orderNumCol, timestampCol, emailCol, statusCol, orderNotesCol);
    if (row.length <= maxOrderHeaderIndex || orderNumCol === -1 || timestampCol === -1 || emailCol === -1 || statusCol === -1 || orderNotesCol === -1) {
        Logger.log(`Warning: getOrdersWithItems: Skipping Order row ${rowIndex + 2} due to missing header(s) or insufficient columns: ${JSON.stringify(row)}`);
        return;
    }

    const order = {};
    order.orderNum = row[orderNumCol] ? row[orderNumCol].toString().trim() : '';
    order.timestamp = (row[timestampCol] instanceof Date) ? row[timestampCol].toISOString() : (row[timestampCol] ? new Date(row[timestampCol]).toISOString() : null);
    order.email = row[emailCol] ? row[emailCol].toString().trim() : '';
    order.status = row[statusCol] ? row[statusCol].toString().trim() : '';
    order.orderNotes = row[orderNotesCol] ? row[orderNotesCol].toString().trim() : '';
    order.items = [];

    if (order.orderNum) {
        ordersMap.set(order.orderNum, order);
    } else {
        Logger.log(`Warning: Order row ${rowIndex + 2} skipped due to empty OrderNum.`);
    }
  });
  Logger.log(`Orders map populated with ${ordersMap.size} entries.`);

  // 3. Get all Order Items data
  const lastOrderItemRow = orderItemsSheet.getLastRow();
  if (lastOrderItemRow > 1) {
    const orderItemValues = orderItemsSheet.getDataRange().getValues();
    const orderItemHeaders = orderItemValues[0].map(h => h.toString().trim()); // Trim headers
    const orderItemData = orderItemValues.slice(1);
    Logger.log("getOrdersWithItems: Order Items Headers: " + JSON.stringify(orderItemHeaders)); // DEBUG

    const oiOrderNumCol = orderItemHeaders.indexOf("Order #");
    const oiItemIdCol = orderItemHeaders.indexOf("Item ID");
    const oiQtyCol = orderItemHeaders.indexOf("Quantity"); // Assumed "Quantity" as per config in emails.gs
    const oiItemNotesCol = orderItemHeaders.indexOf("Item Notes");

    // Validate headers
    if (oiOrderNumCol === -1 || oiItemIdCol === -1 || oiQtyCol === -1 || oiItemNotesCol === -1) {
        Logger.log("Error: Missing required columns (Order #, Item ID, Quantity, or Item Notes) in 'Order items' sheet. Check spelling/existence.");
        throw new Error("Missing required columns in 'Order items' sheet.");
    }

    orderItemData.forEach((row, rowIndex) => {
      // Defensive check for row length
      const maxOrderItemHeaderIndex = Math.max(oiOrderNumCol, oiItemIdCol, oiQtyCol, oiItemNotesCol);
      if (row.length <= maxOrderItemHeaderIndex) {
          Logger.log(`Warning: getOrdersWithItems: Skipping Order Item row ${rowIndex + 2} due to insufficient columns: ${JSON.stringify(row)}`);
          return;
      }

      const orderNum = row[oiOrderNumCol] ? row[oiOrderNumCol].toString().trim() : '';
      const itemId = row[oiItemIdCol] ? row[oiItemIdCol].toString().trim() : '';
      const qty = row[oiQtyCol] ? Number(row[oiQtyCol]) : 0; // Ensure qty is number
      const itemNotes = row[oiItemNotesCol] ? row[oiItemNotesCol].toString().trim() : '';

      if (ordersMap.has(orderNum)) {
        const order = ordersMap.get(orderNum);
        const itemName = itemDetailsMap.has(itemId ? itemId.toLowerCase() : '')
                         ? itemDetailsMap.get(itemId.toLowerCase())
                         : 'Unknown Item Name';
        order.items.push({ itemId: itemId, itemName: itemName, qty: qty, itemNotes: itemNotes });
      } else {
          Logger.log(`Warning: Order item (Order #${orderNum}, Item ${itemId}) found without a matching main order. (Row ${rowIndex + 2} in Order items).`);
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
 * Retrieves a single order and its associated items from the 'Orders' and 'Order items' sheets.
 * Includes Item Names where available from the 'Product info' sheet.
 * @param {string} orderNum - The order number to retrieve.
 * @returns {Object|null} A single order object, or null if not found.
 */
function getSingleOrderDetails(orderNum) {
    Logger.log(`Attempting to retrieve details for order: ${orderNum}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName("Orders");
    const orderItemsSheet = ss.getSheetByName("Order items");
    const productInfoSheet = ss.getSheetByName("Product info");

    if (!orderSheet || !orderItemsSheet || !productInfoSheet) {
        Logger.log("Error: Required sheets (Orders, Order items, Product info) not found in getSingleOrderDetails.");
        throw new Error("Required sheets 'Orders', 'Order items', or 'Product info' not found.");
    }

    // 1. Get Product Info details (ID and Name) for quick lookup
    const lastProductInfoRow = productInfoSheet.getLastRow();
    const itemDetailsMap = new Map();

    if (lastProductInfoRow > 1) {
        const productInfoValues = productInfoSheet.getDataRange().getValues();
        const productInfoHeaders = productInfoValues[0].map(h => h.toString().trim());
        const pIdCol = productInfoHeaders.indexOf("Item ID");
        const pNameCol = productInfoHeaders.indexOf("Item Name");

        if (pIdCol === -1 || pNameCol === -1) {
            Logger.log("Error: 'Item ID' or 'Item Name' columns not found in 'Product info' sheet for single order item lookup. Check spelling/existence.");
            // Continue, but item names might be 'Unknown'
        } else {
            for (let r = 1; r < productInfoValues.length; r++) {
                if (productInfoValues[r].length <= Math.max(pIdCol, pNameCol)) continue; // Defensive check
                const id = productInfoValues[r][pIdCol] ? productInfoValues[r][pIdCol].toString().trim() : '';
                const name = productInfoValues[r][pNameCol] ? productInfoValues[r][pNameCol].toString().trim() : '';
                if (id) itemDetailsMap.set(id.toLowerCase(), name);
            }
        }
    }

    // 2. Find the specific order in the 'Orders' sheet
    const orderValues = orderSheet.getDataRange().getValues();
    const orderHeaders = orderValues[0].map(h => h.toString().trim());
    const orderNumCol = orderHeaders.indexOf("Order #");
    const timestampCol = orderHeaders.indexOf("Timestamp");
    const emailCol = orderHeaders.indexOf("Email");
    const statusCol = orderHeaders.indexOf("Status");
    const orderNotesCol = orderHeaders.indexOf("Order Notes");

    if (orderNumCol === -1 || timestampCol === -1 || emailCol === -1 || statusCol === -1 || orderNotesCol === -1) {
        Logger.log("Error: Missing required columns in 'Orders' sheet for getSingleOrderDetails. Check spelling/existence.");
        throw new Error("Missing required columns in 'Orders' sheet.");
    }

    let foundOrder = null;
    for (let i = 1; i < orderValues.length; i++) {
        if (orderValues[i].length <= Math.max(orderNumCol, timestampCol, emailCol, statusCol, orderNotesCol)) continue; // Defensive check
        if (orderValues[i][orderNumCol] && orderValues[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
            foundOrder = {
                orderNum: row[orderNumCol].toString().trim(),
                timestamp: (row[timestampCol] instanceof Date) ? row[timestampCol].toISOString() : (row[timestampCol] ? new Date(row[timestampCol]).toISOString() : null),
                email: row[emailCol] ? row[emailCol].toString().trim() : '',
                status: row[statusCol] ? row[statusCol].toString().trim() : '',
                orderNotes: row[orderNotesCol] ? row[orderNotesCol].toString().trim() : '',
                items: []
            };
            break;
        }
    }

    if (!foundOrder) {
        Logger.log(`Order ${orderNum} not found in 'Orders' sheet.`);
        return null;
    }

    // 3. Get all items for the found order from 'Order items' sheet
    const orderItemValues = orderItemsSheet.getDataRange().getValues();
    const orderItemHeaders = orderItemValues[0].map(h => h.toString().trim());
    const oiOrderNumCol = orderItemHeaders.indexOf("Order #");
    const oiItemIdCol = orderItemHeaders.indexOf("Item ID");
    const oiQtyCol = orderItemHeaders.indexOf("Quantity"); // Assumed "Quantity" as per config in emails.gs
    const oiItemNotesCol = orderItemHeaders.indexOf("Item Notes");

    if (oiOrderNumCol === -1 || oiItemIdCol === -1 || oiQtyCol === -1 || oiItemNotesCol === -1) {
        Logger.log("Error: Missing required columns (Order #, Item ID, Quantity, or Item Notes) in 'Order items' sheet for getSingleOrderDetails. Check spelling/existence.");
        throw new Error("Missing required columns in 'Order items' sheet.");
    }

    for (let i = 1; i < orderItemValues.length; i++) {
        if (orderItemValues[i].length <= Math.max(oiOrderNumCol, oiItemIdCol, oiQtyCol, oiItemNotesCol)) continue; // Defensive check
        if (orderItemValues[i][oiOrderNumCol] && orderItemValues[i][oiOrderNumCol].toString().trim() === orderNum.toString().trim()) {
            const itemId = orderItemValues[i][oiItemIdCol] ? orderItemValues[i][oiItemIdCol].toString().trim() : '';
            const qty = orderItemValues[i][oiQtyCol] ? Number(orderItemValues[i][oiQtyCol]) : 0;
            const itemNotes = orderItemValues[i][oiItemNotesCol] ? orderItemValues[i][oiItemNotesCol].toString().trim() : '';
            const itemName = itemDetailsMap.has(itemId ? itemId.toLowerCase() : '')
                             ? itemDetailsMap.get(itemId.toLowerCase())
                             : 'Unknown Item Name';
            foundOrder.items.push({ itemId: itemId, itemName: itemName, qty: qty, itemNotes: itemNotes });
        }
    }

    Logger.log(`Finished retrieving details for order ${orderNum}: ${JSON.stringify(foundOrder)}`);
    return foundOrder;
}

/**
 * Updates an existing order in the 'Orders' and 'Order items' sheets.
 * This function will overwrite the items for the given order number in 'Order items'.
 * @param {string} orderNum - The order number to update.
 * @param {string} newOrderNotes - The updated notes for the entire order.
 * @param {Array<Object>} updatedItems - An array of item objects {itemId: string, itemName: string, qty: number, itemNotes: string}
 * representing the new set of items for this order.
 * @returns {Object} An object with success status and a message.
 */
function updateOrder(orderNum, newOrderNotes, updatedItems) {
    Logger.log(`Attempting to update order: ${orderNum}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName("Orders");
    const orderItemsSheet = ss.getSheetByName("Order items");

    if (!orderSheet || !orderItemsSheet) {
        Logger.log("Error: Required sheets (Orders, Order items) not found in updateOrder.");
        throw new Error("Required sheets 'Orders' or 'Order items' not found.");
    }

    let orderRowIndex = -1;
    const orderValues = orderSheet.getDataRange().getValues();
    const orderHeaders = orderValues[0].map(h => h.toString().trim());
    const orderNumCol = orderHeaders.indexOf("Order #");
    const orderNotesCol = orderHeaders.indexOf("Order Notes");

    if (orderNumCol === -1 || orderNotesCol === -1) {
        Logger.log("Error: Missing 'Order #' or 'Order Notes' column in 'Orders' sheet for updateOrder.");
        throw new Error("Missing required columns in 'Orders' sheet.");
    }

    // Find the order row in the 'Orders' sheet
    for (let i = 1; i < orderValues.length; i++) {
        if (orderValues[i].length <= Math.max(orderNumCol, orderNotesCol)) continue; // Defensive check
        if (orderValues[i][orderNumCol] && orderValues[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
            orderRowIndex = i;
            break;
        }
    }

    if (orderRowIndex === -1) {
        Logger.log(`Order ${orderNum} not found in 'Orders' sheet for update.`);
        return { success: false, message: `Order ${orderNum} not found.` };
    }

    // Update Order Notes in 'Orders' sheet
    orderSheet.getRange(orderRowIndex + 1, orderNotesCol + 1).setValue(newOrderNotes);
    Logger.log(`Order notes for ${orderNum} updated.`);

    // --- Update 'Order items' sheet ---
    const allOrderItemsValues = orderItemsSheet.getDataRange().getValues();
    const orderItemHeaders = allOrderItemsValues[0].map(h => h.toString().trim());
    const currentOrderItemsData = allOrderItemsValues.slice(1);

    const oiOrderNumCol = orderItemHeaders.indexOf("Order #");
    const oiItemIdCol = orderItemHeaders.indexOf("Item ID");
    const oiQtyCol = orderItemHeaders.indexOf("Quantity"); // Assumed "Quantity"
    const oiItemNotesCol = orderItemHeaders.indexOf("Item Notes");

    if (oiOrderNumCol === -1 || oiItemIdCol === -1 || oiQtyCol === -1 || oiItemNotesCol === -1) {
        Logger.log("Error: Missing required columns in 'Order items' sheet for updateOrder.");
        throw new Error("Missing required columns in 'Order items' sheet.");
    }

    // 2. Filter out old items for this order number
    const filteredOrderItems = currentOrderItemsData.filter(row => {
        if (row.length <= oiOrderNumCol) return false; // Defensive check
        return row[oiOrderNumCol] && row[oiOrderNumCol].toString().trim() !== orderNum.toString().trim();
    });
    Logger.log(`Removed existing items for order ${orderNum} from 'Order items' sheet.`);

    // 3. Prepare new items to be added for this order
    const newItemsForSheet = [];
    if (Array.isArray(updatedItems) && updatedItems.length > 0) {
        updatedItems.forEach(item => {
            const itemId = item.itemId ? item.itemId.toString().trim() : '';
            const qty = item.qty ? Number(item.qty) : 0;
            const itemNotes = item.itemNotes ? item.itemNotes.toString().trim() : '';

            if (itemId && !isNaN(qty) && qty > 0) {
                // Ensure array length matches sheet columns expected
                const newRow = new Array(orderItemHeaders.length).fill('');
                newRow[oiOrderNumCol] = orderNum;
                newRow[oiItemIdCol] = itemId;
                newRow[oiQtyCol] = qty;
                newRow[oiItemNotesCol] = itemNotes;
                newItemsForSheet.push(newRow);
            } else {
                Logger.log(`Warning: Skipping invalid item during update for order ${orderNum}: ${JSON.stringify(item)}`);
            }
        });
    }

    // 4. Clear and rewrite 'Order items' sheet
    // Use clearContents() from the first cell to ensure proper clear if ranges change
    if (orderItemsSheet.getLastRow() > 1) {
        orderItemsSheet.getRange(2, 1, orderItemsSheet.getLastRow() -1, orderItemsSheet.getLastColumn()).clearContent();
    }
    
    // Rewrite headers (optional but good for robustness if headers were cleared for some reason)
    orderItemsSheet.getRange(1, 1, 1, orderItemHeaders.length).setValues([orderItemHeaders]);

    const finalDataForOrderItemsSheet = filteredOrderItems.concat(newItemsForSheet);
    if (finalDataForOrderItemsSheet.length > 0) {
        // Use the maximum width of the new data or existing headers
        const colsToWrite = Math.max(orderItemHeaders.length, finalDataForOrderItemsSheet[0].length);
        orderItemsSheet.getRange(2, 1, finalDataForOrderItemsSheet.length, colsToWrite).setValues(finalDataForOrderItemsSheet);
    }
    Logger.log(`Rewrote 'Order items' sheet with ${newItemsForSheet.length} new items for order ${orderNum}.`);

    SpreadsheetApp.flush();

    return { success: true, message: `Order ${orderNum} updated successfully.` };
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
  const orderItemsSheet = ss.getSheetByName("Order items");
  const configSheet = ss.getSheetByName("Config");

  if (!orderSheet || !orderItemsSheet || !configSheet) {
    Logger.log("Error: Required sheets (Orders, Order items, Config) not found in updateOrderStatusByOrderNum.");
    throw new Error("Required sheets 'Orders', 'Order items', or 'Config' not found.");
  }

  const values = orderSheet.getDataRange().getValues();
  const headers = values[0].map(h => h.toString().trim());
  const orderNumCol = headers.indexOf("Order #");
  const statusCol = headers.indexOf("Status");

  if (orderNumCol === -1 || statusCol === -1) {
    Logger.log("Error: Required columns (Order # or Status) not found in Orders sheet for updateOrderStatusByOrderNum.");
    throw new Error("Missing required columns in 'Orders' sheet.");
  }

  for (let i = 1; i < values.length; i++) {
    if (values[i].length <= Math.max(orderNumCol, statusCol)) continue; // Defensive check
    if (values[i][orderNumCol] && values[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
      orderSheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      Logger.log(`Order ${orderNum} status updated successfully.`);

      // If order is completed, clear the "On Order Flag" for its items
      if (newStatus === 'Completed') {
          const itemIdsInOrder = getItemIDsForOrder(orderNum);
          updateOnOrderFlag(itemIdsInOrder, false);
      }
      return true;
    }
  }
  Logger.log(`Order ${orderNum} not found for status update.`);
  return false;
}

/**
 * Helper function to get all Item IDs associated with a specific order.
 * @param {string} orderNum The order number.
 * @returns {Array<string>} An array of Item IDs.
 */
function getItemIDsForOrder(orderNum) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderItemsSheet = ss.getSheetByName("Order items");
    if (!orderItemsSheet) {
        Logger.log("Error: 'Order items' sheet not found in getItemIDsForOrder.");
        throw new Error("Required sheet 'Order items' not found.");
    }

    const values = orderItemsSheet.getDataRange().getValues();
    const headers = values[0].map(h => h.toString().trim());
    const oiOrderNumCol = headers.indexOf("Order #");
    const oiItemIdCol = headers.indexOf("Item ID");

    if (oiOrderNumCol === -1 || oiItemIdCol === -1) {
        Logger.log("Error: Missing 'Order #' or 'Item ID' columns in 'Order items' sheet for getItemIDsForOrder.");
        throw new Error("Missing required columns in 'Order items' sheet.");
    }

    const itemIds = [];
    for (let i = 1; i < values.length; i++) {
        if (values[i].length <= Math.max(oiOrderNumCol, oiItemIdCol)) continue; // Defensive check
        if (values[i][oiOrderNumCol] && values[i][oiOrderNumCol].toString().trim() === orderNum.toString().trim()) {
            itemIds.push(values[i][oiItemIdCol].toString().trim());
        }
    }
    return itemIds;
}

/**
 * Updates the "On Order Flag" for specified items in the "Config" sheet.
 * @param {Array<string>} itemIds - An array of Item IDs to update.
 * @param {boolean} isOnOrder - True to set the flag, false to clear it.
 */
function updateOnOrderFlag(itemIds, isOnOrder) {
    Logger.log(`Updating On Order Flag for items: ${itemIds.join(', ')} to ${isOnOrder}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Config");

    if (!configSheet) {
        Logger.log("Error: 'Config' sheet not found in updateOnOrderFlag.");
        throw new Error("Required sheet 'Config' not found.");
    }

    const configData = configSheet.getDataRange().getValues();
    const configHeaders = configData[0].map(h => h.toString().trim()); // Trim headers
    const configRows = configData.slice(1);

    const configItemIdCol = configHeaders.indexOf("Item ID");
    const onOrderFlagCol = configHeaders.indexOf("On Order Flag"); // Column H

    if (configItemIdCol === -1 || onOrderFlagCol === -1) {
        Logger.log("Error: Missing 'Item ID' or 'On Order Flag' column in 'Config' sheet for updateOnOrderFlag. Check spelling/existence.");
        throw new Error("Missing required columns in 'Config' sheet.");
    }

    let updatesMade = false;
    configRows.forEach((row, index) => {
        if (row.length <= Math.max(configItemIdCol, onOrderFlagCol)) return; // Defensive check
        const itemIdInConfig = row[configItemIdCol] ? row[configItemIdCol].toString().trim() : '';
        if (itemIds.includes(itemIdInConfig)) {
            const newValue = isOnOrder ? "TRUE" : "";
            if (row[onOrderFlagCol] !== newValue) {
                configSheet.getRange(index + 2, onOrderFlagCol + 1).setValue(newValue);
                Logger.log(`Set On Order Flag for ${itemIdInConfig} to ${newValue}.`);
                updatesMade = true;
            }
        }
    });

    if (updatesMade) {
        SpreadsheetApp.flush();
        Logger.log("Finished updating On Order Flags.");
    } else {
        Logger.log("No On Order Flag updates needed.");
    }
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
  const headers = values[0].map(h => h.toString().trim());
  const data = values.slice(1);

  const orderNumCol = headers.indexOf("Order #");
  const itemIdCol = headers.indexOf("Item ID");

  if (orderNumCol === -1 || itemIdCol === -1) {
    Logger.log("Error: Required columns (Order # or Item ID) not found in Order items sheet for deleteOrderItems.");
    throw new Error("Required columns (Order # or Item ID) not found in Order items sheet.");
  }

  let rowsToDelete = [];
  data.forEach((row, index) => {
    if (row.length <= Math.max(orderNumCol, itemIdCol)) return; // Defensive check
    if (row[orderNumCol] && row[orderNumCol].toString().trim() === orderNum.toString().trim() && itemIdsToDelete.includes(row[itemIdCol].toString().trim())) {
      rowsToDelete.push(index + 2);
    }
  });

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
  const headers = values[0].map(h => h.toString().trim());
  const orderNumCol = headers.indexOf("Order #");
  const emailCol = headers.indexOf("Email");

  if (orderNumCol === -1 || emailCol === -1) {
    Logger.log("Error: Missing 'Order #' or 'Email' column in 'Orders' sheet for removeEmailFromOrder.");
    throw new Error("Missing required columns in 'Orders' sheet.");
  }

  for (let i = 1; i < values.length; i++) {
    if (values[i].length <= Math.max(orderNumCol, emailCol)) continue; // Defensive check
    if (values[i][orderNumCol] && values[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
      let currentEmails = values[i][emailCol] ? values[i][emailCol].toString().split(',').map(e => e.trim()).filter(Boolean) : [];
      const initialLength = currentEmails.length;

      const newEmails = currentEmails.filter(email => email !== emailToRemove);

      if (newEmails.length === initialLength) {
        Logger.log(`Email "${emailToRemove}" not found for order ${orderNum}. No change made.`);
        return false;
      }
      if (newEmails.length === 0) {
        Logger.log(`Cannot remove last email "${emailToRemove}" from order ${orderNum}. An order must have at least one recipient.`);
        return false;
      }

      orderSheet.getRange(i + 1, emailCol + 1).setValue(newEmails.join(','));
      Logger.log(`Email "${emailToRemove}" removed successfully from order ${orderNum}. New emails: ${newEmails.join(',')}`);
      return true;
    }
  }
  Logger.log(`Order ${orderNum} not found for email removal.`);
  return false;
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
    const headers = values[0].map(h => h.toString().trim());
    const orderNumCol = headers.indexOf("Order #");
    const statusCol = headers.indexOf("Status");

    if (orderNumCol === -1 || statusCol === -1) {
        Logger.log("Error: Required columns ('Order #' or 'Status') not found in 'Orders' sheet for cancelOrder.");
        throw new Error("Missing required columns in 'Orders' sheet.");
    }

    for (let i = 1; i < values.length; i++) {
        if (values[i].length <= Math.max(orderNumCol, statusCol)) continue; // Defensive check
        if (values[i][orderNumCol] && values[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
            orderSheet.getRange(i + 1, statusCol + 1).setValue("Cancelled");
            Logger.log(`Order ${orderNum} cancelled successfully.`);
            return true;
        }
    }
    Logger.log(`Order ${orderNum} not found for cancellation.`);
    return false;
}

/**
 * Permanently deletes an order and its associated items from the sheets.
 * This action is irreversible.
 * @param {string} orderNum - The order number to delete.
 * @returns {boolean} True if deletion was successful, false otherwise.
 */
function deleteOrderPermanently(orderNum) {
    Logger.log(`Attempting to permanently delete order: ${orderNum}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName("Orders");
    const orderItemsSheet = ss.getSheetByName("Order items");

    if (!orderSheet || !orderItemsSheet) {
        Logger.log("Error: Required sheets (Orders, Order items) not found for permanent deletion.");
        throw new Error("Required sheets 'Orders' or 'Order items' not found.");
    }

    let orderDeleted = false;
    let itemsDeleted = false;

    // 1. Delete from 'Orders' sheet
    const orderValues = orderSheet.getDataRange().getValues();
    const orderHeaders = orderValues[0].map(h => h.toString().trim());
    const orderNumCol = orderHeaders.indexOf("Order #");

    if (orderNumCol === -1) {
        Logger.log("Error: 'Order #' column not found in 'Orders' sheet for permanent deletion.");
        throw new Error("Missing 'Order #' column in 'Orders' sheet.");
    }

    for (let i = orderValues.length - 1; i >= 1; i--) {
        if (orderValues[i].length <= orderNumCol) continue; // Defensive check
        if (orderValues[i][orderNumCol] && orderValues[i][orderNumCol].toString().trim() === orderNum.toString().trim()) {
            orderSheet.deleteRow(i + 1);
            orderDeleted = true;
            Logger.log(`Order ${orderNum} deleted from 'Orders' sheet.`);
            break;
        }
    }

    // 2. Delete associated items from 'Order items' sheet
    const orderItemValues = orderItemsSheet.getDataRange().getValues();
    const orderItemHeaders = orderItemValues[0].map(h => h.toString().trim());
    const oiOrderNumCol = orderItemHeaders.indexOf("Order #");

    if (oiOrderNumCol === -1) {
        Logger.log("Error: 'Order #' column not found in 'Order items' sheet for permanent deletion.");
        throw new Error("Missing 'Order #' column in 'Order items' sheet.");
    }

    let rowsToDeleteInOrderItems = [];
    for (let i = orderItemValues.length - 1; i >= 1; i--) {
        if (orderItemValues[i].length <= oiOrderNumCol) continue; // Defensive check
        if (orderItemValues[i][oiOrderNumCol] && orderItemValues[i][oiOrderNumCol].toString().trim() === orderNum.toString().trim()) {
            rowsToDeleteInOrderItems.push(i + 1);
        }
    }

    for (let i = 0; i < rowsToDeleteInOrderItems.length; i++) {
        orderItemsSheet.deleteRow(rowsToDeleteInOrderItems[i]);
        itemsDeleted = true;
    }
    if (itemsDeleted) {
        Logger.log(`${rowsToDeleteInOrderItems.length} items for order ${orderNum} deleted from 'Order items' sheet.`);
    } else {
        Logger.log(`No items found for order ${orderNum} in 'Order items' sheet.`);
    }

    SpreadsheetApp.flush();

    return orderDeleted;
}

/**
 * Retrieves all item configurations from the "Config" sheet.
 * Assumes the columns based on your provided screenshot:
 * A: Item ID, B: Urgent Threshold, C: Urgent Comparison, D: Daily Threshold, E: Daily Comparison,
 * F: Notify Type, G: Emails, H: On Order Flag, I: Last Urgent Sent, J: Notes.
 * @returns {Array<Object>} An array of item configuration objects.
 */
function getIndividualItemConfigs() {
  Logger.log("getIndividualItemConfigs: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found in getIndividualItemConfigs.");
    throw new Error("Required sheet 'Config' not found.");
  }

  const lastRow = configSheet.getLastRow();
  const lastColumn = configSheet.getLastColumn();
  // Read all headers from row 1, trimming each one
  const allHeadersInSheet = configSheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => h.toString().trim());
  Logger.log("getIndividualItemConfigs: All headers found in sheet: " + JSON.stringify(allHeadersInSheet)); // DEBUG

  // Dynamically find column indices based on trimmed headers
  const headerMap = {};
  const requiredHeaders = [
    "Item ID", "Urgent Threshold", "Urgent Comparison", "Daily Threshold", "Daily Comparison",
    "Notify Type", "Emails", "On Order Flag", "Last Urgent Sent", "Notes"
  ];
  requiredHeaders.forEach(h => {
    headerMap[h] = allHeadersInSheet.indexOf(h);
  });
  Logger.log("getIndividualItemConfigs: Found header indices: " + JSON.stringify(headerMap)); // DEBUG

  // Validate all required headers are found (after trimming)
  const missingHeaders = requiredHeaders.filter(h => headerMap[h] === -1);
  if (missingHeaders.length > 0) {
    Logger.log("Error: Missing or misspelled required headers in 'Config' sheet: " + missingHeaders.join(", "));
    throw new Error("Missing required headers in 'Config' sheet. Please verify exact spelling and no extra spaces.");
  }

  // Determine the number of columns to read based on the highest index of a required header
  const numColumnsToRead = Math.max(...Object.values(headerMap)) + 1;
  Logger.log("getIndividualItemConfigs: Reading up to column index: " + (numColumnsToRead - 1)); // DEBUG

  if (lastRow < 2 || numColumnsToRead < 1) {
    Logger.log("No data rows or readable columns found in 'Config' sheet (after header check).");
    return [];
  }

  // Read data starting from row 2 up to numColumnsToRead
  const configValues = configSheet.getRange(2, 1, lastRow - 1, numColumnsToRead).getValues();
  Logger.log("getIndividualItemConfigs: Raw configValues (data rows): " + JSON.stringify(configValues)); // DEBUG

  const configs = [];
  configValues.forEach((row, rowIndex) => { // Added rowIndex for debugging
    // Defensive check for row length before accessing index
    if (row.length <= maxManagedColIndex) { // Using maxManagedColIndex from save function perspective for safety
        Logger.log(`Warning: getIndividualItemConfigs: Skipping config row ${rowIndex + 2} due to insufficient columns to read all expected headers: ${JSON.stringify(row)}`);
        return;
    }

    const itemId = row[headerMap["Item ID"]] ? row[headerMap["Item ID"]].toString().trim() : '';
    if (itemId) {
      const urgentThreshold = row[headerMap["Urgent Threshold"]] !== undefined && !isNaN(parseInt(row[headerMap["Urgent Threshold"]])) ? parseInt(row[headerMap["Urgent Threshold"]]) : 0;
      const urgentComparison = row[headerMap["Urgent Comparison"]] ? row[headerMap["Urgent Comparison"]].toString().trim() : 'less_than_or_equal';
      const dailyThreshold = row[headerMap["Daily Threshold"]] !== undefined && !isNaN(parseInt(row[headerMap["Daily Threshold"]])) ? parseInt(row[headerMap["Daily Threshold"]]) : 5;
      const dailyComparison = row[headerMap["Daily Comparison"]] ? row[headerMap["Daily Comparison"].toString().trim()] : 'less_than_or_equal';
      const notifyType = row[headerMap["Notify Type"]] ? row[headerMap["Notify Type"]].toString().trim() : 'both';
      const emails = row[headerMap["Emails"]] ? row[headerMap["Emails"]].toString().trim() : '';
      const onOrderFlag = row[headerMap["On Order Flag"]] ? row[headerMap["On Order Flag"].toString().trim()] : '';
      const notes = row[headerMap["Notes"]] ? row[headerMap["Notes"].toString().trim()] : '';
      const lastUrgentSent = row[headerMap["Last Urgent Sent"]] || ''; 

      configs.push({
        itemId: itemId,
        urgentThreshold: urgentThreshold,
        urgentComparison: urgentComparison,
        dailyThreshold: dailyThreshold,
        dailyComparison: dailyComparison,
        notifyType: notifyType,
        notes: notes,
        emails: emails,
        onOrderFlag: onOrderFlag,
        lastUrgentSent: lastUrgentSent
      });
      Logger.log(`getIndividualItemConfigs: Successfully parsed item config for ${itemId}: ${JSON.stringify(configs[configs.length - 1])}`); // DEBUG
    } else {
        Logger.log(`Warning: getIndividualItemConfigs: Skipping config row ${rowIndex + 2} due to empty Item ID.`); // DEBUG
    }
  });

  Logger.log(`getIndividualItemConfigs: Function finished. Returning ${configs.length} item configurations.`); // DEBUG
  return configs;
}

/**
 * Saves an array of item configurations to the "Config" sheet.
 * Assumes the columns based on your provided screenshot:
 * A: Item ID, B: Urgent Threshold, C: Urgent Comparison, D: Daily Threshold, E: Daily Comparison,
 * F: Notify Type, G: Emails, H: On Order Flag, I: Last Urgent Sent, J: Notes.
 * @param {Array<Object>} configs - An array of item configuration objects.
 */
function saveIndividualItemConfigs(configs) {
  Logger.log("saveIndividualItemConfigs: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");

  if (!configSheet) {
    Logger.log("Error: 'Config' sheet not found in saveIndividualItemConfigs.");
    throw new Error("Required sheet 'Config' not found.");
  }

  const existingData = configSheet.getDataRange().getValues();
  const existingHeaders = existingData[0].map(h => h.toString().trim()); // Trim headers
  const existingRows = existingData.slice(1);
  Logger.log("saveIndividualItemConfigs: Existing headers: " + JSON.stringify(existingHeaders)); // DEBUG

  // Dynamically find column indices for preservation based on current headers
  const existingItemIdCol = existingHeaders.indexOf("Item ID");
  const existingLastUrgentSentCol = existingHeaders.indexOf("Last Urgent Sent");
  const existingOnOrderFlagCol = existingHeaders.indexOf("On Order Flag");
  const existingNotesCol = existingHeaders.indexOf("Notes");

  const lastUrgentSentMap = new Map();
  const onOrderFlagMap = new Map();
  const notesMap = new Map();

  if (existingItemIdCol !== -1) {
    existingRows.forEach(row => {
      // Defensive check for row length before accessing index
      const maxColToCheck = Math.max(existingItemIdCol, existingLastUrgentSentCol, existingOnOrderFlagCol, existingNotesCol);
      if (row.length <= maxColToCheck) {
          Logger.log(`Warning: saveIndividualItemConfigs: Skipping existing row due to insufficient columns for flags/notes: ${JSON.stringify(row)}`);
          return;
      }
      const itemId = row[existingItemIdCol] ? row[existingItemIdCol].toString().trim() : '';
      if (itemId) {
        if (existingLastUrgentSentCol !== -1) {
          const lastSentDate = row[existingLastUrgentSentCol];
          if (lastSentDate) lastUrgentSentMap.set(itemId, lastSentDate);
        }
        if (existingOnOrderFlagCol !== -1) {
          const onOrderFlag = row[existingOnOrderFlagCol];
          if (onOrderFlag) onOrderFlagMap.set(itemId, onOrderFlag);
        }
        if (existingNotesCol !== -1) {
          const note = row[existingNotesCol];
          if (note) notesMap.set(itemId, note);
        }
      }
    });
  } else {
    Logger.log("Warning: 'Item ID' column not found in existing Config data during save. Cannot preserve flags/notes. (This might be okay if sheet is empty).");
  }
  Logger.log("saveIndividualItemConfigs: Preserved lastUrgentSentMap: " + JSON.stringify(Array.from(lastUrgentSentMap.entries()))); // DEBUG
  Logger.log("saveIndividualItemConfigs: Preserved onOrderFlagMap: " + JSON.stringify(Array.from(onOrderFlagMap.entries()))); // DEBUG
  Logger.log("saveIndividualItemConfigs: Preserved notesMap: " + JSON.stringify(Array.from(notesMap.entries()))); // DEBUG


  // Define the ordered list of headers that we manage in the UI, matching the config.html table
  const managedHeaderNames = [
      "Item ID", "Urgent Threshold", "Urgent Comparison", "Daily Threshold", "Daily Comparison",
      "Notify Type", "Emails", "On Order Flag", "Last Urgent Sent", "Notes"
  ];

  // Calculate the highest index used by these managed headers in the *existing* sheet
  let maxManagedColIndex = -1;
  managedHeaderNames.forEach(name => {
      const idx = existingHeaders.indexOf(name);
      if (idx > maxManagedColIndex) {
          maxManagedColIndex = idx;
      }
  });

  const columnsToManage = maxManagedColIndex !== -1 ? maxManagedColIndex + 1 : 0;
  Logger.log("saveIndividualItemConfigs: Number of columns to manage (A to J equivalent): " + columnsToManage); // DEBUG

  // Clear only the columns that we manage, from row 2 to lastRow
  // This clears current data where our managed headers are.
  if (configSheet.getLastRow() > 1 && columnsToManage > 0) {
    configSheet.getRange(2, 1, configSheet.getLastRow() - 1, columnsToManage).clearContent();
    Logger.log(`saveIndividualItemConfigs: Cleared content from row 2 to ${configSheet.getLastRow()} in columns 1 to ${columnsToManage}.`); // DEBUG
  }

  const dataToSave = [];
  if (Array.isArray(configs) && configs.length > 0) {
    // Determine the full width of the sheet needed for writing,
    // including columns beyond what we manage (e.g., Default Order Emails)
    const totalColsInSheet = configSheet.getLastColumn();
    const finalSheetWidth = Math.max(totalColsInSheet, columnsToManage); // Use the wider of current sheet or managed cols
    Logger.log("saveIndividualItemConfigs: Final sheet width for writing: " + finalSheetWidth); // DEBUG

    // Create a map to store output column indices, derived from actual headers
    const outputColumnIndices = {};
    existingHeaders.forEach((header, index) => {
        outputColumnIndices[header] = index;
    });

    configs.forEach((config, configIndex) => { // Added configIndex for debugging
      const rowData = new Array(finalSheetWidth).fill(''); // Initialize with empty strings for full final width
      
      // If there's an existing row for this item (from previous state), copy its full content first
      // This is crucial to preserve columns that aren't managed by `managedHeaderNames` (like "Default Order Emails")
      const existingRowData = existingRows.find(r => r[outputColumnIndices["Item ID"]] === config.itemId);
      if (existingRowData) {
          for (let j = 0; j < existingRowData.length; j++) {
              if (j < finalSheetWidth) { // Only copy if target column exists in our new rowData
                  rowData[j] = existingRowData[j];
              }
          }
      }

      // Assign values based on their correct column index, using values from config object
      // and preserved values for Last Urgent Sent, On Order Flag, and Notes if not set in config
      if (outputColumnIndices["Item ID"] !== -1) rowData[outputColumnIndices["Item ID"]] = config.itemId;
      if (outputColumnIndices["Urgent Threshold"] !== -1) rowData[outputColumnIndices["Urgent Threshold"]] = config.urgentThreshold;
      if (outputColumnIndices["Urgent Comparison"] !== -1) rowData[outputColumnIndices["Urgent Comparison"]] = config.urgentComparison;
      if (outputColumnIndices["Daily Threshold"] !== -1) rowData[outputColumnIndices["Daily Threshold"]] = config.dailyThreshold;
      if (outputColumnIndices["Daily Comparison"] !== -1) rowData[outputColumnIndices["Daily Comparison"]] = config.dailyComparison;
      if (outputColumnIndices["Notify Type"] !== -1) rowData[outputColumnIndices["Notify Type"]] = config.notifyType;
      if (outputColumnIndices["Emails"] !== -1) rowData[outputColumnIndices["Emails"]] = config.emails;
      
      // Use preserved values for flags/dates/notes if they were not explicitly passed from UI (e.g., config object doesn't have them)
      if (outputColumnIndices["On Order Flag"] !== -1) rowData[outputColumnIndices["On Order Flag"]] = config.onOrderFlag || onOrderFlagMap.get(config.itemId) || '';
      if (outputColumnIndices["Last Urgent Sent"] !== -1) rowData[outputColumnIndices["Last Urgent Sent"]] = config.lastUrgentSent || lastUrgentSentMap.get(config.itemId) || '';
      if (outputColumnIndices["Notes"] !== -1) rowData[outputColumnIndices["Notes"]] = config.notes || notesMap.get(config.itemId) || '';


      dataToSave.push(rowData);
      Logger.log(`saveIndividualItemConfigs: Preparing row ${configIndex + 1}: ${JSON.stringify(rowData.slice(0, columnsToManage))}`); // DEBUG: Log only managed columns
    });
  }

  if (dataToSave.length > 0) {
    // Write the new data starting from row 2.
    // Use the determined finalSheetWidth to ensure consistent column writing.
    configSheet.getRange(2, 1, dataToSave.length, finalSheetWidth).setValues(dataToSave);
  }
  
  SpreadsheetApp.flush();
  Logger.log(`saveIndividualItemConfigs: Function finished. Saved ${dataToSave.length} item configurations.`); // DEBUG
}

/**
 * Consolidated function to run all stock-related checks and email queuing.
 * This function will be called by onEdit and onFormSubmit triggers.
 * It calls functions from other .gs files (urgentEmails.gs, emails.gs).
 * Ensure these files are part of the same Apps Script project.
 */
function runStockChecks() {
  Logger.log("Running consolidated stock checks...");
  try {
    const stockStatus = getStockStatus();
    Logger.log("Stock Status Results: " + JSON.stringify(stockStatus));

    // Check if queueUrgentStockAlerts function exists before calling
    if (typeof queueUrgentStockAlerts === 'function') {
      queueUrgentStockAlerts();
      Logger.log("Urgent stock alerts queued.");
    } else {
      Logger.log("Warning: queueUrgentStockAlerts function not found. Urgent emails will not be sent.");
    }

    // Check if sendDailyStockEmail function exists before calling
    if (typeof sendDailyStockEmail === 'function') {
      sendDailyStockEmail();
      Logger.log("Daily stock summary emails processed.");
    } else {
      Logger.log("Warning: sendDailyStockEmail function not found. Daily emails will not be sent.");
    }

    // Clear Last Urgent Sent for items that are now above threshold
    clearLastUrgentSentForRecoveredStock();

  } catch (e) {
    Logger.log("Error during runStockChecks: " + e.message);
    // Optionally, send an error email to an admin here
  }
}

/**
 * Clears the 'Last Urgent Sent' date for items whose current stock quantity
 * is now above their urgent threshold. This allows new urgent alerts to be sent
 * if their stock drops again later.
 * Assumes the columns based on your provided screenshot.
 */
function clearLastUrgentSentForRecoveredStock() {
  Logger.log("clearLastUrgentSentForRecoveredStock: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker");
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker");

  if (!configSheet || !stockTrackerSheet || !inksTrackerSheet) {
    Logger.log("Error: Missing required sheets for clearLastUrgentSentForRecoveredStock.");
    throw new Error("Required sheet(s) not found for stock status calculation.");
  }

  const configData = configSheet.getDataRange().getValues();
  const configHeaders = configData[0].map(h => h.toString().trim()); // Trim headers
  const configRows = configData.slice(1);
  Logger.log("clearLastUrgentSentForRecoveredStock: Trimmed configHeaders: " + JSON.stringify(configHeaders)); // DEBUG

  // Dynamically find column indices based on headers from your screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID");
  const urgentThresholdCol = configHeaders.indexOf("Urgent Threshold");
  const urgentComparisonCol = configHeaders.indexOf("Urgent Comparison");
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent"); // Column I

  if ([configItemIdCol, urgentThresholdCol, urgentComparisonCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("Error: Missing one or more required headers in 'Config' sheet for clearing Last Urgent Sent dates. Check spelling/existence.");
    const missingHeaders = [];
    if (configHeaders.indexOf("Item ID") === -1) missingHeaders.push("Item ID");
    if (configHeaders.indexOf("Urgent Threshold") === -1) missingHeaders.push("Urgent Threshold");
    if (configHeaders.indexOf("Urgent Comparison") === -1) missingHeaders.push("Urgent Comparison");
    if (configHeaders.indexOf("Last Urgent Sent") === -1) missingHeaders.push("Last Urgent Sent");
    Logger.log("Missing headers detected by clearLastUrgentSentForRecoveredStock: " + missingHeaders.join(", "));
    throw new Error("Missing required headers in 'Config' sheet. Please check the sheet headers.");
  }

  // Get current quantities for all items from both stock sheets
  const currentQuantitiesMap = new Map(); // itemId -> qty

  const processSheetForQuantities = (sheet, qtyColumnIndex) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // No data to process
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => h.toString().trim()); // Trim headers
    const itemIdCol = headers.indexOf("Item ID");
    const qtyCol = qtyColumnIndex;

    if (itemIdCol === -1 || qtyCol === -1) {
      Logger.log(`Warning: Missing 'Item ID' or 'Quantity' header in ${sheet.getName()} for quantity lookup.`);
      return;
    }

    values.slice(1).forEach(row => {
      if (row.length <= Math.max(itemIdCol, qtyCol)) return; // Defensive check
      const itemId = row[itemIdCol] ? row[itemIdCol].toString().trim() : '';
      const qty = row[qtyCol];
      if (itemId && typeof qty === 'number') {
        currentQuantitiesMap.set(itemId, qty);
      }
    });
  };

  processSheetForQuantities(stockTrackerSheet, 5); // Column F for Stock Tracker (index 5)
  processSheetForQuantities(inksTrackerSheet, 2); // Column C for Inks Tracker (index 2)
  Logger.log("clearLastUrgentSentForRecoveredStock: currentQuantitiesMap: " + JSON.stringify(Array.from(currentQuantitiesMap.entries()))); // DEBUG

  let updatesMade = false;
  configRows.forEach((row, index) => {
    if (row.length <= Math.max(configItemIdCol, urgentThresholdCol, urgentComparisonCol, lastUrgentSentCol)) return; // Defensive check
    
    const itemId = row[configItemIdCol] ? row[configItemIdCol].toString().trim() : '';
    const urgentThreshold = !isNaN(parseInt(row[urgentThresholdCol])) ? parseInt(row[urgentThresholdCol]) : 0;
    const urgentComparison = row[urgentComparisonCol] ? row[urgentComparisonCol].toString().trim() : 'less_than_or_equal';
    const lastUrgentSent = row[lastUrgentSentCol];

    if (itemId && lastUrgentSent) { // Only check if there's an item ID and a sent date
      const currentQty = currentQuantitiesMap.get(itemId);

      if (currentQty !== undefined) {
        let isAboveThreshold = false;
        // If the old comparison was 'less_than', then to be "above" it, new qty must be >= threshold
        // If the old comparison was 'less_than_or_equal', then to be "above" it, new qty must be > threshold
        if (urgentComparison === 'less_than') {
            isAboveThreshold = (currentQty >= urgentThreshold); 
        } else { // default to less_than_or_equal
            isAboveThreshold = (currentQty > urgentThreshold); 
        }
        Logger.log(`clearLastUrgentSent: Item ${itemId}: Qty=${currentQty}, Thresh=${urgentThreshold}, Comp=${urgentComparison}, isAbove=${isAboveThreshold}, lastSent=${lastUrgentSent}`); // DEBUG

        if (isAboveThreshold) {
            // Clear the Last Urgent Sent date
            configSheet.getRange(index + 2, lastUrgentSentCol + 1).clearContent();
            Logger.log(`Cleared Last Urgent Sent for ${itemId} (stock recovered to ${currentQty}).`);
            updatesMade = true;
        }
      } else {
          Logger.log(`Warning: clearLastUrgentSent: Item ${itemId} from Config not found in stock sheets. Cannot check for recovery.`); // DEBUG
      }
    }
  });

  if (updatesMade) {
    SpreadsheetApp.flush();
    Logger.log("Finished clearing 'Last Urgent Sent' dates for recovered stock.");
  } else {
    Logger.log("No 'Last Urgent Sent' dates needed clearing for recovered stock.");
  }
}


/**
 * Simple trigger that runs when a spreadsheet is edited.
 * It checks if the edit was in a relevant stock sheet and triggers stock checks.
 * @param {Object} e The event object.
 */
function onEdit(e) {
  Logger.log("onEdit: Trigger fired."); // DEBUG
  if (!e || !e.range || !e.range.getSheet) {
    Logger.log("onEdit: Event object or range is invalid.");
    return;
  }
  const sheetName = e.range.getSheet().getName().trim();
  // List of sheets where an edit should trigger a stock check
  const allowedSheets = ["Stock Tracker", "Inks Tracker", "Materials Log", "Inks Log", "Config"]; 

  if (allowedSheets.includes(sheetName)) {
    Logger.log(`onEdit: Edit detected in ${sheetName}. Running stock checks.`);
    runStockChecks();
  } else {
    Logger.log(`onEdit: Edit in unrelated sheet: ${sheetName}. Skipping stock checks.`); // DEBUG
  }
}

/**
 * Simple trigger that runs when a form is submitted to the spreadsheet.
 * This function should be linked to an installable trigger for "On form submit".
 * @param {Object} e The event object.
 */
function onFormSubmit(e) {
  Logger.log("onFormSubmit: Trigger fired."); // DEBUG
  runStockChecks();
}
