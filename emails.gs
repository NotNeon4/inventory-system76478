function getItemQuantity(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Stock Tracker");
  const inksSheet = ss.getSheetByName("Inks Tracker");
  // Stock Tracker (Quantity in column F, index 5)
  if (stockSheet) {
    const stockData = stockSheet.getDataRange().getValues().slice(1);
    // skip header
    const stockHeaders = stockSheet.getDataRange().getValues()[0];
    const itemIdCol = stockHeaders.indexOf("Item ID");
    const qtyCol = stockHeaders.indexOf("Quantity");
    // Assuming "Quantity" is the header for Column F

    if (itemIdCol !== -1 && qtyCol !== -1) {
      for (let row of stockData) {
        if (row[itemIdCol] && row[itemIdCol].toString().trim() === itemId.toString().trim()) {
          return Number(row[qtyCol]) || 0;
        }
      }
    } else {
      Logger.log("Warning: 'Item ID' or 'Quantity' header not found in Stock Tracker.");
    }
  }

  // Inks Tracker (Quantity in column C, index 2)
  if (inksSheet) {
    const inksData = inksSheet.getDataRange().getValues().slice(1);
    const inksHeaders = inksSheet.getDataRange().getValues()[0];
    const itemIdCol = inksHeaders.indexOf("Item ID");
    const qtyCol = inksHeaders.indexOf("Quantity");
    // Assuming "Quantity" is the header for Column C

    if (itemIdCol !== -1 && qtyCol !== -1) {
      for (let row of inksData) {
        if (row[itemIdCol] && row[itemIdCol].toString().trim() === itemId.toString().trim()) {
          return Number(row[qtyCol]) || 0;
        }
      }
    } else {
      Logger.log("Warning: 'Item ID' or 'Quantity' header not found in Inks Tracker.");
    }
  }

  return undefined; // Item not found in either sheet
}

function sendDailyStockEmail() {
  Logger.log("sendDailyStockEmail: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker");
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker");
  const urgentQueueSheet = ss.getSheetByName("Urgent Queue"); // NEW: Get Urgent Queue Sheet

  if (!configSheet || !stockTrackerSheet || !inksTrackerSheet || !urgentQueueSheet) {
    Logger.log("❌ Missing required sheet(s) for daily stock email (Config, Stock Tracker, Inks Tracker, or Urgent Queue)."); // DEBUG
    return;
  }

  const configData = configSheet.getDataRange().getValues();
  const configHeaders = configData[0].map(h => h.toString().trim()); // Trim headers
  const itemThresholdsMap = new Map();

  // Dynamically find column indices based on headers from screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID");
  const dailyThresholdCol = configHeaders.indexOf("Daily Threshold");
  const dailyComparisonCol = configHeaders.indexOf("Daily Comparison");
  const notifyTypeCol = configHeaders.indexOf("Notify Type");
  const emailsCol = configHeaders.indexOf("Emails");
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent");

  if ([configItemIdCol, dailyThresholdCol, dailyComparisonCol, notifyTypeCol, emailsCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing one or more required headers in 'Config' sheet for daily stock email. Check spelling/existence."); // DEBUG
    return;
  }
  Logger.log("Debug: All required config headers for daily email found."); // DEBUG

  // Populate itemThresholdsMap from Config sheet
  configData.slice(1).forEach(row => { // Skip header row
    // Defensive check for row length
    const maxConfigColIndex = Math.max(configItemIdCol, dailyThresholdCol, dailyComparisonCol, notifyTypeCol, emailsCol, lastUrgentSentCol);
    if (row.length <= maxConfigColIndex) {
      Logger.log(`Warning: Daily Email: Skipping config row due to insufficient columns: ${JSON.stringify(row)}`); // DEBUG
      return;
    }
    const itemId = row[configItemIdCol];
    if (itemId) {
      const daily = !isNaN(parseInt(row[dailyThresholdCol])) ? parseInt(row[dailyThresholdCol]) : 5;
      const dailyComp = row[dailyComparisonCol] ? row[dailyComparisonCol].toString().trim() : 'less_than_or_equal';
      const notifyType = row[notifyTypeCol] ? row[notifyTypeCol].toString().trim().toLowerCase() : 'both';
      const emails = row[emailsCol] ? row[emailsCol].toString().trim() : '';
      const lastUrgentSent = row[lastUrgentSentCol]; // Get the date

      itemThresholdsMap.set(itemId.toString().trim(), { daily, dailyComparison: dailyComp, notifyType, emails, lastUrgentSent });
      Logger.log(`Debug: Daily Email: Mapped config for ${itemId}: ${JSON.stringify(itemThresholdsMap.get(itemId.toString().trim()))}`); // DEBUG
    }
  });

  const lowStockItemsByRecipient = new Map();
  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24 hours ago

  // NEW: Get items currently in the Urgent Queue
  const urgentQueueItems = new Set();
  const lastQueueRow = urgentQueueSheet.getLastRow();
  if (lastQueueRow > 1) {
    const queueData = urgentQueueSheet.getRange(2, 1, lastQueueRow - 1, 1).getValues();
    queueData.forEach(row => {
      if (row[0]) urgentQueueItems.add(row[0].toString().trim());
    });
  }
  Logger.log(`Debug: Daily Email: Items currently in Urgent Queue: ${JSON.stringify(Array.from(urgentQueueItems))}`); // DEBUG


  // Helper function to process a single stock sheet
  const processStockSheet = (sheet, qtyColumnIndex) => {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const stockValues = sheet.getDataRange().getValues();
      const stockHeaders = stockValues[0].map(h => h.toString().trim()); // Trim headers
      const stockDataRows = stockValues.slice(1);

      const stockItemIdCol = stockHeaders.indexOf("Item ID");
      const stockQtyCol = qtyColumnIndex;
      if (stockItemIdCol === -1 || stockQtyCol === -1) {
        Logger.log(`Warning: Daily Email: Missing 'Item ID' or 'Qty' header in ${sheet.getName()}.`); // DEBUG
        return;
      }

      stockDataRows.forEach((row, rowIndex) => { // Added rowIndex for debug
        // Defensive check for row length
        if (row.length <= Math.max(stockItemIdCol, stockQtyCol)) {
          Logger.log(`Warning: Daily Email: Skipping row ${rowIndex + 2} in ${sheet.getName()} due to insufficient columns: ${JSON.stringify(row)}`); // DEBUG
          return;
        }

        const itemId = row[stockItemIdCol] ? row[stockItemIdCol].toString().trim() : '';
        const qty = row[stockQtyCol];

        Logger.log(`Debug: Daily Email: Checking ${sheet.getName()} row ${rowIndex + 2}: ItemID=${itemId}, Qty=${qty}`); // DEBUG

        if (typeof itemId === 'string' && itemId.trim() !== '' && typeof qty === 'number') {
          const config = itemThresholdsMap.get(itemId.trim());

          if (config && ['daily', 'both'].includes(config.notifyType)) {
            // Check if an urgent email was sent recently for this item (from Config sheet)
            let urgentEmailSentRecently = false;
            if (config.lastUrgentSent instanceof Date) {
              if (config.lastUrgentSent.getTime() > oneDayAgo.getTime()) {
                urgentEmailSentRecently = true;
              }
            }
            Logger.log(`Debug: Daily Email: Item ${itemId}: urgentEmailSentRecently (from Config)=${urgentEmailSentRecently}`); // DEBUG

            // NEW: Check if item is currently in the Urgent Queue
            const isCurrentlyUrgent = urgentQueueItems.has(itemId);
            Logger.log(`Debug: Daily Email: Item ${itemId}: isCurrentlyUrgent (from Queue)=${isCurrentlyUrgent}`); // DEBUG

            // If an urgent email was sent recently OR is currently being queued, suppress the daily email for this item
            if (urgentEmailSentRecently || isCurrentlyUrgent) {
              Logger.log(`ℹ️ Daily email for ${itemId} suppressed as urgent email was sent recently or is currently queued.`);
              return; // Skip this item for daily email
            }

            let isDailyLow = false;
            if (config.dailyComparison === 'less_than') {
                isDailyLow = (qty < config.daily);
            } else { // default to less_than_or_equal
                isDailyLow = (qty <= config.daily);
            }
            Logger.log(`Debug: Daily Email: Item ${itemId}: IsDailyLow=${isDailyLow}`); // DEBUG
            
            if (isDailyLow) { 
              const itemDetails = { itemId: itemId, qty: qty, threshold: config.daily };
              const recipients = config.emails.split(',').map(e => e.trim()).filter(Boolean);

              recipients.forEach(email => {
                if (!lowStockItemsByRecipient.has(email)) {
                  lowStockItemsByRecipient.set(email, []);
                }
                lowStockItemsByRecipient.get(email).push(itemDetails);
              });
              Logger.log(`Debug: Daily Email: Item ${itemId} added to daily email list for recipients.`); // DEBUG
            }
          }
        } else {
          Logger.log(`Warning: Daily Email: Skipping row due to invalid ItemID or Quantity (ItemID: '${itemId}', Qty: '${qty}').`); // DEBUG
        }
      });
    } else {
      Logger.log(`Debug: No data in ${sheet.getName()} to check for daily low stock.`); // DEBUG
    }
  };

  // Process both stock sheets
  processStockSheet(stockTrackerSheet, 5); // Assuming Qty is in Column F (index 5) for Stock Tracker
  processStockSheet(inksTrackerSheet, 2); // Assuming Qty is in Column C (index 2) for Inks Tracker

  if (lowStockItemsByRecipient.size === 0) {
    Logger.log("No items found below daily threshold for any recipient. Skipping daily emails.");
    return;
  }
  Logger.log(`Debug: Daily Email: Sending emails to ${lowStockItemsByRecipient.size} recipients.`); // DEBUG

  // Send emails to each recipient with their specific low stock items
  lowStockItemsByRecipient.forEach((itemsForRecipient, recipientEmail) => {
    try {
      const template = HtmlService.createTemplateFromFile('DailyEmailTemplate');
      template.lowStockItems = itemsForRecipient;
      template.appUrl = ScriptApp.getService().getUrl();
      template.googleSheetUrl = "https://docs.google.com/spreadsheets/d/1yWypw_j9PBtQRND_m6ayAr_I9gDXsJvGQ8eIu5XGEcY/edit?gid=0#gid=0"; // Pass Google Sheet URL

      const htmlBody = template.evaluate().getContent();

      MailApp.sendEmail({
        to: recipientEmail,
        subject: "📊 Daily Inventory Stock Summary",
        htmlBody: htmlBody
      });
      Logger.log(`✅ Daily stock summary email sent to: ${recipientEmail}`);
    } catch (e) {
      Logger.log(`❌ Error sending daily stock summary email to ${recipientEmail}: ${e.message}`);
    }
  });
  Logger.log("sendDailyStockEmail: Function finished."); // DEBUG
}