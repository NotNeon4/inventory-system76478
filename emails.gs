function getItemQuantity(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Stock Tracker");
  const inksSheet = ss.getSheetByName("Inks Tracker");

  // Stock Tracker (Quantity in column F, index 5)
  if (stockSheet) {
    const stockData = stockSheet.getDataRange().getValues().slice(1); // skip header
    const stockHeaders = stockSheet.getDataRange().getValues()[0];
    const itemIdCol = stockHeaders.indexOf("Item ID");
    const qtyCol = stockHeaders.indexOf("Quantity"); // Assuming "Quantity" is the header for Column F

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
    const qtyCol = inksHeaders.indexOf("Quantity"); // Assuming "Quantity" is the header for Column C

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker");
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker");

  if (!configSheet || !stockTrackerSheet || !inksTrackerSheet) {
    Logger.log("❌ Missing required sheet(s) for daily stock email (Config, Stock Tracker, or Inks Tracker).");
    return;
  }

  const configData = configSheet.getDataRange().getValues();
  const configHeaders = configData[0];
  const itemThresholdsMap = new Map();

  // Dynamically find column indices based on headers from screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID"); // A
  const dailyThresholdCol = configHeaders.indexOf("Daily Threshold"); // D
  const dailyComparisonCol = configHeaders.indexOf("Daily Comparison"); // E
  const notifyTypeCol = configHeaders.indexOf("Notify Type"); // F
  const emailsCol = configHeaders.indexOf("Emails"); // G
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent"); // I

  if ([configItemIdCol, dailyThresholdCol, dailyComparisonCol, notifyTypeCol, emailsCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing one or more required headers in 'Config' sheet for daily stock email based on screenshot.");
    return;
  }

  // Populate itemThresholdsMap from Config sheet
  configData.slice(1).forEach(row => { // Skip header row
    const itemId = row[configItemIdCol];
    if (itemId) {
      const daily = !isNaN(parseInt(row[dailyThresholdCol])) ? parseInt(row[dailyThresholdCol]) : 5;
      const dailyComp = row[dailyComparisonCol] ? row[dailyComparisonCol].toString().trim() : 'less_than_or_equal';
      const notifyType = row[notifyTypeCol] ? row[notifyTypeCol].toString().trim().toLowerCase() : 'both';
      const emails = row[emailsCol] ? row[emailsCol].toString().trim() : '';
      const lastUrgentSent = row[lastUrgentSentCol]; // Get the date

      itemThresholdsMap.set(itemId.toString().trim(), { daily, dailyComparison: dailyComp, notifyType, emails, lastUrgentSent });
    }
  });

  const lowStockItemsByRecipient = new Map();
  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24 hours ago

  // Helper function to process a single stock sheet
  const processStockSheet = (sheet, qtyColumnIndex) => {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const stockValues = sheet.getDataRange().getValues();
      const stockHeaders = stockValues[0];
      const stockDataRows = stockValues.slice(1);

      const stockItemIdCol = stockHeaders.indexOf("Item ID");
      const stockQtyCol = qtyColumnIndex; 

      if (stockItemIdCol === -1 || stockQtyCol === -1) {
        Logger.log(`Warning: Missing 'Item ID' or 'Qty' header in ${sheet.getName()} for daily email check.`);
        return;
      }

      stockDataRows.forEach(row => {
        const itemId = row[stockItemIdCol];
        const qty = row[stockQtyCol];

        if (typeof itemId === 'string' && itemId.trim() !== '' && typeof qty === 'number') {
          const config = itemThresholdsMap.get(itemId.trim());

          if (config && ['daily', 'both'].includes(config.notifyType)) {
            // Check if an urgent email was sent recently for this item
            let urgentEmailSentRecently = false;
            if (config.lastUrgentSent instanceof Date) {
              if (config.lastUrgentSent.getTime() > oneDayAgo.getTime()) {
                urgentEmailSentRecently = true;
              }
            }

            // If an urgent email was sent recently, suppress the daily email for this item
            if (urgentEmailSentRecently) {
              Logger.log(`ℹ️ Daily email for ${itemId} suppressed as urgent email was sent recently.`);
              return; // Skip this item for daily email
            }

            let isDailyLow = false;
            if (config.dailyComparison === 'less_than') {
                isDailyLow = (qty < config.daily); 
            } else { // default to less_than_or_equal
                isDailyLow = (qty <= config.daily);
            }
            
            if (isDailyLow) { 
              const itemDetails = { itemId: itemId, qty: qty, threshold: config.daily };
              const recipients = config.emails.split(',').map(e => e.trim()).filter(Boolean);

              recipients.forEach(email => {
                if (!lowStockItemsByRecipient.has(email)) {
                  lowStockItemsByRecipient.set(email, []);
                }
                lowStockItemsByRecipient.get(email).push(itemDetails);
              });
            }
          }
        }
      });
    } else {
      Logger.log(`No data in ${sheet.getName()} to check for daily low stock.`);
    }
  };

  // Process both stock sheets
  processStockSheet(stockTrackerSheet, 5); // Assuming Qty is in Column F (index 5) for Stock Tracker
  processStockSheet(inksTrackerSheet, 2); // Assuming Qty is in Column C (index 2) for Inks Tracker

  if (lowStockItemsByRecipient.size === 0) {
    Logger.log("No items found below daily threshold for any recipient. Skipping daily emails.");
    return;
  }

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
}
