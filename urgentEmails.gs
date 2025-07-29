function queueUrgentStockAlerts() {
  Logger.log("queueUrgentStockAlerts: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const queueSheet = ss.getSheetByName("Urgent Queue");
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker"); // Needed for getItemQuantity
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker"); // Needed for getItemQuantity

  if (!configSheet || !queueSheet || !stockTrackerSheet || !inksTrackerSheet) {
    Logger.log("❌ Missing required sheet(s) for urgent stock alerts (Config, Urgent Queue, Stock Tracker, or Inks Tracker)."); // DEBUG
    return;
  }

  const configData = configSheet.getDataRange().getValues(); // Get all config data including headers
  const configHeaders = configData[0].map(h => h.toString().trim()); // Get headers and trim
  Logger.log("Debug: Config Headers: " + JSON.stringify(configHeaders)); // DEBUG
  const configRows = configData.slice(1);

  // Dynamically find column indices based on headers
  const configItemIdCol = configHeaders.indexOf("Item ID"); // A
  const urgentThresholdCol = configHeaders.indexOf("Urgent Threshold"); // B
  const urgentComparisonCol = configHeaders.indexOf("Urgent Comparison"); // C
  const notifyTypeCol = configHeaders.indexOf("Notify Type"); // F
  const emailsCol = configHeaders.indexOf("Emails"); // G
  const onOrderFlagCol = configHeaders.indexOf("On Order Flag"); // H
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent"); // I

  if ([configItemIdCol, urgentThresholdCol, urgentComparisonCol, notifyTypeCol, emailsCol, onOrderFlagCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing one or more required headers in 'Config' sheet for urgent stock alerts. Check spelling/existence."); // DEBUG
    return;
  }
  Logger.log("Debug: All required config headers found."); // DEBUG

  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24 hours ago

  configRows.forEach((row, rowIndex) => { // rowIndex is 0-based for configRows array
    // Defensive check for row length
    const maxConfigColIndex = Math.max(configItemIdCol, urgentThresholdCol, urgentComparisonCol, notifyTypeCol, emailsCol, onOrderFlagCol, lastUrgentSentCol);
    if (row.length <= maxConfigColIndex) {
      Logger.log(`Warning: queueUrgentStockAlerts: Skipping config row ${rowIndex + 2} due to insufficient columns. Row data: ${JSON.stringify(row)}`); // DEBUG
      return;
    }

    const itemId = row[configItemIdCol] ? row[configItemIdCol].toString().trim() : '';
    const urgentThreshold = !isNaN(parseInt(row[urgentThresholdCol])) ? parseInt(row[urgentThresholdCol]) : 0;
    const urgentComparison = row[urgentComparisonCol] ? row[urgentComparisonCol].toString().trim() : 'less_than_or_equal';
    const notifyType = String(row[notifyTypeCol] || '').toLowerCase();
    const emails = row[emailsCol] ? row[emailsCol].toString().trim() : '';
    const lastUrgentSent = row[lastUrgentSentCol];
    const onOrderFlag = row[onOrderFlagCol] ? row[onOrderFlagCol].toString().trim() : ''; // Get the raw value and trim it

    Logger.log(`Debug: Processing item ${itemId} (row ${rowIndex + 2}): NotifyType='${notifyType}', Emails='${emails}', Raw OnOrderFlag: '${row[onOrderFlagCol]}', Trimmed OnOrderFlag: '${onOrderFlag}', LastUrgentSent='${lastUrgentSent}'`); // DEBUG

    if (!itemId || !['urgent', 'both'].includes(notifyType) || !emails) {
      Logger.log(`Debug: Skipping ${itemId} due to missing ID, irrelevant notify type, or no emails.`); // DEBUG
      return;
    }

    // Check if an urgent email was sent recently (within 24 hours) for this item
    let sentRecently = false;
    if (lastUrgentSent instanceof Date) {
      if (lastUrgentSent.getTime() > oneDayAgo.getTime()) {
        sentRecently = true;
      }
    }
    Logger.log(`Debug: Item ${itemId}: SentRecently=${sentRecently}`); // DEBUG

    const currentQty = getItemQuantity(itemId);
    if (currentQty === undefined) {
      Logger.log(`Warning: Item ${itemId} quantity not found in stock sheets. Skipping.`); // DEBUG
      return;
    }
    Logger.log(`Debug: Item ${itemId}: CurrentQty=${currentQty}, UrgentThreshold=${urgentThreshold}, UrgentComparison='${urgentComparison}'`); // DEBUG

    // Evaluate Urgent Threshold based on comparison type
    let isUrgent = false;
    if (urgentComparison === 'less_than') {
        isUrgent = (currentQty < urgentThreshold);
    } else { // default to less_than_or_equal
        isUrgent = (currentQty <= urgentThreshold);
    }
    Logger.log(`Debug: Item ${itemId}: IsUrgent=${isUrgent}`); // DEBUG

    // Suppress urgent if On Order Flag is true (case-insensitive check now)
    const suppressUrgent = (onOrderFlag.toLowerCase() === 'true'); // Changed to case-insensitive comparison
    Logger.log(`Debug: Item ${itemId}: SuppressUrgent (based on On Order Flag) = ${suppressUrgent}`); // DEBUG

    if (!isUrgent) {
      Logger.log(`Debug: Item ${itemId} is NOT urgent. Skipping queueing.`); // DEBUG
      return; // If not urgent, no need to queue
    }

    // If already sent recently OR suppressed by On Order Flag, skip queuing
    if (sentRecently) {
      Logger.log(`ℹ️ Urgent email for ${itemId} was sent recently. Skipping queueing.`);
      return;
    }
    if (suppressUrgent) {
      Logger.log(`ℹ️ Urgent email for ${itemId} suppressed because replenishment is on order (On Order Flag is 'true').`); // Updated log message for clarity
      return; // Skip this item due to On Order Flag
    }

    const lastQueueRow = queueSheet.getLastRow();
    let alreadyQueued = false;

    if (lastQueueRow > 1) {
      const queuedItems = queueSheet.getRange(2, 1, lastQueueRow - 1, 1).getValues().flat();
      alreadyQueued = queuedItems.includes(itemId);
    }
    Logger.log(`Debug: Item ${itemId}: AlreadyQueued=${alreadyQueued}`); // DEBUG

    if (!alreadyQueued) {
      queueSheet.appendRow([itemId, currentQty, urgentThreshold, now]);
      Logger.log(`✅ Queued urgent alert for ${itemId}.`);
    } else {
      Logger.log(`ℹ️ ${itemId} is already in the urgent queue.`);
    }
  });

  ensureUrgentEmailTriggerExists();
  Logger.log("queueUrgentStockAlerts: Function finished."); // DEBUG
}

function ensureUrgentEmailTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Delete any existing triggers for sendBatchedUrgentEmail
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendBatchedUrgentEmail") {
      Logger.log(`⚠️ Deleting existing trigger for sendBatchedUrgentEmail (to ensure fresh 5-min trigger).`);
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create a new one-time trigger that fires 5 minutes from now
  ScriptApp.newTrigger("sendBatchedUrgentEmail")
           .timeBased()
           .after(1 * 60 * 1000) // 5 minutes in milliseconds
           .create();
  Logger.log("✅ Created new ONE-TIME time-based trigger for sendBatchedUrgentEmail (fires in 5 minutes).");
}


function sendBatchedUrgentEmail() {
  Logger.log("sendBatchedUrgentEmail: Function started."); // DEBUG
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("Urgent Queue");
  const configSheet = ss.getSheetByName("Config");
  if (!queueSheet || !configSheet) {
    Logger.log("❌ Missing required sheet(s) for sending batched urgent email (Urgent Queue or Config)."); // DEBUG
    return;
  }

  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No items in Urgent Queue. Skipping email."); // DEBUG
    return;
  }

  const queueData = queueSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  Logger.log("Debug: Raw queueData from Urgent Queue: " + JSON.stringify(queueData)); // DEBUG
  
  const fullConfigData = configSheet.getDataRange().getValues(); 
  const configHeaders = fullConfigData[0].map(h => h.toString().trim()); // Trim headers
  const configRows = fullConfigData.slice(1);

  // Dynamically find column indices based on headers from screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID");
  const emailsCol = configHeaders.indexOf("Emails");
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent");

  if ([configItemIdCol, emailsCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing required headers in 'Config' sheet for sending batched urgent email. Check spelling/existence."); // DEBUG
    return;
  }
  Logger.log("Debug: All required config headers for urgent email found."); // DEBUG

  // Use a Map to store unique urgent items by their Item ID
  const uniqueUrgentItems = new Map(); // itemId -> {itemId, qty, threshold}
  const allRecipients = new Set();

  queueData.forEach(([itemId, qty, threshold]) => {
    const trimmedItemId = itemId.toString().trim();
    // Only add if not already in the map, or update if a later entry has different details (though usually they should be same for same item)
    if (!uniqueUrgentItems.has(trimmedItemId)) {
      uniqueUrgentItems.set(trimmedItemId, {
        itemId: trimmedItemId,
        qty: Number(qty) || 0, // Ensure quantity is a number
        threshold: Number(threshold) || 0 // Ensure threshold is a number
      });
      Logger.log(`Debug: Added unique item to urgent email map: ${trimmedItemId}`); // DEBUG
    } else {
      Logger.log(`Debug: Skipping duplicate item in queue data for email: ${trimmedItemId}`); // DEBUG
    }

    const configRow = configRows.find(row => row[configItemIdCol] && row[configItemIdCol].toString().trim() === trimmedItemId);
    if (configRow) {
      const itemEmailsRaw = configRow[emailsCol];
      if (typeof itemEmailsRaw === 'string' && itemEmailsRaw.trim() !== '') {
        itemEmailsRaw.split(",").map(e => e.trim()).filter(Boolean).forEach(e => allRecipients.add(e));
      }
    } else {
      Logger.log(`Warning: Config row not found for item ${trimmedItemId} while collecting recipients for urgent email.`); // DEBUG
    }
  });

  const urgentItemsForEmail = Array.from(uniqueUrgentItems.values()); // Convert map values to an array
  Logger.log("Debug: Final unique urgent items for email: " + JSON.stringify(urgentItemsForEmail)); // DEBUG

  const recipients = Array.from(allRecipients).join(",");
  if (!recipients) {
    Logger.log("No valid recipients found for urgent alert. Skipping email."); // DEBUG
    return;
  }
  Logger.log("Debug: Recipients for urgent email: " + recipients); // DEBUG

  try {
    const template = HtmlService.createTemplateFromFile('UrgentEmailTemplate');
    template.urgentItems = urgentItemsForEmail;
    template.appUrl = ScriptApp.getService().getUrl();
    template.googleSheetUrl = "https://docs.google.com/spreadsheets/d/1yWypw_j9PBtQRND_m6ayAr_I9gDXsJvGQ8eIu5XGEcY/edit?gid=0#gid=0";
    const htmlBody = template.evaluate().getContent();

    MailApp.sendEmail({
      to: recipients,
      subject: "🚨 Urgent Inventory Alert – Immediate Attention Required",
      htmlBody: htmlBody
    });
    Logger.log(`✅ Urgent stock alert email sent to: ${recipients}`);

    const now = new Date();
    // Update 'Last Urgent Sent' only for items that were *actually* included in this batched email
    urgentItemsForEmail.forEach(item => {
      const itemId = item.itemId; // Use the itemId from the unique list
      const actualConfigRowIndex = fullConfigData.findIndex(row => row[configItemIdCol] && row[configItemIdCol].toString().trim() === itemId);
      
      if (actualConfigRowIndex !== -1) {
        // +1 to convert 0-based array index to 1-based sheet row number
        // +1 to convert 0-based column index to 1-based sheet column number
        configSheet.getRange(actualConfigRowIndex + 1, lastUrgentSentCol + 1).setValue(now);
        Logger.log(`Debug: Updated Last Urgent Sent for ${itemId} in Config sheet.`); // DEBUG
      } else {
        Logger.log(`Warning: Item ID ${itemId} from sent urgent email not found in Config sheet for Last Urgent Sent update.`); // DEBUG
      }
    });

    // Clear the queue sheet only after successful email sending and Last Urgent Sent updates
    queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).clearContent();
    Logger.log("Urgent queue cleared."); // DEBUG
  } catch (e) {
    Logger.log(`❌ Error sending urgent stock alert email: ${e.message}. Stack: ${e.stack}`); // Log stack trace for more details
  }
  Logger.log("sendBatchedUrgentEmail: Function finished."); // DEBUG
}