function queueUrgentStockAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const queueSheet = ss.getSheetByName("Urgent Queue");
  const stockTrackerSheet = ss.getSheetByName("Stock Tracker"); // Needed for getItemQuantity
  const inksTrackerSheet = ss.getSheetByName("Inks Tracker"); // Needed for getItemQuantity

  if (!configSheet || !queueSheet || !stockTrackerSheet || !inksTrackerSheet) {
    Logger.log("❌ Missing required sheet(s) for urgent stock alerts.");
    return;
  }

  const configData = configSheet.getDataRange().getValues(); // Get all config data including headers
  const configHeaders = configData[0];
  const configRows = configData.slice(1);

  // Dynamically find column indices based on headers from screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID"); // A
  const urgentThresholdCol = configHeaders.indexOf("Urgent Threshold"); // B
  const urgentComparisonCol = configHeaders.indexOf("Urgent Comparison"); // C
  const notifyTypeCol = configHeaders.indexOf("Notify Type"); // F
  const emailsCol = configHeaders.indexOf("Emails"); // G
  const onOrderFlagCol = configHeaders.indexOf("On Order Flag"); // H
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent"); // I

  if ([configItemIdCol, urgentThresholdCol, urgentComparisonCol, notifyTypeCol, emailsCol, onOrderFlagCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing one or more required headers in 'Config' sheet for urgent stock alerts based on screenshot.");
    return;
  }

  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24 hours ago

  configRows.forEach((row, index) => { // index is 0-based for configRows array
    const itemId = row[configItemIdCol];
    const urgentThreshold = !isNaN(parseInt(row[urgentThresholdCol])) ? parseInt(row[urgentThresholdCol]) : 0;
    const urgentComparison = row[urgentComparisonCol] ? row[urgentComparisonCol].toString().trim() : 'less_than_or_equal';
    const notifyType = String(row[notifyTypeCol] || '').toLowerCase();
    const emails = row[emailsCol] ? row[emailsCol].toString().trim() : '';
    const lastUrgentSent = row[lastUrgentSentCol];
    const onOrderFlag = row[onOrderFlagCol] ? row[onOrderFlagCol].toString().trim() : '';

    if (!itemId || !['urgent', 'both'].includes(notifyType) || !emails) return;

    // Check if an urgent email was sent recently (within 24 hours) for this item
    let sentRecently = false;
    if (lastUrgentSent instanceof Date) {
      if (lastUrgentSent.getTime() > oneDayAgo.getTime()) {
        sentRecently = true;
      }
    }

    const currentQty = getItemQuantity(itemId);
    if (currentQty === undefined) return;

    // Evaluate Urgent Threshold based on comparison type
    let isUrgent = false;
    if (urgentComparison === 'less_than') {
        isUrgent = (currentQty < urgentThreshold);
    } else { // default to less_than_or_equal
        isUrgent = (currentQty <= urgentThreshold);
    }

    // Suppress urgent if On Order Flag is true
    const suppressUrgent = (onOrderFlag === 'TRUE');

    if (!isUrgent) return; // If not urgent, no need to queue

    // If already sent recently OR suppressed by On Order Flag, skip queuing
    if (sentRecently) {
      Logger.log(`ℹ️ Urgent email for ${itemId} was sent recently. Skipping queueing.`);
      return;
    }
    if (suppressUrgent) {
      Logger.log(`ℹ️ Urgent email for ${itemId} suppressed because replenishment is on order.`);
      return; // Skip this item due to On Order Flag
    }

    const lastQueueRow = queueSheet.getLastRow();
    let alreadyQueued = false;

    if (lastQueueRow > 1) {
      const queuedItems = queueSheet.getRange(2, 1, lastQueueRow - 1, 1).getValues().flat();
      alreadyQueued = queuedItems.includes(itemId);
    }

    if (!alreadyQueued) {
      queueSheet.appendRow([itemId, currentQty, urgentThreshold, now]);
      Logger.log(`✅ Queued urgent alert for ${itemId}.`);
    } else {
      Logger.log(`ℹ️ ${itemId} is already in the urgent queue.`);
    }
  });

  ensureUrgentEmailTriggerExists();
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
           .after(5 * 60 * 1000) // 5 minutes in milliseconds
           .create();
  Logger.log("✅ Created new ONE-TIME time-based trigger for sendBatchedUrgentEmail (fires in 5 minutes).");
}


function sendBatchedUrgentEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("Urgent Queue");
  const configSheet = ss.getSheetByName("Config");
  if (!queueSheet || !configSheet) {
    Logger.log("❌ Missing required sheet(s) for sending batched urgent email.");
    return;
  }

  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No items in Urgent Queue. Skipping email.");
    return;
  }

  const queueData = queueSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  const fullConfigData = configSheet.getDataRange().getValues(); 
  const configHeaders = fullConfigData[0];
  const configRows = fullConfigData.slice(1);

  // Dynamically find column indices based on headers from screenshot
  const configItemIdCol = configHeaders.indexOf("Item ID"); // A
  const emailsCol = configHeaders.indexOf("Emails"); // G
  const lastUrgentSentCol = configHeaders.indexOf("Last Urgent Sent"); // I

  if ([configItemIdCol, emailsCol, lastUrgentSentCol].some(col => col === -1)) {
    Logger.log("❌ Missing required headers in 'Config' sheet for sending batched urgent email based on screenshot.");
    return;
  }

  const urgentItemsForEmail = [];
  const allRecipients = new Set();

  queueData.forEach(([itemId, qty, threshold]) => {
    urgentItemsForEmail.push({
      itemId: itemId,
      qty: qty,
      threshold: threshold
    });

    const configRow = configRows.find(row => row[configItemIdCol] === itemId);
    if (configRow) {
      const itemEmailsRaw = configRow[emailsCol];
      if (typeof itemEmailsRaw === 'string' && itemEmailsRaw.trim() !== '') {
        itemEmailsRaw.split(",").map(e => e.trim()).forEach(e => allRecipients.add(e));
      }
    }
  });

  const recipients = Array.from(allRecipients).join(",");
  if (!recipients) {
    Logger.log("No valid recipients found for urgent alert. Skipping email.");
    return;
  }

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
    queueData.forEach(([itemId]) => {
      const actualConfigRowIndex = fullConfigData.findIndex(row => row[configItemIdCol] === itemId);
      
      if (actualConfigRowIndex !== -1) {
        configSheet.getRange(actualConfigRowIndex + 1, lastUrgentSentCol + 1).setValue(now);
        Logger.log(`Updated Last Urgent Sent for ${itemId} in Config sheet.`);
      } else {
        Logger.log(`Warning: Item ID ${itemId} from queue not found in Config sheet for Last Urgent Sent update.`);
      }
    });

    queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).clearContent();
    Logger.log("Urgent queue cleared.");

  } catch (e) {
    Logger.log(`❌ Error sending urgent stock alert email: ${e.message}`);
  }
}
