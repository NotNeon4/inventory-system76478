function queueUrgentStockAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const queueSheet = ss.getSheetByName("Urgent Queue");
  if (!configSheet || !queueSheet) return;

  const config = configSheet.getDataRange().getValues().slice(1); // Skip header
  const now = new Date();

  config.forEach(row => {
    const [itemId, urgentThreshold, , typeRaw] = row;
    const type = String(typeRaw || '').toLowerCase();

    if (!itemId || !['urgent', 'both'].includes(type)) return;

    const currentQty = getItemQuantity(itemId);
    if (currentQty === undefined || currentQty > urgentThreshold) return;

    const lastRow = queueSheet.getLastRow();
    let alreadyQueued = false;

    if (lastRow > 1) {
      const queuedItems = queueSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      alreadyQueued = queuedItems.includes(itemId);
    }

    if (!alreadyQueued) {
      queueSheet.appendRow([itemId, currentQty, urgentThreshold, now]);
    }
  });

  // Ensure a valid 5-min delayed trigger exists
  ensureUrgentEmailTriggerExists();
}


function ensureUrgentEmailTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendBatchedUrgentEmail") {
      try {
        // Validate trigger by checking its source
        const triggerSource = trigger.getTriggerSource(); // throws if invalid
        found = true;
      } catch (e) {
        // It's a broken trigger — remove it
        ScriptApp.deleteTrigger(trigger);
        found = false;
      }
    }
  });

  if (!found) {
    ScriptApp.newTrigger("sendBatchedUrgentEmail")
             .timeBased()
             .after(5 * 60 * 1000)
             .create();
    Logger.log("✅ Created new time-based trigger for sendBatchedUrgentEmail.");
  } else {
    Logger.log("⏱️ Valid trigger already exists for sendBatchedUrgentEmail.");
  }
}


function sendBatchedUrgentEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("Urgent Queue");
  const configSheet = ss.getSheetByName("Config");
  if (!queueSheet || !configSheet) return;

  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) return;

  const queueData = queueSheet.getRange(2, 1, lastRow - 1, 4).getValues(); // [itemId, qty, threshold, timestamp]
  const configData = configSheet.getDataRange().getValues(); // include header

  const itemIds = [];
  let body = "🚨 *URGENT STOCK ALERT*\n\nThe following items are below their urgent thresholds:\n\n";

  queueData.forEach(([itemId, qty, threshold]) => {
    itemIds.push(itemId);
    body += `🔸 *${itemId}*\n     • Quantity: ${qty}\n     • Threshold: ${threshold}\n\n`;
  });

  const emailsSet = new Set();

  for (let i = 1; i < configData.length; i++) {
    const [itemId, , , , , , emailRaw] = configData[i];
    if (itemIds.includes(itemId) && typeof emailRaw === 'string') {
      emailRaw.split(",").map(e => e.trim()).forEach(e => emailsSet.add(e));
    }
  }

  const recipients = Array.from(emailsSet).join(",");
  if (!recipients) return;

  MailApp.sendEmail({
    to: recipients,
    subject: "🚨 Urgent Inventory Alert – Immediate Attention Required",
    body: body
  });

  // ✅ Update "Last Urgent Sent" (column F = index 5, zero-based)
  const now = new Date();
  for (let i = 1; i < configData.length; i++) {
    const [itemId] = configData[i];
    if (itemIds.includes(itemId)) {
      configSheet.getRange(i + 1, 6).setValue(now); // row = i+1 (because of header), col = 6 (F)
    }
  }

  // ✅ Clear the urgent queue
  queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).clearContent();
}

