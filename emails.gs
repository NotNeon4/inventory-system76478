function getItemQuantity(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("Stock Tracker");
  const inksSheet = ss.getSheetByName("Inks Tracker");

  // Stock Tracker (Quantity in column F, index 5)
  if (stockSheet) {
    const stockData = stockSheet.getDataRange().getValues().slice(1); // skip header
    for (let row of stockData) {
      if (row[0] === itemId) {
        return Number(row[5]) || 0;
      }
    }
  }

  // Inks Tracker (Quantity in column C, index 2)
  if (inksSheet) {
    const inksData = inksSheet.getDataRange().getValues().slice(1);
    for (let row of inksData) {
      if (row[0] === itemId) {
        return Number(row[2]) || 0;
      }
    }
  }

  return undefined;
}

function sendDailyStockEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) return;

  const config = configSheet.getDataRange().getValues().slice(1);
  const lowStockItemsByEmail = {};

  config.forEach(row => {
    const [itemId, , dailyThreshold, notifyType, , , email] = row;
    const type = (notifyType || "").toString().toLowerCase();

    if (!itemId || !["daily", "both"].includes(type) || !email) return;

    const currentQty = getItemQuantity(itemId);
    if (currentQty === undefined || currentQty > dailyThreshold) return;

    if (!lowStockItemsByEmail[email]) {
      lowStockItemsByEmail[email] = [];
    }

    lowStockItemsByEmail[email].push(
      `📝 ${itemId} - Current: ${currentQty}, Daily Threshold: ${dailyThreshold}`
    );
  });

  Object.keys(lowStockItemsByEmail).forEach(email => {
    const message = lowStockItemsByEmail[email].join('\n');
    MailApp.sendEmail({
      to: email,
      subject: "📬 Daily Stock Summary",
      body: `Here are today's low stock items:\n\n${message}`
    });
  });
}

function runUrgentStockCheck() {
  queueUrgentStockAlerts();
  SpreadsheetApp.getUi().alert("✅ Urgent stock check complete.\nIf any items are low, an email will be sent in 5 minutes.");
}

function runDailySummaryEmail() {
  sendDailyStockSummary();
  SpreadsheetApp.getUi().alert("✅ Daily stock summary email sent.");
}

function processOrder(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Orders") || ss.insertSheet("Orders");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Item", "Quantity", "Emails"]);
  }

  const timestamp = new Date();

  // Force arrays (even if only 1 row)
  const items = [].concat(formData["item[]"]);
  const quantities = [].concat(formData["quantity[]"]);
  const emails = formData["emails"];

  for (let i = 0; i < items.length; i++) {
    if (!items[i]) continue;

    sheet.appendRow([
      timestamp,
      items[i],
      quantities[i] || "",
      emails
    ]);
  }

  // Send summary email
  if (emails) {
    const rows = items.map((item, i) => `<li>${item}: ${quantities[i]}</li>`).join("");
    const htmlBody = `
      <p>📦 A new inventory order was submitted:</p>
      <ul>${rows}</ul>
      <p>Sent to: <b>${emails}</b></p>
    `;

    MailApp.sendEmail({
      to: emails,
      subject: `🧾 New Inventory Order Submitted`,
      htmlBody
    });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

