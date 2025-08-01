// --- Constants for Sticker Layout (Adjust these if your physical setup changes) ---
// A4 dimensions in points (1 inch = 72 points, 1mm = ~2.83465 points)
const A4_WIDTH_PTS = 595.28;  // 210 mm
const A4_HEIGHT_PTS = 841.89; // 297 mm

// Specific sticker dimensions provided by user: 80mm x 45mm
const INDIVIDUAL_STICKER_WIDTH_MM = 80;
const INDIVIDUAL_STICKER_HEIGHT_MM = 45;

const INDIVIDUAL_STICKER_WIDTH_PTS = INDIVIDUAL_STICKER_WIDTH_MM * 2.83465;  // ~226.772 pts
const INDIVIDUAL_STICKER_HEIGHT_PTS = INDIVIDUAL_STICKER_HEIGHT_MM * 2.83465; // ~127.559 pts

// User requested layout: 2 columns x 6 rows for 12 stickers on A4
const STICKER_COLS = 2;
const STICKER_ROWS = 6;
const TOTAL_STICKERS_PER_PAGE = STICKER_COLS * STICKER_ROWS;

// Calculated Margins and Gutters for even distribution on A4
const LEFT_MARGIN_PTS = 47.245;
const TOP_MARGIN_PTS = 10.91;

const GUTTER_HORIZONTAL_PTS = 47.245;
const GUTTER_VERTICAL_PTS = 10.91;

// Crop Mark constants
const CROP_MARK_LENGTH_PTS = 15;
const CROP_MARK_OFFSET_PTS = 5;

// --- Google Drive Folder ID for Fiery Hot Folder Integration ---
// IMPORTANT: Replace with the actual Google Drive Folder ID that your local sync script monitors.
const FIERY_PRINT_QUEUE_FOLDER_ID = '1PfG_zgaMV1k44aYj4RLX7XuVD1u69GQP'; // e.g., '1PfG_zgaMV1k44aYj4RLX7XuVD1u69GQP'

// --- Template ID for individual sticker design if needed (from your existing label generator) ---
const INDIVIDUAL_LABEL_TEMPLATE_ID = '1OxCwLQhFWbZc5UzcSKSC1YYfe0NfAVsyBBX3bl2pJKM'; // Your existing label template ID

// Logo URL (from your existing label generator)
const LOGO_IMAGE_URL = "https://drive.google.com/uc?export=view&id=1n8_F8SSx5HIGJs56o-kWo6xHcb2m5NWn";

// NEW CONSTANT: ID of the blank A4-sized Google Slides template
// YOU MUST REPLACE THIS WITH THE ID OF THE A4 TEMPLATE YOU CREATED IN STEP 1
const A4_BLANK_SLIDES_TEMPLATE_ID = '1tbT1G2Q63NDOU1yTVsOAu8m7zGleeVhYzjUNWwNumFM'; // <--- IMPORTANT: REPLACE THIS!

// --- End Constants ---


/**
 * Helper function to retrieve Item Name, Length, and Width from the "Product info" sheet.
 * @param {string} itemId The Item ID to look up.
 * @returns {Object|null} An object {itemName, length, width} or null if not found.
 */
function getProductItemDetails(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productInfoSheet = ss.getSheetByName("Product info");
  if (!productInfoSheet) {
    Logger.log("Error: 'Product info' sheet not found in getProductItemDetails.");
    return null;
  }

  const lastRow = productInfoSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No product data found in 'Product info' sheet.");
    return null;
  }

  const headers = productInfoSheet.getRange(1, 1, 1, productInfoSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const itemIdCol = headers.indexOf("Item ID");
  const lengthCol = headers.indexOf("Length");
  const widthCol = headers.indexOf("Width");
  const itemNameCol = headers.indexOf("Item Name");

  if (itemIdCol === -1 || lengthCol === -1 || widthCol === -1 || itemNameCol === -1) {
    Logger.log("Warning: Missing required columns in 'Product info' sheet for getProductItemDetails.");
    return null;
  }

  const data = productInfoSheet.getRange(2, 1, lastRow - 1, productInfoSheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[itemIdCol] && row[itemIdCol].toString().trim() === itemId.toString().trim()) {
      return {
        itemName: row[itemNameCol] ? row[itemNameCol].toString().trim() : '',
        length: row[lengthCol] ? row[lengthCol].toString().trim() : '',
        width: row[widthCol] ? row[widthCol].toString().trim() : ''
      };
    }
  }
  Logger.log(`Item ID '${itemId}' not found in 'Product info' sheet.`);
  return null;
}


function generateLabelViaSlides(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock Tracker");
  const [itemId, itemName, , width, length] = sheet.getRange(row, 1, 1, 6).getValues()[0];
  const labelUrlCell = sheet.getRange(row, 7); // Column G
  const statusCell = sheet.getRange(row, 8);   // Column H
  statusCell.setValue("⏳ Generating...");
  try {
    if (!itemId || !itemName) {
      statusCell.setValue("❌ Missing Item ID or Name");
      return;
    }

    const templateId = '1OxCwLQhFWbZc5UzcSKSC1YYfe0NfAVsyBBX3bl2pJKM';
    const formURL = `https://docs.google.com/forms/d/e/1FAIpQLScNhuWKlDj3Scb3I0L0otsYljcATJxKRZvr_dF7MJenvVWjOg/viewform?usp=pp_url&entry.118779491=${encodeURIComponent(itemId)}`;
    const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(formURL)}`;
    const logoUrl = "https://drive.google.com/uc?export=view&id=1n8_F8SSx5HIGJs56o-kWo6xHcb2m5NWn";
    const folderId = '1rZ4-9W5ogddHdsGV4wp1jQQtVSU3-x8D';
    const folder = DriveApp.getFolderById(folderId);

    const presentation = DriveApp.getFileById(templateId).makeCopy(`${itemId}`);
    const slideDeck = SlidesApp.openById(presentation.getId());
    const slide = slideDeck.getSlides()[0];
    slide.getPageElements().forEach(el => el.remove());

    // Insert QR
    slide.insertImage(qrImageUrl).setLeft(13.68).setTop(13.68).setWidth(87.84).setHeight(87.84);
    // Insert Logo
    slide.insertImage(logoUrl).setLeft(148.32).setTop(5.76).setWidth(30.24).setHeight(29.52);

    // Item Name
    slide.insertTextBox(`${itemName}`, 103.68, 30.96, 123.12, 25.20)
         .getText().getTextStyle()
         .setFontFamily("Lexend")
         .setFontSize(9)
         .setBold(true)
         .setForegroundColor("#000000");
    // Length
    slide.insertTextBox(`Length: ${length}`, 103.68, 55.36, 116.64, 25.2)
         .getText().getTextStyle()
         .setFontFamily("Lexend")
         .setFontSize(9)
         .setBold(false)
         .setForegroundColor("#000000");
    // Width
    slide.insertTextBox(`Width: ${width}`, 103.68, 66.88, 116.64, 25.2)
         .getText().getTextStyle()
         .setFontFamily("Lexend")
         .setFontSize(9)
         .setBold(false)
         .setForegroundColor("#000000");
    // Item ID at bottom
    const idBox = slide.insertTextBox(`${itemId}`, 0, 101.52, 226.8, 25.92);
    idBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    const idText = idBox.getText();
    idText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    idText.getTextStyle()
          .setFontFamily("Lexend")
          .setFontSize(12)
          .setBold(true)
          .setForegroundColor("#000000");
    slideDeck.saveAndClose();

    // Generate PDF
    const pdfBlob = DriveApp.getFileById(presentation.getId()).getAs('application/pdf');
    const fileName = `${itemId}.pdf`;
    // Remove existing
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    // Save new file
    const pdf = folder.createFile(pdfBlob).setName(fileName);
    labelUrlCell.setFormula(`=HYPERLINK("${pdf.getUrl()}", "Label PDF")`);
    statusCell.setValue("✅ Label Created");
    // Clean up slide file
    DriveApp.getFileById(presentation.getId()).setTrashed(true);

  } catch (err) {
    statusCell.setValue(`❌ Error: ${err.message}`);
    Logger.log(err.stack);
  }
}

function generateMissingLabels() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stock Tracker");
  const data = sheet.getDataRange().getValues();
  const headerRow = 1;

  let generatedCount = 0;
  for (let row = 2; row <= data.length; row++) {
    const itemId = data[row - 1][0];
    const labelUrl = data[row - 1][6];

    if (itemId && !labelUrl) {
      Logger.log(`Generating label for row ${row} → ${itemId}`);
      generateLabelViaSlides(row);
      generatedCount++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Finished: ${generatedCount} label(s) generated.`);
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("More")
    .addItem("Generate Label for Selected Row", "generateLabelFromSelection_Slides")
    .addItem("Generate Missing Labels", "generateMissingLabels")
    .addItem("📊 Run Full Stock Check & Reconciliation", "runStockChecks")
    .addItem("🔴 Send Urgent Stock Alert Email", "triggerUrgentStockEmail")
    .addItem("📬 Send Daily Stock Summary Email", "triggerDailyStockEmail")
    .addSeparator()
    .addItem("🖥️ Open Inventory Portal", "openInventoryPortal")
    .addToUi();
}

function openInventoryPortal() {
  const url = "https://script.google.com/macros/s/AKfycbziqITrCbW3YhZOzK4t6TPgtC0GspVp3fI6gMssCac/dev";
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<a href="${url}" target="_blank">Click here to open the Inventory Portal</a>`)
      .setWidth(300)
      .setHeight(100),
    "Inventory Portal"
  );
}


function generateLabelFromSelection_Slides() {
  const row = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please select a data row.");
    return;
  }
  generateLabelViaSlides(row);
}

function generateStickerPrintJob(itemsToPrint) {
  Logger.log("generateStickerPrintJob: Function started.");
  if (!itemsToPrint || itemsToPrint.length === 0) { // Corrected: itemsToPrint instead of itemsToToPrint
    Logger.log("Error: No items provided for sticker print job.");
    return { success: false, message: "No items to print." };
  }

  // Get the Google Drive folder for Fiery integration (or where you want to save the CSV)
  let outputFolder;
  try {
    // Using FIERY_PRINT_QUEUE_FOLDER_ID as the destination folder, as per previous context.
    outputFolder = DriveApp.getFolderById(FIERY_PRINT_QUEUE_FOLDER_ID);
  } catch (e) {
    Logger.log(`Error: Output Folder ID '${FIERY_PRINT_QUEUE_FOLDER_ID}' not found or inaccessible. Error: ${e.message}`);
    return { success: false, message: `Output folder not found or inaccessible. Please check the ID in the script config.` };
  }

  let fileContent = '';
  let fileName = `Sticker_Print_Job_${new Date().toISOString().replace(/[:.]/g, '-')}.csv`;

  // Build CSV header - matching the provided example file
  fileContent += `"Item ID","Item Name","Length","Width","Quantity"\n`;

  // Build CSV rows from the itemsToPrint array
  itemsToPrint.forEach(item => {
    // Get product details for Length and Width (these are still needed from Product info sheet)
    const productDetails = getProductItemDetails(item.itemId); // Assuming getProductItemDetails is accessible
    const length = productDetails ? String(productDetails.length).replace(/"/g, '""') : '';
    const width = productDetails ? String(productDetails.width).replace(/"/g, '""') : '';

    // Ensure values are properly quoted and escaped for CSV
    const itemId = item.itemId ? `"${String(item.itemId).replace(/"/g, '""')}"` : '""';
    const itemName = item.itemName ? `"${String(item.itemName).replace(/"/g, '""')}"` : '""';
    const qty = (item.qty !== undefined && item.qty !== null) ? String(item.qty) : ''; // Convert to string
    
    // Add Length and Width to the row, matching the order in the header
    fileContent += `${itemId},${itemName},"${length}","${width}",${qty}\n`; 
  });

  try {
    const csvFile = outputFolder.createFile(fileName, fileContent, MimeType.CSV);
    Logger.log(`✅ CSV file saved to Drive: ${csvFile.getUrl()}`);

    return { success: true, message: `CSV file '${fileName}' saved to Google Drive.`, fileUrl: csvFile.getUrl() };
  } catch (e) {
    Logger.log(`❌ Error saving CSV file to Drive: ${e.message}. Stack: ${e.stack}`);
    return { success: false, message: `Failed to save CSV file: ${e.message}` };
  }
}