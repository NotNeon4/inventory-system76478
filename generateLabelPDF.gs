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
// ... (existing constants and helper functions remain the same) ...

/**
 * Generates a printable PDF with multiple stickers (2x6 layout for 12-up on A4)
 * including crop marks, and saves it to a Google Drive folder for Fiery integration.
 *
 * @param {Array<Object>} itemsToPrint - An array of objects, each with { itemId: string, itemName: string, qty: number }.
 * @returns {Object} An object with success status and the generated file name/URL.
 */
function generateStickerPrintJob(itemsToPrint) {
  Logger.log("generateStickerPrintJob: Function started.");
  if (!itemsToPrint || itemsToPrint.length === 0) {
    Logger.log("Error: No items provided for sticker print job.");
    return { success: false, message: "No items to print." };
  }

  // Get the Google Drive folder for Fiery integration
  let fieryPrintFolder;
  try {
    fieryPrintFolder = DriveApp.getFolderById(FIERY_PRINT_QUEUE_FOLDER_ID);
  } catch (e) {
    Logger.log(`Error: Fiery Print Queue Folder ID '${FIERY_PRINT_QUEUE_FOLDER_ID}' not found or inaccessible. Error: ${e.message}`);
    return { success: false, message: `Fiery print folder not found or inaccessible. Please check the ID in the script config.` };
  }

  let tempPresentationId = null; // Will store the ID of the new presentation
  let tempDriveFileObject = null;     // This will hold the DriveApp.File object for cleanup
  let fileBlob = null;
  let fileName = '';

  try {
    // --- ADVANCED SERVICE: Create Presentation ---
    // Use the Slides API directly to create the presentation
    Logger.log("Attempting to create presentation via Slides.Presentations.create");
    const newPresentationBody = {
      title: `Sticker Print Job - ${new Date().toLocaleString()}`
    };
    const newPresentation = Slides.Presentations.create(newPresentationBody);
    tempPresentationId = newPresentation.presentationId;
    Logger.log("Debug: Presentation ID created via Advanced Slides API: " + tempPresentationId);

    if (!tempPresentationId) {
      throw new Error("Failed to create presentation via Advanced Slides API. No ID returned.");
    }

    // Get the DriveApp.File object for this presentation for later PDF conversion and trashing
    tempDriveFileObject = DriveApp.getFileById(tempPresentationId);
    Logger.log("Debug: Drive File object for presentation: " + tempDriveFileObject);
    
    // --- ADVANCED SERVICE: Set Page Size (using updatePresentationProperties) ---
    // This request sets the page size for the whole presentation.
    // The previous error "Unknown name "updatePresentationProperties"" is highly unusual if field name is correct.
    // Let's re-confirm the exact structure based on API docs.
    const setPageSizeRequest = {
      "updatePresentationProperties": { // This should be the correct field name (camelCase)
        "presentationProperties": {
          "pageSize": {
            "width": { "magnitude": A4_WIDTH_PTS, "unit": "PT" },
            "height": { "magnitude": A4_HEIGHT_PTS, "unit": "PT" }
          }
        },
        "fields": "pageSize"
      }
    };

    Logger.log("Debug: setPageSizeRequest payload: " + JSON.stringify(setPageSizeRequest));
    Slides.Presentations.batchUpdate({ requests: [setPageSizeRequest] }, tempPresentationId);
    Logger.log(`Presentation page size updated to A4: ${A4_WIDTH_PTS}x${A4_HEIGHT_PTS} pts`);

    // --- Now, open the presentation using SlidesApp for content manipulation ---
    // We create it with the API, set size with API, then open it with SlidesApp for easier scripting
    const slidePresentationObject = SlidesApp.openById(tempPresentationId);
    
    // Clear the default first slide's content.
    // The Slides API creates one default slide. We will clear its content here.
    if (slidePresentationObject.getSlides().length > 0) {
      const defaultSlide = slidePresentationObject.getSlides()[0];
      defaultSlide.getPageElements().forEach(el => el.remove()); // Remove all elements
      Logger.log("Default slide content cleared.");
    }

    let currentPage = null;
    let stickerCountOnCurrentPage = 0;
    let totalPages = 0;

    itemsToPrint.forEach(item => {
      const itemId = item.itemId;
      const qtyToPrint = item.qty;
      
      const productDetails = getProductItemDetails(itemId); // Fetch detailed info
      const itemName = productDetails ? productDetails.itemName : (item.itemName || 'Unknown Item');
      const itemLength = productDetails ? productDetails.length : 'N/A';
      const itemWidth = productDetails ? productDetails.width : 'N/A';

      Logger.log(`Processing item: ${itemId}, Quantity: ${qtyToPrint}, Details: ${JSON.stringify(productDetails)}`);

      for (let i = 0; i < qtyToPrint; i++) {
        // Check if a new page is needed
        if (stickerCountOnCurrentPage === 0 || stickerCountOnCurrentPage >= TOTAL_STICKERS_PER_PAGE) {
          currentPage = slidePresentationObject.appendSlide(SlidesApp.PredefinedLayout.BLANK);
          stickerCountOnCurrentPage = 0;
          totalPages++;
          Logger.log(`Starting new page ${totalPages}`);

          // Add A4 crop marks to the new page (for the entire A4 sheet)
          addA4PageCropMarks(currentPage);
        }

        const col = stickerCountOnCurrentPage % STICKER_COLS;
        const row = Math.floor(stickerCountOnCurrentPage / STICKER_COLS);

        const xPos = LEFT_MARGIN_PTS + (col * (INDIVIDUAL_STICKER_WIDTH_PTS + GUTTER_HORIZONTAL_PTS));
        const yPos = TOP_MARGIN_PTS + (row * (INDIVIDUAL_STICKER_HEIGHT_PTS + GUTTER_VERTICAL_PTS));

        Logger.log(`Placing sticker ${i + 1}/${qtyToPrint} for ${itemId} at (x: ${xPos}, y: ${yPos}) on page ${totalPages}`);

        // --- Render individual sticker content (QR, Text, Logo) ---
        const formURL = `https://docs.google.com/forms/d/e/1FAIpQLScNhuWKlDj3Scb3I0L0otsYljcATJxKRZvr_dF7MJenvVWjOg/viewform?usp=pp_url&entry.118779491=${encodeURIComponent(itemId)}`;
        const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(formURL)}`;
        
        const qrSize = INDIVIDUAL_STICKER_WIDTH_PTS * 0.4;
        currentPage.insertImage(qrImageUrl)
          .setLeft(xPos + (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05))
          .setTop(yPos + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.05))
          .setWidth(qrSize)
          .setHeight(qrSize);

        const logoWidth = INDIVIDUAL_STICKER_WIDTH_PTS * 0.2;
        const logoHeight = INDIVIDUAL_STICKER_HEIGHT_PTS * 0.2;
        currentPage.insertImage(LOGO_IMAGE_URL)
          .setLeft(xPos + INDIVIDUAL_STICKER_WIDTH_PTS - logoWidth - (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05))
          .setTop(yPos + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.05))
          .setWidth(logoWidth)
          .setHeight(logoHeight);

        currentPage.insertTextBox(itemName, 
          xPos + (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05), 
          yPos + (qrSize) + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.08),
          INDIVIDUAL_STICKER_WIDTH_PTS * 0.9,
          20
        )
        .getText().getTextStyle().setFontSize(10).setBold(true);

        currentPage.insertTextBox(`Item ID: ${itemId}`, 
          xPos + (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05),
          yPos + (qrSize) + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.08) + 20,
          INDIVIDUAL_STICKER_WIDTH_PTS * 0.9,
          20
        )
        .getText().getTextStyle().setFontSize(9);
        
        currentPage.insertTextBox(`Length: ${itemLength}`, 
          xPos + (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05),
          yPos + (qrSize) + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.08) + 40,
          INDIVIDUAL_STICKER_WIDTH_PTS * 0.9,
          20
        )
        .getText().getTextStyle().setFontSize(9);

        currentPage.insertTextBox(`Width: ${itemWidth}`, 
          xPos + (INDIVIDUAL_STICKER_WIDTH_PTS * 0.05),
          yPos + (qrSize) + (INDIVIDUAL_STICKER_HEIGHT_PTS * 0.08) + 60,
          INDIVIDUAL_STICKER_WIDTH_PTS * 0.9,
          20
        )
        .getText().getTextStyle().setFontSize(9);

        stickerCountOnCurrentPage++;
      }
    });

    slidePresentationObject.saveAndClose();

    // Generate PDF name
    fileName = `Sticker_Print_Job_${new Date().toISOString().replace(/[:.]/g, '-')}.pdf`;
    
    // Convert the presentation to PDF Blob using the Drive File object
    fileBlob = tempDriveFileObject.getAs(MimeType.PDF);
    
    // Save the PDF to the designated folder
    const pdfFile = fieryPrintFolder.createFile(fileBlob).setName(fileName);
    Logger.log(`Generated PDF saved to Drive: ${pdfFile.getUrl()}`);

    return { success: true, fileName: fileName, fileUrl: pdfFile.getUrl() };

  } catch (e) {
    Logger.log(`❌ Error generating sticker print job: ${e.message}. Stack: ${e.stack}`);
    return { success: false, message: e.message };
  } finally {
    // Clean up the temporary Google Slides presentation using the Drive File object
    if (tempDriveFileObject) {
      try {
        tempDriveFileObject.setTrashed(true);
        Logger.log("Temporary presentation trashed.");
      } catch (e) {
        Logger.log("Error trashing temporary presentation: " + e.message);
      }
    }
  }
}
/**
 * Adds crop marks to the corners of the A4 slide to indicate the full page dimensions.
 * These are for the overall A4 sheet, not individual stickers.
 * @param {GoogleAppsScript.Slides.Slide} slide The slide to add crop marks to.
 */
function addA4PageCropMarks(slide) {
  const width = slide.getPageWidth();
  const height = slide.getPageHeight();

  // Top-left corner
  slide.insertLine(CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, CROP_MARK_LENGTH_PTS + CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS); // Horizontal
  slide.insertLine(CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, CROP_MARK_LENGTH_PTS + CROP_MARK_OFFSET_PTS); // Vertical

  // Top-right corner
  slide.insertLine(width - CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, width - CROP_MARK_LENGTH_PTS - CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS); // Horizontal
  slide.insertLine(width - CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, width - CROP_MARK_OFFSET_PTS, CROP_MARK_LENGTH_PTS + CROP_MARK_OFFSET_PTS); // Vertical

  // Bottom-left corner
  slide.insertLine(CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS, CROP_MARK_LENGTH_PTS + CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS); // Horizontal
  slide.insertLine(CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS, CROP_MARK_OFFSET_PTS, height - CROP_MARK_LENGTH_PTS - CROP_MARK_OFFSET_PTS); // Vertical

  // Bottom-right corner
  slide.insertLine(width - CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS, width - CROP_MARK_LENGTH_PTS - CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS); // Horizontal
  slide.insertLine(width - CROP_MARK_OFFSET_PTS, height - CROP_MARK_OFFSET_PTS, width - CROP_MARK_OFFSET_PTS, height - CROP_MARK_LENGTH_PTS - CROP_MARK_OFFSET_PTS); // Vertical

  Logger.log("Added A4 page crop marks.");

  // Add internal cut marks (between stickers)
  // Horizontal cut lines
  for (let r = 1; r < STICKER_ROWS; r++) {
    const yCutPos = TOP_MARGIN_PTS + (r * INDIVIDUAL_STICKER_HEIGHT_PTS) + ((r - 1) * GUTTER_VERTICAL_PTS) + (GUTTER_VERTICAL_PTS / 2);
    slide.insertLine(LEFT_MARGIN_PTS, yCutPos, A4_WIDTH_PTS - LEFT_MARGIN_PTS, yCutPos);
  }

  // Vertical cut lines
  for (let c = 1; c < STICKER_COLS; c++) {
    const xCutPos = LEFT_MARGIN_PTS + (c * INDIVIDUAL_STICKER_WIDTH_PTS) + ((c - 1) * GUTTER_HORIZONTAL_PTS) + (GUTTER_HORIZONTAL_PTS / 2);
    slide.insertLine(xCutPos, TOP_MARGIN_PTS, xCutPos, A4_HEIGHT_PTS - TOP_MARGIN_PTS);
  }
  Logger.log("Added internal cut marks between stickers.");
}