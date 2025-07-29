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
    const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=600x600&data=${encodeURIComponent(formURL)}`;
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
  const data = sheet.getDataRange().getValues(); // all data
  const headerRow = 1;

  let generatedCount = 0;

  for (let row = 2; row <= data.length; row++) {
    const itemId = data[row - 1][0]; // Column A
    const labelUrl = data[row - 1][6]; // Column G

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
    .addItem("🔴 Check Urgent Stock (Send Email)", "runUrgentStockCheck")
    .addItem("📬 Send Daily Stock Summary", "sendDailyStockEmail")
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
