/**
 * Moves invoice PDF files from a temporary Drive folder to destination folders
 * defined in a Google Sheet, and writes the Drive link back to the sheet.
 *
 * Uses Google Drive Advanced API (supports Shared Drives).
 */
function moveInvoiceFilesFromTempFolder() {
  // ========= CONFIGURATION =========
  const SHEET_NAME = "Invoice_List";          // Sheet containing invoice data
  const TEMP_FOLDER_ID = "TEMP_FOLDER_ID";    // Temporary Drive folder ID

  const COL_INVOICE_NUMBER = 1; // Column A
  const COL_DEST_FOLDER_ID = 13; // Column M
  const COL_FILE_LINK = 16; // Column P
  // =================================

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  // ===== INDEX FILES IN TEMP FOLDER =====
  const filesByName = {};
  let pageToken = null;

  do {
    const response = Drive.Files.list({
      q: `'${TEMP_FOLDER_ID}' in parents and trashed = false`,
      fields: "nextPageToken, files(id, name, parents)",
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      pageToken: pageToken
    });

    (response.files || []).forEach(file => {
      filesByName[file.name] = file;
    });

    pageToken = response.nextPageToken;
  } while (pageToken);

  // ===== PROCESS SHEET ROWS =====
  let stats = {
    moved: 0,
    skippedNoPdf: 0,
    skippedNoFolder: 0,
    skippedHasLink: 0,
    errors: 0
  };

  data.forEach((row, index) => {
    const rowNumber = index + 2;

    const invoiceNumber = row[COL_INVOICE_NUMBER - 1];
    const destinationFolderId = row[COL_DEST_FOLDER_ID - 1];
    const existingLink = row[COL_FILE_LINK - 1];

    // Skip if already processed
    if (existingLink && existingLink.toString().trim() !== "") {
      stats.skippedHasLink++;
      return;
    }

    if (!invoiceNumber) return;

    const expectedFileName = `Invoice_${invoiceNumber}.pdf`;
    const file = filesByName[expectedFileName];

    if (!file) {
      stats.skippedNoPdf++;
      return;
    }

    if (!destinationFolderId || destinationFolderId.toString().trim() === "") {
      stats.skippedNoFolder++;
      return;
    }

    try {
      // Move file to destination folder
      Drive.Files.update(
        {},
        file.id,
        null,
        {
          addParents: destinationFolderId.toString().trim(),
          removeParents: file.parents.join(","),
          supportsAllDrives: true
        }
      );

      // Write Drive link back to sheet
      const fileLink = `https://drive.google.com/file/d/${file.id}/view`;
      sheet.getRange(rowNumber, COL_FILE_LINK).setValue(fileLink);

      stats.moved++;

    } catch (error) {
      Logger.log(`Row ${rowNumber} | Error: ${error.message}`);
      stats.errors++;
    }
  });

  // ===== SUMMARY LOG =====
  Logger.log("===== PROCESS SUMMARY =====");
  Object.keys(stats).forEach(key => {
    Logger.log(`${key}: ${stats[key]}`);
  });
}
