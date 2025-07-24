/**
 * processCODDataStep2.gs
 * Author: K (Logic Designer) | Script written with AI assistance
 * 
 * Description:
 * End-to-end Google Apps Script automation for:
 * - Processing COD data from multiple sheets
 * - Mapping Rider & Vendor names
 * - Assigning COD types
 * - Highlighting missing mappings
 * - Generating and emailing Excel file
 * - Cleaning up temp files
 */

function processCODDataStep2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName("Raw Data");
    const vendorSheet = ss.getSheetByName("Rider's Vendor Data");
    const codSheet = ss.getSheetByName("COD Customers Data");

    if (!rawSheet || !vendorSheet || !codSheet) {
      console.log("❌ One or more sheets missing!");
      return;
    }

    const processedSheetName = "Processed Data";
    let processedSheet = ss.getSheetByName(processedSheetName);
    if (processedSheet) ss.deleteSheet(processedSheet);
    processedSheet = ss.insertSheet(processedSheetName);

    const rawData = rawSheet.getDataRange().getValues();
    const headers = rawData[0];
    const dataRows = rawData.slice(1);

    // Create mapping for Rider ID → [Rider Name, Vendor Name]
    const riderMap = new Map();
    const vendorData = vendorSheet.getDataRange().getValues().slice(1);
    vendorData.forEach(row => {
      riderMap.set(row[0], [row[2], row[8]]);
    });

    // Create mapping for brand name → COD type
    const codMap = new Map();
    const codData = codSheet.getDataRange().getValues().slice(1);
    codData.forEach(row => {
      codMap.set(row[0].toString().trim().toLowerCase(), row[1]);
    });

    const riderIdIndex = headers.indexOf("rider_id");
    const brandNameIndex = headers.indexOf("brand_name");
    const transactionTypeIndex = headers.indexOf("transaction_type");

    const newHeaders = [...headers];
    newHeaders.splice(riderIdIndex + 1, 0, "Rider Name", "Vendor Name");
    newHeaders.push("COD_type");

    const outputRows = [newHeaders];
    const backgroundColors = [new Array(newHeaders.length).fill(null)];
    let logCount = 0;

    // Loop through data rows
    dataRows.forEach(row => {
      const riderId = row[riderIdIndex];
      const brandNameRaw = row[brandNameIndex];
      const transactionType = row[transactionTypeIndex];

      let riderName = "To be checked Manually";
      let vendorName = "To be checked Manually";

      if (riderMap.has(riderId)) {
        [riderName, vendorName] = riderMap.get(riderId);
      }

      let cleanedBrand = brandNameRaw ? brandNameRaw.toString().trim().toLowerCase() : "";
      let codType = "";

      if (cleanedBrand && codMap.has(cleanedBrand)) {
        codType = codMap.get(cleanedBrand);
      } else if (!brandNameRaw || brandNameRaw === "") {
        if (transactionType == 300) {
          row[brandNameIndex] = "Rider Paid Online";
          codType = "Yes";
        } else if (transactionType == 100) {
          codType = "To be removed";
        }
      }

      const newRow = [...row];
      newRow.splice(riderIdIndex + 1, 0, riderName, vendorName);
      newRow.push(codType);
      outputRows.push(newRow);

      const bgRow = new Array(newHeaders.length).fill(null);
      if (riderName === "To be checked Manually" || vendorName === "To be checked Manually") {
        bgRow[riderIdIndex + 1] = "#FFF3CD";
        bgRow[riderIdIndex + 2] = "#FFF3CD";
      }
      backgroundColors.push(bgRow);
      logCount++;
    });

    processedSheet.getRange(1, 1, outputRows.length, newHeaders.length).setValues(outputRows);
    processedSheet.getRange(1, 1, backgroundColors.length, newHeaders.length).setBackgrounds(backgroundColors);

    console.log(`✅ Processed ${logCount} rows.`);

    // Generate Excel file
    const exportFileName = `COD Processed Data – ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy')}`;
    const tempSpreadsheet = SpreadsheetApp.create(exportFileName);
    const tempSheet = tempSpreadsheet.getActiveSheet();
    tempSheet.getRange(1, 1, outputRows.length, newHeaders.length).setValues(outputRows);
    tempSheet.setName("Processed Data");

    // Export to XLSX using Drive API
    const tempFileId = tempSpreadsheet.getId();
    const exportUrl = `https://www.googleapis.com/drive/v3/files/${tempFileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const blob = response.getBlob().setName(`${exportFileName}.xlsx`);

    // Send Email
    GmailApp.sendEmail("shahbaz.khan@r-indventures.com", exportFileName, "Please find attached the latest COD processed data.", {
      attachments: [blob],
      name: "COD Automation Bot"
    });

    // Clean up temp file
    DriveApp.getFileById(tempFileId).setTrashed(true);

    // Delete temporary sheet from active spreadsheet
    const finalProcessedSheet = ss.getSheetByName(processedSheetName);
    if (finalProcessedSheet) {
      ss.deleteSheet(finalProcessedSheet);
      console.log("🗑️ Deleted 'Processed Data' sheet after sending email.");
    }

    console.log("📧 Email sent successfully with proper Excel file!");

  } catch (e) {
    console.error("❌ Error:", e.toString());
  }
}
