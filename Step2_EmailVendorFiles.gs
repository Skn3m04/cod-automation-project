function sendCODReportEmails() {
  const rawDataFileId = '1_wLdRUwEiC1E1jcVf1eXrs5ZFqIYLweEyk26EceVbeM';
  const masterFileId = '1ljNim13--HZSdTs7-nRLjFKKK4O65P1aHFvg_UaZIUA';

  const rawDataSS = SpreadsheetApp.openById(rawDataFileId);
  const masterSS = SpreadsheetApp.openById(masterFileId);

  const summarySheet = masterSS.getSheetByName("Master Summary Sheet");
  const rawSheet = rawDataSS.getSheetByName("Raw Data Sheet"); // Update if needed
  const bankSheet = rawDataSS.getSheetByName("Received in Pidge Bank");
  const creditSheet = rawDataSS.getSheetByName("Credit Notes");
  const vendorMapSheet = rawDataSS.getSheetByName("VendorMap");

  if (!summarySheet || !rawSheet || !bankSheet || !creditSheet || !vendorMapSheet) {
    throw new Error("One or more sheets not found. Please check all sheet names.");
  }

  const summaryData = summarySheet.getDataRange().getValues();
  const vendorMap = vendorMapSheet.getDataRange().getValues();
  const rawData = rawSheet.getDataRange().getValues();
  const bankData = bankSheet.getDataRange().getValues();
  const creditData = creditSheet.getDataRange().getValues();

  // ✅ Use YESTERDAY'S DATE
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const today = Utilities.formatDate(yesterday, "Asia/Kolkata", "dd MMMM yyyy");

  const findVendorIDIndex = (headerRow) => {
    const keyTerms = ["vendor", "partner", "seller"];
    for (let i = 0; i < headerRow.length; i++) {
      const cell = headerRow[i].toString().toLowerCase();
      if (keyTerms.some(term => cell.includes(term))) return i;
    }
    throw new Error("Vendor ID column not found.");
  };

  const rawVendorCol = findVendorIDIndex(rawData[0]);
  const bankVendorCol = findVendorIDIndex(bankData[0]);
  const creditVendorCol = findVendorIDIndex(creditData[0]);

  for (let i = 1; i < vendorMap.length; i++) {
    const vendorID = vendorMap[i][0];
    const vendorName = vendorMap[i][1];
    const emailRaw = vendorMap[i][2];
    const ccRaw = vendorMap[i][3];

    if (!emailRaw) continue;

    const emailList = emailRaw.toString().split(",").map(e => e.trim()).filter(e => e.length > 0);
    const ccList = ccRaw ? ccRaw.toString().split(",").map(e => e.trim()).filter(e => e.length > 0) : [];

    if (emailList.length === 0) continue;

    const summaryHeader = summaryData[0];
    const vendorSummaryRow = summaryData.find(row => row[0] == vendorID);

    const filterByVendor = (data, vendorColIndex) => {
      const headers = data[0];
      const rows = data.slice(1).filter(r => r[vendorColIndex] === vendorID);
      return rows.length ? [headers, ...rows] : [];
    };

    const rawFiltered = filterByVendor(rawData, rawVendorCol);
    const bankFiltered = filterByVendor(bankData, bankVendorCol);
    const creditFiltered = filterByVendor(creditData, creditVendorCol);

    // ✅ Skip empty vendors
    const hasNoData =
      !vendorSummaryRow &&
      rawFiltered.length === 0 &&
      bankFiltered.length === 0 &&
      creditFiltered.length === 0;

    if (hasNoData) {
      Logger.log(`Skipped vendor "${vendorName}" (ID: ${vendorID}) - No data found.`);
      continue;
    }

    Logger.log(`VendorID: ${vendorID}, Name: ${vendorName}, RawRows: ${rawFiltered.length}, BankRows: ${bankFiltered.length}, CreditRows: ${creditFiltered.length}`);

    const tempSpreadsheet = SpreadsheetApp.create(`COD Report - ${vendorName}`);

    if (vendorSummaryRow) {
      const summarySheet = tempSpreadsheet.insertSheet("COD Summary");
      summarySheet.getRange(1, 1, 1, summaryHeader.length).setValues([summaryHeader]);
      summarySheet.getRange(2, 1, 1, vendorSummaryRow.length).setValues([vendorSummaryRow]);
    }

    if (rawFiltered.length) {
      const s = tempSpreadsheet.insertSheet("Raw Data Sheet");
      s.getRange(1, 1, rawFiltered.length, rawFiltered[0].length).setValues(rawFiltered);
    }

    if (bankFiltered.length) {
      const s = tempSpreadsheet.insertSheet("Received in Pidge Bank");
      s.getRange(1, 1, bankFiltered.length, bankFiltered[0].length).setValues(bankFiltered);
    }

    if (creditFiltered.length) {
      const s = tempSpreadsheet.insertSheet("Credit Notes");
      s.getRange(1, 1, creditFiltered.length, creditFiltered[0].length).setValues(creditFiltered);
    }

    const defaultSheet = tempSpreadsheet.getSheetByName("Sheet1");
    if (defaultSheet && tempSpreadsheet.getSheets().length > 1) {
      tempSpreadsheet.deleteSheet(defaultSheet);
    }

    const exportUrl = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${tempSpreadsheet.getId()}&exportFormat=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });

    const blob = response.getBlob().setName(`${vendorName}_COD_Report_${today}.xlsx`);

    let htmlBody = `<p>Dear ${vendorName},</p>`;
    htmlBody += `<p>Please find below your COD Summary as of <strong>${today}</strong>:</p>`;

    if (vendorSummaryRow) {
      htmlBody += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;">`;
      htmlBody += `<tr>${summaryHeader.map(h => `<th>${h}</th>`).join('')}</tr>`;
      htmlBody += `<tr>${vendorSummaryRow.map(cell => `<td>${cell}</td>`).join('')}</tr>`;
      htmlBody += `</table><br>`;
    } else {
      htmlBody += `<p>No summary data found.</p>`;
    }

    htmlBody += `<p>Attached is your detailed COD report.<br>Regards,<br>Shahbaz Khan</p>`;

    try {
      GmailApp.sendEmail(emailList.join(","), `COD Report - Data till ${today}`, "Please find your COD report attached.", {
        htmlBody: htmlBody,
        attachments: [blob],
        cc: ccList.join(",")
      });
      Logger.log(`✅ Email sent to: ${emailList.join(", ")} for vendor: ${vendorName}`);
    } catch (emailErr) {
      Logger.log(`❌ Failed to email vendor ${vendorName}: ${emailErr.message}`);
    }

    try {
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    } catch (err) {
      Logger.log(`⚠️ Failed to delete spreadsheet for vendor ${vendorName}: ${err.message}`);
    }
  }
}
