function generateVendorIDs() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("CRM");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const vendorId = data[i][0];   // Column A
    const vendorName = data[i][1]; // Column B

    if (vendorName && !vendorId) {
      const newId = "v_" + Utilities.getUuid().slice(0, 8);
      sheet.getRange(i + 1, 1).setValue(newId);
    }
  }
}