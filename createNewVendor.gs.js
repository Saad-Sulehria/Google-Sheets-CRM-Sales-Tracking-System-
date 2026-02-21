function createNewVendor() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const crmSheet = ss.getSheetByName("CRM");
  const notesSheet = ss.getSheetByName("Meeting Notes");
  const templateSheet = ss.getSheetByName("Template_LeftBlock");

  if (!crmSheet || !notesSheet || !templateSheet) {
    SpreadsheetApp.getUi().alert("Missing sheet: CRM, Meeting Notes, or Template_LeftBlock");
    return;
  }

  const row = crmSheet.getActiveCell().getRow();

  if (row === 1) {
    SpreadsheetApp.getUi().alert("Select a vendor row (not header).");
    return;
  }

  const vendorName = crmSheet.getRange(row, 2).getValue(); // Column B

  if (!vendorName) {
    SpreadsheetApp.getUi().alert("Vendor name missing in column B.");
    return;
  }

  const vendorId = "v_" + Utilities.getUuid().slice(0, 8);
  crmSheet.getRange(row, 1).setValue(vendorId);

  // Use reusable core function
  createVendorBlockInNotes(vendorName, notesSheet, templateSheet);

  SpreadsheetApp.getUi().alert("Vendor created successfully!");
}
