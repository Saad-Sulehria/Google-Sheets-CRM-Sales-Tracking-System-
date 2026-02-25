function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("⚡ CRM Tools")
    .addItem("Create Filtered View", "createFilteredView")
    .addSeparator()
    .addItem("Rank Vendors", "rankVendors")
    .addItem("Move Vendor to Sheet", "moveVendorToSheet")
    .addItem("Remove Vendor", "removeVendor")
    .addItem("Sync Sheets", "syncSheets")
    .addSeparator()
    .addItem("Expand New Sheet Vendors", "expandNewSheetVendors")
    .addToUi();

  // Show batch update completion message if a batch finished while sheet was closed
  try {
    const props = PropertiesService.getDocumentProperties();
    const completionFlag = props.getProperty("BATCH_UPDATE_COMPLETE");
    if (completionFlag) {
      props.deleteProperty("BATCH_UPDATE_COMPLETE");
      const info = JSON.parse(completionFlag);
      ui.alert(
        "✅ Batch Update Complete\n\n" +
        info.vendorCount + " vendor blocks were updated successfully.\n" +
        "Notes Links have been reconnected automatically."
      );
    }
  } catch (e) {
    // Non-fatal — don't block the menu from opening
    Logger.log("Could not check batch completion flag: " + e.message);
  }
}
