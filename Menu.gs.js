function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ CRM Tools")
    .addItem("Create New Vendor", "createNewVendor")
    .addItem("Update ALL Vendor Blocks", "updateAllVendorBlocks")
    .addItem("Sync Mailing List", "manualSyncMailingList")
    .addItem("Reconnect Notes Links", "reconnectNotesLinks")
    .addItem("Rename Header", "renameHeader")
    .addSeparator()
    .addItem("Rename Sheet", "renameSheet")
    .addItem("Manual Backup", "manualBackup")
    .addToUi();
}
