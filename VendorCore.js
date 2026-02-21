/**
 * Core Logic for creating vendors
 * Decoupled from UI for reusability in Mailing List Sync
 */

/**
 * Creates a new vendor in Meeting Notes sheet
 * @param {string} vendorName - Name of the vendor
 * @param {Sheet} notesSheet - Meeting Notes sheet object
 * @param {Sheet} templateSheet - Template sheet object
 * @returns {number} The starting row of the new block
 */
function createVendorBlockInNotes(vendorName, notesSheet, templateSheet) {
    // Determine start row with configured spacing
    const lastRow = notesSheet.getLastRow();
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const startRow = lastRow > 0 ? lastRow + blankRows + 1 : 1;

    // Compute template dimensions (include merged cells)
    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);

    // Check if sheet needs expansion
    // We need space for: Vendor Header (1 row) + Template (templateHeight)
    // Total needed rows = startRow + templateHeight
    const neededRows = startRow + templateHeight;
    const currentMaxRows = notesSheet.getMaxRows();

    if (neededRows > currentMaxRows) {
        const rowsToAdd = neededRows - currentMaxRows + 10; // Add buffer
        notesSheet.insertRowsAfter(currentMaxRows, rowsToAdd);
        Logger.log(`Expanded Meeting Notes sheet by ${rowsToAdd} rows.`);
    }

    // Use consistency in header formatting (magenta row)
    const headerCell = notesSheet.getRange(startRow, 1);
    headerCell.setValue(vendorName);
    headerCell.setFontWeight("bold");
    headerCell.setFontColor("#ffffff");
    notesSheet.getRange(startRow, 1, 1, fullHeaderWidth).setBackground(CONFIG.vendorHeaderBgColor);

    // Copy template dynamically (include merged cell width)
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);

    // IMPORTANT: Template copied to column **2 (B)**, starting **1 ROW BELOW** vendor header.
    // Matching new "Exclusive Header Row" logic.
    // Row X: Vendor Name (A)
    // Row X+1: Template (B)
    templateRange.copyTo(
        notesSheet.getRange(startRow + 1, 2),
        { contentsOnly: false }
    );

    return startRow;
}
