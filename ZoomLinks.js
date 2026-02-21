/**
 * Zoom Links Module
 * 
 * Handles creation and linking of Zoom meeting links between
 * Meeting Notes and the dedicated Zoom Links sheet.
 */

const ZOOM_CONFIG = {
    sheetName: "Zoom Links",
    blockSize: 12,       // 1 vendor header + 1 column header + 9 data rows + 1 blank
    dataRows: 9,         // Number of person/link rows per vendor
    vendorNameCol: 1,    // Column A
    personNameCol: 2,    // Column B
    meetingLinkCol: 3,   // Column C
    zoomLabel: "Zoom"    // Label to find in template
};

/**
 * One-time setup: Creates vendor blocks in Zoom Links sheet
 * for all existing vendors in Meeting Notes.
 * 
 * Call this from menu, then remove after initial setup.
 */
function setupZoomLinksSheet() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create Zoom Links sheet
    let zoomSheet = ss.getSheetByName(ZOOM_CONFIG.sheetName);
    if (!zoomSheet) {
        zoomSheet = ss.insertSheet(ZOOM_CONFIG.sheetName);
    }

    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        ui.alert("Missing sheets: " + CONFIG.notesSheetName + " or " + CONFIG.templateSheetName);
        return;
    }

    // Extract vendor names from Meeting Notes
    const templateHeight = templateSheet.getLastRow();
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockTotalRows = 1 + templateHeight + blankRows;

    const vendorNames = [];
    const lastRow = notesSheet.getLastRow();
    let currentRow = 1;

    while (currentRow <= lastRow) {
        const vendorName = notesSheet.getRange(currentRow, 1).getValue();
        if (vendorName && String(vendorName).trim() !== "") {
            vendorNames.push(String(vendorName).trim());
        }
        currentRow += blockTotalRows;
    }

    if (vendorNames.length === 0) {
        ui.alert("No vendors found in " + CONFIG.notesSheetName);
        return;
    }

    // Confirm with user
    const result = ui.alert(
        "Setup Zoom Links",
        "This will create " + vendorNames.length + " vendor blocks in the Zoom Links sheet.\n\nExisting data will be cleared. Proceed?",
        ui.ButtonSet.YES_NO
    );

    if (result !== ui.Button.YES) {
        return;
    }

    // Clear and setup sheet
    zoomSheet.clear();

    // Ensure enough rows
    const requiredRows = vendorNames.length * ZOOM_CONFIG.blockSize;
    const currentMaxRows = zoomSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        zoomSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    // Create vendor blocks
    let zoomRow = 1;
    for (const vendorName of vendorNames) {
        createZoomVendorBlock_(zoomSheet, zoomRow, vendorName);
        zoomRow += ZOOM_CONFIG.blockSize;
    }

    // Set column widths
    zoomSheet.setColumnWidth(ZOOM_CONFIG.vendorNameCol, 200);
    zoomSheet.setColumnWidth(ZOOM_CONFIG.personNameCol, 150);
    zoomSheet.setColumnWidth(ZOOM_CONFIG.meetingLinkCol, 300);

    ui.alert("✅ Zoom Links sheet setup complete!\n\n" + vendorNames.length + " vendor blocks created.");
}

/**
 * Creates a single vendor block in Zoom Links sheet
 */
function createZoomVendorBlock_(sheet, startRow, vendorName) {
    // Row 1: Vendor Name (Column A, bold, magenta background across 3 columns)
    const vendorHeaderRange = sheet.getRange(startRow, 1, 1, 3);
    vendorHeaderRange.setBackground(CONFIG.vendorHeaderBgColor);

    const vendorCell = sheet.getRange(startRow, ZOOM_CONFIG.vendorNameCol);
    vendorCell.setValue(vendorName);
    vendorCell.setFontWeight("bold");
    vendorCell.setFontColor("#ffffff");

    // Row 2: Column Headers (bold, #d9d2e9 background across columns B and C)
    const headerRow = startRow + 1;
    sheet.getRange(headerRow, ZOOM_CONFIG.personNameCol).setValue("Person Name").setFontWeight("bold");
    sheet.getRange(headerRow, ZOOM_CONFIG.meetingLinkCol).setValue("Meeting Link").setFontWeight("bold");
    sheet.getRange(headerRow, ZOOM_CONFIG.personNameCol, 1, 2).setBackground("#d9d2e9");

    // Rows 3-11: Data rows (empty, ready for input)
    // No action needed, cells are already empty
}

/**
 * Creates a vendor block in Zoom Links sheet for a new vendor.
 * Called from createNewVendor().
 */
function createZoomVendorBlockForNewVendor(vendorName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let zoomSheet = ss.getSheetByName(ZOOM_CONFIG.sheetName);

    if (!zoomSheet) {
        // Create sheet if it doesn't exist
        zoomSheet = ss.insertSheet(ZOOM_CONFIG.sheetName);
        zoomSheet.setColumnWidth(ZOOM_CONFIG.vendorNameCol, 200);
        zoomSheet.setColumnWidth(ZOOM_CONFIG.personNameCol, 150);
        zoomSheet.setColumnWidth(ZOOM_CONFIG.meetingLinkCol, 300);
    }

    // Calculate next block start position based on existing vendor count
    // Each block is ZOOM_CONFIG.blockSize rows (12 = 1 vendor header + 1 column header + 9 data rows + 1 blank)
    const lastRow = zoomSheet.getLastRow();

    let startRow = 1;
    if (lastRow > 0) {
        // Count existing vendor blocks (each starts at row 1, 13, 25, ...)
        // Next block starts at: ((current blocks) * blockSize) + 1
        const existingBlocks = Math.ceil(lastRow / ZOOM_CONFIG.blockSize);
        startRow = existingBlocks * ZOOM_CONFIG.blockSize + 1;
    }

    // Ensure enough rows
    const requiredRows = startRow + ZOOM_CONFIG.blockSize;
    const currentMaxRows = zoomSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        zoomSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    createZoomVendorBlock_(zoomSheet, startRow, vendorName);
}

/**
 * Syncs person names and links Zoom cells in Meeting Notes to Zoom Links.
 * 
 * 1. Copies "Name" values from Contact Info to "Person Name" in Zoom Links
 * 2. Links non-empty Zoom cells to matching person's Meeting Link row
 * 
 * Called from reconnectNotesLinksInternal().
 */
function linkZoomCells(notesSheet, templateSheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const zoomSheet = ss.getSheetByName(ZOOM_CONFIG.sheetName);

    if (!zoomSheet) {
        Logger.log("Zoom Links sheet not found. Skipping Zoom linking.");
        return;
    }

    // Find "Name" and "Zoom" columns in template
    const templateHeight = templateSheet.getLastRow();
    const templateWidth = templateSheet.getLastColumn();
    const templateData = templateSheet.getRange(1, 1, templateHeight, templateWidth).getValues();

    let nameColOffset = -1;  // Column offset for "Name" header
    let zoomColOffset = -1;  // Column offset for "Zoom" header
    let contactInfoHeaderRow = -1;  // Row where Contact Info headers are

    // Find the Contact Info section headers (Name, Title, Email, LinkedIn, GEO, Zoom)
    for (let r = 0; r < templateHeight; r++) {
        for (let c = 0; c < templateWidth; c++) {
            const val = String(templateData[r][c]).trim();
            if (val === "Name" && nameColOffset === -1) {
                nameColOffset = c + 2;  // +2 because template is in B onwards in notes
                contactInfoHeaderRow = r;
            }
            if (val === ZOOM_CONFIG.zoomLabel && zoomColOffset === -1) {
                zoomColOffset = c + 2;
            }
        }
        if (nameColOffset !== -1 && zoomColOffset !== -1) break;
    }

    if (nameColOffset === -1 || zoomColOffset === -1) {
        Logger.log("Name or Zoom column not found in template. Name: " + nameColOffset + ", Zoom: " + zoomColOffset);
        return;
    }

    Logger.log("Found Name at col " + nameColOffset + ", Zoom at col " + zoomColOffset + ", header row " + contactInfoHeaderRow);

    // Build vendor position map in Zoom Links sheet
    const zoomLastRow = zoomSheet.getLastRow();
    const zoomVendorMap = {}; // vendorName -> startRow in Zoom Links

    let zoomCurrentRow = 1;
    while (zoomCurrentRow <= zoomLastRow) {
        const vendorName = zoomSheet.getRange(zoomCurrentRow, ZOOM_CONFIG.vendorNameCol).getValue();
        if (vendorName && String(vendorName).trim() !== "") {
            zoomVendorMap[String(vendorName).trim()] = zoomCurrentRow;
        }
        zoomCurrentRow += ZOOM_CONFIG.blockSize;
    }

    // Calculate Meeting Notes block structure
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockTotalRows = 1 + templateHeight + blankRows;
    const notesLastRow = notesSheet.getLastRow();
    let notesCurrentRow = 1;

    // Iterate through each vendor in Meeting Notes
    while (notesCurrentRow <= notesLastRow) {
        const vendorName = notesSheet.getRange(notesCurrentRow, 1).getValue();
        if (!vendorName || String(vendorName).trim() === "") {
            notesCurrentRow += blockTotalRows;
            continue;
        }

        const vendorKey = String(vendorName).trim();
        const zoomBlockStart = zoomVendorMap[vendorKey];

        if (!zoomBlockStart) {
            Logger.log("Vendor not found in Zoom Links: " + vendorKey);
            notesCurrentRow += blockTotalRows;
            continue;
        }

        // Process Contact Info rows (rows after the header row in template)
        // Header row is at contactInfoHeaderRow, data starts at contactInfoHeaderRow + 1
        for (let i = 0; i < ZOOM_CONFIG.dataRows; i++) {
            const dataRowOffset = contactInfoHeaderRow + 1 + i;  // Template row (0-indexed)
            if (dataRowOffset >= templateHeight) break;

            // Calculate actual row in Meeting Notes
            const notesDataRow = notesCurrentRow + 1 + dataRowOffset;  // +1 for vendor header

            // Get Name value from Meeting Notes
            const nameValue = notesSheet.getRange(notesDataRow, nameColOffset).getValue();
            const nameStr = nameValue ? String(nameValue).trim() : "";

            // Get Zoom value from Meeting Notes
            const zoomValue = notesSheet.getRange(notesDataRow, zoomColOffset).getValue();
            const zoomStr = zoomValue ? String(zoomValue).trim() : "";

            // Calculate corresponding row in Zoom Links
            // Data rows start at zoomBlockStart + 2 (vendor header + column header)
            const zoomDataRow = zoomBlockStart + 2 + i;

            // Sync person name to Zoom Links
            if (nameStr !== "") {
                zoomSheet.getRange(zoomDataRow, ZOOM_CONFIG.personNameCol).setValue(nameStr);
            }

            // Link Zoom cell only if non-empty and not already linked
            if (zoomStr !== "") {
                const zoomCell = notesSheet.getRange(notesDataRow, zoomColOffset);

                // Check if already has a hyperlink (Rich Text)
                const existingRichText = zoomCell.getRichTextValue();
                const existingLink = existingRichText ? existingRichText.getLinkUrl() : null;

                if (!existingLink) {
                    // Create hyperlink to Zoom Links sheet, preserving existing text
                    const linkUrl = "#gid=" + zoomSheet.getSheetId() + "&range=C" + zoomDataRow;
                    const richText = SpreadsheetApp.newRichTextValue()
                        .setText(zoomStr)
                        .setLinkUrl(linkUrl)
                        .build();
                    zoomCell.setRichTextValue(richText);
                }
            }
        }

        notesCurrentRow += blockTotalRows;
    }

    Logger.log("Zoom cells synced and linked successfully.");
}

/**
 * Safely rename a sheet and update CONFIG references.
 * Prevents code breakage when users want to change sheet names.
 */
function renameSheet() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Define the sheets that can be renamed (those referenced in CONFIG)
    const managedSheets = {
        "Meeting Notes": "notesSheetName",
        "Template_LeftBlock": "templateSheetName",
        "Zoom Links": "zoomSheetName (ZOOM_CONFIG)"
    };

    // Build selection prompt
    let sheetList = "Which sheet do you want to rename?\n\n";
    const sheetNames = Object.keys(managedSheets);
    for (let i = 0; i < sheetNames.length; i++) {
        sheetList += (i + 1) + ". " + sheetNames[i] + "\n";
    }
    sheetList += "\nEnter the number:";

    const selectionResponse = ui.prompt("🔄 Rename Sheet", sheetList, ui.ButtonSet.OK_CANCEL);
    if (selectionResponse.getSelectedButton() !== ui.Button.OK) return;

    const selection = parseInt(selectionResponse.getResponseText().trim());
    if (isNaN(selection) || selection < 1 || selection > sheetNames.length) {
        ui.alert("Invalid selection.");
        return;
    }

    const oldName = sheetNames[selection - 1];
    const sheet = ss.getSheetByName(oldName);

    if (!sheet) {
        ui.alert("Sheet '" + oldName + "' not found. It may have been renamed already.");
        return;
    }

    // Get new name
    const newNameResponse = ui.prompt(
        "🔄 Rename Sheet",
        "Current name: " + oldName + "\n\nEnter the NEW name:",
        ui.ButtonSet.OK_CANCEL
    );
    if (newNameResponse.getSelectedButton() !== ui.Button.OK) return;

    const newName = newNameResponse.getResponseText().trim();
    if (!newName || newName === oldName) {
        ui.alert("Invalid or same name. No changes made.");
        return;
    }

    // Check if new name already exists
    if (ss.getSheetByName(newName)) {
        ui.alert("A sheet with the name '" + newName + "' already exists. Choose a different name.");
        return;
    }

    // Rename the sheet
    sheet.setName(newName);

    // Show instruction for updating code
    ui.alert(
        "✅ Sheet Renamed Successfully!\n\n" +
        "The sheet has been renamed from:\n'" + oldName + "' → '" + newName + "'\n\n" +
        "⚠️ IMPORTANT: You must also update the code!\n\n" +
        "Go to Extensions > Apps Script and update the CONFIG object:\n" +
        "- For Meeting Notes: CONFIG.notesSheetName\n" +
        "- For Template: CONFIG.templateSheetName\n" +
        "- For Zoom Links: ZOOM_CONFIG.sheetName\n\n" +
        "Change the value to: \"" + newName + "\""
    );

    Logger.log("Sheet renamed: " + oldName + " → " + newName);
}
