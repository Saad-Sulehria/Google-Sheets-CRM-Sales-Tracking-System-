/**
 * Mailing List Sync Module
 * 
 * Synchronizes vendors from an external "Mailing List" spreadsheet to the CRM.
 * - Reads external sheet (ID provided by user)
 * - Checks for "Ready" vendors (Status Richard's Email OR TF Email not empty)
 * - Adds NEW vendors to CRM and Meeting Notes
 * - Prevents duplicates (Case-Insensitive)
 */

const SYNC_CONFIG = {
    sourceSpreadsheetId: "1Q54gLxKZJvnSDUn5Ackyg08a4SOnhIaJiNu783j-T-o",
    sourceSheetName: "CEOs HC -300",
    headerNames: {
        vendor: "Vendor",              // Target header for Vendor Name
        title: "Title",                // Target header for Title
        email: "Email",                // Target header for Email (Exact)
        status: "Status Richard's Email", // Trigger column 1
        tfEmail: "TF Email"            // Trigger column 2
    }
};

/**
 * Main function to sync vendors from Mailing List
 * Can be run manually from menu or via time trigger
 */
function syncMailingList() {


    try {
        Logger.log("Starting Mailing List Sync...");

        // 1. Open Source Spreadsheet
        const sourceSS = SpreadsheetApp.openById(SYNC_CONFIG.sourceSpreadsheetId);
        const sourceSheet = sourceSS.getSheetByName(SYNC_CONFIG.sourceSheetName);

        if (!sourceSheet) {
            console.error(`Missing Sheet: ${SYNC_CONFIG.sourceSheetName}`);
            throw new Error(`Source sheet "${SYNC_CONFIG.sourceSheetName}" not found. Available sheets: ` +
                sourceSS.getSheets().map(s => s.getName()).join(", "));
        }

        // 2. Read Headers & Map Columns Dynamically
        const lastCol = sourceSheet.getLastColumn();
        const headerRow = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];

        const colMap = {};

        // Helper to find column index (0-based) by EXACT name (case-insensitive trim)
        function findColIndex(name) {
            for (let i = 0; i < headerRow.length; i++) {
                if (String(headerRow[i]).trim().toLowerCase() === name.toLowerCase()) {
                    return i;
                }
            }
            return -1;
        }

        colMap.vendor = findColIndex(SYNC_CONFIG.headerNames.vendor);
        colMap.title = findColIndex(SYNC_CONFIG.headerNames.title);
        colMap.email = findColIndex(SYNC_CONFIG.headerNames.email);
        colMap.status = findColIndex(SYNC_CONFIG.headerNames.status);
        colMap.tfEmail = findColIndex(SYNC_CONFIG.headerNames.tfEmail);

        Logger.log("Column Mapping Found: " + JSON.stringify(colMap));

        // Validate critical columns
        if (colMap.vendor === -1) {
            let errorMsg = `Could not find required column "${SYNC_CONFIG.headerNames.vendor}"\n`;
            errorMsg += `\nFound headers: ${headerRow.slice(0, 10).join(", ")}...`;
            throw new Error(errorMsg);
        }

        // Note: Triggers are optional in theory but we need at least one to be useful?
        // But logic below checks if idx > -1 so it's safe if missing, just won't trigger.
        // Let's warn if BOTH are missing.
        if (colMap.status === -1 && colMap.tfEmail === -1) {
            throw new Error("Could not find either Trigger Column: 'Status Richard's Email' or 'TF Email'");
        }

        // 3. Read Source Data
        const lastRow = sourceSheet.getLastRow();
        if (lastRow < 2) {
            return { count: 0, message: "No data in source sheet." };
        }

        // Fetch all data
        const sourceValues = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

        // 4. Process Source Data
        const newCandidates = new Map(); // Key is LOWERCASE Vendor Name

        sourceValues.forEach((row, index) => {
            const vendorName = String(row[colMap.vendor] || "").trim();
            const vendorKey = vendorName.toLowerCase();

            const status1 = colMap.status > -1 ? String(row[colMap.status] || "").trim() : "";
            const status2 = colMap.tfEmail > -1 ? String(row[colMap.tfEmail] || "").trim() : "";

            // Criteria: (Status1 OR Status2 not empty) AND Vendor Name not empty
            if ((status1 !== "" || status2 !== "") && vendorName !== "") {
                if (!newCandidates.has(vendorKey)) {
                    newCandidates.set(vendorKey, {
                        name: vendorName, // Keep original casing for display/creation
                        title: colMap.title > -1 ? String(row[colMap.title]).trim() : "",
                        email: colMap.email > -1 ? String(row[colMap.email]).trim() : ""
                    });
                }
            }
        });

        Logger.log(`Found ${newCandidates.size} candidates with trigger conditions met.`);

        // 5. Read Existing CRM Vendors to check for duplicates
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const crmSheet = ss.getSheetByName("CRM");
        const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
        const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

        if (!crmSheet || !notesSheet || !templateSheet) {
            throw new Error("Missing local sheets: CRM, Meeting Notes, or Template.");
        }

        const crmLastRow = crmSheet.getLastRow();
        const existingVendorsLower = new Set();

        if (crmLastRow > 1) {
            // Read Column B (Vendor Name)
            const crmValues = crmSheet.getRange(2, 2, crmLastRow - 1, 1).getValues();
            crmValues.forEach(row => {
                const vName = String(row[0]).trim();
                if (vName) existingVendorsLower.add(vName.toLowerCase());
            });
        }

        Logger.log(`Found ${existingVendorsLower.size} existing vendors in CRM.`);

        // 6. Identify Truly New Vendors
        const vendorsToAdd = [];
        for (const [key, data] of newCandidates) {
            if (!existingVendorsLower.has(key)) {
                vendorsToAdd.push(data);
            }
        }

        if (vendorsToAdd.length === 0) {
            const msg = `Checked ${newCandidates.size} candidates. All matches existing CRM vendors.`;
            Logger.log(msg);
            return { count: 0, message: msg };
        }

        Logger.log(`Adding ${vendorsToAdd.length} NEW vendors...`);

        // 7. Create New Vendors
        let createdCount = 0;

        vendorsToAdd.forEach(vendor => {
            try {
                // A. Add to CRM Sheet - Using "Fill Holes" logic
                const targetRow = getNextAvailableRow(crmSheet);

                // Check for CRM sheet expansion
                if (targetRow > crmSheet.getMaxRows()) {
                    const rowsToAdd = targetRow - crmSheet.getMaxRows() + 20; // Add buffer
                    crmSheet.insertRowsAfter(crmSheet.getMaxRows(), rowsToAdd);
                    Logger.log(`Expanded CRM Sheet by ${rowsToAdd} rows.`);
                }

                const vendorId = "v_" + Utilities.getUuid().slice(0, 8);

                crmSheet.getRange(targetRow, 1).setValue(vendorId);      // Col A: ID
                crmSheet.getRange(targetRow, 2).setValue(vendor.name);   // Col B: Name
                crmSheet.getRange(targetRow, 5).setValue(vendor.title);  // Col E: Title
                crmSheet.getRange(targetRow, 6).setValue(vendor.email);  // Col F: Email

                // B. Add to Meeting Notes Sheet
                createVendorBlockInNotes(vendor.name, notesSheet, templateSheet);

                createdCount++;
                Logger.log(`Created vendor: ${vendor.name} at CRM Row ${targetRow}`);

            } catch (err) {
                Logger.log(`Failed to create vendor ${vendor.name}: ${err.message}`);
                // Continue loop
            }
        });

        // 8. Results
        const msg = `Sync Complete.\nStats:\n- Candidates in List: ${newCandidates.size}\n- Already in CRM: ${newCandidates.size - vendorsToAdd.length}\n- Newly Added: ${createdCount}`;
        Logger.log(msg);

        return { count: createdCount, message: msg };

    } catch (e) {
        Logger.log("Error in syncMailingList: " + e.toString());
        throw e; // Rethrow to let caller handle
    }
}

/**
 * Setup 2-Hour Trigger
 */
function setupMailingListTrigger() {
    const ui = SpreadsheetApp.getUi();
    const functionName = "syncMailingList";

    // 1. Delete existing triggers for this function to prevent duplicates
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === functionName) {
            ScriptApp.deleteTrigger(triggers[i]);
            deletedCount++;
        }
    }

    // 2. Create new trigger (Every 2 Hours)
    ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyHours(2)
        .create();

    ui.alert(`✅ Sync Trigger Set!\n\nWill run every 2 hours.\n(Replaced ${deletedCount} existing triggers)`);
}

/**
 * Helper to find the first row where Column B (Vendor Name) is empty.
 * Starts checking from Row 2.
 */
function getNextAvailableRow(sheet) {
    const lastRow = sheet.getLastRow();
    // Start from row 2 (skip header)
    if (lastRow < 2) return 2;

    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Get Column B

    for (let i = 0; i < values.length; i++) {
        if (values[i][0] === "") {
            return i + 2; // i is 0-indexed relative to range starting at Row 2
        }
    }

    // If no holes found, append to end
    return lastRow + 1;
}

/**
 * Menu wrapper for manual sync
 */
function manualSyncMailingList() {
    const ui = SpreadsheetApp.getUi();
    try {
        const result = syncMailingList();
        if (result && result.count > 0) {
            ui.alert(`✅ Sync Complete\n\n${result.count} new vendors added.`);
        } else if (result) {
            ui.alert(`✅ Sync Complete\n\n${result.message}`);
        }
    } catch (e) {
        ui.alert(`❌ Sync Failed: ${e.message}`);
    }
}

/**
 * Debug Helper: Check Columns
 */
function debugMailingListColumnMapping() {
    const ui = SpreadsheetApp.getUi();
    try {
        const sourceSS = SpreadsheetApp.openById(SYNC_CONFIG.sourceSpreadsheetId);
        const sourceSheet = sourceSS.getSheetByName(SYNC_CONFIG.sourceSheetName);
        if (!sourceSheet) throw new Error("Sheet '" + SYNC_CONFIG.sourceSheetName + "' not found.");

        const lastCol = sourceSheet.getLastColumn();
        const headerRow = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];

        let msg = `[DEBUG MAPPING]\nSheet: ${SYNC_CONFIG.sourceSheetName}\nTotal Cols: ${lastCol}\n\n`;

        // Show indices of searches
        const targets = [
            SYNC_CONFIG.headerNames.vendor,
            SYNC_CONFIG.headerNames.status,
            SYNC_CONFIG.headerNames.tfEmail,
            SYNC_CONFIG.headerNames.email
        ];

        targets.forEach(t => {
            const idx = headerRow.findIndex(h => String(h).trim().toLowerCase() === t.toLowerCase());
            msg += `Header "${t}": ${idx > -1 ? "FOUND at Col " + (idx + 1) : "NOT FOUND"}\n`;
        });

        // Show first 5 headers
        msg += `\nFirst 5 Headers: ${headerRow.slice(0, 5).join(", ")}`;

        ui.alert(msg);
    } catch (e) {
        ui.alert("Debug Failed: " + e.message);
    }
}
