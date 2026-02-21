/**
 * PhilSync.js — Temporary bi-directional sync between Meeting Notes and a "Phil" sheet.
 * 
 * Vendors with "Phil" in column E of Meeting Notes are cloned to a separate sheet.
 * A 30-minute trigger keeps both sheets in sync using a 3-way merge pattern.
 * 
 * Functions:
 *   createPhilSheet()   — Create the Phil sheet, snapshot, and trigger
 *   syncPhilSheet()     — 30-min sync (headless-safe)
 *   removePhilSync()    — Cleanup: remove trigger, cache, optionally Phil sheet
 */

const PHIL_CONFIG = {
    philSheetName: "Phil",
    cacheSheetName: "_PhilSyncCache",
    triggerFunctionName: "syncPhilSheet",
    triggerMinutes: 30,
    filterColumn: 5,          // Column E (1-indexed)
    filterKeyword: "phil"     // Case-insensitive match
};

// ─────────────────────────────────────────────────────────────────────────────
// 1. CREATE PHIL SHEET
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create the Phil sheet from Meeting Notes, store a sync snapshot, and set up
 * the 30-minute trigger. Called from the menu.
 */
function createPhilSheet() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        ui.alert("Missing required sheets: " + CONFIG.notesSheetName + " or " + CONFIG.templateSheetName);
        return;
    }

    // Find Phil vendors from raw sheet data
    const philVendorNames = findPhilVendorNames_(notesSheet);
    if (philVendorNames.size === 0) {
        ui.alert("No vendors found with \"Phil\" in column E.");
        return;
    }

    // Extract all vendor data, then filter to Phil vendors
    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const allVendors = extractAllVendorData(notesSheet, templateHeight);
    const philVendors = allVendors.filter(v => philVendorNames.has(v.name));

    if (philVendors.length === 0) {
        ui.alert("No matching vendor data found for Phil vendors.");
        return;
    }

    // Create or clear the Phil sheet
    let philSheet = ss.getSheetByName(PHIL_CONFIG.philSheetName);
    if (philSheet) {
        philSheet.clear();
    } else {
        philSheet = ss.insertSheet(PHIL_CONFIG.philSheetName);
    }

    // Ensure enough rows
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;
    const requiredRows = philVendors.length * blockHeight;
    const currentMaxRows = philSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        philSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    // Write vendor blocks
    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);
    rebuildVendorBlocks(philSheet, templateSheet, templateRange, philVendors, fullHeaderWidth);

    // Inject data
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);
    injectDataByLabel(philSheet, philVendors, labelMap, templateHeight, templateContentWidth);

    // Copy column widths
    const lastUsedCol = 1 + Math.max(templateContentWidth, templateCopyWidth);
    for (let c = 1; c <= lastUsedCol; c++) {
        philSheet.setColumnWidth(c, notesSheet.getColumnWidth(c));
    }

    // Copy row heights from first block in Meeting Notes
    const srcBlockHeight = 1 + templateHeight + blankRows;
    const srcHeights = [];
    for (let r = 1; r <= srcBlockHeight; r++) {
        srcHeights.push(notesSheet.getRowHeight(r));
    }
    for (let v = 0; v < philVendors.length; v++) {
        for (let r = 0; r < srcBlockHeight; r++) {
            philSheet.setRowHeight(v * blockHeight + r + 1, srcHeights[r]);
        }
    }

    // Clear trailing rows
    const lastVendorEndRow = philVendors.length * blockHeight;
    const sheetMaxRow = philSheet.getMaxRows();
    if (sheetMaxRow > lastVendorEndRow) {
        philSheet.getRange(lastVendorEndRow + 1, 1, sheetMaxRow - lastVendorEndRow, philSheet.getMaxColumns())
            .clearContent().clearDataValidations();
    }

    // Store initial snapshot for 3-way merge
    storeSnapshot_(ss, philVendors);

    // Set up 30-minute trigger (remove existing first)
    removeTriggers_(PHIL_CONFIG.triggerFunctionName);
    ScriptApp.newTrigger(PHIL_CONFIG.triggerFunctionName)
        .timeBased()
        .everyMinutes(PHIL_CONFIG.triggerMinutes)
        .create();

    ui.alert(
        "✅ Phil Sheet Created\n\n" +
        "Vendors: " + philVendors.length + "\n" +
        "Sync trigger: every " + PHIL_CONFIG.triggerMinutes + " minutes\n\n" +
        "Changes in either sheet will be synced automatically."
    );
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. SYNC (30-MINUTE TRIGGER)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Bi-directional sync between Meeting Notes and Phil sheet.
 * Uses 3-way merge: compares both sheets against a stored snapshot to detect
 * which side changed. Headless-safe (no UI calls).
 */
function syncPhilSheet() {
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const philSheet = ss.getSheetByName(PHIL_CONFIG.philSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !philSheet || !templateSheet) {
        Logger.log("PhilSync: Missing sheets, skipping sync.");
        return;
    }

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);

    // Extract current data from both sheets
    const allNotesVendors = extractAllVendorData(notesSheet, templateHeight);
    const philVendors = extractAllVendorData(philSheet, templateHeight);

    // Build lookup by name
    const notesMap = {};
    for (const v of allNotesVendors) notesMap[v.name] = v;
    const philMap = {};
    for (const v of philVendors) philMap[v.name] = v;

    // Load snapshot
    const snapshot = loadSnapshot_(ss);
    if (!snapshot || Object.keys(snapshot).length === 0) {
        Logger.log("PhilSync: No snapshot found, storing current state.");
        storeSnapshot_(ss, philVendors);
        return;
    }

    let notesToPhilCount = 0;
    let philToNotesCount = 0;

    // For each vendor, build a MERGED version that combines field-level changes
    // from both sides. This ensures true bi-directional sync.
    const vendorsToUpdateInPhil = [];
    const vendorsToUpdateInNotes = [];

    for (const philVendor of philVendors) {
        const name = philVendor.name;
        const notesVendor = notesMap[name];
        const snapshotVendor = snapshot[name];

        if (!notesVendor || !snapshotVendor) continue;

        let needsPhilUpdate = false;   // Notes changed some field → push to Phil
        let needsNotesUpdate = false;  // Phil changed some field → push to Notes

        // Build merged vendor data for EACH side at the FIELD level
        // mergedForPhil = what Phil should become (start from Phil, overlay Notes changes)
        // mergedForNotes = what Notes should become (start from Notes, overlay Phil changes)
        const mergedForPhil = JSON.parse(JSON.stringify(philVendor));
        const mergedForNotes = JSON.parse(JSON.stringify(notesVendor));

        for (let d = 0; d < philVendor.data.length; d++) {
            const philItem = philVendor.data[d];
            const label = philItem.label;

            const notesItem = notesVendor.data.find(x => x.label === label);
            const snapItem = snapshotVendor.find(x => x.label === label);
            if (!notesItem || !snapItem) continue;

            const philVals = philItem.values || [];
            const notesVals = notesItem.values || [];
            const snapVals = snapItem.values || [];

            // Find matching data entries in our merged copies
            const mergedPhilItem = mergedForPhil.data.find(x => x.label === label);
            const mergedNotesItem = mergedForNotes.data.find(x => x.label === label);

            // Only compare positions that BOTH live sheets have — trailing indices
            // beyond the shorter array are just artifacts of different sheet widths
            // and must NOT be treated as "changed to empty".
            const compareLen = Math.min(
                Math.max(philVals.length, snapVals.length),
                Math.max(notesVals.length, snapVals.length)
            );

            for (let v = 0; v < compareLen; v++) {
                const pv = normalizeForCompare_(philVals[v]);
                const nv = normalizeForCompare_(notesVals[v]);
                const sv = normalizeForCompare_(snapVals[v]);

                if (nv === pv) continue; // No difference — skip

                if (nv !== sv && pv === sv) {
                    // Meeting Notes changed this field → push to Phil
                    if (mergedPhilItem && mergedPhilItem.values) {
                        mergedPhilItem.values[v] = notesVals[v];
                        if (notesItem.valueFormulas && notesItem.valueFormulas[v]) {
                            mergedPhilItem.valueFormulas = mergedPhilItem.valueFormulas || [];
                            mergedPhilItem.valueFormulas[v] = notesItem.valueFormulas[v];
                        }
                    }
                    needsPhilUpdate = true;
                } else if (pv !== sv && nv === sv) {
                    // Phil changed this field → push to Notes
                    if (mergedNotesItem && mergedNotesItem.values) {
                        mergedNotesItem.values[v] = philVals[v];
                        if (philItem.valueFormulas && philItem.valueFormulas[v]) {
                            mergedNotesItem.valueFormulas = mergedNotesItem.valueFormulas || [];
                            mergedNotesItem.valueFormulas[v] = philItem.valueFormulas[v];
                        }
                    }
                    needsNotesUpdate = true;
                } else {
                    // Both changed → Meeting Notes wins
                    if (mergedPhilItem && mergedPhilItem.values) {
                        mergedPhilItem.values[v] = notesVals[v];
                    }
                    needsPhilUpdate = true;
                }
            }

            // Merged cell content (same field-level logic)
            const pMerged = normalizeForCompare_(philItem.mergedCellContent);
            const nMerged = normalizeForCompare_(notesItem.mergedCellContent);
            const sMerged = normalizeForCompare_(snapItem.mergedCellContent);

            if (nMerged !== pMerged) {
                if (nMerged !== sMerged && pMerged === sMerged) {
                    if (mergedPhilItem) mergedPhilItem.mergedCellContent = notesItem.mergedCellContent;
                    needsPhilUpdate = true;
                } else if (pMerged !== sMerged && nMerged === sMerged) {
                    if (mergedNotesItem) mergedNotesItem.mergedCellContent = philItem.mergedCellContent;
                    needsNotesUpdate = true;
                } else {
                    if (mergedPhilItem) mergedPhilItem.mergedCellContent = notesItem.mergedCellContent;
                    needsPhilUpdate = true;
                }
            }
        }

        if (needsPhilUpdate) {
            vendorsToUpdateInPhil.push(mergedForPhil);
            notesToPhilCount++;
        }
        if (needsNotesUpdate) {
            vendorsToUpdateInNotes.push(mergedForNotes);
            philToNotesCount++;
        }
    }

    // Apply changes
    const blankRows = CONFIG.blankRowsAfterVendor || 0;

    if (vendorsToUpdateInPhil.length > 0) {
        applyUpdatesToSheet_(philSheet, philVendors, vendorsToUpdateInPhil, labelMap, templateHeight, templateContentWidth, blankRows);
        Logger.log("PhilSync: Updated " + vendorsToUpdateInPhil.length + " vendors in Phil sheet (from Meeting Notes).");
    }

    if (vendorsToUpdateInNotes.length > 0) {
        applyUpdatesToSheet_(notesSheet, allNotesVendors, vendorsToUpdateInNotes, labelMap, templateHeight, templateContentWidth, blankRows);
        Logger.log("PhilSync: Updated " + vendorsToUpdateInNotes.length + " vendors in Meeting Notes (from Phil sheet).");
    }

    // Re-read and store updated snapshot
    const updatedPhilVendors = extractAllVendorData(philSheet, templateHeight);
    storeSnapshot_(ss, updatedPhilVendors);

    Logger.log("PhilSync: Sync complete. Notes→Phil: " + notesToPhilCount + ", Phil→Notes: " + philToNotesCount);
}

/**
 * Manually trigger the sync with UI feedback.
 */
function manualPhilSync() {
    const ui = SpreadsheetApp.getUi();
    try {
        syncPhilSheet();
        ui.alert("✅ Phil Sync Complete\n\nBi-directional changes have been synchronized.");
    } catch (e) {
        ui.alert("❌ Sync Failed: " + e.message);
    }
}

/**
 * Apply updated vendor data to specific vendors in a sheet.
 * Only writes to the vendor blocks that need updating (not full rebuild).
 */
function applyUpdatesToSheet_(sheet, allVendorsInSheet, updatedVendors, labelMap, templateHeight, templateContentWidth, blankRows) {
    const blockHeight = 1 + templateHeight + blankRows;
    const updatedNames = new Set(updatedVendors.map(v => v.name));

    for (let i = 0; i < allVendorsInSheet.length; i++) {
        const vendor = allVendorsInSheet[i];
        if (!updatedNames.has(vendor.name)) continue;

        // Find the updated data
        const updatedVendor = updatedVendors.find(v => v.name === vendor.name);
        if (!updatedVendor) continue;

        // Calculate the block start row for this vendor
        const blockStartRow = i * blockHeight + 1; // 1-indexed
        const dataStartRow = blockStartRow + 1;     // Template data starts after header

        // Write updated data using the injection pattern
        const blockRange = sheet.getRange(dataStartRow, 2, templateHeight, templateContentWidth);
        const currentValues = blockRange.getValues();
        const newValues = currentValues.map(r => [...r]);
        const formulasToSet = Array.from({ length: templateHeight }, () => Array(templateContentWidth).fill(""));
        let hasFormulas = false;

        for (const item of updatedVendor.data) {
            const templatePos = labelMap[item.label];
            if (!templatePos) continue;

            const relativeRow = templatePos.row;
            if (relativeRow < 0 || relativeRow >= templateHeight) continue;

            const valueStartCol = templatePos.valueCol;
            const values = item.values || [];
            const valueFormulas = item.valueFormulas || [];

            for (let v = 0; v < values.length; v++) {
                let val = values[v];
                const targetCol = valueStartCol + v;
                if (val === null) continue;
                if (targetCol >= newValues[relativeRow].length) continue;

                // Reconvert ISO date strings
                if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) {
                    val = new Date(val);
                }

                newValues[relativeRow][targetCol] = val;

                const formula = valueFormulas[v];
                if (formula && String(formula).trim() !== "") {
                    formulasToSet[relativeRow][targetCol] = formula;
                    hasFormulas = true;
                }
            }

            // Merged cell content
            if (item.mergedCellContent) {
                const mergedRow = relativeRow + 1;
                const mergedCol = templatePos.labelCol;
                if (mergedRow < newValues.length && mergedCol < newValues[mergedRow].length) {
                    newValues[mergedRow][mergedCol] = item.mergedCellContent;

                    const mergedFormula = item.mergedCellFormula;
                    if (mergedFormula && String(mergedFormula).trim() !== "") {
                        formulasToSet[mergedRow][mergedCol] = mergedFormula;
                        hasFormulas = true;
                    } else if (item.mergedCellRichText) {
                        formulasToSet[mergedRow][mergedCol] = item.mergedCellRichText;
                        hasFormulas = true;
                    }
                }
            }
        }

        // Save data validations, clear them (to avoid dropdown errors), write, then restore
        const savedValidations = blockRange.getDataValidations();
        blockRange.clearDataValidations();
        blockRange.setValues(newValues);

        // Write formulas cell-by-cell — batch setFormulas() overwrites ALL cells
        // with empty strings, which clears values and causes phantom drift on re-sync.
        if (hasFormulas) {
            for (let r = 0; r < templateHeight; r++) {
                for (let c = 0; c < templateContentWidth; c++) {
                    const f = formulasToSet[r][c];
                    if (f && String(f).trim() !== "") {
                        sheet.getRange(dataStartRow + r, 2 + c).setFormula(f);
                    }
                }
            }
        }

        blockRange.setDataValidations(savedValidations);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. REMOVE PHIL SYNC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Remove the Phil sync trigger and cache sheet. Optionally delete Phil sheet.
 */
function removePhilSync() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    // Remove trigger
    removeTriggers_(PHIL_CONFIG.triggerFunctionName);
    Logger.log("PhilSync: Trigger removed.");

    // Delete cache sheet
    const cacheSheet = ss.getSheetByName(PHIL_CONFIG.cacheSheetName);
    if (cacheSheet) {
        ss.deleteSheet(cacheSheet);
        Logger.log("PhilSync: Cache sheet deleted.");
    }

    // Ask about Phil sheet
    const response = ui.alert(
        "Remove Phil Sheet?",
        "The sync trigger and cache have been removed.\n\nDo you also want to delete the \"" + PHIL_CONFIG.philSheetName + "\" sheet?",
        ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
        const philSheet = ss.getSheetByName(PHIL_CONFIG.philSheetName);
        if (philSheet) {
            ss.deleteSheet(philSheet);
            Logger.log("PhilSync: Phil sheet deleted.");
        }
        ui.alert("✅ Phil Sync fully removed (sheet deleted).");
    } else {
        ui.alert("✅ Phil Sync removed. The Phil sheet has been kept for reference.");
    }
}

/**
 * Remove a specific vendor from the Phil sheet by name.
 * Deletes the entire vendor block (header row + template rows + blank row)
 * and updates the sync snapshot.
 */
function removeVendorFromPhil() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();
    const philSheet = ss.getSheetByName(PHIL_CONFIG.philSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!philSheet) {
        ui.alert("Phil sheet not found.");
        return;
    }

    // Ask for vendor name
    const response = ui.prompt(
        "Remove Vendor from Phil Sheet",
        "Enter the exact vendor name to remove:",
        ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const vendorName = response.getResponseText().trim();
    if (!vendorName) {
        ui.alert("Vendor name cannot be empty.");
        return;
    }

    // Find vendor block boundaries using bold text in column A
    const lastRow = philSheet.getLastRow();
    if (lastRow === 0) {
        ui.alert("Phil sheet is empty.");
        return;
    }

    const rawData = philSheet.getRange(1, 1, lastRow, 1).getValues();
    const rawWeights = philSheet.getRange(1, 1, lastRow, 1).getFontWeights();

    const blocks = [];
    for (let i = 0; i < rawData.length; i++) {
        const cellA = rawData[i][0];
        if (cellA && String(cellA).trim() !== "" && rawWeights[i][0] === "bold") {
            blocks.push({ name: String(cellA).trim(), startRow: i + 1 }); // 1-indexed
        }
    }

    // Find the target vendor (case-insensitive)
    const targetBlock = blocks.find(b => b.name.toLowerCase() === vendorName.toLowerCase());
    if (!targetBlock) {
        const available = blocks.map(b => b.name).join("\n• ");
        ui.alert("Vendor \"" + vendorName + "\" not found.\n\nAvailable vendors:\n• " + available);
        return;
    }

    // Calculate how many rows to delete (header + template + blank rows)
    const templateHeight = templateSheet ? templateSheet.getLastRow() : 0;
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;

    // Delete rows
    philSheet.deleteRows(targetBlock.startRow, blockHeight);
    Logger.log("PhilSync: Removed vendor \"" + targetBlock.name + "\" (" + blockHeight + " rows starting at row " + targetBlock.startRow + ").");

    // Update the sync snapshot (re-extract from Phil sheet)
    if (templateSheet) {
        const updatedPhilVendors = extractAllVendorData(philSheet, templateHeight);
        storeSnapshot_(ss, updatedPhilVendors);
    }

    ui.alert("✅ Vendor \"" + targetBlock.name + "\" has been removed from the Phil sheet.\n\nSync snapshot updated.");
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Find vendor names whose blocks contain "Phil" in column E.
 */
function findPhilVendorNames_(notesSheet) {
    const lastRow = notesSheet.getLastRow();
    const lastCol = notesSheet.getLastColumn();
    if (lastRow === 0) return new Set();

    const rawData = notesSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const rawWeights = notesSheet.getRange(1, 1, lastRow, lastCol).getFontWeights();

    // Build vendor block boundaries
    const blocks = [];
    for (let i = 0; i < rawData.length; i++) {
        const cellA = rawData[i][0];
        if (cellA && String(cellA).trim() !== "" && rawWeights[i][0] === "bold") {
            blocks.push({ name: String(cellA).trim(), startRow: i });
        }
    }
    for (let i = 0; i < blocks.length; i++) {
        blocks[i].endRow = (i + 1 < blocks.length) ? blocks[i + 1].startRow : rawData.length;
    }

    // Check column E (index 4) for "Phil"
    const matchingNames = new Set();
    for (const block of blocks) {
        for (let r = block.startRow; r < block.endRow; r++) {
            const colEVal = String(rawData[r][PHIL_CONFIG.filterColumn - 1] || "").toLowerCase();
            if (colEVal.indexOf(PHIL_CONFIG.filterKeyword) !== -1) {
                matchingNames.add(block.name);
                break;
            }
        }
    }

    Logger.log("PhilSync: Found " + matchingNames.size + " vendors with 'Phil' in column E.");
    return matchingNames;
}

/**
 * Store a snapshot of vendor data for 3-way merge.
 * Uses a hidden cache sheet to avoid DocumentProperties size limits.
 */
function storeSnapshot_(ss, vendors) {
    let cacheSheet = ss.getSheetByName(PHIL_CONFIG.cacheSheetName);
    if (cacheSheet) {
        cacheSheet.clear();
    } else {
        cacheSheet = ss.insertSheet(PHIL_CONFIG.cacheSheetName);
        cacheSheet.hideSheet();
    }

    // Store as name -> serialized data, one vendor per row
    const rows = vendors.map(v => {
        // Only store the fields we need for comparison
        const compactData = v.data.map(item => ({
            label: item.label,
            values: item.values,
            mergedCellContent: item.mergedCellContent || null
        }));
        return [v.name, JSON.stringify(compactData)];
    });

    if (rows.length > 0) {
        cacheSheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
}

/**
 * Load the snapshot from the cache sheet.
 * Returns { vendorName: [{ label, values, mergedCellContent }] }
 */
function loadSnapshot_(ss) {
    const cacheSheet = ss.getSheetByName(PHIL_CONFIG.cacheSheetName);
    if (!cacheSheet || cacheSheet.getLastRow() === 0) return null;

    const data = cacheSheet.getRange(1, 1, cacheSheet.getLastRow(), 2).getValues();
    const snapshot = {};

    for (const row of data) {
        const name = row[0];
        try {
            snapshot[name] = JSON.parse(row[1]);
        } catch (e) {
            Logger.log("PhilSync: Failed to parse snapshot for " + name);
        }
    }

    return snapshot;
}

/**
 * Normalize a value for safe comparison in the 3-way merge.
 * Handles nulls, dates, booleans, and strings.
 */
function normalizeForCompare_(val) {
    if (val === null || val === undefined || val === "") return "";
    if (val instanceof Date) return val.getTime().toString();
    // Handle ISO date strings from JSON cache
    if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) {
        return new Date(val).getTime().toString();
    }
    if (typeof val === 'boolean') return val ? "TRUE" : "FALSE";
    return String(val).trim();
}
