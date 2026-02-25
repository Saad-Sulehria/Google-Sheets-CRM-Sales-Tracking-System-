/**
 * SheetTools.js — General-purpose vendor sheet management tools.
 *
 * Functions (all menu-accessible):
 *   rankVendors()          — Sort vendors by rank number in column E row 2
 *   moveVendorToSheet()    — Copy a vendor from Meeting Notes to another sheet
 *   removeVendor()         — Delete a vendor block from any sheet
 *   syncSheets()           — One-shot bi-directional sync with Meeting Notes
 */

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Get sheet or alert
// ─────────────────────────────────────────────────────────────────────────────
function getSheetOrAlert_(ss, name, ui) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) ui.alert("Sheet not found: \"" + name + "\"");
    return sheet;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Prompt for sheet name (with optional default)
// ─────────────────────────────────────────────────────────────────────────────
function promptSheetName_(ui, title, message) {
    const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return null;
    return response.getResponseText().trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Find vendor blocks in a sheet (name + startRow 1-indexed)
// ─────────────────────────────────────────────────────────────────────────────
function findVendorBlocks_(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return [];
    const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    const weights = sheet.getRange(1, 1, lastRow, 1).getFontWeights();
    const blocks = [];
    for (let i = 0; i < colA.length; i++) {
        const val = String(colA[i][0] || "").trim();
        if (val !== "" && weights[i][0] === "bold") {
            blocks.push({ name: val, startRow: i + 1 }); // 1-indexed
        }
    }
    return blocks;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Copy column widths and row heights from source to dest sheet
// ─────────────────────────────────────────────────────────────────────────────
function copyDimensions_(srcSheet, destSheet, colCount, blockHeight, vendorCount) {
    // Column widths
    for (let c = 1; c <= colCount; c++) {
        destSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
    }
    // Row heights — use first block as pattern, repeat for all vendors
    const srcHeights = [];
    for (let r = 1; r <= blockHeight; r++) {
        srcHeights.push(srcSheet.getRowHeight(r));
    }
    for (let v = 0; v < vendorCount; v++) {
        for (let r = 0; r < blockHeight; r++) {
            destSheet.setRowHeight(v * blockHeight + r + 1, srcHeights[r]);
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// 1. RANK VENDORS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Sort all vendors in a sheet by the rank number in column E (row 2 of each block).
 * Vendors with rank 1 come first, then 2, 3, 4, 5. Missing rank → placed at end.
 */
function rankVendors() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    const sheetName = promptSheetName_(ui,
        "Rank Vendors",
        "Enter the sheet name to rank (leave blank for \"" + CONFIG.notesSheetName + "\"):"
    );
    if (sheetName === null) return;

    const targetSheet = getSheetOrAlert_(ss, sheetName || CONFIG.notesSheetName, ui);
    if (!targetSheet) return;

    const templateSheet = getSheetOrAlert_(ss, CONFIG.templateSheetName, ui);
    if (!templateSheet) return;

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;

    // Extract vendor data
    const vendors = extractAllVendorData(targetSheet, templateHeight);
    if (vendors.length === 0) {
        ui.alert("No vendors found in \"" + (sheetName || CONFIG.notesSheetName) + "\".");
        return;
    }

    // Read rank from column E on the first data row of each block (row 2 of vendor = index 1)
    // Column E = column 5 (1-indexed) = raw data index 4
    const lastRow = targetSheet.getLastRow();
    const rawColE = targetSheet.getRange(1, 5, lastRow, 1).getValues();

    // Attach rank to each vendor
    const vendorsWithRank = vendors.map((v, idx) => {
        // Header row is at idx * blockHeight, data rows start at idx * blockHeight + 1
        const dataRow = idx * blockHeight + 1; // 0-indexed row in rawColE
        const rawRank = rawColE[dataRow] ? rawColE[dataRow][0] : "";
        const rank = parseInt(String(rawRank).trim(), 10);
        return { vendor: v, rank: isNaN(rank) ? Infinity : rank };
    });

    // Stable sort ascending by rank (missing = Infinity = end)
    vendorsWithRank.sort((a, b) => a.rank - b.rank);
    const sortedVendors = vendorsWithRank.map(x => x.vendor);

    // Rebuild the sheet with sorted order
    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);
    const requiredRows = sortedVendors.length * blockHeight;
    const currentMaxRows = targetSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        targetSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    rebuildVendorBlocks(targetSheet, templateSheet, templateRange, sortedVendors, fullHeaderWidth);

    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);
    injectDataByLabel(targetSheet, sortedVendors, labelMap, templateHeight, templateContentWidth);

    // Copy dimensions from source (first block pattern)
    const lastUsedCol = 1 + Math.max(templateContentWidth, templateCopyWidth);
    copyDimensions_(targetSheet, targetSheet, lastUsedCol, blockHeight, sortedVendors.length);

    const ranksUsed = vendorsWithRank.map(x => x.rank === Infinity ? "—" : x.rank).join(", ");
    ui.alert("✅ Vendors Ranked\n\nSheet: \"" + (sheetName || CONFIG.notesSheetName) + "\"\nOrder: " + ranksUsed);
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. MOVE VENDOR TO SHEET
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Copy a specific vendor from Meeting Notes to another sheet (appended at end).
 */
function moveVendorToSheet() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    const destSheetName = promptSheetName_(ui,
        "Move Vendor — Step 1 of 2",
        "Enter the DESTINATION sheet name:"
    );
    if (!destSheetName) { if (destSheetName === null) return; ui.alert("Sheet name is required."); return; }

    const vendorName = promptSheetName_(ui,
        "Move Vendor — Step 2 of 2",
        "Enter the vendor name to copy from \"" + CONFIG.notesSheetName + "\":"
    );
    if (!vendorName) { if (vendorName === null) return; ui.alert("Vendor name is required."); return; }

    const notesSheet = getSheetOrAlert_(ss, CONFIG.notesSheetName, ui);
    if (!notesSheet) return;
    const templateSheet = getSheetOrAlert_(ss, CONFIG.templateSheetName, ui);
    if (!templateSheet) return;

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;

    const allVendors = extractAllVendorData(notesSheet, templateHeight);
    const vendor = allVendors.find(v => v.name.toLowerCase() === vendorName.toLowerCase());

    if (!vendor) {
        const names = allVendors.map(v => v.name).slice(0, 20).join("\n• ");
        ui.alert("Vendor \"" + vendorName + "\" not found in Meeting Notes.\n\nExamples:\n• " + names);
        return;
    }

    // Get or create destination sheet
    let destSheet = ss.getSheetByName(destSheetName);
    if (!destSheet) destSheet = ss.insertSheet(destSheetName);

    // Find current last block in destination (to append after)
    const existingBlocks = findVendorBlocks_(destSheet);
    const numExisting = existingBlocks.length;
    const appendIndex = numExisting; // the new vendor will be at this 0-based index

    // Ensure enough rows in destination
    const requiredRows = (numExisting + 1) * blockHeight;
    const currentMaxRows = destSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        destSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    // Extract existing vendor data from dest sheet (to rebuild all + appended)
    const existingVendors = numExisting > 0 ? extractAllVendorData(destSheet, templateHeight) : [];
    const allDestVendors = [...existingVendors, vendor];

    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);

    // Rebuild: we only add the new block at the end to minimize disruption
    // Write new vendor block directly at the append position
    const newBlockStartRow = appendIndex * blockHeight + 1;
    const headerCell = destSheet.getRange(newBlockStartRow, 1);
    headerCell.setValue(vendor.name);
    headerCell.setFontWeight("bold");
    headerCell.setFontColor("#ffffff");
    headerCell.setBackground(CONFIG.vendorHeaderBgColor || "#ff00ff");

    // Fill rest of header row
    const headerFillWidth = Math.max(templateContentWidth, templateCopyWidth);
    if (headerFillWidth > 0) {
        destSheet.getRange(newBlockStartRow, 2, 1, headerFillWidth)
            .setBackground(CONFIG.vendorHeaderBgColor || "#ff00ff")
            .setValue("");
    }

    // Sanitize header row: clear checkboxes/validations that might ghost into data columns
    destSheet.getRange(newBlockStartRow, 2, 1, headerFillWidth).clearContent().clearDataValidations().removeCheckboxes();

    // Copy template into the vendor block — starts at column B (col 2), NOT col A
    const destBlockDataStart = newBlockStartRow + 1;
    const templateWidth = templateRange.getNumColumns();
    const targetRange = destSheet.getRange(destBlockDataStart, 2, templateHeight, templateWidth);
    templateRange.copyTo(targetRange, { contentsOnly: false });

    // Inject this vendor's data
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);
    // Use a single-vendor array with the correct startRow override
    // We'll do injection manually for a single block
    const blockRange = destSheet.getRange(destBlockDataStart, 2, templateHeight, templateContentWidth);
    const currentValues = blockRange.getValues();
    const newValues = currentValues.map(r => [...r]);
    const formulasToSet = Array.from({ length: templateHeight }, () => Array(templateContentWidth).fill(""));
    let hasFormulas = false;

    for (const item of vendor.data) {
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
            if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) val = new Date(val);
            newValues[relativeRow][targetCol] = val;
            const formula = valueFormulas[v];
            if (formula && String(formula).trim() !== "") {
                formulasToSet[relativeRow][targetCol] = formula;
                hasFormulas = true;
            }
        }

        if (item.mergedCellContent) {
            const mergedRow = relativeRow + 1;
            const mergedCol = templatePos.labelCol;
            if (mergedRow < newValues.length && mergedCol < newValues[mergedRow].length) {
                newValues[mergedRow][mergedCol] = item.mergedCellContent;
                const mf = item.mergedCellFormula;
                if (mf && String(mf).trim() !== "") { formulasToSet[mergedRow][mergedCol] = mf; hasFormulas = true; }
            }
        }
    }

    blockRange.setValues(newValues);
    if (hasFormulas) {
        for (let r = 0; r < templateHeight; r++) {
            for (let c = 0; c < templateContentWidth; c++) {
                const f = formulasToSet[r][c];
                if (f && String(f).trim() !== "") destSheet.getRange(destBlockDataStart + r, 2 + c).setFormula(f);
            }
        }
    }

    // Copy column widths from Meeting Notes
    const lastUsedCol = 1 + Math.max(templateContentWidth, templateCopyWidth);
    for (let c = 1; c <= lastUsedCol; c++) {
        destSheet.setColumnWidth(c, notesSheet.getColumnWidth(c));
    }
    // Copy row heights for the new block only
    for (let r = 0; r < blockHeight; r++) {
        destSheet.setRowHeight(newBlockStartRow + r, notesSheet.getRowHeight(r + 1));
    }

    ui.alert("✅ Vendor Copied\n\nVendor: \"" + vendor.name + "\"\nDestination: \"" + destSheetName + "\"");
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. REMOVE VENDOR
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Remove a specific vendor block (header + template + blank rows) from any sheet.
 */
function removeVendor() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    const sheetName = promptSheetName_(ui,
        "Remove Vendor — Step 1 of 2",
        "Enter the sheet name (leave blank for \"" + CONFIG.notesSheetName + "\"):"
    );
    if (sheetName === null) return;

    const vendorName = promptSheetName_(ui,
        "Remove Vendor — Step 2 of 2",
        "Enter the vendor name to remove:"
    );
    if (!vendorName) { if (vendorName === null) return; ui.alert("Vendor name is required."); return; }

    const targetSheet = getSheetOrAlert_(ss, sheetName || CONFIG.notesSheetName, ui);
    if (!targetSheet) return;
    const templateSheet = getSheetOrAlert_(ss, CONFIG.templateSheetName, ui);
    if (!templateSheet) return;

    const blocks = findVendorBlocks_(targetSheet);
    const target = blocks.find(b => b.name.toLowerCase() === vendorName.toLowerCase());

    if (!target) {
        const available = blocks.map(b => b.name).slice(0, 20).join("\n• ");
        ui.alert("Vendor \"" + vendorName + "\" not found.\n\nAvailable:\n• " + available);
        return;
    }

    const templateHeight = templateSheet.getLastRow();
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;

    targetSheet.deleteRows(target.startRow, blockHeight);

    ui.alert("✅ Vendor Removed\n\nVendor: \"" + target.name + "\"\nSheet: \"" + (sheetName || CONFIG.notesSheetName) + "\"");
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. SYNC SHEETS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * One-shot bi-directional sync between any sheet and Meeting Notes.
 *
 * Strategy (no persistent cache needed):
 *   - For each vendor that exists in BOTH sheets:
 *     - Field present in other sheet but empty in Meeting Notes → push to Meeting Notes
 *     - Field present in Meeting Notes → Meeting Notes wins → overwrite other sheet
 *   - Cache sheet is created temporarily and deleted after sync.
 */
function syncSheets() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    const sheetName = promptSheetName_(ui,
        "Sync Sheets",
        "Enter the sheet name to sync with \"" + CONFIG.notesSheetName + "\":"
    );
    if (!sheetName) { if (sheetName === null) return; ui.alert("Sheet name is required."); return; }

    const otherSheet = getSheetOrAlert_(ss, sheetName, ui);
    if (!otherSheet) return;
    const notesSheet = getSheetOrAlert_(ss, CONFIG.notesSheetName, ui);
    if (!notesSheet) return;
    const templateSheet = getSheetOrAlert_(ss, CONFIG.templateSheetName, ui);
    if (!templateSheet) return;

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);
    const blankRows = CONFIG.blankRowsAfterVendor || 0;

    // Extract data from both sheets
    const notesVendors = extractAllVendorData(notesSheet, templateHeight);
    const otherVendors = extractAllVendorData(otherSheet, templateHeight);

    const notesMap = {};
    for (const v of notesVendors) notesMap[v.name] = v;
    const otherMap = {};
    for (const v of otherVendors) otherMap[v.name] = v;

    // Vendors that exist in both sheets
    const commonNames = otherVendors.map(v => v.name).filter(n => notesMap[n]);

    if (commonNames.length === 0) {
        ui.alert("No common vendors found between \"" + sheetName + "\" and Meeting Notes.\n\nNo changes made.");
        return;
    }

    const vendorsToUpdateInNotes = [];
    const vendorsToUpdateInOther = [];

    for (const name of commonNames) {
        const notesVendor = notesMap[name];
        const otherVendor = otherMap[name];

        const mergedForNotes = JSON.parse(JSON.stringify(notesVendor));
        const mergedForOther = JSON.parse(JSON.stringify(otherVendor));

        let needsNotesUpdate = false;
        let needsOtherUpdate = false;

        for (const otherItem of otherVendor.data) {
            const label = otherItem.label;
            const notesItem = notesVendor.data.find(x => x.label === label);
            if (!notesItem) continue;

            const mergedNotesItem = mergedForNotes.data.find(x => x.label === label);
            const mergedOtherItem = mergedForOther.data.find(x => x.label === label);

            const otherVals = otherItem.values || [];
            const notesVals = notesItem.values || [];

            // Compare only within the shorter live array bounds (prevent phantom nulls)
            const compareLen = Math.min(
                Math.max(otherVals.length, notesVals.length),
                Math.max(otherVals.length, notesVals.length)
            );

            for (let v = 0; v < compareLen; v++) {
                const ov = normalizeForSyncCompare_(otherVals[v]);
                const nv = normalizeForSyncCompare_(notesVals[v]);

                if (ov === nv) continue;

                if (nv !== "" && ov === "") {
                    // Meeting Notes has value, other is empty → push Notes → Other
                    if (mergedOtherItem && mergedOtherItem.values) {
                        mergedOtherItem.values[v] = notesVals[v];
                        if (notesItem.valueFormulas && notesItem.valueFormulas[v]) {
                            mergedOtherItem.valueFormulas = mergedOtherItem.valueFormulas || [];
                            mergedOtherItem.valueFormulas[v] = notesItem.valueFormulas[v];
                        }
                    }
                    needsOtherUpdate = true;
                } else if (ov !== "" && nv === "") {
                    // Other has value, Notes is empty → push Other → Notes
                    if (mergedNotesItem && mergedNotesItem.values) {
                        mergedNotesItem.values[v] = otherVals[v];
                        if (otherItem.valueFormulas && otherItem.valueFormulas[v]) {
                            mergedNotesItem.valueFormulas = mergedNotesItem.valueFormulas || [];
                            mergedNotesItem.valueFormulas[v] = otherItem.valueFormulas[v];
                        }
                    }
                    needsNotesUpdate = true;
                } else {
                    // Both have different non-empty values → Meeting Notes wins
                    if (mergedOtherItem && mergedOtherItem.values) {
                        mergedOtherItem.values[v] = notesVals[v];
                    }
                    needsOtherUpdate = true;
                }
            }

            // Merged cell content
            const ov_m = normalizeForSyncCompare_(otherItem.mergedCellContent);
            const nv_m = normalizeForSyncCompare_(notesItem.mergedCellContent);
            if (ov_m !== nv_m) {
                if (nv_m !== "" && ov_m === "") {
                    if (mergedOtherItem) mergedOtherItem.mergedCellContent = notesItem.mergedCellContent;
                    needsOtherUpdate = true;
                } else if (ov_m !== "" && nv_m === "") {
                    if (mergedNotesItem) mergedNotesItem.mergedCellContent = otherItem.mergedCellContent;
                    needsNotesUpdate = true;
                } else {
                    if (mergedOtherItem) mergedOtherItem.mergedCellContent = notesItem.mergedCellContent;
                    needsOtherUpdate = true;
                }
            }
        }

        if (needsOtherUpdate) vendorsToUpdateInOther.push(mergedForOther);
        if (needsNotesUpdate) vendorsToUpdateInNotes.push(mergedForNotes);
    }

    // Apply changes
    if (vendorsToUpdateInOther.length > 0) {
        applySyncUpdates_(otherSheet, otherVendors, vendorsToUpdateInOther, labelMap, templateHeight, templateContentWidth, blankRows);
    }
    if (vendorsToUpdateInNotes.length > 0) {
        applySyncUpdates_(notesSheet, notesVendors, vendorsToUpdateInNotes, labelMap, templateHeight, templateContentWidth, blankRows);
    }

    const summary = "✅ Sync Complete\n\n" +
        "Sheet: \"" + sheetName + "\"\n" +
        "Common vendors synced: " + commonNames.length + "\n" +
        "Updated in \"" + sheetName + "\": " + vendorsToUpdateInOther.length + "\n" +
        "Updated in Meeting Notes: " + vendorsToUpdateInNotes.length;

    ui.alert(summary);
}

/**
 * Apply sync updates to specific vendor blocks in a sheet.
 * Only writes to vendors that changed. Preserves data validations.
 * Uses cell-by-cell formula writes to avoid phantom drift.
 */
function applySyncUpdates_(sheet, allVendorsInSheet, updatedVendors, labelMap, templateHeight, templateContentWidth, blankRows) {
    const blockHeight = 1 + templateHeight + blankRows;
    const updatedNames = new Set(updatedVendors.map(v => v.name));

    for (let i = 0; i < allVendorsInSheet.length; i++) {
        const vendor = allVendorsInSheet[i];
        if (!updatedNames.has(vendor.name)) continue;

        const updatedVendor = updatedVendors.find(v => v.name === vendor.name);
        if (!updatedVendor) continue;

        const blockStartRow = i * blockHeight + 1;
        const dataStartRow = blockStartRow + 1;

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
                if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) val = new Date(val);
                newValues[relativeRow][targetCol] = val;
                const formula = valueFormulas[v];
                if (formula && String(formula).trim() !== "") {
                    formulasToSet[relativeRow][targetCol] = formula;
                    hasFormulas = true;
                }
            }

            if (item.mergedCellContent) {
                const mergedRow = relativeRow + 1;
                const mergedCol = templatePos.labelCol;
                if (mergedRow < newValues.length && mergedCol < newValues[mergedRow].length) {
                    newValues[mergedRow][mergedCol] = item.mergedCellContent;
                    const mf = item.mergedCellFormula;
                    if (mf && String(mf).trim() !== "") { formulasToSet[mergedRow][mergedCol] = mf; hasFormulas = true; }
                }
            }
        }

        const savedValidations = blockRange.getDataValidations();
        blockRange.clearDataValidations();
        blockRange.setValues(newValues);
        if (hasFormulas) {
            for (let r = 0; r < templateHeight; r++) {
                for (let c = 0; c < templateContentWidth; c++) {
                    const f = formulasToSet[r][c];
                    if (f && String(f).trim() !== "") sheet.getRange(dataStartRow + r, 2 + c).setFormula(f);
                }
            }
        }
        blockRange.setDataValidations(savedValidations);
    }
}

/**
 * Normalize a value for sync comparison (strings, dates, booleans, nulls).
 */
function normalizeForSyncCompare_(val) {
    if (val === null || val === undefined || val === "") return "";
    if (val instanceof Date) return val.getTime().toString();
    if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) return new Date(val).getTime().toString();
    if (typeof val === 'boolean') return val ? "TRUE" : "FALSE";
    return String(val).trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. EXPAND NEW SHEET VENDORS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads a flat vendor list (one row per contact) from a source sheet
 * and creates structured vendor blocks in a destination sheet.
 * Handles multiple contacts per vendor by inserting extra rows.
 */
function expandNewSheetVendors() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    // 1. Prompt for source sheet
    const sourceSheetName = promptSheetName_(ui,
        "Expand Vendors — Step 1 of 2",
        "Enter the SOURCE sheet name (contains flat rows):"
    );
    if (!sourceSheetName) { if (sourceSheetName === null) return; ui.alert("Source sheet name is required."); return; }

    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) { ui.alert("Source sheet not found: \"" + sourceSheetName + "\""); return; }

    // 2. Prompt for destination sheet
    const destSheetName = promptSheetName_(ui,
        "Expand Vendors — Step 2 of 2",
        "Enter the DESTINATION sheet name (will be created if needed):"
    );
    if (!destSheetName) { if (destSheetName === null) return; ui.alert("Destination sheet name is required."); return; }

    // 3. Get templates
    const templateSheet = getSheetOrAlert_(ss, CONFIG.templateSheetName, ui);
    if (!templateSheet) return;

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = Math.max(templateContentWidth, getTemplateMaxColumn(templateSheet, templateHeight));
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);

    // Scan template to build a dynamic coordinate map for column-header style fields
    const posMap = {};
    const tData = templateRange.getValues();
    for (let r = 0; r < tData.length; r++) {
        for (let c = 0; c < tData[r].length; c++) {
            const val = String(tData[r][c] || "").trim().toLowerCase();
            if (val === "name" && !posMap.name) posMap.name = { r: r + 1, c: c };
            else if (val === "title" && !posMap.title) posMap.title = { r: r + 1, c: c };
            else if (val === "email" && !posMap.email) posMap.email = { r: r + 1, c: c };
            else if (val === "linkedin" && !posMap.linkedin) posMap.linkedin = { r: r + 1, c: c };
            else if (val === "geo" && !posMap.geo) posMap.geo = { r: r + 1, c: c };
            else if (val === "hc" && !posMap.hc) posMap.hc = { r: r + 1, c: c };
        }
    }

    // 4. Read Source Data
    const srcLastRow = sourceSheet.getLastRow();
    const srcLastCol = sourceSheet.getLastColumn();
    if (srcLastRow < 2) {
        ui.alert("Source sheet has no data rows (needs header in row 1 + data).");
        return;
    }

    const srcData = sourceSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();
    const headers = srcData[0].map(h => String(h).trim().toLowerCase());

    // Header indices mapping - using includes for maximum robustness against extra spaces/newlines
    const colIdx = {
        vendor: headers.findIndex(h => h.includes("vendor")),
        exclude: headers.findIndex(h => h.includes("exclude")),
        firstName: headers.findIndex(h => h.includes("first name")),
        lastName: headers.findIndex(h => h.includes("last name")),
        title: headers.findIndex(h => h.includes("title")),
        email: headers.findIndex(h => h === "email"), // Exact match to avoid "got email" or "email?"
        linkedin: headers.findIndex(h => h === "li" || h.includes("linkedin")), // 'LI' is tricky, strict or linkedin
        geo: headers.findIndex(h => h.includes("geo")),
        hc: headers.findIndex(h => h.includes("hc") || h.includes("headcount"))
    };

    if (colIdx.vendor === -1) {
        ui.alert("Could not find 'Vendor' column in source sheet header.");
        return;
    }

    if (colIdx.exclude === -1) {
        ui.alert("Notice: Could not find 'Exclude' column in the header. Vendors will not be skipped.");
    }

    // 5. Group by Vendor
    const vendorsMap = new Map(); // vendorName -> array of contact objects
    const vendorNamesInOrder = []; // To preserve source order
    const globallyExcludedVendors = new Set(); // Track vendors that have 'Y' on ANY row

    for (let r = 1; r < srcData.length; r++) {
        const row = srcData[r];
        const vendorName = String(row[colIdx.vendor] || "").trim();
        const excludeFlag = colIdx.exclude !== -1 ? String(row[colIdx.exclude] || "").trim().toUpperCase() : "";

        if (!vendorName) continue;

        // If ANY row for this vendor has a 'Y', the entire vendor is blacklisted
        if (excludeFlag.startsWith("Y")) {
            globallyExcludedVendors.add(vendorName);
            continue;
        }

        const contact = {
            firstName: colIdx.firstName !== -1 ? String(row[colIdx.firstName] || "").trim() : "",
            lastName: colIdx.lastName !== -1 ? String(row[colIdx.lastName] || "").trim() : "",
            title: colIdx.title !== -1 ? String(row[colIdx.title] || "").trim() : "",
            email: colIdx.email !== -1 ? String(row[colIdx.email] || "").trim() : "",
            linkedin: colIdx.linkedin !== -1 ? String(row[colIdx.linkedin] || "").trim() : "",
            geo: colIdx.geo !== -1 ? String(row[colIdx.geo] || "").trim() : "",
            hc: colIdx.hc !== -1 ? String(row[colIdx.hc] || "").trim() : ""
        };

        contact.fullName = [contact.firstName, contact.lastName].filter(Boolean).join(" ");

        if (!vendorsMap.has(vendorName)) {
            vendorsMap.set(vendorName, []);
            vendorNamesInOrder.push(vendorName);
        }

        const existingContacts = vendorsMap.get(vendorName);
        const isDuplicate = existingContacts.some(
            c => c.fullName.toLowerCase() === contact.fullName.toLowerCase()
        );

        if (!isDuplicate && contact.fullName) {
            existingContacts.push(contact);
        } else if (existingContacts.length === 0 && !contact.fullName) {
            // Push empty contact to preserve the vendor block if it's the only row
            existingContacts.push(contact);
        }
    }

    // Filter out any vendors that were flagged for exclusion on ANY row
    const finalVendorsList = vendorNamesInOrder.filter(v => !globallyExcludedVendors.has(v));

    if (finalVendorsList.length === 0) {
        ui.alert("No valid vendors found to process (all were excluded or empty).");
        return;
    }

    // 6. Setup Destination Sheet
    let destSheet = ss.getSheetByName(destSheetName);
    if (!destSheet) {
        destSheet = ss.insertSheet(destSheetName);
    } else {
        const response = ui.alert(
            "Destination sheet exists",
            "Sheet \"" + destSheetName + "\" already exists. This will append to the bottom. Continue?",
            ui.ButtonSet.YES_NO
        );
        if (response !== ui.Button.YES) return;
    }

    // Prepare block formatting info
    const existingBlocks = findVendorBlocks_(destSheet);
    const numExisting = existingBlocks.length;
    const blankRows = CONFIG.blankRowsAfterVendor || 0;

    // Assuming worst case space needed (base block + 1 row per extra contact)
    const requiredExtraRows = vendorNamesInOrder.length * (templateHeight + blankRows + 5) + 10;
    destSheet.insertRowsAfter(destSheet.getMaxRows(), requiredExtraRows);

    let currentDestRow = destSheet.getLastRow() + (destSheet.getLastRow() > 0 ? 1 : 1);
    const fullHeaderWidth = 1 + templateCopyWidth;

    let processedCount = 0;

    // 7. Write Blocks
    for (const vName of finalVendorsList) {
        const contacts = vendorsMap.get(vName);
        if (contacts.length === 0) continue;

        const mainContact = contacts[0];
        const extraContacts = contacts.slice(1);

        // --- Write Header ---
        const headerCell = destSheet.getRange(currentDestRow, 1);
        headerCell.setValue(vName);
        headerCell.setFontWeight("bold");
        headerCell.setFontColor("#ffffff");

        const headerRowRange = destSheet.getRange(currentDestRow, 1, 1, fullHeaderWidth);
        headerRowRange.setBackground(CONFIG.vendorHeaderBgColor || "#ff00ff");
        if (fullHeaderWidth > 1) {
            destSheet.getRange(currentDestRow, 2, 1, fullHeaderWidth - 1).clearContent().clearDataValidations().removeCheckboxes();
        }

        // --- Copy Base Template ---
        const blockStartRow = currentDestRow + 1;
        const targetRange = destSheet.getRange(blockStartRow, 2, templateHeight, templateCopyWidth);
        templateRange.copyTo(targetRange, { contentsOnly: false });

        // Calculate actual height of this block (might increase if we add contacts)
        let currentBlockHeight = templateHeight;

        // --- Inject First Contact & Company Info ---
        const fieldsToInject = [
            { key: "name", val: mainContact.fullName },
            { key: "title", val: mainContact.title },
            { key: "email", val: mainContact.email },
            { key: "linkedin", val: mainContact.linkedin },
            { key: "geo", val: mainContact.geo },
            { key: "hc", val: mainContact.hc }
        ];

        for (const field of fieldsToInject) {
            if (!field.val) continue;
            const pos = posMap[field.key];
            if (pos && pos.r < templateHeight) {
                // pos.r is the row index (0-based) for the data row.
                // Paste happens at column 2 (B), so absolute col = 2 + pos.c
                destSheet.getRange(blockStartRow + pos.r, 2 + pos.c).setValue(field.val);
            }
        }

        // --- Handle Multiple Contacts ---
        if (extraContacts.length > 0) {
            // Find where to add extra contacts (after the main Name row)
            const namePos = posMap.name;
            if (namePos) {
                const absoluteInsertRow = blockStartRow + namePos.r; // The first data row

                // For each extra contact, write data into the existing rows below it
                for (let i = 0; i < extraContacts.length; i++) {
                    const ec = extraContacts[i];
                    const ecRow = absoluteInsertRow + i + 1; // move down one row per extra contact

                    // If we exceed the template layout height (meaning no more blank rows in the template box),
                    // only then we insert a row to avoid overlapping with section below
                    if (namePos.r + i + 1 >= templateHeight) {
                        destSheet.insertRowAfter(ecRow - 1);
                        currentBlockHeight++;
                        // Copy formatting from the first contact's data row to the new row
                        destSheet.getRange(absoluteInsertRow, 2, 1, templateCopyWidth)
                            .copyTo(destSheet.getRange(ecRow, 2), { formatOnly: true });
                    }

                    // Write secondary contact details into the exact same columns
                    if (posMap.name) destSheet.getRange(ecRow, 2 + posMap.name.c).setValue(ec.fullName + " (Secondary)");
                    if (posMap.title) destSheet.getRange(ecRow, 2 + posMap.title.c).setValue(ec.title);
                    if (posMap.email) destSheet.getRange(ecRow, 2 + posMap.email.c).setValue(ec.email);
                    if (posMap.linkedin) destSheet.getRange(ecRow, 2 + posMap.linkedin.c).setValue(ec.linkedin);
                    if (posMap.geo) destSheet.getRange(ecRow, 2 + posMap.geo.c).setValue(ec.geo);
                }
            }
        }

        currentDestRow += 1 + currentBlockHeight + blankRows;
        processedCount++;
    }

    // Clean up extra rows at the bottom
    const finalMaxRows = destSheet.getMaxRows();
    const finalUsedRow = destSheet.getLastRow();
    if (finalMaxRows > finalUsedRow + 10) {
        destSheet.deleteRows(finalUsedRow + 5, finalMaxRows - finalUsedRow - 5);
    }

    // Copy column widths from template
    for (let c = 1; c <= fullHeaderWidth; c++) {
        destSheet.setColumnWidth(c + 1, templateSheet.getColumnWidth(c));
    }

    ui.alert("✅ Expansion Complete\n\nProcessed " + processedCount + " vendors into sheet: \"" + destSheetName + "\"");
}
