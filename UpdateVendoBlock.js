/**
 * Configuration for the CRM Update Engine
 */
const CONFIG = {
    notesSheetName: "Meeting Notes",
    templateSheetName: "Template_LeftBlock",
    // templateRange and templateHeight are now dynamic
    labelColumns: [1, 2],           // Labels in columns A and B
    valueStartColumn: 2,            // Values start in column B
    vendorHeaderColumn: 1,          // Column A (vendor name)
    mergedRanges: ["H26:J27", "H32:J36"], // Can be detected dynamically if needed, but keeping primarily for known fields
    maxColumns: 11,                 // Columns A through K
    headerColor: "#1f4e79",         // Dark Blue for Vendor Headers (used for links)
    headerFontColor: "#ffffff",     // White text for Vendor Headers
    // Vendor block layout settings
    blankRowsAfterVendor: 1,        // Number of blank rows after each vendor block
    vendorHeaderBgColor: "#ff00ff", // Magenta background for vendor header row (full width)
    updateBatchSize: 10,            // Vendors per batch when using batched update
    preserveRichTextLinks: false    // Set true to preserve rich-text links (slower)
};

/**
 * Get the maximum column in the template, including merged cell extent
 * This is needed because getLastColumn() and getDataRange() don't include
 * merged cells that extend beyond the last content column
 * @param {Sheet} sheet - The template sheet
 * @param {number} lastRow - The last row to check
 * @returns {number} The maximum column number
 */
function getTemplateMaxColumn(sheet, lastRow) {
    // Start with the last column that has content
    let maxCol = sheet.getLastColumn();

    // Get all merged ranges in the sheet
    // Use bounded range (50 cols max) instead of getMaxColumns() for efficiency
    const maxColsToCheck = Math.min(50, sheet.getMaxColumns());
    const mergedRanges = sheet.getRange(1, 1, lastRow, maxColsToCheck).getMergedRanges();

    // Check each merged range to find the rightmost column
    for (let i = 0; i < mergedRanges.length; i++) {
        const range = mergedRanges[i];
        const lastMergedCol = range.getLastColumn();
        if (lastMergedCol > maxCol) {
            maxCol = lastMergedCol;
            Logger.log("Found merged range extending to column " + lastMergedCol + ": " + range.getA1Notation());
        }
    }

    Logger.log("Template max column (including merges): " + maxCol);
    return maxCol;
}

/**
 * Main function to update all vendor blocks with data preservation
 */
function updateAllVendorBlocks() {
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        SpreadsheetApp.getUi().alert("Missing sheets: " + CONFIG.notesSheetName + " or " + CONFIG.templateSheetName);
        return;
    }

    // Dynamic Template Dimensions Calculation
    // We need TWO widths:
    // 1. templateCopyWidth: Includes merged cells (for correct template copying)
    // 2. templateContentWidth: Only actual content columns (for injection bounds)
    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();  // Actual content/labels
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);  // Including merged cells
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);

    Logger.log("Template Dimensions: " + templateHeight + " rows, Content width: " + templateContentWidth + " cols, Copy width: " + templateCopyWidth + " cols");

    // Phase 1: Extract all vendor data before any destructive operations
    const vendorData = extractAllVendorData(notesSheet, templateHeight);
    const vendorCount = vendorData.length;

    if (vendorCount === 0) {
        SpreadsheetApp.getUi().alert("No vendors found in " + CONFIG.notesSheetName + ". Nothing to update.");
        return;
    }

    // Phase 2: Show confirmation dialog
    const confirmed = showConfirmationDialog(vendorCount, templateHeight);
    if (!confirmed) {
        return;
    }

    // Phase 3: Create backup before destructive operations
    const backupName = backupSpreadsheet();
    Logger.log("Backup created: " + backupName);

    // Phase 4: Build label map from new template
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);

    // Phase 5: Wipe and rebuild blocks
    // Calculate fullHeaderWidth: Column A (vendor name) + max of content width and copy width
    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);
    rebuildVendorBlocks(notesSheet, templateSheet, templateRange, vendorData, fullHeaderWidth);

    // Phase 6: Inject data back into new blocks
    // Use templateContentWidth to limit injection to actual content columns (not merged cell extent)
    const warnings = injectDataByLabel(notesSheet, vendorData, labelMap, templateHeight, templateContentWidth);

    // Phase 6.5: Clear trailing rows after last vendor block
    // Prevents leftover checkboxes/content when template shrinks (rows)
    // Use getMaxRows() instead of getLastRow() - checkboxes without content aren't counted by getLastRow
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const lastVendorEndRow = vendorCount * (1 + templateHeight + blankRows); // Each vendor: 1 header + template + blank rows
    const sheetMaxRow = notesSheet.getMaxRows();

    if (sheetMaxRow > lastVendorEndRow) {
        const rowsToClear = sheetMaxRow - lastVendorEndRow;
        Logger.log("Clearing " + rowsToClear + " trailing rows after row " + lastVendorEndRow);

        // Clear content, formatting, and data validations (checkboxes)
        const trailingRowsRange = notesSheet.getRange(lastVendorEndRow + 1, 1, rowsToClear, notesSheet.getMaxColumns());
        trailingRowsRange.clearContent();
        trailingRowsRange.clearDataValidations();
    }

    // Phase 6.55: Clear trailing columns after template CONTENT width
    // Prevents leftover data when template shrinks (columns removed)
    // Uses templateContentWidth (actual content columns) rather than templateCopyWidth (merged extent)
    const sheetMaxCols = notesSheet.getMaxColumns();
    const lastUsedCol = 1 + templateContentWidth; // Column A (vendor name) + template content width (B onwards)

    if (sheetMaxCols > lastUsedCol) {
        const colsToClear = sheetMaxCols - lastUsedCol;
        Logger.log("Clearing " + colsToClear + " trailing columns after column " + lastUsedCol);

        // Clear all rows in the trailing columns
        const trailingColsRange = notesSheet.getRange(1, lastUsedCol + 1, sheetMaxRow, colsToClear);
        trailingColsRange.clearContent();
        trailingColsRange.clearDataValidations();
    }

    // Phase 7: Show summary with reminder to run Reconnect Notes Links
    showSummary(vendorCount, warnings, backupName);
}

/**
 * Batched update to avoid execution time limits
 */
function startUpdateAllVendorBlocksBatched() {
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        SpreadsheetApp.getUi().alert("Missing sheets: " + CONFIG.notesSheetName + " or " + CONFIG.templateSheetName);
        return;
    }

    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);

    const vendorData = extractAllVendorData(notesSheet, templateHeight);
    const vendorCount = vendorData.length;

    if (vendorCount === 0) {
        SpreadsheetApp.getUi().alert("No vendors found in " + CONFIG.notesSheetName + ". Nothing to update.");
        return;
    }

    const confirmed = showConfirmationDialog(vendorCount, templateHeight);
    if (!confirmed) {
        return;
    }

    const backupName = backupSpreadsheet();
    Logger.log("Backup created: " + backupName);

    // Cache vendor data to a hidden sheet to avoid large in-memory state
    const cacheSheet = getOrCreateUpdateCacheSheet_(ss);
    cacheSheet.clear();
    cacheSheet.getRange(1, 1, 1, 2).setValues([["Vendor", "Data"]]);
    const cacheRows = vendorData.map(vendor => [vendor.name, JSON.stringify(vendor)]);
    if (cacheRows.length > 0) {
        cacheSheet.getRange(2, 1, cacheRows.length, 2).setValues(cacheRows);
    }
    cacheSheet.hideSheet();

    // Clear Meeting Notes and ensure enough rows for rebuild
    notesSheet.clear();
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;
    const requiredRows = vendorCount * blockHeight;
    const maxRows = notesSheet.getMaxRows();
    if (maxRows < requiredRows) {
        notesSheet.insertRowsAfter(maxRows, requiredRows - maxRows);
    }

    const state = {
        vendorCount: vendorCount,
        templateHeight: templateHeight,
        templateContentWidth: templateContentWidth,
        templateCopyWidth: templateCopyWidth,
        fullHeaderWidth: 1 + Math.max(templateContentWidth, templateCopyWidth),
        blankRows: blankRows,
        batchSize: CONFIG.updateBatchSize || 10,
        startIndex: 0
    };

    PropertiesService.getDocumentProperties().setProperty("UPDATE_ALL_BATCH_STATE", JSON.stringify(state));
    removeTriggers_("processUpdateAllVendorBlocksBatch");

    // Run first batch immediately
    processUpdateAllVendorBlocksBatch(true);
}

/**
 * Processes a batch of vendors for the batched update flow.
 * @param {boolean} isManual - True when called directly by the user.
 */
function processUpdateAllVendorBlocksBatch(isManual) {
    const ss = SpreadsheetApp.getActive();
    const props = PropertiesService.getDocumentProperties();
    const stateJson = props.getProperty("UPDATE_ALL_BATCH_STATE");
    if (!stateJson) {
        return;
    }

    const state = JSON.parse(stateJson);
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);
    const cacheSheet = ss.getSheetByName("_UpdateCache");

    if (!notesSheet || !templateSheet || !cacheSheet) {
        props.deleteProperty("UPDATE_ALL_BATCH_STATE");
        return;
    }

    const currentTemplateHeight = templateSheet.getLastRow();
    const currentTemplateContentWidth = templateSheet.getLastColumn();
    if (currentTemplateHeight !== state.templateHeight || currentTemplateContentWidth !== state.templateContentWidth) {
        props.deleteProperty("UPDATE_ALL_BATCH_STATE");
        return;
    }

    const templateRange = templateSheet.getRange(1, 1, state.templateHeight, state.templateCopyWidth);
    const labelMapRange = templateSheet.getRange(1, 1, state.templateHeight, state.templateContentWidth);
    const labelMap = buildTemplateLabelMap(templateSheet, labelMapRange);

    const startIndex = state.startIndex;
    const endIndex = Math.min(startIndex + state.batchSize, state.vendorCount);
    if (startIndex >= endIndex) {
        finalizeBatchedUpdate_(ss, cacheSheet, props, isManual);
        return;
    }

    const rowsToRead = endIndex - startIndex;
    const cacheRows = cacheSheet.getRange(2 + startIndex, 1, rowsToRead, 2).getValues();

    const blockHeight = 1 + state.templateHeight + state.blankRows;

    for (let i = 0; i < cacheRows.length; i++) {
        const rowIndex = startIndex + i;
        const payload = cacheRows[i][1];
        if (!payload) {
            continue;
        }

        const vendor = JSON.parse(payload);
        const startRow = 1 + rowIndex * blockHeight;
        writeVendorBlock_(notesSheet, templateRange, labelMap, state, vendor, startRow);
    }

    state.startIndex = endIndex;
    props.setProperty("UPDATE_ALL_BATCH_STATE", JSON.stringify(state));

    if (endIndex < state.vendorCount) {
        scheduleNextBatch_();
    } else {
        finalizeBatchedUpdate_(ss, cacheSheet, props, isManual);
    }
}

function finalizeBatchedUpdate_(ss, cacheSheet, props, isManual) {
    const state = JSON.parse(props.getProperty("UPDATE_ALL_BATCH_STATE") || "{}");
    props.deleteProperty("UPDATE_ALL_BATCH_STATE");
    removeTriggers_("processUpdateAllVendorBlocksBatch");

    try {
        ss.deleteSheet(cacheSheet);
    } catch (e) {
        Logger.log("Failed to delete cache sheet: " + e.message);
    }

    // Phase 6.5 + 6.55: Clear trailing rows and columns (same as non-batched path)
    try {
        const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
        if (notesSheet && state.vendorCount && state.templateHeight != null) {
            const blankRows = state.blankRows || 0;
            const lastVendorEndRow = state.vendorCount * (1 + state.templateHeight + blankRows);
            const sheetMaxRow = notesSheet.getMaxRows();

            if (sheetMaxRow > lastVendorEndRow) {
                const rowsToClear = sheetMaxRow - lastVendorEndRow;
                Logger.log("Clearing " + rowsToClear + " trailing rows after row " + lastVendorEndRow);
                const trailingRowsRange = notesSheet.getRange(lastVendorEndRow + 1, 1, rowsToClear, notesSheet.getMaxColumns());
                trailingRowsRange.clearContent();
                trailingRowsRange.clearDataValidations();
            }

            const sheetMaxCols = notesSheet.getMaxColumns();
            const lastUsedCol = 1 + (state.templateContentWidth || state.templateCopyWidth || 0);
            if (sheetMaxCols > lastUsedCol) {
                const colsToClear = sheetMaxCols - lastUsedCol;
                Logger.log("Clearing " + colsToClear + " trailing columns after column " + lastUsedCol);
                const trailingColsRange = notesSheet.getRange(1, lastUsedCol + 1, sheetMaxRow, colsToClear);
                trailingColsRange.clearContent();
                trailingColsRange.clearDataValidations();
            }
        }
    } catch (e) {
        Logger.log("Trailing cleanup error (non-fatal): " + e.message);
    }

    // Auto-reconnect Notes Links (safe for headless triggers - no getUi() inside)
    try {
        reconnectNotesLinksInternal(ss);
        Logger.log("Notes Links reconnected after batch update.");
    } catch (e) {
        Logger.log("Notes Links reconnect error (non-fatal): " + e.message);
    }

    // Store completion flag — shown to user on next sheet open (onOpen has a valid UI context)
    try {
        PropertiesService.getDocumentProperties().setProperty(
            "BATCH_UPDATE_COMPLETE",
            JSON.stringify({ vendorCount: state.vendorCount, timestamp: new Date().toISOString() })
        );
    } catch (e) {
        Logger.log("Could not store completion flag: " + e.message);
    }
}

function scheduleNextBatch_() {
    removeTriggers_("processUpdateAllVendorBlocksBatch");
    ScriptApp.newTrigger("processUpdateAllVendorBlocksBatch")
        .timeBased()
        .after(60 * 1000)
        .create();
}

function removeTriggers_(handlerName) {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === handlerName) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}

function getOrCreateUpdateCacheSheet_(ss) {
    const existing = ss.getSheetByName("_UpdateCache");
    if (existing) {
        return existing;
    }
    return ss.insertSheet("_UpdateCache");
}

function writeVendorBlock_(notesSheet, templateRange, labelMap, state, vendor, startRow) {
    const headerCell = notesSheet.getRange(startRow, 1);
    headerCell.setValue(vendor.name);
    headerCell.setFontWeight("bold");
    headerCell.setFontColor("#ffffff");

    const headerRowRange = notesSheet.getRange(startRow, 1, 1, state.fullHeaderWidth);
    headerRowRange.setBackground(CONFIG.vendorHeaderBgColor);

    if (state.fullHeaderWidth > 1) {
        notesSheet.getRange(startRow, 2, 1, state.fullHeaderWidth - 1).clearContent().clearDataValidations().removeCheckboxes();
    }

    const targetRange = notesSheet.getRange(startRow + 1, 2, state.templateHeight, state.templateCopyWidth);
    templateRange.copyTo(targetRange, { contentsOnly: false });

    injectVendorData_(notesSheet, vendor, labelMap, startRow + 1, state.templateHeight, state.templateContentWidth);
}

function injectVendorData_(sheet, vendor, labelMap, blockStartRow, templateHeight, templateWidth) {
    const blockRange = sheet.getRange(blockStartRow, 2, templateHeight, templateWidth);
    const currentValues = blockRange.getValues();
    const newValuesToSet = currentValues.map(r => [...r]);
    const formulasToSet = Array.from({ length: templateHeight }, () => Array(templateWidth).fill(""));
    let hasFormulas = false;

    for (let i = 0; i < vendor.data.length; i++) {
        const item = vendor.data[i];
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
            if (targetCol >= newValuesToSet[relativeRow].length) continue;

            // Reconvert ISO date strings back to Date objects (lost during JSON cache serialization)
            if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}T[\d:.]+Z$/.test(val)) {
                val = new Date(val);
            }

            newValuesToSet[relativeRow][targetCol] = val;

            const formula = valueFormulas[v];
            if (formula && String(formula).trim() !== "") {
                formulasToSet[relativeRow][targetCol] = formula;
                hasFormulas = true;
            }
        }

        if (item.mergedCellContent) {
            const mergedRow = relativeRow + 1;
            const mergedCol = templatePos.labelCol;
            if (mergedRow < newValuesToSet.length && mergedCol < newValuesToSet[mergedRow].length) {
                newValuesToSet[mergedRow][mergedCol] = item.mergedCellContent;

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

    blockRange.setValues(newValuesToSet);
    if (hasFormulas) {
        blockRange.setFormulas(formulasToSet);
    }
}

/**
 * Convert a RichTextValue hyperlink to a HYPERLINK formula, when possible
 * @param {RichTextValue} richText - Rich text value to inspect
 * @returns {string|null} HYPERLINK formula or null when not supported
 */
function richTextToHyperlinkFormula(richText) {
    if (!richText) return null;

    const directLink = richText.getLinkUrl && richText.getLinkUrl();
    if (directLink) {
        return '=HYPERLINK("' + directLink.replace(/"/g, '""') + '","' + richText.getText().replace(/"/g, '""') + '")';
    }

    const runs = richText.getRuns ? richText.getRuns() : null;
    if (!runs || runs.length === 0) return null;

    let linkUrl = null;
    for (let i = 0; i < runs.length; i++) {
        const runLink = runs[i].getLinkUrl && runs[i].getLinkUrl();
        if (runLink) {
            if (!linkUrl) {
                linkUrl = runLink;
            } else if (linkUrl !== runLink) {
                return null;
            }
        }
    }

    if (!linkUrl) return null;
    return '=HYPERLINK("' + linkUrl.replace(/"/g, '""') + '","' + richText.getText().replace(/"/g, '""') + '")';
}
/**
 * Extract all vendor data from Meeting Notes sheet
 * V10.1: Pure Positional Extraction with Merged Cell Support
 * 
 * Strategy:
 * 1. Find each labeled row (first non-empty cell from column B)
 * 2. Extract ALL values to the right of the label
 * 3. For special labels (Notes Field, Follow Up Action), look at next rows
 *    for merged cell content when current row values are empty
 * 
 * This approach is simpler and more reliable than header-keyed approaches.
 */
function extractAllVendorData(sheet, templateHeight) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0) return [];

    const fullRange = sheet.getRange(1, 1, lastRow, Math.max(lastCol, CONFIG.maxColumns + 10));
    const allData = fullRange.getValues();
    const allWeights = fullRange.getFontWeights();
    const allFormulas = fullRange.getFormulas();
    const allRichText = CONFIG.preserveRichTextLinks ? fullRange.getRichTextValues() : null;

    const vendors = [];
    let i = 0;

    // Labels that have content in merged cells on subsequent rows
    const MERGED_CONTENT_LABELS = ['Notes Field', 'Follow Up Action'];

    while (i < allData.length) {
        const cellA = allData[i][0];
        const weightA = allWeights[i][0];

        // Vendor Detection: Bold text in Column A
        if (cellA && String(cellA).trim() !== "" && weightA === 'bold') {
            const vendorName = String(cellA).trim();
            const vendorStart = i;

            // Find end of this vendor block
            let vendorEnd = i + 1;
            while (vendorEnd < allData.length) {
                const nextCellA = allData[vendorEnd][0];
                const nextWeightA = allWeights[vendorEnd][0];
                if (nextCellA && String(nextCellA).trim() !== "" && nextWeightA === 'bold') {
                    break;
                }
                vendorEnd++;
            }

            // Extract data from this vendor block
            const blockData = [];
            const processedRows = new Set(); // Track rows we've already processed

            for (let r = vendorStart; r < vendorEnd; r++) {
                // Skip already processed rows (e.g., merged content rows)
                if (processedRows.has(r)) continue;

                const rowData = allData[r];

                // Skip vendor header row
                if (r === vendorStart) continue;

                // Find label: first non-empty cell from column B onwards
                let labelCol = -1;
                let label = null;
                for (let c = 1; c < rowData.length; c++) {
                    const val = rowData[c];
                    if (val !== null && val !== "" && val !== undefined) {
                        label = String(val).trim();
                        labelCol = c - 1; // Template-relative (0 = template column A)
                        break;
                    }
                }

                if (!label || labelCol === -1) continue;

                // Extract ALL values after the label (pure positional)
                const sheetLabelCol = labelCol + 1; // Convert back to sheet column index
                const values = rowData.slice(sheetLabelCol + 1); // Everything after the label

                // Map values, preserving types (booleans stay booleans)
                let positionalValues = values.map(v => {
                    if (v === "") return null;
                    return v;
                });

                // Capture formulas and rich text (for hyperlinks) in positional fields
                const positionalFormulas = [];
                for (let v = 0; v < values.length; v++) {
                    const colIndex = sheetLabelCol + 1 + v;
                    const formula = allFormulas[r][colIndex];
                    if (formula && String(formula).trim() !== "") {
                        positionalFormulas[v] = formula;
                        continue;
                    }

                    if (allRichText) {
                        const richText = allRichText[r][colIndex];
                        const rtFormula = richTextToHyperlinkFormula(richText);
                        positionalFormulas[v] = rtFormula || null;
                    } else {
                        positionalFormulas[v] = null;
                    }
                }

                // For merged cell content, we'll store it separately
                let mergedCellContent = null;
                let mergedCellFormula = null;
                let mergedCellRichText = null;

                // Special handling for merged cell content labels
                // If this is a Notes Field or Follow Up Action label, check subsequent rows
                if (MERGED_CONTENT_LABELS.includes(label)) {
                    const hasContent = positionalValues.some(v => v !== null);

                    if (!hasContent) {
                        // Look at next rows for merged cell content
                        // The content is in a merged cell BELOW the label, at the SAME column
                        for (let nextR = r + 1; nextR < vendorEnd && nextR < r + 6; nextR++) {
                            const nextRowData = allData[nextR];

                            // Check if this row has content at the label's column position
                            // (Merged cell content appears in the top-left cell of the merge)
                            const contentAtLabelCol = nextRowData[sheetLabelCol];

                            if (contentAtLabelCol && String(contentAtLabelCol).trim() !== "") {
                                // Found merged cell content - store separately
                                mergedCellContent = contentAtLabelCol;
                                const mergedFormula = allFormulas[nextR][sheetLabelCol];
                                if (mergedFormula && String(mergedFormula).trim() !== "") {
                                    mergedCellFormula = mergedFormula;
                                } else {
                                    if (allRichText) {
                                        const mergedRichText = allRichText[nextR][sheetLabelCol];
                                        mergedCellRichText = richTextToHyperlinkFormula(mergedRichText);
                                    }
                                }
                                processedRows.add(nextR);
                                break;
                            }
                        }
                    }
                }

                blockData.push({
                    label: label,
                    labelCol: labelCol,
                    values: positionalValues,
                    valueFormulas: positionalFormulas,
                    mergedCellContent: mergedCellContent,  // Content for merged cell below label
                    mergedCellFormula: mergedCellFormula,
                    mergedCellRichText: mergedCellRichText
                });
            }

            vendors.push({ name: vendorName, data: blockData });
            i = vendorEnd;
        } else {
            i++;
        }
    }

    Logger.log("V10.1 Extracted " + vendors.length + " vendors with merged cell support");
    return vendors;
}

/**
 * Build a map of label positions from the template
 * V10: Simplified - just builds label → row/column mapping
 */
function buildTemplateLabelMap(templateSheet, templateRange) {
    const values = templateRange.getValues();
    const labelMap = {};

    for (let r = 0; r < values.length; r++) {
        const row = values[r];

        // Find label: first non-empty cell in row
        let label = null;
        let labelCol = -1;

        for (let c = 0; c < row.length; c++) {
            const val = row[c];
            if (val !== null && val !== "" && val !== undefined) {
                label = String(val).trim();
                labelCol = c;
                break;
            }
        }

        if (!label || labelCol === -1) continue;

        labelMap[label] = {
            row: r,
            labelCol: labelCol,
            valueCol: labelCol + 1
        };
    }

    Logger.log("V10 Built label map with " + Object.keys(labelMap).length + " labels");
    return labelMap;
}

/**
 * Safely set a value, handling merged cells
 * @param {Sheet} sheet - The target sheet
 * @param {number} row - 1-indexed row
 * @param {number} col - 1-indexed column
 * @param {any} value - The value to set
 */
function safeSetValue(sheet, row, col, value) {
    if (value === null || value === undefined) return;

    try {
        const range = sheet.getRange(row, col);

        // Check if this cell is part of a merged range
        if (range.isPartOfMerge()) {
            const mergedRange = range.getMergedRanges()[0];
            if (mergedRange) {
                // Write to top-left cell of the merge
                mergedRange.getCell(1, 1).setValue(value);
            } else {
                range.setValue(value);
            }
        } else {
            range.setValue(value);
        }
    } catch (e) {
        Logger.log("Error setting value at row " + row + ", col " + col + ": " + e.message);
    }
}

/**
 * Create a backup copy of the spreadsheet
 * @returns {string} The name of the backup file
 */
function backupSpreadsheet() {
    const ss = SpreadsheetApp.getActive();
    const name = ss.getName();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const backupName = "Backup - " + name + " - " + timestamp;

    const file = DriveApp.getFileById(ss.getId());
    file.makeCopy(backupName);

    return backupName;
}

/**
 * Show confirmation dialog before destructive operations
 * @param {number} vendorCount - Number of vendors found
 * @param {number} templateHeight - Rows per template block
 * @returns {boolean} True if user confirmed
 */
function showConfirmationDialog(vendorCount, templateHeight) {
    const ui = SpreadsheetApp.getUi();
    const message =
        "⚠️ GLOBAL UPDATE\n\n" +
        "This will rebuild ALL vendor blocks.\n\n" +
        "• Vendors found: " + vendorCount + "\n" +
        "• New template height: " + templateHeight + " rows\n" +
        "• A backup will be created first.\n\n" +
        "Data will be preserved by label matching.\n\n" +
        "Continue?";

    const response = ui.alert("Confirm Global Update", message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
}

/**
 * Clear and rebuild vendor blocks with new template
 * @param {Sheet} notesSheet - The Meeting Notes sheet
 * @param {Sheet} templateSheet - The Template sheet
 * @param {Range} templateRange - The template range to copy
 * @param {Array} vendorData - Array of vendor objects
 */
function rebuildVendorBlocks(notesSheet, templateSheet, templateRange, vendorData, fullHeaderWidth) {
    const templateHeight = templateRange.getNumRows();
    const templateWidth = templateRange.getNumColumns();
    const lastRow = notesSheet.getLastRow();
    const blankRows = CONFIG.blankRowsAfterVendor || 0;

    // Clear sheet completely except header row if exists? No, global update rebuilds everything.
    if (lastRow > 0) {
        notesSheet.clear();
    }

    let currentRow = 1;

    for (let i = 0; i < vendorData.length; i++) {
        const vendor = vendorData[i];

        // 1. Write Vendor Name in Column A
        const headerCell = notesSheet.getRange(currentRow, 1);
        headerCell.setValue(vendor.name);
        headerCell.setFontWeight("bold");
        headerCell.setFontColor("#ffffff");  // White text on magenta

        // 2. Color the ENTIRE header row (#ff00ff) for the full content width
        // This includes Column A (vendor name) + all template columns
        const headerRowRange = notesSheet.getRange(currentRow, 1, 1, fullHeaderWidth);
        headerRowRange.setBackground(CONFIG.vendorHeaderBgColor);

        // Sanitize the header row: Remove checkboxes and validations that might have ghosted
        notesSheet.getRange(currentRow, 2, 1, fullHeaderWidth - 1).clearContent().clearDataValidations().removeCheckboxes();

        // 3. Copy Template starting at B(currentRow + 1)
        const targetRange = notesSheet.getRange(currentRow + 1, 2, templateHeight, templateWidth);
        templateRange.copyTo(targetRange, { contentsOnly: false });

        // 4. Add blank rows after the vendor block (if configured)
        // Block layout: Vendor Header (1) + Template (N) + Blank Rows (B)
        // Next vendor starts at: currentRow + 1 + templateHeight + blankRows

        currentRow += 1 + templateHeight + blankRows;
    }

    Logger.log("Rebuilt " + vendorData.length + " vendor blocks with " + blankRows + " blank row(s) between each");
}

/**
 * Inject extracted data back into rebuilt blocks by label matching
 * V10: Pure Positional Injection (Simplified)
 * 
 * Strategy:
 * 1. For each vendor block, find the template row for each label
 * 2. Write ALL values at the same relative positions
 * 3. Preserve value types (booleans stay booleans)
 */
function injectDataByLabel(sheet, vendorData, labelMap, templateHeight, templateWidth) {
    const warnings = [];
    let currentRow = 1;

    // Use the actual template width passed in, not a fixed value
    // This ensures we only inject within the template bounds
    const injectionWidth = templateWidth;

    for (let i = 0; i < vendorData.length; i++) {
        const vendor = vendorData[i];
        const blockStartRow = currentRow + 1;

        // Read current block values from Column B (template width only)
        const blockRange = sheet.getRange(blockStartRow, 2, templateHeight, injectionWidth);
        const currentValues = blockRange.getValues();
        const newValuesToSet = currentValues.map(r => [...r]);
        const formulasToSet = Array.from({ length: templateHeight }, () => Array(injectionWidth).fill(""));
        let hasFormulas = false;

        for (const item of vendor.data) {
            const label = item.label;
            const templatePos = labelMap[label];

            if (!templatePos) {
                warnings.push("Label not in template: " + label);
                continue;
            }

            const relativeRow = templatePos.row;

            if (relativeRow < 0 || relativeRow >= templateHeight) {
                warnings.push("Row out of bounds: " + label);
                continue;
            }

            const valueStartCol = templatePos.valueCol;
            const values = item.values || [];
            const valueFormulas = item.valueFormulas || [];

            // Write ALL values at same relative positions
            for (let v = 0; v < values.length; v++) {
                const val = values[v];
                const targetCol = valueStartCol + v;

                // Skip null values (preserve template defaults)
                if (val === null) continue;

                // Bounds check
                if (targetCol >= newValuesToSet[relativeRow].length) continue;

                // Write the value (preserving type - booleans stay booleans)
                newValuesToSet[relativeRow][targetCol] = val;

                // Restore formulas or rich text hyperlinks when present
                const formula = valueFormulas[v];
                if (formula && String(formula).trim() !== "") {
                    formulasToSet[relativeRow][targetCol] = formula;
                    hasFormulas = true;
                }
            }

            // Handle merged cell content (Notes Field, Follow Up Action)
            // This content goes to the NEXT ROW at the LABEL column (not valueCol)
            if (item.mergedCellContent) {
                const mergedRow = relativeRow + 1;  // Next row
                const mergedCol = templatePos.labelCol;  // Same column as label

                if (mergedRow < newValuesToSet.length && mergedCol < newValuesToSet[mergedRow].length) {
                    newValuesToSet[mergedRow][mergedCol] = item.mergedCellContent;

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

        // Write the modified block back
        blockRange.setValues(newValuesToSet);
        if (hasFormulas) {
            blockRange.setFormulas(formulasToSet);
        }
        currentRow += 1 + templateHeight + (CONFIG.blankRowsAfterVendor || 0);
    }

    if (warnings.length > 0) {
        Logger.log("V10 Injection warnings (" + warnings.length + "): " + warnings.slice(0, 5).join("; "));
    }

    return warnings;
}

/**
 * Create a filtered view of vendor blocks in a separate sheet.
 * Prompts for destination sheet name and up to 2 label/value filter criteria.
 * Only vendors matching ALL filters are written to the destination sheet.
 * Formatting, template structure, and blank rows are preserved identically.
 */
function createFilteredView() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        ui.alert("Missing required sheets: " + CONFIG.notesSheetName + " or " + CONFIG.templateSheetName);
        return;
    }

    // Step 1: Ask for destination sheet name
    const destResponse = ui.prompt(
        "Create Filtered View (1/4)",
        "Enter destination sheet name (will be created if it doesn't exist):",
        ui.ButtonSet.OK_CANCEL
    );
    if (destResponse.getSelectedButton() !== ui.Button.OK) return;
    const destSheetName = destResponse.getResponseText().trim();
    if (!destSheetName) {
        ui.alert("Sheet name cannot be empty.");
        return;
    }

    // Step 2: Ask for first filter header
    const header1Response = ui.prompt(
        "Create Filtered View (2/4)",
        "Enter filter header label (e.g. \"Category\"):",
        ui.ButtonSet.OK_CANCEL
    );
    if (header1Response.getSelectedButton() !== ui.Button.OK) return;
    const filterHeader1 = header1Response.getResponseText().trim();
    if (!filterHeader1) {
        ui.alert("Filter header cannot be empty.");
        return;
    }

    // Step 3: Ask for first filter value
    const value1Response = ui.prompt(
        "Create Filtered View (3/4)",
        "Enter value to match for \"" + filterHeader1 + "\" (case-insensitive):",
        ui.ButtonSet.OK_CANCEL
    );
    if (value1Response.getSelectedButton() !== ui.Button.OK) return;
    const filterValue1 = value1Response.getResponseText().trim();

    const filters = [{ header: filterHeader1, value: filterValue1 }];

    // Step 4: Optional second filter
    const addSecondResponse = ui.alert(
        "Create Filtered View (4/4)",
        "Do you want to add a second filter? (AND logic — vendor must match both)\n\nCurrent filter: " + filterHeader1 + " = \"" + filterValue1 + "\"",
        ui.ButtonSet.YES_NO
    );

    if (addSecondResponse === ui.Button.YES) {
        const header2Response = ui.prompt(
            "Second Filter — Header",
            "Enter second filter header label:",
            ui.ButtonSet.OK_CANCEL
        );
        if (header2Response.getSelectedButton() !== ui.Button.OK) return;
        const filterHeader2 = header2Response.getResponseText().trim();

        const value2Response = ui.prompt(
            "Second Filter — Value",
            "Enter value to match for \"" + filterHeader2 + "\":",
            ui.ButtonSet.OK_CANCEL
        );
        if (value2Response.getSelectedButton() !== ui.Button.OK) return;
        const filterValue2 = value2Response.getResponseText().trim();

        if (filterHeader2) {
            filters.push({ header: filterHeader2, value: filterValue2 });
        }
    }

    // Extract all vendor data from Meeting Notes
    const templateHeight = templateSheet.getLastRow();
    const templateContentWidth = templateSheet.getLastColumn();
    const templateCopyWidth = getTemplateMaxColumn(templateSheet, templateHeight);
    const templateRange = templateSheet.getRange(1, 1, templateHeight, templateCopyWidth);

    const allVendors = extractAllVendorData(notesSheet, templateHeight);

    // Read raw sheet data for filter matching (needed because filter headers like
    // "Category" may be column sub-headers, not root-level labels).
    const lastRow = notesSheet.getLastRow();
    const lastCol = notesSheet.getLastColumn();
    const rawData = lastRow > 0 ? notesSheet.getRange(1, 1, lastRow, lastCol).getValues() : [];
    const rawWeights = lastRow > 0 ? notesSheet.getRange(1, 1, lastRow, lastCol).getFontWeights() : [];

    // Build vendor block boundaries from raw data (bold text in column A)
    const vendorBlocks = [];
    for (let i = 0; i < rawData.length; i++) {
        const cellA = rawData[i][0];
        if (cellA && String(cellA).trim() !== "" && rawWeights[i][0] === "bold") {
            vendorBlocks.push({ name: String(cellA).trim(), startRow: i });
        }
    }

    // Set end rows
    for (let i = 0; i < vendorBlocks.length; i++) {
        vendorBlocks[i].endRow = (i + 1 < vendorBlocks.length)
            ? vendorBlocks[i + 1].startRow
            : rawData.length;
    }

    // Filter: for each vendor block, scan all cells for the header text,
    // then check the cell DIRECTLY BELOW for the filter value.
    const matchingVendorNames = new Set();

    for (const block of vendorBlocks) {
        const passesAll = filters.every(filter => {
            const headerLower = filter.header.toLowerCase();
            const valueLower = filter.value.toLowerCase();

            for (let r = block.startRow; r < block.endRow; r++) {
                for (let c = 0; c < rawData[r].length; c++) {
                    const cellVal = String(rawData[r][c] || "").trim().toLowerCase();
                    if (cellVal === headerLower) {
                        // Found the header — check the cell directly below
                        const dataRow = r + 1;
                        if (dataRow < block.endRow) {
                            const dataVal = String(rawData[dataRow][c] || "").trim().toLowerCase();
                            if (dataVal === valueLower) return true;
                        }
                        // Also check same-row value to the right (for label: value pairs)
                        if (c + 1 < rawData[r].length) {
                            const rightVal = String(rawData[r][c + 1] || "").trim().toLowerCase();
                            if (rightVal === valueLower) return true;
                        }
                    }
                }
            }
            return false;
        });

        if (passesAll) {
            matchingVendorNames.add(block.name);
        }
    }

    // Filter the extracted vendor data by matching names
    const matchingVendors = allVendors.filter(v => matchingVendorNames.has(v.name));

    if (matchingVendors.length === 0) {
        const filterSummary = filters.map(f => '"' + f.header + '" = "' + f.value + '"').join(" AND ");
        ui.alert("No vendors matched the filter: " + filterSummary + "\n\nNo changes were made.");
        return;
    }

    // Get or create the destination sheet
    let destSheet = ss.getSheetByName(destSheetName);
    if (destSheet) {
        destSheet.clear();
    } else {
        destSheet = ss.insertSheet(destSheetName);
    }

    // Ensure sheet has enough rows
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockHeight = 1 + templateHeight + blankRows;
    const requiredRows = matchingVendors.length * blockHeight;
    const currentMaxRows = destSheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
        destSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }

    // Write vendor blocks (same as the update flow)
    const fullHeaderWidth = 1 + Math.max(templateContentWidth, templateCopyWidth);
    rebuildVendorBlocks(destSheet, templateSheet, templateRange, matchingVendors, fullHeaderWidth);

    // Inject data back
    const labelMap = buildTemplateLabelMap(templateSheet, templateRange);
    injectDataByLabel(destSheet, matchingVendors, labelMap, templateHeight, templateContentWidth);

    // Copy column widths from Meeting Notes (uniform across all vendor blocks)
    const lastUsedCol = 1 + Math.max(templateContentWidth, templateCopyWidth);
    for (let c = 1; c <= lastUsedCol; c++) {
        destSheet.setColumnWidth(c, notesSheet.getColumnWidth(c));
    }

    // Copy row heights from Meeting Notes — use the first block as the pattern
    // Block layout: 1 header row + templateHeight rows + blankRows
    if (allVendors.length > 0) {
        const srcBlockHeight = 1 + templateHeight + blankRows;
        const srcHeights = [];
        for (let r = 1; r <= srcBlockHeight; r++) {
            srcHeights.push(notesSheet.getRowHeight(r));
        }
        // Apply the same row height pattern to each vendor block in the destination
        for (let v = 0; v < matchingVendors.length; v++) {
            for (let r = 0; r < srcBlockHeight; r++) {
                const destRow = v * blockHeight + r + 1;
                destSheet.setRowHeight(destRow, srcHeights[r]);
            }
        }
    }

    // Clear trailing rows after last vendor block
    const lastVendorEndRow = matchingVendors.length * blockHeight;
    const sheetMaxRow = destSheet.getMaxRows();
    if (sheetMaxRow > lastVendorEndRow) {
        const rowsToClear = sheetMaxRow - lastVendorEndRow;
        destSheet.getRange(lastVendorEndRow + 1, 1, rowsToClear, destSheet.getMaxColumns())
            .clearContent().clearDataValidations();
    }

    // Show summary
    const filterSummary = filters.map(f => '"' + f.header + '" = "' + f.value + '"').join(" AND ");
    ui.alert(
        "✅ Filtered View Created\n\n" +
        "Sheet: \"" + destSheetName + "\"\n" +
        "Filter: " + filterSummary + "\n" +
        "Vendors found: " + matchingVendors.length + " of " + allVendors.length
    );
}

/**
 * Show summary after update completes
 * @param {number} vendorCount - Number of vendors processed
 * @param {Array} warnings - Array of warning messages
 * @param {string} backupName - Name of the backup file
 */
function showSummary(vendorCount, warnings, backupName) {
    const ui = SpreadsheetApp.getUi();
    let message =
        "✅ UPDATE COMPLETE\n\n" +
        "• Vendors processed: " + vendorCount + "\n" +
        "• Backup file: " + backupName + "\n";

    if (warnings.length > 0) {
        message += "• Warnings: " + warnings.length + " (see Logs)\n";
    }

    message += "\n⚠️ IMPORTANT: Please run 'Reconnect Notes Links' now!\n";
    message += "Go to: ⚡ CRM Tools > Reconnect Notes Links";

    ui.alert("Global Update Complete", message, ui.ButtonSet.OK);
}

/**
 * Manual backup function accessible from menu
 */
function manualBackup() {
    const backupName = backupSpreadsheet();
    SpreadsheetApp.getUi().alert("Backup created: " + backupName);
}

/**
 * Reconnect Notes Links in CRM sheet (User-facing with alerts)
 * Wrapper around reconnectNotesLinksInternal that shows UI feedback
 */
function reconnectNotesLinks() {
    const ss = SpreadsheetApp.getActive();
    const result = reconnectNotesLinksInternal(ss);

    if (result.error) {
        SpreadsheetApp.getUi().alert(result.error);
        return;
    }

    SpreadsheetApp.getUi().alert(
        "Bidirectional Links Reconnected\n\n" +
        "CRM → Meeting Notes:\n" +
        "✅ Linked: " + result.linkedCount + " vendors\n" +
        (result.missingCount > 0 ? "⚠️ Not found in Meeting Notes: " + result.missingCount + " vendors\n" : "") +
        "\nMeeting Notes → CRM:\n" +
        "✅ Reverse linked: " + result.reverseLinkedCount + " vendors\n" +
        "\nAll links now point to correct rows."
    );
}

/**
 * Internal function to reconnect Notes Links (no UI alerts)
 * Called by updateAllVendorBlocks for automatic reconnection
 * @param {Spreadsheet} ss - The spreadsheet to work on
 * @returns {Object} Result with linkedCount, missingCount, or error
 */
function reconnectNotesLinksInternal(ss) {
    const crmSheet = ss.getSheetByName("CRM");
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!crmSheet || !notesSheet) {
        return { error: "Missing sheets: CRM or " + CONFIG.notesSheetName };
    }

    // Get Meeting Notes sheet ID for hyperlink
    const notesSheetId = notesSheet.getSheetId();

    // Build vendor → row map from Meeting Notes (scan Column A for bold text)
    const notesLastRow = notesSheet.getLastRow();
    if (notesLastRow === 0) {
        return { error: "Meeting Notes sheet is empty." };
    }

    const notesColA = notesSheet.getRange(1, 1, notesLastRow, 1).getValues();
    const notesWeights = notesSheet.getRange(1, 1, notesLastRow, 1).getFontWeights();

    const vendorRowMap = {}; // vendorName → row number in Meeting Notes (1-indexed)
    for (let i = 0; i < notesColA.length; i++) {
        const cellValue = notesColA[i][0];
        const fontWeight = notesWeights[i][0];

        if (cellValue && String(cellValue).trim() !== "" && fontWeight === 'bold') {
            const vendorName = String(cellValue).trim();
            vendorRowMap[vendorName] = i + 1; // 1-indexed row
        }
    }

    Logger.log("Found " + Object.keys(vendorRowMap).length + " vendors in Meeting Notes");

    // Get CRM data (Column B = vendor names, Column N = links)
    const crmLastRow = crmSheet.getLastRow();
    if (crmLastRow < 2) {
        return { error: "CRM sheet has no data rows." };
    }

    // Get CRM sheet ID for reverse links
    const crmSheetId = crmSheet.getSheetId();

    // Build CRM vendor → row map (for reverse links)
    const crmVendorNames = crmSheet.getRange(2, 2, crmLastRow - 1, 1).getValues(); // Column B, starting row 2
    const crmVendorRowMap = {}; // vendorName → row number in CRM (1-indexed)
    for (let i = 0; i < crmVendorNames.length; i++) {
        const vendorName = crmVendorNames[i][0];
        if (vendorName && String(vendorName).trim() !== "") {
            crmVendorRowMap[String(vendorName).trim()] = i + 2; // Row 2 + offset
        }
    }

    // === PART 1: Update CRM → Meeting Notes links (Column N) ===
    const linksToWrite = [];
    let linkedCount = 0;
    let missingCount = 0;

    for (let i = 0; i < crmVendorNames.length; i++) {
        const vendorName = crmVendorNames[i][0];

        if (!vendorName || String(vendorName).trim() === "") {
            linksToWrite.push([""]);
            continue;
        }

        const vendorNameTrimmed = String(vendorName).trim();
        const notesRow = vendorRowMap[vendorNameTrimmed];

        if (notesRow) {
            // Create HYPERLINK formula: =HYPERLINK("#gid=SHEET_ID&range=A{ROW}", "View Notes")
            const formula = '=HYPERLINK("#gid=' + notesSheetId + '&range=A' + notesRow + '", "View Notes")';
            linksToWrite.push([formula]);
            linkedCount++;
        } else {
            // Vendor not found in Meeting Notes - keep as plain text (no link)
            linksToWrite.push(["View Notes"]);
            missingCount++;
            Logger.log("Warning: Vendor not found in Meeting Notes: " + vendorNameTrimmed);
        }
    }

    // Write all CRM links at once (Column N = column 14)
    const linksRange = crmSheet.getRange(2, 14, linksToWrite.length, 1);
    linksRange.setValues(linksToWrite);

    Logger.log("CRM → Meeting Notes: Reconnected " + linkedCount + " links, " + missingCount + " vendors not found");

    // === PART 2: Update Meeting Notes → CRM links (Column A vendor headers) ===
    let reverseLinkedCount = 0;

    // Process each vendor in Meeting Notes
    for (const vendorName in vendorRowMap) {
        const notesRow = vendorRowMap[vendorName];
        const crmRow = crmVendorRowMap[vendorName];

        if (crmRow) {
            // Create HYPERLINK formula: =HYPERLINK("#gid=CRM_SHEET_ID&range=B{ROW}", "VendorName")
            const formula = '=HYPERLINK("#gid=' + crmSheetId + '&range=B' + crmRow + '", "' + vendorName.replace(/"/g, '""') + '")';

            // Write the hyperlink formula to the vendor header cell
            const vendorCell = notesSheet.getRange(notesRow, 1);
            vendorCell.setFormula(formula);

            // Preserve the header formatting (bold, magenta background, black text)
            vendorCell.setFontWeight("bold");
            vendorCell.setBackground(CONFIG.vendorHeaderBgColor);  // Magenta
            vendorCell.setFontColor("#ffffff");  // White text

            reverseLinkedCount++;
        } else {
            Logger.log("Warning: Vendor not found in CRM: " + vendorName);
        }
    }

    Logger.log("Meeting Notes → CRM: Reconnected " + reverseLinkedCount + " reverse links");

    return {
        linkedCount: linkedCount,
        missingCount: missingCount,
        reverseLinkedCount: reverseLinkedCount
    };
}

/**
 * Rename a header label across all vendor blocks
 * This is a SAFE operation that directly modifies cells without data extraction
 * The header is renamed in both Meeting Notes AND Template_LeftBlock
 */
function renameHeader() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActive();

    // Step 1: Get old header name from user
    const oldNameResponse = ui.prompt(
        '🔄 Rename Header - Step 1 of 2',
        'Enter the CURRENT header name (exactly as it appears):',
        ui.ButtonSet.OK_CANCEL
    );

    if (oldNameResponse.getSelectedButton() !== ui.Button.OK) {
        return; // User cancelled
    }

    const oldName = oldNameResponse.getResponseText().trim();
    if (!oldName) {
        ui.alert('Error', 'Header name cannot be empty.', ui.ButtonSet.OK);
        return;
    }

    // Step 2: Get new header name from user
    const newNameResponse = ui.prompt(
        '🔄 Rename Header - Step 2 of 2',
        'Enter the NEW header name:\n\nRenaming: "' + oldName + '"',
        ui.ButtonSet.OK_CANCEL
    );

    if (newNameResponse.getSelectedButton() !== ui.Button.OK) {
        return; // User cancelled
    }

    const newName = newNameResponse.getResponseText().trim();
    if (!newName) {
        ui.alert('Error', 'New header name cannot be empty.', ui.ButtonSet.OK);
        return;
    }

    if (oldName === newName) {
        ui.alert('No Change', 'Old and new names are the same. Nothing to do.', ui.ButtonSet.OK);
        return;
    }

    // Step 3: Confirm the rename
    const confirmResult = ui.alert(
        'Confirm Rename',
        'This will rename ALL occurrences of:\n\n' +
        '"' + oldName + '"\n\n' +
        'to:\n\n' +
        '"' + newName + '"\n\n' +
        'in both Meeting Notes and Template_LeftBlock sheets.\n\n' +
        'Continue?',
        ui.ButtonSet.YES_NO
    );

    if (confirmResult !== ui.Button.YES) {
        return; // User cancelled
    }

    // Step 4: Perform the rename
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet) {
        ui.alert('Error', 'Meeting Notes sheet not found.', ui.ButtonSet.OK);
        return;
    }

    if (!templateSheet) {
        ui.alert('Error', 'Template_LeftBlock sheet not found.', ui.ButtonSet.OK);
        return;
    }

    let notesCount = 0;
    let templateCount = 0;

    // Rename in Meeting Notes
    const notesLastRow = notesSheet.getLastRow();
    const notesLastCol = notesSheet.getLastColumn();

    if (notesLastRow > 0 && notesLastCol > 0) {
        const notesRange = notesSheet.getRange(1, 1, notesLastRow, notesLastCol);
        const notesValues = notesRange.getValues();
        let modified = false;

        for (let r = 0; r < notesValues.length; r++) {
            for (let c = 0; c < notesValues[r].length; c++) {
                const cellValue = notesValues[r][c];
                if (cellValue !== null && String(cellValue).trim() === oldName) {
                    notesValues[r][c] = newName;
                    notesCount++;
                    modified = true;
                }
            }
        }

        if (modified) {
            notesRange.setValues(notesValues);
        }
    }

    // Rename in Template_LeftBlock
    const templateLastRow = templateSheet.getLastRow();
    const templateLastCol = templateSheet.getLastColumn();

    if (templateLastRow > 0 && templateLastCol > 0) {
        const templateRange = templateSheet.getRange(1, 1, templateLastRow, templateLastCol);
        const templateValues = templateRange.getValues();
        let modified = false;

        for (let r = 0; r < templateValues.length; r++) {
            for (let c = 0; c < templateValues[r].length; c++) {
                const cellValue = templateValues[r][c];
                if (cellValue !== null && String(cellValue).trim() === oldName) {
                    templateValues[r][c] = newName;
                    templateCount++;
                    modified = true;
                }
            }
        }

        if (modified) {
            templateRange.setValues(templateValues);
        }
    }

    // Step 5: Show summary
    const totalCount = notesCount + templateCount;

    if (totalCount === 0) {
        ui.alert(
            'No Matches Found',
            'Could not find any cells containing:\n\n"' + oldName + '"\n\n' +
            'Make sure you typed the header name exactly as it appears (including punctuation and capitalization).',
            ui.ButtonSet.OK
        );
    } else {
        ui.alert(
            '✅ Rename Complete',
            'Successfully renamed "' + oldName + '" to "' + newName + '"\n\n' +
            '• Meeting Notes: ' + notesCount + ' cells updated\n' +
            '• Template: ' + templateCount + ' cells updated\n' +
            '• Total: ' + totalCount + ' cells\n\n' +
            'All vendor blocks and the template are now in sync.',
            ui.ButtonSet.OK
        );
    }

    Logger.log('Header rename complete: "' + oldName + '" → "' + newName + '" (' + totalCount + ' cells)');
}
