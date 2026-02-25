# Context Brief — Google Sheets CRM

## Project Overview

A Google Apps Script-based CRM system built inside Google Sheets. The primary sheet is **"Meeting Notes"**, which contains vendor blocks — each block has a **bold vendor name in column A** followed by template rows (labels in column B, values in columns C+). Blocks are separated by blank rows.

**Workspace:** `/Users/saadtariq/Documents/AG/`
**Deployment:** Uses `clasp` to push to Google Apps Script. Run with:
```bash
export PATH="$HOME/node-js/bin:$HOME/npm-global/bin:$PATH" && clasp push --force
```
**Current target script ID** (in `.clasp.json`): `1RJvcS8RhhlTTILpJkux5X6Sjdrm6eUN02jc8aZs8HPlvBM-DtHy12gv5`

---

## File Structure

| File | Purpose |
|---|---|
| `UpdateVendoBlock.js` | Core engine: `extractAllVendorData`, `rebuildVendorBlocks`, `injectDataByLabel`, `createFilteredView`, `buildTemplateLabelMap`, `getTemplateMaxColumn`, batch update system |
| `SheetTools.js` | General tools: `rankVendors`, `moveVendorToSheet`, `removeVendor`, `syncSheets` |
| `PhilSync.js` | Legacy bi-directional sync for "Phil" sheet (temporary feature, mostly superseded by `SheetTools.js`) |
| `Menu.gs.js` | `onOpen()` menu builder + batch completion notification |
| `createNewVendor.gs.js` | Creates a new vendor block |
| `MailingListSync.js` | Syncs mailing list data |
| `VendorCore.js` | Core vendor utilities |
| `Vendor_ID.js` | Vendor ID management |
| `Reminders.js` | Reminder system |
| `SlackNotifications.js` | Slack integration |
| `ZoomLinks.js` | Zoom link handling |
| `Template_LeftBlock.js` | Template block utilities |

---

## Key Architecture Concepts

### Vendor Block Layout
```
Row N:   [Vendor Name (bold, magenta bg)] [magenta fill across columns...]
Row N+1: [Template row 1 - labels in col B, values in C+]
Row N+2: [Template row 2...]
...
Row N+T: [Last template row]
Row N+T+1: [Blank row separator]
```
- **Column A**: ONLY vendor name (header row). Template data starts at column **B**.
- **Column E, row 2 of block** (first data row): Contains rank number (1-5).
- Template height = `templateSheet.getLastRow()`
- Block height = `1 + templateHeight + CONFIG.blankRowsAfterVendor`

### CONFIG Object (in UpdateVendoBlock.js)
```javascript
const CONFIG = {
    notesSheetName: "Meeting Notes",
    templateSheetName: "Template",
    blankRowsAfterVendor: 1,
    vendorHeaderBgColor: "#ff00ff",
    updateBatchSize: 100,
    preserveRichTextLinks: true,
    maxColumns: 15
};
```

### Core Functions (UpdateVendoBlock.js)
- **`extractAllVendorData(sheet, templateHeight)`** — Reads all vendor blocks from a sheet. Returns `[{name, data: [{label, values, valueFormulas, mergedCellContent, mergedCellFormula}]}]`
- **`rebuildVendorBlocks(sheet, templateSheet, templateRange, vendorData, fullHeaderWidth)`** — Clears sheet and writes formatted vendor blocks with template
- **`injectDataByLabel(sheet, vendorData, labelMap, templateHeight, templateContentWidth)`** — Writes extracted data back into rebuilt blocks by label matching
- **`buildTemplateLabelMap(templateSheet, templateRange)`** — Maps label names to `{row, labelCol, valueCol}` positions
- **`getTemplateMaxColumn(templateSheet, templateHeight)`** — Finds rightmost used column including merges

---

## Current Menu (Menu.gs.js)
```
⚡ CRM Tools
├── Create New Vendor
├── Update ALL Vendor Blocks
├── Sync Mailing List
├── Reconnect Notes Links
├── Rename Header
├── Create Filtered View         ← uses OR logic for multiple filters
├── ─────────────
├── Rank Vendors                 ← NEW
├── Move Vendor to Sheet         ← NEW
├── Remove Vendor                ← NEW
├── Sync Sheets                  ← NEW
├── ─────────────
├── Rename Sheet
└── Manual Backup
```

---

## Recently Implemented Features (SheetTools.js)

### 1. Rank Vendors
Sorts vendors in a sheet by the number in column E (row 2 of each block). Ascending order, blanks at end. Fully rebuilds the sheet.

### 2. Move Vendor to Sheet
Copies a vendor from Meeting Notes to any destination sheet. Appends at end. Creates the sheet if it doesn't exist.

### 3. Remove Vendor
Deletes an entire vendor block from any sheet by name.

### 4. Sync Sheets
One-shot bi-directional sync between any sheet and Meeting Notes:
- Empty in one sheet + filled in other → filled value wins
- Both differ with non-empty values → Meeting Notes wins
- No persistent cache — direct comparison each time

### 5. Expand New Sheet Vendors
Reads a flat sheet of vendor rows, groups contacts by vendor, and creates fully formatted template blocks in a destination sheet, auto-populating fields and inserting extra rows for multiple contacts.

### 6. Create Filtered View (AND → OR)
Changed from `filters.every()` to `filters.some()` — vendors matching ANY filter condition are included.

---

## Critical Bugs Fixed (Reference for Future Issues)

### 1. Data Validation Crash
**Problem:** `setValues()` crashes on cells with dropdown validations.
**Fix:** Save validations → clear → write values → restore validations.

### 2. Checkbox → TRUE/FALSE
**Problem:** `clearDataValidations()` removes checkbox formatting.
**Fix:** Same save/restore pattern above preserves checkbox validations.

### 3. Phantom Sync Updates (setFormulas)
**Problem:** Batch `setFormulas()` with empty strings `""` clears non-formula cell values.
**Fix:** Write formulas **cell-by-cell**, only for cells that actually have a formula:
```javascript
for (let r = 0; r < templateHeight; r++) {
    for (let c = 0; c < templateContentWidth; c++) {
        const f = formulasToSet[r][c];
        if (f && String(f).trim() !== "") sheet.getRange(row + r, 2 + c).setFormula(f);
    }
}
```

### 4. Phantom Sync Updates (Array Length Mismatch)
**Problem:** `extractAllVendorData` produces different-length value arrays depending on sheet column count. Using `Math.max()` caused out-of-bounds indices to appear as null → false "Notes changed" diffs.
**Fix:** Use `Math.min()` to cap comparison to positions both sheets actually have.

### 5. Template Column Offset
**Problem:** Template copied to column A instead of column B.
**Fix:** Always copy template to column 2: `destSheet.getRange(row, 2, height, width)`

---

## Pending / Known Items

1. **PhilSync.js cleanup** — The Phil-specific menu items were removed, but the file still exists. It can be deleted when no longer needed for reference.
2. **Sync Sheets** — Currently no persistent snapshot, so it can't detect "who changed what" in conflicts — it uses a simple rule (Notes wins). If true 3-way merge is needed later, the PhilSync cache pattern can be reintroduced.
3. **Rank Vendors** — Had not been tested by the user yet at session end.
4. **Testing all new features** on the target project ID `1RJvcS8RhhlTTILpJkux5X6Sjdrm6eUN02jc8aZs8HPlvBM-DtHy12gv5`.

---

## Deployment Notes

- Always use `clasp push --force` (force overwrites remote)
- The `.clasp.json` scriptId gets changed frequently when pushing to different spreadsheets
- `skipSubdirectories: true` in `.clasp.json` prevents pushing non-GAS files
- Running locally: `export PATH="$HOME/node-js/bin:$HOME/npm-global/bin:$PATH"`
