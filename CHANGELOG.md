# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). 

## [0.1.3] - 2026-01-29

### Added
- **Rename Header Feature**:
  - New menu item: `⚡ CRM Tools > Rename Header`
  - Guided UI prompts for old and new header names
  - Safely renames headers across ALL vendor blocks in Meeting Notes AND Template_LeftBlock
  - No data extraction/injection required (zero data loss risk)
  - Shows summary of cells updated

- **Vendor Block Spacing**:
  - Added `blankRowsAfterVendor` config option (default: 1)
  - 1 blank row automatically added between each vendor block
  - Improves visual separation between vendors

- **Full-Width Header Row Coloring**:
  - Added `vendorHeaderBgColor` config option (default: #ff00ff magenta)
  - Entire vendor header row colored magenta across full content width
  - Vendor name cell included (Column A)
  - Black text (#000000) for visibility on magenta background

### Changed
- Updated `rebuildVendorBlocks()` to accept `fullHeaderWidth` parameter
- Updated `injectDataByLabel()` row calculations to include blank rows
- Updated Phase 6.5 cleanup to account for blank rows
- Updated `reconnectNotesLinksInternal()` to use magenta/black for vendor cells

---

## [0.1.2] - 2026-01-28

### Added
- **V10 Pure Positional Extraction/Injection (Major Simplification)**:
  - **Abandoned Header-Keyed Approach**: After analysis revealed that template rows have MULTIPLE independent label zones (left zone in column B, right zone in columns F-I), the complex header-keyed approaches (V6-V9) were abandoned.
  - **Simple Positional Strategy**: Extract ALL values after each label as a positional array. Inject ALL values at the same relative positions. This is more reliable for the actual use case.
  - **Removed Split Left/Right Logic**: The LEFT_SIDE_END = 3 split point didn't work for all row types (some rows like Source labels have different structures).

- **V10.1 Merged Cell Support**:
  - **Special Handling for Notes Field & Follow Up Action**: These labels have content in merged cells on subsequent rows, not in the same row.
  - **Look-Ahead Extraction**: When extracting these labels and finding no content on the label row, the code looks at subsequent rows to find the merged cell content.

- **V10.2 Merged Cell Location Fix**:
  - **Fixed Content Placement**: V10.1 wrote merged content NEXT TO the label (wrong). V10.2 correctly writes to the NEXT ROW at the LABEL column position (where the merged cell area is).
  - **`mergedCellContent` Field**: Added separate field to store merged cell content, with special injection logic.

- **Trailing Row Cleanup**:
  - Added **Phase 6.5** to `updateAllVendorBlocks()` that clears all rows after the last vendor block.
  - Removes leftover checkboxes and content when the template shrinks (fewer rows).
  - Uses `clearContent()` and `clearDataValidations()` for complete cleanup.

- **Reconnect Notes Links**:
  - New menu item: `⚡ CRM Tools > Reconnect Notes Links`
  - Regenerates "View Notes" hyperlinks in CRM Column N to point to correct vendor rows in Meeting Notes.
  - Matches vendor names between CRM (Column B) and Meeting Notes (Column A bold headers).
  - **Auto-Reconnect**: Automatically runs after "Update ALL Vendor Blocks" to keep links in sync.
- **Bidirectional Links**: Meeting Notes vendor names are now clickable links pointing to CRM.
  - Click vendor header in Meeting Notes → jumps to CRM row
  - Formatting (bold, blue background) is preserved

### Fixed
- **Notes Field Data Loss**: The Notes Field and Follow Up Action content was not being extracted because the label is at column I and the content is in a merged cell on subsequent rows. Now correctly extracted and injected.
- **FALSE Text in Source Fields**: V8-V9 had issues with booleans appearing as text. V10 pure positional approach preserves types correctly.
- **Contact Info Data Loss**: V9 split logic caused data loss for multi-zone rows. V10 pure positional fixes this.
- **Trailing Checkboxes Not Cleared**: Fixed Phase 6.5 to use `getMaxRows()` instead of `getLastRow()`. Checkboxes without content are now properly detected and cleared.
- **Removed Columns Persisting**: Added Phase 6.55 that clears trailing columns when template width shrinks. Data validations (dropdowns) in removed columns are now cleared.
- **Merged Cell Width Not Copied**: Added `getTemplateMaxColumn()` function that scans merged ranges to find the true template width. Merged cells extending beyond the last content column are now included.

### Known Limitations
- **Column Insertion**: When modifying the template, **new columns must be added at the END of each section** (e.g., after "Zoom" in Contact Info, after "Category" in Company Info). Adding columns in the MIDDLE of a section will cause data to shift incorrectly.
  - ✅ Supported: Add column at end of section
  - ✅ Supported: Add/remove rows anywhere
  - ⚠️ Not Supported: Insert column in middle of section

## [0.1.1] - 2026-01-28

### Added
- **V4 Simple Positional Mapping (Fixes Duplicates + Notes + Source)**:
  - **Expanded Label Detection**: Now scans ALL columns for the first non-empty cell as label. This fixes "Notes Field" and "Follow Up Action" which have labels in Column I, not A/B.
  - **Simple Positional Injection**: Values are written at positions relative to the label column. No complex header-keyed logic that caused duplicates.
  - **Boolean Preservation**: Boolean values (checkbox states) are preserved as booleans, not converted to "TRUE"/"FALSE" text.
  - **Reverted V3**: Removed header-keyed extraction/injection which caused duplicate columns when templates changed.
- **V3 Structure-Aware Field Mapping**: Complete redesign of extraction and injection engine.
  - **Header-Keyed Extraction**: Data is now extracted as `{headerName: value}` pairs (e.g., `{"Name": "Tony", "Title": "CEO"}`), not positional arrays.
  - **Sub-Header Detection**: Rows with 3+ text values are detected as sub-headers and used to build column mappings.
  - **Header-Aware Injection**: On injection, values are written to columns by looking up where each header exists in the NEW template. This eliminates duplicate columns when template changes.
  - **Automatic Header Row Skipping**: Header rows are not overwritten during injection, trusting the template's structure.
- **Simplified Injection Engine V2**: Complete rewrite of the data injection logic. Removed complex Header Alignment heuristics. Now uses direct label→row mapping: for each extracted label, values are written directly to the matching template row. No rows are skipped, ensuring ALL data (including "Source", "Notes", etc.) is preserved.
- **Terminal Column Preservation**: Fixed a bug where data in columns beyond the template's explicit width (like "Zoom") was being truncated. The injection loop now iterates over the maximum of new template columns AND old data columns, ensuring all data is written.
- **Stricter Header Detection**: Increased the threshold for detecting "header rows" from 1 match to 2 matches. This prevents single-word values (like a data entry "Zoom") from being mistakenly identified as headers, which caused data loss.
- **Header Alignment Protocol**: The system now dynamically detects column shifts (e.g., deleted or reordered columns) by matching extracted data against the new template's headers. If a row contains matching headers (e.g., "Name", "Title"), the system recalculates the column mapping for that section, ensuring that subsequent data rows are injected into the correct columns, preventing data shifting.
- **Batch Data Injection**: Optimization of the injection phase to use `setValues()` on the entire vendor block at once, rather than updating cell-by-cell. This drastically reduces API calls (from ~1000 to ~50 per update) and eliminates "Execution Time Exceeded" errors for large datasets.
- **Corrected Template Label Mapping**: Adjusted the template scanning logic to prioritize Column A of the source template. This ensures that labels are correctly identified even when the template is intended to be pasted into Column B of the target sheet, solving the "Extraction Works, Injection Fails" data loss issue.
- **Dual-Column Data Extraction**: Implemented intelligent feedback loops in the extraction engine that scan BOTH Column A and Column B for labels. This ensures zero data loss when migrating from older layouts (where labels were in Col A) to the new layout (Labels in Col B).
- **Ghost Artifact Cleanup**: Added a sanitation step that actively removes checkboxes, validations, and content from the Vendor Header row, preventing "ghost" checkboxes from appearing next to the vendor name.
- **Exclusive Header Row Layout**: Refined the layout so that the Vendor Name occupies its own exclusive row (Column A), and the Template Block starts on the *next row* (Column B). This improves visual separation and clarity.
- **Ghost Artifact Cleanup**: Added a sanitation step that actively removes checkboxes, validations, and content from the Vendor Header row, preventing "ghost" checkboxes from appearing next to the vendor name.
- **Exclusive Header Row Layout**: Refined the layout so that the Vendor Name occupies its own exclusive row (Column A), and the Template Block starts on the *next row* (Column B). This improves visual separation and clarity.
- **Ghost Artifact Cleanup**: Added a sanitation step that actively removes checkboxes, validations, and content from the Vendor Header row, preventing "ghost" checkboxes from appearing next to the vendor name.
- **Exclusive Header Row Layout**: Refined the layout so that the Vendor Name occupies its own exclusive row (Column A), and the Template Block starts on the *next row* (Column B). This improves visual separation and clarity.
- **Robust Data Extraction**: Fixed a critical extraction bug where data located on the same row as the Vendor Name (in previous layouts) was being skipped, causing data loss. The extractor now strictly scans the entire block range including the header row.
- **Column-Aware Data Mapping**: The engine now maps data based on the *column relationship* between the label and the value (e.g., "Value is 1 column right of Label") rather than fixed grid positions. This prevents data shifting when new columns are inserted into the template.
- **B1-Offset Layout Support**: Updated the system to support a cleaner layout where:
    - **Column A** is exclusively for Vendor Headers.
    - **Column B+** contains the Template structure (Labels & Data).
    - The Template block now starts at the *same row* as the Vendor Header, optimizing vertical space.
- **Dynamic Template Sizing**: The engine now automatically detects the height of the `Template_LeftBlock` sheet, allowing for updates to templates of any size without code changes.
- **Strict Header Formatting**: Enforced `Bold` + `Dark Blue (#1f4e79)` + `White Text` formatting for vendor headers to eliminate false positives during detection.
- **Testing Suite**: Comprehensive Jest-based test suite for `UpdateVendoBlock.js` covering unit and integration scenarios.

### Fixed
- **Vendor Extraction Logic**: Fixed a critical bug where labels in Column A were incorrectly identified as vendor headers. The logic now strictly requires **Bold** font weight for headers, ensuring robust extraction of existing data.

## [0.1.0] - 2026-01-28

### Added
- **Manual Backup**: Added a "Manual Backup" option to the "⚡ CRM Tools" menu to allow users to create timestamped backups of the spreadsheet on demand.
- **Safety Interlocks**: Implemented a confirmation dialog that shows the number of vendors to be updated and the new template height before any destructive operation begins.
- **Data Preservation Engine**:
    - Added `extractAllVendorData()` to safely read all existing vendor block data (labels and values) before block deletion.
    - Added `injectDataByLabel()` to restore data into the new block structure by matching labels.
    - Added `buildTemplateLabelMap()` to dynamically understand the new template's structure.
- **Merged Cell Handling**: Introduced `safeSetValue()` utility to write data correctly into merged cells (anchoring to top-left).
- **Configuration**: Added `CONFIG` object to centralize sheet names, column mappings, and template dimensions.

### Changed
- **Update Logic Refactor**: Completely rewrote `updateAllVendorBlocks()` in `UpdateVendoBlock.js`.
    - **Old Behavior**: Identified vendor header, deleted rows, copied template (caused 100% data loss of values).
    - **New Behavior**: 
        1. Validates sheets.
        2. Extracts all data.
        3. Prompts for confirmation.
        4. Creates a backup.
        5. Rebuilds blocks with new template.
        6. Injects preserved data.
        7. Reports summary of changes and warnings.
- **Menu**: Renamed menu to "⚡ CRM Tools" and added a separator for better UX.

### Fixed
- **Data Loss Bug**: Resolved critical issue where updating vendor blocks would wipe all entered data.
- **Merged Cell Errors**: Prevented script failures when attempting to set values in merged ranges.
