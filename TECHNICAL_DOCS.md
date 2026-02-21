# CRM Tools - Technical Documentation

> **Version:** 0.1.2  
> **Last Updated:** January 28, 2026

---

## Architecture Overview

### File Structure

```
/Users/saadtariq/Documents/AG/
├── .clasp.json              # CLASP configuration (script ID)
├── appsscript.json          # Apps Script manifest
├── Menu.gs.js               # Menu creation (onOpen trigger)
├── UpdateVendoBlock.js      # Core update engine
├── createNewVendor.gs.js    # New vendor creation
├── Vendor_ID.js             # Vendor ID utilities
├── Reminders.js             # Reminder functionality
├── Template_LeftBlock.js    # Template utilities
├── testTemplate.gs.js       # Test utilities
├── CHANGELOG.md             # Version history
└── USER_GUIDE.md            # User documentation
```

### Sheet Dependencies

| Sheet Name | Required | Purpose |
|------------|----------|---------|
| CRM | Yes | Master vendor list |
| Meeting Notes | Yes | Vendor data blocks |
| Template_LeftBlock | Yes | Block template source |
| Stages | No | Dropdown values |

---

## Core Functions

### UpdateVendoBlock.js

#### `updateAllVendorBlocks()`
Main orchestration function with 7 phases:

1. **Phase 1:** Extract all vendor data (`extractAllVendorData`)
2. **Phase 2:** Confirmation dialog (`showConfirmationDialog`)
3. **Phase 3:** Create backup (`backupSpreadsheet`)
4. **Phase 4:** Build label map (`buildTemplateLabelMap`)
5. **Phase 5:** Rebuild blocks (`rebuildVendorBlocks`)
6. **Phase 6:** Inject data (`injectDataByLabel`)
   - 6.5: Clear trailing rows
   - 6.55: Clear trailing columns
   - 6.6: Reconnect notes links
7. **Phase 7:** Show summary (`showSummary`)

#### `extractAllVendorData(sheet, templateHeight)`
V10.1 Pure Positional Extraction:
- Scans Column A for bold headers (vendor names)
- Extracts all values after each label as positional arrays
- Special handling for merged cell labels (Notes Field, Follow Up Action)

#### `injectDataByLabel(sheet, vendorData, labelMap, templateHeight)`
V10 Pure Positional Injection:
- Writes values at same relative positions
- Handles merged cell content separately
- Returns warnings for missing labels

#### `reconnectNotesLinks()` / `reconnectNotesLinksInternal(ss)`
- Scans CRM Column B for vendor names
- Finds matching bold text in Meeting Notes Column A
- Generates HYPERLINK formulas: `=HYPERLINK("#gid=xxx&range=Ayyy", "View Notes")`

---

## Configuration

```javascript
const CONFIG = {
  notesSheetName: "Meeting Notes",
  templateSheetName: "Template_LeftBlock",
  labelColumns: [1, 2],           // Labels in columns A and B
  valueStartColumn: 2,            // Values start in column B
  vendorHeaderColumn: 1,          // Column A (vendor name)
  maxColumns: 11,                 // Deprecated - now dynamic
  headerColor: "#1f4e79",         // Dark Blue
  headerFontColor: "#ffffff"      // White
};
```

---

## Data Flow

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│ Template_       │────→│ Meeting Notes    │←───→│ CRM             │
│ LeftBlock       │     │ (Vendor Blocks)  │     │ (Master List)   │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                        │                        │
        │   copy template        │  extract/inject        │  links
        └────────────────────────┴────────────────────────┘
```

---

## Vendor Detection

Vendors are identified by:
1. **Bold text** in Column A
2. **Dark blue background** (#1f4e79)
3. **White font color** (#ffffff)

```javascript
// Detection logic in extractAllVendorData
if (fontWeight === 'bold' && firstNonEmpty !== "") {
  // This is a vendor header
}
```

---

## Merged Cell Handling

Special labels with content in merged cells below:
- **Notes Field** (Column I)
- **Follow Up Action** (Column I)

```javascript
const MERGED_CONTENT_LABELS = ['Notes Field', 'Follow Up Action'];

// Look-ahead for merged content
if (MERGED_CONTENT_LABELS.includes(label) && !hasContent) {
  // Check next rows for content
  mergedCellContent = contentAtLabelCol;
}
```

---

## Known Limitations

1. **Column Insertion:** New columns must be added at END of sections
2. **Positional Mapping:** Data mapped by position, not by header name
3. **Vendor Name Matching:** Must be exact match between CRM and Meeting Notes

---

## Deployment

### Push to Apps Script
```bash
export PATH="$HOME/node-js/bin:$HOME/npm-global/bin:$PATH"
clasp push --force
```

### Current Script ID
```
1nT7pcDbt8BBhFvqJn2Cj_634Dw-YbBPxyQA3UD5LR-v7KxUihRH23bfo
```

---

## Version History

See [CHANGELOG.md](file:///Users/saadtariq/Documents/AG/CHANGELOG.md) for full version history.
