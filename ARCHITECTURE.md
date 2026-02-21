# CRM Automation System - Complete Architecture Document

> **Version**: 2.0  
> **Last Updated**: January 31, 2026  
> **Purpose**: Comprehensive technical reference for developers, maintainers, and users

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [Sheet Architecture](#2-sheet-architecture)
3. [Code Files & Functions](#3-code-files--functions)
4. [Data Flow & Connections](#4-data-flow--connections)
5. [Configuration Reference](#5-configuration-reference)
6. [User Guide](#6-user-guide)
7. [Constraints & Limitations](#7-constraints--limitations)
8. [⚠️ Things to Avoid](#8-️-things-to-avoid)
9. [Troubleshooting](#9-troubleshooting)
10. [Future Development Notes](#10-future-development-notes)

---

## 1. System Overview

### 1.1 Purpose
This Google Apps Script system automates CRM operations for a vendor management workflow built on Google Sheets. It handles:
- **Vendor Block Management**: Creating, updating, and rebuilding structured data blocks in Meeting Notes
- **Bidirectional Linking**: Hyperlinks between CRM sheet and Meeting Notes
- **Zoom Integration**: Automatic linking of Zoom meeting cells to a dedicated Zoom Links sheet
- **Slack Notifications**: Daily automated alerts for date-based triggers (touches, follow-ups, timelines)
- **Template Synchronization**: Applying template changes to all existing vendor blocks

### 1.2 Technology Stack
| Component | Technology |
|-----------|------------|
| Platform | Google Apps Script (V8 runtime) |
| Storage | Google Sheets |
| Notifications | Slack Incoming Webhooks |
| Deployment | clasp (Command Line Apps Script) |
| Version Control | Local Git + clasp push |

### 1.3 File Structure
```
/Users/saadtariq/Documents/AG/
├── Menu.gs.js              # Menu definition (onOpen)
├── createNewVendor.gs.js   # Create new vendor workflow
├── UpdateVendoBlock.js     # Main update engine (24 functions)
├── ZoomLinks.js            # Zoom Links sheet integration
├── SlackNotifications.js   # Slack daily notifications
├── Reminders.js            # Legacy reminder system
├── Template_LeftBlock.js   # Template utilities
├── Vendor_ID.js            # Vendor ID helpers
├── testTemplate.gs.js      # Test utilities
├── appsscript.json         # Apps Script manifest
└── .clasp.json             # clasp configuration
```

---

## 2. Sheet Architecture

### 2.1 Sheet Overview
| Sheet Name | Purpose | Structure |
|------------|---------|-----------|
| **CRM** | Master vendor list | Row-per-vendor, columns A-Z+ |
| **Meeting Notes** | Detailed vendor blocks | Vendor header + template blocks |
| **Template_LeftBlock** | Reusable template | Single block defining structure |
| **Zoom Links** | Zoom meeting link storage | 12-row blocks per vendor |
| **Account Notes** | Legacy notes (for reminders) | Free-form |

### 2.2 Meeting Notes Structure

Each vendor occupies a **fixed-size block** in Meeting Notes:

```
Row 1:  [VENDOR NAME]  │ ← Magenta background, white text, bold
Row 2:  │ Template Row 1 (copied from Template_LeftBlock)
Row 3:  │ Template Row 2
...     │ ...
Row N:  │ Template Row (N-1)
Row N+1:│ (blank row - spacing)
```

**Block Calculation**:
```javascript
blockSize = 1 (vendor header) + templateHeight + blankRowsAfterVendor
```

### 2.3 Zoom Links Structure

Each vendor occupies a **12-row block**:

```
Row 1:  [VENDOR NAME]  │                │               ← Magenta, 3 cols
Row 2:                 │ Person Name    │ Meeting Link  ← #d9d2e9 header
Row 3:                 │ Person 1       │ https://...
Row 4:                 │ Person 2       │ https://...
...
Row 11:                │ Person 9       │ https://...
Row 12: (blank)
```

### 2.4 CRM Sheet Columns
| Column | Field | Notes |
|--------|-------|-------|
| A | Vendor ID | Auto-generated: `v_XXXXXXXX` |
| B | Company Name | Primary identifier |
| C-... | Various fields | Configurable |
| Notes column | Hyperlink | Links to Meeting Notes vendor row |

---

## 3. Code Files & Functions

### 3.1 Menu.gs.js
**Purpose**: Defines the custom menu that appears when the sheet opens.

```javascript
function onOpen() {
  // Creates "⚡ CRM Tools" menu with all available actions
}
```

**Menu Structure**:
```
⚡ CRM Tools
├── Create New Vendor
├── Update ALL Vendor Blocks
├── Reconnect Notes Links
├── Rename Header
├── ─────────────────────
├── Slack Notifications ▶
│   ├── Run Test Scan (Check Today)
│   └── Setup Daily Trigger (4:30 AM)
├── ─────────────────────
├── Rename Sheet
└── Manual Backup
```

---

### 3.2 createNewVendor.gs.js
**Purpose**: Creates a new vendor in both CRM and Meeting Notes sheets.

#### Function: `createNewVendor()`
| Step | Action |
|------|--------|
| 1 | Get active row in CRM sheet |
| 2 | Validate selection (not header row) |
| 3 | Extract vendor name from column B |
| 4 | Generate unique vendor ID (`v_XXXXXXXX`) |
| 5 | Write vendor ID to column A |
| 6 | Calculate start row in Meeting Notes (with spacing) |
| 7 | Write vendor header (magenta background, white text) |
| 8 | Copy template below vendor header |
| 9 | Create vendor block in Zoom Links sheet |
| 10 | Show success message |

**Dependencies**:
- `CONFIG` (from UpdateVendoBlock.js)
- `getTemplateMaxColumn()` (from UpdateVendoBlock.js)
- `createZoomVendorBlockForNewVendor()` (from ZoomLinks.js)

---

### 3.3 UpdateVendoBlock.js (Main Engine)

**Purpose**: Core update engine for vendor block management.

#### Configuration Object: `CONFIG`
```javascript
const CONFIG = {
  notesSheetName: "Meeting Notes",
  templateSheetName: "Template_LeftBlock",
  blankRowsAfterVendor: 1,
  vendorHeaderBgColor: "#ff00ff",  // Magenta
  updateBatchSize: 10,
  preserveRichTextLinks: false
};
```

#### Core Functions

| Function | Purpose | Triggers |
|----------|---------|----------|
| `updateAllVendorBlocks()` | Main update workflow | Menu button |
| `reconnectNotesLinks()` | UI wrapper for linking | Menu button |
| `reconnectNotesLinksInternal(ss)` | Actual linking logic | Called internally |
| `renameHeader()` | Rename labels across all blocks | Menu button |
| `manualBackup()` | Create spreadsheet copy | Menu button |

#### Data Extraction Functions

| Function | Purpose |
|----------|---------|
| `extractAllVendorData(sheet, templateHeight)` | Extract all vendor data into objects |
| `buildTemplateLabelMap(templateSheet, templateRange)` | Map label positions |

#### Data Injection Functions

| Function | Purpose |
|----------|---------|
| `rebuildVendorBlocks(...)` | Clear and rebuild all blocks |
| `injectDataByLabel(...)` | Write extracted data back |
| `injectVendorData_(...)` | Write single vendor data (batched) |
| `writeVendorBlock_(...)` | Write single block (batched) |
| `safeSetValue(sheet, row, col, value)` | Merged-cell-safe value setter |

#### Batched Update Functions

| Function | Purpose |
|----------|---------|
| `startUpdateAllVendorBlocksBatched()` | Initialize batch process |
| `processUpdateAllVendorBlocksBatch(isManual)` | Process next batch |
| `finalizeBatchedUpdate_(...)` | Cleanup after batching |
| `scheduleNextBatch_()` | Create time trigger for next batch |
| `removeTriggers_(handlerName)` | Delete triggers by function name |
| `getOrCreateUpdateCacheSheet_(ss)` | Manage cache sheet |

#### Utility Functions

| Function | Purpose |
|----------|---------|
| `getTemplateMaxColumn(sheet, lastRow)` | Get true template width (including merges) |
| `richTextToHyperlinkFormula(richText)` | Convert rich text to HYPERLINK formula |
| `backupSpreadsheet()` | Create timestamped backup |
| `showConfirmationDialog(vendorCount, templateHeight)` | Confirm destructive action |
| `showSummary(vendorCount, warnings, backupName)` | Display completion message |

---

### 3.4 ZoomLinks.js

**Purpose**: Manage Zoom Links sheet and integrate with Meeting Notes.

#### Configuration Object: `ZOOM_CONFIG`
```javascript
const ZOOM_CONFIG = {
  sheetName: "Zoom Links",
  blockSize: 12,       // 1 header + 1 subheader + 9 data + 1 blank
  dataRows: 9,
  vendorNameCol: 1,    // Column A
  personNameCol: 2,    // Column B
  meetingLinkCol: 3,   // Column C
  zoomLabel: "Zoom"
};
```

#### Functions

| Function | Purpose | Triggers |
|----------|---------|----------|
| `setupZoomLinksSheet()` | One-time setup for all vendors | (Manual/Removed from menu) |
| `createZoomVendorBlock_(sheet, startRow, vendorName)` | Create single block | Internal |
| `createZoomVendorBlockForNewVendor(vendorName)` | Create block for new vendor | `createNewVendor()` |
| `linkZoomCells(notesSheet, templateSheet)` | Sync names + link cells | `reconnectNotesLinksInternal()` |
| `renameSheet()` | Safe sheet renaming | Menu button |

#### `linkZoomCells()` Logic
1. Find "Name" and "Zoom" columns in template
2. Build vendor position map in Zoom Links
3. For each vendor in Meeting Notes:
   - Copy person names to Zoom Links "Person Name" column
   - Add Rich Text hyperlink to non-empty Zoom cells

---

### 3.5 SlackNotifications.js

**Purpose**: Daily automated Slack alerts for date-based events.

#### Configuration Object: `SLACK_CONFIG`
```javascript
const SLACK_CONFIG = {
  webhooks: {
    touches: "https://hooks.slack.com/services/...",
    followUp: "https://hooks.slack.com/services/...",
    timeline: "https://hooks.slack.com/services/..."
  },
  timezone: "GMT-5",
  triggerHour: 4,
  triggerMinute: 30
};
```

#### Functions

| Function | Purpose | Triggers |
|----------|---------|----------|
| `checkDatesAndNotify()` | Main scan function | Daily trigger |
| `mapNotificationLabels(sheet, height, width)` | Find label positions | Internal |
| `sendSlackAlert(vendor, label, type)` | POST to Slack webhook | Internal |
| `setupSlackTrigger()` | Create/reset daily trigger | Menu button |
| `testSlackNotifications()` | Manual test scan | Menu button |

#### Monitored Labels
| Label | Type | Webhook Channel |
|-------|------|-----------------|
| 1st Touch | touches | touches channel |
| 2nd Touch | touches | touches channel |
| 3rd Touch | touches | touches channel |
| 4th Touch | touches | touches channel |
| Follow Up Date | followUp | followUp channel |
| Timeline for consideration | timeline | timeline channel (offset: 2) |

#### Special Handling: "Timeline for consideration"
This label has an **empty cell** between the label and date value. The code uses `offset: 2` instead of the default `offset: 1`.

---

### 3.6 Reminders.js (Legacy)

**Purpose**: Separate reminder system for CRM and Account Notes sheets.

| Function | Sheet | Trigger |
|----------|-------|---------|
| `sendSlackReminders()` | CRM | `runAllSlackReminders()` |
| `sendTimelineReminders()` | Account Notes | `runAllSlackReminders()` |
| `runAllSlackReminders()` | Both | Time-based trigger |

> **Note**: This is a legacy system that runs separately from `SlackNotifications.js`.

---

## 4. Data Flow & Connections

### 4.1 Create New Vendor Flow
```
User clicks "Create New Vendor"
       │
       ▼
┌─────────────────────┐
│ createNewVendor()   │
└─────────────────────┘
       │
       ├──► CRM Sheet: Write vendor ID (Col A)
       │
       ├──► Meeting Notes: Write vendor header + template
       │
       └──► Zoom Links: Create 12-row block
              │
              └──► createZoomVendorBlockForNewVendor()
```

### 4.2 Update All Vendor Blocks Flow
```
User clicks "Update ALL Vendor Blocks"
       │
       ▼
┌─────────────────────────────┐
│ updateAllVendorBlocks()     │
└─────────────────────────────┘
       │
       ├──► backupSpreadsheet()
       │
       ├──► extractAllVendorData()
       │         │
       │         └──► Reads all vendor blocks from Meeting Notes
       │
       ├──► buildTemplateLabelMap()
       │         │
       │         └──► Maps labels to positions in template
       │
       ├──► rebuildVendorBlocks()
       │         │
       │         └──► Clears Meeting Notes, writes fresh blocks
       │
       ├──► injectDataByLabel()
       │         │
       │         └──► Writes extracted data back into new blocks
       │
       └──► showSummary()
              │
              └──► ⚠️ Reminds user to run Reconnect Notes Links
```

### 4.3 Reconnect Notes Links Flow
```
User clicks "Reconnect Notes Links"
       │
       ▼
┌─────────────────────────────┐
│ reconnectNotesLinks()       │
└─────────────────────────────┘
       │
       ▼
┌─────────────────────────────────┐
│ reconnectNotesLinksInternal()   │
└─────────────────────────────────┘
       │
       ├──► CRM → Meeting Notes links
       │         │
       │         └──► For each vendor in CRM, create HYPERLINK to Meeting Notes row
       │
       ├──► Meeting Notes → CRM links (reverse)
       │         │
       │         └──► For each vendor header, create link back to CRM row
       │
       └──► linkZoomCells()
              │
              ├──► Sync "Name" values to Zoom Links "Person Name"
              │
              └──► Add Rich Text hyperlinks to non-empty Zoom cells
```

### 4.4 Slack Notifications Flow (Daily)
```
Time Trigger (4:30 AM GMT-5)
       │
       ▼
┌─────────────────────────────┐
│ checkDatesAndNotify()       │
└─────────────────────────────┘
       │
       ├──► Get Template_LeftBlock
       │
       ├──► mapNotificationLabels()
       │         │
       │         └──► Find positions of 1st Touch, 2nd Touch, etc.
       │
       ├──► Loop through Meeting Notes vendors
       │         │
       │         └──► For each label position, check if date == TODAY
       │
       └──► sendSlackAlert()
              │
              └──► POST to appropriate Slack webhook
```

---

## 5. Configuration Reference

### 5.1 CONFIG (UpdateVendoBlock.js)
| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `notesSheetName` | string | "Meeting Notes" | Name of Meeting Notes sheet |
| `templateSheetName` | string | "Template_LeftBlock" | Name of template sheet |
| `blankRowsAfterVendor` | number | 1 | Blank rows between vendor blocks |
| `vendorHeaderBgColor` | string | "#ff00ff" | Magenta background for headers |
| `updateBatchSize` | number | 10 | Vendors per batch (batched update) |
| `preserveRichTextLinks` | boolean | false | Preserve rich text (slower) |

### 5.2 ZOOM_CONFIG (ZoomLinks.js)
| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `sheetName` | string | "Zoom Links" | Name of Zoom Links sheet |
| `blockSize` | number | 12 | Rows per vendor block |
| `dataRows` | number | 9 | Data rows for person/links |
| `vendorNameCol` | number | 1 | Column A |
| `personNameCol` | number | 2 | Column B |
| `meetingLinkCol` | number | 3 | Column C |
| `zoomLabel` | string | "Zoom" | Label to find in template |

### 5.3 SLACK_CONFIG (SlackNotifications.js)
| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `webhooks.touches` | string | URL | Webhook for touch notifications |
| `webhooks.followUp` | string | URL | Webhook for follow-up notifications |
| `webhooks.timeline` | string | URL | Webhook for timeline notifications |
| `timezone` | string | "GMT-5" | Timezone for trigger |
| `triggerHour` | number | 4 | Hour (0-23) for daily trigger |
| `triggerMinute` | number | 30 | Minute (0-59) for daily trigger |

---

## 6. User Guide

### 6.1 Creating a New Vendor
1. Go to the **CRM** sheet
2. Select any cell in the row of the vendor you want to add
3. Click `⚡ CRM Tools > Create New Vendor`
4. The system will:
   - Generate a unique vendor ID
   - Create a block in Meeting Notes
   - Create a block in Zoom Links
   - Show a success message

### 6.2 Updating Template Changes
When you modify `Template_LeftBlock`:

1. Click `⚡ CRM Tools > Update ALL Vendor Blocks`
2. Confirm the backup prompt
3. Wait for completion (~230 seconds for ~50 vendors)
4. **IMPORTANT**: Click `⚡ CRM Tools > Reconnect Notes Links` immediately after
5. This re-establishes all hyperlinks

### 6.3 Reconnecting Links
Use this after any major update or if links appear broken:
1. Click `⚡ CRM Tools > Reconnect Notes Links`
2. Wait for completion
3. Links in CRM and Meeting Notes are restored

### 6.4 Renaming Headers
To rename a label across all vendor blocks:
1. Click `⚡ CRM Tools > Rename Header`
2. Enter the current header name (exact match)
3. Enter the new header name
4. Confirm the change

### 6.5 Slack Notifications
**Setup (One-Time)**:
1. Click `⚡ CRM Tools > Slack Notifications > Setup Daily Trigger (4:30 AM)`
2. Confirm the trigger is set

**Test Manually**:
1. Add a date value matching TODAY to a monitored label (e.g., "1st Touch")
2. Click `⚡ CRM Tools > Slack Notifications > Run Test Scan`
3. Check Slack for the notification

### 6.6 Manual Backup
1. Click `⚡ CRM Tools > Manual Backup`
2. A copy is created in the same folder with timestamp

### 6.7 Renaming Sheets
**NEVER** rename sheets directly. Use:
1. Click `⚡ CRM Tools > Rename Sheet`
2. Select the sheet to rename
3. Enter the new name
4. Follow instructions to update code (if needed)

---

## 7. Constraints & Limitations

### 7.1 Execution Time Limits
| Limit | Value | Impact |
|-------|-------|--------|
| Max execution time | 6 minutes (360 seconds) | Large updates may timeout |
| Current typical run | ~230 seconds | Safe margin for ~50 vendors |

**Mitigation**: Update ALL and Reconnect Links are now **decoupled**. Run them separately.

### 7.2 Data Limits
| Limit | Value | Notes |
|-------|-------|-------|
| Max rows per sheet | 10,000,000 | Practical limit ~100,000 |
| Max cells per sheet | 10,000,000 | Shared across all rows/columns |
| Max triggers per user | 20 | Each function can have 1 trigger |

### 7.3 Trigger Ownership
- **Scripts only see triggers owned by the current user**
- If another collaborator created a trigger, you cannot delete it via code
- Solution: Manually delete old triggers before setting up new ones

### 7.4 Template Structure Requirements
- Template must start at cell A1
- Template must not have empty first row
- Column A in Meeting Notes is **reserved** for vendor names
- Template content starts at Column B in Meeting Notes

### 7.5 Merged Cells
- Merged cells in template are preserved
- `getTemplateMaxColumn()` accounts for merge extent
- `safeSetValue()` handles merged cell writes safely

---

## 8. ⚠️ Things to Avoid

### 8.1 NEVER Do These

| Action | Consequence | Solution |
|--------|-------------|----------|
| Rename sheets directly | Code will break | Use "Rename Sheet" button |
| Add columns in MIDDLE of template | Data shifts incorrectly | Add columns at END of sections |
| Delete Template_LeftBlock | Update will fail | Restore from backup |
| Run Update ALL on very large sheets | Timeout/corruption | Use batched update or split data |
| Create duplicate triggers | Multiple notifications | Delete old triggers first |

### 8.2 Template Modification Rules

**✅ CORRECT**: Add new column at END of section
```
| Name | Title | Email | LinkedIn | GEO | Zoom | NewField ← HERE |
```

**❌ WRONG**: Add new column in MIDDLE
```
| Name | NewField ← WRONG | Title | Email | LinkedIn | GEO | Zoom |
```

### 8.3 Link Preservation
- Running "Update ALL" will **break all hyperlinks**
- You **MUST** run "Reconnect Notes Links" after every update
- Failing to do so leaves CRM and Meeting Notes unlinked

### 8.4 Trigger Management
- Only one `checkDatesAndNotify` trigger should exist per user
- Delete old triggers before creating new ones
- Check `Extensions > Apps Script > Triggers` to verify

---

## 9. Troubleshooting

### 9.1 "Missing sheets" Error
**Cause**: Sheet name mismatch in CONFIG
**Solution**: 
1. Check CONFIG values match actual sheet names
2. Use "Rename Sheet" if you changed names

### 9.2 Hyperlinks Not Working
**Cause**: Links broken after update
**Solution**: Run `Reconnect Notes Links`

### 9.3 Slack Notifications Not Sending
**Cause 1**: Trigger not set up
**Solution**: Run `Setup Daily Trigger`

**Cause 2**: Wrong timezone causing trigger to fire at unexpected time
**Solution**: Verify `SLACK_CONFIG.timezone` matches your expectation

**Cause 3**: Date format mismatch
**Solution**: Ensure dates are actual Date objects, not text strings

### 9.4 Duplicate Triggers
**Cause**: Multiple users set up triggers, or old triggers not deleted
**Solution**: Go to `Extensions > Apps Script > Triggers` and delete duplicates

### 9.5 Execution Timeout
**Cause**: Too many vendors or slow network
**Solution**: 
1. Run update during low-usage hours
2. Consider batched update (hidden but still available in code)
3. Split data into multiple sheets

### 9.6 Zoom Links Not Syncing
**Cause 1**: Zoom Links sheet doesn't exist
**Solution**: Run `setupZoomLinksSheet()` manually from Apps Script editor

**Cause 2**: Vendor names don't match exactly
**Solution**: Ensure vendor names in Meeting Notes match Zoom Links exactly

---

## 10. Future Development Notes

### 10.1 Potential Improvements
1. **Auto-update CONFIG on sheet rename**: Store sheet names in Document Properties instead of hardcoded CONFIG
2. **Progress indicator**: Show real-time progress during long updates
3. **Incremental updates**: Only update changed vendors instead of all
4. **Error recovery**: Resume from last processed vendor after timeout

### 10.2 Code Architecture Principles
- **Separation of Concerns**: Each file handles one domain (Menu, Vendors, Zoom, Slack)
- **Internal vs External Functions**: `*Internal()` functions have no UI, used programmatically
- **Private Functions**: `*_()` suffix indicates internal/helper functions
- **Configuration Objects**: All constants in `CONFIG`, `ZOOM_CONFIG`, `SLACK_CONFIG`

### 10.3 Adding New Features
1. Create new `.js` file for the feature
2. Add configuration object if needed
3. Add menu item in `Menu.gs.js`
4. Document in this architecture file
5. Test thoroughly before pushing to production

### 10.4 Deployment
```bash
# Push to production
export PATH="$HOME/node-js/bin:$HOME/npm-global/bin:$PATH"
clasp push --force

# Switch script ID (edit .clasp.json)
{
  "scriptId": "YOUR_SCRIPT_ID_HERE"
}
```

---

## Document Revision History

| Date | Version | Changes |
|------|---------|---------|
| 2026-01-31 | 2.0 | Complete rewrite with full function documentation |
| 2026-01-29 | 1.0 | Initial architecture document |

---

*End of Architecture Document*
