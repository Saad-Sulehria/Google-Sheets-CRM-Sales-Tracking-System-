# CRM Tools - User Guide

> **Version:** 0.1.2  
> **Last Updated:** January 28, 2026

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [Menu Features](#menu-features)
3. [Update ALL Vendor Blocks](#update-all-vendor-blocks)
4. [Reconnect Notes Links](#reconnect-notes-links)
5. [Create New Vendor](#create-new-vendor)
6. [Manual Backup](#manual-backup)
7. [⚠️ Critical Rules & Constraints](#critical-rules--constraints)
8. [Troubleshooting](#troubleshooting)
9. [Best Practices](#best-practices)

---

## Overview

The **CRM Tools** system manages vendor data across two key sheets:

| Sheet | Purpose |
|-------|---------|
| **CRM** | Master list of vendors with contact info and status |
| **Meeting Notes** | Detailed notes, checkboxes, and tracking per vendor |
| **Template_LeftBlock** | The template structure for each vendor block |

The system preserves your data when the template structure changes, allowing you to add rows/columns without losing existing information.

---

## Menu Features

Access all features from: **⚡ CRM Tools** menu

```
⚡ CRM Tools
├── Create New Vendor
├── Update ALL Vendor Blocks
├── Reconnect Notes Links
├── ─────────────────────
└── Manual Backup
```

---

## Update ALL Vendor Blocks

### What It Does
Rebuilds all vendor blocks in "Meeting Notes" using the current template structure, while **preserving all existing data**.

### When to Use
- After modifying the `Template_LeftBlock` (adding rows/columns)
- To apply formatting changes to all vendors
- To clean up corrupted blocks

### How to Use
1. **Modify your template** in `Template_LeftBlock` sheet
2. Click **⚡ CRM Tools → Update ALL Vendor Blocks**
3. Review the confirmation dialog (shows vendor count)
4. Click **Yes** to proceed
5. Wait for completion message

### What Happens Automatically
- ✅ Creates a timestamped backup
- ✅ Extracts all existing data
- ✅ Rebuilds blocks with new template
- ✅ Injects data back into matching positions
- ✅ Clears trailing rows (leftover checkboxes)
- ✅ Clears trailing columns (removed columns)
- ✅ Reconnects "View Notes" links in CRM sheet

---

## Reconnect Notes Links

### What It Does
Updates the "View Notes" hyperlinks in the CRM sheet (Column N) to point to the correct vendor rows in Meeting Notes.

### When to Use
- After manually reordering vendors
- If links become broken
- After adding/removing vendors manually

### How to Use
1. Click **⚡ CRM Tools → Reconnect Notes Links**
2. Wait for completion message showing linked count

> **Note:** This runs automatically after "Update ALL Vendor Blocks"

---

## Create New Vendor

### What It Does
Creates a new vendor entry in Meeting Notes with the correct template structure.

### How to Use
1. Go to the **CRM** sheet
2. Fill in the vendor name in **Column B** of a new row
3. Keep your cursor on that row
4. Click **⚡ CRM Tools → Create New Vendor**
5. A new vendor block appears in Meeting Notes

### What Happens Automatically
- ✅ Generates a unique Vendor ID (Column A)
- ✅ Creates vendor header in Meeting Notes
- ✅ Copies the template below the header

---

## Manual Backup

### What It Does
Creates a timestamped copy of the entire spreadsheet.

### How to Use
1. Click **⚡ CRM Tools → Manual Backup**
2. Find backup in Google Drive (same folder)

### Backup Naming
`Backup_[SpreadsheetName]_[YYYY-MM-DD_HH-MM-SS]`

---

## ⚠️ Critical Rules & Constraints

### 🚨 FOLLOW THESE RULES TO AVOID DATA LOSS

### Template Modification Rules

| ✅ SAFE | ⚠️ UNSAFE |
|---------|-----------|
| Add rows anywhere | - |
| Remove rows anywhere | - |
| Add columns **at the END** of sections | Add columns **in the MIDDLE** |
| Remove columns | - |
| Change formatting | - |
| Change cell formulas | - |

### Adding New Columns

**IMPORTANT: Always add new columns at the END of each section, not in the middle.**

**Correct:** Add "Phone" column AFTER "Zoom" in Contact Info
```
| Name | Title | Email | LinkedIn | GEO | Zoom | Phone ← NEW |
```

**Incorrect:** Insert "Phone" column BETWEEN Title and Email
```
| Name | Title | Phone ← WRONG | Email | LinkedIn | GEO | Zoom |
```

**Why?** Data is mapped by position. Inserting in the middle shifts all subsequent data.

### Vendor Name Matching

**WARNING: Vendor names must match EXACTLY between CRM and Meeting Notes.**

The system matches vendors by the name in:
- **CRM:** Column B
- **Meeting Notes:** Column A (bold header)

If names don't match:
- "View Notes" links won't work
- "Reconnect Notes Links" will show warnings

### Never Do These

| ❌ DON'T | Why |
|----------|-----|
| Delete Meeting Notes sheet during update | Data loss |
| Rename Template_LeftBlock sheet | Script can't find template |
| Run "Update ALL" while another update is running | Data corruption |
| Change vendor header formatting (bold/color) | Vendor detection fails |
| Add content to Column A (except vendor names) | Breaks vendor detection |

---

## Troubleshooting

### "Missing sheets" Error
**Cause:** Required sheet renamed or deleted.
**Fix:** Ensure these sheets exist:
- `CRM`
- `Meeting Notes`
- `Template_LeftBlock`

### "No vendors found" Error
**Cause:** No bold text in Column A of Meeting Notes.
**Fix:** Vendor headers must be bold with dark blue background.

### Data in Wrong Columns After Update
**Cause:** Column added in middle of template section.
**Fix:** 
1. Restore from backup
2. Add column at END of section instead
3. Run Update ALL again

### Leftover Checkboxes After Update
**Cause:** Should not happen in current version.
**Fix:** Run "Update ALL Vendor Blocks" again.

### "View Notes" Links Broken
**Cause:** Vendor names don't match between CRM and Meeting Notes.
**Fix:** 
1. Ensure names match exactly
2. Run "Reconnect Notes Links"

### Dates Showing Wrong Format
**Cause:** Date columns in template not formatted.
**Fix:**
1. Open `Template_LeftBlock`
2. Select date columns
3. Format → Number → Custom: `MM/dd/yy`
4. Run "Update ALL Vendor Blocks"

---

## Best Practices

### Before Major Changes
1. **Always run Manual Backup first**
2. Test changes on one vendor before running Update ALL

### Template Modifications
1. Make ALL template changes at once
2. Run Update ALL only once (not after each small change)
3. Review the confirmation dialog carefully

### Regular Maintenance
- Run "Reconnect Notes Links" weekly if adding vendors manually
- Create manual backups before bulk changes
- Review the backup folder periodically (old backups can be deleted)

### Data Entry
- Use consistent date format (system displays as MM/DD/YY)
- Don't merge cells manually in Meeting Notes
- Keep vendor names short and unique

---

## Quick Reference Card

| Action | Menu Item |
|--------|-----------|
| Update all vendor blocks | ⚡ CRM Tools → Update ALL Vendor Blocks |
| Fix broken links | ⚡ CRM Tools → Reconnect Notes Links |
| Add new vendor | ⚡ CRM Tools → Create New Vendor |
| Create backup | ⚡ CRM Tools → Manual Backup |

| Rule | Remember |
|------|----------|
| New columns | Add at END of section only |
| Vendor names | Must match CRM ↔ Meeting Notes |
| Before changes | Create Manual Backup |

---

*For technical issues, contact your administrator.*
