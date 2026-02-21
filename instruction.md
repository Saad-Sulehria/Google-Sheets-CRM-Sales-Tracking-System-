# User Instructions & Edge Cases Guide

This document explains how the **Dynamic Schema Update Engine** works, what actions you can take, and importantly, what precise behaviors (cause & effect) you can expect.

---

## ⚡ Core Concepts

### 1. Vendor Identification (Critical)
The system identifies where a "Vendor Block" starts by looking for a very specific format in **Column A**:
-   **Text**: Vendor Name
-   **Formatting**: **Bold**, **Dark Blue Background (#1f4e79)**, **White Text**.

> **⚠️ Edge Case:** If you manually change the background color of a vendor header to white or remove the bold formatting, the system **will not recognize it as a vendor**.
> *   **Result**: That vendor block might be skipped or treated as part of the previous vendor's data.
> *   **Fix**: Use the "Create New Vendor" tool to ensure correct formatting, or copy format from another valid header.

### 2. Field Matching (Labels)
The system preserves your data by matching "Labels".
-   A **Label** is the text in Column A (e.g., "Meetings") or Column B (e.g., "KR Booked Meeting").
-   A **Value** is the data entered to the right of that label.

> **⚠️ Edge Case:** If you **rename a label** in the template (e.g., changing `Project Name` to `Project Title`), the system considers `Project Title` a *new, empty field*.
> *   **Result**: The data that was under `Project Name` will be orphaned (lost) because `Project Name` no longer exists in the template.
> *   **Warning**: The update summary will warn you: `Label 'Project Name' not found in new template`.

---

## 📋 Action -> Result Matrix

| Your Action | System Result | Note |
| :--- | :--- | :--- |
| **Add a new row** to Template | All vendor blocks grow by 1 row. | Existing data below the insertion point is shifted down correctly. |
| **Remove a row** from Template | All vendor blocks shrink by 1 row. | Data associated with the removed label is deleted. |
| **Rename a Label** in Template | A new empty field appears with the new name. | **DATA LOSS**: Old data associated with the old name is dropped. |
| **Reorder rows** in Template | Data follows the label to the new position. | **Safe**: Values move with their labels. |
| **Change Formatting** (colors/fonts) in Template | New format is applied to ALL blocks. | Data remains safe. |
| **Enter Data** in Column A (not header) | System treats it as a Label. | If it duplicates a header style, it might break the block detection. |

---

## 🚨 Known Edge Cases

### 1. "Orphaned" Data
*   **Scenario**: You have data in a row in "Meeting Notes" that DOES NOT have a corresponding label in Columns A/B (e.g., a loose note in Column C).
*   **Result**: Since the system extracts data *paired with labels*, any data sitting in a row without a label **will be lost** during an update.
*   **Best Practice**: Always ensure data is associated with a row that has a label in Col A or B.

### 2. Duplicate Labels
*   **Scenario**: You have two fields named "Date" in the same vendor block.
*   **Result**: The system may strictly overwrite the second "Date" with the value of the first, or treat them unpredictably.
*   **Best Practice**: Ensure all labels within a single block are unique (e.g., "Meeting Date", "Funding Date").

### 3. Merged Cells
*   **Scenario**: You unmerge a cell in the template that was holding data.
*   **Result**: The data will be placed in the top-left cell of the previously merged range.
*   **Best Practice**: Use `CRM Tools > Manual Backup` before changing merged cell structures significantly.

---

## 🛠 Troubleshooting

**"I ran an update and my vendor disappeared!"**
*   **Check**: Did the vendor header lose its Bold/Blue formatting?
*   **Fix**: Restore from Backup, fix the formatting, and run update again.

**"I see a warning: 'Label X not found' in the logs."**
*   **Cause**: You likely renamed label 'X' to something else in the template.
*   **Fix**: If you need that data, rename the label back, or manually migrate the data before updating.

**"Can I undo an update?"**
*   **Yes**: Go to `File > Version History` in Google Sheets, OR check your Google Drive for the timestamped Backup file created automatically before the update.
