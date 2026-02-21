/**
 * Slack Notifications Module
 * 
 * Sends automated alerts to Slack when specific dates arrive.
 * Runs daily via time-driven trigger.
 */

const SLACK_CONFIG = {
    // Replace these with your actual Incoming Webhook URLs
    webhooks: {
        touches: "https://hooks.slack.com/services/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",
        followUp: "https://hooks.slack.com/services/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",
        timeline: "https://hooks.slack.com/services/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"
    },
    timezone: "GMT-5", // Adjusted to GMT-5 per request
    triggerHour: 4,    // 4 AM
    triggerMinute: 30  // 30 minutes
};

// ... (rest of the file remains unchanged until setupSlackTrigger) ...

/**
 * Creates the daily trigger
 * Run this function ONCE manually
 */
function setupSlackTrigger() {
    // Clear existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
        if (t.getHandlerFunction() === "checkDatesAndNotify") {
            ScriptApp.deleteTrigger(t);
        }
    }

    ScriptApp.newTrigger("checkDatesAndNotify")
        .timeBased()
        .everyDays(1)
        .inTimezone(SLACK_CONFIG.timezone)
        .atHour(SLACK_CONFIG.triggerHour)
        .nearMinute(SLACK_CONFIG.triggerMinute || 0)
        .create();

    SpreadsheetApp.getUi().alert(`✅ Daily Slack trigger set for ${SLACK_CONFIG.triggerHour}:${SLACK_CONFIG.triggerMinute} ${SLACK_CONFIG.timezone}`);
}

/**
 * Main function to check dates and send notifications
 * Can be run manually or via trigger
 */
function checkDatesAndNotify() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const notesSheet = ss.getSheetByName(CONFIG.notesSheetName);
    const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);

    if (!notesSheet || !templateSheet) {
        Logger.log("Missing sheets");
        return;
    }

    // 1. Determine Template Dimensions & Layout
    const templateHeight = templateSheet.getLastRow();
    const templateWidth = getTemplateMaxColumn(templateSheet, templateHeight); // Use helper from UpdateVendoBlock
    const blankRows = CONFIG.blankRowsAfterVendor || 0;
    const blockTotalRows = 1 + templateHeight + blankRows; // Header + Template + Blank

    // 2. Map Label Locations relative to the block
    // We need to find the specific rows and columns for our target dates
    const labelMap = mapNotificationLabels(templateSheet, templateHeight, templateWidth);

    if (Object.keys(labelMap).length === 0) {
        Logger.log("No value labels found in template");
        return;
    }

    Logger.log("Label Map: " + JSON.stringify(labelMap));

    // 3. Loop through all vendors
    const lastRow = notesSheet.getLastRow();
    let currentRow = 1;
    const today = new Date();
    today.setHours(0, 0, 0, 0); // normalize to midnight

    Logger.log("=== SLACK SCAN STARTED ===");
    Logger.log("Config: TemplateHeight=" + templateHeight + ", BlankRows=" + blankRows + ", BlockTotal=" + blockTotalRows);
    Logger.log("Scanning range: Row " + currentRow + " to " + lastRow);
    Logger.log("Target Date: " + today.toDateString());

    while (currentRow <= lastRow) {
        const vendorName = notesSheet.getRange(currentRow, 1).getValue();

        // Safety check: Ensure we are on a valid block (Vendor name should accept non-empty string)
        if (vendorName === "") {
            currentRow += blockTotalRows;
            continue;
        }

        Logger.log("Checking Vendor: " + vendorName + " at Row " + currentRow);

        // Check each mapped label for this vendor
        for (const [key, pos] of Object.entries(labelMap)) {
            // Calculate absolute position: 
            // Row = Current Header Row + Template Label Row relative to B1
            // Col = Value Column (Label Column + 1, assuming simple layout for now, or use mapped value col)

            // Note: Template starts at B2 (relative to Header A1). 
            // So if label is at Row 5 in Template sheet (which starts at Row 1), 
            // In the block (header at R1), the template starts at R2.
            // So Block Row = HeaderRow + TemplateRow

            const checkRow = currentRow + pos.row;
            const checkCol = pos.valueCol;

            const cell = notesSheet.getRange(checkRow, checkCol);
            const val = cell.getValue();

            if (val) {
                Logger.log(`  > Found value for ${key} at [${checkRow}, ${checkCol}]: "${val}" (Type: ${typeof val})`);

                if (val instanceof Date) {
                    const checkDate = new Date(val);
                    checkDate.setHours(0, 0, 0, 0);

                    Logger.log(`    > Comparing [${checkDate.toDateString()}] == Today [${today.toDateString()}]`);

                    if (checkDate.getTime() === today.getTime()) {
                        Logger.log("    >>> MATCH! Sending alert...");
                        sendSlackAlert(vendorName, key, pos.type);
                    }
                } else {
                    Logger.log("    > Not a Date object. formatted value: " + cell.getDisplayValue());
                    // Fallback: Try to parse string date if needed
                }
            }
        }

        currentRow += blockTotalRows;
    }
}

/**
 * Maps template labels to their relative positions
 * Returns objects with { row, valueCol, type }
 */
function mapNotificationLabels(sheet, height, width) {
    const map = {};
    const data = sheet.getRange(1, 1, height, width).getValues();

    // Define target labels to look for
    const targets = [
        { label: "1st Touch", type: "touches" },
        { label: "2nd Touch", type: "touches" },
        { label: "3rd Touch", type: "touches" },
        { label: "4th Touch", type: "touches" },
        { label: "Follow Up Date", type: "followUp" },
        { label: "Timeline for consideration", type: "timeline", offset: 2 } // Offset 2: Label | Empty | Value
    ];

    for (let r = 0; r < height; r++) {
        for (let c = 0; c < width; c++) {
            const cellValue = data[r][c];
            if (!cellValue) continue;

            const strVal = String(cellValue).trim();
            const target = targets.find(t => t.label === strVal);

            if (target) {
                // Determine offset (default 1 for next cell)
                const colOffset = target.offset || 1;

                map[target.label] = {
                    row: r + 1,      // 1-based row index in template
                    valueCol: c + 2 + colOffset, // TemplateCol(c) -> NotesLabel(c+2) -> Value(c+2+offset)
                    type: target.type
                };
            }
        }
    }
    return map;
}

/**
 * Sends POST request to Slack
 */
function sendSlackAlert(vendor, label, type) {
    const url = SLACK_CONFIG.webhooks[type];
    if (!url || url.includes("YOUR")) {
        Logger.log("Skipping Slack alert (URL not configured): " + label + " for " + vendor);
        return;
    }

    const payload = {
        text: `🔔 *${label}* has arrived for *${vendor}*`
    };

    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };

    try {
        UrlFetchApp.fetch(url, options);
        Logger.log("Sent Slack alert: " + label + " for " + vendor);
    } catch (e) {
        Logger.log("Error sending to Slack: " + e.toString());
    }
}

/**
 * Creates the daily trigger
 * Run this function ONCE manually
 */
function setupSlackTrigger() {
    // Clear existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (const t of triggers) {
        if (t.getHandlerFunction() === "checkDatesAndNotify") {
            ScriptApp.deleteTrigger(t);
        }
    }

    ScriptApp.newTrigger("checkDatesAndNotify")
        .timeBased()
        .everyDays(1)
        .inTimezone(SLACK_CONFIG.timezone)
        .atHour(SLACK_CONFIG.triggerHour)
        .nearMinute(SLACK_CONFIG.triggerMinute)
        .create();

    const hour = SLACK_CONFIG.triggerHour;
    const minute = SLACK_CONFIG.triggerMinute < 10 ? "0" + SLACK_CONFIG.triggerMinute : SLACK_CONFIG.triggerMinute;
    SpreadsheetApp.getUi().alert("✅ Daily Slack trigger set for " + hour + ":" + minute + " AM (" + SLACK_CONFIG.timezone + ")");
}

/**
 * Menu hook to test manually
 */
function testSlackNotifications() {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
        "Test Slack Notifications",
        "This will scan the current sheet for TODAY's dates and send alerts to the configured webhooks. Proceed?",
        ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
        checkDatesAndNotify();
        ui.alert("Scan complete. Check Logs (View > Logs) for details.");
    }
}
