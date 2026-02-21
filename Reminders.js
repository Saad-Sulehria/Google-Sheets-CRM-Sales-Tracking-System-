/*******************************
 * 1) FOLLOW UPS (CRM)
 *******************************/
function sendSlackReminders() {
  const WEBHOOK_URL_FOLLOWUPS = "https://hooks.slack.com/services/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";

  const sheet = SpreadsheetApp.getActive().getSheetByName("CRM");
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  for (let i = 1; i < data.length; i++) {
    const company = data[i][0];     // Column A
    const contact = data[i][1];     // Column B
    const stage = data[i][6];       // Column G
    const nextStep = data[i][8];    // Column I
    const actionDate = data[i][9];  // Column J

    if (!actionDate || stage === "Closed" || stage === "Not a Fit") continue;

    const due = new Date(actionDate);
    due.setHours(0,0,0,0);

    if (due.getTime() === today.getTime()) {
      const message = {
        text: `🔔 *Overdue Follow Ups*\n*${company}* (${contact})\nNext step: ${nextStep}`
      };

      UrlFetchApp.fetch(WEBHOOK_URL_FOLLOWUPS, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(message)
      });
    }
  }
}


/*******************************
 * 2) TIMELINE FOR CONSIDERATION (Account Notes)
 *******************************/
function sendTimelineReminders() {
  const WEBHOOK_URL_TIMELINE = "https://hooks.slack.com/services/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";

  const sheet = SpreadsheetApp.getActive().getSheetByName("Account Notes");
  const values = sheet.getDataRange().getDisplayValues();
  const raw = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0,0,0,0);

  let reminders = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    // buscar label en cualquier celda
    let found = -1;
    for (let c = 0; c < row.length; c++) {
      if (row[c] && String(row[c]).toLowerCase().includes("timeline for consideration")) {
        found = c;
        break;
      }
    }
    if (found === -1) continue;

    // buscar la fecha a la derecha
    let timelineDate = null;
    for (let c = found + 1; c < Math.min(found + 8, row.length); c++) {
      const rawVal = raw[r][c];
      const dispVal = values[r][c];

      if (rawVal instanceof Date) {
        timelineDate = rawVal;
        break;
      }

      if (dispVal && String(dispVal).match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
        timelineDate = new Date(dispVal);
        break;
      }
    }

    if (!timelineDate || isNaN(new Date(timelineDate).getTime())) continue;

    const d = new Date(timelineDate);
    d.setHours(0,0,0,0);

    if (d.getTime() !== today.getTime()) continue;

    // vendor hacia arriba col A
    let vendor = "";
    for (let up = r; up >= 0; up--) {
      const v = values[up][0];
      if (v && String(v).trim() !== "") {
        vendor = String(v).trim();
        break;
      }
    }

    reminders.push(`⏳ *Timeline due today:* ${vendor || "(Vendor not found)"}`);
  }

  if (reminders.length === 0) return;

  UrlFetchApp.fetch(WEBHOOK_URL_TIMELINE, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      text: `📌 *Timeline Reminders*\n\n${reminders.join("\n")}`
    })
  });
}


/*******************************
 * 3) RUN BOTH
 *******************************/
function runAllSlackReminders() {
  try {
    sendSlackReminders(); // CRM → followups
    Logger.log("✅ CRM reminders executed");
  } catch (e) {
    Logger.log("❌ CRM reminders failed: " + e);
  }

  try {
    sendTimelineReminders(); // Account Notes → timeline channel
    Logger.log("✅ Timeline reminders executed");
  } catch (e) {
    Logger.log("❌ Timeline reminders failed: " + e);
  }
}