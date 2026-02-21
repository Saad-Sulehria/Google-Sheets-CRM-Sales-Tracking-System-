function syncLeftBlockFromTemplate() {

  const ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const templateSheet = ss.getSheetByName("Template_LeftBlock");
  const targetSheet = ss.getSheetByName("Meeting Notes");

  const template = templateSheet.getRange("A1:A").getValues()
    .flat()
    .filter(x => x !== "");

  const data = targetSheet.getDataRange().getValues();

  for (let r = 0; r < data.length; r++) {

    // Vendor row (blue header)
    if (data[r][0] && r === 0 || (data[r][0] && data[r][0].toString().trim() !== "" && data[r][0] === data[r][0].toString().toUpperCase())) {

      let insertAt = r + 1;

      // Collect existing labels under vendor
      let existing = [];
      let rr = insertAt;
      while (rr < data.length && !data[rr][0]) {
        if (data[rr][1]) existing.push(data[rr][1]);
        rr++;
      }

      // Insert missing labels
      template.forEach(label => {
        if (!existing.includes(label)) {
          targetSheet.insertRowBefore(insertAt + 1);
          targetSheet.getRange(insertAt + 1, 2).setValue(label);
          insertAt++;
        }
      });
    }
  }
}