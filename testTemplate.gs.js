function debugTemplateRead() {
  const ss = SpreadsheetApp.getActive();
  const template = ss.getSheetByName("Template_LeftBlock");

  const values = template.getDataRange().getValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      if (String(values[r][c]).toUpperCase().includes("TEST")) {
        Logger.log(`FOUND TEST at row ${r+1}, col ${c+1}`);
      }
    }
  }

  Logger.log("DONE SCANNING");
}