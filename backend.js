function getAllData() {
  return getMainData();
}

function getMainData() {
  try {
    const ss = SpreadsheetApp.openById("1DRteBSFT1cj4R_OUPMoDxeLMzAIJexWF3HPT-rpMOoM");
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return [];
    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    if (lastrow <= 1) return [];
    const data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
    return data.map(row => row.map(cell =>
      cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy") : cell
    ));
  } catch (e) {
    Logger.log("Lỗi getMainData: " + e.toString());
    return [];
  }
}

function getReleaseData() {
  try {
    const ss = SpreadsheetApp.openById("1t5PWyoJHrxElWP3QgmB16BEMIvEC015NHq0tsxu_TpE");
    const sheet = ss.getSheetByName("data");
    if (!sheet) return [];
    const lastrow = sheet.getLastRow();
    const lastcol = sheet.getLastColumn();
    if (lastrow <= 1) return [];
    const data = sheet.getRange(2, 1, lastrow - 1, lastcol).getValues();
    return data.map(row => row.map(cell =>
      cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy") : cell
    ));
  } catch (e) {
    Logger.log("Lỗi getReleaseData: " + e.toString());
    return [];
  }
}