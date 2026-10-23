// Google Sheet > Extensions > Apps Script > paste this > Deploy > New deployment > Web app
// Execute as: Me | Who has access: Anyone > Deploy > copy the /exec URL into the website (⚙︎ button)
const TABS = ["Customers", "Riders", "Leave"];

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), out = {};
  TABS.forEach(t => {
    const v = ss.getSheetByName(t).getDataRange().getValues();
    out[t] = v.slice(1).filter(r => r.join("") !== "").map(r => r.map(x => String(x)));
  });
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), d = JSON.parse(e.postData.contents);
  TABS.forEach(t => {
    const sh = ss.getSheetByName(t), cols = sh.getLastColumn(), last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, cols).clearContent();
    const rows = (d[t] || []).map(r => { r = r.slice(0, cols); while (r.length < cols) r.push(""); return r; });
    if (rows.length) sh.getRange(2, 1, rows.length, cols).setNumberFormat("@").setValues(rows);
  });
  return ContentService.createTextOutput("ok");
}
