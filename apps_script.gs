// Apps Script web app for storing/loading color mappings in a Google Sheet
// Fill the SHEET_ID constant with your spreadsheet ID before deploying.
const SHEET_ID = '11iOV4yU7yD13XAqeaZqhJZjHqfWcX5W11gZRsGwvNqg'; // e.g. '1AbC...'
const DEFAULT_SHEET_NAME = 'Colors';

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  return sh;
}

function doGet(e) {
  try {
    const sheetName = (e && e.parameter && e.parameter.sheet) || DEFAULT_SHEET_NAME;
    const sh = getSheet(sheetName);
    const values = sh.getDataRange().getValues();
    // Expect header row [name, hex, label]
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      rows.push({ name: r[0], hex: r[1], label: r[2] });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = payload.action || (e.parameter && e.parameter.action) || 'set';
    if (action !== 'set') {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const sheetName = payload.sheet || DEFAULT_SHEET_NAME;
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const sh = getSheet(sheetName);
    // clear and write header + rows
    sh.clearContents();
    sh.appendRow(['name', 'hex', 'label']);
    rows.forEach(r => {
      sh.appendRow([r.name || '', r.hex || '', r.label || '']);
    });
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', rowsWritten: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
