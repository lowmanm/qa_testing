// ====================
// System Settings Module
// ====================

/**
 * Fetches system-level configuration from the "settings" sheet.
 */
function getSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('settings');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key) settings[key] = value;
  }

  return settings;
}

/**
 * Persists system-level configuration to the "settings" sheet.
 */
function saveSystemSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('settings');
  if (!sheet) {
    sheet = ss.insertSheet('settings');
    sheet.appendRow(['key', 'value']);
  }

  const keys = Object.keys(data);
  keys.forEach((key, idx) => {
    let rowIdx = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .findIndex(k => k === key);

    if (rowIdx === -1) {
      sheet.appendRow([key, data[key]]);
    } else {
      sheet.getRange(rowIdx + 2, 2).setValue(data[key]);
    }
  });

  return { success: true };
}
