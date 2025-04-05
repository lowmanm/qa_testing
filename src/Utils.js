// src/Utils.gs
// Shared helper functions for data manipulation and formatting

/**
 * Convert sheet data into an array of objects using headers from row 1
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object[]}
 */
function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

/**
 * Convert snake_case or underscore_case to Title Case
 * @param {string} str
 * @returns {string}
 */
function toTitleCase(str) {
  return (str || '')
    .split('_')
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ');
}

/**
 * Generate a unique ID with optional prefix
 * @param {string} prefix
 * @returns {string}
 */
function generateId(prefix = '') {
  return `${prefix}${Date.now()}`;
}
