// ====================
// Users Module
// ====================

/**
 * Fetch all users from the 'users' sheet.
 */
function getAllUsers() {
  return getCachedOrFetch('all_users', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Return the current user object based on email.
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const users = getAllUsers();
  const user = users.find(u => u.email === email);

  return user || {
    id: 'unknown',
    name: 'Unknown User',
    email,
    role: 'qa_analyst' // default fallback role
  };
}

/**
 * Create a new user in the 'users' sheet.
 */
function createUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');

  if (!userData.id) {
    userData.id = 'user_' + Date.now();
  }

  userData.createdTimestamp = new Date().toISOString();
  userData.createdBy = Session.getActiveUser().getEmail() || 'system';

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => userData[h] || '');
  sheet.appendRow(row);

  CacheService.getScriptCache().remove('all_users');
  return userData;
}

/**
 * Update existing user row.
 */
function updateUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === userData.id);
  if (rowIndex === -1) throw new Error(`User ID ${userData.id} not found`);

  headers.forEach((header, i) => {
    if (header in userData && header !== 'createdTimestamp' && header !== 'createdBy') {
      sheet.getRange(rowIndex + 1, i + 1).setValue(userData[header]);
    }
  });

  CacheService.getScriptCache().remove('all_users');
  return userData;
}

/**
 * Delete a user from the sheet.
 */
function deleteUser(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === userId);
  if (rowIndex === -1) throw new Error(`User ID ${userId} not found`);

  sheet.deleteRow(rowIndex + 1);
  CacheService.getScriptCache().remove('all_users');

  return { success: true, message: 'User deleted successfully' };
}
