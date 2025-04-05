// src/SpreadsheetSetup.gs
// Creates and ensures consistent structure for all required sheets

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToCreate = {
    users: ['id', 'name', 'email', 'managerEmail', 'role', 'createdBy', 'createdTimestamp'],
    auditQueue: [
      'auditId', 'taskId', 'referenceNumber', 'auditStatus', 'agentEmail',
      'requestType', 'taskType', 'outcome', 'taskTimestamp', 'auditTimestamp', 'locked'
    ],
    evalSummary: [
      'id', 'evalId', 'referenceNumber', 'taskType', 'outcome',
      'qaEmail', 'startTimestamp', 'stopTimestamp', 'totalPoints',
      'totalPointsPossible', 'status', 'feedback', 'evalScore'
    ],
    evalQuest: [
      'id', 'evalId', 'questionId', 'questionText', 'response',
      'pointsEarned', 'pointsPossible', 'feedback'
    ],
    questions: [
      'id', 'setId', 'taskType', 'questionText', 'pointsPossible',
      'createdBy', 'createdTimestamp'
    ],
    disputesQueue: [
      'id', 'evalId', 'userEmail', 'disputeTimestamp', 'reason',
      'questionIds', 'status', 'resolutionNotes', 'resolvedBy', 'resolutionTimestamp'
    ],
    settings: [
      'key', 'value'
    ]
  };

  for (const [sheetName, headers] of Object.entries(sheetsToCreate)) {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // Try to rename legacy sheets
      const legacyName = sheetName.charAt(0).toUpperCase() + sheetName.slice(1);
      const legacySheet = ss.getSheetByName(legacyName);
      sheet = legacySheet || ss.insertSheet(sheetName);
      if (legacySheet) legacySheet.setName(sheetName);
    }

    // Ensure correct headers
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existingHeaders.join(',') !== headers.join(',')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
}
