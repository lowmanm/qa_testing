// ====================
// Constants
// ====================

const SHEET_USERS = 'users';
const SHEET_AUDIT_QUEUE = 'auditQueue';
const SHEET_EVAL_SUMMARY = 'evalSummary';
const SHEET_EVAL_QUEST = 'evalQuest';
const SHEET_QUESTIONS = 'questions';
const SHEET_DISPUTES_QUEUE = 'disputesQueue';

const MENU_QA_SYSTEM = 'QA System';
const ITEM_OPEN_QA_APP = 'Open QA App';
const ITEM_SETUP_SPREADSHEET = 'Setup Spreadsheet';
const ITEM_IMPORT_DATA_FROM_EMAIL = 'Import Data from Email';

const CACHE_DURATION = 300; // seconds (5 minutes)

// ====================
// App Entry Points
// ====================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper function to include HTML files.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu(MENU_QA_SYSTEM)
    .addItem(ITEM_OPEN_QA_APP, 'openQaApp')
    .addItem(ITEM_SETUP_SPREADSHEET, 'setupSpreadsheet')
    .addItem(ITEM_IMPORT_DATA_FROM_EMAIL, 'importDataFromEmail')
    .addToUi();
}

/**
 * Opens the QA App as a modal dialog.
 */
function openQaApp() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Evaluation System');
}

// ====================
// Spreadsheet Setup
// ====================

/**
 * Sets up the spreadsheet by creating necessary sheets and headers.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToCreate = {
    [SHEET_USERS]: ['id', 'name', 'email', 'managerEmail', 'role', 'createdBy', 'createdTimestamp', 'avatarUrl'],
    [SHEET_AUDIT_QUEUE]: [
      'auditId', 'taskId', 'referenceNumber', 'auditStatus', 'agentEmail',
      'requestType', 'taskType', 'outcome', 'taskTimestamp', 'auditTimestamp',
      'lockedBy', 'lockedAt'
    ],

    [SHEET_EVAL_SUMMARY]: [
      'id', 'auditId', 'referenceNumber', 'taskType', 'outcome',
      'qaEmail', 'startTimestamp', 'stopTimestamp', 'totalPoints',
      'totalPointsPossible', 'status', 'feedback', 'evalScore'
    ],
    [SHEET_EVAL_QUEST]: [
      'id', 'evalId', 'questionId', 'questionText', 'response',
      'pointsEarned', 'pointsPossible', 'feedback'
    ],
    [SHEET_QUESTIONS]: [
      'id', 'sequenceId', 'setId', 'requestType', 'taskType',
      'questionText', 'pointsPossible', 'createdBy', 'createdTimestamp', 'active'
    ],
    [SHEET_DISPUTES_QUEUE]: [
      'id', 'evalId', 'userEmail', 'disputeTimestamp', 'reason',
      'questionIds', 'status', 'resolutionNotes', 'resolvedBy', 'resolutionTimestamp'
    ]
  };

  for (const [sheetName, headers] of Object.entries(sheetsToCreate)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      // Try to rename matching legacy sheet
      const legacySheet = ss.getSheetByName(sheetName.charAt(0).toUpperCase() + sheetName.slice(1));
      sheet = legacySheet || ss.insertSheet(sheetName);
      if (legacySheet) legacySheet.setName(sheetName);
    }

    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (JSON.stringify(existingHeaders) !== JSON.stringify(headers)) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
}

// ====================
// Caching Utilities
// ====================

/**
 * Generic caching wrapper to fetch data or retrieve from cache.
 */
function getCachedOrFetch(key, fetchFn) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);

  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      Logger.log(`✅ Cache hit for key: ${key}`);
      return parsed;
    } catch (e) {
      Logger.log(`❌ Error parsing cache for ${key}: ${e.message}`);
      // Proceed to fetch fresh
    }
  } else {
    Logger.log(`⚠️ Cache miss for key: ${key}`);
  }

  // Fetch fresh data
  let fresh;
  try {
    fresh = fetchFn();
  } catch (e) {
    Logger.log(`❌ fetchFn for key ${key} threw an error: ${e.message}`);
    return [];
  }

  if (!Array.isArray(fresh)) {
    Logger.log(`⚠️ Fetched data for ${key} is not an array. Returning empty list.`);
    return [];
  }

  try {
  const json = JSON.stringify(fresh);
  if (json.length < 90000) {
    cache.put(key, json, CACHE_DURATION);
    Logger.log(`✅ Cached fresh value for key: ${key}`);
  } else {
    Logger.log(`⚠️ Skipped caching for ${key}: data too large (${json.length} bytes)`);
  }
} catch (e) {
  Logger.log(`❌ Failed to process caching for ${key}: ${e.message}`);
}

  return fresh;
}

function clearCache(keys) {
  const cache = CacheService.getScriptCache();

  if (!keys) return;

  if (Array.isArray(keys)) {
    try {
      cache.removeAll(keys);
      Logger.log(`✅ Cleared cache keys: ${keys.join(', ')}`);
    } catch (e) {
      Logger.log(`❌ Error clearing multiple cache keys: ${e.message}`);
    }
  } else {
    try {
      cache.remove(keys);
      Logger.log(`✅ Cleared cache key: ${keys}`);
    } catch (e) {
      Logger.log(`❌ Error clearing cache key "${keys}": ${e.message}`);
    }
  }
}

function clearQaCaches() {
  clearCache(['all_disputes', 'all_evaluations', 'pending_audits','all_audits']);
}

// ====================
// Sheet Data Helpers
// ====================

/**
 * Converts sheet data to an array of objects with headers.
 */
function getSheetDataAsObjects(sheet) {
  if (!sheet) {
    Logger.log('❌ getSheetDataAsObjects: No sheet provided');
    return [];
  }

  const range = sheet.getDataRange();
  const data = range.getValues();

  Logger.log(`📄 getSheetDataAsObjects: Rows = ${data.length}, Columns = ${data[0]?.length || 0}`);

  if (!data.length || !Array.isArray(data[0])) {
    Logger.log('⚠️ Sheet data is empty or malformed.');
    return [];
  }

  const [headers, ...values] = data;

  if (!headers || headers.length === 0) {
    Logger.log('⚠️ getSheetDataAsObjects: No headers found in the first row.');
    return [];
  }

  const objects = values.map((row, rowIndex) => {
    return headers.reduce((obj, header, i) => {
      if (header) obj[header] = row[i];
      return obj;
    }, {});
  });

  Logger.log(`✅ Parsed ${objects.length} data rows into objects.`);
  return objects;
}

// ====================
// Users Module
// ====================

/**
 * Retrieves all users from the users sheet.
 */
function getAllUsers() {
  return getCachedOrFetch('all_users', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Retrieves the current user based on their email.
 */
function getCurrentUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    var users = getAllUsers();
    var user = users.find(function(user) {
      return user.email === email;
    });

    if (!user && users.length > 0) {
      user = users[0];
    }

    if (!user) {
      throw new Error('No matching user found and no users available for default.');
    }

    return user;
  } catch (e) {
    Logger.log(`Error in getCurrentUser: ${e.message}`);
    return {
      id: 'unknown',
      name: 'Unknown User',
      email: email || 'unknown',
      role: 'qa_analyst'
    };
  }
}

/**
 * Creates a new user and adds to the users sheet.
 */
function createUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);

  if (!userData.id) {
    userData.id = 'user' + Date.now();
  }

  userData.createdTimestamp = new Date().toISOString();
  userData.createdBy = userData.createdBy || Session.getActiveUser().getEmail() || 'system';

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => userData[header] || '');

  sheet.appendRow(row);
  clearCache('all_users');

  return userData;
}

/**
 * Updates an existing user in the users sheet.
 */
function updateUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
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

  clearCache('all_users');
  return userData;
}

/**
 * Deletes a user from the users sheet.
 */
function deleteUser(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === userId);
  if (rowIndex === -1) throw new Error(`User ID ${userId} not found`);

  sheet.deleteRow(rowIndex + 1);
  clearCache('all_users');

  return { success: true, message: 'User deleted successfully' };
}

// ====================
// Questions Module
// ====================

/**
 * Retrieves all questions from the questions sheet.
 */
function getAllQuestions() {
  return getCachedOrFetch('all_questions', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Marks an audit as misconfigured.
 */
function markAuditAsMisconfigured(auditId, requestType, taskType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  if (!sheet) throw new Error(`Sheet "${SHEET_AUDIT_QUEUE}" not found`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf('auditId');
  const requestCol = headers.indexOf('requestType');
  const taskCol = headers.indexOf('taskType');
  const statusCol = headers.indexOf('auditStatus');

  if ([idCol, requestCol, taskCol, statusCol].includes(-1)) {
    throw new Error('Missing one or more required columns (auditId, requestType, taskType, auditStatus)');
  }

  let updatedCount = 0;
  let foundAudit = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentId = row[idCol];
    const currentRequest = row[requestCol];
    const currentTask = row[taskCol];
    const currentStatus = row[statusCol];

    const isTargetRecord = currentId === auditId;
    const isMatchingType = currentRequest === requestType && currentTask === taskType;

    if (isTargetRecord) {
      foundAudit = true;
    }

    if ((isTargetRecord || isMatchingType) && currentStatus !== 'misconfigured') {
      sheet.getRange(i + 1, statusCol + 1).setValue('misconfigured');
      updatedCount++;
    }
  }

  if (!foundAudit) throw new Error(`Audit ID ${auditId} not found in ${SHEET_AUDIT_QUEUE}.`);

  Logger.log(`✅ Marked ${updatedCount} audits as misconfigured.`);
  return updatedCount;
}

// ====================
// Audit Queue Module
// ====================

/**
 * Retrieves all audits from the audit queue sheet.
 */
function getAllAudits() {
  return getCachedOrFetch('all_audits', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
    if (!sheet) {
      Logger.log('❌ Sheet not found: SHEET_AUDIT_QUEUE');
      return [];
    }

    const values = getSheetDataAsObjects(sheet);

    if (values.length > 1000) {
      Logger.log(`⚠️ Not caching all_audits because size = ${values.length}`);
      // Return directly without caching
      return values;
    }

    Logger.log(`✅ Loaded ${values.length} audits`);
    return values;
  });
}


/**
 * Retrieves pending audits from the audit queue sheet.
 */
function getPendingAudits() {
  Logger.log('📥 Fetching pending audits (from cache or fresh)...');

  return getCachedOrFetch('pending_audits', () => {
    let audits = [];
    let evaluations = [];

    try {
      audits = getAllAudits();
      if (!Array.isArray(audits)) {
        Logger.log('❌ getAllAudits returned non-array or null');
        audits = [];
      }
    } catch (e) {
      Logger.log(`❌ getAllAudits threw error: ${e.message}`);
    }

    try {
      evaluations = getAllEvaluations();
      if (!Array.isArray(evaluations)) {
        Logger.log('❌ getAllEvaluations returned non-array or null');
        evaluations = [];
      }
    } catch (e) {
      Logger.log(`❌ getAllEvaluations threw error: ${e.message}`);
    }

    const evaluatedIds = new Set(evaluations.map(e => e.evalId));

    const pending = audits.filter(a =>
      a.auditStatus?.toLowerCase() === 'pending' &&
      !evaluatedIds.has(a.auditId)
    );

    Logger.log(`✅ getPendingAudits: ${pending.length} audits returned.`);
    return pending;
  });
}

/**
 * Updates the status of an audit in the audit queue sheet.
 */
function updateAuditStatus(auditId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === auditId);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  clearCache('all_audits');
}


/**
 * Updates the status of an audit and locks it.
 */
function updateAuditStatusAndLock(auditId, status) {
  if (!auditId || typeof auditId !== 'string') {
    Logger.log(`❌ Skipping updateAuditStatusAndLock: Invalid auditId provided: ${auditId}`);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');
  const lockedByIdx = headers.indexOf('lockedBy');
  const lockedAtIdx = headers.indexOf('lockedAt');

  if ([idIdx, statusIdx, lockedByIdx, lockedAtIdx].some(idx => idx === -1)) {
    Logger.log('❌ One or more required columns (auditId, auditStatus, lockedBy, lockedAt) are missing.');
    return;
  }

  let updated = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === auditId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(status);
      sheet.getRange(i + 1, lockedByIdx + 1).setValue('');
      sheet.getRange(i + 1, lockedAtIdx + 1).setValue('');
      Logger.log(`✅ Audit ${auditId} status updated to '${status}' and lock cleared.`);
      updated = true;
      break;
    }
  }

  if (!updated) {
    Logger.log(`⚠️ Audit with ID ${auditId} not found in sheet.`);
  }

  clearCache(['all_audits', 'pending_audits']);
}

/*
Automatically unlocks audits stuck "In Process" for too long.
{time-driven trigger}
*/
function unlockStaleAudits() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const statusIdx = headers.indexOf('auditStatus');
  const lockedByIdx = headers.indexOf('lockedBy');
  const lockedAtIdx = headers.indexOf('lockedAt');

  const now = new Date();
  const heartbeatIntervalMinutes = 5;
  const graceBufferMinutes = 1;
  const thresholdMinutes = heartbeatIntervalMinutes + graceBufferMinutes;
  let unlockedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusIdx];
    const lockedBy = data[i][lockedByIdx];
    const lockedAtRaw = data[i][lockedAtIdx];

    if (status === 'In Process' && lockedBy && lockedAtRaw) {
      const lockedAt = new Date(lockedAtRaw);
      if (isNaN(lockedAt.getTime())) continue; // skip invalid dates

      const minutesLocked = (now - lockedAt) / 60000;

      if (minutesLocked > thresholdMinutes) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('pending');
        sheet.getRange(i + 1, lockedByIdx + 1).setValue('');
        sheet.getRange(i + 1, lockedAtIdx + 1).setValue('');
        unlockedCount++;
      }
    }
  }

  Logger.log(`✅ unlockStaleAudits: ${unlockedCount} stale audits unlocked.`);
  clearCache(['all_audits', 'pending_audits']);
}

function fullyUnlockAudit(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');
  const lockedByIdx = headers.indexOf('lockedBy');
  const lockedAtIdx = headers.indexOf('lockedAt');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === auditId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue('pending');
      sheet.getRange(i + 1, lockedByIdx + 1).setValue('');
      sheet.getRange(i + 1, lockedAtIdx + 1).setValue('');
      break;
    }
  }

  clearCache(['all_audits', 'pending_audits']);
}

function keepAuditLockAlive(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const lockedAtIdx = headers.indexOf('lockedAt');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === auditId) {
      sheet.getRange(i + 1, lockedAtIdx + 1).setValue(new Date().toISOString());
      break;
    }
  }
}

function checkIfAlreadyEvaluated(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const auditIdIdx = headers.indexOf('auditId');

  return data.some(row => row[auditIdIdx] === auditId);
}

/**
 * Prepares an evaluation by updating the audit status to 'In Process'.
 */
function prepareEvaluation(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx = headers.indexOf('auditId');
  const lockedByIdx = headers.indexOf('lockedBy');
  const lockedAtIdx = headers.indexOf('lockedAt');
  const statusIdx = headers.indexOf('auditStatus');

  const rowIndex = data.findIndex((row, i) => i > 0 && row[idIdx] === auditId);
  if (rowIndex === -1) throw new Error('Audit not found');

  const lockedBy = data[rowIndex][lockedByIdx];
  if (lockedBy && lockedBy !== '') {
    throw new Error('This audit is currently being evaluated by another user.');
  }

  const userEmail = Session.getActiveUser().getEmail();
  const now = new Date().toISOString();

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue('In Process');
  sheet.getRange(rowIndex + 1, lockedByIdx + 1).setValue(userEmail);
  sheet.getRange(rowIndex + 1, lockedAtIdx + 1).setValue(now);

  clearCache('all_audits');

  const audits = getAllAudits();
  return audits.find(a => a.auditId === auditId);
}

// ====================
// Evaluations Module
// ====================

/**
 * Retrieves all evaluations from the evaluation summary sheet.
 */
function getAllEvaluations() {
  return getCachedOrFetch('all_evaluations', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const evalSheet = ss.getSheetByName(SHEET_EVAL_SUMMARY);
    const questSheet = ss.getSheetByName(SHEET_EVAL_QUEST);

    const summaries = getSheetDataAsObjects(evalSheet);   // Main evals
    const questions = getSheetDataAsObjects(questSheet);  // Detailed questions

    // Group questions by evalId
    const map = {};
    questions.forEach(q => {
      if (!map[q.evalId]) map[q.evalId] = [];
      map[q.evalId].push({
        id: q.id,
        questionId: q.questionId,
        questionText: q.questionText,
        response: q.response,
        pointsEarned: q.pointsEarned,
        pointsPossible: q.pointsPossible,
        feedback: q.feedback
      });
    });

    // Attach grouped questions to each summary
    summaries.forEach(s => {
      s.questions = map[s.id] || [];
    });

    return summaries;
  });
}

/**
 * Saves a new evaluation and updates relevant sheets.
 */
function saveEvaluation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName(SHEET_EVAL_SUMMARY);
  const questSheet = ss.getSheetByName(SHEET_EVAL_QUEST);

  const evalId = 'eval' + Date.now();
  const stopTime = new Date().toISOString();
  const qaEmail = data.qaEmail || Session.getActiveUser().getEmail();

  // Step 1: Write to evalQuest first
  const questRows = data.questions.map((q, i) => [
    `${evalId}-q${i + 1}`,
    evalId,
    q.questionId,
    q.questionText,
    q.response,
    q.response === 'yes' ? q.pointsPossible : 0,
    q.pointsPossible,
    q.feedback || ''
  ]);

  if (questRows.length > 0) {
    questSheet.getRange(questSheet.getLastRow() + 1, 1, questRows.length, 8).setValues(questRows);
  }

  // Step 2: Recalculate totals dynamically
  const totalPoints = questRows.reduce((sum, row) => sum + (parseFloat(row[5]) || 0), 0);
  const totalPossible = questRows.reduce((sum, row) => sum + (parseFloat(row[6]) || 0), 0);
  const evalScore = totalPossible > 0 ? totalPoints / totalPossible : 0;

  // Step 3: Write to evalSummary
  evalSheet.appendRow([
    evalId,
    data.evalId || data.auditId,
    data.referenceNumber,
    data.taskType,
    data.outcome,
    qaEmail,
    data.startTimestamp || new Date().toISOString(),
    stopTime,
    totalPoints,
    totalPossible,
    'completed',
    data.feedback || '',
    evalScore
  ]);

  // Step 4: Mark audit as evaluated
  updateAuditStatus(data.evalId || data.auditId, 'evaluated');
  clearCache(['all_evaluations', 'all_audits', 'pending_audits']);

  // Step 5: Trigger email and return payload
  const evaluation = {
    id: evalId,
    ...data,
    stopTimestamp: stopTime,
    evalScore,
    totalPoints,
    totalPointsPossible: totalPossible,
    status: 'completed',
    qaEmail,
    questions: data.questions
  };

  sendEvaluationNotification(evaluation);
  return evaluation;
}

function updateEvaluationStatus(auditId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const auditIdIdx = headers.indexOf('auditId'); // ✅ updated header
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[auditIdIdx] === auditId);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  clearCache('all_evaluations');
}

function resetEvalStatusFromReviewing(evalId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  if (idIdx === -1 || statusIdx === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === evalId && data[i][statusIdx] === 'reviewing') {
      sheet.getRange(i + 1, statusIdx + 1).setValue('completed');
      break;
    }
  }
}

// ====================
// Disputes Module
// ====================

function getAllDisputes() {
  return getCachedOrFetch('all_disputes', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
    const data = getSheetDataAsObjects(sheet);

    data.forEach(d => {
      if (typeof d.questionIds === 'string') {
        d.questionIds = d.questionIds.split(',').map(q => q.trim()).filter(Boolean);
      } else {
        d.questionIds = [];
      }
    });

    // ✅ Improved logging to detect empty/missing questionIds
    Logger.log(`✅ Disputes Retrieved (${data.length}):`);
data.forEach((d, i) => {
  Logger.log(`Dispute #${i + 1}: ID=${d.id}, EvalID=${d.evalId}, questionIds=${Array.isArray(d.questionIds) ? d.questionIds.join(', ') : '[INVALID]'}`);
});


    return data;
  });
}

function checkAndSetDisputeReviewStatus(evalId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  if (idIdx === -1 || statusIdx === -1) return { success: false, error: "Missing columns" };

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === evalId) {
      const currentStatus = data[i][statusIdx].toLowerCase();
      if (currentStatus === 'disputed' || currentStatus === 'reviewing') {
        return { success: false, status: currentStatus };
      }

      // Otherwise, mark it as under review
      sheet.getRange(i + 1, statusIdx + 1).setValue('reviewing');
      return { success: true };
    }
  }

  return { success: false, error: "Evaluation not found" };
}

function checkIfAlreadyDisputed(evalId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const evalIdIdx = headers.indexOf('evalId');
  if (evalIdIdx === -1) return false;

  return data.some(row => row[evalIdIdx] === evalId);
}

function saveDispute(dispute) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
  const id = `dispute${Date.now()}`;
  const timestamp = new Date().toISOString();
  const userEmail = dispute.userEmail || Session.getActiveUser().getEmail();
  const status = dispute.status || 'pending';

  // Defensive: Ensure evalId exists
  if (!dispute.evalId) {
    throw new Error("Missing evaluation ID for dispute.");
  }

  const questionIds = (dispute.questionIds || []).join(',');

  // Save to Disputes Sheet
  sheet.appendRow([
    id,
    dispute.evalId,     // corresponds to evalSummary.id
    userEmail,
    timestamp,
    dispute.reason,
    questionIds,
    status
  ]);

  // Update evalSummary status via evalId (which is actually the 'id' field in evalSummary)
  updateEvaluationStatus(dispute.evalId, 'disputed');

  clearCache(['all_disputes', 'all_evaluations']);

  return {
    id,
    evalId: dispute.evalId,
    userEmail,
    disputeTimestamp: timestamp,
    reason: dispute.reason,
    questionIds: dispute.questionIds,
    status
  };
}

function checkDisputeReviewStatus(disputeId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
    if (!sheet) return { success: false, reason: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, reason: 'No data available' };

    const headers = data[0];
    const idIdx = headers.indexOf('id');
    const statusIdx = headers.indexOf('status');

    if (idIdx === -1 || statusIdx === -1) {
      return { success: false, reason: 'Missing required columns: id or status' };
    }

    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][idIdx];
      if (rowId === disputeId) {
        const status = String(data[i][statusIdx] || '').toLowerCase();
        const isLocked = status === 'reviewing';
        const isResolved = ['overturned', 'upheld', 'partial overturn'].includes(status);

        return { 
          success: !isLocked && !isResolved, // Only allow proceed if it's not locked or resolved
          status: status
        };
      }
    }

    return { success: false, reason: 'Dispute not found' };

  } catch (err) {
    console.error('checkDisputeReviewStatus failed:', err);
    return { success: false, reason: 'Unexpected error occurred' };
  }
}

/**
 * Updates the status of a dispute in the disputesQueue sheet.
 */
function updateDisputeStatus(disputeId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === disputeId);
  if (rowIndex === -1) {
    Logger.log(`❌ Dispute ID ${disputeId} not found.`);
    return { success: false, message: 'Dispute not found' };
  }

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  clearCache('all_disputes');

  Logger.log(`✅ Dispute ${disputeId} status updated to "${newStatus}"`);
  return { success: true };
}

function updateEvaluationStatus(evalId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  if (idIdx === -1 || statusIdx === -1) {
    Logger.log('❌ Missing expected headers in evalSummary');
    return;
  }

  const rowIndex = data.findIndex((row, i) => i > 0 && row[idIdx] === evalId);

  if (rowIndex === -1) {
    Logger.log(`⚠️ Evaluation with ID ${evalId} not found in evalSummary.`);
    return;
  }

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  clearCache('all_evaluations');
  Logger.log(`✅ Evaluation status updated: ID ${evalId} → ${newStatus}`);
}


/**
 * Resolves a dispute and updates relevant sheets.
 */
function resolveDispute(resolution) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const questSheet = ss.getSheetByName(SHEET_EVAL_QUEST);
  const summarySheet = ss.getSheetByName(SHEET_EVAL_SUMMARY);
  const disputeSheet = ss.getSheetByName(SHEET_DISPUTES_QUEUE);

  const { disputeId, evalId, decisions, resolutionNotes, status } = resolution;
  const resolvedBy = Session.getActiveUser().getEmail();
  const resolvedAt = new Date().toISOString();

  // 1. Update evalQuest rows
  const questData = questSheet.getDataRange().getValues();
  const headers = questData[0];
  const evalIdx = headers.indexOf('evalId');
  const qIdIdx = headers.indexOf('questionId');
  const respIdx = headers.indexOf('response');
  const ptsEarnedIdx = headers.indexOf('pointsEarned');
  const ptsPossibleIdx = headers.indexOf('pointsPossible');
  const feedbackIdx = headers.indexOf('feedback');

  for (let i = 1; i < questData.length; i++) {
    const row = questData[i];
    if (row[evalIdx] !== evalId) continue;

    const match = decisions.find(d => d.questionId === row[qIdIdx]);
    if (!match) continue;

    if (match.resolution === 'overturned') {
      row[respIdx] = 'yes';
      row[ptsEarnedIdx] = row[ptsPossibleIdx];
    } else {
      row[respIdx] = 'no';
      row[ptsEarnedIdx] = 0;
    }

    row[feedbackIdx] = match.note || '';
    questSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
  }

  // 2. Recalculate totals based on updated evalQuest
  const updatedQuestData = questSheet.getDataRange().getValues();
  const relevant = updatedQuestData.filter(r => r[evalIdx] === evalId);
  const totalPoints = relevant.reduce((sum, row) => sum + (parseFloat(row[ptsEarnedIdx]) || 0), 0);
  const totalPossible = relevant.reduce((sum, row) => sum + (parseFloat(row[ptsPossibleIdx]) || 0), 0);
  const evalScore = totalPossible ? totalPoints / totalPossible : 0;

  // 3. Update evalSummary with recalculated score and updated status
  const summaryData = summarySheet.getDataRange().getValues();
  const sHeaders = summaryData[0];
  const idIdx = sHeaders.indexOf('id');
  const ptsIdx = sHeaders.indexOf('totalPoints');
  const scoreIdx = sHeaders.indexOf('evalScore');
  const statusIdx = sHeaders.indexOf('status');

  for (let i = 1; i < summaryData.length; i++) {
    const row = summaryData[i];
    if (row[idIdx] === evalId) {
      summarySheet.getRange(i + 1, ptsIdx + 1).setValue(totalPoints);
      summarySheet.getRange(i + 1, scoreIdx + 1).setValue(evalScore);
      summarySheet.getRange(i + 1, statusIdx + 1).setValue(status || 'resolved');
      break;
    }
  }

  // 4. Update disputesQueue
  const disputeData = disputeSheet.getDataRange().getValues();
  const dHeaders = disputeData[0];
  const dIdIdx = dHeaders.indexOf('id');
  const dStatusIdx = dHeaders.indexOf('status');
  const dNotesIdx = dHeaders.indexOf('resolutionNotes');
  const dByIdx = dHeaders.indexOf('resolvedBy');
  const dTimeIdx = dHeaders.indexOf('resolutionTimestamp');

  for (let i = 1; i < disputeData.length; i++) {
    const row = disputeData[i];
    if (row[dIdIdx] === disputeId) {
      if (dStatusIdx !== -1) disputeSheet.getRange(i + 1, dStatusIdx + 1).setValue(status || 'resolved');
      if (dNotesIdx !== -1) disputeSheet.getRange(i + 1, dNotesIdx + 1).setValue(resolutionNotes || '');
      if (dByIdx !== -1) disputeSheet.getRange(i + 1, dByIdx + 1).setValue(resolvedBy);
      if (dTimeIdx !== -1) disputeSheet.getRange(i + 1, dTimeIdx + 1).setValue(resolvedAt);
      break;
    }
  }

  clearCache(['all_disputes', 'all_evaluations']);
  return true;
}

function getAllEvaluationsAndDisputes() {
  const evaluations = getAllEvaluations(); // already uses getCachedOrFetch inside
  const disputes = getAllDisputes();       // already uses getCachedOrFetch inside

  return { evaluations, disputes };
}

/* THIS WAS TO TEST BUT SEEMS LIKE DOUBLE CACHEING
function getAllEvaluationsAndDisputes() {
  const evaluations = getCachedOrFetch('all_evaluations', getAllEvaluations);
  const disputes = getCachedOrFetch('all_disputes', getAllDisputes);

  Logger.log("✅ getAllEvaluationsAndDisputes returned:");
  Logger.log('✅ Disputes:', JSON.stringify(disputes, null, 2));
  Logger.log('✅ Evaluations:', JSON.stringify(evaluations, null, 2));


  return {
    evaluations,
    disputes
  };
}
*/

function getDisputeStats() {
  const data = getAllDisputes();

  let total = 0;
  let partialOverturns = 0;
  let overturned = 0;
  let upheld = 0;

  data.forEach(row => {
    total++;
    const s = (row.status || '').toLowerCase();
    if (s === 'partial overturn') partialOverturns++;
    else if (s === 'overturned') overturned++;
    else if (s === 'upheld') upheld++;
  });

  return {
    total,
    partialOverturns,
    totalOverturned: overturned,
    disputesUpheld: upheld
  };
}

// Convert snake_case to Title Case for display
function toTitleCase(str) {
  return (str || '').split('_').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

function sendEvaluationNotification(evaluation) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const auditsSheet = ss.getSheetByName(SHEET_AUDIT_QUEUE);

  const usersData = usersSheet.getDataRange().getValues();
  const userHeaders = usersData[0];
  const emailIdx = userHeaders.indexOf('email');
  const managerIdx = userHeaders.indexOf('managerEmail');

  const agentEmail = getAuditField(auditsSheet, evaluation.auditId, 'agentEmail');
  const agentRow = usersData.find(row => row[emailIdx] === agentEmail);
  const managerEmail = agentRow ? agentRow[managerIdx] : '';

  const html = buildScorecardHtml(evaluation);
  const stopTime = new Date(evaluation.stopTimestamp).toLocaleString();
  const subject = `Evaluation has been completed for ${evaluation.referenceNumber || 'N/A'} at ${stopTime}`;

  GmailApp.sendEmail(agentEmail, subject, '', {
    htmlBody: html,
    cc: managerEmail || '',
    name: 'QA Team',
    replyTo: 'qa-team@equifax.com',
    noReply: true
  });
}

function buildScorecardHtml(evaluation) {
  const scorePercentage = Math.round((evaluation.totalPoints / evaluation.totalPointsPossible) * 100);
  const dateStr = new Date(evaluation.stopTimestamp || evaluation.startTimestamp).toLocaleDateString();

  const headerStyle = "font-weight: bold; color: #333;";
  const labelStyle = "font-size: 13px; color: #777;";
  const valueStyle = "font-size: 14px; color: #222; font-weight: 500;";
  const tableHeadStyle = "background-color:#f5f5f5; text-align:left; font-size:13px; border-bottom:1px solid #ddd;";
  const highlightRed = "background-color:#fff0f0;";

  let questionHtml = '';
  evaluation.questions.forEach(q => {
    const highlight = q.response === 'no' ? highlightRed : '';
    questionHtml += `
      <tr style="${highlight}">
        <td style="padding:10px; border-bottom:1px solid #eee;">${q.questionText}</td>
        <td style="padding:10px; text-align:center; border-bottom:1px solid #eee;">${q.response.toUpperCase()}</td>
        <td style="padding:10px; text-align:center; border-bottom:1px solid #eee;">${q.pointsEarned}/${q.pointsPossible}</td>
        <td style="padding:10px; border-bottom:1px solid #eee;">${q.feedback || ''}</td>
      </tr>
    `;
  });

  return `
    <div style="font-family:Arial, sans-serif; max-width:800px; margin:0 auto; padding:20px; color:#333;">
      <div style="border-radius:10px; padding:20px; background-color:#f8faff; border:1px solid #d0e6f9;">
        <h2 style="color:#007298; margin-top:0;">Evaluation Summary</h2>
        <table style="width:100%; margin-top:10px;">
          <tr><td style="${labelStyle}">Reference Number:</td><td style="${valueStyle}">${evaluation.referenceNumber || 'N/A'}</td></tr>
          <tr><td style="${labelStyle}">Task Type:</td><td style="${valueStyle}">${evaluation.taskType || 'N/A'}</td></tr>
          <tr><td style="${labelStyle}">Outcome:</td><td style="${valueStyle}">${evaluation.outcome || 'N/A'}</td></tr>
          <tr><td style="${labelStyle}">Score:</td><td style="${valueStyle}">${evaluation.totalPoints}/${evaluation.totalPointsPossible} (${scorePercentage}%)</td></tr>
          <tr><td style="${labelStyle}">Evaluator:</td><td style="${valueStyle}">${evaluation.qaEmail || 'QA Team'}</td></tr>
          <tr><td style="${labelStyle}">Date:</td><td style="${valueStyle}">${dateStr}</td></tr>
        </table>
      </div>

      <h3 style="color:#007298; margin-top:30px;">Evaluation Details</h3>
      <table style="width:100%; border-collapse:collapse; border:1px solid #ddd; margin-top:10px;">
        <thead>
          <tr style="${tableHeadStyle}">
            <th style="padding:10px;">Question</th>
            <th style="padding:10px; text-align:center;">Response</th>
            <th style="padding:10px; text-align:center;">Score</th>
            <th style="padding:10px;">Feedback</th>
          </tr>
        </thead>
        <tbody>
          ${questionHtml}
        </tbody>
      </table>

      ${evaluation.feedback ? `
        <div style="margin-top:25px;">
          <h4 style="color:#007298;">Overall Feedback</h4>
          <p style="background-color:#f9f9f9; padding:12px; border-left:4px solid #007298; border-radius:4px;">
            ${evaluation.feedback}
          </p>
        </div>` : ''}
    </div>
  `;
}

function getAuditField(sheet, auditId, columnName) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const colIdx = headers.indexOf(columnName);

  const row = data.find((r, i) => i > 0 && r[idIdx] === auditId);
  return row ? row[colIdx] : '';
}

function getUniqueTaskTypes() {
  return getUniqueColumnValues(SHEET_AUDIT_QUEUE, 'taskType');
}

function getUniqueRequestTypes() {
  return getUniqueColumnValues(SHEET_AUDIT_QUEUE, 'requestType');
}

function getUniqueColumnValues(sheetName, columnName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = headers.indexOf(columnName);

  return data
    .slice(1)
    .map(row => row[colIdx])
    .filter(v => v)
    .filter((v, i, arr) => arr.indexOf(v) === i)
    .sort();
}

function saveQuestion(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const headers = sheet.getDataRange().getValues()[0];
  const existingData = sheet.getDataRange().getValues();
  const idIndex = headers.indexOf('id');

  if (data.id) {
    // Update existing question
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][idIndex] === data.id) {
        const row = i + 1;
        sheet.getRange(row, headers.indexOf('sequenceId') + 1).setValue(data.sequenceId);
        sheet.getRange(row, headers.indexOf('requestType') + 1).setValue(data.requestType);
        sheet.getRange(row, headers.indexOf('taskType') + 1).setValue(data.taskType);
        sheet.getRange(row, headers.indexOf('questionText') + 1).setValue(data.questionText);
        sheet.getRange(row, headers.indexOf('pointsPossible') + 1).setValue(data.pointsPossible);
        return true;
      }
    }
  } else {
    // Create new question
    const newId = 'q_' + new Date().getTime();
    const setId = `${data.requestType}_${data.taskType}`;
    const createdBy = Session.getActiveUser().getEmail();
    const createdTimestamp = new Date().toISOString();
    const active = true;

    const newRow = [];
    headers.forEach(header => {
      switch (header) {
        case 'id': newRow.push(newId); break;
        case 'sequenceId': newRow.push(data.sequenceId); break;
        case 'requestType': newRow.push(data.requestType); break;
        case 'taskType': newRow.push(data.taskType); break;
        case 'setId': newRow.push(setId); break;
        case 'questionText': newRow.push(data.questionText); break;
        case 'pointsPossible': newRow.push(data.pointsPossible); break;
        case 'createdBy': newRow.push(createdBy); break;
        case 'createdTimestamp': newRow.push(createdTimestamp); break;
        case 'active': newRow.push(active); break;
        default: newRow.push('');
      }
    });

    sheet.appendRow(newRow);
    clearCache('all_questions');
    return true;
  }
}

function getQuestionsBySet(requestType, taskType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const questions = data.slice(1)
    .filter(row => row[headers.indexOf('active')])
    .filter(row =>
      row[headers.indexOf('requestType')] === requestType &&
      row[headers.indexOf('taskType')] === taskType
    )
    .sort((a, b) => a[headers.indexOf('sequenceId')] - b[headers.indexOf('sequenceId')])
    .map(row => ({
      id: row[headers.indexOf('id')],
      sequenceId: row[headers.indexOf('sequenceId')],
      setId: row[headers.indexOf('setId')],
      requestType: row[headers.indexOf('requestType')],
      taskType: row[headers.indexOf('taskType')],
      questionText: row[headers.indexOf('questionText')],
      pointsPossible: row[headers.indexOf('pointsPossible')],
      active: row[headers.indexOf('active')]
    }));

  return questions;
}

function getQuestionById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');

  const row = data.find((r, i) => i > 0 && r[idIdx] === id);
  if (!row) return null;

  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function toggleQuestionActive(id, isActive) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const activeIdx = headers.indexOf('active');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === id) {
      sheet.getRange(i + 1, activeIdx + 1).setValue(isActive);
      clearCache('all_questions');
      return true;
    }
  }

  return false; // not found
}

function getUniqueRequestTypes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const reqIndex = headers.indexOf('requestType');

  const values = data.slice(1).map(row => row[reqIndex]);
  const unique = [...new Set(values)].filter(v => v);
  return unique.sort();
}

function getUniqueTaskTypes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const taskIndex = headers.indexOf('taskType');

  const values = data.slice(1).map(row => row[taskIndex]);
  const unique = [...new Set(values)].filter(v => v);
  return unique.sort();
}

function recalculateEvaluationScores(evalId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const questSheet = ss.getSheetByName(SHEET_EVAL_QUEST);
  const summarySheet = ss.getSheetByName(SHEET_EVAL_SUMMARY);

  const questData = questSheet.getDataRange().getValues();
  const qHeaders = questData[0];
  const evalIdx = qHeaders.indexOf('evalId');
  const respIdx = qHeaders.indexOf('response');
  const ptsEarnedIdx = qHeaders.indexOf('pointsEarned');
  const ptsPossibleIdx = qHeaders.indexOf('pointsPossible');

  let totalPoints = 0;
  let totalPossible = 0;

  for (let i = 1; i < questData.length; i++) {
    const row = questData[i];
    if (row[evalIdx] !== evalId) continue;

    const response = String(row[respIdx] || '').toLowerCase();
    const possible = parseFloat(row[ptsPossibleIdx]) || 0;
    const earned = response === 'yes' ? possible : 0;

    row[ptsEarnedIdx] = earned;
    totalPoints += earned;
    totalPossible += possible;

    // Update the row in the sheet
    questSheet.getRange(i + 1, ptsEarnedIdx + 1).setValue(earned);
  }

  const evalScore = totalPossible > 0 ? totalPoints / totalPossible : 0;

  // Update summary
  const summaryData = summarySheet.getDataRange().getValues();
  const sHeaders = summaryData[0];
  const idIdx = sHeaders.indexOf('id');
  const scoreIdx = sHeaders.indexOf('evalScore');
  const ptsIdx = sHeaders.indexOf('totalPoints');

  for (let i = 1; i < summaryData.length; i++) {
    if (summaryData[i][idIdx] === evalId) {
      summarySheet.getRange(i + 1, ptsIdx + 1).setValue(totalPoints);
      summarySheet.getRange(i + 1, scoreIdx + 1).setValue(evalScore);
      break;
    }
  }

  return { totalPoints, totalPossible, evalScore };
}
