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
      'id', 'evalId', 'referenceNumber', 'taskType', 'outcome',
      'qaEmail', 'startTimestamp', 'stopTimestamp', 'totalPoints',
      'totalPointsPossible', 'status', 'feedback', 'evalScore'
    ],
    [SHEET_EVAL_QUEST]: [
      'id', 'evalId', 'questionId', 'questionText', 'response',
      'pointsEarned', 'pointsPossible', 'feedback'
    ],
    [SHEET_QUESTIONS]: [
      'id', 'setId', 'taskType', 'questionText', 'pointsPossible',
      'createdBy', 'createdTimestamp'
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
      return JSON.parse(cached);
    } catch (e) {
      Logger.log(`Error parsing cache for ${key}: ${e.message}`);
    }
  }

  const fresh = fetchFn();
  try {
    cache.put(key, JSON.stringify(fresh), CACHE_DURATION);
  } catch (e) {
    Logger.log(`Failed to cache ${key}: ${e.message}`);
  }

  return fresh;
}

// ====================
// Sheet Data Helpers
// ====================

/**
 * Converts sheet data to an array of objects with headers.
 */
function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];

  const [headers, ...values] = sheet.getDataRange().getValues();
  return values.map(row => {
    return headers.reduce((obj, header, i) => {
      if (header) obj[header] = row[i];
      return obj;
    }, {});
  });
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
  CacheService.getScriptCache().remove('all_users');

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

  CacheService.getScriptCache().remove('all_users');
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
  CacheService.getScriptCache().remove('all_users');

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
 * Retrieves questions for a specific task type.
 */
function getQuestionsForTaskType(taskType) {
  const cacheKey = 'questions_' + taskType;
  return getCachedOrFetch(cacheKey, () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
    const data = getSheetDataAsObjects(sheet);
    return data.filter(q => q.taskType === taskType);
  });
}

/**
 * Marks an audit as misconfigured.
 */
function markAuditAsMisconfigured(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  if (!sheet) throw new Error(`Sheet "${SHEET_AUDIT_QUEUE}" not found`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf('auditId');
  const statusCol = headers.indexOf('auditStatus');

  if (idCol === -1 || statusCol === -1) {
    throw new Error('Missing auditId or auditStatus column.');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === auditId) {
      sheet.getRange(i + 1, statusCol + 1).setValue('misconfigured');
      return;
    }
  }

  throw new Error(`Audit ID ${auditId} not found in ${SHEET_AUDIT_QUEUE}.`);
}

/**
 * Creates a new question and adds to the questions sheet.
 */
function createQuestion(questionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);

  if (!questionData.id) {
    questionData.id = 'q' + Date.now() + '-' + questionData.taskType;
  }

  questionData.createdTimestamp = new Date().toISOString();
  questionData.createdBy = questionData.createdBy || Session.getActiveUser().getEmail() || 'system';

  if (!questionData.setId) {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const setIdIdx = headers.indexOf('setId');
    const taskTypeIdx = headers.indexOf('taskType');

    const setIds = [...new Set(data.filter(r => r[taskTypeIdx] === questionData.taskType).map(r => r[setIdIdx]))];
    questionData.setId = setIds.length > 0 ? setIds[0] : 'set' + (Date.now() % 1000);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => questionData[h] || '');
  sheet.appendRow(row);

  CacheService.getScriptCache().remove('all_questions');
  return questionData;
}

/**
 * Updates an existing question in the questions sheet.
 */
function updateQuestion(questionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === questionData.id);
  if (rowIndex === -1) throw new Error(`Question ID ${questionData.id} not found`);

  headers.forEach((header, i) => {
    if (header in questionData && header !== 'createdTimestamp' && header !== 'createdBy') {
      sheet.getRange(rowIndex + 1, i + 1).setValue(questionData[header]);
    }
  });

  CacheService.getScriptCache().remove('all_questions');
  return questionData;
}

/**
 * Deletes a question from the questions sheet.
 */
function deleteQuestion(questionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_QUESTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === questionId);
  if (rowIndex === -1) throw new Error(`Question ID ${questionId} not found`);

  sheet.deleteRow(rowIndex + 1);
  CacheService.getScriptCache().remove('all_questions');
  return { success: true, message: 'Question deleted successfully' };
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
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Retrieves pending audits from the audit queue sheet.
 */
function getPendingAudits() {
  Logger.log('Fetching pending audits...');
  return getCachedOrFetch('pending_audits', () => {
    const audits = getAllAudits();
    const evaluations = getAllEvaluations();

    const evaluatedIds = new Set(evaluations.map(e => e.evalId));
    return audits.filter(a =>
      a.auditStatus.toLowerCase() === 'pending' &&
      !evaluatedIds.has(a.auditId)
    );
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
  CacheService.getScriptCache().remove('all_audits');
}


/**
 * Updates the status of an audit and locks it.
 */
function updateAuditStatusAndLock(auditId, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_AUDIT_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');
  const lockedByIdx = headers.indexOf('lockedBy');
  const lockedAtIdx = headers.indexOf('lockedAt');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === auditId);
  if (rowIndex === -1) throw new Error('Audit not found');

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(status);
  sheet.getRange(rowIndex + 1, lockedByIdx + 1).setValue('');
  sheet.getRange(rowIndex + 1, lockedAtIdx + 1).setValue('');

  CacheService.getScriptCache().remove('all_audits');

  return { success: true };
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

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusIdx];
    const lockedBy = data[i][lockedByIdx];
    const lockedAtRaw = data[i][lockedAtIdx];

    if (status === 'In Process' && lockedBy && lockedAtRaw) {
      const lockedAt = new Date(lockedAtRaw);
      const minutesLocked = (now - lockedAt) / 60000;

      if (minutesLocked > 30) {
        sheet.getRange(i + 1, statusIdx + 1).setValue('pending');
        sheet.getRange(i + 1, lockedByIdx + 1).setValue('');
        sheet.getRange(i + 1, lockedAtIdx + 1).setValue('');
      }
    }
  }

  CacheService.getScriptCache().remove('all_audits');
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

  CacheService.getScriptCache().remove('all_audits');

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

    const summaries = getSheetDataAsObjects(evalSheet);
    const questions = getSheetDataAsObjects(questSheet);

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

    summaries.forEach(s => s.questions = map[s.id] || []);
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
  const score = data.totalPoints / data.totalPointsPossible;

  evalSheet.appendRow([
    evalId,
    data.evalId || data.auditId,
    data.referenceNumber,
    data.taskType,
    data.outcome,
    data.qaEmail || Session.getActiveUser().getEmail(),
    data.startTimestamp || new Date().toISOString(),
    stopTime,
    data.totalPoints,
    data.totalPointsPossible,
    data.status || 'completed',
    data.feedback || '',
    score
  ]);

  const questRows = data.questions.map((q, i) => [
    `${evalId}-q${i + 1}`,
    evalId,
    q.questionId,
    q.questionText,
    q.response,
    q.pointsEarned,
    q.pointsPossible,
    q.feedback || ''
  ]);

  if (questRows.length > 0) {
    questSheet.getRange(questSheet.getLastRow() + 1, 1, questRows.length, 8).setValues(questRows);
  }

  updateAuditStatus(data.evalId || data.auditId, 'evaluated');
  CacheService.getScriptCache().removeAll(['all_evaluations', 'all_audits']);

  return {
    id: evalId,
    ...data,
    stopTimestamp: stopTime,
    evalScore: score,
    questions: data.questions
  };
}

function updateEvaluationStatus(evalId, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EVAL_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === id);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(status);
  CacheService.getScriptCache().remove('all_evaluations');
}

// ====================
// Disputes Module
// ====================

function getAllDisputes() {
  return getCachedOrFetch('all_disputes', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
    const data = getSheetDataAsObjects(sheet);

    // Convert questionIds from string to array
    data.forEach(d => {
      if (d.questionIds && typeof d.questionIds === 'string') {
        d.questionIds = d.questionIds.split(',');
      }
    });

    return data;
  });
}

function saveDispute(dispute) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
  const id = 'dispute' + Date.now();
  const timestamp = new Date().toISOString();
  const userEmail = dispute.userEmail || Session.getActiveUser().getEmail();
  const questionIds = dispute.questionIds.join(',');

  sheet.appendRow([
    id,
    dispute.evalId,
    userEmail,
    timestamp,
    dispute.reason,
    questionIds,
    dispute.status || 'pending'
  ]);

  updateEvaluationStatus(dispute.evalId, 'disputed');
  CacheService.getScriptCache().removeAll(['all_disputes', 'all_evaluations']);

  return {
    id,
    evalId: dispute.evalId,
    userEmail,
    disputeTimestamp: timestamp,
    reason: dispute.reason,
    questionIds: dispute.questionIds,
    status: dispute.status || 'pending'
  };
}

function updateDisputeStatus(disputeId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DISPUTES_QUEUE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === disputeId);
  if (rowIndex === -1) return { success: false, message: "Dispute not found" };

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  return { success: true };
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

  // Get evalQuest data
  const questData = questSheet.getDataRange().getValues();
  const headers = questData[0];
  const idIndex = headers.indexOf('evalId');
  const questionIdIndex = headers.indexOf('questionId');
  const responseIndex = headers.indexOf('response');
  const pointsEarnedIndex = headers.indexOf('pointsEarned');
  const feedbackIndex = headers.indexOf('feedback');
  const pointsPossibleIndex = headers.indexOf('pointsPossible');

  // Update affected questions
  const updatedRows = [];
  for (let i = 1; i < questData.length; i++) {
    const row = questData[i];
    if (row[idIndex] !== evalId) continue;

    const decision = decisions.find(d => d.questionId === row[questionIdIndex]);
    if (!decision) continue;

    if (decision.resolution === 'overturned') {
      row[responseIndex] = 'yes';
      row[pointsEarnedIndex] = row[pointsPossibleIndex];
    }

    row[feedbackIndex] = decision.note || '';
    updatedRows.push(row);
  }

  if (updatedRows.length) {
    questSheet.getRange(2, 1, updatedRows.length, headers.length).setValues(updatedRows);
  }

  // Recalculate score totals
  const updatedQuestData = questSheet.getDataRange().getValues().filter(row => row[idIndex] === evalId);
  const totalPoints = updatedQuestData.reduce((sum, row) => sum + (parseFloat(row[pointsEarnedIndex]) || 0), 0);
  const totalPossible = updatedQuestData.reduce((sum, row) => sum + (parseFloat(row[pointsPossibleIndex]) || 0), 0);
  const evalScore = totalPossible > 0 ? totalPoints / totalPossible : 0;

  // Update evalSummary
  const summaryData = summarySheet.getDataRange().getValues();
  const sHeaders = summaryData[0];
  const sIdIndex = sHeaders.indexOf('id');
  const totalPointsIndex = sHeaders.indexOf('totalPoints');
  const evalScoreIndex = sHeaders.indexOf('evalScore');
  const statusIndex = sHeaders.indexOf('status');

  for (let i = 1; i < summaryData.length; i++) {
    const row = summaryData[i];
    if (row[sIdIndex] === evalId) {
      summarySheet.getRange(i + 1, totalPointsIndex + 1).setValue(totalPoints);
      summarySheet.getRange(i + 1, evalScoreIndex + 1).setValue(evalScore);
      summarySheet.getRange(i + 1, statusIndex + 1).setValue(status || 'resolved');
      break;
    }
  }

  // Update dispute row
  const disputeData = disputeSheet.getDataRange().getValues();
  const dHeaders = disputeData[0];
  const dIdIndex = dHeaders.indexOf('id');
  const dStatusIndex = dHeaders.indexOf('status');
  const dNotesIndex = dHeaders.indexOf('resolutionNotes');
  const dByIndex = dHeaders.indexOf('resolvedBy');
  const dTimeIndex = dHeaders.indexOf('resolutionTimestamp');

  for (let i = 1; i < disputeData.length; i++) {
    const row = disputeData[i];
    if (row[dIdIndex] === disputeId) {
      disputeSheet.getRange(i + 1, dStatusIndex + 1).setValue(status || 'resolved');
      disputeSheet.getRange(i + 1, dNotesIndex + 1).setValue(resolutionNotes || '');
      disputeSheet.getRange(i + 1, dByIndex + 1).setValue(resolvedBy);
      disputeSheet.getRange(i + 1, dTimeIndex + 1).setValue(resolvedAt);
      break;
    }
  }

  // Invalidate caches
  CacheService.getScriptCache().removeAll(['all_disputes', 'all_evaluations']);
  return true;
}

function getAllEvaluationsAndDisputes() {
  return {
    evaluations: getAllEvaluations(),
    disputes: getAllDisputes()
  };
}

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
