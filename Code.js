// Refactored Code.gs

// ====================
// App Entry Points
// ====================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('QA System')
    .addItem('Open QA App', 'openQaApp')
    .addItem('Setup Spreadsheet', 'setupSpreadsheet')
    .addItem('Import Data from Email', 'importDataFromEmail')
    .addToUi();
}

function openQaApp() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'QA Evaluation System');
}

// ====================
// Spreadsheet Setup
// ====================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToCreate = {
    users: ['id', 'name', 'email', 'managerEmail', 'role', 'createdBy', 'createdTimestamp', 'avatarUrl'],
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
    if (existingHeaders.join(',') !== headers.join(',')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
}

// ====================
// Caching Utilities
// ====================

const CACHE_DURATION = 300; // seconds (5 minutes)

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

function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((key, i) => {
      if (key) obj[key] = row[i];
    });
    return obj;
  });
}

// ====================
// Users Module
// ====================

function getAllUsers() {
  return getCachedOrFetch('all_users', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
    return getSheetDataAsObjects(sheet);
  });
}

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const users = getAllUsers();
  const user = users.find(function(user) {
    return user.email === email;
  });

  // Default to the first user for testing/demo purposes if no matching user
  if (!user && users.length > 0) {
    user = users[0];
  }

  return user || {
    id: 'unknown',
    name: 'Unknown User',
    email: email,
    role: 'qa_analyst'
  };
}


function createUser(userData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');

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

// ====================
// Questions Module
// ====================

function getAllQuestions() {
  return getCachedOrFetch('all_questions', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
    return getSheetDataAsObjects(sheet);
  });
}

function getQuestionsForTaskType(taskType) {
  const cacheKey = 'questions_' + taskType;
  return getCachedOrFetch(cacheKey, () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
    const data = getSheetDataAsObjects(sheet);
    return data.filter(q => q.taskType === taskType);
  });
}

function markAuditAsMisconfigured(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  if (!sheet) throw new Error('Sheet "auditQueue" not found');

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

  throw new Error('Audit ID not found in auditsQueue.');
}

function createQuestion(questionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');

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

function updateQuestion(questionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
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

function deleteQuestion(questionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
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

function getAllAudits() {
  return getCachedOrFetch('all_audits', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
    return getSheetDataAsObjects(sheet);
  });
}

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

function updateAuditStatus(auditId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === auditId);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  CacheService.getScriptCache().remove('all_audits');
}

function updateAuditStatusAndLock(auditId, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');
  const lockedIdx = headers.indexOf('locked');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === auditId);
  if (rowIndex === -1) throw new Error('Audit not found');

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(status);
  sheet.getRange(rowIndex + 1, lockedIdx + 1).setValue(status === 'evaluated' ? false : true);
  CacheService.getScriptCache().remove('all_audits');

  return { success: true };
}

function prepareEvaluation(auditId) {
  const result = updateAuditStatusAndLock(auditId, 'In Process');
  if (!result.success) throw new Error('Failed to update audit status');

  const audits = getAllAudits();
  const audit = audits.find(a => a.auditId === auditId);
  if (!audit) throw new Error('Audit not found');

  return audit;
}

// ====================
// Evaluations Module
// ====================

function getAllEvaluations() {
  return getCachedOrFetch('all_evaluations', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const evalSheet = ss.getSheetByName('evalSummary');
    const questSheet = ss.getSheetByName('evalQuest');
    const disputeSheet = ss.getSheetByName('disputesQueue');

    const summaries = getSheetDataAsObjects(evalSheet);
    const questions = getSheetDataAsObjects(questSheet);
    const disputes = getSheetDataAsObjects(disputeSheet);

    // Map questions by evalId
    const questionMap = {};
    questions.forEach(q => {
      if (!questionMap[q.evalId]) questionMap[q.evalId] = [];
      questionMap[q.evalId].push({
        id: q.id,
        questionId: q.questionId,
        questionText: q.questionText,
        response: q.response,
        pointsEarned: q.pointsEarned,
        pointsPossible: q.pointsPossible,
        feedback: q.feedback
      });
    });

    // Set of disputed evalIds
    const disputedEvalIds = new Set(disputes.map(d => d.evalId));

    // Add questions and disputed flag to evaluations
    summaries.forEach(s => {
      s.questions = questionMap[s.id] || [];
      s.isDisputed = disputedEvalIds.has(s.id);
    });

    return summaries;
  });
}

function saveEvaluation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName('evalSummary');
  const questSheet = ss.getSheetByName('evalQuest');

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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('evalId');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === evalId);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(status);
  CacheService.getScriptCache().remove('all_evaluations');
}

// ====================
// Disputes Module
// ====================

function getAllDisputes() {
  return getCachedOrFetch('all_disputes', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === disputeId);
  if (rowIndex === -1) return { success: false, message: "Dispute not found" };

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  return { success: true };
}

function resolveDispute(resolution) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const questSheet = ss.getSheetByName('evalQuest');
  const summarySheet = ss.getSheetByName('evalSummary');
  const disputeSheet = ss.getSheetByName('disputesQueue');

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
    questSheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
  }

  // Recalculate score totals
  const updatedQuestData = questSheet.getDataRange().getValues().filter(row => row[idIndex] === evalId);
  const totalPoints = updatedQuestData.reduce((sum, row) => sum + (parseFloat(row[pointsEarnedIndex]) || 0), 0);
  const totalPossible = updatedQuestData.reduce((sum, row) => sum + (parseFloat(row[pointsPossibleIndex]) || 0), 0);
  const evalScore = totalPossible > 0 ? totalPoints / totalPossible : 0;

  // Update evalSummary —
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

// ==============================
// Utility & Shared Helper Functions
// ==============================

// Convert snake_case to Title Case for display
function toTitleCase(str) {
  return (str || '').split('_').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

