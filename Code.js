// Cleaned and Optimized Code.js

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  const page = template.evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  if (e.parameter && e.parameter.evaluationId) {
    page.addMetaTag('evaluationId', e.parameter.evaluationId);
  }

  return page;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const CACHE_DURATION = 60; // seconds

function getCachedOrFetch(cacheKey, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch (e) {
      Logger.log('Error parsing cached data: ' + e.message);
    }
  }
  const freshData = fetchFunction();
  try {
    cache.put(cacheKey, JSON.stringify(freshData), CACHE_DURATION);
  } catch (e) {
    Logger.log('Error caching data: ' + e.message);
  }
  return freshData;
}

function getSheetDataAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

// USERS
function getAllUsers() {
  return getCachedOrFetch('all_users', () => getSheetDataAsObjects('users'));
}

function getUserByEmail(email) {
  return getAllUsers().find(u => u.email === email);
}

// AUDITS
function getPendingAudits() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const auditStatusIndex = headers.indexOf('auditStatus');
  const lockedIndex = headers.indexOf('locked');
  const lockedAtIndex = headers.indexOf('lockedAt');

  const now = new Date();
  const audits = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const isLocked = row[lockedIndex] === true || row[lockedIndex] === 'TRUE';
    const lockedAt = new Date(row[lockedAtIndex]);

    if (isLocked && !isNaN(lockedAt.getTime()) && (now - lockedAt) > 10 * 60 * 1000) {
      sheet.getRange(i + 1, lockedIndex + 1).setValue(false);
      sheet.getRange(i + 1, lockedAtIndex + 1).setValue('');
      sheet.getRange(i + 1, auditStatusIndex + 1).setValue('pending');
      row[lockedIndex] = false;
    }

    if (row[auditStatusIndex] !== 'evaluated' && !row[lockedIndex]) {
      const audit = {};
      headers.forEach((key, j) => audit[key] = row[j]);
      audits.push(audit);
    }
  }

  return audits;
}

function getAuditById(auditId) {
  const audits = getAllAudits();
  return audits.find(a => a.auditId === auditId);
}

function getAllAudits() {
  return getSheetDataAsObjects('auditQueue');
}

function updateAuditStatusAndLock(auditId, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('auditId');
  const statusIndex = headers.indexOf('auditStatus');
  const lockedIndex = headers.indexOf('locked');
  const lockedAtIndex = headers.indexOf('lockedAt');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === auditId) {
      sheet.getRange(i + 1, statusIndex + 1).setValue(status);
      sheet.getRange(i + 1, lockedIndex + 1).setValue(status === 'evaluated' ? false : true);
      sheet.getRange(i + 1, lockedAtIndex + 1).setValue(status === 'evaluated' ? '' : new Date().toISOString());
      return { success: true };
    }
  }
  throw new Error('Audit not found');
}

// QUESTIONS
function getAllQuestions() {
  return getCachedOrFetch('all_questions', () => getSheetDataAsObjects('questions'));
}

function getQuestionsForTaskType(taskType) {
  return getAllQuestions().filter(q => q.taskType === taskType);
}

// EVALUATION FLOW
function prepareEvaluationWithQuestions(auditId) {
  const audit = getAuditById(auditId);
  const questions = getQuestionsForTaskType(audit.taskType);
  return { audit, questions };
}

function saveEvaluation(evaluation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
  const questSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalQuest');

  evaluation.stopTimestamp = new Date().toISOString();
  evaluation.evalScore = Math.round((evaluation.totalPoints / evaluation.totalPointsPossible) * 100);

  const summaryHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const summaryRow = summaryHeaders.map(h => evaluation[h] || '');
  sheet.appendRow(summaryRow);

  evaluation.questions.forEach(q => {
    const questHeaders = questSheet.getRange(1, 1, 1, questSheet.getLastColumn()).getValues()[0];
    const questRow = questHeaders.map(h => q[h] || evaluation.evalId || '');
    questSheet.appendRow(questRow);
  });

  notifyAgentAndManager(evaluation);
  return { success: true };
}

function notifyAgentAndManager(evaluation) {
  const user = getUserByEmail(evaluation.agentEmail);
  const managerEmail = user?.managerEmail;

  const agentHtml = `<p>Your score: ${evaluation.totalPoints}/${evaluation.totalPointsPossible}</p>`;
  GmailApp.sendEmail(user.email, 'Your Evaluation Results', '', { htmlBody: agentHtml });

  if (managerEmail) {
    const managerHtml = `<p>Evaluation for ${user.name} completed.</p>`;
    GmailApp.sendEmail(managerEmail, 'Evaluation Completed: ' + user.name, '', { htmlBody: managerHtml });
  }
}

// EVALUATIONS VIEW
function getCompletedEvaluations() {
  const all = getSheetDataAsObjects('evalSummary');
  return all.filter(e => e.status === 'completed');
}

// DISPUTES
function prepareDisputeForm(evalId) {
  const evalData = getSheetDataAsObjects('evalSummary').find(e => e.evalId === evalId);
  const questionData = getSheetDataAsObjects('evalQuest').filter(q => q.evalId === evalId);
  return { evaluation: evalData, questions: questionData };
}

function submitDispute(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const newRow = headers.map(h => {
    if (h === 'id') return 'dispute_' + new Date().getTime();
    if (h === 'evalId') return payload.evalId;
    if (h === 'userEmail') return payload.userEmail;
    if (h === 'disputeTimestamp') return payload.timestamp;
    if (h === 'reason') return payload.reason;
    if (h === 'questionIds') return payload.questionIds.join(',');
    if (h === 'status') return 'pending';
    return '';
  });

  sheet.appendRow(newRow);
  return { success: true };
}

function getAllDisputes() {
  return getSheetDataAsObjects('disputesQueue').filter(d => d.status === 'pending');
}

function resolveDispute(disputeId, resolution, feedback) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIndex = headers.indexOf('id');
  const statusIndex = headers.indexOf('status');
  const resolutionIndex = headers.indexOf('resolution');
  const feedbackIndex = headers.indexOf('resolutionFeedback');
  const resolvedAtIndex = headers.indexOf('resolutionTimestamp');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === disputeId) {
      sheet.getRange(i + 1, statusIndex + 1).setValue('resolved');
      sheet.getRange(i + 1, resolutionIndex + 1).setValue(resolution);
      sheet.getRange(i + 1, feedbackIndex + 1).setValue(feedback);
      sheet.getRange(i + 1, resolvedAtIndex + 1).setValue(new Date().toISOString()); // renamed to resolutionTimestamp
      return { success: true };
    }
  }
  throw new Error('Dispute not found');
}
