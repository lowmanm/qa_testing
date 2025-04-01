// Cleaned and Optimized Code.js

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  const page = template.evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  // Deep-linking support for specific evaluation
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

  // Save summary
  const summaryHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const summaryRow = summaryHeaders.map(h => evaluation[h] || '');
  sheet.appendRow(summaryRow);

  // Save questions
  evaluation.questions.forEach(q => {
    const questHeaders = questSheet.getRange(1, 1, 1, questSheet.getLastColumn()).getValues()[0];
    const questRow = questHeaders.map(h => q[h] || evaluation.evalId || '');
    questSheet.appendRow(questRow);
  });

  notifyAgentAndManager(evaluation);
  return { success: true };
}

// NOTIFICATIONS
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

// UTILITIES
function getSheetDataAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}
