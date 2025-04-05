// ====================
// Evaluations Module
// ====================

/**
 * Get all evaluations with nested questions.
 */
function getAllEvaluations() {
  return getCachedOrFetch('all_evaluations', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const evalSheet = ss.getSheetByName('evalSummary');
    const questSheet = ss.getSheetByName('evalQuest');

    const summaries = getSheetDataAsObjects(evalSheet);
    const questions = getSheetDataAsObjects(questSheet);

    const grouped = {};
    questions.forEach(q => {
      if (!grouped[q.evalId]) grouped[q.evalId] = [];
      grouped[q.evalId].push({
        id: q.id,
        questionId: q.questionId,
        questionText: q.questionText,
        response: q.response,
        pointsEarned: q.pointsEarned,
        pointsPossible: q.pointsPossible,
        feedback: q.feedback
      });
    });

    summaries.forEach(s => s.questions = grouped[s.id] || []);
    return summaries;
  });
}

/**
 * Save a new evaluation and its questions.
 */
function saveEvaluation(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName('evalSummary');
  const questSheet = ss.getSheetByName('evalQuest');

  const evalId = 'eval' + Date.now();
  const stopTime = new Date().toISOString();
  const score = data.totalPoints / data.totalPointsPossible;

  // Save evaluation summary
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

  // Save individual questions
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

/**
 * Update evaluation status (e.g., to disputed or resolved).
 */
function updateEvaluationStatus(evalId, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('evalId');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === evalId);
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(status);
    CacheService.getScriptCache().remove('all_evaluations');
  }
}
