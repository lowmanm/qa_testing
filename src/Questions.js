// ====================
// Questions Module
// ====================

/**
 * Fetch all questions from the 'questions' sheet.
 */
function getAllQuestions() {
  return getCachedOrFetch('all_questions', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Fetch questions by task type.
 */
function getQuestionsForTaskType(taskType) {
  const cacheKey = `questions_${taskType}`;
  return getCachedOrFetch(cacheKey, () => {
    return getAllQuestions().filter(q => q.taskType === taskType);
  });
}

/**
 * Create a new question.
 */
function createQuestion(questionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');

  if (!questionData.id) {
    questionData.id = `q_${Date.now()}_${questionData.taskType}`;
  }

  questionData.createdTimestamp = new Date().toISOString();
  questionData.createdBy = Session.getActiveUser().getEmail() || 'system';

  // Assign a default setId if not provided
  if (!questionData.setId) {
    const data = getSheetDataAsObjects(sheet);
    const existingSetIds = [...new Set(data.filter(q => q.taskType === questionData.taskType).map(q => q.setId))];
    questionData.setId = existingSetIds.length > 0 ? existingSetIds[0] : `set_${Date.now()}`;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => questionData[h] || '');
  sheet.appendRow(row);

  CacheService.getScriptCache().removeAll(['all_questions', `questions_${questionData.taskType}`]);
  return questionData;
}

/**
 * Update an existing question.
 */
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

  CacheService.getScriptCache().removeAll(['all_questions', `questions_${questionData.taskType}`]);
  return questionData;
}

/**
 * Delete a question by ID.
 */
function deleteQuestion(questionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idCol] === questionId);
  if (rowIndex === -1) throw new Error(`Question ID ${questionId} not found`);

  const taskType = data[rowIndex][headers.indexOf('taskType')];
  sheet.deleteRow(rowIndex + 1);

  CacheService.getScriptCache().removeAll(['all_questions', `questions_${taskType}`]);
  return { success: true, message: 'Question deleted successfully' };
}
