
// Code.gs - Main script file
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QA Evaluation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Optimized data structure setup without mock data
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create users sheet
  if (!ss.getSheetByName('users')) {
    var usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      // Rename existing Users sheet
      usersSheet.setName('users');
      // Update headers to match new schema
      usersSheet.getRange(1, 1, 1, 8).setValues([['id', 'name', 'email', 'managerEmail', 'role', 'createdBy', 'createdTimestamp', 'avatarUrl']]);
    } else {
      // Create new users sheet
      usersSheet = ss.insertSheet('users');
      usersSheet.appendRow(['id', 'name', 'email', 'managerEmail', 'role', 'createdBy', 'createdTimestamp', 'avatarUrl']);
    }
  }

  // Create auditQueue sheet
  if (!ss.getSheetByName('auditQueue')) {
    var auditQueueSheet = ss.getSheetByName('Tasks');
    if (auditQueueSheet) {
      // Rename existing Tasks sheet
      auditQueueSheet.setName('auditQueue');
      // Update headers to match new schema
      auditQueueSheet.getRange(1, 1, 1, 10).setValues([[
        'auditId',
        'taskId',
        'referenceNumber',
        'auditStatus',
        'agentEmail',
        'requestType',
        'taskType',
        'outcome',
        'taskTimestamp',
        'auditTimestamp',
        'locked' // New column 2025-03-24
      ]]);
    } else {
      // Create new auditQueue sheet
      auditQueueSheet = ss.insertSheet('auditQueue');
      auditQueueSheet.appendRow([
        'auditId',
        'taskId',
        'referenceNumber',
        'auditStatus',
        'agentEmail',
        'requestType',
        'taskType',
        'outcome',
        'taskTimestamp',
        'auditTimestamp',
        'locked' // New column 2025-03-24
      ]);
    }
  }

  // Create evalSummary sheet
  if (!ss.getSheetByName('evalSummary')) {
    var evalSummarySheet = ss.getSheetByName('Evaluations');
    if (evalSummarySheet) {
      // Rename existing Evaluations sheet
      evalSummarySheet.setName('evalSummary');
      // Update headers to match new schema
      evalSummarySheet.getRange(1, 1, 1, 13).setValues([[
        'id',
        'evalId',
        'referenceNumber',
        'taskType',
        'outcome',
        'qaEmail',
        'startTimestamp',
        'stopTimestamp',
        'totalPoints',
        'totalPointsPossible',
        'status',
        'feedback',
        'evalScore'
      ]]);
    } else {
      // Create new evalSummary sheet
      evalSummarySheet = ss.insertSheet('evalSummary');
      evalSummarySheet.appendRow([
        'id',
        'evalId',
        'referenceNumber',
        'taskType',
        'outcome',
        'qaEmail',
        'startTimestamp',
        'stopTimestamp',
        'totalPoints',
        'totalPointsPossible',
        'status',
        'feedback',
        'evalScore'
      ]);
    }
  }

  // Create evalQuest sheet
  if (!ss.getSheetByName('evalQuest')) {
    var evalQuestSheet = ss.insertSheet('evalQuest');
    evalQuestSheet.appendRow([
      'id',
      'evalId',
      'questionId',
      'questionText',
      'response',
      'pointsEarned',
      'pointsPossible',
      'feedback'
    ]);
  }

  // Create questions sheet
  if (!ss.getSheetByName('questions')) {
    var questionsSheet = ss.getSheetByName('Questions');
    if (questionsSheet) {
      // Rename and update existing Questions sheet
      questionsSheet.setName('questions');
      // Update headers to match new schema
      questionsSheet.getRange(1, 1, 1, 7).setValues([[
        'id',
        'setId',
        'taskType',
        'questionText',
        'pointsPossible',
        'createdBy',
        'createdTimestamp'
      ]]);
    } else {
      // Create new questions sheet
      questionsSheet = ss.insertSheet('questions');
      questionsSheet.appendRow([
        'id',
        'setId',
        'taskType',
        'questionText',
        'pointsPossible',
        'createdBy',
        'createdTimestamp'
      ]);
    }
  }

  // Create disputesQueue sheet
  if (!ss.getSheetByName('disputesQueue')) {
    var disputesQueueSheet = ss.getSheetByName('Disputes');
    if (disputesQueueSheet) {
      // Rename existing Disputes sheet
      disputesQueueSheet.setName('disputesQueue');
      // Update headers to match new schema
      disputesQueueSheet.getRange(1, 1, 1, 10).setValues([[
        'id',
        'evalId',
        'userEmail',
        'disputeTimestamp',
        'reason',
        'questionIds',
        'status',
        'resolutionNotes',         // NEW
        'resolvedBy',              // NEW
        'resolutionTimestamp'      // NEW
      ]]);
    } else {
      // Create new disputesQueue sheet
      disputesQueueSheet = ss.insertSheet('disputesQueue');
      disputesQueueSheet.appendRow([
        'id',
        'evalId',
        'userEmail',
        'disputeTimestamp',
        'reason',
        'questionIds',
        'status',
        'resolutionNotes',         // NEW
        'resolvedBy',              // NEW
        'resolutionTimestamp'      // NEW
      ]);
    }
  }
}

// Create a new menu when the spreadsheet opens
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('QA System')
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

// Optimized function to import data from Gmail
function importDataFromEmail() {
  try {
    // Search for emails with the specified subject
    var query = 'subject:"NVS Audit File" has:attachment filename:nvs_qa_audit.csv';
    var threads = GmailApp.search(query, 0, 1);  // Get the most recent matching thread

    if (threads.length === 0) {
      throw new Error("No emails found with subject 'NVS Audit File' and attachment 'nvs_qa_audit.csv'");
    }

    // Get the first message from the thread
    var messages = threads[0].getMessages();
    var message = messages[0];

    // Get attachments
    var attachments = message.getAttachments();
    var csvAttachment = null;

    // Find the CSV attachment with the correct name
    for (var i = 0; i < attachments.length; i++) {
      if (attachments[i].getName() === 'nvs_qa_audit.csv') {
        csvAttachment = attachments[i];
        break;
      }
    }

    if (!csvAttachment) {
      throw new Error("No attachment named 'nvs_qa_audit.csv' found in the email");
    }

    // Parse the CSV data
    var content = csvAttachment.getDataAsString();
    var csvData = Utilities.parseCsv(content);

    // Get header row
    var headers = csvData[0];
    Logger.log("Found CSV headers: " + headers.join(", "));

    // Get or create the auditQueue sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var auditQueueSheet = ss.getSheetByName('auditQueue');

    // If auditQueue sheet doesn't exist, create it or rename from Tasks
    if (!auditQueueSheet) {
      if (ss.getSheetByName('Tasks')) {
        ss.getSheetByName('Tasks').setName('auditQueue');
        auditQueueSheet = ss.getSheetByName('auditQueue');
      } else {
        auditQueueSheet = ss.insertSheet('auditQueue');
        auditQueueSheet.appendRow([
          'auditId',
          'taskId',
          'referenceNumber',
          'auditStatus',
          'agentEmail',
          'requestType',
          'taskType',
          'outcome',
          'taskTimestamp',
          'auditTimestamp'
        ]);
      }
    }

    var auditQueueHeaders = auditQueueSheet.getRange(1, 1, 1, auditQueueSheet.getLastColumn()).getValues()[0];
    Logger.log("Sheet headers: " + auditQueueHeaders.join(", "));

    // Improved CSV to sheet mapping - case insensitive and flexible matching
    var headerMapping = {};
    headers.forEach(function(header, index) {
      var normalizedHeader = header.toLowerCase().trim().replace(/\s+/g, '');

      // Map CSV headers to our sheet headers based on normalized strings
      if (normalizedHeader.includes('taskid') || normalizedHeader.includes('task_id')) {
        headerMapping['taskId'] = index;
      } else if (normalizedHeader.includes('reference') || normalizedHeader.includes('ref')) {
        headerMapping['referenceNumber'] = index;
      } else if (normalizedHeader.includes('status')) {
        headerMapping['auditStatus'] = index;
      } else if ((normalizedHeader.includes('agent') && normalizedHeader.includes('email')) || normalizedHeader.includes('agentemail')) {
        headerMapping['agentEmail'] = index;
      } else if (normalizedHeader.includes('request') && normalizedHeader.includes('type')) {
        headerMapping['requestType'] = index;
      } else if (normalizedHeader.includes('task') && normalizedHeader.includes('type')) {
        headerMapping['taskType'] = index;
      } else if (normalizedHeader.includes('outcome') || normalizedHeader.includes('result')) {
        headerMapping['outcome'] = index;
      } else if ((normalizedHeader.includes('task') && normalizedHeader.includes('time')) ||
                 (normalizedHeader.includes('task') && normalizedHeader.includes('date'))) {
        headerMapping['taskTimestamp'] = index;
      }
    });

    Logger.log("Header mapping: " + JSON.stringify(headerMapping));

    // Bulk import preparation
    var auditsToImport = [];
    var now = new Date().toISOString(); // Store timestamp in UTC
    var auditsImported = 0;

    // Process each row in the CSV (skip the header row)
    for (var i = 1; i < csvData.length; i++) {
      var row = csvData[i];

      // Skip empty rows
      if (row.length === 0 || (row.length === 1 && row[0] === '')) continue;

      // Generate a unique auditId
      var auditId = 'audit' + (new Date().getTime() + i);

      // Create audit row
      var auditRow = [];

      // Fill row with mapped data
      for (var j = 0; j < auditQueueHeaders.length; j++) {
        var header = auditQueueHeaders[j];
        if (header === 'auditId') {
          auditRow.push(auditId);
        } else if (header === 'auditTimestamp') {
          auditRow.push(now);
        } else if (header === 'taskTimestamp' && headerMapping[header] !== undefined) {
          var rawTimestamp = row[headerMapping[header]];
          try {
            // Try to parse the date and format to ISO
            var parsedDate = new Date(rawTimestamp);
            if (!isNaN(parsedDate.getTime())) {
              auditRow.push(parsedDate.toISOString());
            } else {
              auditRow.push(rawTimestamp); // Keep original if parsing fails
              Logger.log("Warning: Couldn't parse date: " + rawTimestamp);
            }
          } catch (e) {
            auditRow.push(rawTimestamp); // Keep original if exception occurs
            Logger.log("Error parsing date: " + e.message);
          }
        } else if (headerMapping[header] !== undefined) {
          auditRow.push(row[headerMapping[header]]);
        } else {
          auditRow.push('');
        }
      }

      // Add to bulk import array
      auditsToImport.push(auditRow);
      auditsImported++;

      // Process in batches of 100 to avoid memory issues
      if (auditsToImport.length >= 100) {
        if (auditsToImport.length > 0) {
          auditQueueSheet.getRange(auditQueueSheet.getLastRow() + 1, 1, auditsToImport.length, auditQueueHeaders.length)
            .setValues(auditsToImport);
          auditsToImport = [];
        }
      }
    }

    // Import any remaining rows
    if (auditsToImport.length > 0) {
      auditQueueSheet.getRange(auditQueueSheet.getLastRow() + 1, 1, auditsToImport.length, auditQueueHeaders.length)
        .setValues(auditsToImport);
    }

    // Mark the email as read
    message.markRead();

    return {
      success: true,
      message: `Successfully imported ${auditsImported} audits from email.`,
      date: new Date().toISOString() // Return timestamp in UTC
    };
  } catch (error) {
    Logger.log('Error importing data from email: ' + error.message);
    return {
      success: false,
      message: 'Error importing data: ' + error.message,
      date: new Date().toISOString() // Return timestamp in UTC
    };
  }
}

// Optimized data access functions using caching for better performance
// Cache duration in seconds
var CACHE_DURATION = 300; // 5 minutes

// Helper function to get cache or fetch fresh data
function getCachedOrFetch(cacheKey, fetchFunction) {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get(cacheKey);

  if (cachedData) {
    try {
      return JSON.parse(cachedData);
    } catch (e) {
      Logger.log('Error parsing cached data: ' + e.message);
      // Proceed to fetch fresh data if parsing fails
    }
  }

  // Fetch fresh data
  var freshData = fetchFunction();

  // Store in cache
  try {
    cache.put(cacheKey, JSON.stringify(freshData), CACHE_DURATION);
  } catch (e) {
    Logger.log('Error caching data: ' + e.message);
    // Continue even if caching fails
  }

  return freshData;
}

// Optimized function to get all users
function getAllUsers() {
  return getCachedOrFetch('all_users', function() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
    return getSheetDataAsObjects(sheet);
  });
}

function createUser(userData) {
  var usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');

  // Generate ID if not provided
  if (!userData.id) {
    userData.id = 'user' + new Date().getTime();
  }

  // Add metadata
  userData.createdTimestamp = new Date().toISOString();
  userData.createdBy = userData.createdBy || Session.getActiveUser().getEmail() || 'system';

  // Log input object
  Logger.log('[createUser] userData received:\n' + JSON.stringify(userData, null, 2));

  // Get sheet headers
  var headers = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  Logger.log('[createUser] Sheet Headers: ' + headers.join(', '));

  // Construct row based on header order
  var row = headers.map(function(header) {
    var value = userData[header] || '';
    Logger.log(`[createUser] Mapping header "${header}" to value: "${value}"`);
    return value;
  });

  // Log the full row before insertion
  Logger.log('[createUser] Row to insert: ' + JSON.stringify(row));

  // Append row to sheet
  usersSheet.appendRow(row);

  // Clear cache
  CacheService.getScriptCache().remove('all_users');

  return userData;
}

function updateUser(userData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('id');

  if (idColIndex === -1) {
    throw new Error('Could not find "id" column in users sheet');
  }

  // Find the user row
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === userData.id) {
      rowIndex = i + 2; // +2 because of 0-indexing and header row
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('User with ID ' + userData.id + ' not found');
  }

  // Update each cell
  headers.forEach(function(header, colIndex) {
    if (header in userData && header !== 'createdTimestamp' && header !== 'createdBy') {
      sheet.getRange(rowIndex, colIndex + 1).setValue(userData[header]);
    }
  });

  // Invalidate users cache
  CacheService.getScriptCache().remove('all_users');

  return userData;
}

function deleteUser(userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('id');

  if (idColIndex === -1) {
    throw new Error('Could not find "id" column in users sheet');
  }

  // Find the user row
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === userId) {
      rowIndex = i + 2; // +2 because of 0-indexing and header row
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('User with ID ' + userId + ' not found');
  }

  // Delete the row
  sheet.deleteRow(rowIndex);

  // Invalidate users cache
  CacheService.getScriptCache().remove('all_users');

  return { success: true, message: 'User deleted successfully' };
}

// Use caching for question data
function getAllQuestions() {
  return getCachedOrFetch('all_questions', function() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
    return getSheetDataAsObjects(sheet);
  });
}

// Use caching for audit data
function getAllAudits() {
  return getCachedOrFetch('all_audits', function() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
    return getSheetDataAsObjects(sheet);
  });
}

// Server-side function to update audit status and lock the record
function updateAuditStatusAndLock(auditId, status) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('auditId');
  var statusColIndex = headers.indexOf('auditStatus');
  var lockedColIndex = headers.indexOf('locked');

  if (idColIndex === -1 || statusColIndex === -1 || lockedColIndex === -1) {
    Logger.log('Warning: Could not find required columns in auditQueue sheet');
    return;
  }

  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === auditId) {
      sheet.getRange(i + 2, statusColIndex + 1).setValue(status);
      sheet.getRange(i + 2, lockedColIndex + 1).setValue(status === 'evaluated' ? false : true);

      // Invalidate audits cache
      CacheService.getScriptCache().remove('all_audits');
      return { success: true };
    }
  }

  Logger.log('Warning: Could not find audit with ID ' + auditId);
  throw new Error('Audit not found');
}

function createQuestion(questionData) {
  var questionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');

  // Create a unique question ID if not provided
  if (!questionData.id) {
    var baseId = 'q' + new Date().getTime();
    questionData.id = baseId + '-' + questionData.taskType;
  }

  // Set created timestamp in UTC
  questionData.createdTimestamp = new Date().toISOString();

  // Set createdBy if not provided
  if (!questionData.createdBy) {
    questionData.createdBy = Session.getActiveUser().getEmail() || 'system';
  }

  // Generate a setId if not provided
  if (!questionData.setId) {
    // Find existing sets for this task type and create a new one
    var existingSets = questionsSheet.getDataRange().getValues();
    var headers = existingSets.shift();
    var setIdColIndex = headers.indexOf('setId');
    var taskTypeColIndex = headers.indexOf('taskType');

    var existingSetIds = [];
    for (var i = 0; i < existingSets.length; i++) {
      if (existingSets[i][taskTypeColIndex] === questionData.taskType) {
        var setId = existingSets[i][setIdColIndex];
        if (setId && !existingSetIds.includes(setId)) {
          existingSetIds.push(setId);
        }
      }
    }

    // Use the first set for this task type or create a new one
    if (existingSetIds.length > 0) {
      questionData.setId = existingSetIds[0];
    } else {
      questionData.setId = 'set' + (new Date().getTime() % 1000);
    }
  }

  // Get the headers
  var headers = questionsSheet.getRange(1, 1, 1, questionsSheet.getLastColumn()).getValues()[0];

  // Create a row with data in the correct order
  var row = [];
  headers.forEach(function(header) {
    row.push(questionData[header] || '');
  });

  // Append the row
  questionsSheet.appendRow(row);

  // Invalidate questions cache
  CacheService.getScriptCache().remove('all_questions');

  return questionData;
}

function updateQuestion(questionData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('id');

  if (idColIndex === -1) {
    throw new Error('Could not find "id" column in questions sheet');
  }

  // Find the question row
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === questionData.id) {
      rowIndex = i + 2; // +2 because of 0-indexing and header row
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Question with ID ' + questionData.id + ' not found');
  }

  // Update each cell
  headers.forEach(function(header, colIndex) {
    if (header in questionData && header !== 'createdTimestamp' && header !== 'createdBy') {
      sheet.getRange(rowIndex, colIndex + 1).setValue(questionData[header]);
    }
  });

  // Invalidate questions cache
  CacheService.getScriptCache().remove('all_questions');

  return questionData;
}

function deleteQuestion(questionId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('id');

  if (idColIndex === -1) {
    throw new Error('Could not find "id" column in questions sheet');
  }

  // Find the question row
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === questionId) {
      rowIndex = i + 2; // +2 because of 0-indexing and header row
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Question with ID ' + questionId + ' not found');
  }

  // Delete the row
  sheet.deleteRow(rowIndex);

  // Invalidate questions cache
  CacheService.getScriptCache().remove('all_questions');

  return { success: true, message: 'Question deleted successfully' };
}

function getQuestionsForTaskType(taskType) {
  var cacheKey = 'questions_' + taskType;

  return getCachedOrFetch(cacheKey, function() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('questions');
    var data = getSheetDataAsObjects(sheet);

    return data.filter(function(question) {
      return question.taskType === taskType;
    });
  });
}

function getAllEvaluations() {
  return getCachedOrFetch('all_evaluations', function() {
    var evalSummarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
    var evalQuestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalQuest');

    var summaries = getSheetDataAsObjects(evalSummarySheet);
    var questions = getSheetDataAsObjects(evalQuestSheet);

    // Log the retrieved questions
    console.log('Retrieved questions:', questions);

    // Group the questions by evalId
    var questionsMap = {};
    questions.forEach(function(q) {
      if (!questionsMap[q.evalId]) {
        questionsMap[q.evalId] = [];
      }
      questionsMap[q.evalId].push({
        id: q.id,
        questionId: q.questionId,
        questionText: q.questionText,
        response: q.response,
        pointsEarned: q.pointsEarned,
        pointsPossible: q.pointsPossible,
        feedback: q.feedback
      });
    });

    // Attach questions to summaries
    summaries.forEach(function(summary) {
      summary.questions = questionsMap[summary.id] || [];
    });

     // Log the final summaries with questions attached
    console.log('Final summaries with questions:', summaries);

    return summaries;
  });
}

function getAllDisputes() {
  return getCachedOrFetch('all_disputes', function() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
    var data = getSheetDataAsObjects(sheet);

    // Parse the questionIds string into an array
    data.forEach(function(dispute) {
      if (dispute.questionIds) {
        dispute.questionIds = dispute.questionIds.split(',');
      }
    });

    return data;
  });
}

// More efficient helper function to convert sheet data to an array of objects
function getSheetDataAsObjects(sheet) {
  if (!sheet) return [];

  // Get all data at once to minimize API calls
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  if (values.length <= 1) return []; // Just headers or empty

  var headers = values[0];
  var objects = [];

  // Process in batches for large datasets
  var batchSize = 1000;
  for (var i = 1; i < values.length; i += batchSize) {
    var endIdx = Math.min(i + batchSize, values.length);
    for (var j = i; j < endIdx; j++) {
      var obj = {};
      for (var k = 0; k < headers.length; k++) {
        // Skip empty headers
        if (headers[k]) {
          obj[headers[k]] = values[j][k];
        }
      }
      objects.push(obj);
    }
  }

  return objects;
}

// Function to save a new evaluation - updated for new schema
function saveEvaluation(evaluationData) {
  var evalSummarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
  var evalQuestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalQuest');

  // Create a unique ID for the evaluation
  var evalId = 'eval' + (new Date().getTime());
  var stopTimestamp = new Date().toISOString(); // Store in UTC

  // Calculate evalScore as a decimal
  var evalScore = evaluationData.totalPoints / evaluationData.totalPointsPossible;

  // Save the summary
  evalSummarySheet.appendRow([
    evalId,
    evaluationData.evalId || evaluationData.auditId,
    evaluationData.referenceNumber,
    evaluationData.taskType,
    evaluationData.outcome,
    evaluationData.qaEmail || Session.getActiveUser().getEmail(),
    evaluationData.startTimestamp || evaluationData.date || new Date().toISOString(), // UTC
    stopTimestamp,
    evaluationData.totalPoints,
    evaluationData.totalPointsPossible,
    evaluationData.status || 'completed',
    evaluationData.feedback || '',
    evalScore
  ]);

  // Prepare batch insert for questions
  var questRows = [];
  evaluationData.questions.forEach(function(question, index) {
    var questId = evalId + '-q' + (index + 1);
    questRows.push([
      questId,
      evalId,
      question.questionId,
      question.questionText,
      question.response,
      question.pointsEarned,
      question.pointsPossible,
      question.feedback || ''
    ]);
  });

  // Batch insert all question responses at once
  if (questRows.length > 0) {
    evalQuestSheet.getRange(evalQuestSheet.getLastRow() + 1, 1, questRows.length, 8)
      .setValues(questRows);
  }

  // Update the audit status in the auditQueue sheet
  updateAuditStatus(evaluationData.evalId || evaluationData.auditId, 'evaluated');

  // Invalidate relevant caches
  var cache = CacheService.getScriptCache();
  cache.remove('all_evaluations');

  return {
    id: evalId,
    evalId: evaluationData.evalId || evaluationData.auditId,
    referenceNumber: evaluationData.referenceNumber,
    taskType: evaluationData.taskType,
    outcome: evaluationData.outcome,
    qaEmail: evaluationData.qaEmail || Session.getActiveUser().getEmail(),
    startTimestamp: evaluationData.startTimestamp || evaluationData.date || new Date().toISOString(),
    stopTimestamp: stopTimestamp,
    totalPoints: evaluationData.totalPoints,
    totalPointsPossible: evaluationData.totalPointsPossible,
    status: evaluationData.status || 'completed',
    feedback: evaluationData.feedback || '',
    evalScore: evalScore,
    questions: evaluationData.questions
  };
}

// Function to save a new dispute - updated for new schema
function saveDispute(disputeData) {
  var disputesQueueSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');

  // Create a unique ID for the dispute
  var disputeId = 'dispute' + (new Date().getTime());
  var disputeTimestamp = new Date().toISOString(); // Store in UTC

  // Convert questionIds array to string for storage
  var questionIdsString = disputeData.questionIds.join(',');

  // Log dispute data for debugging
  Logger.log('Saving dispute with data: ' + JSON.stringify(disputeData));

  disputesQueueSheet.appendRow([
    disputeId,
    disputeData.evalId,
    disputeData.userEmail || Session.getActiveUser().getEmail(),
    disputeTimestamp,
    disputeData.reason,
    questionIdsString,
    disputeData.status || 'pending'
  ]);

  // Update the evaluation status in the evalSummary sheet
  updateEvaluationStatus(disputeData.evalId, 'disputed');

  // Invalidate relevant caches
  var cache = CacheService.getScriptCache();
  cache.remove('all_disputes');
  cache.remove('all_evaluations');

  return {
    id: disputeId,
    evalId: disputeData.evalId,
    userEmail: disputeData.userEmail || Session.getActiveUser().getEmail(),
    disputeTimestamp: disputeTimestamp,
    reason: disputeData.reason,
    questionIds: disputeData.questionIds,
    status: disputeData.status || 'pending'
  };
}

// Function to update audit status with improved performance
function updateAuditStatus(auditId, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('auditId');
  var statusColIndex = headers.indexOf('auditStatus');

  if (idColIndex === -1 || statusColIndex === -1) {
    Logger.log('Warning: Could not find required columns in auditQueue sheet');
    return;
  }

  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === auditId) {
      sheet.getRange(i + 2, statusColIndex + 1).setValue(newStatus);

      // Invalidate audits cache
      CacheService.getScriptCache().remove('all_audits');
      return;
    }
  }

  Logger.log('Warning: Could not find audit with ID ' + auditId);
}

// Function to update evaluation status with improved performance
function updateEvaluationStatus(evalId, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('evalSummary');
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var idColIndex = headers.indexOf('evalId');
  var statusColIndex = headers.indexOf('status');

  if (idColIndex === -1 || statusColIndex === -1) {
    Logger.log('Warning: Could not find required columns in evalSummary sheet');
    return;
  }

  Logger.log('Updating evaluation status: ' + evalId + ' to ' + newStatus);

  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === evalId) {
      sheet.getRange(i + 2, statusColIndex + 1).setValue(newStatus);
      Logger.log('Updated evaluation status successfully');

      // Invalidate evaluations cache
      CacheService.getScriptCache().remove('all_evaluations');
      return;
    }
  }

  Logger.log('Warning: Could not find evaluation with ID ' + evalId);
}

// Get pending audits for evaluation with improved performance
function getPendingAudits() {
  Logger.log('Fetching pending audits...');
  return getCachedOrFetch('pending_audits', function() {
    var audits = getAllAudits();
    Logger.log('Total audits fetched: ' + audits.length);

    var evaluations = getAllEvaluations();
    Logger.log('Total evaluations fetched: ' + evaluations.length);

    // Create a Set of evaluated audit IDs for faster lookups
    var evaluatedAuditIds = new Set(evaluations.map(evaluation => evaluation.evalId));
    Logger.log('Evaluated audit IDs: ' + Array.from(evaluatedAuditIds).join(', '));

    // Filter audits efficiently
    var pendingAudits = audits.filter(audit =>
      (audit.auditStatus.toLowerCase() === 'pending') && !evaluatedAuditIds.has(audit.auditId)
    );
    Logger.log('Pending audits count: ' + pendingAudits.length);

    return pendingAudits;
  });
}

// Function to get current user information
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  var users = getAllUsers();
  var user = users.find(function(user) {
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

function prepareEvaluation(auditId) {
    // Update the audit status to "In Process" and lock the record
    var audit = updateAuditStatusAndLock(auditId, 'In Process');
    if (!audit) {
        throw new Error('Failed to update audit status');
    }

    // Get the audit details
    var audits = getAllAudits();
    var auditDetails = audits.find(function(a) { return a.auditId === auditId; });

    if (!auditDetails) {
        throw new Error('Audit not found');
    }

    return auditDetails;
}

function resolveDispute(payload) {
  const { disputeId, evalId, decisions, resolutionNotes, status } = payload;
  const now = new Date().toISOString();
  const user = Session.getActiveUser().getEmail();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const disputesSheet = ss.getSheetByName('disputesQueue');
  const evalSheet = ss.getSheetByName('evalSummary');
  const questSheet = ss.getSheetByName('evalQuest');

  // 1. Update evalQuest scores & feedback
  const qHeaders = questSheet.getRange(1, 1, 1, questSheet.getLastColumn()).getValues()[0];
  const data = questSheet.getDataRange().getValues();

  const evalIdCol = qHeaders.indexOf('evalId');
  const qIdCol = qHeaders.indexOf('questionId');
  const earnedCol = qHeaders.indexOf('pointsEarned');
  const feedbackCol = qHeaders.indexOf('feedback');
  const pointsPossibleCol = qHeaders.indexOf('pointsPossible');

  data.forEach((row, i) => {
    if (row[evalIdCol] !== evalId) return;
    const qMatch = decisions.find(d => d.questionId === row[qIdCol]);
    if (!qMatch) return;

    const rowIndex = i + 1;
    if (qMatch.resolution === 'overturned') {
      questSheet.getRange(rowIndex, earnedCol + 1).setValue(row[pointsPossibleCol]);
    }
    if (qMatch.note) {
      questSheet.getRange(rowIndex, feedbackCol + 1).setValue(qMatch.note);
    }
  });

  // 2. Update evalSummary score
  const filtered = data.filter(r => r[evalIdCol] === evalId);
  const totalEarned = filtered.reduce((sum, r) => sum + (parseInt(r[earnedCol]) || 0), 0);
  const totalPossible = filtered.reduce((sum, r) => sum + (parseInt(r[pointsPossibleCol]) || 0), 0);
  const evalHeaders = evalSheet.getRange(1, 1, 1, evalSheet.getLastColumn()).getValues()[0];
  const rowIndexEval = evalSheet.getDataRange().getValues().findIndex(r => r[evalHeaders.indexOf('id')] === evalId);
  if (rowIndexEval >= 1) {
    evalSheet.getRange(rowIndexEval + 1, evalHeaders.indexOf('totalPoints') + 1).setValue(totalEarned);
    evalSheet.getRange(rowIndexEval + 1, evalHeaders.indexOf('totalPointsPossible') + 1).setValue(totalPossible);
    evalSheet.getRange(rowIndexEval + 1, evalHeaders.indexOf('evalScore') + 1)
      .setValue(Math.round((totalEarned / totalPossible) * 100) + '%');
  }

  // 3. Update dispute resolution metadata
  const dispHeaders = disputesSheet.getRange(1, 1, 1, disputesSheet.getLastColumn()).getValues()[0];
  const rowIndexDispute = disputesSheet.getDataRange().getValues().findIndex(r => r[dispHeaders.indexOf('id')] === disputeId);
  if (rowIndexDispute >= 1) {
    disputesSheet.getRange(rowIndexDispute + 1, dispHeaders.indexOf('status') + 1).setValue(status);
    disputesSheet.getRange(rowIndexDispute + 1, dispHeaders.indexOf('resolutionNotes') + 1).setValue(resolutionNotes);
    disputesSheet.getRange(rowIndexDispute + 1, dispHeaders.indexOf('resolvedBy') + 1).setValue(user);
    disputesSheet.getRange(rowIndexDispute + 1, dispHeaders.indexOf('resolutionTimestamp') + 1).setValue(now);
  }

  CacheService.getScriptCache().removeAll(['all_evaluations', 'all_disputes']);
}

function getAllEvaluationsAndDisputes() {
  return {
    evaluations: getAllEvaluations(), // assumes this function already exists
    disputes: getAllDisputes()        // assumes this function already exists
  };
}

function updateDisputeStatus(disputeId, newStatus) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Disputes");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == disputeId) { // assuming column A has the dispute ID
      sheet.getRange(i + 1, 5).setValue(newStatus); // assuming column E is 'status'
      return { success: true };
    }
  }

  return { success: false, message: "Dispute not found." };
}



