// ====================
// Disputes Module
// ====================

/**
 * Fetch all disputes with questionIds as arrays.
 */
function getAllDisputes() {
  return getCachedOrFetch('all_disputes', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
    const data = getSheetDataAsObjects(sheet);

    data.forEach(d => {
      if (typeof d.questionIds === 'string') {
        d.questionIds = d.questionIds.split(',');
      }
    });

    return data;
  });
}

/**
 * Save a new dispute and update evaluation status.
 */
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

/**
 * Update dispute status (e.g., reviewing, resolved).
 */
function updateDisputeStatus(disputeId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('disputesQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('id');
  const statusIdx = headers.indexOf('status');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === disputeId);
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
    return { success: true };
  }

  return { success: false, message: 'Dispute not found' };
}

/**
 * Apply dispute resolution to evaluation.
 */
function resolveDispute(resolution) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const questSheet = ss.getSheetByName('evalQuest');
  const summarySheet = ss.getSheetByName('evalSummary');
  const disputeSheet = ss.getSheetByName('disputesQueue');

  const { disputeId, evalId, decisions, resolutionNotes, status } = resolution;
  const resolvedBy = Session.getActiveUser().getEmail();
  const resolvedAt = new Date().toISOString();

  const questData = questSheet.getDataRange().getValues();
  const headers = questData[0];
  const indices = {
    evalId: headers.indexOf('evalId'),
    questionId: headers.indexOf('questionId'),
    response: headers.indexOf('response'),
    pointsEarned: headers.indexOf('pointsEarned'),
    feedback: headers.indexOf('feedback'),
    pointsPossible: headers.indexOf('pointsPossible')
  };

  for (let i = 1; i < questData.length; i++) {
    const row = questData[i];
    if (row[indices.evalId] !== evalId) continue;

    const decision = decisions.find(d => d.questionId === row[indices.questionId]);
    if (!decision) continue;

    if (decision.resolution === 'overturned') {
      row[indices.response] = 'yes';
      row[indices.pointsEarned] = row[indices.pointsPossible];
    }

    row[indices.feedback] = decision.note || '';
    questSheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
  }

  // Recalculate total score
  const updated = getSheetDataAsObjects(questSheet).filter(q => q.evalId === evalId);
  const totalPoints = updated.reduce((sum, q) => sum + (parseFloat(q.pointsEarned) || 0), 0);
  const totalPossible = updated.reduce((sum, q) => sum + (parseFloat(q.pointsPossible) || 0), 0);
  const evalScore = totalPossible > 0 ? totalPoints / totalPossible : 0;

  const summaryData = summarySheet.getDataRange().getValues();
  const sHeaders = summaryData[0];
  const sIdx = {
    id: sHeaders.indexOf('id'),
    totalPoints: sHeaders.indexOf('totalPoints'),
    evalScore: sHeaders.indexOf('evalScore'),
    status: sHeaders.indexOf('status')
  };

  for (let i = 1; i < summaryData.length; i++) {
    const row = summaryData[i];
    if (row[sIdx.id] === evalId) {
      summarySheet.getRange(i + 1, sIdx.totalPoints + 1).setValue(totalPoints);
      summarySheet.getRange(i + 1, sIdx.evalScore + 1).setValue(evalScore);
      summarySheet.getRange(i + 1, sIdx.status + 1).setValue(status || 'resolved');
      break;
    }
  }

  const disputeData = disputeSheet.getDataRange().getValues();
  const dHeaders = disputeData[0];
  const dIdx = {
    id: dHeaders.indexOf('id'),
    status: dHeaders.indexOf('status'),
    notes: dHeaders.indexOf('resolutionNotes'),
    by: dHeaders.indexOf('resolvedBy'),
    timestamp: dHeaders.indexOf('resolutionTimestamp')
  };

  for (let i = 1; i < disputeData.length; i++) {
    const row = disputeData[i];
    if (row[dIdx.id] === disputeId) {
      disputeSheet.getRange(i + 1, dIdx.status + 1).setValue(status || 'resolved');
      disputeSheet.getRange(i + 1, dIdx.notes + 1).setValue(resolutionNotes || '');
      disputeSheet.getRange(i + 1, dIdx.by + 1).setValue(resolvedBy);
      disputeSheet.getRange(i + 1, dIdx.timestamp + 1).setValue(resolvedAt);
      break;
    }
  }

  CacheService.getScriptCache().removeAll(['all_disputes', 'all_evaluations']);
  return true;
}

/**
 * Combine all evaluations and disputes.
 */
function getAllEvaluationsAndDisputes() {
  return {
    evaluations: getAllEvaluations(),
    disputes: getAllDisputes()
  };
}

/**
 * Generate stats for dispute summary.
 */
function getDisputeStats() {
  const data = getAllDisputes();

  let total = 0, partialOverturns = 0, overturned = 0, upheld = 0;

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
