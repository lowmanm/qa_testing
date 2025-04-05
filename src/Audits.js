// ====================
// Audit Module
// ====================

/**
 * Fetch all audits.
 */
function getAllAudits() {
  return getCachedOrFetch('all_audits', () => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
    return getSheetDataAsObjects(sheet);
  });
}

/**
 * Fetch pending audits (not evaluated yet).
 */
function getPendingAudits() {
  return getCachedOrFetch('pending_audits', () => {
    const audits = getAllAudits();
    const evaluations = getAllEvaluations();
    const evaluatedIds = new Set(evaluations.map(e => e.evalId));

    return audits.filter(a =>
      String(a.auditStatus).toLowerCase() === 'pending' &&
      !evaluatedIds.has(a.auditId)
    );
  });
}

/**
 * Mark audit as misconfigured when no questions are found.
 */
function markAuditAsMisconfigured(auditId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('auditId');
  const statusCol = headers.indexOf('auditStatus');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === auditId) {
      sheet.getRange(i + 1, statusCol + 1).setValue('misconfigured');
      CacheService.getScriptCache().removeAll(['all_audits', 'pending_audits']);
      return;
    }
  }

  throw new Error(`Audit ID ${auditId} not found`);
}

/**
 * Update audit status only.
 */
function updateAuditStatus(auditId, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('auditQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf('auditId');
  const statusIdx = headers.indexOf('auditStatus');

  const rowIndex = data.findIndex((r, i) => i > 0 && r[idIdx] === auditId);
  if (rowIndex === -1) return;

  sheet.getRange(rowIndex + 1, statusIdx + 1).setValue(newStatus);
  CacheService.getScriptCache().removeAll(['all_audits', 'pending_audits']);
}

/**
 * Update audit status and lock field (used when entering evaluation).
 */
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

  CacheService.getScriptCache().removeAll(['all_audits', 'pending_audits']);
  return { success: true };
}

/**
 * Lock the audit and return its details.
 */
function prepareEvaluation(auditId) {
  const result = updateAuditStatusAndLock(auditId, 'in_process');
  if (!result.success) throw new Error('Failed to update audit status');

  const audit = getAllAudits().find(a => a.auditId === auditId);
  if (!audit) throw new Error('Audit not found');

  return audit;
}
