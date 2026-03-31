// ========================================
// UTILITIES
// Shared helper functions
// ========================================

function getLabelConfigs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.labelConfigSheetName);

  if (!sheet) {
    throw new Error('Label Config sheet not found. Run Initial Setup first.');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  return data.slice(1).map(row => ({
    labelId:          row[col('label_id')],
    labelNameCurrent: row[col('label_name_current')],
    routeKey:         row[col('route_key')],
    driveFolderId:    row[col('drive_folder_id')],
    captureToQueue:   row[col('capture_to_queue')] === true || row[col('capture_to_queue')] === 'TRUE',
    sendToN8n:        row[col('send_to_n8n')] === true || row[col('send_to_n8n')] === 'TRUE',
    active:           row[col('active')] === true || row[col('active')] === 'TRUE',
    markRead:         col('mark_read') < 0 || (row[col('mark_read')] !== false && row[col('mark_read')] !== 'FALSE'),
    archive:          row[col('archive')] === true || row[col('archive')] === 'TRUE'
  })).filter(c => c.labelId && c.labelId.startsWith('Label_'));
}

/**
 * Build a webhook payload from a queue sheet row (used by retry and initial send).
 */
function buildPayloadFromRow(headers, row) {
  const col = name => headers.indexOf(name);

  return {
    queueId:         row[col('id')],
    routeKey:        row[col('route_key')],
    labelName:       row[col('label_name')],
    labelId:         row[col('label_id')],
    messageId:       row[col('message_id')],
    threadId:        row[col('thread_id')],
    subject:         row[col('subject')],
    from:            row[col('from_email')],
    fromName:        row[col('from_name')],
    receivedDate:    row[col('received_date')],
    bodyPreview:     row[col('body_preview')],
    fileIds:         row[col('file_ids')],
    fileDetails:     row[col('file_details')],
    attachmentCount: row[col('attachment_count')],
    timestamp:       new Date().toISOString(),
    source:          'apps_script_automated'
  };
}

/**
 * Get queue sheet headers and all data rows in one read.
 */
function getQueueData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.queueSheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return { sheet: null, headers: [], rows: [], data: [] };
  }

  const data = sheet.getDataRange().getValues();
  return { sheet, headers: data[0], rows: data.slice(1), data };
}
