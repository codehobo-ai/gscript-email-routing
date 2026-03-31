// ========================================
// QUEUE MANAGEMENT
// Queue operations and maintenance
// ========================================

/**
 * Statuses:
 *   sent   — webhook fired and n8n confirmed receipt
 *   logged — capture only (send_to_n8n = false), no webhook sent
 *   error  — webhook failed; see error_message column
 */
function addToQueue(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.queueSheetName);
  if (!sheet) throw new Error('Queue sheet not found. Run Initial Setup first.');

  const uniqueId = Utilities.getUuid();
  const now = new Date();
  const gmailLink = `https://mail.google.com/mail/u/0/#all/${data.threadId}`;

  let status = 'logged';
  let executionId = '';
  let errorMessage = '';

  if (data.sendToN8n) {
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tempRow = currentHeaders.map(h => {
      const map = {
        id: uniqueId, route_key: data.routeKey, label_name: data.labelName,
        label_id: data.labelId, message_id: data.messageId, thread_id: data.threadId,
        subject: data.subject, from_email: data.senderEmail, from_name: data.senderName,
        received_date: data.receivedDate, body_preview: data.body, file_ids: data.fileIds,
        file_details: data.fileDetails, attachment_count: data.attachmentCount
      };
      return map.hasOwnProperty(h) ? map[h] : '';
    });

    const result = sendToN8nWebhook(buildPayloadFromRow(currentHeaders, tempRow));
    status       = result.success ? 'sent' : 'error';
    executionId  = result.executionId || '';
    errorMessage = result.error || '';
  }

  const valueMap = {
    id:               uniqueId,
    created_at:       now,
    label_id:         data.labelId,
    route_key:        data.routeKey,
    label_name:       data.labelName,
    message_id:       data.messageId,
    thread_id:        data.threadId,
    gmail_link:       gmailLink,
    subject:          data.subject,
    from_email:       data.senderEmail,
    from_name:        data.senderName,
    received_date:    data.receivedDate,
    body_preview:     data.body,
    file_ids:         data.fileIds,
    file_details:     data.fileDetails,
    attachment_count: data.attachmentCount,
    status:           status,
    n8n_execution_id: executionId,
    error_message:    errorMessage,
    retry_count:      0
  };

  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = currentHeaders.map(h => valueMap.hasOwnProperty(h) ? valueMap[h] : '');
  sheet.appendRow(row);
}

function ensureQueueSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.queueSheetName);

  if (sheet && sheet.getLastRow() > 0) {
    Logger.log('✅ Queue sheet already exists');
    return sheet;
  }

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.queueSheetName);
  }

  sheet.appendRow(CONFIG.queueHeaders);
  sheet.getRange(1, 1, 1, CONFIG.queueHeaders.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, CONFIG.queueHeaders.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#f0faf0');

  Logger.log('✅ Queue sheet created');
  return sheet;
}

// --- RETRY ---

/**
 * Retry all error rows (up to 3 attempts each).
 * Updates status to 'sent' on success or keeps 'error' on continued failure.
 */
function retryErrors() {
  const { sheet, headers, data } = getQueueData();
  if (!sheet) return 0;

  const col = name => headers.indexOf(name);
  const maxRetries = 3;
  let count = 0;

  for (let i = data.length - 1; i > 0; i--) {
    const row = data[i];
    if (row[col('status')] !== 'error') continue;

    const retryCount = row[col('retry_count')] || 0;
    if (retryCount >= maxRetries) continue;

    const payload = buildPayloadFromRow(headers, row);
    payload.isRetry = true;
    payload.retryCount = retryCount + 1;

    const result = sendToN8nWebhook(payload);
    const newRetryCount = retryCount + 1;

    sheet.getRange(i + 1, col('status') + 1).setValue(result.success ? 'sent' : 'error');
    sheet.getRange(i + 1, col('retry_count') + 1).setValue(newRetryCount);
    sheet.getRange(i + 1, col('n8n_execution_id') + 1).setValue(result.executionId || '');
    sheet.getRange(i + 1, col('error_message') + 1).setValue(result.error || '');

    count++;
  }

  Logger.log(`📊 Retried ${count} error(s)`);
  return count;
}

// --- CLEANUP ---

function smartCleanup() {
  Logger.log('=== SMART CLEANUP ===\n');

  const { sheet, headers, data } = getQueueData();
  if (!sheet) {
    Logger.log('No queue items');
    return;
  }

  const configs = getLabelConfigs();
  const col = name => headers.indexOf(name);

  // 1. Update route keys for unsent items
  const labelToRouteKey = new Map(configs.map(c => [c.labelNameCurrent, c.routeKey]));
  let routeKeysUpdated = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][col('status')] === 'sent') continue;

    const correctKey = labelToRouteKey.get(data[i][col('label_name')]);
    if (correctKey && correctKey !== data[i][col('route_key')]) {
      sheet.getRange(i + 1, col('route_key') + 1).setValue(correctKey);
      routeKeysUpdated++;
    }
  }

  // 2. Reset error → logged for disabled labels
  const disabledLabels = new Set(configs.filter(c => !c.sendToN8n).map(c => c.labelNameCurrent));
  let itemsReset = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][col('status')] !== 'error') continue;
    if (disabledLabels.has(data[i][col('label_name')])) {
      sheet.getRange(i + 1, col('status') + 1).setValue('logged');
      itemsReset++;
    }
  }

  Logger.log(`Route keys updated: ${routeKeysUpdated}`);
  Logger.log(`Error items reset to logged (disabled labels): ${itemsReset}`);
}

function cleanupOldQueueItems() {
  const { sheet, headers, data } = getQueueData();
  if (!sheet) return;

  const col = name => headers.indexOf(name);
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  const rowsToDelete = [];
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][col('status')] === 'sent' && new Date(data[i][col('created_at')]) < thirtyDaysAgo) {
      rowsToDelete.push(i + 1);
    }
  }

  rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));

  Logger.log(`✅ Cleaned up ${rowsToDelete.length} old sent items`);
}

// --- HEADER SYNC ---

function syncQueueHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.queueSheetName);

  if (!sheet) {
    Logger.log('Queue sheet not found — skipping header sync');
    return [];
  }

  const expectedHeaders = CONFIG.queueHeaders;

  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  const missing = expectedHeaders.filter(h => !currentHeaders.includes(h));

  if (missing.length === 0) {
    Logger.log('✅ Queue headers are up to date');
    return [];
  }

  missing.forEach(header => {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol)
      .setValue(header)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
  });

  Logger.log(`✅ Queue header sync — added ${missing.length} column(s): ${missing.join(', ')}`);
  return missing;
}
