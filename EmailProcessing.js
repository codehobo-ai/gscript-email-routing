// ========================================
// EMAIL PROCESSING
// Gmail History API-based processing
// ========================================

/**
 * Main trigger function — processes emails newly labeled since last run.
 * Uses Gmail History API (labelsAdded events): only fires on new label events,
 * not on repeat scans. Already-queued messages are skipped via cached dedup set.
 */
function processNewEmails() {
  try {
    const configs = getLabelConfigs().filter(c => c.active && c.captureToQueue
      && !String(c.labelNameCurrent).includes('⚠️ DELETED'));

    if (configs.length === 0) return;

    const props = PropertiesService.getScriptProperties();
    const lastHistoryId = props.getProperty('gmail_last_history_id');

    if (!lastHistoryId) {
      _initializeHistoryId(props);
      Logger.log('ℹ️ History tracker initialized. Use 📥 Backfill Existing Labels to process pre-existing labeled emails.');
      return;
    }

    const labelConfigMap = new Map(configs.map(c => [c.labelId, c]));
    const monitoredLabelIds = new Set(configs.map(c => c.labelId));

    // Fetch history since last run
    let historyPage;
    try {
      historyPage = Gmail.Users.History.list('me', {
        startHistoryId: lastHistoryId,
        historyTypes: ['labelAdded'],
        maxResults: 500
      });
    } catch (e) {
      if (e.message && (e.message.includes('404') || e.message.includes('startHistoryId'))) {
        Logger.log('⚠️ History ID expired (>30 days). Resetting — run Backfill if needed.');
        props.deleteProperty('gmail_last_history_id');
        return;
      }
      throw e;
    }

    // Always advance the history cursor, even if no matching events
    if (historyPage.historyId) {
      props.setProperty('gmail_last_history_id', historyPage.historyId);
    }

    if (!historyPage.history || historyPage.history.length === 0) return;

    // Collect labelsAdded events for monitored labels (dedupe by messageId+labelId)
    const seen = new Set();
    const toProcess = [];

    historyPage.history.forEach(record => {
      if (!record.labelsAdded) return;
      record.labelsAdded.forEach(event => {
        const matchedLabelId = event.labelIds && event.labelIds.find(id => monitoredLabelIds.has(id));
        if (!matchedLabelId) return;
        const key = `${event.message.id}:${matchedLabelId}`;
        if (seen.has(key)) return;
        seen.add(key);
        toProcess.push({
          messageId: event.message.id,
          labelId: matchedLabelId,
          config: labelConfigMap.get(matchedLabelId)
        });
      });
    });

    if (toProcess.length === 0) return;

    Logger.log(`Found ${toProcess.length} newly labeled message(s)`);

    // Build processed set once — avoids re-reading the queue sheet per message
    const { headers, rows } = getQueueData();
    const processedSet = new Set();
    if (rows.length > 0) {
      const msgCol = headers.indexOf('message_id');
      rows.forEach(row => processedSet.add(row[msgCol]));
    }

    // Cache Drive folders per label to avoid redundant API calls
    const folderCache = new Map();
    let processed = 0, logged = 0;

    toProcess.forEach(item => {
      if (processedSet.has(item.messageId)) return;

      try {
        const message = GmailApp.getMessageById(item.messageId);
        if (!message) return;

        if (!folderCache.has(item.labelId)) {
          folderCache.set(item.labelId, _getDriveFolder(item.config));
        }

        const queueData = _extractMessageData(message, item.config, folderCache.get(item.labelId));
        addToQueue(queueData);
        processedSet.add(item.messageId);

        if (item.config.markRead) message.markRead();
        if (item.config.archive) message.getThread().moveToArchive();

        item.config.sendToN8n ? processed++ : logged++;
      } catch (e) {
        Logger.log(`❌ Failed to process ${item.messageId}: ${e.message}`);
      }
    });

    if (processed > 0 || logged > 0) {
      Logger.log(`📊 ${processed} processed, ${logged} logged`);
    }

  } catch (e) {
    Logger.log(`💥 FATAL ERROR: ${e.message}\n${e.stack}`);
  }
}

/**
 * One-time backfill: searches for all currently labeled emails and queues them.
 * Safe to run multiple times — already-queued messages are skipped via dedup.
 */
function backfillExistingEmails() {
  const configs = getLabelConfigs().filter(c => c.active && c.captureToQueue
    && !String(c.labelNameCurrent).includes('⚠️ DELETED'));

  if (configs.length === 0) {
    Logger.log('No active labels to backfill');
    return 0;
  }

  // Build processed set once for the whole backfill
  const { headers, rows } = getQueueData();
  const processedSet = new Set();
  if (rows.length > 0) {
    const msgCol = headers.indexOf('message_id');
    rows.forEach(row => processedSet.add(row[msgCol]));
  }

  let total = 0;

  configs.forEach(config => {
    let start = 0;
    const pageSize = 50;
    const folder = _getDriveFolder(config);

    while (true) {
      const threads = GmailApp.search(`label:"${config.labelNameCurrent}"`, start, pageSize);
      if (threads.length === 0) break;

      threads.forEach(thread => {
        thread.getMessages().forEach(message => {
          if (processedSet.has(message.getId())) return;
          try {
            addToQueue(_extractMessageData(message, config, folder));
            processedSet.add(message.getId());
            if (config.markRead) message.markRead();
            if (config.archive) thread.moveToArchive();
            total++;
          } catch (e) {
            Logger.log(`❌ Backfill failed for ${message.getId()}: ${e.message}`);
          }
        });
      });

      if (threads.length < pageSize) break;
      start += pageSize;
    }

    Logger.log(`✅ Backfilled: ${config.labelNameCurrent}`);
  });

  Logger.log(`📊 Backfill complete — ${total} new message(s) queued`);
  return total;
}

// --- INTERNAL HELPERS ---

function _initializeHistoryId(props) {
  const profile = Gmail.Users.getProfile('me');
  props.setProperty('gmail_last_history_id', String(profile.historyId));
  Logger.log(`✅ History ID initialized: ${profile.historyId}`);
}

function _getDriveFolder(config) {
  if (!config.sendToN8n || !config.driveFolderId) return null;
  try {
    return DriveApp.getFolderById(config.driveFolderId);
  } catch (e) {
    Logger.log(`⚠️ Drive folder not found: ${config.driveFolderId}`);
    return null;
  }
}

function _extractMessageData(message, config, folder) {
  const fromField = message.getFrom();
  let senderName = '', senderEmail = '';
  const match = fromField.match(/^"?([^"<]+)"?\s*<([^>]+)>$/);
  if (match) {
    senderName = match[1].trim();
    senderEmail = match[2].trim();
  } else {
    senderEmail = fromField;
  }

  const fileIds = [], fileDetails = [];
  if (config.sendToN8n && folder) {
    message.getAttachments().forEach(att => {
      if (att.getContentType() !== 'application/pdf') return;
      try {
        const file = folder.createFile(att);
        fileIds.push(file.getId());
        fileDetails.push({ id: file.getId(), name: file.getName(), url: file.getUrl(), size: file.getSize() });
        Logger.log(`📎 Uploaded: ${file.getName()}`);
      } catch (e) {
        Logger.log(`✗ Attachment upload failed: ${e.message}`);
      }
    });
  }

  return {
    labelId: config.labelId,
    labelName: config.labelNameCurrent,
    routeKey: config.routeKey,
    messageId: message.getId(),
    threadId: message.getThread().getId(),
    subject: message.getSubject() || '(No Subject)',
    senderEmail,
    senderName,
    receivedDate: message.getDate().toISOString(),
    body: message.getPlainBody().substring(0, 1000),
    fileIds: fileIds.join(','),
    fileDetails: JSON.stringify(fileDetails),
    attachmentCount: fileIds.length,
    sendToN8n: config.sendToN8n
  };
}
