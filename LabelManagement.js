// ========================================
// LABEL MANAGEMENT
// Setup, sync, discovery, and email count tracking
// ========================================

// --- SETUP ---

function autoSetupLabels() {
  Logger.log('=== AUTO SETUP STARTING ===\n');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.labelConfigSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.labelConfigSheetName);
  }

  sheet.clear();

  const headers = [
    'label_id',
    'label_name_current',
    'route_key',
    'drive_folder_id',
    'capture_to_queue',
    'send_to_n8n',
    'active',
    'mark_read',
    'archive',
    'email_count',
    'prev_email_count',
    'last_synced',
    'notes'
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  const labels = Gmail.Users.Labels.list('me').labels;
  const userLabels = labels.filter(l => l.type === 'user');
  userLabels.sort((a, b) => a.name.localeCompare(b.name));

  Logger.log(`Found ${userLabels.length} Gmail labels\n`);

  userLabels.forEach(label => {
    const labelDetail = Gmail.Users.Labels.get('me', label.id);
    const emailCount = labelDetail.messagesTotal || 0;

    sheet.appendRow([
      label.id,
      label.name,
      '',           // route_key — user fills in
      '',           // drive_folder_id
      true,         // capture_to_queue
      false,        // send_to_n8n
      false,        // active
      true,         // mark_read
      false,        // archive
      emailCount,
      emailCount,   // prev_email_count (same on first setup)
      new Date(),
      'Auto-imported'
    ]);

    Logger.log(`Added: ${label.name} (${emailCount} emails)`);
  });

  // Format sheet
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const checkboxCols = ['capture_to_queue', 'send_to_n8n', 'active', 'mark_read', 'archive'];
    checkboxCols.forEach(colName => {
      const colIdx = headers.indexOf(colName) + 1;
      sheet.getRange(2, colIdx, lastRow - 1, 1).insertCheckboxes();
    });

    const activeCol = headers.indexOf('active') + 1;
    sheet.getRange(2, activeCol, lastRow - 1, 1).setBackground('#fff2cc');
  }

  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);

  // Alternating row banding
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#e8f0fe');

  ss.setActiveSheet(sheet);

  Logger.log(`\n✅ Imported ${userLabels.length} labels`);
  return sheet;
}

// --- SYNC ---

function syncLabelNames(gmailLabels) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.labelConfigSheetName);
  if (!sheet) throw new Error('Label Config sheet not found');

  if (!gmailLabels) gmailLabels = Gmail.Users.Labels.list('me').labels;
  const labelMap = new Map(gmailLabels.map(l => [l.id, l.name]));

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  let updatedCount = 0;
  let missingCount = 0;

  for (let i = 1; i < data.length; i++) {
    const labelId = data[i][col('label_id')];
    if (!labelId || !labelId.startsWith('Label_')) continue;

    const gmailName = labelMap.get(labelId);
    const sheetName = data[i][col('label_name_current')];

    if (!gmailName) {
      sheet.getRange(i + 1, col('label_name_current') + 1).setValue('⚠️ DELETED').setBackground('#ffcccc');
      sheet.getRange(i + 1, col('route_key') + 1).setValue('');
      missingCount++;
    } else if (gmailName !== sheetName) {
      sheet.getRange(i + 1, col('label_name_current') + 1).setValue(gmailName).setBackground('#fff2cc');
      updatedCount++;
    }

    sheet.getRange(i + 1, col('last_synced') + 1).setValue(new Date());
  }

  Logger.log(`✅ Label sync: ${updatedCount} renamed, ${missingCount} missing`);
  return { updatedCount, missingCount };
}

function discoverNewLabels(gmailLabels) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.labelConfigSheetName);
  if (!sheet) throw new Error('Label Config sheet not found');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const labelIdCol = headers.indexOf('label_id');
  const existingIds = new Set(data.slice(1).map(row => row[labelIdCol]).filter(Boolean));

  if (!gmailLabels) gmailLabels = Gmail.Users.Labels.list('me').labels;
  const newLabels = gmailLabels.filter(l => l.type === 'user' && !existingIds.has(l.id));

  if (newLabels.length === 0) {
    Logger.log('✅ No new labels found');
    return [];
  }

  newLabels.forEach(label => {
    const labelDetail = Gmail.Users.Labels.get('me', label.id);
    const emailCount = labelDetail.messagesTotal || 0;

    sheet.appendRow([
      label.id,
      label.name,
      '',           // route_key
      '',           // drive_folder_id
      true,         // capture_to_queue
      false,        // send_to_n8n
      false,        // active
      true,         // mark_read
      false,        // archive
      emailCount,
      emailCount,
      new Date(),
      '🆕 Auto-discovered'
    ]);

    Logger.log(`➕ Added: ${label.name} (${emailCount} emails)`);
  });

  const lastRow = sheet.getLastRow();
  const newRowStart = lastRow - newLabels.length + 1;
  sheet.getRange(newRowStart, 1, newLabels.length, headers.length).setBackground('#e8f5e9');

  Logger.log(`✅ Discovered ${newLabels.length} new labels`);
  return newLabels;
}

// --- EMAIL COUNT TRACKING ---

function updateEmailCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.labelConfigSheetName);
  if (!sheet) throw new Error('Label Config sheet not found');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  const changes = [];

  for (let i = 1; i < data.length; i++) {
    const labelId = data[i][col('label_id')];
    if (!labelId || !labelId.startsWith('Label_')) continue;

    const labelName = data[i][col('label_name_current')];
    if (labelName && labelName.includes('⚠️ DELETED')) continue;

    try {
      const labelDetail = Gmail.Users.Labels.get('me', labelId);
      const newCount = labelDetail.messagesTotal || 0;
      const prevCount = data[i][col('email_count')] || 0;

      sheet.getRange(i + 1, col('prev_email_count') + 1).setValue(prevCount);
      sheet.getRange(i + 1, col('email_count') + 1).setValue(newCount);

      if (newCount !== prevCount) {
        changes.push({
          labelName,
          labelId,
          previousCount: prevCount,
          currentCount: newCount,
          delta: newCount - prevCount
        });
      }
    } catch (e) {
      Logger.log(`⚠️ Could not get count for ${labelId}: ${e.message}`);
    }
  }

  Logger.log(`📊 Email counts updated. ${changes.length} label(s) changed.`);
  return changes;
}

// --- FULL SYNC (runs on schedule) ---

function scheduledSync() {
  Logger.log('=== SCHEDULED SYNC STARTING ===\n');

  // Fetch Gmail labels once for both sync and discovery
  const gmailLabels = Gmail.Users.Labels.list('me').labels;

  const { updatedCount, missingCount } = syncLabelNames(gmailLabels);
  const newLabels = discoverNewLabels(gmailLabels);
  const countChanges = updateEmailCounts();

  syncQueueHeaders();

  const hasChanges = updatedCount > 0 || missingCount > 0 || newLabels.length > 0 || countChanges.length > 0;

  if (hasChanges) {
    sendChangeNotification({
      renamedCount: updatedCount,
      missingCount,
      newLabels,
      countChanges
    });
  } else {
    Logger.log('✅ No changes detected — no notification sent.');
  }

  Logger.log('\n=== SYNC COMPLETE ===');
}

// --- CHANGE-ONLY NOTIFICATION ---

function sendChangeNotification(changes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let subject = '📋 Gmail Label Changes Detected';
    let body = `Label sync completed at ${new Date().toLocaleString()}\n\n`;

    if (changes.renamedCount > 0) {
      body += `✏️ ${changes.renamedCount} label(s) renamed (highlighted yellow in sheet)\n`;
    }

    if (changes.missingCount > 0) {
      body += `⚠️ ${changes.missingCount} label(s) deleted from Gmail (highlighted red in sheet)\n`;
      subject = '⚠️ Gmail Labels Changed';
    }

    if (changes.newLabels.length > 0) {
      body += `\n🆕 ${changes.newLabels.length} new label(s) discovered:\n`;
      changes.newLabels.forEach(l => {
        body += `  • ${l.name}\n`;
      });
      body += '\nNew labels are inactive by default — configure them in the Label Config sheet.\n';
    }

    if (changes.countChanges.length > 0) {
      body += `\n📊 Email count changes:\n`;
      changes.countChanges.forEach(c => {
        const direction = c.delta > 0 ? `+${c.delta}` : `${c.delta}`;
        body += `  • ${c.labelName}: ${c.previousCount} → ${c.currentCount} (${direction})\n`;
      });
    }

    body += `\nSpreadsheet: ${ss.getUrl()}`;

    MailApp.sendEmail({
      to: CONFIG.notificationEmail,
      subject,
      body
    });

    Logger.log('📧 Change notification sent');
  } catch (e) {
    Logger.log(`Failed to send notification: ${e.message}`);
  }
}
