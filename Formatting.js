// ========================================
// FORMATTING
// Sheet formatting, column widths, date formats, conditional rules
// ========================================

/**
 * Apply consistent formatting to both the Label Config and Queue sheets.
 * Safe to run multiple times — conditional rules are replaced, not stacked.
 */
function formatTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const labelSheet = ss.getSheetByName(CONFIG.labelConfigSheetName);
  if (labelSheet) _formatLabelConfigSheet(labelSheet);

  const queueSheet = ss.getSheetByName(CONFIG.queueSheetName);
  if (queueSheet) _formatQueueSheet(queueSheet);

  const rulesSheet = ss.getSheetByName(CONFIG.labelRulesSheetName);
  if (rulesSheet) _formatLabelRulesSheet(rulesSheet);

  Logger.log('✅ Tables formatted');
}

function formatTablesUI() {
  formatTables();
  SpreadsheetApp.getUi().alert(
    'Tables Formatted',
    'Column widths, date formats, and conditional highlighting applied to all sheets.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --- INTERNAL HELPERS ---

function _formatLabelConfigSheet(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = name => headers.indexOf(name) + 1;

  const widths = {
    label_id: 120,
    label_name_current: 200,
    route_key: 120,
    drive_folder_id: 160,
    capture_to_queue: 90,
    send_to_n8n: 90,
    active: 70,
    mark_read: 80,
    archive: 70,
    email_count: 90,
    prev_email_count: 110,
    last_synced: 155,
    notes: 200
  };
  headers.forEach((h, i) => {
    if (widths[h]) sheet.setColumnWidth(i + 1, widths[h]);
  });

  // Ensure checkbox columns have checkboxes on all data rows
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    ['capture_to_queue', 'send_to_n8n', 'active', 'mark_read', 'archive'].forEach(name => {
      const c = col(name);
      if (c > 0) sheet.getRange(2, c, lastRow - 1, 1).insertCheckboxes();
    });

    ['last_synced'].forEach(name => {
      const c = col(name);
      if (c > 0) sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    });
  }

  const maxRows = sheet.getMaxRows();
  const dataRange = sheet.getRange(2, 1, maxRows - 1, lastCol);
  const rules = [];

  const emailCountCol = col('email_count');
  if (emailCountCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${_colLetter(emailCountCol)}2=0`)
      .setBackground('#fce8e6')
      .setRanges([dataRange])
      .build());
  }

  const routeKeyCol = col('route_key');
  if (routeKeyCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${_colLetter(routeKeyCol)}2=""`)
      .setBackground('#fff8e1')
      .setRanges([dataRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
}

function _formatQueueSheet(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = name => headers.indexOf(name) + 1;

  const widths = {
    id: 220,
    created_at: 155,
    label_id: 110,
    route_key: 110,
    label_name: 160,
    message_id: 140,
    thread_id: 140,
    gmail_link: 80,
    subject: 250,
    from_email: 180,
    from_name: 140,
    received_date: 155,
    body_preview: 220,
    file_ids: 110,
    file_details: 110,
    attachment_count: 80,
    status: 90,
    n8n_execution_id: 160,
    error_message: 200,
    retry_count: 70
  };
  headers.forEach((h, i) => {
    if (widths[h]) sheet.setColumnWidth(i + 1, widths[h]);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    ['created_at', 'received_date'].forEach(name => {
      const c = col(name);
      if (c > 0) sheet.getRange(2, c, lastRow - 1, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    });
  }

  const maxRows = sheet.getMaxRows();
  const dataRange = sheet.getRange(2, 1, maxRows - 1, lastCol);
  const rules = [];

  const statusCol = col('status');
  if (statusCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${_colLetter(statusCol)}2="error"`)
      .setBackground('#fce8e6')
      .setRanges([dataRange])
      .build());
  }

  const routeKeyCol = col('route_key');
  if (routeKeyCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${_colLetter(routeKeyCol)}2=""`)
      .setBackground('#fff8e1')
      .setRanges([dataRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
}

function _formatLabelRulesSheet(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const widths = {
    rule_name:     180,
    active:         70,
    priority:       70,
    query:         300,
    label_to_apply: 180,
    stop_on_match:  90,
    unread_only:    90,
    notes:         220
  };
  headers.forEach((h, i) => {
    if (widths[h]) sheet.setColumnWidth(i + 1, widths[h]);
  });

  // Ensure checkbox columns have checkboxes on all data rows
  const maxRows = sheet.getMaxRows();
  const col = name => headers.indexOf(name) + 1;

  ['active', 'stop_on_match', 'unread_only'].forEach(name => {
    const c = col(name);
    if (c > 0) sheet.getRange(2, c, maxRows - 1, 1).insertCheckboxes();
  });

  const dataRange = sheet.getRange(2, 1, maxRows - 1, lastCol);
  const rules = [];

  const activeCol = col('active');
  if (activeCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${_colLetter(activeCol)}2=FALSE`)
      .setBackground('#f3f3f3')
      .setFontColor('#aaaaaa')
      .setRanges([dataRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);

  // label_to_apply dropdown from Label Config label_name_current
  const labelToApplyCol = col('label_to_apply');
  if (labelToApplyCol > 0) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const labelConfigSheet = ss.getSheetByName(CONFIG.labelConfigSheetName);
    if (labelConfigSheet && labelConfigSheet.getLastRow() > 1) {
      const labelNameCol = labelConfigSheet.getRange(1, 1, 1, labelConfigSheet.getLastColumn()).getValues()[0].indexOf('label_name_current') + 1;
      if (labelNameCol > 0) {
        const labelRange = labelConfigSheet.getRange(2, labelNameCol, labelConfigSheet.getLastRow() - 1, 1);
        const validation = SpreadsheetApp.newDataValidation()
          .requireValueInRange(labelRange, true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, labelToApplyCol, maxRows - 1, 1).setDataValidation(validation);
      }
    }
  }
}

/**
 * Convert a 1-based column index to a column letter (1=A, 2=B, 27=AA, etc.)
 */
function _colLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
