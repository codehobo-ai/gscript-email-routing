// ========================================
// AUTO LABEL
// Query-based auto-labeling using Gmail search syntax
// ========================================

/**
 * Runs at the start of each 1-min trigger cycle.
 * Applies labels to recent unread inbox emails matching active rules.
 * Labels applied here are picked up by the History API on the next run.
 *
 * Uses newer_than:30m to handle occasional missed trigger runs without
 * scanning the full inbox. Already-labeled threads are skipped.
 */
function autoLabelEmails() {
  const rules = _getActiveRules();
  if (rules.length === 0) return;

  const stoppedThreads = new Set(); // thread IDs where stop_on_match fired
  let labeled = 0;

  rules.forEach(rule => {
    const query = `in:inbox is:unread newer_than:30m ${rule.query}`;
    let threads;
    try {
      threads = GmailApp.search(query, 0, 50);
    } catch (e) {
      Logger.log(`⚠️ Auto-label rule "${rule.ruleName}" search failed: ${e.message}`);
      return;
    }

    const label = _resolveLabel(rule.labelToApply);
    if (!label) return;

    threads.forEach(thread => {
      if (stoppedThreads.has(thread.getId())) return;

      const existingLabels = new Set(thread.getLabels().map(l => l.getName()));
      if (existingLabels.has(rule.labelToApply)) return;

      thread.addLabel(label);
      if (rule.markRead) thread.getMessages().forEach(m => m.markRead());
      if (rule.archive) thread.moveToArchive();
      labeled++;

      if (rule.stopOnMatch) stoppedThreads.add(thread.getId());
    });
  });

  if (labeled > 0) Logger.log(`🏷️ Auto-labeled ${labeled} thread(s)`);
}

/**
 * Dry run: shows which pre-existing emails would be labeled by each rule,
 * then offers to apply labels to all of them (backfill).
 *
 * unread_only column controls whether backfill is limited to unread emails.
 * Ongoing auto-labeling always targets unread; backfill respects the toggle.
 */
function dryRunAutoLabelUI() {
  const ui = SpreadsheetApp.getUi();
  const rules = _getActiveRules();

  if (rules.length === 0) {
    ui.alert('No Active Rules', 'Add rules to the Label Rules sheet and set them active first.', ui.ButtonSet.OK);
    return;
  }

  ui.alert('Scanning...', 'Checking pre-existing emails against your rules.\nThis may take a moment for large inboxes.', ui.ButtonSet.OK);

  const results = [];
  let totalNew = 0;

  rules.forEach(rule => {
    const query = rule.unreadOnly ? `is:unread ${rule.query}` : rule.query;
    let newCount = 0;
    const samples = [];
    let start = 0;
    const cap = 500;

    while (start < cap) {
      let threads;
      try {
        threads = GmailApp.search(query, start, 50);
      } catch (e) {
        Logger.log(`⚠️ Dry run rule "${rule.ruleName}" failed: ${e.message}`);
        break;
      }
      if (threads.length === 0) break;

      threads.forEach(thread => {
        const existingLabels = new Set(thread.getLabels().map(l => l.getName()));
        if (!existingLabels.has(rule.labelToApply)) {
          newCount++;
          if (samples.length < 3) samples.push(thread.getFirstMessageSubject() || '(no subject)');
        }
      });

      if (threads.length < 50) break;
      start += 50;
    }

    results.push({ rule, newCount, samples, capped: start >= cap });
    totalNew += newCount;
    Logger.log(`📋 "${rule.ruleName}": ${newCount} unlabeled match(es)`);
  });

  if (totalNew === 0) {
    ui.alert('No New Matches', 'All emails matching your rules are already labeled.', ui.ButtonSet.OK);
    return;
  }

  const summary = results
    .filter(r => r.newCount > 0)
    .map(r => {
      const cap = r.capped ? ' (capped at 500)' : '';
      const examples = r.samples.length > 0 ? `\n  e.g. ${r.samples.map(s => `"${s.substring(0, 50)}"`).join(', ')}` : '';
      return `• "${r.rule.ruleName}" → ${r.rule.labelToApply}: ${r.newCount} email(s)${cap}${examples}`;
    })
    .join('\n\n');

  const response = ui.alert(
    'Dry Run Results',
    `Found ${totalNew} unlabeled email(s) matching your rules:\n\n${summary}\n\n` +
    `Apply labels to all these pre-existing emails now?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const count = _runBackfill(rules);
  ui.alert('Done', `Labels applied to ${count} email(s).\nThe History API will process them on the next trigger run.`, ui.ButtonSet.OK);
}

// --- INTERNAL ---

function _getActiveRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.labelRulesSheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  return data.slice(1)
    .map(row => ({
      ruleName:     String(row[col('rule_name')] || '').trim(),
      active:       row[col('active')] === true || row[col('active')] === 'TRUE',
      priority:     Number(row[col('priority')]) || 99,
      query:        String(row[col('query')] || '').trim(),
      labelToApply: String(row[col('label_to_apply')] || '').trim(),
      stopOnMatch:  row[col('stop_on_match')] === true || row[col('stop_on_match')] === 'TRUE',
      unreadOnly:   row[col('unread_only')] === true || row[col('unread_only')] === 'TRUE',
      markRead:     row[col('mark_read')] === true || row[col('mark_read')] === 'TRUE',
      archive:      row[col('archive')] === true || row[col('archive')] === 'TRUE'
    }))
    .filter(r => r.ruleName && r.query && r.labelToApply && r.active)
    .sort((a, b) => a.priority - b.priority);
}

function _resolveLabel(labelName) {
  try {
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) Logger.log(`⚠️ Label not found: "${labelName}" — create it in Gmail first`);
    return label;
  } catch (e) {
    Logger.log(`⚠️ Could not resolve label "${labelName}": ${e.message}`);
    return null;
  }
}

function _runBackfill(rules) {
  const stoppedThreads = new Set();
  let total = 0;

  rules.forEach(rule => {
    const query = rule.unreadOnly ? `is:unread ${rule.query}` : rule.query;
    const label = _resolveLabel(rule.labelToApply);
    if (!label) return;

    let start = 0;
    while (true) {
      let threads;
      try {
        threads = GmailApp.search(query, start, 50);
      } catch (e) {
        Logger.log(`❌ Backfill rule "${rule.ruleName}" failed: ${e.message}`);
        break;
      }
      if (threads.length === 0) break;

      threads.forEach(thread => {
        if (stoppedThreads.has(thread.getId())) return;

        const existingLabels = new Set(thread.getLabels().map(l => l.getName()));
        if (existingLabels.has(rule.labelToApply)) return;

        thread.addLabel(label);
        if (rule.markRead) thread.getMessages().forEach(m => m.markRead());
        if (rule.archive) thread.moveToArchive();
        total++;

        if (rule.stopOnMatch) stoppedThreads.add(thread.getId());
      });

      if (threads.length < 50) break;
      start += 50;
    }

    Logger.log(`✅ Backfilled rule "${rule.ruleName}": applied "${rule.labelToApply}"`);
  });

  Logger.log(`📊 Auto-label backfill complete — ${total} label(s) applied`);
  return total;
}

// --- SHEET SETUP ---

function ensureAutoLabelRulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.labelRulesSheetName);

  if (sheet && sheet.getLastRow() > 0) {
    Logger.log('✅ Label Rules sheet already exists');
    return sheet;
  }

  if (!sheet) sheet = ss.insertSheet(CONFIG.labelRulesSheetName);

  const headers = [
    'rule_name', 'active', 'priority', 'query',
    'label_to_apply', 'stop_on_match', 'unread_only', 'mark_read', 'archive', 'notes'
  ];

  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // Checkbox columns
  const lastDataRow = Math.max(sheet.getMaxRows(), 100);
  ['active', 'stop_on_match', 'unread_only', 'mark_read', 'archive'].forEach(name => {
    const c = headers.indexOf(name) + 1;
    if (c > 0) {
      sheet.getRange(2, c, lastDataRow - 1, 1)
        .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
  });

  Logger.log('✅ Label Rules sheet created');
  return sheet;
}
