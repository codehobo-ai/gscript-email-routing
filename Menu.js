// ========================================
// MENU & UI
// Custom menu and user-facing functions
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📧 Gmail Tools')
    .addItem('🚀 Initial Setup', 'runInitialSetup')
    .addItem('🧭 Resume Setup Wizard', 'resumeWizard')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 Label Management')
      .addItem('Sync Labels & Counts', 'scheduledSync')
      .addItem('Discover New Labels', 'discoverNewLabelsUI'))
    .addItem('⚙️ Setup Triggers', 'setupTriggersUI')
    .addSeparator()
    .addItem('▶️ Run Processing Now', 'runEmailProcessing')
    .addItem('📥 Backfill Existing Labels', 'backfillExistingEmailsUI')
    .addItem('📊 View Status', 'showStatus')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏷️ Auto Label')
      .addItem('Dry Run Auto Label', 'dryRunAutoLabelUI')
      .addItem('Create Label Rules Sheet', 'createLabelRulesSheetUI'))
    .addSubMenu(ui.createMenu('🧹 Maintenance')
      .addItem('Smart Cleanup', 'smartCleanupUI')
      .addItem('Retry Errors', 'retryErrorsUI')
      .addItem('Sync Queue Headers', 'syncQueueHeadersUI')
      .addItem('Reset History Tracker', 'resetHistoryTrackerUI'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔗 Webhook')
      .addItem('Configure n8n Webhook', 'configureN8nWebhook')
      .addItem('Test Webhook', 'testWebhook'))
    .addItem('⏱️ View Active Triggers', 'viewTriggers')
    .addItem('🎨 Format Tables', 'formatTablesUI')
    .addToUi();

  // Check for updates (non-blocking toast)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const update = checkForUpdates();
  if (update) {
    const files = update.changedFiles && update.changedFiles.length > 0
      ? ` (${update.changedFiles.length} file${update.changedFiles.length > 1 ? 's' : ''} changed)`
      : '';
    ss.toast(
      `v${update.version} available${files}: ${update.changelog}\nSee: ${update.url}`,
      '📦 Update Available',
      15
    );
  }

  // First-run detection: show welcome dialog if setup hasn't been run yet
  if (!ss.getSheetByName(CONFIG.labelConfigSheetName)) {
    const response = ui.alert(
      '👋 Welcome to Gmail Email Router!',
      'It looks like this is your first time here.\n\n' +
      'Would you like to run the initial setup now?\n\n' +
      'This will import your Gmail labels and create everything you need to get started.',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      try {
        autoSetupLabels();
        ensureQueueSheet();
        ensureAutoLabelRulesSheet();
        createSetupGuideSheet();
        formatTables();
        postSetupWizard();
      } catch (e) {
        ui.alert(
          '⚠️ Setup Error',
          'Something went wrong: ' + e.message + '\n\n' +
          'Please try again via Gmail Tools → 🚀 Initial Setup.',
          ui.ButtonSet.OK
        );
      }
    }
  }
}

// --- SETUP ---

function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🚀 Initial Setup',
    'This will:\n' +
    '1. Import all Gmail labels into Label Config sheet\n' +
    '2. Create the Queue sheet\n' +
    '3. Create the Label Rules sheet (for auto-labeling)\n' +
    '4. Create the Setup Guide sheet\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  autoSetupLabels();
  ensureQueueSheet();
  ensureAutoLabelRulesSheet();
  createSetupGuideSheet();
  formatTables();

  postSetupWizard();
}

// --- POST-SETUP WIZARD ---

/**
 * Guided prompt sequence after initial setup.
 * Each step can be skipped — the user can always finish later from the menu.
 */
function postSetupWizard() {
  const ui = SpreadsheetApp.getUi();

  // --- Step 1: Activate labels ---
  const labelConfigs = getLabelConfigs();
  const activeCount = labelConfigs.filter(c => c.active).length;

  if (activeCount === 0) {
    const step1 = ui.alert(
      'Step 1 of 4: Activate Labels',
      'Setup created your Label Config sheet with ' + labelConfigs.length + ' Gmail labels.\n\n' +
      'Next, go to the "Label Config" sheet and check "active" for labels you want to monitor.\n\n' +
      'Click OK when done, or Cancel to finish later.',
      ui.ButtonSet.OK_CANCEL
    );
    if (step1 === ui.Button.CANCEL) {
      showWizardExitMessage_('Step 1: Activate labels in the Label Config sheet');
      return;
    }
  }

  // --- Step 2: Configure webhook ---
  const existingUrl = PropertiesService.getScriptProperties().getProperty('n8n_webhook_url');

  if (!existingUrl) {
    const step2 = ui.alert(
      'Step 2 of 4: Configure Webhook',
      'Do you have an n8n (or other) webhook URL to receive email data?\n\n' +
      'YES — enter your webhook URL now\n' +
      'NO — skip this (emails will be logged to the Queue sheet only)',
      ui.ButtonSet.YES_NO
    );

    if (step2 === ui.Button.YES) {
      const urlResponse = ui.prompt(
        'Enter Webhook URL',
        'Paste your webhook URL:',
        ui.ButtonSet.OK_CANCEL
      );

      if (urlResponse.getSelectedButton() === ui.Button.OK) {
        const url = urlResponse.getResponseText().trim();
        if (url.startsWith('http')) {
          PropertiesService.getScriptProperties().setProperty('n8n_webhook_url', url);
          ui.alert('Webhook saved.', '', ui.ButtonSet.OK);
        } else if (url) {
          ui.alert('Invalid URL — skipped. You can configure it later via Gmail Tools → Webhook.', '', ui.ButtonSet.OK);
        }
      }
    }
  }

  // --- Step 3: Test webhook (if configured) ---
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('n8n_webhook_url');

  if (webhookUrl) {
    const step3 = ui.alert(
      'Step 3 of 4: Test Webhook',
      'Your webhook is configured. Want to send a test payload to verify it works?\n\n' +
      'URL: ' + webhookUrl.substring(0, 50) + (webhookUrl.length > 50 ? '...' : ''),
      ui.ButtonSet.YES_NO
    );

    if (step3 === ui.Button.YES) {
      testWebhook();
    }
  }

  // --- Step 4: Setup triggers ---
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    const step4 = ui.alert(
      (webhookUrl ? 'Step 4 of 4' : 'Step 3 of 3') + ': Start Automation',
      'Ready to start the automated triggers?\n\n' +
      '• Auto-label + email processing: every 1 minute\n' +
      '• Label sync + counts: daily at 3 AM\n' +
      '• Queue cleanup: Sundays at 2 AM\n\n' +
      'The history tracker will initialize on the first trigger run.\n\n' +
      'You can also do this later via Gmail Tools → Setup Triggers.',
      ui.ButtonSet.YES_NO
    );

    if (step4 === ui.Button.YES) {
      setupTriggers();
      ui.alert(
        '🎉 You\'re all set!',
        'Automation is running. Emails from your active labels will be processed automatically.\n\n' +
        'Use Gmail Tools → Run Processing Now to verify, or check the Queue sheet anytime.\n\n' +
        'The "📋 Setup Guide" sheet has a full reference if you need it.',
        ui.ButtonSet.OK
      );
      return;
    }
  }

  if (triggers.length > 0) {
    ui.alert(
      '🎉 Setup Complete!',
      'Triggers are already running. You\'re good to go.\n\n' +
      'Check the "📋 Setup Guide" sheet for a full reference.',
      ui.ButtonSet.OK
    );
  } else {
    showWizardExitMessage_('Start triggers via Gmail Tools → Setup Triggers');
  }
}

/**
 * Re-entry point for users who left the wizard partway through.
 * Checks what's already configured and picks up from the next incomplete step.
 */
function resumeWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(CONFIG.labelConfigSheetName)) {
    const response = ui.alert(
      'Setup Not Started',
      'Initial setup hasn\'t been run yet. Run it now?',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) runInitialSetup();
    return;
  }

  const configs = getLabelConfigs();
  const activeCount = configs.filter(c => c.active).length;
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('n8n_webhook_url');
  const triggers = ScriptApp.getProjectTriggers();
  const historyId = PropertiesService.getScriptProperties().getProperty('gmail_last_history_id');

  const status = [];
  if (activeCount > 0) status.push('✅ ' + activeCount + ' label(s) active');
  else status.push('❌ No labels activated');

  if (webhookUrl) status.push('✅ Webhook configured');
  else status.push('⬜ No webhook (optional)');

  if (triggers.length > 0) status.push('✅ Triggers running (' + triggers.length + ')');
  else status.push('❌ Triggers not started');

  if (historyId) status.push('✅ History tracker active');
  else status.push('⬜ History tracker not initialized (starts on first trigger run)');

  const allDone = activeCount > 0 && triggers.length > 0;

  if (allDone) {
    ui.alert(
      '🎉 Everything Looks Good',
      'Current status:\n\n' + status.join('\n') + '\n\n' +
      'Your email router is fully configured and running.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '🧭 Setup Status',
    'Current status:\n\n' + status.join('\n') + '\n\n' +
    'Continue the setup wizard from where you left off?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    postSetupWizard();
  }
}

/**
 * Shows a consistent exit message when the user bails from the wizard.
 */
function showWizardExitMessage_(nextStep) {
  SpreadsheetApp.getUi().alert(
    'Setup Paused',
    'No problem — you can pick up where you left off anytime.\n\n' +
    'Next step: ' + nextStep + '\n\n' +
    'The "📋 Setup Guide" sheet has the full checklist.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --- STATUS ---

function showStatus() {
  const configs = getLabelConfigs();
  const activeLabels = configs.filter(c => c.active);
  const captureLabels = activeLabels.filter(c => c.captureToQueue);
  const n8nLabels = activeLabels.filter(c => c.sendToN8n);
  const triggers = ScriptApp.getProjectTriggers();
  const historyId = PropertiesService.getScriptProperties().getProperty('gmail_last_history_id');

  const { headers, rows } = getQueueData();
  const queueStats = { sent: 0, logged: 0, error: 0 };

  if (rows.length > 0) {
    const statusCol = headers.indexOf('status');
    rows.forEach(row => {
      const s = row[statusCol];
      if (queueStats.hasOwnProperty(s)) queueStats[s]++;
    });
  }

  const status =
    `📊 EMAIL PROCESSING STATUS\n\n` +
    `═══ CONFIGURATION ═══\n` +
    `Active: ${activeLabels.length} | Capturing: ${captureLabels.length} | n8n: ${n8nLabels.length}\n` +
    `Total labels: ${configs.length}\n\n` +
    `═══ AUTOMATION ═══\n` +
    `Triggers: ${triggers.length > 0 ? '✅ Active (' + triggers.length + ')' : '❌ Not Setup'}\n` +
    `History tracker: ${historyId ? '✅ Active' : '❌ Not initialized (run Setup Triggers)'}\n\n` +
    `═══ QUEUE ═══\n` +
    `Sent: ${queueStats.sent} | Logged: ${queueStats.logged} | Errors: ${queueStats.error}\n\n` +
    (activeLabels.length > 0
      ? `═══ ACTIVE LABELS ═══\n` +
        activeLabels.map(l => `• ${l.labelNameCurrent} → ${l.routeKey || '(no route key)'} ${l.sendToN8n ? '→ n8n' : '(log only)'}`).join('\n')
      : '');

  Logger.log(status);
  SpreadsheetApp.getUi().alert('System Status', status, SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- UI WRAPPERS ---

function setupTriggersUI() {
  const ui = SpreadsheetApp.getUi();
  setupTriggers();
  ui.alert(
    '✅ Automation Started',
    'Triggers:\n\n' +
    '• Auto-label + email processing: Every 1 minute\n' +
    '• Label sync + counts: Daily at 3 AM\n' +
    '• Queue cleanup: Sundays at 2 AM\n\n' +
    'The history tracker will initialize on the first trigger run.',
    ui.ButtonSet.OK
  );
}

function backfillExistingEmailsUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Backfill Existing Labels',
    'This will search for all emails currently carrying your active labels and add them to the Queue.\n\n' +
    'Already-queued messages are skipped automatically.\n\n' +
    'For large inboxes this may take a minute. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const count = backfillExistingEmails();
    ui.alert('Backfill Complete', `${count} new message(s) added to the Queue.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Backfill Failed', error.message, ui.ButtonSet.OK);
  }
}

function resetHistoryTrackerUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset History Tracker',
    'This clears the saved history ID. The next trigger run will re-initialize it to the current inbox state.\n\n' +
    'Use this if you see "History ID expired" errors, or to force a fresh start.\n\n' +
    'Run Backfill after resetting to catch any emails labeled during the gap.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  PropertiesService.getScriptProperties().deleteProperty('gmail_last_history_id');
  ui.alert('History tracker reset.', 'The next trigger run will re-initialize it.', ui.ButtonSet.OK);
}

function retryErrorsUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Retry Errors',
    'Retry all error rows (up to 3 attempts each)?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const count = retryErrors();
  ui.alert('Done', `${count} item(s) retried.`, ui.ButtonSet.OK);
}

function smartCleanupUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Smart Cleanup',
    'Update route keys for unsent items + reset error items for disabled labels?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  smartCleanup();
  ui.alert('Cleanup complete. Check Execution Log for details.', '', ui.ButtonSet.OK);
}

function createLabelRulesSheetUI() {
  ensureAutoLabelRulesSheet();
  formatTables();
  SpreadsheetApp.getUi().alert(
    'Label Rules Sheet Ready',
    'The "Label Rules" sheet is ready. Add rules and set them active to begin auto-labeling.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function syncQueueHeadersUI() {
  const missing = syncQueueHeaders();
  const ui = SpreadsheetApp.getUi();

  if (missing.length === 0) {
    ui.alert('No Changes', 'Queue sheet headers are already up to date.', ui.ButtonSet.OK);
  } else {
    ui.alert(
      'Headers Updated',
      `Added ${missing.length} missing column(s):\n\n• ${missing.join('\n• ')}\n\nExisting columns and data were not affected.`,
      ui.ButtonSet.OK
    );
  }
}

function discoverNewLabelsUI() {
  const newLabels = discoverNewLabels();
  const ui = SpreadsheetApp.getUi();

  if (newLabels.length === 0) {
    ui.alert('No New Labels', 'All Gmail labels are already in your config.', ui.ButtonSet.OK);
  } else {
    ui.alert('Found New Labels', `${newLabels.length} new label(s) added.\nThey're inactive by default — configure them in Label Config.`, ui.ButtonSet.OK);
  }
}
