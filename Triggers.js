// ========================================
// TRIGGERS
// Trigger management and scheduling
// ========================================

function setupTriggers() {
  // Clear existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Auto-label + process newly labeled emails every minute
  ScriptApp.newTrigger('runEmailProcessing')
    .timeBased()
    .everyMinutes(1)
    .create();

  // Sync label names, discover new labels, update counts — daily at 3 AM
  ScriptApp.newTrigger('scheduledSync')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();

  // Cleanup old processed items weekly
  ScriptApp.newTrigger('cleanupOldQueueItems')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();

  Logger.log('✅ Triggers created');
}

/**
 * Main 1-min trigger entry point.
 * Auto-labels first, then History API picks up newly applied labels.
 */
function runEmailProcessing() {
  autoLabelEmails();
  processNewEmails();
}

function viewTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const ui = SpreadsheetApp.getUi();

  if (triggers.length === 0) {
    ui.alert('No Triggers', 'No triggers are configured.\n\nGo to Gmail Tools → Setup Triggers to start automation.', ui.ButtonSet.OK);
    return;
  }

  const lines = triggers.map(trigger =>
    `• ${trigger.getHandlerFunction()} (${trigger.getEventType()})`
  );

  ui.alert(
    'Active Triggers',
    `${triggers.length} trigger(s) configured:\n\n${lines.join('\n')}`,
    ui.ButtonSet.OK
  );
}
