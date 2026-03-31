// ========================================
// CONFIGURATION
// Global constants and settings
// ========================================

const VERSION = '4.0.0';
const VERSION_CHECK_URL = 'https://raw.githubusercontent.com/codehobo-ai/gscript-email-routing/main/version.json';

const CONFIG = {
  queueSheetName: 'Queue',
  labelConfigSheetName: 'Label Config',
  labelRulesSheetName: 'Label Rules',
  notificationEmail: Session.getActiveUser().getEmail(),
  queueHeaders: [
    'id', 'created_at', 'label_id', 'route_key', 'label_name',
    'message_id', 'thread_id', 'gmail_link', 'subject', 'from_email', 'from_name',
    'received_date', 'body_preview', 'file_ids', 'file_details',
    'attachment_count', 'status', 'n8n_execution_id', 'error_message', 'retry_count'
  ]
};

/**
 * Check for updates against the published version.json.
 * Non-blocking — failures are silently ignored.
 */
function checkForUpdates() {
  try {
    const response = UrlFetchApp.fetch(VERSION_CHECK_URL, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;

    const remote = JSON.parse(response.getContentText());
    if (remote.version === VERSION) return null;

    return remote;
  } catch (e) {
    // Network error, blocked by corporate firewall, etc. — fail silently
    return null;
  }
}
