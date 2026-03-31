// ========================================
// WEBHOOK INTEGRATION
// n8n webhook functions
// ========================================

function viewWebhookSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('n8n_webhook_url');
  const authHeader = props.getProperty('n8n_webhook_auth_header');
  const authValue = props.getProperty('n8n_webhook_auth_value');

  if (!url) {
    ui.alert('No Webhook Configured', 'No webhook URL is set.\n\nUse Gmail Tools → Webhook → Configure to set one up.', ui.ButtonSet.OK);
    return;
  }

  const lines = [`URL: ${url}`];
  if (authHeader && authValue) {
    lines.push(`Auth header: ${authHeader}`);
    lines.push(`Auth value: ${'*'.repeat(Math.min(authValue.length, 12))}`);
  } else {
    lines.push('Auth: none');
  }

  const response = ui.alert(
    'Webhook Settings',
    lines.join('\n') + '\n\nWould you like to change these settings?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    configureN8nWebhook();
  }
}

function clearWebhookSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Webhook Settings',
    'This will remove the webhook URL and auth credentials.\n\nEmails will still be logged to the Queue sheet but no webhooks will be sent.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('n8n_webhook_url');
  props.deleteProperty('n8n_webhook_auth_header');
  props.deleteProperty('n8n_webhook_auth_value');
  ui.alert('Webhook settings cleared.', '', ui.ButtonSet.OK);
}

function configureN8nWebhook() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentUrl = props.getProperty('n8n_webhook_url');

  // Step 1: URL
  const promptMsg = currentUrl
    ? `Current URL: ${currentUrl.substring(0, 40)}...\n\nEnter a new webhook URL (or leave blank to keep current):`
    : 'Enter your n8n webhook URL:';

  const urlResponse = ui.prompt('Configure n8n Webhook (1/2)', promptMsg, ui.ButtonSet.OK_CANCEL);
  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;

  const url = urlResponse.getResponseText().trim();

  if (!url && currentUrl) {
    // Keep existing URL, move to auth step
  } else if (!url) {
    return;
  } else if (!url.startsWith('http')) {
    ui.alert('Invalid URL', 'URL must start with http:// or https://', ui.ButtonSet.OK);
    return;
  } else {
    props.setProperty('n8n_webhook_url', url);
  }

  // Step 2: Optional auth header name
  const headerNameResponse = ui.prompt(
    'Configure n8n Webhook (2/2) — Auth Header (optional)',
    'Enter the auth header name, e.g.  X-API-Key  or  Authorization\n\nLeave blank to send no auth header.',
    ui.ButtonSet.OK_CANCEL
  );
  if (headerNameResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Webhook Configured', `URL saved.\nNo auth header set.`, ui.ButtonSet.OK);
    return;
  }

  const headerName = headerNameResponse.getResponseText().trim();

  if (!headerName) {
    props.deleteProperty('n8n_webhook_auth_header');
    props.deleteProperty('n8n_webhook_auth_value');
    ui.alert('Webhook Configured', `URL saved.\nAuth header cleared.`, ui.ButtonSet.OK);
    return;
  }

  // Step 3: Auth header value
  const headerValueResponse = ui.prompt(
    `Auth Header Value for "${headerName}"`,
    'Enter the header value (e.g. your API key or "Bearer <token>"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (headerValueResponse.getSelectedButton() !== ui.Button.OK) return;

  const headerValue = headerValueResponse.getResponseText().trim();
  if (!headerValue) {
    ui.alert('No value entered', 'Auth header was not saved.', ui.ButtonSet.OK);
    return;
  }

  props.setProperty('n8n_webhook_auth_header', headerName);
  props.setProperty('n8n_webhook_auth_value', headerValue);

  ui.alert(
    'Webhook Configured',
    `URL saved.\nAuth header: ${headerName}: ${'*'.repeat(Math.min(headerValue.length, 8))}`,
    ui.ButtonSet.OK
  );
}

/**
 * Send a payload to the configured n8n webhook.
 * Returns { success, executionId, error } so the caller can write final status.
 */
function sendToN8nWebhook(payload) {
  const props = PropertiesService.getScriptProperties();
  const webhookUrl = props.getProperty('n8n_webhook_url');

  if (!webhookUrl) {
    Logger.log('⚠️ n8n webhook URL not configured');
    return { success: false, executionId: null, error: 'Webhook URL not configured' };
  }

  const authHeader = props.getProperty('n8n_webhook_auth_header');
  const authValue  = props.getProperty('n8n_webhook_auth_value');

  const headers = { 'Content-Type': 'application/json' };
  if (authHeader && authValue) headers[authHeader] = authValue;

  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      let executionId = null;
      try {
        const body = JSON.parse(response.getContentText());
        executionId = body.executionId || null;
      } catch (e) {
        // Non-JSON response — still a success
      }
      Logger.log(`✅ Webhook sent: ${payload.queueId || payload.messageId}${executionId ? ' (exec: ' + executionId + ')' : ''}`);
      return { success: true, executionId, error: null };
    } else {
      const error = `HTTP ${statusCode}: ${response.getContentText()}`;
      Logger.log(`⚠️ Webhook failed: ${error}`);
      return { success: false, executionId: null, error };
    }
  } catch (e) {
    Logger.log(`❌ Webhook error: ${e.message}`);
    return { success: false, executionId: null, error: e.message };
  }
}

function testWebhook() {
  const ui = SpreadsheetApp.getUi();
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('n8n_webhook_url');

  if (!webhookUrl) {
    ui.alert('No Webhook', 'Configure a webhook URL first via Gmail Tools → Webhook → Configure.', ui.ButtonSet.OK);
    return;
  }

  const payload = {
    queueId: 'test-' + Utilities.getUuid(),
    routeKey: 'test',
    labelName: 'Test Label',
    subject: 'Test message from Gmail Email Router',
    from: Session.getActiveUser().getEmail(),
    timestamp: new Date().toISOString(),
    source: 'apps_script_test'
  };

  const result = sendToN8nWebhook(payload);

  if (result.success) {
    ui.alert('Webhook OK', `Test payload sent successfully.${result.executionId ? '\n\nExecution ID: ' + result.executionId : ''}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Webhook Failed', `Error: ${result.error}`, ui.ButtonSet.OK);
  }
}
